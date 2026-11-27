"""Durable, per-run crawler checkpoints. Data stays on the WorkLab host."""
from __future__ import annotations

import json
import os
import re
import sqlite3
import uuid
from contextlib import closing
from datetime import datetime
from pathlib import Path


def checkpoint_directory() -> Path:
    local = os.environ.get("LOCALAPPDATA")
    root = Path(local) / "WorkLab_PAOS" if local else Path.home() / ".worklab_paos"
    return root / "crawl_checkpoints"


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _dump(value) -> str:
    return json.dumps(value, ensure_ascii=False, default=str)


def _path(job_id: str) -> Path:
    if not re.fullmatch(r"[0-9a-f]{32}", job_id):
        raise FileNotFoundError("Invalid checkpoint ID")
    return checkpoint_directory() / f"{job_id}.sqlite3"


def _connect(path: Path):
    # Each call opens its own connection: live downloads never share the writer connection.
    conn = sqlite3.connect(path, timeout=5)
    conn.execute("PRAGMA synchronous=FULL")
    return conn


def create_checkpoint(kind: str, inputs: list[str], columns: list[str], options: dict, clean_sheet: bool, mode: str) -> tuple[str, str]:
    folder = checkpoint_directory()
    folder.mkdir(parents=True, exist_ok=True)
    job_id = uuid.uuid4().hex
    path = _path(job_id)
    created = _now()
    meta = {
        "kind": kind, "mode": mode, "columns": columns, "options": options,
        "clean_sheet": clean_sheet, "created_at": created, "updated_at": created,
        "total": len(inputs), "processed": 0, "state": "running", "row_count": 0,
    }
    with closing(_connect(path)) as conn, conn:
        conn.execute("CREATE TABLE meta (id INTEGER PRIMARY KEY, payload TEXT NOT NULL)")
        conn.execute("CREATE TABLE results (seq INTEGER PRIMARY KEY, payload TEXT NOT NULL)")
        conn.execute("CREATE TABLE inputs (seq INTEGER PRIMARY KEY, value TEXT NOT NULL)")
        conn.execute("INSERT INTO meta VALUES (1, ?)", (_dump(meta),))
        conn.executemany("INSERT INTO inputs VALUES (?, ?)", enumerate(inputs, 1))
    return job_id, created


def update_checkpoint(job_id: str, rows: dict[int, dict], changes: dict) -> None:
    path = _path(job_id)
    if not path.is_file():
        raise FileNotFoundError("Checkpoint not found")
    with closing(_connect(path)) as conn, conn:
        meta = json.loads(conn.execute("SELECT payload FROM meta WHERE id=1").fetchone()[0])
        if rows:
            conn.executemany("INSERT OR REPLACE INTO results VALUES (?, ?)",
                             ((seq, _dump(row)) for seq, row in rows.items()))
        meta.update(changes)
        meta["row_count"] = conn.execute("SELECT COUNT(*) FROM results").fetchone()[0]
        meta["updated_at"] = _now()
        conn.execute("UPDATE meta SET payload=? WHERE id=1", (_dump(meta),))


def read_checkpoint(job_id: str) -> tuple[dict, list[dict], list[str]]:
    path = _path(job_id)
    if not path.is_file():
        raise FileNotFoundError("Checkpoint not found")
    # One read transaction for a consistent snapshot while the crawler keeps writing.
    with closing(_connect(path)) as conn, conn:
        conn.execute("BEGIN")
        meta = json.loads(conn.execute("SELECT payload FROM meta WHERE id=1").fetchone()[0])
        rows = [json.loads(row[0]) for row in conn.execute("SELECT payload FROM results ORDER BY seq")]
        remaining = [row[0] for row in conn.execute("SELECT value FROM inputs WHERE seq > ? ORDER BY seq", (meta["processed"],))]
    return meta, rows, remaining


def tracked_crawl(events, *, kind, inputs, columns, options, clean_sheet, mode):
    """Commit each collected row before advancing; expose only lightweight events to UI."""
    try:
        job_id, created = create_checkpoint(kind, inputs, columns, options, clean_sheet, mode)
    except Exception as exc:
        yield {"type": "error", "message": f"중간저장 파일을 만들 수 없어 크롤링을 시작하지 않았습니다: {exc}"}
        return

    pending: dict[int, dict] = {}
    changes: dict = {}
    terminal = False
    warned = False
    saved_count = 0

    def flush():
        nonlocal warned, saved_count
        try:
            update_checkpoint(job_id, pending, changes)
        except Exception as exc:
            if not warned:
                warned = True
                return False, f"⚠️ 중간저장 실패: {exc}. 다음 수집 시 다시 저장합니다. 완료 결과는 기존 방식으로 다운로드할 수 있습니다."
            return False, None
        if pending:
            saved_count = max(saved_count, max(pending))
        pending.clear()
        changes.clear()
        message = "✅ 중간저장 연결 복구" if warned else None
        warned = False
        return True, message

    try:
        yield {"type": "checkpoint_start", "job_id": job_id, "created_at": created, "mode": mode, "total": len(inputs)}
        for event in events:
            event_type = event.get("type")
            if event_type == "checkpoint_rows":
                for seq, row in enumerate(event["rows"], event["offset"] + 1):
                    pending[seq] = row
                if event.get("columns"):
                    changes["columns"] = event["columns"]
            elif event_type == "progress":
                changes["processed"] = event["current"]
            elif event_type == "complete":
                # Reconcile with the final in-memory result, even if disk writes failed earlier.
                pending.update(enumerate(event.get("results", []), 1))
                changes.update(state="complete", processed=len(inputs), columns=event.get("columns") or columns)
                terminal = True
            elif event_type == "error":
                changes["state"] = "error"
                terminal = True
            else:
                yield event
                continue

            ok, message = flush()
            if message:
                yield {"type": "log", "message": message}
            yield {"type": "checkpoint_status", "job_id": job_id, "saved_count": saved_count, "ok": ok}
            if event_type != "checkpoint_rows":
                yield event
    finally:
        # Browser refresh, disconnect or generator close: committed rows remain on disk.
        if not terminal:
            changes["state"] = "interrupted"
        if changes or pending:
            flush()
        close = getattr(events, "close", None)
        if close:
            close()
