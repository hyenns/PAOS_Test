from __future__ import annotations

import base64
import io
import json
import os
from pathlib import Path

import pandas as pd
from flask import Flask, Response, jsonify, request, send_from_directory, stream_with_context

from backend.iros_crawler import IROS_RESULT_COLUMNS, run_iros_crawler_events
from backend.saramin_crawler import SARAMIN_RESULT_COLUMNS, get_saramin_result_columns, run_saramin_crawler_events

BASE_DIR = Path(__file__).resolve().parent

app = Flask(__name__)
# JSON 응답의 키 순서가 자동 정렬되면 화면 미리보기 열 순서가 엑셀과 달라질 수 있어 비활성화합니다.
app.json.sort_keys = False
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30MB


def excel_column_letter(index: int) -> str:
    index += 1
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def read_search_values(file_bytes: bytes, sheet_name: str, header_mode: str, column_index: int) -> list[str]:
    header = 0 if header_mode == "header" else None
    df = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=sheet_name,
        header=header,
        dtype=str,
    )
    df = df.dropna(how="all").reset_index(drop=True)

    if df.empty or len(df.columns) == 0:
        return []

    if column_index < 0 or column_index >= len(df.columns):
        raise ValueError("선택한 열 번호가 엑셀 범위를 벗어났습니다.")

    values: list[str] = []
    for value in df.iloc[:, column_index].tolist():
        if pd.isna(value):
            continue
        cleaned = str(value).strip()
        if not cleaned or cleaned.lower() == "nan":
            continue
        values.append(cleaned)
    return values


def apply_result_excel_style(worksheet) -> None:
    """크롤링 결과 엑셀 파일의 가독성을 높이는 공통 서식입니다."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    max_row = worksheet.max_row or 1
    max_column = worksheet.max_column or 1

    # 첫 행 고정 + 필터
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    # 요청 반영: 헤더는 초록색이 아닌 회색으로 표시
    header_fill = PatternFill(fill_type="solid", fgColor="D9D9D9")
    header_font = Font(color="000000", bold=True)
    thin_border = Border(
        left=Side(style="thin", color="D9E2EC"),
        right=Side(style="thin", color="D9E2EC"),
        top=Side(style="thin", color="D9E2EC"),
        bottom=Side(style="thin", color="D9E2EC"),
    )
    body_alignment = Alignment(vertical="center", wrap_text=False)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    worksheet.row_dimensions[1].height = 22

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border

    for row in worksheet.iter_rows(min_row=2, max_row=max_row, max_col=max_column):
        for cell in row:
            cell.alignment = body_alignment
            cell.border = thin_border

    # 내용 길이에 맞춰 열 너비 조정
    for column_index in range(1, max_column + 1):
        column_letter = get_column_letter(column_index)
        header_value = worksheet.cell(row=1, column=column_index).value
        header_text = str(header_value or "")
        max_length = len(header_text)

        for cell in worksheet[column_letter]:
            value = cell.value
            if value is None:
                continue
            value_text = str(value)
            # 빈 문자열인 셀이 있어도 오류가 나지 않도록 안전하게 계산합니다.
            parts = value_text.splitlines() or [value_text]
            max_length = max(max_length, max(len(part) for part in parts))

        width = max_length + 3
        if any(keyword in header_text for keyword in ["주소", "소재지", "사업내용"]):
            width = min(max(width, 18), 55)
        elif "URL" in header_text or "홈페이지" in header_text:
            width = min(max(width, 18), 50)
        elif any(keyword in header_text for keyword in ["법인등록번호", "사업자등록번호"]):
            width = min(max(width, 18), 24)
        else:
            width = min(max(width, 10), 28)

        worksheet.column_dimensions[column_letter].width = width


def dataframe_to_excel_base64(rows: list[dict], columns: list[str] | None = None) -> str:
    df = pd.DataFrame(rows, columns=columns) if columns else pd.DataFrame(rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="결과")
        worksheet = writer.sheets["결과"]
        apply_result_excel_style(worksheet)
    output.seek(0)
    return base64.b64encode(output.read()).decode("utf-8")


@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/<path:path>")
def static_files(path: str):
    target = BASE_DIR / path
    if target.exists() and target.is_file():
        return send_from_directory(BASE_DIR, path)
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/api/health")
def health():
    return jsonify({"ok": True, "app": "WorkLab PAOS", "mode": "team-test"})


@app.route("/api/iros/run", methods=["POST"])
def api_iros_run():
    uploaded = request.files.get("file")
    if not uploaded:
        return jsonify({"ok": False, "message": "엑셀 파일이 업로드되지 않았습니다."}), 400

    sheet_name = request.form.get("sheet_name", "")
    header_mode = request.form.get("header_mode", "header")

    try:
        column_index = int(request.form.get("column_index", "0"))
    except ValueError:
        return jsonify({"ok": False, "message": "열 번호가 올바르지 않습니다."}), 400

    try:
        file_bytes = uploaded.read()
        search_values = read_search_values(file_bytes, sheet_name, header_mode, column_index)
    except Exception as e:
        return jsonify({"ok": False, "message": f"엑셀 파일을 읽을 수 없습니다: {str(e)}"}), 400

    if not search_values:
        return jsonify({"ok": False, "message": "선택한 열에 크롤링할 회사명이 없습니다."}), 400

    include_closed_records = request.form.get("include_closed_records", "false").lower() == "true"
    include_erased_names = request.form.get("include_erased_names", "false").lower() == "true"

    # 기본값은 Chrome 창을 띄우지 않는 백그라운드 모드입니다.
    # 추후 문제가 생기면 프론트에서 headless=false를 넘기거나 환경변수/코드로 조정할 수 있습니다.
    headless = request.form.get("headless", "true").lower() != "false"

    def ndjson(event: dict) -> str:
        return json.dumps(event, ensure_ascii=False) + "\n"

    @stream_with_context
    def generate():
        for event in run_iros_crawler_events(
            search_values,
            include_closed_records=include_closed_records,
            include_erased_names=include_erased_names,
            headless=headless,
        ):
            if event.get("type") == "complete":
                results = event.get("results", [])
                excel_base64 = dataframe_to_excel_base64(results, columns=IROS_RESULT_COLUMNS)

                status_columns = ["등기상태", "상호말소상태", "주말 여부"]
                status_has_value = any(
                    str(row.get(col, "")).strip()
                    for row in results
                    for col in status_columns
                )

                event.update({
                    "ok": True,
                    "total_input": len(search_values),
                    "total_result": len(results),
                    "status_has_value": status_has_value,
                    "columns": IROS_RESULT_COLUMNS,
                    "excel_base64": excel_base64,
                    "filename": "등기소_크롤링_결과.xlsx",
                })

            yield ndjson(event)

    return Response(
        generate(),
        mimetype="application/x-ndjson; charset=utf-8",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


@app.route("/api/saramin/run", methods=["POST"])
def api_saramin_run():
    uploaded = request.files.get("file")
    if not uploaded:
        return jsonify({"ok": False, "message": "엑셀 파일이 업로드되지 않았습니다."}), 400

    sheet_name = request.form.get("sheet_name", "")
    header_mode = request.form.get("header_mode", "header")

    try:
        column_index = int(request.form.get("column_index", "0"))
    except ValueError:
        return jsonify({"ok": False, "message": "열 번호가 올바르지 않습니다."}), 400

    try:
        max_results_per_keyword = int(request.form.get("max_results_per_keyword", "5"))
        max_results_per_keyword = max(1, min(max_results_per_keyword, 5))
    except ValueError:
        max_results_per_keyword = 5

    try:
        file_bytes = uploaded.read()
        search_values = read_search_values(file_bytes, sheet_name, header_mode, column_index)
    except Exception as e:
        return jsonify({"ok": False, "message": f"엑셀 파일을 읽을 수 없습니다: {str(e)}"}), 400

    if not search_values:
        return jsonify({"ok": False, "message": "선택한 열에 크롤링할 회사명이 없습니다."}), 400

    headless = request.form.get("headless", "true").lower() != "false"
    collect_finance = request.form.get("collect_finance", "false").lower() == "true"
    result_columns = get_saramin_result_columns(collect_finance)

    def ndjson(event: dict) -> str:
        return json.dumps(event, ensure_ascii=False) + "\n"

    @stream_with_context
    def generate():
        for event in run_saramin_crawler_events(
            search_values,
            max_results_per_keyword=max_results_per_keyword,
            collect_finance=collect_finance,
            headless=headless,
        ):
            if event.get("type") == "complete":
                results = event.get("results", [])
                excel_base64 = dataframe_to_excel_base64(results, columns=result_columns)

                event.update({
                    "ok": True,
                    "total_input": len(search_values),
                    "total_result": len(results),
                    "columns": result_columns,
                    "collect_finance": collect_finance,
                    "excel_base64": excel_base64,
                    "filename": "사람인_기업정보_크롤링_결과.xlsx",
                })

            yield ndjson(event)

    return Response(
        generate(),
        mimetype="application/x-ndjson; charset=utf-8",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )

if __name__ == "__main__":
    host = os.environ.get("WORKLAB_HOST", "0.0.0.0")
    port = int(os.environ.get("WORKLAB_PORT", "8000"))
    app.run(host=host, port=port, debug=False, threaded=True)
