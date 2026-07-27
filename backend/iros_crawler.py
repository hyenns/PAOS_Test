from __future__ import annotations

import re
import time
from typing import Iterable, Iterator, Any

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

IROS_URL = "https://www.iros.go.kr"

IROS_RESULT_COLUMNS = [
    "검색값",
    "법인종류",
    "상호(명칭)",
    "법인등록번호",
    "본점소재지",
    "등기상태",
    "상호말소상태",
    "주말 여부",
]


def make_event(event_type: str, **payload: Any) -> dict[str, Any]:
    return {"type": event_type, **payload}


def get_driver(*, headless: bool = True, page_load_strategy: str | None = None) -> webdriver.Chrome:
    options = Options()

    if page_load_strategy:
        options.page_load_strategy = page_load_strategy

    if headless:
        # Chrome 창을 띄우지 않고 백그라운드에서 실행합니다.
        # 최신 Chrome 기준으로 --headless=new가 일반 headless보다 화면 렌더링 호환성이 좋습니다.
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
    else:
        options.add_argument("--start-maximized")

    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--remote-allow-origins=*")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


def set_iros_option(driver: webdriver.Chrome, wait: WebDriverWait, checkbox_id: str, enabled: bool) -> None:
    """인터넷등기소 체크박스를 원하는 상태로 맞춥니다."""
    checkbox = wait.until(EC.presence_of_element_located((By.ID, checkbox_id)))

    def click_checkbox() -> None:
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
        time.sleep(0.3)
        try:
            label = driver.find_element(By.CSS_SELECTOR, f"label[for='{checkbox_id}']")
            driver.execute_script("arguments[0].click();", label)
        except Exception:
            driver.execute_script("arguments[0].click();", checkbox)
        time.sleep(0.3)

    if checkbox.is_selected() != enabled:
        click_checkbox()

    # WebSquare 화면에서 클릭 이벤트가 정상 반영되지 않는 경우 대비
    if checkbox.is_selected() != enabled:
        driver.execute_script(
            """
            arguments[0].checked = arguments[1];
            if (arguments[1]) {
                arguments[0].setAttribute('checked', 'checked');
            } else {
                arguments[0].removeAttribute('checked');
            }
            arguments[0].dispatchEvent(new Event('click', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """,
            checkbox,
            enabled,
        )
        time.sleep(0.3)


def enter_iros_business_search(
    driver: webdriver.Chrome,
    wait: WebDriverWait,
    *,
    include_closed_records: bool = False,
    include_erased_names: bool = False,
) -> None:
    """기존 Streamlit/bat에서 정상 동작하던 방식과 동일한 상호검색 진입 로직."""
    driver.get(IROS_URL)

    menu_element = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '등기지원 서비스')]"))
    )
    ActionChains(driver).move_to_element(menu_element).perform()
    driver.execute_script("arguments[0].click();", menu_element)
    time.sleep(2)

    business_search = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), '상호검색')]"))
    )
    driver.execute_script("arguments[0].click();", business_search)
    time.sleep(2)

    Select(
        wait.until(
            EC.presence_of_element_located(
                (By.ID, "mf_wfm_potal_main_wfm_content_sel_conm_name_juris_regt")
            )
        )
    ).select_by_visible_text("전체등기소")

    # 검색 옵션: 사용자가 선택한 경우에만 체크합니다.
    set_iros_option(
        driver,
        wait,
        "mf_wfm_potal_main_wfm_content_cbx_conm_name_close_rgs_rec_incld_yn_input_0",
        include_closed_records,
    )
    set_iros_option(
        driver,
        wait,
        "mf_wfm_potal_main_wfm_content_cbx_conm_name_eras_conm_incld_yn_input_0",
        include_erased_names,
    )


def normalize_iros_header(value: str) -> str:
    """표 머리글 비교를 위해 공백과 줄바꿈을 제거합니다."""
    return re.sub(r"\s+", "", str(value or "")).strip()


def get_iros_header_map(row_element) -> dict[str, int]:
    """현재 검색 결과 표의 머리글명과 실제 td 순번을 연결합니다."""
    header_map: dict[str, int] = {}
    try:
        table = row_element.find_element(By.XPATH, "./ancestor::table[1]")
        headers = table.find_elements(By.XPATH, ".//thead//th")
        for index, header in enumerate(headers, start=1):
            header_text = normalize_iros_header(header.get_attribute("textContent") or header.text)
            if header_text:
                header_map[header_text] = index
    except Exception:
        pass
    return header_map


def get_iros_cell_text(row_element, header_map: dict[str, int], header_name: str, fallback_index: int) -> str:
    """머리글 기준으로 셀을 찾고, 실패하면 기존 고정 순번을 사용합니다."""
    try:
        cells = row_element.find_elements(By.XPATH, "./td")
        td_index = header_map.get(normalize_iros_header(header_name), fallback_index)
        if td_index < 1 or len(cells) < td_index:
            return ""

        cell = cells[td_index - 1]
        value = cell.text or cell.get_attribute("textContent") or ""
        return re.sub(r"\s+", " ", value).strip()
    except Exception:
        return ""


def run_iros_crawler_events(
    company_inputs: Iterable[str],
    *,
    include_closed_records: bool = False,
    include_erased_names: bool = False,
    headless: bool = True,
) -> Iterator[dict[str, Any]]:
    """등기소 크롤링을 실행하면서 진행 로그를 실시간 이벤트로 반환합니다."""
    company_inputs = [str(v).strip() for v in company_inputs if str(v).strip()]
    results: list[dict] = []
    logs: list[str] = []
    total = len(company_inputs)

    def log(message: str) -> dict[str, Any]:
        logs.append(message)
        return make_event("log", message=message, logs=logs.copy())

    driver = None
    try:
        yield log("🌐 인터넷등기소 접속 중...")

        yield log(
            "검색 옵션: "
            f"폐쇄등기기록 포함={'ON' if include_closed_records else 'OFF'}, "
            f"주말된 상호(명칭) 포함={'ON' if include_erased_names else 'OFF'}"
        )

        driver = get_driver(headless=headless)
        wait = WebDriverWait(driver, 10)

        enter_iros_business_search(
            driver,
            wait,
            include_closed_records=include_closed_records,
            include_erased_names=include_erased_names,
        )
        yield log("✅ 인터넷등기소 상호검색 화면 진입 완료")

        search_box = wait.until(
            EC.presence_of_element_located(
                (By.ID, "mf_wfm_potal_main_wfm_content_sbx_conm_name_swrd___input")
            )
        )

        for idx, company in enumerate(company_inputs):
            current_no = idx + 1
            try:
                yield log(f"🔎 [{current_no}/{total}] {company} 검색 중...")

                search_box.clear()
                search_box.send_keys(company)

                search_button = wait.until(
                    EC.element_to_be_clickable(
                        (By.ID, "mf_wfm_potal_main_wfm_content_btn_conm_name_search")
                    )
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
                time.sleep(1)
                driver.execute_script("arguments[0].click();", search_button)
                time.sleep(2)

                # 팝업 처리
                try:
                    popup_button = WebDriverWait(driver, 1).until(
                        EC.presence_of_element_located(
                            (
                                By.XPATH,
                                "//a[contains(text(), '확인') and contains(@class, 'btn solid medium color-main')]",
                            )
                        )
                    )
                    driver.execute_script("arguments[0].click();", popup_button)
                    time.sleep(1)
                except Exception:
                    pass

                rows = driver.find_elements(
                    By.XPATH,
                    "//td[@data-col_id='no']/span[text()='*']/parent::td/parent::tr",
                )

                if not rows:
                    yield log(f"⏭️ [{current_no}/{total}] {company} — 유효 데이터 없음")
                    yield make_event("progress", current=current_no, total=total, result_count=len(results))
                    continue

                count = 0
                for row in rows:
                    try:
                        # 머리글명으로 열 위치를 찾고, 실패하면 현재 화면 기준 순번을 사용합니다.
                        header_map = get_iros_header_map(row)

                        corp_type = get_iros_cell_text(row, header_map, "법인종류", 3)
                        company_name = get_iros_cell_text(row, header_map, "상호(명칭)", 4)
                        corp_reg_num = get_iros_cell_text(row, header_map, "법인등록번호", 5)
                        company_address = get_iros_cell_text(row, header_map, "본점소재지", 6)
                        registration_status = get_iros_cell_text(row, header_map, "등기상태", 11)
                        name_cancellation_status = get_iros_cell_text(row, header_map, "상호말소상태", 12)
                        weekend_status = get_iros_cell_text(row, header_map, "주말여부", 13)

                        if not all([corp_type, company_name, corp_reg_num, company_address]):
                            continue

                        results.append(
                            {
                                "검색값": company,
                                "법인종류": corp_type,
                                "상호(명칭)": company_name,
                                "법인등록번호": corp_reg_num,
                                "본점소재지": company_address,
                                "등기상태": registration_status,
                                "상호말소상태": name_cancellation_status,
                                "주말 여부": weekend_status,
                            }
                        )
                        count += 1
                    except Exception:
                        continue

                yield log(f"✅ [{current_no}/{total}] {company} — {count}건")
                yield make_event("progress", current=current_no, total=total, result_count=len(results))

            except TimeoutException:
                yield log(f"❌ [{current_no}/{total}] {company} — 조회 시간 초과")
                yield make_event("progress", current=current_no, total=total, result_count=len(results))
            except Exception as e:
                yield log(f"❌ [{current_no}/{total}] {company} — 조회 실패: {str(e)[:80]}")
                yield make_event("progress", current=current_no, total=total, result_count=len(results))

        if results:
            status_has_value = any(
                str(row.get(col, "")).strip()
                for row in results
                for col in ["등기상태", "상호말소상태", "주말 여부"]
            )
            if not status_has_value:
                yield log("⚠️ 상태값 열은 생성됐지만 값이 비어 있습니다. 인터넷등기소 결과표의 상태 열 위치가 변경됐을 수 있습니다.")

        yield log("🎉 크롤링 완료")
        yield make_event("complete", results=results, logs=logs.copy())

    except Exception as e:
        yield make_event("error", message=f"등기소 크롤링 실패: {str(e)}", logs=logs.copy())
    finally:
        if driver:
            driver.quit()


def run_iros_crawler(
    company_inputs: Iterable[str],
    *,
    include_closed_records: bool = False,
    include_erased_names: bool = False,
    headless: bool = True,
) -> tuple[list[dict], list[str]]:
    """기존 API 호환용 함수. 내부적으로 실시간 이벤트 함수를 실행하고 최종 결과만 반환합니다."""
    final_results: list[dict] = []
    final_logs: list[str] = []
    for event in run_iros_crawler_events(
        company_inputs,
        include_closed_records=include_closed_records,
        include_erased_names=include_erased_names,
        headless=headless,
    ):
        if event.get("type") in {"log", "error", "complete"}:
            final_logs = event.get("logs", final_logs)
        if event.get("type") == "complete":
            final_results = event.get("results", [])
        if event.get("type") == "error":
            raise RuntimeError(event.get("message", "등기소 크롤링 실패"))
    return final_results, final_logs
