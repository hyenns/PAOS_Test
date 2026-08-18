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
from selenium.webdriver.common.keys import Keys
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



def _visible_exact_text_elements(driver: webdriver.Chrome, text: str):
    xpath = (
        "//*[self::a or self::button or self::span or self::div]"
        f"[normalize-space(.)='{text}']"
    )
    elements = []
    for element in driver.find_elements(By.XPATH, xpath):
        try:
            if element.is_displayed():
                elements.append(element)
        except Exception:
            continue
    return elements


def _click_element(driver: webdriver.Chrome, element) -> None:
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    except Exception:
        pass
    try:
        driver.execute_script("arguments[0].click();", element)
    except Exception:
        element.click()


def _hover_element(driver: webdriver.Chrome, element) -> None:
    """메가메뉴가 hover 이벤트로 열리는 경우를 위해 mouseover/mouseenter를 함께 발생시킵니다."""
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    except Exception:
        pass

    try:
        ActionChains(driver).move_to_element(element).pause(0.3).perform()
    except Exception:
        pass

    try:
        driver.execute_script(
            """
            ['mouseover', 'mouseenter'].forEach(function(type) {
                arguments[0].dispatchEvent(new MouseEvent(type, {bubbles: true, view: window}));
            });
            """,
            element,
        )
    except Exception:
        pass


def _wait_visible_exact_text(driver: webdriver.Chrome, text: str, timeout: int = 12):
    """화면에 실제로 보이는 동일 텍스트 요소가 나타날 때까지 기다립니다."""
    try:
        return WebDriverWait(driver, timeout).until(
            lambda d: (_visible_exact_text_elements(d, text) or [False])[0]
        )
    except TimeoutException as exc:
        raise TimeoutException(f"'{text}' 요소를 찾지 못했습니다.") from exc

def _wait_top_navigation_text(driver: webdriver.Chrome, text: str, timeout: int = 15):
    """같은 문구가 여러 곳에 있을 때 화면 상단의 네비게이션 요소를 우선합니다."""
    def locate(d):
        candidates = _visible_exact_text_elements(d, text)
        if not candidates:
            return False
        ranked = []
        for element in candidates:
            try:
                y = element.location.get("y", 99999)
                x = element.location.get("x", 99999)
                ranked.append((y, x, element))
            except Exception:
                continue
        if not ranked:
            return False
        ranked.sort(key=lambda item: (item[0], item[1]))
        return ranked[0][2]

    try:
        return WebDriverWait(driver, timeout).until(locate)
    except TimeoutException as exc:
        raise TimeoutException(f"상단 '{text}' 메뉴를 찾지 못했습니다.") from exc


def _find_corporation_issue_link(driver: webdriver.Chrome, corp_heading, top_menu):
    """메가메뉴의 '법인' 열 안에 있는 '열람·발급' 링크를 우선해서 찾습니다."""
    # 1순위: '법인' 제목의 가장 가까운 부모 영역 안에서 열람·발급 링크 탐색
    try:
        container = corp_heading.find_element(
            By.XPATH,
            "ancestor::*[self::li or self::div or self::section][.//*[self::a or self::button][normalize-space(.)='열람·발급']][1]",
        )
        links = container.find_elements(
            By.XPATH,
            ".//*[self::a or self::button][normalize-space(.)='열람·발급']",
        )
        for link in links:
            try:
                if link.is_displayed() and link.is_enabled() and link != top_menu:
                    return link
            except Exception:
                continue
    except Exception:
        pass

    # 2순위: 법인 제목과 화면 좌표가 가장 가까운 열람·발급 요소
    submenu_candidates = _visible_exact_text_elements(driver, "열람·발급")
    try:
        hx = corp_heading.location.get("x", 0)
        hy = corp_heading.location.get("y", 0)
        ranked = []
        for candidate in submenu_candidates:
            if candidate == top_menu:
                continue
            try:
                cx = candidate.location.get("x", 0)
                cy = candidate.location.get("y", 0)
                if cy >= hy - 15:
                    distance = abs(cx - hx) + abs(cy - hy) * 0.35
                    ranked.append((distance, candidate))
            except Exception:
                continue
        if ranked:
            ranked.sort(key=lambda item: item[0])
            return ranked[0][1]
    except Exception:
        pass

    return None


def enter_iros_registration_search(driver: webdriver.Chrome, wait: WebDriverWait) -> None:
    """열람·발급 > 법인 > 열람·발급 > 등록번호검색 화면으로 진입합니다.

    상단 메뉴는 클릭이 아니라 hover로 메가메뉴가 열리는 구조라 먼저 hover를 발생시키고,
    그 안의 법인 열에 있는 '열람·발급' 링크를 찾아 클릭합니다.
    """
    driver.get(IROS_URL)

    # 1) 상단 '열람·발급' 메뉴: 클릭하면 메가메뉴가 닫히거나 다른 화면으로 이동할 수 있으므로 hover 우선
    try:
        top_menu = _wait_top_navigation_text(driver, "열람·발급", timeout=15)
        _hover_element(driver, top_menu)

        # 메가메뉴가 실제로 펼쳐질 때까지 '법인' 제목을 기다립니다.
        try:
            corp_heading = WebDriverWait(driver, 5).until(
                lambda d: (_visible_exact_text_elements(d, "법인") or [False])[0]
            )
        except TimeoutException:
            # hover 이벤트가 브라우저 환경에 따라 먹지 않을 때 한 번만 클릭 fallback
            _click_element(driver, top_menu)
            time.sleep(0.8)
            corp_heading = _wait_visible_exact_text(driver, "법인", timeout=6)
    except TimeoutException as exc:
        raise TimeoutException("상단 '열람·발급' 메뉴 또는 법인 메뉴를 열지 못했습니다.") from exc

    # 2) 펼쳐진 메가메뉴의 '법인' 열에서 '열람·발급' 선택
    corp_issue_link = _find_corporation_issue_link(driver, corp_heading, top_menu)
    if corp_issue_link is None:
        raise TimeoutException("법인 메뉴의 '열람·발급' 항목을 찾지 못했습니다.")

    try:
        _click_element(driver, corp_issue_link)
    except Exception as exc:
        raise RuntimeError("법인 '열람·발급' 메뉴 클릭에 실패했습니다.") from exc

    # 신청 화면 로딩 대기
    try:
        WebDriverWait(driver, 15).until(
            lambda d: "법인 등기사항증명서" in d.find_element(By.TAG_NAME, "body").text
        )
    except TimeoutException as exc:
        raise TimeoutException("법인 등기사항증명서 열람·발급 신청 화면으로 이동하지 못했습니다.") from exc

    # 3) 등록번호검색 탭
    try:
        tab = _wait_visible_exact_text(driver, "등록번호검색", timeout=10)
        _click_element(driver, tab)
        find_registration_number_input(driver, WebDriverWait(driver, 8))
    except TimeoutException as exc:
        raise TimeoutException("등록번호검색 탭 또는 등록번호 입력창을 찾지 못했습니다.") from exc
    except Exception:
        # 입력창 확인 과정에서 일시적인 WebDriver 예외가 나도 탭 클릭까지 성공했다면 이후 검색 단계에서 재탐색합니다.
        pass

    time.sleep(0.5)

def normalize_registration_number(value: str) -> str:
    """법인등록번호를 인터넷등기소 입력 형식(6자리-7자리)으로 맞춥니다."""
    text = str(value or "").strip()
    digits = re.sub(r"\D", "", text)
    if len(digits) == 13:
        return f"{digits[:6]}-{digits[6:]}"
    return text


def find_registration_number_input(driver: webdriver.Chrome, wait: WebDriverWait):
    """등록번호검색 화면의 등록번호 입력창을 찾습니다."""
    def locate(d):
        # label의 for 속성이 연결되어 있는 경우
        labels = d.find_elements(
            By.XPATH,
            "//*[self::label or self::span or self::div][contains(normalize-space(.), '등록번호')]",
        )
        for label in labels:
            try:
                if not label.is_displayed():
                    continue
                target_id = label.get_attribute("for")
                if target_id:
                    target = d.find_element(By.ID, target_id)
                    if target.is_displayed() and target.tag_name.lower() == "input":
                        return target
                nearby = label.find_elements(By.XPATH, "following::input[1]")
                if nearby and nearby[0].is_displayed():
                    return nearby[0]
            except Exception:
                continue

        # 화면 안내문상 하이픈 포함 14자리 입력이므로 maxlength=14인 입력창을 우선합니다.
        for element in d.find_elements(By.CSS_SELECTOR, "input[maxlength='14']"):
            try:
                if element.is_displayed() and element.is_enabled():
                    return element
            except Exception:
                continue

        # 마지막 fallback: 현재 화면의 보이는 텍스트 입력 중 폭이 가장 큰 입력창
        candidates = []
        for element in d.find_elements(By.CSS_SELECTOR, "input[type='text'], input:not([type])"):
            try:
                if element.is_displayed() and element.is_enabled():
                    candidates.append((element.size.get("width", 0), element))
            except Exception:
                continue
        if candidates:
            candidates.sort(key=lambda item: item[0], reverse=True)
            return candidates[0][1]
        return False

    return wait.until(locate)


def find_nearest_search_button(driver: webdriver.Chrome, input_element, wait: WebDriverWait):
    """등록번호 입력창과 가장 가까운 '검색' 버튼을 찾습니다."""
    def locate(d):
        candidates = []
        xpath = "//*[self::button or self::a or @role='button'][normalize-space(.)='검색']"
        try:
            ix = input_element.location.get("x", 0)
            iy = input_element.location.get("y", 0)
        except Exception:
            ix, iy = 0, 0

        for element in d.find_elements(By.XPATH, xpath):
            try:
                if not element.is_displayed() or not element.is_enabled():
                    continue
                ex = element.location.get("x", 0)
                ey = element.location.get("y", 0)
                distance = abs(ex - ix) + abs(ey - iy) * 2
                candidates.append((distance, element))
            except Exception:
                continue
        if candidates:
            candidates.sort(key=lambda item: item[0])
            return candidates[0][1]
        return False

    return wait.until(locate)


def set_registration_number_input(driver: webdriver.Chrome, input_element, value: str) -> str:
    """등록번호 입력값을 WebSquare 내부 상태까지 확실히 갱신합니다.

    일반 Selenium clear()/send_keys()만 사용하면 첫 검색값이 화면/내부 모델에 남아
    다음 검색에서도 같은 번호가 조회되는 경우가 있어 input/change 이벤트를 함께 발생시킵니다.
    최종적으로 입력창의 실제 값을 다시 읽어 목표 번호가 들어갔는지 확인합니다.
    """
    target = normalize_registration_number(value)
    target_digits = re.sub(r"\D", "", target)

    for _ in range(3):
        # 사람 입력과 비슷하게 전체 선택 후 삭제
        try:
            input_element.click()
            input_element.send_keys(Keys.CONTROL, "a")
            input_element.send_keys(Keys.BACKSPACE)
        except Exception:
            pass

        # WebSquare/SPA 바인딩이 값을 인식하도록 native value setter + 이벤트 발생
        try:
            driver.execute_script(
                """
                const el = arguments[0];
                const value = arguments[1];
                el.focus();
                const proto = Object.getPrototypeOf(el);
                const desc = Object.getOwnPropertyDescriptor(proto, 'value')
                    || Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
                if (desc && desc.set) {
                    desc.set.call(el, value);
                } else {
                    el.value = value;
                }
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: '0' }));
                el.blur();
                """,
                input_element,
                target,
            )
        except Exception:
            try:
                input_element.clear()
                input_element.send_keys(target)
            except Exception:
                pass

        time.sleep(0.25)
        try:
            actual = input_element.get_attribute("value") or ""
        except Exception:
            actual = ""

        if re.sub(r"\D", "", actual) == target_digits:
            return actual

        # DOM이 다시 렌더링된 경우 입력창을 재탐색해서 한 번 더 시도
        try:
            input_element = find_registration_number_input(driver, WebDriverWait(driver, 3))
        except Exception:
            pass

    raise RuntimeError(f"등록번호 입력값 반영 실패: {target}")


def registration_result_signature(driver: webdriver.Chrome) -> str:
    """현재 화면에 표시된 등록번호 검색 결과표의 내용 서명값을 반환합니다."""
    try:
        rows, columns = extract_registration_search_results(driver, "")
    except Exception:
        return ""

    if not rows:
        return ""

    usable_columns = [column for column in columns if column != "검색값"]
    lines: list[str] = []
    for row in rows:
        lines.append("|".join(str(row.get(column, "")).strip() for column in usable_columns))
    return "\n".join(lines)


def _dedupe_headers(headers: list[str], count: int) -> list[str]:
    output: list[str] = []
    seen: dict[str, int] = {}
    for idx in range(count):
        base = re.sub(r"\s+", " ", str(headers[idx] if idx < len(headers) else "")).strip()
        if not base:
            base = f"열{idx + 1}"
        seen[base] = seen.get(base, 0) + 1
        name = base if seen[base] == 1 else f"{base}_{seen[base]}"
        output.append(name)
    return output


def extract_registration_search_results(driver: webdriver.Chrome, original_value: str) -> tuple[list[dict], list[str]]:
    """등록번호 검색 후 '법인상호 선택' 결과표의 실제 데이터 행만 수집합니다.

    인터넷등기소 결과표는 반응형 화면용 머리글을 td 한 칸짜리 행으로도 렌더링합니다.
    기존 로직은 이 행들까지 데이터로 읽어 '관할등기소', '법인구분' 등이 결과 행으로
    저장되는 문제가 있었습니다. 실제 결과 행과 동일한 셀 수를 가진 행만 남기고,
    UI 선택용 '선택' 열은 결과 엑셀에서 제외합니다.
    """
    best = None
    keyword_tokens = ["상호", "법인", "등록번호", "본점", "등기", "관할", "주말", "폐쇄"]

    for table in driver.find_elements(By.TAG_NAME, "table"):
        try:
            if not table.is_displayed():
                continue

            tr_elements = table.find_elements(By.XPATH, ".//tr")
            row_candidates: list[list[str]] = []
            max_cells = 0

            for tr in tr_elements:
                tds = tr.find_elements(By.XPATH, "./td")
                if not tds:
                    continue

                values = [
                    re.sub(r"\s+", " ", (td.text or td.get_attribute("textContent") or "")).strip()
                    for td in tds
                ]
                if not any(values):
                    continue

                row_candidates.append(values)
                max_cells = max(max_cells, len(values))

            if not row_candidates or max_cells < 2:
                continue

            # 반응형 머리글은 보통 td가 1개인 별도 행으로 반복됩니다.
            # 실제 결과 행은 결과표의 최대 셀 수와 동일하므로 그 행만 남깁니다.
            data_rows = [values for values in row_candidates if len(values) == max_cells]
            if not data_rows:
                continue

            header_elements = table.find_elements(By.XPATH, ".//thead//th")
            if not header_elements:
                header_elements = table.find_elements(By.XPATH, ".//tr[1]/th")
            raw_headers = [
                re.sub(r"\s+", " ", (h.text or h.get_attribute("textContent") or "")).strip()
                for h in header_elements
            ]
            headers = _dedupe_headers(raw_headers, max_cells)

            # 일부 화면에서는 머리글 자체가 전체 폭 td 행으로 한 번 더 렌더링될 수 있어 제거합니다.
            normalized_headers = [normalize_iros_header(h) for h in headers]
            filtered_rows: list[list[str]] = []
            for values in data_rows:
                normalized_values = [normalize_iros_header(v) for v in values]
                if normalized_values == normalized_headers:
                    continue
                filtered_rows.append(values)
            data_rows = filtered_rows
            if not data_rows:
                continue

            table_text = re.sub(r"\s+", " ", table.text or "")
            score = sum(3 for token in keyword_tokens if token in " ".join(headers))
            score += sum(1 for token in keyword_tokens if token in table_text)
            score += min(len(data_rows), 5)

            candidate = (score, headers, data_rows)
            if best is None or candidate[0] > best[0]:
                best = candidate
        except Exception:
            continue

    if best is None:
        return [], []

    _, headers, data_rows = best

    # '선택'은 라디오 버튼용 UI 열이므로 수집 결과에서는 제외합니다.
    keep_indices = [
        idx for idx, header in enumerate(headers)
        if normalize_iros_header(header) not in {"선택"}
    ]
    clean_headers = [headers[idx] for idx in keep_indices]

    results: list[dict] = []
    for values in data_rows:
        padded = values + [""] * (len(headers) - len(values))
        clean_values = [padded[idx] for idx in keep_indices]
        row = {header: clean_values[idx] for idx, header in enumerate(clean_headers)}
        row = {"검색값": original_value, **row}
        results.append(row)

    return results, ["검색값", *clean_headers]

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
    search_mode: str = "company",
    include_closed_records: bool = False,
    include_erased_names: bool = False,
    headless: bool = True,
) -> Iterator[dict[str, Any]]:
    """등기소 크롤링을 실행하면서 진행 로그를 실시간 이벤트로 반환합니다.

    search_mode:
      - company: 기존 상호검색
      - registration: 열람·발급 > 법인 > 등록번호검색
    """
    company_inputs = [str(v).strip() for v in company_inputs if str(v).strip()]
    results: list[dict] = []
    logs: list[str] = []
    total = len(company_inputs)
    search_mode = "registration" if search_mode == "registration" else "company"
    result_columns: list[str] = IROS_RESULT_COLUMNS.copy()

    def log(message: str) -> dict[str, Any]:
        logs.append(message)
        return make_event("log", message=message, logs=logs.copy())

    def merge_columns(columns: list[str]) -> None:
        nonlocal result_columns
        if search_mode == "registration" and result_columns == IROS_RESULT_COLUMNS:
            result_columns = []
        for column in columns:
            if column not in result_columns:
                result_columns.append(column)

    driver = None
    try:
        yield log("🌐 인터넷등기소 접속 중...")
        yield log(f"검색 방식: {'등록번호검색' if search_mode == 'registration' else '상호검색'}")
        if search_mode == "registration":
            yield log(f"📋 법인등록번호 {total}건을 순차 검색합니다.")

        if search_mode == "company":
            yield log(
                "검색 옵션: "
                f"폐쇄등기기록 포함={'ON' if include_closed_records else 'OFF'}, "
                f"주말된 상호(명칭) 포함={'ON' if include_erased_names else 'OFF'}"
            )

        driver = get_driver(headless=headless)
        wait = WebDriverWait(driver, 12)

        if search_mode == "registration":
            enter_iros_registration_search(driver, wait)
            yield log("✅ 인터넷등기소 등록번호검색 화면 진입 완료")
        else:
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

                if search_mode == "registration":
                    search_value = normalize_registration_number(company)
                    per_rows: list[dict] = []
                    per_columns: list[str] = []
                    no_result = False

                    # 첫 검색 결과가 DOM에 남아 있는 상태에서 다음 번호를 검색하면,
                    # 기존 표를 새 결과로 오인하거나 WebSquare 내부 값이 첫 번호로 유지되는 문제가 있었습니다.
                    # 입력값을 이벤트까지 포함해 갱신하고, 결과표 내용이 실제로 바뀔 때까지 기다립니다.
                    for attempt in range(2):
                        previous_signature = registration_result_signature(driver)
                        registration_input = find_registration_number_input(driver, wait)
                        set_registration_number_input(driver, registration_input, search_value)

                        # 값 설정 과정에서 DOM이 다시 그려질 수 있어 버튼 탐색 직전에 입력창을 다시 잡습니다.
                        registration_input = find_registration_number_input(driver, wait)
                        search_button = find_nearest_search_button(driver, registration_input, wait)
                        _click_element(driver, search_button)

                        # 이전 결과표가 잠깐 남아 있어도 바로 수집하지 않고 실제 갱신을 기다립니다.
                        time.sleep(0.7)
                        deadline = time.time() + 8.0
                        while time.time() < deadline:
                            body_text = driver.find_element(By.TAG_NAME, "body").text
                            current_signature = registration_result_signature(driver)

                            # 첫 검색이거나, 기존 표와 다른 내용으로 갱신된 경우에만 결과를 확정합니다.
                            if current_signature and (not previous_signature or current_signature != previous_signature):
                                per_rows, per_columns = extract_registration_search_results(driver, company)
                                if per_rows:
                                    break

                            if any(text in body_text for text in ["검색결과가 없습니다", "조회된 결과가 없습니다", "검색 결과가 없습니다"]):
                                no_result = True
                                break

                            time.sleep(0.35)

                        if per_rows or no_result:
                            break

                        if attempt == 0:
                            yield log(f"↻ [{current_no}/{total}] 결과 갱신이 확인되지 않아 검색 화면을 새로 열고 재시도합니다.")
                            enter_iros_registration_search(driver, wait)
                            time.sleep(0.4)

                    if not per_rows:
                        status_text = "검색결과 없음" if no_result else "검색 결과 갱신 실패"
                        yield log(f"⏭️ [{current_no}/{total}] {company} — {status_text}")
                        yield make_event("progress", current=current_no, total=total, result_count=len(results))
                        continue

                    merge_columns(per_columns)
                    results.extend(per_rows)
                    yield log(f"✅ [{current_no}/{total}] {company} — {len(per_rows)}건")
                    yield make_event("progress", current=current_no, total=total, result_count=len(results))
                    continue

                # 기존 상호검색 로직
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
                yield log(f"❌ [{current_no}/{total}] {company} — 조회 실패: {str(e)[:100]}")
                yield make_event("progress", current=current_no, total=total, result_count=len(results))

        if search_mode == "company" and results:
            status_has_value = any(
                str(row.get(col, "")).strip()
                for row in results
                for col in ["등기상태", "상호말소상태", "주말 여부"]
            )
            if not status_has_value:
                yield log("⚠️ 상태값 열은 생성됐지만 값이 비어 있습니다. 인터넷등기소 결과표의 상태 열 위치가 변경됐을 수 있습니다.")

        if search_mode == "registration" and not result_columns:
            result_columns = ["검색값"]

        yield log("🎉 크롤링 완료")
        yield make_event(
            "complete",
            results=results,
            columns=result_columns,
            search_mode=search_mode,
            logs=logs.copy(),
        )

    except Exception as e:
        yield make_event("error", message=f"등기소 크롤링 실패: {str(e)}", logs=logs.copy())
    finally:
        if driver:
            driver.quit()


def run_iros_crawler(
    company_inputs: Iterable[str],
    *,
    search_mode: str = "company",
    include_closed_records: bool = False,
    include_erased_names: bool = False,
    headless: bool = True,
) -> tuple[list[dict], list[str]]:
    """기존 API 호환용 함수. 내부적으로 실시간 이벤트 함수를 실행하고 최종 결과만 반환합니다."""
    final_results: list[dict] = []
    final_logs: list[str] = []
    for event in run_iros_crawler_events(
        company_inputs,
        search_mode=search_mode,
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
