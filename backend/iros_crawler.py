from __future__ import annotations

import os
import re
import time
import threading
import json
import shutil
from pathlib import Path
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


class IrosLoginRequiredError(RuntimeError):
    """인터넷등기소 로그인 세션이 없거나 만료된 경우 사용합니다."""


_IROS_LOGIN_DRIVER = None
_IROS_LOGIN_LOCK = threading.Lock()
_IROS_BROWSER_USE_LOCK = threading.Lock()
_IROS_LOGIN_READY = False


def get_iros_profile_dir() -> Path:
    """WorkLab 전용 인터넷등기소 Chrome 프로필 경로를 반환합니다.

    계정정보를 코드나 GitHub에 저장하지 않고, 사용자의 로컬 Chrome 프로필 안에
    쿠키/세션만 보관합니다.
    """
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        base_dir = Path(local_app_data) / "WorkLab_PAOS"
    else:
        base_dir = Path.home() / ".worklab_paos"

    profile_dir = base_dir / "iros_chrome_profile"
    profile_dir.mkdir(parents=True, exist_ok=True)
    return profile_dir


def get_regular_chrome_user_data_dir() -> Path:
    """Windows에서 사용자가 평소 사용하는 Chrome User Data 경로를 반환합니다."""
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "Google" / "Chrome" / "User Data"
    return Path.home() / "AppData" / "Local" / "Google" / "Chrome" / "User Data"


def get_regular_chrome_last_profile() -> tuple[Path, str]:
    """평소 Chrome에서 마지막으로 사용한 프로필 폴더를 찾습니다."""
    user_data_dir = get_regular_chrome_user_data_dir()
    local_state_path = user_data_dir / "Local State"

    profile_name = "Default"
    try:
        if local_state_path.exists():
            data = json.loads(local_state_path.read_text(encoding="utf-8"))
            last_used = (
                data.get("profile", {}).get("last_used")
                or data.get("profile", {}).get("last_active_profiles", [None])[0]
            )
            if last_used and (user_data_dir / str(last_used)).exists():
                profile_name = str(last_used)
    except Exception:
        pass

    profile_dir = user_data_dir / profile_name
    if not profile_dir.exists():
        fallback = user_data_dir / "Default"
        if fallback.exists():
            profile_dir = fallback
            profile_name = "Default"

    return profile_dir, profile_name


def sync_regular_chrome_security_environment() -> dict[str, Any]:
    """평소 Chrome의 확장/보안 환경을 WorkLab 전용 프로필에 복사합니다.

    인터넷등기소 보안프로그램은 Windows에 설치되어 있어도 Chrome 프로필에 연결된
    확장프로그램/설정이 없으면 '설치 필요'로 오인될 수 있습니다. Selenium이 실제
    기본 Chrome 프로필을 직접 사용하면 프로필 잠금과 최신 Chrome 보안 정책 때문에
    불안정하므로, 확장프로그램 관련 파일만 WorkLab 전용 프로필로 동기화합니다.

    로그인 쿠키/비밀번호/방문기록은 의도적으로 복사하지 않습니다.
    """
    source_root = get_regular_chrome_user_data_dir()
    source_profile, source_profile_name = get_regular_chrome_last_profile()
    target_root = get_iros_profile_dir()
    target_profile = target_root / "Default"
    target_profile.mkdir(parents=True, exist_ok=True)

    if not source_root.exists() or not source_profile.exists():
        return {
            "ok": False,
            "message": "평소 사용하는 Chrome 프로필을 찾지 못했습니다.",
            "profile": source_profile_name,
        }

    copied = []

    # Chrome 확장 프로그램의 보호 설정/등록 정보.
    for filename in ["Preferences", "Secure Preferences"]:
        src = source_profile / filename
        dst = target_profile / filename
        if src.exists():
            try:
                shutil.copy2(src, dst)
                copied.append(filename)
            except Exception:
                pass

    # 보안프로그램 연동 확장과 확장별 저장 설정.
    extension_dirs = [
        "Extensions",
        "Local Extension Settings",
        "Sync Extension Settings",
        "Managed Extension Settings",
        "Extension Rules",
        "Extension State",
    ]
    for dirname in extension_dirs:
        src = source_profile / dirname
        dst = target_profile / dirname
        if not src.exists():
            continue
        try:
            if dst.exists():
                shutil.rmtree(dst, ignore_errors=True)
            shutil.copytree(src, dst, dirs_exist_ok=True)
            copied.append(dirname)
        except Exception:
            # 일부 확장 파일이 Chrome 백그라운드 프로세스에 의해 잠겨 있어도
            # 나머지 보안환경 복사는 계속 진행합니다.
            continue

    # 확장/프로필 보호값 계산에 필요한 Chrome Local State.
    local_state_src = source_root / "Local State"
    local_state_dst = target_root / "Local State"
    if local_state_src.exists():
        try:
            shutil.copy2(local_state_src, local_state_dst)
            copied.append("Local State")
        except Exception:
            pass

    if not copied:
        return {
            "ok": False,
            "message": "Chrome 보안환경을 복사하지 못했습니다. 평소 Chrome을 완전히 종료한 뒤 다시 시도해 주세요.",
            "profile": source_profile_name,
        }

    marker = target_root.parent / "iros_chrome_environment_synced.txt"
    try:
        marker.write_text(
            f"profile={source_profile_name}\ntime={int(time.time())}\n",
            encoding="utf-8",
        )
    except Exception:
        pass

    return {
        "ok": True,
        "profile": source_profile_name,
        "copied_count": len(copied),
        "message": f"평소 Chrome({source_profile_name})의 보안환경을 WorkLab에 연결했습니다.",
    }


def get_iros_session_marker_path() -> Path:
    return get_iros_profile_dir().parent / "iros_session_saved.txt"


def has_iros_saved_session() -> bool:
    """현재 WorkLab 프로세스에서 실제 로그인된 Chrome 세션이 살아있는지 확인합니다."""
    global _IROS_LOGIN_READY
    with _IROS_LOGIN_LOCK:
        return _login_driver_alive() and _IROS_LOGIN_READY


def invalidate_iros_saved_session() -> None:
    """로그인 완료 상태만 해제합니다. 브라우저는 재로그인을 위해 유지할 수 있습니다."""
    global _IROS_LOGIN_READY
    _IROS_LOGIN_READY = False
    try:
        get_iros_session_marker_path().unlink(missing_ok=True)
    except Exception:
        pass


def _login_driver_alive() -> bool:
    global _IROS_LOGIN_DRIVER
    if _IROS_LOGIN_DRIVER is None:
        return False
    try:
        _ = _IROS_LOGIN_DRIVER.current_url
        return True
    except Exception:
        global _IROS_LOGIN_READY
        _IROS_LOGIN_DRIVER = None
        _IROS_LOGIN_READY = False
        return False


def is_iros_login_browser_open() -> bool:
    with _IROS_LOGIN_LOCK:
        return _login_driver_alive()


def _looks_like_iros_login_page(driver: webdriver.Chrome) -> bool:
    """현재 화면이 로그인 요구 화면인지 보수적으로 확인합니다."""
    try:
        current_url = (driver.current_url or "").lower()
        if any(token in current_url for token in ["login", "signin", "memberlogin"]):
            return True
    except Exception:
        pass

    try:
        password_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='password']")
        for element in password_inputs:
            try:
                if element.is_displayed():
                    return True
            except Exception:
                continue
    except Exception:
        pass

    try:
        body_text = driver.find_element(By.TAG_NAME, "body").text
        compact = re.sub(r"\s+", " ", body_text or "")
        login_phrases = ["로그인이 필요", "로그인 후 이용", "회원 로그인", "아이디 비밀번호"]
        if any(phrase in compact for phrase in login_phrases):
            return True
    except Exception:
        pass

    return False


def ensure_iros_logged_in(driver: webdriver.Chrome) -> None:
    if _looks_like_iros_login_page(driver):
        invalidate_iros_saved_session()
        raise IrosLoginRequiredError(
            "인터넷등기소 로그인 세션이 없거나 만료되었습니다. WorkLab의 '등기소 로그인'에서 다시 로그인해 주세요."
        )


def start_iros_login_browser() -> dict[str, Any]:
    """로그인용 Chrome을 엽니다.

    로그인 완료 후에도 이 브라우저를 닫지 않고 동일한 WebDriver 세션을
    등기소 크롤링에 그대로 재사용합니다.
    """
    global _IROS_LOGIN_DRIVER, _IROS_LOGIN_READY
    with _IROS_LOGIN_LOCK:
        if _login_driver_alive():
            try:
                _IROS_LOGIN_DRIVER.set_window_position(50, 50)
                _IROS_LOGIN_DRIVER.maximize_window()
            except Exception:
                pass

            # 로그인 만료 후 다시 로그인 버튼을 누른 경우 같은 창을 재사용합니다.
            if not _IROS_LOGIN_READY:
                try:
                    _IROS_LOGIN_DRIVER.get(IROS_URL)
                except Exception:
                    pass

            return {
                "ok": True,
                "browser_open": True,
                "session_active": bool(_IROS_LOGIN_READY),
                "message": (
                    "현재 로그인 연결이 이미 유지되고 있습니다."
                    if _IROS_LOGIN_READY
                    else "기존 등기소 Chrome 창을 다시 표시했습니다. 로그인 후 '로그인 완료'를 눌러 주세요."
                ),
            }

        try:
            chrome_sync = sync_regular_chrome_security_environment()
            if not chrome_sync.get("ok"):
                return {
                    "ok": False,
                    "browser_open": False,
                    "session_active": False,
                    "message": chrome_sync.get(
                        "message",
                        "평소 Chrome의 보안환경을 WorkLab에 연결하지 못했습니다."
                    ),
                }

            driver = get_driver(headless=False, use_iros_profile=True)
            driver.get(IROS_URL)
            _IROS_LOGIN_DRIVER = driver
            _IROS_LOGIN_READY = False
            profile_name = chrome_sync.get("profile", "Default")
            return {
                "ok": True,
                "browser_open": True,
                "session_active": False,
                "message": (
                    f"평소 Chrome({profile_name})의 보안환경을 연결해 등기소 창을 열었습니다. "
                    "인터넷등기소에 직접 로그인한 뒤 WorkLab에서 '로그인 완료'를 눌러 주세요."
                ),
            }
        except Exception as exc:
            _IROS_LOGIN_DRIVER = None
            _IROS_LOGIN_READY = False
            return {
                "ok": False,
                "browser_open": False,
                "session_active": False,
                "message": f"등기소 로그인 창을 열 수 없습니다: {str(exc)[:160]}",
            }


def complete_iros_login_browser() -> dict[str, Any]:
    """현재 로그인된 Chrome을 닫지 않고 크롤링용 활성 세션으로 전환합니다."""
    global _IROS_LOGIN_DRIVER, _IROS_LOGIN_READY
    with _IROS_LOGIN_LOCK:
        if not _login_driver_alive():
            _IROS_LOGIN_READY = False
            return {
                "ok": False,
                "browser_open": False,
                "session_active": False,
                "session_saved": False,
                "message": "열려 있는 등기소 로그인 창을 찾지 못했습니다. '등기소 로그인'을 다시 눌러 주세요.",
            }

        login_detected = False
        try:
            body_text = _IROS_LOGIN_DRIVER.find_element(By.TAG_NAME, "body").text or ""
            login_detected = "로그아웃" in body_text
        except Exception:
            pass

        # 실제 로그인 화면이 그대로 보이는 상태라면 완료로 처리하지 않습니다.
        if _looks_like_iros_login_page(_IROS_LOGIN_DRIVER):
            _IROS_LOGIN_READY = False
            return {
                "ok": False,
                "browser_open": True,
                "session_active": False,
                "session_saved": False,
                "message": "아직 인터넷등기소 로그인 화면입니다. 로그인한 뒤 다시 '로그인 완료'를 눌러 주세요.",
            }

        # 핵심: quit() 하지 않습니다.
        # 동일한 WebDriver/브라우저를 그대로 사용해야 인터넷등기소 인증 세션이 유지됩니다.
        _IROS_LOGIN_READY = True

        marker = get_iros_session_marker_path()
        marker.parent.mkdir(parents=True, exist_ok=True)
        marker.write_text(str(int(time.time())), encoding="utf-8")

        return {
            "ok": True,
            "browser_open": True,
            "session_active": True,
            "session_saved": True,
            "login_detected": login_detected,
            "message": "로그인 연결이 완료되었습니다. 이 Chrome 창을 닫지 않은 상태에서 크롤링을 실행해 주세요.",
        }


def get_iros_login_status() -> dict[str, Any]:
    global _IROS_LOGIN_READY

    with _IROS_LOGIN_LOCK:
        browser_open = _login_driver_alive()
        session_active = browser_open and bool(_IROS_LOGIN_READY)

    saved_at = None
    marker = get_iros_session_marker_path()
    if session_active and marker.exists():
        try:
            saved_at = int(marker.read_text(encoding="utf-8").strip())
        except Exception:
            saved_at = None
    elif not session_active:
        # 이전 버전의 stale marker가 남아 있어도 로그인 완료로 오인하지 않습니다.
        try:
            marker.unlink(missing_ok=True)
        except Exception:
            pass

    return {
        "ok": True,
        "browser_open": browser_open,
        "session_active": session_active,
        "session_saved": session_active,
        "saved_at": saved_at,
    }


def get_driver(
    *,
    headless: bool = True,
    page_load_strategy: str | None = None,
    use_iros_profile: bool = True,
) -> webdriver.Chrome:
    options = Options()

    if page_load_strategy:
        options.page_load_strategy = page_load_strategy

    if use_iros_profile:
        profile_dir = get_iros_profile_dir()
        options.add_argument(f"--user-data-dir={profile_dir}")
        options.add_argument("--profile-directory=Default")

    if headless:
        # Chrome 창을 띄우지 않고 백그라운드에서 실행합니다.
        # 최신 Chrome 기준으로 --headless=new가 일반 headless보다 화면 렌더링 호환성이 좋습니다.
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
    else:
        options.add_argument("--start-maximized")

    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
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
    ensure_iros_logged_in(driver)

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
    ensure_iros_logged_in(driver)

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
    ensure_iros_logged_in(driver)

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
        time.sleep(0.5)
        ensure_iros_logged_in(driver)
    except IrosLoginRequiredError:
        raise
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
    """등록번호검색 입력창 형식에 맞게 숫자 13자리만 반환합니다.

    실제 등록번호검색 input은 maxlength=13 이므로 하이픈을 넣지 않습니다.
    예: 110111-0003428 -> 1101110003428
    """
    text = str(value or "").strip()
    digits = re.sub(r"\D", "", text)
    return digits if digits else text


def find_registration_number_input(driver: webdriver.Chrome, wait: WebDriverWait):
    """등록번호검색 화면의 법인등록번호 입력창을 찾습니다.

    인터넷등기소 실제 요소 ID를 최우선으로 사용합니다.
    """
    input_id = "mf_wfm_potal_main_wfm_content_sbx_drokno___input"

    def locate(d):
        # 실제 등록번호검색 입력창을 ID로 직접 탐색
        try:
            element = d.find_element(By.ID, input_id)
            if element.is_displayed() and element.is_enabled():
                return element
        except Exception:
            pass

        # ID가 변경될 경우를 대비한 fallback: maxlength=13인 보이는 입력창
        for element in d.find_elements(By.CSS_SELECTOR, "input[maxlength='13']"):
            try:
                if element.is_displayed() and element.is_enabled():
                    return element
            except Exception:
                continue

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
    """같은 등록번호 입력창에 다음 13자리 번호를 넣습니다.

    화면을 다시 열지 않고 Ctrl+A -> 삭제 -> 다음 값 입력 방식으로 반복합니다.
    Selenium 입력이 반영되지 않는 경우에만 JS 이벤트를 fallback으로 사용합니다.
    """
    target = normalize_registration_number(value)
    target_digits = re.sub(r"\D", "", target)
    if len(target_digits) != 13:
        raise ValueError(f"법인등록번호는 숫자 13자리여야 합니다: {value}")

    # 1순위: 실제 사용자 입력과 동일하게 전체선택 -> 삭제 -> 입력
    try:
        input_element.click()
        input_element.send_keys(Keys.CONTROL, "a")
        input_element.send_keys(Keys.BACKSPACE)
        input_element.send_keys(target_digits)
        time.sleep(0.15)
        actual = input_element.get_attribute("value") or ""
        if re.sub(r"\D", "", actual) == target_digits:
            return actual
    except Exception:
        pass

    # 2순위 fallback: WebSquare 내부 이벤트까지 갱신
    driver.execute_script(
        """
        const el = arguments[0];
        const value = arguments[1];
        el.focus();
        el.value = '';
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.value = value;
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.blur();
        """,
        input_element,
        target_digits,
    )
    time.sleep(0.15)
    actual = input_element.get_attribute("value") or ""
    if re.sub(r"\D", "", actual) != target_digits:
        raise RuntimeError(f"등록번호 입력값 반영 실패: {target_digits}")
    return actual

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


def normalize_company_match_text(value: str) -> str:
    """상호검색 결과와 검색값 비교용 문자열을 만듭니다.

    공백/기호 차이 때문에 같은 회사가 불일치로 처리되지 않도록 한글, 영문, 숫자만 남깁니다.
    """
    return re.sub(r"[^0-9A-Za-z가-힣]", "", str(value or "")).lower()


def company_name_matches_search(search_value: str, company_name: str) -> bool:
    """현재 검색어와 실제 상호가 서로 포함 관계인지 확인합니다.

    예: 국보 ↔ 국보(KUKBO CO.,LTD.), 에티버스 ↔ 주식회사에티버스
    """
    search_norm = normalize_company_match_text(search_value)
    company_norm = normalize_company_match_text(company_name)
    if not search_norm or not company_norm:
        return False
    return search_norm in company_norm or company_norm in search_norm


def set_company_search_input(driver: webdriver.Chrome, search_box, value: str) -> None:
    """상호검색 입력값을 확실히 교체하고 실제 반영 여부까지 확인합니다."""
    target = str(value or "").strip()

    try:
        search_box.click()
        search_box.send_keys(Keys.CONTROL, "a")
        search_box.send_keys(Keys.BACKSPACE)
        search_box.send_keys(target)
        time.sleep(0.1)
        actual = search_box.get_attribute("value") or ""
        if normalize_company_match_text(actual) == normalize_company_match_text(target):
            return
    except Exception:
        pass

    # WebSquare 입력값이 Selenium clear/send_keys만으로 갱신되지 않는 경우 대비
    driver.execute_script(
        """
        const el = arguments[0];
        const value = arguments[1];
        el.focus();
        el.value = '';
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.value = value;
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        """,
        search_box,
        target,
    )
    time.sleep(0.1)
    actual = search_box.get_attribute("value") or ""
    if normalize_company_match_text(actual) != normalize_company_match_text(target):
        raise RuntimeError(f"상호검색 입력값 반영 실패: {target}")


def get_visible_company_search_rows(driver: webdriver.Chrome, search_value: str):
    """현재 검색값과 상호가 맞는 상호검색 결과행을 반환합니다.

    WebSquare는 이전 검색 결과 DOM을 숨겨둔 채 남기거나, 새 결과행을 렌더링하는 순간
    Selenium의 is_displayed()가 False로 잡히는 경우가 있습니다. 따라서 표시 여부만으로
    결과를 버리지 않고, 실제 상호(명칭)가 현재 검색값과 맞는지를 최종 기준으로 사용합니다.

    기존 '*' 표시 행을 우선 사용하되, '*' 표시가 없는 정상 결과도 놓치지 않도록
    같은 결과표의 전체 행을 fallback으로 함께 확인합니다.
    """
    candidate_xpaths = [
        "//td[@data-col_id='no']/span[normalize-space(text())='*']/parent::td/parent::tr",
        "//td[@data-col_id='no']/parent::tr",
    ]

    candidates = []
    seen_ids = set()
    for xpath in candidate_xpaths:
        for row in driver.find_elements(By.XPATH, xpath):
            try:
                row_id = getattr(row, "id", None) or str(id(row))
                if row_id in seen_ids:
                    continue
                seen_ids.add(row_id)
                candidates.append(row)
            except Exception:
                continue

    matched_rows = []
    seen_keys = set()
    for row in candidates:
        try:
            header_map = get_iros_header_map(row)
            company_name = get_iros_cell_text(row, header_map, "상호(명칭)", 4)
            if not company_name_matches_search(search_value, company_name):
                continue

            corp_reg_num = get_iros_cell_text(row, header_map, "법인등록번호", 5)
            company_address = get_iros_cell_text(row, header_map, "본점소재지", 6)
            dedupe_key = (
                normalize_company_match_text(company_name),
                re.sub(r"\D", "", corp_reg_num),
                normalize_company_match_text(company_address),
            )
            if dedupe_key in seen_keys:
                continue
            seen_keys.add(dedupe_key)
            matched_rows.append(row)
        except Exception:
            continue

    return matched_rows


def _close_company_search_popup_if_present(driver: webdriver.Chrome) -> None:
    """상호검색 중 표시되는 확인 팝업이 있으면 닫습니다."""
    try:
        popup_buttons = driver.find_elements(
            By.XPATH,
            "//a[contains(text(), '확인') and contains(@class, 'btn solid medium color-main')]",
        )
        for popup_button in popup_buttons:
            try:
                if popup_button.is_displayed():
                    driver.execute_script("arguments[0].click();", popup_button)
                    time.sleep(0.1)
                    return
            except Exception:
                continue
    except Exception:
        pass


def _visible_company_no_result_message(driver: webdriver.Chrome) -> bool:
    """현재 화면에 실제로 보이는 '검색결과 없음' 메시지가 있는지 확인합니다."""
    messages = [
        "검색결과가 없습니다",
        "조회된 결과가 없습니다",
        "검색 결과가 없습니다",
    ]
    for message in messages:
        try:
            elements = driver.find_elements(
                By.XPATH,
                f"//*[contains(normalize-space(.), '{message}')]",
            )
            for element in elements:
                try:
                    if element.is_displayed():
                        return True
                except Exception:
                    continue
        except Exception:
            continue
    return False


def wait_for_company_search_rows(
    driver: webdriver.Chrome,
    search_value: str,
    timeout: float,
) -> tuple[list, bool]:
    """현재 검색어 결과를 짧은 간격으로 확인합니다.

    결과가 나타나면 timeout을 다 채우지 않고 즉시 반환합니다.
    반환값 두 번째 값은 '검색결과 없음' 메시지가 실제로 확인됐는지 여부입니다.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        _close_company_search_popup_if_present(driver)

        rows = get_visible_company_search_rows(driver, search_value)
        if rows:
            return rows, False

        if _visible_company_no_result_message(driver):
            return [], True

        time.sleep(0.2)

    return [], False


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
    browser_lock_acquired = False
    using_live_login_driver = False
    try:
        # 동일 로그인 브라우저를 동시에 두 요청이 조작하지 못하도록 막습니다.
        browser_lock_acquired = _IROS_BROWSER_USE_LOCK.acquire(blocking=False)
        if not browser_lock_acquired:
            yield make_event(
                "error",
                code="IROS_ALREADY_RUNNING",
                message="이미 등기소 크롤링이 실행 중입니다. 현재 작업이 끝난 뒤 다시 시도해 주세요.",
                logs=logs.copy(),
            )
            return

        global _IROS_LOGIN_READY
        with _IROS_LOGIN_LOCK:
            if not _login_driver_alive() or not _IROS_LOGIN_READY:
                yield make_event(
                    "error",
                    code="IROS_LOGIN_REQUIRED",
                    message="인터넷등기소 로그인이 필요합니다. 먼저 '등기소 로그인' 후 '로그인 완료'를 눌러 주세요.",
                    logs=logs.copy(),
                )
                return

            # 새 Chrome을 열지 않고 사용자가 로그인한 바로 그 WebDriver를 사용합니다.
            driver = _IROS_LOGIN_DRIVER
            using_live_login_driver = True

        yield log("🔐 현재 로그인된 인터넷등기소 Chrome 세션을 그대로 사용합니다.")
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

        # 로그인한 브라우저와 동일한 세션을 유지해야 하므로 새 headless Chrome을 만들지 않습니다.
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

                    # 등록번호검색 화면은 최초 1회만 진입합니다.
                    # 이후에는 동일한 실제 입력창에 다음 13자리 법인번호를 넣고 검색 버튼만 반복 클릭합니다.
                    previous_signature = registration_result_signature(driver)
                    registration_input = find_registration_number_input(driver, wait)
                    set_registration_number_input(driver, registration_input, search_value)

                    # 입력값이 정확히 들어갔는지 로그/검증용으로 다시 확인
                    actual_value = registration_input.get_attribute("value") or ""
                    if re.sub(r"\D", "", actual_value) != re.sub(r"\D", "", search_value):
                        raise RuntimeError(f"등록번호 입력 확인 실패: {search_value}")

                    search_button = find_nearest_search_button(driver, registration_input, wait)
                    _click_element(driver, search_button)

                    # 검색 결과 갱신 대기. 화면을 다시 열지는 않습니다.
                    time.sleep(0.5)
                    deadline = time.time() + 8.0
                    while time.time() < deadline:
                        body_text = driver.find_element(By.TAG_NAME, "body").text
                        current_signature = registration_result_signature(driver)

                        if current_signature and (not previous_signature or current_signature != previous_signature):
                            per_rows, per_columns = extract_registration_search_results(driver, company)
                            if per_rows:
                                break

                        if any(text in body_text for text in ["검색결과가 없습니다", "조회된 결과가 없습니다", "검색 결과가 없습니다"]):
                            no_result = True
                            break

                        time.sleep(0.25)

                    if not per_rows:
                        status_text = "검색결과 없음" if no_result else "검색 결과 갱신 확인 실패"
                        yield log(f"⏭️ [{current_no}/{total}] {company} — {status_text}")
                        yield make_event("progress", current=current_no, total=total, result_count=len(results))
                        continue

                    merge_columns(per_columns)
                    results.extend(per_rows)
                    yield log(f"✅ [{current_no}/{total}] {company} — {len(per_rows)}건")
                    yield make_event("progress", current=current_no, total=total, result_count=len(results))
                    continue

                # 상호검색 로직
                # 1차 검색은 최대 5초까지만 확인하되 결과가 뜨는 즉시 다음 단계로 넘어갑니다.
                # 정상 데이터인데 화면 갱신이 늦은 경우에는 같은 검색어로 1회만 재검색하고
                # 추가 최대 4초 동안 확인합니다.
                set_company_search_input(driver, search_box, company)

                search_button = wait.until(
                    EC.element_to_be_clickable(
                        (By.ID, "mf_wfm_potal_main_wfm_content_btn_conm_name_search")
                    )
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", search_button)
                driver.execute_script("arguments[0].click();", search_button)

                rows, explicit_no_result = wait_for_company_search_rows(
                    driver,
                    company,
                    timeout=5.0,
                )

                # '검색결과 없음'이 명확히 뜬 경우에는 불필요한 재검색을 하지 않습니다.
                # 결과가 늦게 갱신된 것으로 보이는 경우에만 동일 검색어로 1회 재시도합니다.
                if not rows and not explicit_no_result:
                    yield log(f"↻ [{current_no}/{total}] {company} — 결과 확인이 늦어 1회 재검색합니다.")
                    set_company_search_input(driver, search_box, company)
                    search_button = wait.until(
                        EC.element_to_be_clickable(
                            (By.ID, "mf_wfm_potal_main_wfm_content_btn_conm_name_search")
                        )
                    )
                    driver.execute_script("arguments[0].click();", search_button)
                    rows, explicit_no_result = wait_for_company_search_rows(
                        driver,
                        company,
                        timeout=4.0,
                    )

                if not rows:
                    status_text = "검색결과 없음" if explicit_no_result else "검색 결과 확인 실패"
                    yield log(f"⏭️ [{current_no}/{total}] {company} — {status_text}")
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

    except IrosLoginRequiredError as e:
        invalidate_iros_saved_session()
        try:
            if driver:
                driver.set_window_position(50, 50)
                driver.maximize_window()
        except Exception:
            pass
        yield make_event(
            "error",
            code="IROS_LOGIN_REQUIRED",
            message=str(e),
            logs=logs.copy(),
        )
    except TimeoutException:
        yield make_event(
            "error",
            code="IROS_SCREEN_TIMEOUT",
            message=(
                "등기소 화면 전환을 확인하지 못했습니다. 로그인된 Chrome 창을 확인해 주세요. "
                "로그인이 풀렸거나 인터넷등기소 화면 구조가 변경된 경우 '등기소 로그인'에서 다시 로그인한 뒤 재시도해 주세요."
            ),
            logs=logs.copy(),
        )
    except Exception as e:
        error_text = str(e).strip()
        if not error_text or error_text.startswith("Message:"):
            error_text = type(e).__name__
        yield make_event(
            "error",
            message=f"등기소 크롤링 실패: {error_text[:240]}",
            logs=logs.copy(),
        )
    finally:
        # 로그인 세션 유지를 위해 live driver는 quit()하지 않습니다.
        if driver and not using_live_login_driver:
            try:
                driver.quit()
            except Exception:
                pass
        if browser_lock_acquired:
            try:
                _IROS_BROWSER_USE_LOCK.release()
            except Exception:
                pass


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
