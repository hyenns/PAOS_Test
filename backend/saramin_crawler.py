from __future__ import annotations

import re
import time
from typing import Iterable, Iterator, Any
from urllib.parse import quote

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException

from backend.iros_crawler import get_driver, make_event

SARAMIN_URL = "https://www.saramin.co.kr/zf_user/"

SARAMIN_RESULT_COLUMNS = [
    "검색값",
    "회사명",
    "설립일",
    "대표자명",
    "업종",
    "기업형태",
    "사원수",
    "매출액",
    "평균연봉",
    "홈페이지",
    "사업내용",
    "주소",
    "사람인URL",
    "사람인_재무정보URL",
    "재무_기준연도",
    "재무_매출액",
    "재무_동종업계순위",
    "재무_영업이익",
    "재무_당기순이익",
    "재무_자본금",
    "신용등급",
]

for _metric in ["매출액", "영업이익", "당기순이익", "자본금"]:
    for _year in ["2019", "2020", "2021", "2022", "2023", "2024", "2025"]:
        SARAMIN_RESULT_COLUMNS.append(f"{_metric}_{_year}")


def clean_text(text: object) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def normalize_company_key(value: object) -> str:
    """검색어와 사람인 결과 회사명을 비교하기 위한 간단 정규화."""
    text = clean_text(value)
    text = text.replace("㈜", "")
    text = re.sub(r"\(주\)|\(유\)|\(사\)|주식회사|유한회사|사단법인|재단법인", "", text)
    text = re.sub(r"[\s\-\_\.,·ㆍ/\\()\[\]{}]", "", text)
    return text.lower()


def is_related_company(keyword: str, corp_name: str) -> bool:
    keyword_key = normalize_company_key(keyword)
    name_key = normalize_company_key(corp_name)
    if not keyword_key or not name_key:
        return False
    return keyword_key in name_key or name_key in keyword_key


def extract_between(body_text: str, label: str, stop_labels: list[str]) -> str:
    pattern = rf"{re.escape(label)}\s*\n\s*(.*?)\s*(?=\n\s*(?:{'|'.join(map(re.escape, stop_labels))})\s*\n|$)"
    match = re.search(pattern, body_text, flags=re.S)
    if not match:
        return ""
    value = match.group(1).strip()
    value = re.sub(r"\n+", " ", value)
    value = re.sub(r"\s{2,}", " ", value)
    return value.replace("지도보기", "").strip()


def get_detail_from_page(driver) -> dict[str, str]:
    body_text = driver.find_element(By.TAG_NAME, "body").text
    stop_labels = [
        "업종", "대표자명", "홈페이지", "사업내용", "주소", "기업비전",
        "출처", "우리는", "채용정보", "연봉정보", "재무정보", "관심기업"
    ]

    # 주소는 HTML의 dt/dd 구조에서 title 속성값으로 직접 추출합니다.
    address = ""
    try:
        groups = driver.find_elements(By.CSS_SELECTOR, "div.company_details_group")
        for group in groups:
            try:
                title = group.find_element(By.CSS_SELECTOR, "dt.tit").text.strip()
                if title == "주소":
                    address = group.find_element(By.CSS_SELECTOR, "dd.desc p.ellipsis").get_attribute("title").strip()
                    break
            except Exception:
                continue
    except Exception:
        pass

    detail = {
        "업종": extract_between(body_text, "업종", stop_labels),
        "대표자명": extract_between(body_text, "대표자명", stop_labels).replace("/", ","),
        "홈페이지": extract_between(body_text, "홈페이지", stop_labels),
        "사업내용": extract_between(body_text, "사업내용", stop_labels),
        "주소": address,
        "사원수": "",
        "설립일": "",
        "기업형태": "",
        "매출액": "",
        "평균연봉": "",
    }

    match = re.search(r"(\d{4}년\s*\d{1,2}월\s*\d{1,2}일\s*설립)", body_text)
    if match:
        detail["설립일"] = match.group(1).replace(" 설립", "")

    match = re.search(r"기업형태:\s*([^\n]+)", body_text)
    if match:
        detail["기업형태"] = clean_text(match.group(1))
    else:
        match = re.search(r"\n([^\n]+)\n기업형태\n", body_text)
        if match:
            detail["기업형태"] = clean_text(match.group(1))

    match = re.search(r"\n([0-9,]+\s*명)\s*\n(?:출처:.*\n)?사원수", body_text)
    if match:
        detail["사원수"] = clean_text(match.group(1))

    match = re.search(r"\n([^\n]*억[^\n]*만원|[^\n]*억)\s*\n매출액", body_text)
    if match:
        detail["매출액"] = clean_text(match.group(1))

    return detail


def make_finance_url(corp_url: str) -> str:
    if not corp_url:
        return ""
    if "view-inner-finance" in corp_url:
        return corp_url
    return corp_url.replace("/zf_user/company-info/view?", "/zf_user/company-info/view-inner-finance?")


def get_finance_from_page(driver) -> dict[str, str]:
    finance: dict[str, str] = {
        "재무_기준연도": "",
        "재무_매출액": "",
        "재무_동종업계순위": "",
        "재무_영업이익": "",
        "재무_당기순이익": "",
        "재무_자본금": "",
        "신용등급": "",
    }

    metrics = ["매출액", "영업이익", "당기순이익", "자본금"]
    years = ["2019", "2020", "2021", "2022", "2023", "2024", "2025"]
    for metric in metrics:
        for year in years:
            finance[f"{metric}_{year}"] = ""

    try:
        WebDriverWait(driver, 5).until(
            lambda d: d.find_elements(By.CSS_SELECTOR, "div.box_finance, section")
        )
    except TimeoutException:
        return finance

    boxes = driver.find_elements(By.CSS_SELECTOR, "div.box_finance")
    for box in boxes:
        try:
            title = clean_text(box.find_element(By.CSS_SELECTOR, "h3.tit_finance").text)
        except Exception:
            continue

        if title not in metrics:
            continue

        try:
            summary_items = box.find_elements(By.CSS_SELECTOR, "ul.list_summary li")
            for li in summary_items:
                li_text = clean_text(li.text)
                year_match = re.search(r"(\d{4})년\s*기준", li_text)
                if year_match and not finance["재무_기준연도"]:
                    finance["재무_기준연도"] = year_match.group(1)

                try:
                    strong = clean_text(li.find_element(By.TAG_NAME, "strong").text)
                except Exception:
                    strong = ""

                try:
                    value = clean_text(li.find_element(By.CSS_SELECTOR, "span.num").text)
                except Exception:
                    value = ""

                if "기준" in strong and value:
                    finance[f"재무_{title}"] = value
                elif "동종업계" in strong and title == "매출액":
                    finance["재무_동종업계순위"] = value
        except Exception:
            pass

        try:
            desc = clean_text(box.find_element(By.CSS_SELECTOR, "p.desc_finance").text)
            if title == "자본금" and desc and not finance["재무_자본금"]:
                finance["재무_자본금"] = desc
        except Exception:
            pass

        try:
            graph_items = box.find_elements(By.CSS_SELECTOR, "div.wrap_graph")
            for graph in graph_items:
                try:
                    year = clean_text(graph.find_element(By.CSS_SELECTOR, "em.tit_graph").text)
                    value = clean_text(graph.find_element(By.CSS_SELECTOR, "span.txt_value").text)
                    if year:
                        finance[f"{title}_{year}"] = value
                except Exception:
                    continue
        except Exception:
            pass

    try:
        grade = clean_text(driver.find_element(By.CSS_SELECTOR, "dl.chart_company_grade_value dd").text)
        finance["신용등급"] = grade
    except Exception:
        pass

    return finance


def run_saramin_crawler_events(
    company_inputs: Iterable[str],
    *,
    max_results_per_keyword: int = 5,
    headless: bool = True,
) -> Iterator[dict[str, Any]]:
    """사람인 기업정보 크롤링을 실행하면서 진행 로그를 실시간 이벤트로 반환합니다."""
    company_inputs = [str(v).strip() for v in company_inputs if str(v).strip()]
    max_results_per_keyword = max(1, min(int(max_results_per_keyword or 5), 5))

    results: list[dict[str, str]] = []
    logs: list[str] = []
    total = len(company_inputs)

    def log(message: str) -> dict[str, Any]:
        logs.append(message)
        return make_event("log", message=message, logs=logs.copy())

    driver = None
    try:
        yield log("🌐 사람인 기업정보 검색을 시작합니다.")
        yield log(f"검색어별 상위 최대 {max_results_per_keyword}개 기업의 기업소개·재무정보를 수집합니다.")

        driver = get_driver(headless=headless)
        wait = WebDriverWait(driver, 10)

        for idx, keyword in enumerate(company_inputs):
            current_no = idx + 1
            keyword = str(keyword).strip()

            try:
                yield log(f"🔎 [{current_no}/{total}] {keyword} 검색 중...")

                search_url = f"https://www.saramin.co.kr/zf_user/search/company?searchword={quote(keyword)}"
                driver.get(search_url)

                try:
                    wait.until(lambda d: "기업정보" in d.find_element(By.TAG_NAME, "body").text)
                except TimeoutException:
                    yield log(f"⏭️ [{current_no}/{total}] {keyword} — 검색 페이지 로딩 실패")
                    yield make_event("progress", current=current_no, total=total, result_count=len(results))
                    continue

                time.sleep(1)

                no_result_boxes = driver.find_elements(By.CSS_SELECTOR, "div.info_no_result, div.no_result")
                no_result_text = "\n".join([box.text for box in no_result_boxes]).strip()
                if no_result_boxes and ("검색결과가 없습니다" in no_result_text or "검색어를 다시 확인" in no_result_text):
                    yield log(f"⏭️ [{current_no}/{total}] {keyword} — 검색 실패")
                    yield make_event("progress", current=current_no, total=total, result_count=len(results))
                    continue

                links = driver.find_elements(By.CSS_SELECTOR, "a[href*='/zf_user/company-info/view?csn=']")
                corp_list: list[dict[str, str]] = []
                seen: set[str] = set()

                for link in links:
                    try:
                        name = clean_text(link.text)
                        url = link.get_attribute("href")
                        if not name or not url or url in seen:
                            continue
                        if "/company-info/" not in url:
                            continue
                        seen.add(url)
                        corp_list.append({"corp_name": name, "corp_url": url})
                    except Exception:
                        continue

                valid_corp_list = [
                    corp for corp in corp_list
                    if is_related_company(keyword, corp.get("corp_name", ""))
                ][:max_results_per_keyword]

                if not valid_corp_list:
                    yield log(f"⏭️ [{current_no}/{total}] {keyword} — 검색결과 없음")
                    yield make_event("progress", current=current_no, total=total, result_count=len(results))
                    continue

                count = 0
                for corp in valid_corp_list:
                    try:
                        yield log(f"  ↳ {corp['corp_name']} 상세정보 수집 중...")

                        driver.get(corp["corp_url"])
                        wait.until(lambda d: d.find_elements(By.TAG_NAME, "body"))
                        time.sleep(1)
                        detail = get_detail_from_page(driver)

                        finance_url = make_finance_url(corp["corp_url"])
                        finance: dict[str, str] = {}
                        try:
                            driver.get(finance_url)
                            wait.until(lambda d: d.find_elements(By.TAG_NAME, "body"))
                            time.sleep(1)
                            finance = get_finance_from_page(driver)
                        except Exception:
                            finance = get_finance_from_page(driver)

                        row = {
                            "검색값": keyword,
                            "회사명": corp["corp_name"],
                            "설립일": detail.get("설립일", ""),
                            "대표자명": detail.get("대표자명", ""),
                            "업종": detail.get("업종", ""),
                            "기업형태": detail.get("기업형태", ""),
                            "사원수": detail.get("사원수", ""),
                            "매출액": detail.get("매출액", ""),
                            "평균연봉": detail.get("평균연봉", ""),
                            "홈페이지": detail.get("홈페이지", ""),
                            "사업내용": detail.get("사업내용", ""),
                            "주소": detail.get("주소", ""),
                            "사람인URL": corp["corp_url"],
                            "사람인_재무정보URL": finance_url,
                        }
                        row.update(finance)
                        results.append(row)
                        count += 1
                    except Exception as e:
                        yield log(f"  ↳ {corp.get('corp_name', '기업명 없음')} — 상세 수집 실패: {str(e)[:60]}")
                        continue

                yield log(f"✅ [{current_no}/{total}] {keyword} — {count}건")
                yield make_event("progress", current=current_no, total=total, result_count=len(results))

            except Exception as e:
                yield log(f"❌ [{current_no}/{total}] {keyword} — 조회 실패: {str(e)[:80]}")
                yield make_event("progress", current=current_no, total=total, result_count=len(results))

        yield log("🎉 사람인 크롤링 완료")
        yield make_event("complete", results=results, logs=logs.copy())

    except Exception as e:
        yield make_event("error", message=f"사람인 크롤링 실패: {str(e)}", logs=logs.copy())
    finally:
        if driver:
            driver.quit()


def run_saramin_crawler(
    company_inputs: Iterable[str],
    *,
    max_results_per_keyword: int = 5,
    headless: bool = True,
) -> tuple[list[dict], list[str]]:
    final_results: list[dict] = []
    final_logs: list[str] = []
    for event in run_saramin_crawler_events(
        company_inputs,
        max_results_per_keyword=max_results_per_keyword,
        headless=headless,
    ):
        if event.get("type") in {"log", "error", "complete"}:
            final_logs = event.get("logs", final_logs)
        if event.get("type") == "complete":
            final_results = event.get("results", [])
        if event.get("type") == "error":
            raise RuntimeError(event.get("message", "사람인 크롤링 실패"))
    return final_results, final_logs
