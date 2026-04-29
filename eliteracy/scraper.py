"""
eliteracy.edu.tw 得獎作品下載器
下載近五年（110-114年）國小階段得獎作品

Usage:
    python scraper.py                    # 下載所有比賽的國小作品
    python scraper.py --comp-ids 3       # 只下載 comp_id=3（112年圖文）
    python scraper.py --comp-ids 5 6 7   # 只下載教案競賽
    python scraper.py --all-groups       # 不限國小，下載全部組別
"""

import requests
import re
import csv
import json
import time
import warnings
import argparse
from pathlib import Path
from lxml import html as lxhtml
from urllib.parse import urljoin

warnings.filterwarnings("ignore", message="Unverified HTTPS request")

# ─────────────────────────────────────────────
# 常數設定
# ─────────────────────────────────────────────

BASE_URL        = "https://eliteracy.edu.tw"
EXHIBITIONS_URL = f"{BASE_URL}/Exhibitions.aspx"

# comp_id → 年度 / 類型 / 說明 / 是否為教案彙整頁
COMPETITION_MAP = {
    1: {"year": 110, "type": "guwen",  "name": "110年得獎作品",       "has_guoxiao": False, "aggregate": False},
    2: {"year": 111, "type": "guwen",  "name": "111年得獎作品",       "has_guoxiao": False, "aggregate": False},
    3: {"year": 112, "type": "guwen",  "name": "112年得獎作品(國小)", "has_guoxiao": True,  "aggregate": False},
    4: {"year": 113, "type": "guwen",  "name": "113年圖文創作比賽",   "has_guoxiao": False, "aggregate": False},
    5: {"year": 114, "type": "jiaoan", "name": "114年教案徵件",       "has_guoxiao": True,  "aggregate": True},
    6: {"year": 113, "type": "jiaoan", "name": "113年教案作品",       "has_guoxiao": True,  "aggregate": True},
    7: {"year": 112, "type": "jiaoan", "name": "112年教案作品",       "has_guoxiao": True,  "aggregate": True},
}

JIAOAN_AWARDS   = ["特優", "優選", "佳作", "優等", "第一名", "第二名", "第三名"]
LEVEL_PATTERN   = re.compile(r"(國小|國中|高中職|大專院校|一般民眾)")

DELAY_BETWEEN_REQUESTS = 1.5   # 秒，每筆詳細頁之間
DELAY_BETWEEN_PAGES    = 2.0   # 秒，每次翻頁 POST 之間
REQUEST_TIMEOUT        = 30    # 秒
MAX_RETRIES            = 3

OUTPUT_DIR    = Path(__file__).parent / "downloads"
STATE_FILE    = Path(__file__).parent / "scrape_state.json"
METADATA_FILE = Path(__file__).parent / "metadata.csv"

METADATA_FIELDS = [
    "entry_id", "competition_id", "year", "competition_name",
    "award_level", "title", "group", "team_name", "description",
    "school_level", "poster_url", "thumb_url",
    "poster_local", "thumb_local",
    "has_download_note", "external_links",
    "detail_url",
]

HEADERS = {
    "User-Agent":      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                       "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
    "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

# ─────────────────────────────────────────────
# Session / HTTP 工具
# ─────────────────────────────────────────────

def make_session() -> requests.Session:
    session = requests.Session()
    session.verify = False
    session.headers.update(HEADERS)
    return session


def get_with_retry(session: requests.Session, url: str, **kwargs) -> requests.Response:
    for attempt in range(MAX_RETRIES):
        try:
            resp = session.get(url, timeout=REQUEST_TIMEOUT, **kwargs)
            resp.raise_for_status()
            resp.encoding = "utf-8"
            return resp
        except requests.RequestException as e:
            if attempt == MAX_RETRIES - 1:
                raise
            wait = 2 ** attempt * 2
            print(f"    [Retry {attempt+1}/{MAX_RETRIES}] {e}，等待 {wait}s")
            time.sleep(wait)


def post_with_retry(session: requests.Session, url: str, data: dict, **kwargs) -> requests.Response:
    post_headers = {**HEADERS, "Content-Type": "application/x-www-form-urlencoded"}
    for attempt in range(MAX_RETRIES):
        try:
            resp = session.post(url, data=data, headers=post_headers,
                                timeout=REQUEST_TIMEOUT, **kwargs)
            resp.raise_for_status()
            resp.encoding = "utf-8"
            return resp
        except requests.RequestException as e:
            if attempt == MAX_RETRIES - 1:
                raise
            wait = 2 ** attempt * 2
            print(f"    [Retry {attempt+1}/{MAX_RETRIES}] {e}，等待 {wait}s")
            time.sleep(wait)

# ─────────────────────────────────────────────
# ASP.NET 表單 / 分頁處理
# ─────────────────────────────────────────────

def extract_viewstate(tree) -> dict:
    """提取 ASP.NET 隱藏表單欄位。"""
    def xval(xpath):
        vals = tree.xpath(xpath)
        return vals[0] if vals else ""
    return {
        "__VIEWSTATE":          xval('//input[@id="__VIEWSTATE"]/@value'),
        "__VIEWSTATEGENERATOR": xval('//input[@id="__VIEWSTATEGENERATOR"]/@value'),
        "__EVENTVALIDATION":    xval('//input[@id="__EVENTVALIDATION"]/@value'),
        "__VIEWSTATEENCRYPTED": xval('//input[@id="__VIEWSTATEENCRYPTED"]/@value'),
    }


def extract_next_page_target(tree) -> str | None:
    """
    找出「下一頁」按鈕的 PostBack 目標字串。
    若在最後一頁（無 href），回傳 None。
    """
    # 先嘗試找分頁器中帶有 href 的「下一頁」連結
    next_links = tree.xpath(
        '//*[contains(@id,"ExhibitionDataPager")]'
        '//a[contains(text(),"下一頁") and @href]'
    )
    if not next_links:
        return None
    href = next_links[0].get("href", "")
    m = re.search(r"__doPostBack\('([^']+)','([^']*)'\)", href)
    return m.group(1) if m else None

# ─────────────────────────────────────────────
# 列表頁解析
# ─────────────────────────────────────────────

def get_listing_page_entries(tree) -> list[dict]:
    """從列表頁解析所有作品卡片。"""
    cards = tree.xpath('//*[@class="grid_std"]')
    entries = []
    for card in cards:
        links = card.xpath('.//a[contains(@href,"Exhibition.aspx")]/@href')
        if not links:
            continue
        href = links[0]
        m = re.search(r"id=(\d+)", href)
        if not m:
            continue
        entry_id = int(m.group(1))
        thumb_src = card.xpath('.//img/@src')
        title_text = card.xpath('.//*[contains(@id,"TitleLabel")]//text()')
        # 若沒有 TitleLabel，取 gridimg_text 內容
        if not title_text:
            title_text = card.xpath('.//*[@class="gridimg_text"]//text()')
        entries.append({
            "entry_id":   entry_id,
            "detail_url": urljoin(BASE_URL + "/", href),
            "thumb_url":  urljoin(BASE_URL + "/", thumb_src[0]) if thumb_src else None,
            "title_raw":  "".join(title_text).strip(),
        })
    return entries


def scrape_competition_listing(session: requests.Session, comp_id: int) -> list[dict]:
    """分頁爬取某競賽的所有列表頁，回傳所有 entry stub。"""
    comp_url = f"{EXHIBITIONS_URL}?id={comp_id}"
    print(f"\n  載入列表第 1 頁：{comp_url}")
    resp = get_with_retry(session, comp_url)
    tree = lxhtml.fromstring(resp.content)

    all_entries = get_listing_page_entries(tree)
    viewstate   = extract_viewstate(tree)
    page_num    = 1

    while True:
        next_target = extract_next_page_target(tree)
        if not next_target:
            break

        time.sleep(DELAY_BETWEEN_PAGES)
        page_num += 1
        print(f"  載入列表第 {page_num} 頁（POST）...")

        form_data = {
            "__EVENTTARGET":                   next_target,
            "__EVENTARGUMENT":                 "",
            "ctl00$SiteContent$AnnualID":      str(comp_id),
            "ctl00$SiteContent$GroupNum":      "0",
            **viewstate,
        }
        resp = post_with_retry(session, comp_url, form_data)
        tree = lxhtml.fromstring(resp.content)

        page_entries = get_listing_page_entries(tree)
        all_entries.extend(page_entries)
        viewstate = extract_viewstate(tree)   # 每次回應都需重新取

    print(f"  列表共找到 {len(all_entries)} 筆")
    return all_entries

# ─────────────────────────────────────────────
# 詳細頁解析（圖文作品）
# ─────────────────────────────────────────────

def parse_detail_page(tree, detail_url: str) -> dict:
    """解析圖文作品詳細頁。"""
    result = {"detail_url": detail_url}

    # 標題（含獎項，例如「金獎：設好密碼不被盜」）
    h1_texts = tree.xpath("//h1//text()")
    title_raw = "".join(h1_texts).strip()
    result["title_raw"] = title_raw

    if "：" in title_raw:
        parts = title_raw.split("：", 1)
        result["award_level"] = parts[0].strip()
        result["title"]       = parts[1].strip()
    else:
        result["award_level"] = ""
        result["title"]       = title_raw

    # 從 <p> 元素取得組別 / 團隊名稱 / 作品簡介
    result["group"]       = ""
    result["team_name"]   = ""
    result["description"] = ""

    paras = tree.xpath("//main//p | //div[@class='article_content']//p")
    for p in paras:
        text = p.text_content().strip()
        if "組別" in text and not result["group"]:
            m = re.search(r"組別[：:]\s*(.+)", text)
            result["group"] = m.group(1).strip() if m else text
        elif "團隊名稱" in text and not result["team_name"]:
            m = re.search(r"團隊名稱[：:]\s*(.+)", text)
            result["team_name"] = m.group(1).strip() if m else ""
        elif "作品簡介" in text and not result["description"]:
            m = re.search(r"作品簡介[：:]\s*([\s\S]+)", text)
            result["description"] = m.group(1).strip() if m else text

    # 海報圖片
    poster_src = tree.xpath('//img[@id="ctl00_SiteContent_ExhibitionFigure"]/@src')
    if not poster_src:
        # 備用：找 main 內最大的圖片
        poster_src = tree.xpath("//main//img[@src]/@src")
    result["poster_url"] = urljoin(BASE_URL + "/", poster_src[0]) if poster_src else None

    # 尋找外部下載連結
    ext_links = []
    for a in tree.xpath("//main//a[@href] | //article//a[@href]"):
        href = a.get("href", "")
        if any(kw in href for kw in ["drive.google", "docs.google", ".pdf", ".docx",
                                      ".pptx", ".zip", "reurl.cc", "bit.ly", "goo.gl"]):
            ext_links.append(href)
    result["external_links"]    = "; ".join(ext_links)
    result["has_download_note"] = bool(ext_links)

    return result


def is_guoxiao(detail: dict) -> bool:
    """判斷是否為國小組（含寬鬆比對）。"""
    group = detail.get("group", "")
    if "國小" in group:
        return True
    # comp_id=3 的組別欄位可能寫「高年級」，需配合競賽已知全為國小
    return False

# ─────────────────────────────────────────────
# 教案彙整頁解析
# ─────────────────────────────────────────────

def parse_jiaoan_aggregate(tree) -> list[dict]:
    """
    解析教案彙整頁（無 HTML 表格，為純文字分行格式）。
    回傳各條目的 dict 列表。
    """
    main_els = tree.xpath("//main") or tree.xpath("//article") or tree.xpath("//body")
    if not main_els:
        return []

    full_text = main_els[0].text_content()

    # 定位表格區段
    start_markers = ["獎項", "得獎作品名稱", "作品名稱"]
    end_markers   = ["返回得獎作品", "回到列表", "回上頁"]

    idx1, idx2 = -1, len(full_text)
    for mk in start_markers:
        pos = full_text.find(mk)
        if pos >= 0:
            idx1 = pos
            break
    for mk in end_markers:
        pos = full_text.find(mk)
        if 0 <= pos < idx2:
            idx2 = pos

    if idx1 < 0:
        return []

    table_text = full_text[idx1:idx2]

    # 切行並去除空行 / 不可視字元
    lines = [
        l.strip().replace("\xa0", "").replace("\u3000", "")
        for l in table_text.splitlines()
        if l.strip().replace("\xa0", "").replace("\u3000", "")
    ]

    # 以獎項關鍵字為新列的起點
    award_pattern = re.compile(r"^(" + "|".join(JIAOAN_AWARDS) + r")$")
    rows = []
    current = []
    header_done = False

    for line in lines:
        if not header_done:
            # 跳過表頭（獎項, 作品名稱, 作品議題, 適用學制, 作品連結）
            if award_pattern.match(line):
                header_done = True
                current = [line]
            continue
        if award_pattern.match(line):
            if current:
                rows.append(current)
            current = [line]
        else:
            current.append(line)
    if current:
        rows.append(current)

    results = []
    for row in rows:
        if not row:
            continue
        award = row[0]
        # 找學制欄位
        level = ""
        for field in row:
            m = LEVEL_PATTERN.search(field)
            if m:
                level = m.group(1)
                break
        # 標題：通常是第 2 個元素
        title = row[1] if len(row) > 1 else ""
        has_dl = any("作品下載" in f or "下載" in f for f in row)
        results.append({
            "award_level":       award,
            "title":             title,
            "school_level":      level,
            "has_download_note": has_dl,
        })

    return results


def scrape_jiaoan_competition(session: requests.Session, comp_id: int,
                               comp_info: dict) -> list[dict]:
    """爬取教案彙整型競賽，回傳國小條目列表。"""
    listing_entries = scrape_competition_listing(session, comp_id)
    results = []

    for stub in listing_entries:
        time.sleep(DELAY_BETWEEN_REQUESTS)
        print(f"    取得教案彙整頁 id={stub['entry_id']}...")
        resp = get_with_retry(session, stub["detail_url"])
        tree = lxhtml.fromstring(resp.content)

        # 彙整頁海報
        poster_src = tree.xpath('//img[@id="ctl00_SiteContent_ExhibitionFigure"]/@src')
        poster_url = urljoin(BASE_URL + "/", poster_src[0]) if poster_src else None

        rows = parse_jiaoan_aggregate(tree)
        for row in rows:
            if "國小" not in row.get("school_level", ""):
                continue
            results.append({
                "entry_id":          f"{stub['entry_id']}_jiaoan",
                "competition_id":    comp_id,
                "year":              comp_info["year"],
                "competition_name":  comp_info["name"],
                "award_level":       row["award_level"],
                "title":             row["title"],
                "group":             "國小",
                "team_name":         "",
                "description":       "",
                "school_level":      row["school_level"],
                "poster_url":        poster_url,
                "thumb_url":         stub.get("thumb_url"),
                "poster_local":      "",
                "thumb_local":       "",
                "has_download_note": row["has_download_note"],
                "external_links":    "",
                "detail_url":        stub["detail_url"],
            })

    print(f"    找到 {len(results)} 筆國小教案條目")
    return results

# ─────────────────────────────────────────────
# 圖文競賽爬取
# ─────────────────────────────────────────────

def scrape_image_competition(session: requests.Session, comp_id: int,
                              comp_info: dict, state: dict,
                              guoxiao_only: bool = True) -> list[dict]:
    """爬取圖文競賽（comp_id=1-4），過濾國小並下載圖片。"""
    listing_entries = scrape_competition_listing(session, comp_id)
    results = []

    for stub in listing_entries:
        entry_id = stub["entry_id"]
        if str(entry_id) in state.get("scraped_entry_ids", []):
            print(f"    跳過 id={entry_id}（已爬）")
            continue

        time.sleep(DELAY_BETWEEN_REQUESTS)
        print(f"    取得詳細頁 id={entry_id}...")

        try:
            resp = get_with_retry(session, stub["detail_url"])
            tree = lxhtml.fromstring(resp.content)
            detail = parse_detail_page(tree, stub["detail_url"])
        except Exception as e:
            print(f"    !! 詳細頁失敗 id={entry_id}: {e}")
            continue

        # comp_id=3 全為國小高年級組，有些組別欄位可能為空
        if guoxiao_only and comp_id != 3 and not is_guoxiao(detail):
            state.setdefault("scraped_entry_ids", []).append(str(entry_id))
            continue

        # 建立下載目錄
        safe_title   = sanitize_filename(detail.get("title") or detail.get("title_raw", ""))
        entry_subdir = (OUTPUT_DIR / str(comp_info["year"])
                        / f"competition_{comp_id}_{comp_info['type']}"
                        / f"{entry_id}_{safe_title}")
        entry_subdir.mkdir(parents=True, exist_ok=True)

        poster_url = detail.get("poster_url")
        thumb_url  = stub.get("thumb_url")

        poster_local = ""
        thumb_local  = ""

        if poster_url:
            ext = Path(poster_url.split("?")[0]).suffix or ".jpg"
            poster_path = entry_subdir / f"poster{ext}"
            if download_image(session, poster_url, poster_path):
                print(f"      下載海報：{poster_path.name}")
            poster_local = str(poster_path)
            time.sleep(0.5)

        if thumb_url:
            thumb_path = entry_subdir / "thumb.jpg"
            if download_image(session, thumb_url, thumb_path):
                print(f"      下載縮圖：{thumb_path.name}")
            thumb_local = str(thumb_path)
            time.sleep(0.3)

        row = {
            "entry_id":          entry_id,
            "competition_id":    comp_id,
            "year":              comp_info["year"],
            "competition_name":  comp_info["name"],
            "award_level":       detail.get("award_level", ""),
            "title":             detail.get("title", ""),
            "group":             detail.get("group", ""),
            "team_name":         detail.get("team_name", ""),
            "description":       detail.get("description", ""),
            "school_level":      "",
            "poster_url":        poster_url or "",
            "thumb_url":         thumb_url or "",
            "poster_local":      poster_local,
            "thumb_local":       thumb_local,
            "has_download_note": detail.get("has_download_note", False),
            "external_links":    detail.get("external_links", ""),
            "detail_url":        stub["detail_url"],
        }
        append_metadata_row(METADATA_FILE, row)
        results.append(row)

        state.setdefault("scraped_entry_ids", []).append(str(entry_id))
        save_state(state)

    return results

# ─────────────────────────────────────────────
# 圖片下載
# ─────────────────────────────────────────────

def download_image(session: requests.Session, url: str, dest_path: Path) -> bool:
    """
    串流下載圖片至 dest_path。
    若已存在且非零大小則跳過（支援續跑）。
    回傳 True 表示新下載，False 表示跳過。
    """
    if dest_path.exists() and dest_path.stat().st_size > 0:
        return False
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        resp = session.get(url, timeout=REQUEST_TIMEOUT, stream=True)
        resp.raise_for_status()
        with open(dest_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                f.write(chunk)
        return True
    except requests.RequestException as e:
        print(f"      !! 下載失敗 {url}: {e}")
        return False

# ─────────────────────────────────────────────
# 工具函數
# ─────────────────────────────────────────────

def sanitize_filename(name: str, max_len: int = 50) -> str:
    """移除 Windows 檔名不合法字元。"""
    for ch in r'<>:"/\|?*':
        name = name.replace(ch, "_")
    name = re.sub(r"\s+", "_", name)
    return name[:max_len]


def load_state() -> dict:
    if STATE_FILE.exists():
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {"scraped_entry_ids": [], "completed_comp_ids": []}


def save_state(state: dict):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def write_metadata_header():
    """若 metadata.csv 不存在，建立並寫入欄位標頭。"""
    if not METADATA_FILE.exists():
        with open(METADATA_FILE, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=METADATA_FIELDS)
            writer.writeheader()


def append_metadata_row(csv_path: Path, row: dict):
    with open(csv_path, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=METADATA_FIELDS, extrasaction="ignore")
        writer.writerow(row)

# ─────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="eliteracy.edu.tw 得獎作品下載器")
    parser.add_argument("--all-groups", action="store_true",
                        help="下載全部組別（不限國小）")
    parser.add_argument("--comp-ids", nargs="+", type=int,
                        default=list(COMPETITION_MAP.keys()),
                        help="指定要爬的競賽 ID（預設全部 1-7）")
    args = parser.parse_args()

    guoxiao_only = not args.all_groups

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    write_metadata_header()

    state   = load_state()
    session = make_session()

    print(f"開始爬取，競賽 IDs: {args.comp_ids}")
    print(f"{'僅國小' if guoxiao_only else '全部組別'}")
    print(f"輸出目錄：{OUTPUT_DIR}")
    print(f"Metadata：{METADATA_FILE}\n")

    total_guoxiao = 0

    for comp_id in args.comp_ids:
        if comp_id not in COMPETITION_MAP:
            print(f"未知競賽 ID: {comp_id}，跳過")
            continue

        if comp_id in state.get("completed_comp_ids", []):
            print(f"[Competition {comp_id}] 已完成，跳過")
            continue

        comp_info = {**COMPETITION_MAP[comp_id], "id": comp_id}
        print(f"\n{'='*60}")
        print(f"[Competition {comp_id}] {comp_info['name']}（{comp_info['year']}年）")
        print(f"{'='*60}")

        try:
            if comp_info.get("aggregate"):
                # 教案彙整頁
                results = scrape_jiaoan_competition(session, comp_id, comp_info)
                for row in results:
                    # 下載彙整頁海報圖片
                    if row.get("poster_url"):
                        poster_path = (OUTPUT_DIR / str(comp_info["year"])
                                       / f"competition_{comp_id}_jiaoan"
                                       / f"{row['entry_id']}_poster.jpg")
                        if download_image(session, row["poster_url"], poster_path):
                            print(f"      下載教案海報：{poster_path.name}")
                        row["poster_local"] = str(poster_path)
                    append_metadata_row(METADATA_FILE, row)
                total_guoxiao += len(results)
            else:
                # 圖文競賽
                results = scrape_image_competition(
                    session, comp_id, comp_info, state, guoxiao_only=guoxiao_only
                )
                total_guoxiao += len(results)

            state.setdefault("completed_comp_ids", []).append(comp_id)
            save_state(state)
            print(f"\n  [Competition {comp_id}] 完成，本場次國小作品：{len(results)} 件")

        except KeyboardInterrupt:
            print("\n\n中斷。狀態已儲存，下次執行將自動從斷點繼續。")
            save_state(state)
            break
        except Exception as e:
            print(f"\n  [Competition {comp_id}] 發生錯誤：{e}，繼續下一場次")
            save_state(state)

    print(f"\n{'='*60}")
    print(f"全部完成！共找到國小得獎作品：{total_guoxiao} 件")
    print(f"Metadata CSV：{METADATA_FILE}")
    print(f"下載目錄：{OUTPUT_DIR}")


if __name__ == "__main__":
    main()
