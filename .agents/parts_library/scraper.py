"""
标准件在线下载器

从 3D ContentCentral、MISUMI 中国、怡合达等平台搜索标准件并建立本地索引。
所有下载的 STEP 文件存入 .agents/parts_library/step_files/ 目录。
"""

import os, time, json, re, urllib.request, urllib.parse
from pathlib import Path
from .parts_db import PartsDB, STEP_DIR, STANDARDS_DIR

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 Chrome/125.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/json,*/*",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
}


def _safe_request(url, timeout=30, headers=None, data=None, method="GET"):
    req = urllib.request.Request(url, data=data,
                                 headers=headers or HEADERS, method=method)
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read()
        except Exception:
            if attempt == 2:
                raise
            time.sleep(1 * (attempt + 1))
    return None


def download_step_file(url, filename, category=""):
    """下载单个 STEP 文件到本地。"""
    cat_dir = STEP_DIR / category if category else STEP_DIR
    cat_dir.mkdir(parents=True, exist_ok=True)
    filepath = cat_dir / filename
    if filepath.exists():
        return str(filepath)
    try:
        data = _safe_request(url)
        if data and len(data) > 1000:
            filepath.write_bytes(data)
            return str(filepath)
    except Exception:
        pass
    return None


# ============================================================
# 3D ContentCentral
# ============================================================

CONTENTCENTRAL_SEARCH = "https://www.3dcontentcentral.com/Search.aspx"


def scrape_contentcentral(keyword, category="", max_parts=10):
    """从 3D ContentCentral 搜索标准件并索引。"""
    db = PartsDB()
    found = 0
    try:
        params = urllib.parse.urlencode({"arg": keyword, "ext": "all"})
        url = f"{CONTENTCENTRAL_SEARCH}?{params}"
        html = _safe_request(url)
        if not html:
            return found
        text = html.decode("utf-8", errors="replace")
        pattern = re.compile(
            r'class="result-title"[^>]*href="([^"]*)"[^>]*>([^<]*)</a>',
            re.IGNORECASE,
        )
        for href, name in pattern.findall(text)[:max_parts]:
            name = name.strip()
            if not name:
                continue
            detail_url = (
                f"https://www.3dcontentcentral.com{href}"
                if href.startswith("/") else href
            )
            db.add_part(
                source="3dcontentcentral",
                category=category or keyword,
                description=name,
                file_url=detail_url,
                tags=f"{keyword},{name}",
            )
            found += 1
        db.log_scraper_run("3dcontentcentral", "success", found)
    except Exception as e:
        db.log_scraper_run("3dcontentcentral", "error", error_msg=str(e))
    return found


# ============================================================
# MISUMI 中国
# ============================================================

MISUMI_API = "https://www.misumi.com.cn/api"


def scrape_misumi(keyword, category="", max_parts=20):
    """从 MISUMI 中国搜索标准件。"""
    db = PartsDB()
    found = 0
    try:
        search_url = f"{MISUMI_API}/product/search"
        params = json.dumps({
            "keyword": keyword, "page": 1,
            "pageSize": min(max_parts, 20), "sort": "relevance",
        }).encode("utf-8")
        headers = {**HEADERS, "Content-Type": "application/json"}
        resp = _safe_request(search_url, data=params, headers=headers, method="POST")
        if not resp:
            return found
        data = json.loads(resp.decode("utf-8", errors="replace"))
        items = data.get("data", {}).get("items", []) if isinstance(data, dict) else []

        for item in items:
            part_no = item.get("partNumber", "")
            name = item.get("partName", "")
            brand = item.get("brandName", "")
            if not part_no:
                continue

            # 尝试下载 STEP
            cad_files = item.get("cadFiles", [])
            step_file = next(
                (f for f in cad_files if "step" in f.get("format", "").lower()),
                None,
            )
            local_path = ""
            if step_file and step_file.get("url", "").startswith("http"):
                local_path = download_step_file(
                    step_file["url"], f"misumi_{part_no}.step", category or keyword
                ) or ""

            db.add_part(
                source="misumi",
                standard_no=part_no,
                category=category or keyword,
                spec=part_no,
                description=f"{brand} {name} ({part_no})",
                file_path=local_path,
                file_url=step_file.get("url", "") if step_file else "",
                tags=f"{keyword},{brand},{name},{part_no}",
            )
            found += 1

        db.log_scraper_run("misumi", "success", found)
    except Exception as e:
        db.log_scraper_run("misumi", "error", error_msg=str(e))
    return found


# ============================================================
# 怡合达
# ============================================================

YIHEDA_SEARCH = "https://www.yiheda.com/search"


def scrape_yiheda(keyword, category="", max_parts=20):
    """从怡合达搜索标准件。"""
    db = PartsDB()
    found = 0
    try:
        params = urllib.parse.urlencode({"keyword": keyword, "page": 1})
        url = f"{YIHEDA_SEARCH}?{params}"
        html = _safe_request(url)
        if not html:
            return found
        text = html.decode("utf-8", errors="replace")

        # 尝试多种解析模式
        for pat in [
            r'class="[^"]*product[^"]*"[^>]*>.*?<a[^>]*href="([^"]*)"[^>]*>([^<]*)</a>',
            r'href="(/product[^"]*)"[^>]*>([^<]{3,100})</a>',
            r'<a[^>]*href="(/[^"]*)"[^>]*class="[^"]*title[^"]*"[^>]*>([^<]+)</a>',
        ]:
            matches = re.findall(pat, text, re.IGNORECASE)
            if matches:
                break

        for href, name in matches[:max_parts]:
            name = name.strip()
            if not name or len(name) < 3:
                continue
            detail_url = (
                f"https://www.yiheda.com{href}"
                if href.startswith("/") else href
            )
            db.add_part(
                source="yiheda",
                category=category or keyword,
                description=name,
                file_url=detail_url,
                tags=f"{keyword},{name}",
            )
            found += 1
        db.log_scraper_run("yiheda", "success", found)
    except Exception as e:
        db.log_scraper_run("yiheda", "error", error_msg=str(e))
    return found


# ============================================================
# 批量下载
# ============================================================

MANUAL_DOWNLOAD_GUIDE = """
=== 标准件手动下载指南 ===

在线平台需要浏览 JavaScript 渲染页面，请手动下载并用 sw_import_step 导入。
下载的 STEP 文件放入 .agents/parts_library/step_files/<类别>/ 目录即可被自动索引。

常用平台：

1. 3D ContentCentral (全球最大 CAD 模型库)
   网址: https://www.3dcontentcentral.com
   操作: 搜索 -> 选择供应商/标准 -> 下载 STEP (.stp)

2. MISUMI 中国 (日系标准件)
   网址: https://www.misumi.com.cn
   操作: 搜索型号 -> 产品页 -> [CAD下载] -> 选择 STEP 格式

3. 怡合达 (国产自动化零部件)
   网址: https://www.yiheda.com
   操作: 搜索 -> 产品详情 -> [下载CAD] -> STEP (.stp)

4. TraceParts (工业零部件库)
   网址: https://www.traceparts.com
   操作: 搜索 -> 选择格式 -> 下载 STEP

5. GrabCAD (社区共享库)
   网址: https://grabcad.com/library
   操作: 搜索 -> 下载 STEP/IGES

常用搜索关键词: 螺栓 bolt, 螺母 nut, 垫圈 washer, 轴承 bearing,
                 齿轮 gear, 键 key, 销 pin, 挡圈 ring, 弹簧 spring,
                 法兰 flange, 导轨 linear guide, 丝杆 ball screw,
                 联轴器 coupling, 同步带轮 timing pulley
"""


def get_download_instructions():
    return MANUAL_DOWNLOAD_GUIDE


PREDEFINED_SEARCHES = [
    ("misumi", "六角螺栓", "bolt"),
    ("misumi", "六角螺母", "nut"),
    ("misumi", "平垫圈", "washer"),
    ("misumi", "深沟球轴承", "bearing"),
    ("yiheda", "直线导轨", "profile"),
    ("yiheda", "滚珠丝杆", "shaft"),
    ("yiheda", "联轴器", "shaft"),
    ("yiheda", "同步带轮", "gear"),
    ("contentcentral", "step motor nema", "shaft"),
    ("contentcentral", "flange bearing", "bearing"),
]


def run_all_scrapers():
    """批量运行所有下载器。"""
    from .parts_db import seed_common_parts

    results = {}
    n = seed_common_parts()
    results["seed"] = n

    scrapers = {
        "misumi": scrape_misumi,
        "yiheda": scrape_yiheda,
        "contentcentral": scrape_contentcentral,
    }

    for source, keyword, category in PREDEFINED_SEARCHES:
        fn = scrapers.get(source)
        if not fn:
            continue
        print(f"[{source}] 搜索: {keyword} ({category}) ...")
        try:
            count = fn(keyword, category=category, max_parts=10)
            key = f"{source}:{keyword}"
            results[key] = count
            print(f"  -> 找到 {count} 个")
            time.sleep(1)
        except Exception as e:
            print(f"  -> 失败: {e}")
            results[key] = f"error: {e}"

    return results


if __name__ == "__main__":
    print("=== Pola 标准件库下载器 ===\n")
    results = run_all_scrapers()
    print("\n=== 统计 ===")
    for k, v in results.items():
        print(f"  {k}: {v}")
    db = PartsDB()
    s = db.stats()
    print(f"\n标准件总数: {s['total_parts']}")
    print("按来源:", s["by_source"])
    print("按类别:", s["by_category"])
