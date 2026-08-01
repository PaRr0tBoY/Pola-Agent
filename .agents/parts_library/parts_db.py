"""
标准件本地索引数据库

使用 SQLite 管理标准件目录，支持按标准号、类别、规格搜索。
数据库文件存储在 .agents/parts_library/parts.db
下载的 STEP 文件存储在 .agents/parts_library/step_files/
"""

import sqlite3
import os
from pathlib import Path

STANDARDS_DIR = Path(__file__).parent
DB_PATH = STANDARDS_DIR / "parts.db"
STEP_DIR = STANDARDS_DIR / "step_files"

# 常用标准件类别及对应的 SolidWorks Toolbox 路径前缀
CATEGORY_MAP = {
    "bolt": "螺栓",
    "nut": "螺母",
    "washer": "垫圈",
    "bearing": "轴承",
    "gear": "齿轮",
    "key": "键",
    "pin": "销",
    "ring": "挡圈",
    "rivet": "铆钉",
    "spring": "弹簧",
    "flange": "法兰",
    "pipe": "管件",
    "profile": "型材",
    "seal": "密封件",
    "shaft": "轴",
}

STANDARD_SYSTEMS = ["GB", "ISO", "ANSI", "DIN", "JIS"]


class PartsDB:
    """标准件库 SQLite 管理器。"""

    def __init__(self, db_path=None):
        self.db_path = str(db_path or DB_PATH)
        STEP_DIR.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("""
                CREATE TABLE IF NOT EXISTS standard_parts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    source TEXT NOT NULL DEFAULT 'manual',
                    standard_system TEXT NOT NULL DEFAULT '',
                    standard_no TEXT NOT NULL DEFAULT '',
                    category TEXT NOT NULL DEFAULT '',
                    spec TEXT NOT NULL DEFAULT '',
                    description TEXT NOT NULL DEFAULT '',
                    material TEXT DEFAULT '',
                    file_path TEXT DEFAULT '',
                    file_url TEXT DEFAULT '',
                    toolbox_ref TEXT DEFAULT '',
                    tags TEXT DEFAULT '',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_parts_category
                    ON standard_parts(category)
            """)
            conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_parts_standard
                    ON standard_parts(standard_system, standard_no)
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS scraper_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    source TEXT NOT NULL,
                    status TEXT NOT NULL,
                    parts_found INTEGER DEFAULT 0,
                    error_msg TEXT DEFAULT '',
                    started_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    finished_at TIMESTAMP
                )
            """)
            conn.commit()

    def add_part(self, **kwargs):
        """添加一个标准件记录。"""
        allowed = [
            "source", "standard_system", "standard_no", "category",
            "spec", "description", "material", "file_path",
            "file_url", "toolbox_ref", "tags",
        ]
        fields = {k: kwargs.get(k, "") for k in allowed}
        columns = ", ".join(fields.keys())
        placeholders = ", ".join("?" for _ in fields)
        values = list(fields.values())
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                f"INSERT INTO standard_parts ({columns}) VALUES ({placeholders})",
                values,
            )
            conn.commit()

    def search(self, keyword="", category="", standard_system="",
               standard_no="", limit=20):
        """搜索标准件库。支持多条件组合过滤。"""
        conditions = []
        params = []
        if keyword:
            conditions.append(
                "(description LIKE ? OR spec LIKE ? OR tags LIKE ? "
                "OR standard_no LIKE ? OR category LIKE ? OR source LIKE ?)"
            )
            kw = f"%{keyword}%"
            params.extend([kw, kw, kw, kw, kw, kw])
        if category:
            conditions.append("category = ?")
            params.append(category)
        if standard_system:
            conditions.append("standard_system = ?")
            params.append(standard_system.upper())
        if standard_no:
            conditions.append("standard_no LIKE ?")
            params.append(f"%{standard_no}%")

        where = ("WHERE " + " AND ".join(conditions)) if conditions else ""
        query = (
            "SELECT id, source, standard_system, standard_no, "
            "category, spec, description, material, file_path, "
            "file_url, toolbox_ref, tags "
            f"FROM standard_parts {where} ORDER BY category, standard_no "
            f"LIMIT {int(limit)}"
        )
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(query, params).fetchall()
        return [dict(r) for r in rows]

    def count_by_category(self):
        """按类别统计零件数量。"""
        with sqlite3.connect(self.db_path) as conn:
            rows = conn.execute(
                "SELECT category, COUNT(*) as cnt "
                "FROM standard_parts GROUP BY category ORDER BY cnt DESC"
            ).fetchall()
        return [(r[0], r[1]) for r in rows]

    def stats(self):
        """数据库统计信息。"""
        with sqlite3.connect(self.db_path) as conn:
            total = conn.execute(
                "SELECT COUNT(*) FROM standard_parts"
            ).fetchone()[0]
            sources = conn.execute(
                "SELECT source, COUNT(*) FROM standard_parts "
                "GROUP BY source"
            ).fetchall()
            categories = conn.execute(
                "SELECT category, COUNT(*) FROM standard_parts "
                "GROUP BY category ORDER BY COUNT(*) DESC"
            ).fetchall()
        return {
            "total_parts": total,
            "by_source": dict(sources),
            "by_category": dict(categories),
        }

    def log_scraper_run(self, source, status, parts_found=0, error_msg=""):
        with sqlite3.connect(self.db_path) as conn:
            conn.execute(
                "INSERT INTO scraper_log (source, status, parts_found, error_msg) "
                "VALUES (?, ?, ?, ?)",
                (source, status, parts_found, error_msg),
            )
            conn.commit()


# 启动时自动种子一些常用 GB 标准件（无 STEP 文件，仅索引）
def seed_common_parts():
    """种子数据：常用中国 GB 标准件索引。"""
    db = PartsDB()
    if db.stats()["total_parts"] > 0:
        return  # 已有数据，不重复种子

    seeds = [
        # 螺栓 GB/T 5782
        ("GB", "GB/T 5782", "bolt", "M6×20", "六角头螺栓 M6×20", "8.8级"),
        ("GB", "GB/T 5782", "bolt", "M8×30", "六角头螺栓 M8×30", "8.8级"),
        ("GB", "GB/T 5782", "bolt", "M10×40", "六角头螺栓 M10×40", "8.8级"),
        ("GB", "GB/T 5782", "bolt", "M12×50", "六角头螺栓 M12×50", "8.8级"),
        ("GB", "GB/T 5782", "bolt", "M16×60", "六角头螺栓 M16×60", "8.8级"),
        # 螺母 GB/T 6170
        ("GB", "GB/T 6170", "nut", "M6", "六角螺母 M6", "8级"),
        ("GB", "GB/T 6170", "nut", "M8", "六角螺母 M8", "8级"),
        ("GB", "GB/T 6170", "nut", "M10", "六角螺母 M10", "8级"),
        ("GB", "GB/T 6170", "nut", "M12", "六角螺母 M12", "8级"),
        # 平垫圈 GB/T 97.1
        ("GB", "GB/T 97.1", "washer", "6", "平垫圈 6mm", "200HV"),
        ("GB", "GB/T 97.1", "washer", "8", "平垫圈 8mm", "200HV"),
        ("GB", "GB/T 97.1", "washer", "10", "平垫圈 10mm", "200HV"),
        ("GB", "GB/T 97.1", "washer", "12", "平垫圈 12mm", "200HV"),
        # 滚动轴承 GB/T 276
        ("GB", "GB/T 276", "bearing", "6204", "深沟球轴承 6204 (20×47×14)", ""),
        ("GB", "GB/T 276", "bearing", "6205", "深沟球轴承 6205 (25×52×15)", ""),
        ("GB", "GB/T 276", "bearing", "6206", "深沟球轴承 6206 (30×62×16)", ""),
        # 平键 GB/T 1096
        ("GB", "GB/T 1096", "key", "8×7×20", "普通平键 8×7×20", ""),
        ("GB", "GB/T 1096", "key", "10×8×30", "普通平键 10×8×30", ""),
        ("GB", "GB/T 1096", "key", "12×8×40", "普通平键 12×8×40", ""),
        # 圆柱销 GB/T 119.1
        ("GB", "GB/T 119.1", "pin", "φ6×30", "圆柱销 φ6×30", ""),
        ("GB", "GB/T 119.1", "pin", "φ8×40", "圆柱销 φ8×40", ""),
    ]
    for std_sys, std_no, cat, spec, desc, mat in seeds:
        db.add_part(
            source="seed",
            standard_system=std_sys,
            standard_no=std_no,
            category=cat,
            spec=spec,
            description=desc,
            material=mat,
            tags=f"{CATEGORY_MAP.get(cat, '')},{spec}",
        )
    return len(seeds)
