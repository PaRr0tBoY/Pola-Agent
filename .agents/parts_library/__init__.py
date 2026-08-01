"""
标准件库模块

提供标准件的本地索引、搜索和在线下载能力。
数据来源：SolidWorks Toolbox / 3D ContentCentral / MISUMI / 怡合达
"""

from .parts_db import PartsDB, STANDARDS_DIR, seed_common_parts
from .scraper import (
    scrape_contentcentral,
    scrape_misumi,
    scrape_yiheda,
    download_step_file,
)

__all__ = [
    "PartsDB",
    "STANDARDS_DIR",
    "seed_common_parts",
    "scrape_contentcentral",
    "scrape_misumi",
    "scrape_yiheda",
    "download_step_file",
]
