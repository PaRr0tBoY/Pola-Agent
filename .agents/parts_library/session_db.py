"""
会话持久化与工具调用统计

SQLite 表结构：
  sessions     — 会话元数据
  messages     — 消息记录（JSON 序列化）
  tool_metrics — 每次工具调用的详细记录
  tool_summary — 按 (tool_name, session_id) 预聚合的统计视图
"""

import sqlite3, json, time, threading
from pathlib import Path
from datetime import datetime

DB_DIR = Path(__file__).parent
SESSION_DB_PATH = DB_DIR / "sessions.db"

# 线程安全的连接池（单连接 + WAL 足够）
_conn_lock = threading.Lock()


def _get_conn():
    # timeout=10：跨进程并发写时等待锁，避免立即 SQLITE_BUSY
    conn = sqlite3.connect(str(SESSION_DB_PATH), timeout=10.0)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn


def _init_tables():
    with _get_conn() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS sessions (
                id TEXT PRIMARY KEY,
                created_at TEXT NOT NULL DEFAULT (datetime('now','localtime')),
                last_active TEXT NOT NULL DEFAULT (datetime('now','localtime')),
                message_count INTEGER NOT NULL DEFAULT 0,
                tool_call_count INTEGER NOT NULL DEFAULT 0,
                status TEXT NOT NULL DEFAULT 'active'
            );

            CREATE TABLE IF NOT EXISTS messages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL REFERENCES sessions(id),
                seq INTEGER NOT NULL,
                role TEXT NOT NULL,
                content_type TEXT NOT NULL DEFAULT 'text',
                content_json TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            );
            CREATE INDEX IF NOT EXISTS idx_msg_session ON messages(session_id, seq);

            CREATE TABLE IF NOT EXISTS tool_metrics (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL REFERENCES sessions(id),
                tool_name TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'ok',
                duration_ms REAL,
                input_json TEXT DEFAULT '{}',
                output_len INTEGER DEFAULT 0,
                error_msg TEXT DEFAULT '',
                created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            );
            CREATE INDEX IF NOT EXISTS idx_tm_session ON tool_metrics(session_id);
            CREATE INDEX IF NOT EXISTS idx_tm_tool ON tool_metrics(tool_name);

            CREATE TABLE IF NOT EXISTS security_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL REFERENCES sessions(id),
                event_type TEXT NOT NULL,
                severity TEXT NOT NULL DEFAULT 'info',
                tool_name TEXT DEFAULT '',
                command TEXT DEFAULT '',
                action TEXT DEFAULT '',
                detail TEXT DEFAULT '',
                created_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            );
            CREATE INDEX IF NOT EXISTS idx_se_session ON security_events(session_id);
        """)
        conn.commit()


_tables_initialized = False


def _ensure_init():
    global _tables_initialized
    if not _tables_initialized:
        _init_tables()
        _tables_initialized = True


# ============================================================
# 会话操作
# ============================================================

def start_session(session_id=None, resume=False):
    """开始或恢复一个会话。返回 session_id 和恢复的消息列表。"""
    _ensure_init()
    if not session_id:
        session_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

    with _get_conn() as conn:
        existing = conn.execute(
            "SELECT id FROM sessions WHERE id = ?", (session_id,)
        ).fetchone()

        if resume and existing:
            conn.execute(
                "UPDATE sessions SET last_active = datetime('now','localtime') WHERE id = ?",
                (session_id,),
            )
            messages = _load_messages(conn, session_id)
            conn.commit()
            return session_id, messages

        if resume and not existing:
            resume = False

        conn.execute(
            "INSERT OR REPLACE INTO sessions (id) VALUES (?)",
            (session_id,),
        )
        conn.commit()

    return session_id, [] if not resume else []


def _load_messages(conn, session_id):
    """从数据库恢复消息列表（转换为 Anthropic API 格式）。"""
    rows = conn.execute(
        "SELECT role, content_type, content_json FROM messages "
        "WHERE session_id = ? ORDER BY seq",
        (session_id,),
    ).fetchall()

    history = []
    for role, ctype, cjson in rows:
        content = json.loads(cjson)
        history.append({"role": role, "content": content})
    return history


def _clean_surrogates(obj):
    """递归清理 surrogate 字符（U+D800-U+DFFF），它们不是合法 Unicode 无法 UTF-8 编码。"""
    if isinstance(obj, str):
        return obj.encode('utf-8', errors='surrogatepass').decode('utf-8', errors='replace')
    if isinstance(obj, list):
        return [_clean_surrogates(item) for item in obj]
    if isinstance(obj, dict):
        return {_clean_surrogates(k): _clean_surrogates(v) for k, v in obj.items()}
    return obj


def _clean_surrogates(obj):
    """递归清理 surrogate 字符（U+D800-U+DFFF），它们不是合法 Unicode 无法 UTF-8 编码。"""
    if isinstance(obj, str):
        return obj.encode('utf-8', errors='surrogatepass').decode('utf-8', errors='replace')
    if isinstance(obj, list):
        return [_clean_surrogates(item) for item in obj]
    if isinstance(obj, dict):
        return {_clean_surrogates(k): _clean_surrogates(v) for k, v in obj.items()}
    return obj


def persist_messages(session_id, history, start_seq=0):
    """持久化消息历史到数据库。增量写入（只写新消息）。
    自动剔除末尾不完整的 tool_use 消息（无对应 tool_result）。
    """
    _ensure_init()

    # 清理末尾：如果最后一条 assistant 消息全都是 tool_use（无 text/thinking），
    # 说明工具尚未执行就被中断，剔除它防止 API 报 tool_use/tool_result 不配对
    clean_history = list(history)
    if clean_history:
        last = clean_history[-1]
        if last.get("role") == "assistant":
            content = last.get("content", [])
            if isinstance(content, list) and content:
                all_tool_use = all(
                    (isinstance(b, dict) and b.get("type") == "tool_use") or
                    (hasattr(b, "type") and getattr(b, "type", None) == "tool_use")
                    for b in content
                )
                if all_tool_use:
                    clean_history.pop()

    with _get_conn() as conn:
        existing_count = conn.execute(
            "SELECT COUNT(*) FROM messages WHERE session_id = ?",
            (session_id,),
        ).fetchone()[0]

        for i, msg in enumerate(clean_history):
            seq = i + start_seq
            if seq < existing_count:
                continue

            role = msg["role"]
            content = msg["content"]
            ctype = "text" if isinstance(content, str) else "list"

            # JSON 序列化：Anthropic SDK 的 ThinkingBlock/TextBlock 等对象需先转 dict
            if isinstance(content, list):
                safe = []
                for item in content:
                    if isinstance(item, dict):
                        safe.append(item)
                    elif hasattr(item, "model_dump"):
                        safe.append(item.model_dump())
                    elif hasattr(item, "to_dict"):
                        safe.append(item.to_dict())
                    else:
                        safe.append(str(item))
                content = safe
            # 清理 surrogate 字符后再 JSON 序列化（否则 UTF-8 encode 会崩溃）
            content = _clean_surrogates(content)
            content = _clean_surrogates(content)
            cjson = json.dumps(content, ensure_ascii=False)

            conn.execute(
                "INSERT INTO messages (session_id, seq, role, content_type, content_json) "
                "VALUES (?, ?, ?, ?, ?)",
                (session_id, seq, role, ctype, cjson),
            )

        conn.execute(
            "UPDATE sessions SET message_count = ?, last_active = datetime('now','localtime') "
            "WHERE id = ?",
            (len(clean_history), session_id),
        )
        conn.commit()


def end_session(session_id, status="completed"):
    """结束会话（标记状态）。"""
    _ensure_init()
    with _get_conn() as conn:
        conn.execute(
            "UPDATE sessions SET status = ?, last_active = datetime('now','localtime') "
            "WHERE id = ?",
            (status, session_id),
        )
        conn.commit()


def list_sessions(limit=10):
    """列出最近的会话。"""
    _ensure_init()
    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT id, created_at, last_active, message_count, "
            "tool_call_count, status "
            "FROM sessions ORDER BY last_active DESC LIMIT ?",
            (limit,),
        ).fetchall()
    return [
        {
            "id": r[0], "created": r[1], "last_active": r[2],
            "messages": r[3], "tool_calls": r[4], "status": r[5],
        }
        for r in rows
    ]


# ============================================================
# 工具调用统计
# ============================================================

def log_tool_call(session_id, tool_name, status="ok",
                  duration_ms=0.0, input_dict=None,
                  output_len=0, error_msg=""):
    """记录一次工具调用。由 PostToolUse Hook 调用。"""
    _ensure_init()
    input_json = json.dumps(input_dict or {}, ensure_ascii=False, default=str)
    with _get_conn() as conn:
        conn.execute(
            "INSERT INTO tool_metrics "
            "(session_id, tool_name, status, duration_ms, input_json, output_len, error_msg) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (session_id, tool_name, status, duration_ms,
             input_json, output_len, error_msg),
        )
        conn.execute(
            "UPDATE sessions SET tool_call_count = tool_call_count + 1, "
            "last_active = datetime('now','localtime') WHERE id = ?",
            (session_id,),
        )
        conn.commit()


def log_security_event(session_id, event_type, severity="info",
                       tool_name="", command="", action="", detail=""):
    """记录安全事件（权限拦截、二次确认等）。"""
    _ensure_init()
    with _get_conn() as conn:
        conn.execute(
            "INSERT INTO security_events "
            "(session_id, event_type, severity, tool_name, command, action, detail) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (session_id, event_type, severity, tool_name, command, action, detail),
        )
        conn.commit()


# ============================================================
# 统计查询
# ============================================================

def tool_stats(session_id=None, limit=20):
    """工具调用统计：次数、成功率、平均耗时。
    如果指定 session_id，只统计该会话；
    否则统计所有会话。
    """
    _ensure_init()
    where = "WHERE session_id = ?" if session_id else ""
    params = (session_id,) if session_id else ()

    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT tool_name, "
            "COUNT(*) AS total, "
            "SUM(CASE WHEN status = 'ok' THEN 1 ELSE 0 END) AS ok_count, "
            "ROUND(AVG(CASE WHEN status = 'ok' THEN duration_ms END), 1) AS avg_ms, "
            "MAX(duration_ms) AS max_ms "
            f"FROM tool_metrics {where} "
            "GROUP BY tool_name ORDER BY total DESC LIMIT ?",
            params + (limit,),
        ).fetchall()

    return [
        {
            "tool": r[0], "total": r[1], "ok": r[2],
            "fail": r[1] - r[2],
            "success_rate": f"{r[2]/r[1]*100:.1f}%" if r[1] else "N/A",
            "avg_ms": r[3] or 0, "max_ms": r[4] or 0,
        }
        for r in rows
    ]


def tool_timeline(session_id=None, limit=50):
    """按时间排列的最近工具调用。"""
    _ensure_init()
    where = "WHERE session_id = ?" if session_id else ""
    params = (session_id,) if session_id else ()

    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT tool_name, status, duration_ms, output_len, "
            "error_msg, created_at "
            f"FROM tool_metrics {where} "
            "ORDER BY id DESC LIMIT ?",
            params + (limit,),
        ).fetchall()

    return [
        {
            "tool": r[0], "status": r[1], "ms": r[2],
            "output_bytes": r[3], "error": r[4][:80] if r[4] else "",
            "time": r[5],
        }
        for r in rows
    ]


def security_stats(session_id=None):
    """安全事件统计。"""
    _ensure_init()
    where = "WHERE session_id = ?" if session_id else ""
    params = (session_id,) if session_id else ()

    with _get_conn() as conn:
        rows = conn.execute(
            "SELECT event_type, severity, COUNT(*) "
            f"FROM security_events {where} "
            "GROUP BY event_type, severity ORDER BY COUNT(*) DESC",
            params,
        ).fetchall()

    return [
        {"type": r[0], "severity": r[1], "count": r[2]}
        for r in rows
    ]


def session_summary(session_id):
    """单个会话的综合摘要。"""
    _ensure_init()
    with _get_conn() as conn:
        session = conn.execute(
            "SELECT * FROM sessions WHERE id = ?", (session_id,)
        ).fetchone()
        if not session:
            return None

        total_tools = conn.execute(
            "SELECT COUNT(*) FROM tool_metrics WHERE session_id = ?",
            (session_id,),
        ).fetchone()[0]

        total_ok = conn.execute(
            "SELECT COUNT(*) FROM tool_metrics WHERE session_id = ? AND status = 'ok'",
            (session_id,),
        ).fetchone()[0]

        total_sec = conn.execute(
            "SELECT COUNT(*) FROM security_events WHERE session_id = ?",
            (session_id,),
        ).fetchone()[0]

        avg_duration = conn.execute(
            "SELECT ROUND(AVG(duration_ms), 1) FROM tool_metrics "
            "WHERE session_id = ? AND status = 'ok'",
            (session_id,),
        ).fetchone()[0] or 0

    return {
        "id": session[0],
        "created": session[1],
        "last_active": session[2],
        "messages": session[3],
        "total_tool_calls": total_tools,
        "ok_tool_calls": total_ok,
        "fail_tool_calls": total_tools - total_ok,
        "success_rate": f"{total_ok / total_tools * 100:.1f}%" if total_tools else "N/A",
        "avg_duration_ms": avg_duration,
        "security_events": total_sec,
        "status": session[5],
    }
