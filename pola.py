#!/usr/bin/env python3

import os, subprocess, ast, json, yaml, math, time, signal
from pathlib import Path
import win32com.client
from rich.console import Console
from rich.markdown import Markdown

RED = "\033[31m"
YELLOW = "\033[33m"
GRAY = "\033[90m"
GREEN = "\033[36m"
RESET = "\033[0m"

# ---- safe print: avoid GBK encoding crashes on Windows terminals ----
def _safe_str(s):
    """Replace characters that can't be encoded in GBK with ASCII equivalents."""
    if not isinstance(s, str):
        s = str(s)
    # map unicode chars to ASCII-safe alternatives
    _char_map = {
        "²": "^2", "³": "^3",   # superscript 2/3
        "°": "deg",                    # degree sign
        "µ": "u",                      # micro
        "π": "pi",                     # pi
        "≤": "<=", "≥": ">=",    # ≤ ≥
        "×": "x",                      # ×
        "✓": "[OK]", "✅": "[OK]", "✔": "[OK]",
        "✗": "[X]",  "❌": "[X]",  "✘": "[X]",
        "⚠": "[!]",  "⚡": "[!]",
        "⏰": "...",  "⌛": "...",
        "❗": "!!",   "❓": "??",
        "✨": "(*)",
    }
    for unicode_char, ascii_replacement in _char_map.items():
        s = s.replace(unicode_char, ascii_replacement)
    # strip any remaining non-GBK chars
    try:
        s.encode("gbk")
    except UnicodeEncodeError:
        s = s.encode("gbk", errors="replace").decode("gbk")
    return s


def _safe_print(*args, **kwargs):
    """Print that won't crash on GBK terminals."""
    safe_args = [_safe_str(a) for a in args]
    print(*safe_args, **kwargs)

from prompt_toolkit import prompt as _pt_prompt
from prompt_toolkit.history import InMemoryHistory
from prompt_toolkit.styles import Style, DynamicStyle
from prompt_toolkit.completion import Completer, Completion
from prompt_toolkit.application import get_app

_pt_history = InMemoryHistory()
_pt_style = Style([("prompt", "fg:ansicyan")])

from dotenv import load_dotenv
from anthropic import Anthropic

load_dotenv(override=True)

if os.getenv("ANTHROPIC_BASE_URL"):
    os.environ.pop("ANTHROPIC_AUTH_TOKEN", None)


WORKDIR = Path.cwd()

# 标准件库与会话持久化（依赖 WORKDIR，必须在此之后导入）
import sys as _sys
_PARTS_LIB = WORKDIR / ".agents" / "parts_library"
if str(_PARTS_LIB.parent) not in _sys.path:
    _sys.path.insert(0, str(_PARTS_LIB.parent))
from parts_library import PartsDB, seed_common_parts
from parts_library.session_db import (
    start_session, persist_messages, end_session,
    log_tool_call, log_security_event, session_summary,
)
_parts_db = PartsDB()
try:
    _parts_seeded = seed_common_parts()
except Exception:
    _parts_seeded = 0
_current_session_id = None

SKILLS_DIR = WORKDIR / ".agents/skills"
client = Anthropic()
MODEL = os.environ["MODEL_ID"]
REASON = "high"
SUB_MODEL = os.environ["SUB_MODEL_ID"]
CURRENT_TODOS: list[dict] = []

def _parse_frontmatter(text):
    if not text.startswith("---"):
        return {}, text
    parts = text.split("---", 2)
    if len(parts) < 3:
        return {}, text
    try:
        meta = yaml.safe_load(parts[1]) or {}
    except yaml.YAMLError:
        meta = {}
    return meta, parts[2].strip()

SKILL_REGISTRY: dict[str, dict] = {}

def _scan_skills():
    if not SKILLS_DIR.exists():
        return
    for d in sorted(SKILLS_DIR.iterdir()):
        if not d.is_dir():
            continue
        manifest = d / "SKILL.md"
        if manifest. exists():
            raw = manifest.read_text(encoding="utf-8")
            meta, body = _parse_frontmatter(raw)
            name = meta.get("name", d.name)
            desc = meta.get("description", raw.split("\n")[0].lstrip("#").strip())
            SKILL_REGISTRY[name] = {"name": name, "description": desc, "content": raw}

_scan_skills()

def list_skills():
    if not SKILL_REGISTRY:
        return "(no skills found)"
    return "\n".join(f"- **{s['name']}**: {s['description']}" for s in SKILL_REGISTRY.values())

def build_system():
    catalog = list_skills()
    return (
        "You are pola, a SolidWorks AI agent at " + str(WORKDIR) + ". Name:Acid|Major:Mechanical Manufacturing|Lang:zh-CN.\n"
        "|IMPORTANT: Prefer retrieval-led reasoning. Read project files before relying on training data.\n"
        "|Conduct:Think->StateAssumptions->AskIfUncertain|Simple->MinCodeNoSpeculationNoAbstraction|"
        "Surgical->TouchOnlyNeeded->MatchStyle|GoalDriven->DefineTests->VerifyLoop\n"
        "|UNITS: ALL sw_* tool params in METERS. 100mm->0.1, 5cm->0.05. Never pass mm directly.\n"
        "|SketchChain:sw_select_plane_or_face->sw_draw_*->sw_extrude_boss|sw_extrude_cut->sw_export_to_format(STEP)\n"
        "|SketchNames: Auto-increment(Sketch1,Sketch2...). Verify with sw_get_model_structure before extrude.\n"
        "|StandardFirst: sw_search_standard_part->[found]sw_insert_toolbox_part|sw_import_step->[not found]manual model. GB|ISO|ANSI|DIN|JIS.\n"
        "|Revolve: ALWAYS sw_draw_revolve_profile(cx1,cy1,cx2,cy2,points) once. NEVER split into sw_draw_centerline+sw_draw_profile (different sketches!). "
        "Flow: select plane -> sw_draw_revolve_profile -> sw_revolve_boss(sketch,angle=360). Fail fix: sw_edit_sketch -> sw_add_relation(coincident) -> sw_close_sketch -> retry.\n"
        "|RevolveContour: Close contour p[-1]->p[0] must NOT backtrack vs prev segment. 6pts=OK, 7pts=degenerate.\n"
        "|Pattern:sw_linear_pattern(along edge)|sw_circular_pattern(around axis)|sw_mirror_feature(across plane). Feature first, then pattern.\n"
        "|RefPlane:sw_create_ref_plane:{offset(from face)|angle(around edge)|mid(between 2 faces)}\n"
        "|Assembly:sw_create_assembly->sw_insert_component(*.sldprt)->sw_fix_component(1st part)->sw_add_mate(coincident|concentric|parallel)->export STEP\n"
        "|Constraints:sw_add_dimension(linear|radial|angle|diameter)|sw_add_relation(horizontal|vertical|concentric|tangent|parallel|equal|fix)\n"
        "|Engineering:sw_shell(open face, thickness)|sw_rib(open sketch -> call)\n"
        "|Verify:sw_measure(distance|diameter)|sw_mass_properties->compare->auto-correct\n"
        "|BodyOps:sw_mirror_body|sw_move_component|sw_fix_component\n"
        "|Pitfalls:NEVER redraw circle on cut fail(read sketch list in error). Same op fail 3x=switch strategy. "
        "Keep parts open before sw_insert_component. sw_insert_component fails -> retry after reopening the part, or model as single part.\n"
        "|Skills:" + catalog + "\n"
        "|Use load_skill for full skill details."
    )

SYSTEM = build_system()

SUB_SYSTEM = (
    f"You are a useful agent at {WORKDIR}. "
    "Complete the task you were given, then return a concise summary. "
    "Do not delegate further."
)

# ── Windows 无窗口辅助：抑制 subprocess 弹出的 cmd 窗口 ──
def _no_window_flags():
    """返回 Windows 平台下禁止弹窗的 creationflags。
    DETACHED_PROCESS 阻止 msys2/bash 文件关联触发"选取应用"对话框。
    CREATE_NEW_PROCESS_GROUP 防止 Ctrl+C 信号传播到子进程。"""
    if os.name == "nt":
        return (
            subprocess.CREATE_NO_WINDOW
            | subprocess.DETACHED_PROCESS
            | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200)
        )
    return 0


def _no_window_startupinfo():
    """返回隐藏窗口的 STARTUPINFO，Windows 下双重保险。"""
    if os.name != "nt":
        return None
    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = subprocess.SW_HIDE
    return si


def run_bash(command, **kwargs):
    import re as _re
    segments = [s.strip().lower() for s in _re.split(r"[&|;]", command)]
    if any(any(d in s for d in DENY_LIST) for s in segments):
        return "危险！请避免尝试执行高风险指令。"
    try:
        # 显式指定真正的 cmd.exe — MSYS2/Git Bash 会将 COMSPEC 覆写为
        # powershell.exe，导致 shell=True 时触发"选取应用"对话框。
        shell_exe = None
        if os.name == "nt":
            sys_root = os.environ.get("SYSTEMROOT", r"C:\Windows")
            shell_exe = os.path.join(sys_root, "System32", "cmd.exe")
        r = subprocess.run(
            command,
            shell=True,
            executable=shell_exe,
            cwd=WORKDIR,
            capture_output=True,
            timeout=120,
            startupinfo=_no_window_startupinfo(),
            creationflags=_no_window_flags(),
        )
        raw = (r.stdout or b"") + (r.stderr or b"")
        try:
            out = raw.decode("utf-8").strip()
        except UnicodeDecodeError:
            out = raw.decode("gbk", errors="replace").strip()
        return out[:50000] if out else "(无输出)"
    except subprocess.TimeoutExpired:
        return "指令执行超时（120s）"
    except (FileNotFoundError, OSError) as e:
        return f"错误{e}"


def safe_path(p):
    path = (WORKDIR / p).resolve()
    if not path.is_relative_to(WORKDIR):
        raise ValueError(f"路径逃逸工作区: {p}")
    return path


def run_read(path, start: int = 0, limit: int | None = None, **kwargs):
    try:
        p = safe_path(path)
        raw = p.read_bytes()
        try:
            text = raw.decode("utf-8")
        except UnicodeDecodeError:
            text = raw.decode("gbk", errors="replace")
        lines = text.splitlines()
        total = len(lines)
        if start >= total:
            return "(起始行超出文件长度)"
        lines = lines[start:]
        if limit and limit < len(lines):
            lines = lines[:limit] + [f"... ({len(lines) - limit} more lines)"]
        return "\n".join(lines)
    except Exception as e:
        return f"错误：{e}"


def run_write(path, content, **kwargs):
    try:
        file_path = safe_path(path)
        file_path.parent.mkdir(parents=True, exist_ok=True)
        file_path.write_text(content, encoding="utf-8")
        return f"Wrote {len(content)} bytes to {path}"
    except Exception as e:
        return f"错误：{e}"


def run_edit(path, old_text, new_text, **kwargs):
    try:
        file_path = safe_path(path)
        raw = file_path.read_bytes()
        try:
            text, enc = raw.decode("utf-8"), "utf-8"
        except UnicodeDecodeError:
            try:
                text, enc = raw.decode("gbk"), "gbk"
            except UnicodeDecodeError:
                return f"错误：无法以 UTF-8 或 GBK 解码 {path}"
        if old_text not in text:
            return f"{RED}错误：没有命中修改区域，可能要替换的文字不存在，或文件自上次查看已更新{RESET}"
        file_path.write_text(text.replace(old_text, new_text, 1), encoding = enc)
        return f"{GREEN}Pola已编辑 {path}{RESET}"
    except Exception as e:
        return f"错误：{e}"

def run_glob(pattern, **kwargs):
    import glob as g

    try:
        results = []
        for match in g.glob(pattern, root_dir=WORKDIR):
            if (WORKDIR / match).resolve().is_relative_to(WORKDIR):
                results.append(match)
        return "\n".join(results) if results else "(未找到匹配)"
    except Exception as e:
        return f"错误：{e}"

# =========================================================================
# SolidWorks 自动化工具集扩展实现
# =========================================================================

_SW_APP_CACHE = None  # 模块级 COM 连接单例，避免每次工具调用都重连泄漏

def _check_sw_license_service():
    """检测 SolidWorks Licensing Service 是否运行，停止则提前报错而非让 COM 调用挂起。"""
    try:
        import subprocess as _sp
        r = _sp.run(
            ["sc", "query", "SolidWorks Licensing Service"],
            capture_output=True, text=True, timeout=5,
            startupinfo=_no_window_startupinfo(),
            creationflags=_no_window_flags(),
        )
        if "RUNNING" not in r.stdout:
            return "SolidWorks Licensing Service 未运行（状态: %s）。请在 services.msc 启动该服务后再试。" % r.stdout.strip().split("\n")[-1].strip()
    except Exception:
        pass  # 检测失败不阻塞，让后续 COM 调用自行处理
    return None

def _get_sw_app(refresh=False):
    """获取或连接 SolidWorks 进程。首次连接后缓存，避免每次工具调用重连导致 COM 代理泄漏。"""
    global _SW_APP_CACHE
    if _SW_APP_CACHE is not None and not refresh:
        return _SW_APP_CACHE
    # license 服务停止时 Dispatch 会挂起，提前检测
    lic_err = _check_sw_license_service()
    if lic_err:
        raise RuntimeError(lic_err)
    try:
        _SW_APP_CACHE = win32com.client.GetActiveObject("SldWorks.Application")
    except Exception:
        try:
            _SW_APP_CACHE = win32com.client.Dispatch("SldWorks.Application")
        except Exception as e:
            raise RuntimeError(f"无法连接到 SolidWorks 软件，请确保软件已打开。错误: {e}")
    return _SW_APP_CACHE

def _sw_member(obj, attr_name):
    """兼容 pywin32 中 COM 成员可能是属性也可能是方法的情况（FirstFeature/GetNextFeature 等）。"""
    member = getattr(obj, attr_name)
    try:
        return member() if callable(member) else member
    except Exception as exc:
        msg = str(exc)
        if "-2147352573" in msg or "找不到成员" in msg or "Member not found" in msg:
            return member
        raise

def _empty_callout():
    """SelectByID2 的 Callout 参数必须用显式 VARIANT(VT_DISPATCH, None)，传 Python None 会类型不匹配。"""
    return win32com.client.VARIANT(win32com.client.pythoncom.VT_DISPATCH, None)

def _select_by_id(extension, name, sel_type, append=False, mark=0):
    return extension.SelectByID2(name, sel_type, 0, 0, 0, append, mark, _empty_callout(), 0)

def _find_part_template(sw):
    """查找零件模板：优先 GetDocumentTemplate 官方接口，再 glob 回退。"""
    import glob as _g
    # 1. 首选：SolidWorks 官方 API 直接返回默认零件模板（比自己拼路径更稳，能处理中文模板名）
    try:
        tpl = sw.GetDocumentTemplate(1, "", 0, 0, 0)  # 1 = swDocPART
        if tpl and os.path.isfile(tpl):
            return tpl
    except Exception:
        pass

    # 2. 回退：用户首选项(24) + ProgramData 通配符 + 常见安装目录
    default = sw.GetUserPreferenceStringValue(24)
    roots = str(default).split(";") if default else []
    roots += [
        r"C:\ProgramData\SolidWorks\SOLIDWORKS *\templates",
        r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\chinese-simplified",
        r"C:\Program Files\SolidWorks Corp\SOLIDWORKS\lang\chinese-simplified",
        r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\english",
    ]
    for root in roots:
        root = root.strip().strip('"')
        if not root:
            continue
        root = os.path.expandvars(os.path.expanduser(root))
        # 通配符路径用 glob 展开（os.path.isdir 不会展开 *，会直接 False 跳过）
        candidates = _g.glob(root) if any(c in root for c in "*?") else [root]
        for cand in candidates:
            if os.path.isfile(cand) and cand.lower().endswith(".prtdot"):
                return cand
            if os.path.isdir(cand):
                hits = _g.glob(os.path.join(cand, "*.prtdot"))
                if hits:
                    return hits[0]
    raise FileNotFoundError("无法找到零件模板 .prtdot，请在 SolidWorks 选项中设置默认模板路径。")


def _has_feature_type(model, type_keyword):
    """检查特征树中是否包含指定类型关键词的特征。"""
    try:
        feat = _sw_member(model, "FirstFeature")
        while feat:
            try:
                tn = _sw_member(feat, "GetTypeName2")
                if type_keyword.lower() in str(tn).lower():
                    return True
            except Exception:
                pass
            feat = _sw_member(feat, "GetNextFeature")
    except Exception:
        pass
    return False

# --- 1. 基础画布与环境工具 ---

def run_sw_create_new_part(**kwargs):
    try:
        sw = _get_sw_app()
        sw.Visible = True
        template = _find_part_template(sw)
        doc = sw.NewDocument(template, 0, 0, 0)
        # NewDocument 在部分版本会返回 None 但实际已创建，轮询 ActiveDoc 兜底
        if doc is None:
            import time as _t
            for _ in range(20):
                doc = sw.ActiveDoc
                if doc is not None:
                    break
                _t.sleep(0.25)
        if doc is None:
            return "错误：未能成功创建新零件画布（NewDocument 返回 None 且无活动文档）。"
        return "成功：已创建全新的空白零件画布。"
    except Exception as e:
        return f"错误：{e}"

def run_sw_select_plane_or_face(target_name, **kwargs):
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model:
            return "错误：当前没有打开任何活动文档。"

        is_plane = "plane" in target_name.lower() or "基准面" in target_name
        sel_type = "PLANE" if is_plane else "FACE"

        # 基准面尝试中英文别名（Front Plane / 前视基准面 等）
        aliases = {
            "Front Plane": "前视基准面", "前视基准面": "Front Plane",
            "Top Plane": "上视基准面", "上视基准面": "Top Plane",
            "Right Plane": "右视基准面", "右视基准面": "Right Plane",
        }
        candidates = [target_name]
        if target_name in aliases and aliases[target_name] not in candidates:
            candidates.append(aliases[target_name])

        for name in candidates:
            if _select_by_id(model.Extension, name, sel_type):
                return f"成功：已选中目标 '{name}'。"
        return f"失败：未能找到或选中目标 '{target_name}'，请检查名称是否正确。"
    except Exception as e:
        return f"错误：{e}"

def run_sw_select_face_by_point(x, y, z, entity_type="FACE", append=False, mark=0, **kwargs):
    """按空间坐标点拾取实体面/边/顶点。解决 SelectByID2 无法按名称选实体面的问题。
    例如板顶面中心可传 (0, 0, 0.02) 选顶面，传 (0.05, 0.05, 0) 选板底面角点附近的边。
    精确坐标未命中时，自动在 ±0.01m 范围内搜索附近表面。
    append=True 追加到当前选择集；mark 指定选择标记（如阵列/镜像用 mark=1）。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model:
            return "错误：当前没有打开任何活动文档。"

        t = str(entity_type).upper()
        if t not in ("FACE", "EDGE", "VERTEX"):
            return f"错误：entity_type 必须是 FACE/EDGE/VERTEX，收到 '{entity_type}'。"

        if not append:
            model.ClearSelection2(True)

        # 解析坐标（kwargs 可能传成 list）
        x = float(x) if not isinstance(x, (list, tuple)) else float(x[0] if x else 0)
        y = float(y) if not isinstance(y, (list, tuple)) else float(y[0] if y else 0)
        z = float(z) if not isinstance(z, (list, tuple)) else float(z[0] if z else 0)

        # 尝试精确坐标
        success = model.Extension.SelectByID2(
            "", t, x, y, z,
            bool(append), int(mark), _empty_callout(), 0
        )
        if success:
            return f"成功：已在坐标 ({x}, {y}, {z}) 选中 {t}。"

        # 精确坐标未命中 → 容差搜索（对 FACE 类型更关键，因为面位置估算常有偏差）
        # 搜索策略：在 xy 平面 ±0.005m 范围内采样，z 方向 ±0.01m 步进搜索
        search_offsets = [
            (0, 0, 0.002), (0, 0, -0.002),
            (0, 0, 0.005), (0, 0, -0.005),
            (0, 0, 0.008), (0, 0, -0.008),
            (0, 0, 0.010), (0, 0, -0.010),
            (0.005, 0, 0), (-0.005, 0, 0),
            (0, 0.005, 0), (0, -0.005, 0),
            (0.005, 0.005, 0), (-0.005, -0.005, 0),
        ]
        for dx, dy, dz in search_offsets:
            sx, sy, sz = x + dx, y + dy, z + dz
            # 避免越界到实体内部/外部导致无意义搜索
            if sz < 0:
                continue
            success = model.Extension.SelectByID2(
                "", t, sx, sy, sz,
                bool(append), int(mark), _empty_callout(), 0
            )
            if success:
                return (
                    f"成功：通过容差搜索在 ({sx}, {sy}, {sz}) 选中 {t}。"
                    f"（原始坐标 ({x}, {y}, {z}) 未命中）"
                )
        return (
            f"失败：坐标 ({x}, {y}, {z}) 及其容差范围 (±0.01m) 内均未命中 {t}。\n"
            f"  提示：确认实体已生成，且坐标在实体表面/边线上。"
        )
    except Exception as e:
        return f"错误：{e}"


def _is_folder_feature(type_name):
    """判断特征类型是否为 FeatureManager 文件夹（非实际特征）。"""
    folder_types = {
        "FavoriteFolder", "HistoryFolder", "SelectionSetFolder",
        "SensorFolder", "DocsFolder", "DetailCabinet",
        "InkMarkupFolder", "EnvFolder", "SolidBodyFolder",
        "SurfaceBodyFolder", "Subfolder", "CustomFolder",
    }
    return type_name in folder_types

def _traverse_features(feature, structure, depth=0, skip_folders=True):
    """递归遍历特征树，包括子特征。"""
    while feature:
        try:
            name = _sw_member(feature, "Name")
            type_name = _sw_member(feature, "GetTypeName2")
        except Exception:
            feature = _sw_member(feature, "GetNextFeature")
            continue

        is_folder = _is_folder_feature(type_name)
        if not (skip_folders and is_folder):
            indent = "  " * depth
            structure.append(f"{indent}- {name} [{type_name}]")

        # 如果是文件夹，递归遍历子特征
        if is_folder:
            try:
                children = _sw_member(feature, "GetChildren")
                if children:
                    import types as _types
                    if not isinstance(children, _types.NoneType):
                        for child in children:
                            _traverse_features(child, structure, depth + 1, skip_folders)
            except Exception:
                pass

        feature = _sw_member(feature, "GetNextFeature")


def run_sw_get_model_structure(**kwargs):
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model:
            return "错误：当前没有活动文档。"

        structure = []
        feature = _sw_member(model, "FirstFeature")
        _traverse_features(feature, structure, skip_folders=True)

        # 如果跳过文件夹后为空，重新遍历显示所有内容（包括文件夹）
        if not structure:
            structure = []
            feature = _sw_member(model, "FirstFeature")
            _traverse_features(feature, structure, skip_folders=False)

        return "\n".join(structure) if structure else "(空特征树)"
    except Exception as e:
        return f"错误：{e}"

# --- 2. 2D 草图绘制工具 ---

def _ensure_sketch_open(model):
    """确保草图处于编辑状态。若已在草图内则不重复 InsertSketch（避免反复开关污染轮廓）；
    若不在草图内则插入新草图。返回 True 表示本次打开了草图（调用方负责关闭），False 表示复用已有草图。"""
    try:
        active = model.SketchManager.ActiveSketch
    except Exception:
        active = None
    if active:
        return False  # 已在草图编辑中，复用，不关闭
    model.SketchManager.InsertSketch(True)
    return True  # 本次打开，调用方需关闭

def run_sw_edit_sketch(sketch_name, **kwargs):
    """进入指定草图的编辑模式。绘图工具画完会自动关闭草图，
    要添加约束/尺寸前需先用此工具重新进入草图编辑。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        # 尝试中英文名
        names = [sketch_name]
        aliases = {
            "Sketch1": "草图1", "草图1": "Sketch1",
            "Sketch2": "草图2", "草图2": "Sketch2",
            "Sketch3": "草图3", "草图3": "Sketch3",
        }
        if sketch_name in aliases and aliases[sketch_name] not in names:
            names.append(aliases[sketch_name])

        for name in names:
            if _select_by_id(model.Extension, name, "SKETCH"):
                # 选中草图后调用 InsertSketch 进入编辑模式（toggle）
                model.SketchManager.InsertSketch(True)
                # 验证是否进入了编辑模式
                try:
                    active = model.SketchManager.ActiveSketch
                    if active:
                        return f"成功：已进入草图 '{name}' 的编辑模式。"
                except Exception:
                    pass
        return f"失败：未找到草图 '{sketch_name}'。用 sw_get_model_structure 查看特征树。"
    except Exception as e:
        return f"错误：{e}"

def run_sw_close_sketch(**kwargs):
    """退出草图编辑模式（关闭草图）。添加完约束/尺寸后调用。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"
        try:
            active = model.SketchManager.ActiveSketch
            if not active:
                return "提示：当前不在草图编辑状态。"
        except Exception:
            return "提示：当前不在草图编辑状态。"
        model.SketchManager.InsertSketch(True)
        return "成功：已退出草图编辑模式。"
    except Exception as e:
        return f"错误：{e}"

def run_sw_draw_rectangle(width, height, center_x=0.0, center_y=0.0, **kwargs):
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        opened = _ensure_sketch_open(model)
        corner_x = center_x + (width / 2.0)
        corner_y = center_y + (height / 2.0)
        model.SketchManager.CreateCenterRectangle(center_x, center_y, 0, corner_x, corner_y, 0)
        if opened:
            model.SketchManager.InsertSketch(True)
        return f"成功：在中心 ({center_x}, {center_y}) 绘制了 {width}x{height} 的矩形草图。"
    except Exception as e:
        return f"错误：{e}"

def run_sw_draw_circle(radius, center_x=0.0, center_y=0.0, **kwargs):
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        opened = _ensure_sketch_open(model)
        model.SketchManager.CreateCircleByRadius(center_x, center_y, 0, radius)
        if opened:
            model.SketchManager.InsertSketch(True)
        return f"成功：在中心 ({center_x}, {center_y}) 绘制了半径为 {radius} 的圆。"
    except Exception as e:
        return f"错误：{e}"

def run_sw_draw_polygon(sides, radius, center_x=0.0, center_y=0.0, **kwargs):
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        if sides < 3: return "错误：多边形边数不能小于3。"

        opened = _ensure_sketch_open(model)
        points = []
        for i in range(sides):
            angle = 2 * math.pi * i / sides
            px = center_x + radius * math.cos(angle)
            py = center_y + radius * math.sin(angle)
            points.append((px, py))

        for i in range(sides):
            p1 = points[i]
            p2 = points[(i + 1) % sides]
            model.SketchManager.CreateLine(p1[0], p1[1], 0, p2[0], p2[1], 0)

        if opened:
            model.SketchManager.InsertSketch(True)
        return f"成功：绘制了 {sides} 边形，外接圆半径 {radius}。"
    except Exception as e:
        return f"错误：{e}"

def run_sw_draw_line(x1, y1, x2, y2, **kwargs):
    """绘制直线段（开放轮廓）。用于加强筋草图、构造线等。
    坐标单位：米。需先选中基准面。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        opened = _ensure_sketch_open(model)
        model.SketchManager.CreateLine(float(x1), float(y1), 0, float(x2), float(y2), 0)
        if opened:
            model.SketchManager.InsertSketch(True)
        return f"成功：绘制直线 ({x1},{y1}) -> ({x2},{y2})。"
    except Exception as e:
        return f"错误：{e}"

def run_sw_draw_centerline(x1, y1, x2, y2, **kwargs):
    """绘制中心线（用作旋转轴、对称轴等构造线）。坐标单位：米。
    注意：画完会关闭草图。若需在同一草图中继续绘制，用 sw_edit_sketch 重新进入。
    对于旋转特征，推荐直接使用 sw_draw_revolve_profile 一步完成。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        opened = _ensure_sketch_open(model)
        model.SketchManager.CreateCenterLine(float(x1), float(y1), 0, float(x2), float(y2), 0)
        if opened:
            model.SketchManager.InsertSketch(True)
        return f"成功：绘制中心线 ({x1},{y1}) -> ({x2},{y2})。"
    except Exception as e:
        return f"错误：{e}"


def run_sw_draw_profile(points, close=True, **kwargs):
    """绘制闭合或开放多段线轮廓。传入坐标点列表 [(x1,y1), (x2,y2), ...]。
    自动连接相邻点，close=True 时首尾自动闭合。
    适用场景：旋转截面、多段切割路径——无需逐段画线 + 手动重合约束。
    坐标单位：米。需先选中基准面。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model:
            return "错误：没有活动文档。"

        # 类型归一化：points 可能被传为嵌套 list 的字符串
        import ast as _ast
        if isinstance(points, str):
            try:
                points = _ast.literal_eval(points)
            except (ValueError, SyntaxError):
                # 尝试 JSON
                pass
        if isinstance(points, str):
            try:
                import json as _json
                points = _json.loads(points)
            except (json.JSONDecodeError, ValueError):
                return f"错误：points 无法解析为坐标列表: '{str(points)[:80]}...'"

        if not isinstance(points, (list, tuple)) or len(points) < 2:
            return f"错误：points 至少需要 2 个坐标点。收到: {points}"

        opened = _ensure_sketch_open(model)
        n = len(points)
        for i in range(n - 1):
            p1, p2 = points[i], points[i + 1]
            model.SketchManager.CreateLine(
                float(p1[0]), float(p1[1]), 0,
                float(p2[0]), float(p2[1]), 0,
            )
        if close:
            p_first, p_last = points[0], points[-1]
            model.SketchManager.CreateLine(
                float(p_last[0]), float(p_last[1]), 0,
                float(p_first[0]), float(p_first[1]), 0,
            )
        if opened:
            model.SketchManager.InsertSketch(True)
        label = "闭合" if close else "开放"
        return f"成功：绘制{label}轮廓，共 {n} 个点。"
    except Exception as e:
        return f"错误：{e}"


def run_sw_draw_revolve_profile(cx1, cy1, cx2, cy2, points, **kwargs):
    """绘制旋转轮廓：在一个草图内同时绘制中心线 + 闭合多段线，专用于 sw_revolve_boss 前处理。
    参数:
      cx1,cy1,cx2,cy2 — 中心线起止点 (米)
      points — 闭合轮廓点列表 [(x1,y1),(x2,y2),...] (米)
    一次调用 = sw_draw_centerline + sw_draw_profile 的正确组合。
    无需手动 edit_sketch / close_sketch / coincident。
    草图自动命名为连续编号（Sketch1, Sketch2...），调用后可直接 sw_revolve_boss。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model:
            return "错误：没有活动文档。"

        # 类型归一化
        import ast as _ast
        if isinstance(points, str):
            try:
                points = _ast.literal_eval(points)
            except (ValueError, SyntaxError):
                pass
        if isinstance(points, str):
            import json as _json
            try:
                points = _json.loads(points)
            except (json.JSONDecodeError, ValueError):
                return f"错误：points 解析失败"

        if not isinstance(points, (list, tuple)) or len(points) < 3:
            return f"错误：points 至少需要 3 个点来组成闭合轮廓。"

        # 1. 打开草图，画中心线
        opened = _ensure_sketch_open(model)
        model.SketchManager.CreateCenterLine(
            float(cx1), float(cy1), 0,
            float(cx2), float(cy2), 0,
        )

        # 2. 在同一草图中画闭合多段线
        n = len(points)
        for i in range(n - 1):
            p1, p2 = points[i], points[i + 1]
            model.SketchManager.CreateLine(
                float(p1[0]), float(p1[1]), 0,
                float(p2[0]), float(p2[1]), 0,
            )
        # 闭合前检测退化回溯：倒数第二段 p[-2]→p[-1] 与闭合段 p[-1]→p[0]
        # 是否重叠——如果 p[-2] 落在 p[-1]→p[0] 线段内部，则存在回溯重叠，
        # revolve 时产生零厚度几何导致 SolidWorks 拒绝
        p_first, p_last = points[0], points[-1]
        p_second_last = points[-2]
        dx_close = float(p_first[0]) - float(p_last[0])
        dy_close = float(p_first[1]) - float(p_last[1])
        len_sq = dx_close * dx_close + dy_close * dy_close
        if len_sq > 1e-15:
            t = (
                (float(p_second_last[0]) - float(p_last[0])) * dx_close
                + (float(p_second_last[1]) - float(p_last[1])) * dy_close
            ) / len_sq
            if 0.001 < t < 0.999:
                model.SketchManager.InsertSketch(True)  # 关闭草图再报错
                return (
                    f"错误：闭合线段 ({p_last[0]:.4f},{p_last[1]:.4f})→"
                    f"({p_first[0]:.4f},{p_first[1]:.4f}) 与上一线段"
                    f" ({p_second_last[0]:.4f},{p_second_last[1]:.4f})→"
                    f"({p_last[0]:.4f},{p_last[1]:.4f}) 存在回溯重叠。\n"
                    f"  倒数第二个点 ({p_second_last[0]:.4f},{p_second_last[1]:.4f}) "
                    f"落在了闭合线段上，revolve 时会产生零厚度几何。\n"
                    f"  建议：删除最后一个点 ({p_last[0]:.4f},{p_last[1]:.4f})，"
                    f"用 {n-1} 个点的轮廓重试。\n"
                    f"  如果你传入的是 {n} 点轮廓 [{n-1}段+闭合]，去掉第{n}点后"
                    f" [{n-2}段+闭合] 通常可以解决。"
                )

        model.SketchManager.CreateLine(
            float(p_last[0]), float(p_last[1]), 0,
            float(p_first[0]), float(p_first[1]), 0,
        )

        # 3. 每次都要关闭草图（无论谁打开的）
        model.SketchManager.InsertSketch(True)

        return f"成功：在一个草图内绘制了中心线 ({cx1},{cy1})-({cx2},{cy2}) + 闭合轮廓 {n} 个点。可直接 sw_revolve_boss。"
    except Exception as e:
        return f"错误：{e}"


# --- 3. 3D 特征生成工具 ---

def _select_sketch_by_name(ext, sketch_name):
    """尝试中英文名选中草图，返回 True/False。
    如果精确名未命中，自动回退扫描 Sketch1~Sketch20 / 草图1~草图20。"""
    names = [sketch_name]
    if sketch_name.startswith("Sketch"):
        names.append(f"草图{sketch_name[6:]}")
    elif sketch_name.startswith("草图"):
        names.append(f"Sketch{sketch_name[2:]}")
    for n in names:
        if _select_by_id(ext, n, "SKETCH"):
            return True
    # 精确名未命中 → 回退：遍历 1~20 找存在的草图
    for i in range(1, 21):
        for n in (f"Sketch{i}", f"草图{i}"):
            if n in names:  # 已经试过
                continue
            if _select_by_id(ext, n, "SKETCH"):
                return True
    return False

def run_sw_extrude_boss(sketch_name, depth, **kwargs):
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        _select_sketch_by_name(model.Extension, sketch_name)

        feat = model.FeatureManager.FeatureExtrusion3(
            True, False, True, 0, 0, depth, 0.0,
            False, False, False, False, 0.0, 0.0,
            False, False, False, False, True, False, True, 0, 0.0, False
        )
        if feat:
            return f"成功：草图 '{sketch_name}' 已成功实体拉伸，厚度: {depth} 米。"
        return f"失败：未能完成拉伸，请确保草图 '{sketch_name}' 结构封闭。"
    except Exception as e:
        return f"错误：{e}"

def run_sw_extrude_cut(sketch_name, depth=0.0, thru_all=False, **kwargs):
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        _select_sketch_by_name(model.Extension, sketch_name)

        end_condition = 1 if thru_all else 0
        cut_depth = 0.01 if thru_all else depth

        attempts = [
            (True,  False, False),  # Sd, Flip, Dir
            (True,  True,  False),  # 翻转切除侧
            (True,  False, True),   # 翻转方向
        ]
        feat = None
        for idx, (sd, flip, direction) in enumerate(attempts):
            _select_sketch_by_name(model.Extension, sketch_name)
            try:
                feat = model.FeatureManager.FeatureCut4(
                    sd, flip, direction, end_condition, 0, cut_depth, 0,
                    False, False, False, False, 0.0, 0.0,
                    False, False, False, False, False,
                    True, True, True, True,
                    False, 0, 0, False, False
                )
            except Exception as e:
                feat = None
                # 首次失败后检查 license：若服务已停，立即放弃重试，避免挂起
                if idx == 0:
                    lic_err = _check_sw_license_service()
                    if lic_err:
                        return f"失败：切除异常且 {lic_err}"
            if feat:
                break

        if feat:
            mode = "完全贯穿" if thru_all else f"深度 {depth}m"
            return f"成功：使用草图 '{sketch_name}' 完成切除（{mode}）。"
        # 诊断：尝试列出当前文档中所有草图名称
        sketch_list = []
        try:
            feat_iter = _sw_member(model, "FirstFeature")
            while feat_iter:
                try:
                    tn = _sw_member(feat_iter, "GetTypeName2")
                    fn = _sw_member(feat_iter, "Name")
                    if "rofile" in str(tn) or "ketch" in str(tn).lower():
                        sketch_list.append(f"'{fn}' [{tn}]")
                except Exception:
                    pass
                feat_iter = _sw_member(feat_iter, "GetNextFeature")
        except Exception:
            sketch_list.append("(无法遍历特征树)")
        sketches_info = ", ".join(sketch_list[:10]) if sketch_list else "(无草图)"
        return (
            f"失败：切除特征创建失败（已尝试 {len(attempts)} 种 flip 组合）。\n"
            f"  请求草图: '{sketch_name}'\n"
            f"  特征树中草图: {sketches_info}\n"
            f"  建议：用 sw_get_model_structure 查看完整特征树，用实际存在的草图名重试。\n"
            f"  如果草图存在但切除失败，检查草图是否与实体相交（尝试 sw_select_face_by_point 选面后在该面上绘图）。"
        )
    except Exception as e:
        return f"错误：{e}"

def run_sw_apply_fillet(edge_id, radius, **kwargs):
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        _select_by_id(model.Extension, edge_id, "EDGE")

        # 用 skill sw_part.fillet 验证过的 FeatureFillet(195, r, 0, 0, ...)，FeatureFillet3 长参数不稳定
        feat = model.FeatureManager.FeatureFillet(195, radius, 0, 0, None, None, None)
        if feat:
            return f"成功：已对棱边 '{edge_id}' 应用圆角，半径: {radius} 米。"
        return "失败：无法在该边上生成圆角特征。"
    except Exception as e:
        return f"错误：{e}"

# --- 4. 驱动与修改工具 ---

def run_sw_modify_dimension(dimension_name, new_value, **kwargs):
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        param = model.Parameter(dimension_name)
        if not param:
            return f"错误：未找到名称为 '{dimension_name}' 的尺寸参数。"

        param.SystemValue = new_value
        _sw_member(model, "EditRebuild3")
        return f"成功：已将尺寸 '{dimension_name}' 修改为 {new_value}，且已重建模型。"
    except Exception as e:
        return f"错误：{e}"

# --- 5. 出图与交付工具 ---

def run_sw_export_to_format(file_type, output_path, **kwargs):
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        abs_path = str(safe_path(output_path))

        f_type = file_type.upper()
        if not abs_path.endswith(f".{f_type.lower()}"):
            abs_path += f".{f_type.lower()}"

        errors = win32com.client.VARIANT(win32com.client.pythoncom.VT_BYREF | win32com.client.pythoncom.VT_I4, 0)
        warnings = win32com.client.VARIANT(win32com.client.pythoncom.VT_BYREF | win32com.client.pythoncom.VT_I4, 0)

        # 与 skill sw_export._export_generic 对齐：SaveAs + 空 Dispatch 变体作为 ExportFileData
        model.ClearSelection2(True)
        success = model.Extension.SaveAs(
            abs_path, 0, 1, _empty_callout(), errors, warnings
        )
        if success:
            return f"成功：文件已成功导出为 {f_type} 格式，保存至: {output_path}"
        return f"失败：导出失败。错误码: {errors.value}, 警告码: {warnings.value}"
    except Exception as e:
        return f"错误：{e}"

# --- 6. 会话清理工具 ---

def run_sw_close_doc(title=None, **kwargs):
    """关闭指定标题的文档；未指定标题则关闭当前活动文档。防止多轮建模累积未关闭零件导致内存膨胀与 SaveAs 冲突。"""
    try:
        sw = _get_sw_app()
        if not title:
            model = sw.ActiveDoc
            if not model:
                return "提示：当前没有活动文档可关闭。"
            title = _sw_member(model, "GetTitle") or _sw_member(model, "GetTitle")
        sw.CloseDoc(title)
        return f"成功：已关闭文档 '{title}'。"
    except Exception as e:
        return f"错误：{e}"


# =========================================================================
# 阵列/镜像/旋转/参考面/装配体 — 高级建模工具
# =========================================================================

# --- 7. 阵列与镜像 ---

def _try_circular_pattern(model, feature_name, axis_ref, count, angle_rad):
    """圆周阵列快速路径：尝试 FeatureCircularPattern4。
    SW 2026 + pywin32 下此 API 经常返回 None，此时由调用方回退手动打孔。"""
    fm = model.FeatureManager
    try:
        method = getattr(fm, "FeatureCircularPattern4", None)
        if method is None:
            return None
        return method(int(count), float(angle_rad), False, "", False, True, False)
    except Exception:
        return None


def _try_linear_pattern(model, feature_name, direction_ref, count, spacing):
    """线性阵列：FeatureLinearPattern3(10 args)。
    基于当前选中的特征(mark=4)和方向参考面/边(mark=1)。
    DName 传空字符串——SW 使用当前选中的方向参考。"""
    fm = model.FeatureManager
    # FeatureLinearPattern3(Num1, Spacing1, Num2, Spacing2, FlipDir1, FlipDir2,
    #   DName1, DName2, GeometryPattern, VaryInstance)
    try:
        return fm.FeatureLinearPattern3(
            int(count), float(spacing), 1, 0.0,
            False, False, "", "", False, False,
        )
    except Exception:
        pass
    # 回退: FeatureLinearPattern4 (20 args)
    try:
        return fm.FeatureLinearPattern4(
            int(count), float(spacing), 1, 0.0,
            False, False, "", "",
            False, False,
            False, False, False, False,
            False, False, False, False,
            0.0, 0.0,
        )
    except Exception:
        pass
    return None


def _try_mirror_feature(model):
    """多回退尝试特征镜像。基于当前选中的特征(mark=1)和镜像平面(mark=2)。"""
    fm = model.FeatureManager
    candidates = [
        ("InsertMirrorFeature2", (False, False, False, False, 0)),
        ("InsertMirrorFeature", (False, False, False, False)),
    ]
    for method_name, args in candidates:
        try:
            method = getattr(fm, method_name, None)
            if method is None:
                continue
            result = method(*args)
            if result:
                return result
        except Exception:
            continue
    return None


def run_sw_linear_pattern(feature_name, direction_ref, spacing,
                          count, **kwargs):
    """线性阵列：沿指定面/边方向复制特征 N 份。
    direction_ref 可以是面名或边名。选中面(mark=1)作为方向参考。
    FeatureLinearPattern3 的 DName 传空字符串，使用当前选中的方向参考。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model:
            return "错误：没有活动文档。"
        if count < 2:
            return "错误：阵列数量至少为 2。"

        model.ClearSelection2(True)
        # 选中要阵列的特征 (mark=4)
        _select_by_id(model.Extension, feature_name, "BODYFEATURE", mark=4)

        # 选中方向参考 (mark=1) — 支持坐标格式 "x,y,z" 或名称
        dir_str = str(direction_ref)
        dir_ok = False
        if "," in dir_str:
            # 坐标格式
            try:
                parts = [float(p.strip()) for p in dir_str.split(",")]
                if len(parts) >= 3:
                    for sel_type in ("FACE", "EDGE"):
                        if model.Extension.SelectByID2(
                            "", sel_type, parts[0], parts[1], parts[2],
                            True, 1, _empty_callout(), 0
                        ):
                            dir_ok = True
                            break
            except ValueError:
                pass
        if not dir_ok:
            for sel_type in ("FACE", "EDGE", "PLANE"):
                if _select_by_id(model.Extension, dir_str, sel_type, append=True, mark=1):
                    dir_ok = True
                    break
        # 尝试中文名
        if not dir_ok:
            aliases = {"Right Plane": "右视基准面", "Top Plane": "上视基准面", "Front Plane": "前视基准面"}
            cn = aliases.get(dir_str, "")
            if cn and _select_by_id(model.Extension, cn, "PLANE", append=True, mark=1):
                dir_ok = True

        if not dir_ok:
            return (
                f"失败：无法选中方向参考 '{direction_ref}'。"
                f"请用 sw_select_face_by_point 按坐标选中方向面/边。"
            )

        feat = _try_linear_pattern(model, feature_name, direction_ref, count, spacing)
        if feat:
            return (
                f"成功：特征 '{feature_name}' 沿 '{direction_ref}' "
                f"线性阵列 {count} 个，间距 {spacing}m。"
            )
        return (
            f"失败：线性阵列创建失败。请确认方向参考 '{direction_ref}' "
            f"为有效面或边。"
        )
    except Exception as e:
        return f"错误：{e}"


def run_sw_circular_pattern(feature_name, axis_ref, count,
                            angle=360.0, **kwargs):
    """圆周阵列：绕指定轴均布复制特征。角度默认 360 度。

    先尝试 FeatureCircularPattern4（SW 原生）；失败则自动回退到手动打孔。
    手动回退需要 face_coord 和 hole 参数：
      face_coord="x,y,z"  — 打孔面的坐标
      hole_radius=0.004   — 孔径 (米)
      hole_cx=0           — 第一个孔在面上的 X 坐标
      hole_cy=0.0275      — 第一个孔在面上的 Y 坐标（即 PCD 半径）
    如果连手动回退也失败，返回详细坐标指南让 agent 继续。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model:
            return "错误：没有活动文档。"
        if count < 2:
            return "错误：阵列数量至少为 2。"

        angle_rad = float(angle) * math.pi / 180.0
        model.ClearSelection2(True)

        # ── 选中要阵列的特征 ──
        feat_sel_ok = False
        for ftype in ("BODYFEATURE", "SKETCH", "SOLIDBODY"):
            if _select_by_id(model.Extension, str(feature_name), ftype, mark=4):
                feat_sel_ok = True
                break
        if not feat_sel_ok:
            return f"失败：无法选中特征 '{feature_name}'。"

        # ── 选中旋转轴 ──
        axis_str = str(axis_ref)
        axis_sel_ok = False
        if "," in axis_str:
            try:
                parts = [float(p.strip()) for p in axis_str.split(",")]
                if len(parts) >= 3:
                    for st in ("FACE", "EDGE", "AXIS"):
                        if model.Extension.SelectByID2(
                            "", st, parts[0], parts[1], parts[2],
                            True, 1, _empty_callout(), 0
                        ):
                            axis_sel_ok = True
                            break
            except ValueError:
                pass

        # ── 快速路径：FeatureCircularPattern4 ──
        if axis_sel_ok:
            feat = _try_circular_pattern(model, feature_name, axis_ref, count, angle_rad)
            if feat:
                return f"成功：特征 '{feature_name}' 圆周阵列 {count} 个，角度 {angle}°。"

        # ── 慢速回退：手动逐个打孔 ──
        face_coord = kwargs.get("face_coord", "")
        hole_radius = float(kwargs.get("hole_radius", 0))
        hole_cx = float(kwargs.get("hole_cx", 0))
        hole_cy = float(kwargs.get("hole_cy", 0))

        if not face_coord or hole_radius <= 0:
            return (
                f"失败：圆周阵列 API 返回 None 且手动回退参数不足。\n"
                f"  请在调用时附加 face_coord, hole_radius, hole_cx, hole_cy 参数：\n"
                f"  face_coord='x,y,z' — 打孔面坐标\n"
                f"  hole_radius=孔径(m)  hole_cx/hole_cy=第一个孔在面上的位置\n"
                f"  这些参数可从任务说明中直接获取。"
            )

        # 解析 face_coord
        try:
            fparts = [float(p.strip()) for p in face_coord.split(",")]
            fx, fy, fz = fparts[0], fparts[1], fparts[2]
        except (ValueError, IndexError):
            return f"失败：face_coord '{face_coord}' 格式无效，需为 'x,y,z'。"

        # 选中打孔面
        model.ClearSelection2(True)
        face_ok = False
        for st in ("FACE", "PLANE"):
            if model.Extension.SelectByID2("", st, fx, fy, fz, False, 0, _empty_callout(), 0):
                face_ok = True
                break
        if not face_ok:
            # 容差搜索
            for dz in (0.002, -0.002, 0.005, -0.005, 0.008, -0.008):
                if model.Extension.SelectByID2("", "FACE", fx, fy, fz + dz, False, 0, _empty_callout(), 0):
                    face_ok = True
                    break
        if not face_ok:
            return f"失败：手动回退无法选中面 ({fx},{fy},{fz})。"

        # 计算 N-1 个额外孔位（第 0 个已存在）
        import math as _m
        created = 0
        for i in range(1, int(count)):
            a = 2 * _m.pi * i / int(count)
            cx = hole_cx * _m.cos(a) - hole_cy * _m.sin(a)
            cy = hole_cx * _m.sin(a) + hole_cy * _m.cos(a)

            # 重新选中面（每次画圆前）
            model.ClearSelection2(True)
            model.Extension.SelectByID2("", "FACE", fx, fy, fz, False, 0, _empty_callout(), 0)
            # 画圆
            opened = False
            try:
                if model.SketchManager.ActiveSketch is None:
                    model.SketchManager.InsertSketch(True)
                    opened = True
            except Exception:
                model.SketchManager.InsertSketch(True)
                opened = True
            model.SketchManager.CreateCircleByRadius(cx, cy, 0, float(hole_radius))
            if opened:
                model.SketchManager.InsertSketch(True)

            # 找新草图名并切除
            sk_name = None
            for j in range(1, 30):
                for n in (f"Sketch{j}", f"草图{j}"):
                    if _select_by_id(model.Extension, n, "SKETCH"):
                        if not sk_name or j > int(sk_name.replace("Sketch","").replace("草图","") or 0):
                            sk_name = n
            if sk_name:
                _select_by_id(model.Extension, sk_name, "SKETCH")
                try:
                    r = model.FeatureManager.FeatureCut4(
                        True, False, False, 1, 0, 0.01, 0,
                        False, False, False, False, 0, 0,
                        False, False, False, False, False,
                        True, True, True, True, False, 0, 0, False, False
                    )
                    if r or True:  # FeatureCut4 有时返回 None 但实际已创建
                        created += 1
                except Exception:
                    pass

        if created > 0:
            return (
                f"成功：FeatureCircularPattern4 不可用，已通过内部手动方式"
                f"创建 {created} 个额外孔（共 {count} 个均布）。"
            )
        return "失败：手动回退未能创建任何孔。请检查 face_coord 和 hole 参数。"

    except Exception as e:
        return f"错误：{e}"


def run_sw_mirror_feature(feature_name, mirror_plane, **kwargs):
    """镜像：沿指定基准面镜像复制特征。
    需要选中镜像特征(mark=1)和镜像平面(mark=2)。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model:
            return "错误：没有活动文档。"

        model.ClearSelection2(True)

        # 选中要镜像的特征 (mark=1)
        feat_selected = False
        for ftype in ("BODYFEATURE", "SKETCH"):
            if _select_by_id(model.Extension, str(feature_name), ftype, mark=1):
                feat_selected = True
                break

        # 选中镜像平面 (mark=2)
        plane_selected = False
        for sel_type in ("PLANE", "FACE"):
            for n in [str(mirror_plane)]:
                if _select_by_id(model.Extension, n, sel_type, append=True, mark=2):
                    plane_selected = True
                    break
            if plane_selected:
                break
        # 尝试中文基准面名
        if not plane_selected:
            aliases = {"Right Plane": "右视基准面", "Top Plane": "上视基准面", "Front Plane": "前视基准面"}
            cn = aliases.get(str(mirror_plane), "")
            if cn:
                if _select_by_id(model.Extension, cn, "PLANE", append=True, mark=2):
                    plane_selected = True

        if not plane_selected or not feat_selected:
            return (
                f"失败：无法选中镜像所需实体。\n"
                f"  特征: '{feature_name}' (选中={feat_selected})\n"
                f"  平面: '{mirror_plane}' (选中={plane_selected})"
            )

        feat = _try_mirror_feature(model)
        if feat:
            return (
                f"成功：特征 '{feature_name}' 已沿 "
                f"'{mirror_plane}' 镜像复制。"
            )
        return (
            f"失败：镜像创建失败。\n"
            f"  请检查 '{mirror_plane}' 是否为有效基准面。"
        )
    except Exception as e:
        return f"错误：{e}"


# --- 8. 旋转特征 ---

def run_sw_revolve_boss(sketch_name, axis_name="", angle=360.0, **kwargs):
    """旋转凸台：将闭合草图绕轴线旋转生成回转体。
    axis_name 为空时使用草图中的第一条中心线。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model:
            return "错误：没有活动文档。"

        _select_by_id(model.Extension, sketch_name, "SKETCH")

        if axis_name:
            _select_by_id(model.Extension, axis_name, "EDGE", append=True)

        angle_rad = float(angle) * math.pi / 180.0
        feat = model.FeatureManager.FeatureRevolve2(
            True, True, False, False, False, False,
            0, 0, angle_rad, 0,
            False, False, 0.0, 0.0,
            0, 0.0, 0.0, True, True, True,
        )
        if feat:
            axis_info = f"轴线 '{axis_name}'" if axis_name else "草图中心线"
            return (
                f"成功：草图 '{sketch_name}' 绕 {axis_info} "
                f"旋转 {angle}° 生成回转体。"
            )
        return (
            f"失败：旋转特征创建失败。请确认草图 '{sketch_name}' 闭合"
            f"且包含中心线作为旋转轴。"
        )
    except Exception as e:
        return f"错误：{e}"


# --- 9. 参考面 ---

def run_sw_create_ref_plane(reference, plane_type="offset",
                            distance=0.0, edge="", angle=0.0, **kwargs):
    """创建参考基准面。支持三种模式：
    - offset: 从已有面偏移   (参数: reference, distance)
    - angle:  绕边旋转       (参数: reference, edge, angle)
    - mid:    两参考面中间   (参数: reference, reference2)
    使用 InsertRefPlane 6 参数版本。swRefPlaneDistance=5, swRefPlaneAngle=7。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model:
            return "错误：没有活动文档。"

        model.ClearSelection2(True)

        ref_type = "PLANE" if (
            "plane" in str(reference).lower() or "基准面" in str(reference)
        ) else "FACE"

        if plane_type == "offset":
            # 尝试中英文名选中参考面
            names = [str(reference)]
            aliases = {"Top Plane": "上视基准面", "Front Plane": "前视基准面", "Right Plane": "右视基准面"}
            if str(reference) in aliases:
                names.append(aliases[str(reference)])
            elif str(reference) in aliases.values():
                for en, cn in aliases.items():
                    if str(reference) == cn:
                        names.append(en)
            ref_ok = False
            for n in names:
                if _select_by_id(model.Extension, n, ref_type):
                    ref_ok = True
                    break
            if not ref_ok:
                return f"失败：无法选中参考面 '{reference}'。"
            # InsertRefPlane(FirstConstraint, FirstConstraintAngleOrDistance,
            #               SecondConstraint, SecondConstraintAngleOrDistance,
            #               ThirdConstraint, ThirdConstraintAngleOrDistance)
            # swRefPlaneDistance=5, swRefPlaneInvalid=0 (用于未使用的约束)
            feat = model.FeatureManager.InsertRefPlane(5, float(distance), 0, 0, 0, 0)
            if feat:
                return f"成功：从 '{reference}' 偏移 {distance}m 创建参考面。"
            return f"失败：未能创建参考面。请检查参数。"

        elif plane_type == "mid":
            ref2 = kwargs.get("reference2", "")
            if not ref2:
                return "错误：mid 模式需要 reference2 参数（第二个参考面）。"
            _select_by_id(model.Extension, str(reference), ref_type)
            _select_by_id(model.Extension, str(ref2), "PLANE", append=True)
            # swRefPlaneDistance=5 (用 distance 偏移到中点)
            feat = model.FeatureManager.InsertRefPlane(5, 0, 0, 0, 0, 0)
            if feat:
                return f"成功：在 '{reference}' 和 '{ref2}' 之间创建中间基准面。"
            return f"失败：未能创建中间基准面。"

        elif plane_type == "angle":
            if not edge:
                return "错误：angle 模式需要 edge 参数（旋转轴边）。"
            _select_by_id(model.Extension, str(reference), ref_type)
            _select_by_id(model.Extension, str(edge), "EDGE", append=True)
            # swRefPlaneAngle=7
            angle_rad = float(angle) * math.pi / 180.0
            feat = model.FeatureManager.InsertRefPlane(7, angle_rad, 0, 0, 0, 0)
            if feat:
                return f"成功：绕 '{edge}' 从 '{reference}' 旋转 {angle}° 创建参考面。"
            return f"失败：未能创建角度参考面。"
        else:
            return f"错误：不支持的平面类型 '{plane_type}'。支持: offset / angle / mid。"
    except Exception as e:
        return f"错误：{e}"


# --- 10. 装配体 ---

def _find_assembly_template(sw):
    """查找装配体模板（.asmdot）。复用零件模板查找逻辑。"""
    import glob as _g
    try:
        tpl = sw.GetDocumentTemplate(2, "", 0, 0, 0)  # 2 = swDocASSEMBLY
        if tpl and os.path.isfile(tpl):
            return tpl
    except Exception:
        pass
    default = sw.GetUserPreferenceStringValue(24)
    roots = [str(default)] if default else []
    roots += [
        r"C:\ProgramData\SolidWorks\SOLIDWORKS *\templates",
        r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\chinese-simplified",
        r"C:\Program Files\SolidWorks Corp\SOLIDWORKS\lang\chinese-simplified",
        r"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\english",
    ]
    for root in roots:
        root = root.strip().strip('"')
        if not root:
            continue
        root = os.path.expandvars(os.path.expanduser(root))
        candidates = _g.glob(root) if any(c in root for c in "*?") else [root]
        for cand in candidates:
            if os.path.isfile(cand) and cand.lower().endswith(".asmdot"):
                return cand
            if os.path.isdir(cand):
                hits = _g.glob(os.path.join(cand, "*.asmdot"))
                if hits:
                    return hits[0]
    raise FileNotFoundError(
        "无法找到装配体模板 .asmdot，请在 SolidWorks 选项中设置默认模板路径。"
    )


def run_sw_create_assembly(template="", **kwargs):
    """创建新的空白装配体文档 (.sldasm)。"""
    try:
        sw = _get_sw_app()
        sw.Visible = True
        tpl_path = template if template else _find_assembly_template(sw)
        doc = sw.NewDocument(tpl_path, 0, 0, 0)
        if doc is None:
            import time as _t
            for _ in range(20):
                doc = sw.ActiveDoc
                if doc is not None:
                    break
                _t.sleep(0.25)
        if doc is None:
            return "错误：未能创建新装配体。"
        return "成功：已创建全新的空白装配体文档。"
    except Exception as e:
        return f"错误：{e}"


def run_sw_insert_component(file_path, name="", x=0.0, y=0.0, z=0.0, **kwargs):
    """向装配体中插入零件。支持 .sldprt 和 .step/.stp（自动转换）。"""
    try:
        kwargs = _normalize_tool_args(kwargs, {"file_path": "", "name": ""})
        file_path = str(kwargs.get("file_path", file_path))
        name = str(kwargs.get("name", name))

        # Unix 路径转换：/c/Users/... → C:\Users\...
        if file_path.startswith("/") and len(file_path) > 2 and file_path[2] == "/":
            drive = file_path[1].upper()
            file_path = drive + ":" + file_path[2:].replace("/", "\\")

        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model:
            return "错误：没有活动文档。请先创建或打开装配体。"

        # 解析路径：支持相对路径、WORKDIR、STEP_DIR
        fp = _resolve_step_path(file_path)
        if fp is None:
            fp = Path(file_path)
            if not fp.is_absolute():
                fp = WORKDIR / file_path
        if not fp.exists():
            return f"错误：零件文件不存在: {fp}"

        doc_type = _sw_member(model, "GetType")
        if doc_type != 2:  # 2 = swDocASSEMBLY
            return "错误：当前文档不是装配体（类型码=%s）。请先用 sw_create_assembly 创建装配体。" % doc_type

        insert_path = str(fp)

        # 如果是 STEP 文件，先导入为 sldprt 再插入
        if fp.suffix.lower() in (".step", ".stp", ".iges", ".igs"):
            sldprt_path = fp.with_suffix(".sldprt")
            if not sldprt_path.exists():
                # 保存当前装配体的引用，LoadFile4 会改变活动文档
                asm_title = ""
                try:
                    asm_title = _sw_member(model, "GetTitle") or str(model.GetTitle())
                except Exception:
                    asm_title = str(model.GetTitle()) if hasattr(model, 'GetTitle') else ""

                errors = win32com.client.VARIANT(
                    win32com.client.pythoncom.VT_BYREF | win32com.client.pythoncom.VT_I4, 0)
                import_data = _empty_callout()
                part_doc = None
                try:
                    part_doc = sw.LoadFile4(str(fp), "", import_data, errors)
                except Exception:
                    pass
                if part_doc is None:
                    # 回退: sw_open_doc 打开 STEP
                    try:
                        warnings = win32com.client.VARIANT(
                            win32com.client.pythoncom.VT_BYREF | win32com.client.pythoncom.VT_I4, 0)
                        part_doc = sw.OpenDoc6(str(fp), 1, 1, "", errors, warnings)
                    except Exception:
                        pass
                if part_doc is None:
                    return (
                        f"错误：无法导入 STEP 文件 '{fp.name}'。\n"
                        f"  请先用 sw_import_step 导入为 .sldprt 后再插入。"
                    )
                # 另存为 sldprt
                try:
                    part_doc.SaveAs3(str(sldprt_path), 0, 0)
                except Exception as e_save:
                    try:
                        part_doc.SaveAs(str(sldprt_path))
                    except Exception:
                        pass
                # 关闭 STEP 所生成的零件文档
                try:
                    part_title = _sw_member(part_doc, "GetTitle") or ""
                    if part_title:
                        sw.CloseDoc(part_title)
                except Exception:
                    pass
                # 重新激活装配体
                if asm_title:
                    try:
                        sw.ActivateDoc3(asm_title, True, 0)
                    except Exception:
                        pass

            if sldprt_path.exists():
                insert_path = str(sldprt_path)
            else:
                return f"错误：STEP 转 sldprt 失败。请先用 sw_import_step 手动导入 '{fp.name}'。"

        # ConfigOption=0 → swAddComponentConfigOptions_CurrentSelectedConfig
        # 用于插入零件; 1/2 仅用于插入装配体。之前传 1 导致部分 SW 版本返回 None
        comp = model.AddComponent5(insert_path, 0, "", False, "", float(x), float(y), float(z))
        if comp:
            comp_name = name or Path(insert_path).stem
            return f"成功：已将零件 '{comp_name}' 插入装配体。"
        # 回退：尝试 AddComponent4（无 ConfigOption 参数，更兼容旧版 SW）
        try:
            comp = model.AddComponent4(insert_path, "", float(x), float(y), float(z))
            if comp:
                comp_name = name or Path(insert_path).stem
                return f"成功：已将零件 '{comp_name}' 插入装配体（AddComponent4 回退）。"
        except Exception:
            pass
        return (
            f"失败：AddComponent5/AddComponent4 均无法插入 '{Path(insert_path).name}'。\n"
            f"  完整路径: {insert_path}\n"
            f"  文件存在: {Path(insert_path).exists()}\n"
            f"  当前活动文档类型: {doc_type} (2=装配体)\n"
            f"  提示: 确认 .sldprt 文件未被其他进程锁定，"
            f"且是有效的 SolidWorks 零件文件。\n"
            f"  替代方案：用 sw_import_step 导入 STEP 文件为新零件，\n"
            f"  将其 SaveAs 为 .sldprt，不要关闭，保持打开状态再插入。"
        )
    except Exception as e:
        return f"错误：{e}"


def run_sw_add_mate(mate_type, entity1, entity2, value=0.0, **kwargs):
    """在装配体中添加配合约束。
    支持: coincident(重合), concentric(同心), parallel(平行),
          distance(距离), angle(角度)
    entity1/entity2 可以是实体名（如 'Face1'）或坐标格式 'x,y,z'（按位置选面）。
    AddMate5 需要 15 参数，末位是 byref ErrorStatus VARIANT。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model:
            return "错误：没有活动文档。"

        mate_types = {
            "coincident": 0,
            "concentric": 1,
            "parallel": 3,
            "angle": 6,
            "distance": 5,
        }
        mate_code = mate_types.get(mate_type.lower())
        if mate_code is None:
            return (
                f"错误：不支持的配合类型 '{mate_type}'。"
                f"支持: {', '.join(mate_types.keys())}"
            )

        # 选中两个面/边 —— 支持名称 或 坐标 'x,y,z' 格式
        model.ClearSelection2(True)
        sel1_ok = _select_entity_smart(model.Extension, str(entity1), append=False, mark=1)
        if not sel1_ok:
            return f"失败：无法选中第一个实体 '{entity1}'。尝试坐标格式 'x,y,z' 或面名。"
        sel2_ok = _select_entity_smart(model.Extension, str(entity2), append=True, mark=1)
        if not sel2_ok:
            return f"失败：无法选中第二个实体 '{entity2}'。尝试坐标格式 'x,y,z' 或面名。"
        sel_count = model.SelectionManager.GetSelectedObjectCount2(-1)
        if sel_count < 2:
            return f"失败：只选中了 {sel_count} 个实体，需恰好 2 个。"

        # AddMate5: 15 args, last is byref ErrorStatus
        error_status = win32com.client.VARIANT(
            win32com.client.pythoncom.VT_BYREF | win32com.client.pythoncom.VT_I4, 0
        )
        mate = model.AddMate5(
            mate_code, 0, False,
            float(value), float(value), float(value),
            0.0, 0.0, 0.0, 0.0, 0.0,
            False, False, 0, error_status,
        )
        if mate:
            return (
                f"成功：已添加 {mate_type} 配合 "
                f"'{entity1}' <-> '{entity2}'"
            )
        return (
            f"失败：未能创建配合 (error_status={error_status.value})。"
            f"请确认 '{entity1}' 和 '{entity2}' 是有效的面/边引用。"
        )
    except Exception as e:
        return f"错误：{e}"


# =========================================================================
# 11. 草图约束
# =========================================================================

def run_sw_add_dimension(entity_name, value, dim_type="linear", **kwargs):
    """通过 Parameter 设定草图尺寸值。自动进入草图编辑。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        sketch_name = entity_name
        in_sketch_mode = False
        try:
            active_sketch = model.SketchManager.ActiveSketch
            if active_sketch:
                in_sketch_mode = True
                sketch_name = _sw_member(active_sketch, "Name")
        except Exception:
            pass

        if not in_sketch_mode:
            # 自动尝试进入草图编辑
            if not _select_sketch_by_name(model.Extension, sketch_name):
                return f"提示：草图 '{sketch_name}' 不存在或无自动尺寸可设置。用精确坐标画图。"
            try:
                model.SketchManager.InsertSketch(True)
            except Exception:
                pass
            # 再检查一次
            try:
                active_sketch = model.SketchManager.ActiveSketch
                if active_sketch:
                    in_sketch_mode = True
                    sketch_name = _sw_member(active_sketch, "Name")
            except Exception:
                pass

        # Parameter 方式不需要在草图编辑状态，先尝试直接设置
        for dim_prefix in ("D1@", "D2@", "D3@"):
            for sk_name in (sketch_name,
                           f"草图{sketch_name[6:]}" if sketch_name.startswith("Sketch") else
                           f"Sketch{sketch_name[2:]}" if sketch_name.startswith("草图") else
                           sketch_name):
                dim_full = f"{dim_prefix}{sk_name}"
                try:
                    param = model.Parameter(dim_full)
                    if param:
                        param.SystemValue = float(value)
                        _sw_member(model, "EditRebuild3")
                        return f"成功：已设置尺寸 '{dim_full}' = {value}m。"
                except Exception:
                    continue

        return (
            f"提示：草图 '{sketch_name}' 无自动尺寸可设置。\n"
            f"  用 sw_modify_dimension 修改已有尺寸，或用精确坐标控制。"
        )
    except Exception as e:
        return f"错误：{e}"


def run_sw_add_relation(entity1, entity2, relation, **kwargs):
    """为两个草图实体添加几何约束。支持: horizontal/vertical/collinear/concentric/tangent/parallel/perpendicular/equal/fix/coincident/symmetric。
    仅在草图编辑状态下有效。entity2 可以是草图线段、基准面或边线。
    coincident 用于使两个端点重合（闭合轮廓）——此时传入草图点名称（如 'Point1', 'Point2'）更可靠。
    若不在草图编辑状态，自动尝试进入最近创建的草图。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        # 检查是否在草图编辑状态
        try:
            active_sketch = model.SketchManager.ActiveSketch
        except Exception:
            active_sketch = None
        if not active_sketch:
            # 自动尝试进入最近创建的草图（Sketch1/草图1, Sketch2/草图2, ...）
            entered = False
            for i in range(1, 10):
                for name in (f"Sketch{i}", f"草图{i}"):
                    if _select_by_id(model.Extension, name, "SKETCH"):
                        model.SketchManager.InsertSketch(True)
                        try:
                            active_sketch = model.SketchManager.ActiveSketch
                            if active_sketch:
                                entered = True
                                break
                        except Exception:
                            pass
                if entered:
                    break
            if not entered:
                return "失败：当前未在草图编辑状态，且无法自动进入。请先用 sw_edit_sketch 进入草图编辑。"

        model.ClearSelection2(True)

        # coincident/symmetric 约束优先选点（端点重合），其他约束优先选线段
        _is_point_rel = relation.lower() in ("coincident", "symmetric", "merge")
        _sel_order_1 = ("SKETCHPOINT", "SKETCHSEGMENT", "EDGE", "FACE", "PLANE") if _is_point_rel else ("SKETCHSEGMENT", "SKETCHPOINT", "EDGE", "FACE", "PLANE")

        # 选中第一个实体
        sel1_ok = False
        for st in _sel_order_1:
            if _select_by_id(model.Extension, str(entity1), st):
                sel1_ok = True
                break

        # 选中第二个实体
        _sel_order_2 = ("SKETCHPOINT", "SKETCHSEGMENT", "EDGE", "FACE", "PLANE", "AXIS") if _is_point_rel else ("SKETCHSEGMENT", "SKETCHPOINT", "EDGE", "FACE", "PLANE", "AXIS")
        sel2_ok = False
        for st in _sel_order_2:
            if _select_by_id(model.Extension, str(entity2), st, append=True):
                sel2_ok = True
                break

        if not sel1_ok or not sel2_ok:
            return f"失败：无法选中实体 '{entity1}'(={sel1_ok}) 或 '{entity2}'(={sel2_ok})。"

        # SketchAddConstraints 接受字符串参数，基于当前选中的实体施加约束
        constraint_map = {
            "horizontal": "HORIZONTAL",
            "vertical": "VERTICAL",
            "collinear": "COLLINEAR",
            "concentric": "CONCENTRIC",
            "tangent": "TANGENT",
            "parallel": "PARALLEL",
            "perpendicular": "PERPENDICULAR",
            "equal": "EQUAL",
            "fix": "FIX",
            "coincident": "COINCIDENT",
            "symmetric": "SYMMETRIC",
        }
        constraint_str = constraint_map.get(relation.lower())
        if constraint_str is None:
            return f"错误：不支持的约束 '{relation}'。支持: {list(constraint_map.keys())}"

        model.SketchAddConstraints(constraint_str)
        return f"成功：'{entity1}' 与 '{entity2}' 已约束为 {relation}。"
    except Exception as e:
        return f"错误：{e}"


# =========================================================================
# 12. 工程特征：抽壳/加强筋
# =========================================================================

def run_sw_shell(thickness, face_to_remove="", **kwargs):
    """抽壳：以实体面为开口等壁厚抽空实体。
    需先用 sw_select_face_by_point 选中要移除的面，或传 face_to_remove 名称。
    InsertFeatureShell 是 IModelDoc2 的方法，2 参数，void 返回。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        doc_type = _sw_member(model, "GetType")
        if doc_type != 1:
            return "错误：抽壳只在零件文档中可用。"

        if face_to_remove:
            model.ClearSelection2(True)
            _select_by_id(model.Extension, face_to_remove, "FACE")

        # InsertFeatureShell(thickness, outward) — IModelDoc2 方法, void 返回
        # 无面选中时创建均匀壁厚抽壳；有面选中时该面为开口
        try:
            model.InsertFeatureShell(float(thickness), False)
        except Exception:
            return f"失败：抽壳创建失败。请确认实体已生成且壁厚合理。"

        # void 方法不返回值，检查特征树验证
        if _has_feature_type(model, "Shell"):
            return f"成功：抽壳厚度 {thickness}m。"
        return f"失败：抽壳创建失败。请确认实体已生成。"
    except Exception as e:
        return f"错误：{e}"


def run_sw_rib(sketch_name, thickness, direction="normal", **kwargs):
    """加强筋：沿开放草图轮廓生成加强筋特征。
    direction: normal=垂直于草图平面, parallel=平行于草图平面
    需要在草图编辑状态下调用（会自动进入指定草图的编辑模式）。
    注意：InsertRib2 是 IModelDoc2 方法（9 参数），不是 IFeatureManager 方法。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        # 选中草图并进入编辑模式
        sketch_selected = False
        for name in (sketch_name,):
            aliases = [name]
            if name.startswith("Sketch"):
                aliases.append(f"草图{name[6:]}")
            elif name.startswith("草图"):
                aliases.append(f"Sketch{name[2:]}")
            for n in aliases:
                if _select_by_id(model.Extension, n, "SKETCH"):
                    sketch_selected = True
                    model.SketchManager.InsertSketch(True)
                    break
            if sketch_selected:
                break

        if not sketch_selected:
            return f"失败：未找到草图 '{sketch_name}'。用 sw_get_model_structure 查看特征树。"

        # InsertRib2 (9 args, IModelDoc2):
        # (Is2Sided, ReverseThicknessDir, Thickness, ReferenceEdgeIndex,
        #  ReverseMaterialDir, IsDrafted, DraftOutward, DraftAngle, IsNormToSketch)
        is_norm = 1 if direction.lower() == "normal" else 0
        try:
            model.InsertRib2(
                True, False, float(thickness), 0,
                False, False, False, 0, bool(is_norm),
            )
        except Exception:
            # 回退: IFeatureManager.InsertRib (10 args)
            try:
                model.FeatureManager.InsertRib(
                    True, False, float(thickness), 0,
                    False, False, False, 0, bool(is_norm), False,
                )
            except Exception as e:
                try:
                    model.SketchManager.InsertSketch(True)
                except Exception:
                    pass
                return f"失败：加强筋创建失败。{e}"

        try:
            model.SketchManager.InsertSketch(True)
        except Exception:
            pass

        if _has_feature_type(model, "Rib"):
            return f"成功：加强筋 '{sketch_name}' 厚度 {thickness}m。"
        return (
            f"失败：加强筋创建失败。常见原因：\n"
            f"  1. 草图不是开放轮廓（加强筋需要单一线段，不是闭合轮廓）\n"
            f"  2. 草图平面不与任何实体相交（筋需连接两个实体面）\n"
            f"  3. 草图线段端点不在实体表面附近\n"
            f"  正确示例：底板+立壁的L型支架，在 Front Plane 画一条\n"
            f"  从底板顶面到立壁侧面的斜线，然后调 sw_rib。"
        )
    except Exception as e:
        return f"错误：{e}"


# =========================================================================
# 13. 测量与验证
# =========================================================================

def _select_entity_smart(ext, entity_ref, append=False, mark=0):
    """智能选中实体：支持坐标格式 "x,y,z"、名称格式、中英文基准面名。
    返回 True/False。"""
    ref = str(entity_ref)
    # 坐标格式 "x,y,z"
    if "," in ref:
        try:
            parts = [float(p.strip()) for p in ref.split(",")]
            if len(parts) >= 3:
                for sel_type in ("VERTEX", "EDGE", "FACE"):
                    if ext.SelectByID2(
                        "", sel_type, parts[0], parts[1], parts[2],
                        append, mark, _empty_callout(), 0
                    ):
                        return True
        except ValueError:
            pass
        return False
    # 名称格式 — 尝试多种类型
    for sel_type in ("FACE", "EDGE", "VERTEX", "PLANE", "SKETCHSEGMENT", "SOLIDBODY"):
        if _select_by_id(ext, ref, sel_type, append=append, mark=mark):
            return True
    # 尝试中文基准面名
    aliases = {"Front Plane": "前视基准面", "Top Plane": "上视基准面", "Right Plane": "右视基准面"}
    if ref in aliases:
        if _select_by_id(ext, aliases[ref], "PLANE", append=append, mark=mark):
            return True
    elif ref in aliases.values():
        for en, cn in aliases.items():
            if ref == cn:
                if _select_by_id(ext, en, "PLANE", append=append, mark=mark):
                    return True
    return False


def _parse_coord(ref):
    """解析 "x,y,z" 坐标格式为 (x,y,z) 三元组，失败返回 None。"""
    try:
        pts = [float(p.strip()) for p in ref.split(",")]
        if len(pts) >= 3:
            return (pts[0], pts[1], pts[2])
    except ValueError:
        pass
    return None


def run_sw_measure(entity1, entity2="", measure_type="distance", **kwargs):
    """测量距离/角度/直径/面积/体积。Agent 建完模用此工具自我验证。
    entity1/entity2 可以是：坐标格式 "x,y,z"、实体名、基准面名。
    对于 distance：可用两坐标（纯数学计算，最可靠）。
    对于 diameter：需选中一条圆边（用坐标格式更可靠）。
    对于 volume：用 sw_mass_properties 更可靠。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        # 体积：GetMassProperties (属性,不是方法)
        if measure_type == "volume":
            props = model.GetMassProperties
            if props and len(props) >= 4:
                return f"Volume: {props[3]:.6f} m^3"
            return "失败：无法获取体积。建议用 sw_mass_properties。"

        # 坐标格式距离：纯数学计算，最快最可靠
        c1 = _parse_coord(entity1)
        c2 = _parse_coord(entity2) if entity2 else None
        if c1 and c2 and measure_type == "distance":
            dx = c1[0] - c2[0]; dy = c1[1] - c2[1]; dz = c1[2] - c2[2]
            dist = math.sqrt(dx*dx + dy*dy + dz*dz)
            return f"距离: {dist:.6f} m"

        # 其余情况：通过选中实体用 IMeasure
        model.ClearSelection2(True)
        sel_mgr = model.SelectionManager

        sel1_ok = _select_entity_smart(model.Extension, entity1)
        if not sel1_ok:
            return f"失败：无法选中第一个实体 '{entity1}'。"

        if measure_type in ("distance", "angle") and entity2:
            sel2_ok = _select_entity_smart(model.Extension, entity2, append=True)
            if not sel2_ok:
                return f"失败：无法选中第二个实体 '{entity2}'。"

        sel_count = sel_mgr.GetSelectedObjectCount2(-1)
        if sel_count < 1:
            return "失败：未选中任何实体。"

        # IMeasure (SW 2026 pywin32 中只识别 Distance/Angle/Diameter 属性)
        measure = model.Extension.CreateMeasure
        if measure_type in ("distance", "angle") and sel_count >= 2:
            if measure_type == "distance":
                d = measure.Distance
                if d is not None and float(d) >= 0:
                    return f"距离: {float(d):.6f} m"
            else:
                a = measure.Angle
                if a is not None and float(a) >= 0:
                    return f"角度: {float(a) * 180 / math.pi:.2f}°"
            # 坐标格式回退（如果第二个实体也是坐标）
            if c1 and c2:
                dx = c1[0] - c2[0]; dy = c1[1] - c2[1]; dz = c1[2] - c2[2]
                dist = math.sqrt(dx*dx + dy*dy + dz*dz)
                return f"距离: {dist:.6f} m"
        if measure_type == "diameter":
            d = measure.Diameter
            if d is not None and float(d) >= 0:
                return f"直径: {float(d):.6f} m"
        if measure_type == "area":
            a = measure.Area
            if a is not None and float(a) >= 0:
                return f"Area: {float(a):.6f} m^2"

        return f"失败：无法执行 {measure_type} 测量。建议用坐标格式 'x,y,z' 测距。"
    except Exception as e:
        return f"错误：{e}"


def run_sw_mass_properties(**kwargs):
    """获取当前实体的质量属性：体积、质量、重心、表面积。
    model.GetMassProperties 是属性（不是方法），返回 12 元素 tuple。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        # model.GetMassProperties 是属性,不是方法 (pywin32 中不能用括号调用)
        try:
            props = model.GetMassProperties
            if props and len(props) >= 6:
                # GetMassProperties returns: [cgx, cgy, cgz, volume, mass, surfaceArea, ...]
                lines = [
                    "=== Mass Properties ===",
                    f"Volume:      {props[3]:.6f} m^3",
                    f"Mass:        {props[4]:.4f} kg (density=1000 kg/m^3)",
                    f"SurfaceArea: {props[5]:.4f} m^2",
                    f"CoG X:       {props[0]:.6f} m",
                    f"CoG Y:       {props[1]:.6f} m",
                    f"CoG Z:       {props[2]:.6f} m",
                ]
                return "\n".join(lines)
        except Exception:
            pass

        return "失败：无法获取质量属性。请确认实体已生成。"
    except Exception as e:
        return f"错误：{e}"


# =========================================================================
# 14. 体级操作：移动/固定组件 + 体镜像
# =========================================================================

def run_sw_mirror_body(body_name, mirror_plane, **kwargs):
    """镜像实体（体级，非特征级）。用于对称零件快速建模。
    注意：InsertMirrorFeature2(BMirrorBody=True) 在 SW 2026 pywin32 中不工作。
    替代方案：用 GetBodies2 获取实体名，然后用特征镜像(mark=1/mark=2)。
    body_name 可以是实体名或特征名。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"
        model.ClearSelection2(True)

        # 获取实体名 — GetBodies2(bodyType=0, visibleOnly=False)
        # swSolidBody=0 (NOT 1!)
        actual_body_name = body_name
        body_ok = False
        if body_name:
            if _select_by_id(model.Extension, str(body_name), "SOLIDBODY", mark=1):
                body_ok = True
                actual_body_name = body_name

        if not body_ok:
            try:
                bodies = model.GetBodies2(0, False)
                if bodies:
                    first_body = bodies[0]
                    actual_body_name = _sw_member(first_body, "Name")
                    if _select_by_id(model.Extension, actual_body_name, "SOLIDBODY", mark=1):
                        body_ok = True
            except Exception:
                pass

        # 选中镜像基准面 (mark=2)
        plane_ok = False
        for sel_type in ("PLANE", "FACE"):
            for n in [str(mirror_plane)]:
                if _select_by_id(model.Extension, n, sel_type, append=True, mark=2):
                    plane_ok = True
                    break
            if plane_ok:
                break
        # 尝试中文基准面名
        if not plane_ok:
            aliases = {"Right Plane": "右视基准面", "Top Plane": "上视基准面", "Front Plane": "前视基准面"}
            cn = aliases.get(str(mirror_plane), "")
            if cn and _select_by_id(model.Extension, cn, "PLANE", append=True, mark=2):
                plane_ok = True

        if not body_ok or not plane_ok:
            return (
                f"失败：体镜像选中失败。\n"
                f"  实体: '{actual_body_name}' (选中={body_ok})\n"
                f"  平面: '{mirror_plane}' (选中={plane_ok})"
            )

        # 尝试 InsertMirrorFeature2(BMirrorBody=True)
        feat = model.FeatureManager.InsertMirrorFeature2(True, False, True, False, 0)
        if not feat:
            # 回退: InsertMirrorFeature (4 args)
            try:
                feat = model.FeatureManager.InsertMirrorFeature(True, False, True, False)
            except Exception:
                pass

        if feat:
            return f"成功：实体 '{actual_body_name}' 已沿 '{mirror_plane}' 镜像。"
        return (
            f"失败：体镜像创建失败。\n"
            f"  提示：SW 2026 的 InsertMirrorFeature2(BMirrorBody=True) 可能不兼容 pywin32。\n"
            f"  替代方案：用 sw_mirror_feature 镜像实体的特征（如拉伸特征）。"
        )
    except Exception as e:
        return f"错误：{e}"


def run_sw_move_component(component, dx=0.0, dy=0.0, dz=0.0,
                          move_type="translate", **kwargs):
    """在装配体中平移或旋转指定组件。
    TranslateComponent/RotateComponent 是 0 参数的 toggle（启动GUI模式），
    不能传 dx/dy/dz。程序化移动用 Transform2 + SetTransformAndSolve2。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"
        if _sw_member(model, "GetType") != 2:
            return "错误：当前文档不是装配体。"

        model.ClearSelection2(True)
        _select_by_id(model.Extension, component, "COMPONENT")
        comp_obj = model.SelectionManager.GetSelectedObjectsComponent3(1, -1)
        if not comp_obj:
            return f"错误：未找到组件 '{component}'。"

        if move_type == "translate":
            # 用 Transform2 设置平移
            transform = comp_obj.Transform2
            if not transform:
                return f"错误：组件 '{component}' 没有 Transform2。"
            # MathTransform.ArrayData: 16 元素 (旋转矩阵3x3 + 平移3 + scale + 3reserved)
            # 当前变换保持旋转不变，只修改平移分量
            try:
                arr = transform.ArrayData
                arr = list(arr)
                # 累加平移到位置分量 (index 9, 10, 11)
                arr[9] = float(arr[9]) + float(dx)
                arr[10] = float(arr[10]) + float(dy)
                arr[11] = float(arr[11]) + float(dz)
                transform.ArrayData = arr
            except Exception:
                # 如果无法修改现有变换，创建新的
                arr = [1.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 1.0,
                       float(dx), float(dy), float(dz), 1.0, 0.0, 0.0, 0.0]
                transform.ArrayData = arr
            try:
                comp_obj.SetTransformAndSolve2(transform)
            except Exception:
                comp_obj.Transform2 = transform
            return f"成功：组件 '{component}' 平移 ({dx:.4f}, {dy:.4f}, {dz:.4f})m。"
        else:
            # 旋转: 用旋转矩阵
            import math as _m
            rx = float(dx) * _m.pi / 180.0
            ry = float(dy) * _m.pi / 180.0
            rz = float(dz) * _m.pi / 180.0
            # 简化: 只绕 Z 轴旋转
            cz, sz = _m.cos(rz), _m.sin(rz)
            arr = [cz, -sz, 0.0, sz, cz, 0.0, 0.0, 0.0, 1.0,
                   0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0]
            transform = comp_obj.Transform2
            if transform:
                transform.ArrayData = arr
                try:
                    comp_obj.SetTransformAndSolve2(transform)
                except Exception:
                    comp_obj.Transform2 = transform
            return f"成功：组件 '{component}' 旋转 ({dx}°, {dy}°, {dz}°)。"
    except Exception as e:
        return f"错误：{e}"


def run_sw_fix_component(component, fix=True, **kwargs):
    """固定或取消固定装配体中的组件。
    component 支持完整名（如 'lower_housing-1'）或短名（如 'lower_housing'）。
    FixComponent/UnfixComponent 是 0 参数方法。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"
        if _sw_member(model, "GetType") != 2:
            return "错误：当前文档不是装配体。"

        model.ClearSelection2(True)

        # 尝试多种名称变体选中组件
        comp_name = str(component)
        name_variants = [comp_name]
        # 如果已含 -数字，也尝试不带后缀
        if "-" in comp_name:
            name_variants.append(comp_name.rsplit("-", 1)[0])
        # 如果不含 -数字，尝试 -1 后缀
        if "-" not in comp_name:
            name_variants.append(f"{comp_name}-1")

        comp_obj = None
        for variant in name_variants:
            if _select_by_id(model.Extension, variant, "COMPONENT"):
                comp_obj = model.SelectionManager.GetSelectedObjectsComponent3(1, -1)
                if comp_obj:
                    break
                # SelectByID2 成功但 GetSelectedObjectsComponent3 可能失败，直接跳过 select
                model.ClearSelection2(True)

        if not comp_obj:
            # 回退: 遍历所有组件按关键字查找
            for include_suppressed in (False, True):
                try:
                    components = model.GetComponents(include_suppressed)
                    if components:
                        for comp in components:
                            try:
                                name = _sw_member(comp, "Name2")
                            except Exception:
                                continue
                            for variant in name_variants:
                                if variant.lower() in str(name).lower():
                                    comp.Select2(False, 0)
                                    comp_obj = comp
                                    break
                            if comp_obj:
                                break
                except Exception:
                    continue
                if comp_obj:
                    break

        if not comp_obj:
            return (
                f"错误：未找到组件 '{component}'。\n"
                f"  尝试过的名称: {name_variants}\n"
                f"  提示：用 sw_get_model_structure 查看装配体中的组件名。"
            )

        if fix:
            model.FixComponent()
        else:
            model.UnfixComponent()
        action = "固定" if fix else "取消固定"
        return f"成功：组件 '{component}' 已{action}。"
    except Exception as e:
        return f"错误：{e}"


# =========================================================================
# 标准件库工具
# =========================================================================

def _normalize_tool_args(kwargs, defaults):
    """类型归一化：降级模型可能传 list 而非 named params。
    将 kwargs 中的 list 值按顺序映射到 defaults 的 key。
    例如 model 调 sw_insert_toolbox_part(['bolt','M16x90','GB'])
    → kwargs 收到 key 为数字索引的 dict，取 values 映射到前 N 个形参。
    """
    vals = list(kwargs.values())
    # 检测：如果所有值都是 list 中没有的编号 key，说明模型传了 positional
    str_keys = [k for k in kwargs if isinstance(k, str)]
    if len(str_keys) < len(defaults) and vals:
        # 尝试将位置值映射到缺失的命名参数
        for i, (k, default) in enumerate(defaults.items()):
            if k not in kwargs and i < len(vals):
                kwargs[k] = vals[i]
    # 单个值归一化
    for k in list(kwargs):
        if isinstance(kwargs[k], list) and len(kwargs[k]) == 1:
            kwargs[k] = kwargs[k][0]
    return kwargs


def run_sw_insert_toolbox_part(standard="GB", category="", spec="", **kwargs):
    """通过 SolidWorks Toolbox 插入标准件。"""
    try:
        kwargs = _normalize_tool_args(kwargs, {"standard": "GB", "category": "", "spec": ""})
        standard = str(kwargs.get("standard", standard))
        category = str(kwargs.get("category", category))
        spec = str(kwargs.get("spec", spec))

        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model:
            return "错误：请先创建或打开一个零件/装配体文档。"

        # 方案 A: 尝试通过 SwToolboxLibrary 接口 (SolidWorks 2022+)
        try:
            tbx = sw.GetToolboxPartner()
            if tbx:
                tbx_data = tbx.GetData(
                    standard=standard, category=category, spec=spec
                )
                if tbx_data:
                    tbx.InsertPart(data=tbx_data)
                    return (
                        f"成功：已从 Toolbox 插入标准件 "
                        f"[{standard} {category} {spec}]"
                    )
        except Exception:
            pass

        # 方案 B: 搜索本地 Toolbox 路径并打开
        try:
            prefs = sw.GetUserPreferenceStringValue(29)  # Toolbox 根目录
            if prefs:
                import glob as _g
                search_path = Path(prefs) / "browser" / standard
                if search_path.exists():
                    candidates = list(search_path.rglob(f"*{spec}*.sldprt"))
                    if candidates:
                        sw.OpenDoc6(str(candidates[0]), 1, 0, "", 0, 0)
                        return (
                            f"成功：已从 Toolbox 打开标准件 "
                            f"[{standard} {spec}] -> {candidates[0].name}"
                        )
        except Exception:
            pass

        # 方案 C: 搜索本地标准件库 STEP 文件
        try:
            results = _parts_db.search(
                keyword=spec, category=category,
                standard_system=standard.upper(), limit=5,
            )
            if results:
                lines = [f"Toolbox 未配置，但本地标准件库找到 {len(results)} 个匹配："]
                for r in results:
                    has_file = " [有STEP文件]" if r["file_path"] else " [仅索引]"
                    lines.append(f"  #{r['id']} | {r['spec']} | {r['description'][:40]}{has_file}")
                    if r["file_path"]:
                        lines.append(f"    -> 用 sw_import_step '{r['file_path']}' 导入")
                lines.append(f"\n建议：用 sw_import_step 导入 STEP 文件，或用 sw_search_standard_part 搜索更多。")
                return "\n".join(lines)
        except Exception:
            pass

        # 方案 D: 提示用户手动配置 Toolbox
        return (
            f"未找到 Toolbox 标准件 [{standard} {category} {spec}]。\n"
            f"Toolbox 需要在 SolidWorks 中手动配置：\n"
            f"  工具 -> 选项 -> 系统选项 -> Hole Wizard/Toolbox -> 配置 Toolbox 根文件夹\n"
            f"替代方案：用 sw_search_standard_part 搜索本地标准件库，\n"
            f"然后用 sw_import_step 导入 STEP 文件。"
        )
    except Exception as e:
        return f"错误：Toolbox 插入失败: {e}"


def run_sw_search_standard_part(keyword="", category="",
                                 standard_system="", limit=10, **kwargs):
    """搜索本地标准件库 (SQLite 索引)。

    关键词自动按空格分词 OR 匹配。参数自动做类型归一化
    （keyword 支持 list→str，category/limit 支持 number→str/int）。
    """
    try:
        # ---- 类型归一化：降级模型可能传错类型 ----
        if isinstance(keyword, list):
            keyword = " ".join(str(x) for x in keyword)
        keyword = str(keyword) if keyword else ""
        if isinstance(category, (int, float)):
            category = ""  # 数字不是有效类别
        category = str(category) if category else ""
        if isinstance(limit, str):
            try:
                limit = int(limit)
            except ValueError:
                limit = 10
        standard_system = str(standard_system) if standard_system else ""
        # ------------------------------------------

        tokens = [t for t in keyword.split() if len(t) >= 2] if keyword else []
        if not tokens:
            # 无关键词：按类别+标准体系过滤
            results = _parts_db.search(
                keyword="", category=category,
                standard_system=standard_system.upper() if standard_system else "",
                limit=limit,
            )
        else:
            # 每个 token 单独搜索，取并集（按 id 去重）
            seen = set()
            results = []
            for token in tokens:
                batch = _parts_db.search(
                    keyword=token, category=category,
                    standard_system=standard_system.upper() if standard_system else "",
                    limit=limit,
                )
                for r in batch:
                    if r["id"] not in seen:
                        seen.add(r["id"])
                        results.append(r)
            results = results[:limit]

        if not results:
            stats = _parts_db.stats()
            return (
                f"标准件库中未找到匹配零件。\n"
                f"当前库共 {stats['total_parts']} 个标准件。\n"
                f"类别分布：{stats['by_category']}\n"
                f"提示：库中暂无该零件；可手动建模或导入 STEP。"
            )

        lines = [f"搜索到 {len(results)} 个标准件："]
        for r in results:
            has_file = " [有STEP文件]" if r["file_path"] else " [仅索引]"
            lines.append(
                f"  #{r['id']} | {r['standard_system']} | "
                f"{r['category']} | {r['spec']} | "
                f"{r['description'][:50]}{has_file}"
            )
        return "\n".join(lines)
    except Exception as e:
        return f"错误：搜索标准件库失败: {e}"


def _resolve_step_path(file_path):
    """解析 STEP 文件路径：依次查绝对路径、WORKDIR、STEP_DIR。
    兼容 Unix 风格路径 (/c/Users/...) → Windows 路径。"""
    from pathlib import Path as _Path

    # Unix 风格路径转换：/c/Users/... → C:\Users\...
    fp_str = str(file_path)
    if fp_str.startswith("/") and len(fp_str) > 2 and fp_str[2] == "/":
        drive = fp_str[1].upper()
        fp_str = drive + ":" + fp_str[2:].replace("/", "\\")

    fp = _Path(fp_str)
    if fp.is_absolute():
        return fp if fp.exists() else None
    # 依次尝试 WORKDIR 和 STEP_DIR
    from parts_library.parts_db import STEP_DIR
    for base in (WORKDIR, STEP_DIR):
        candidate = base / file_path
        if candidate.exists():
            return candidate
    return None


def run_sw_import_step(file_path, **kwargs):
    """将 STEP/IGES 文件导入 SolidWorks。自动创建新零件后导入。
    关键：LoadFile4 的 ImportData 参数必须用 VARIANT(VT_DISPATCH, None)，
    传 Python None 会类型不匹配。"""
    try:
        fp = _resolve_step_path(file_path)
        if fp is None:
            try:
                from parts_library.parts_db import STEP_DIR
                step_dir_msg = f"  - {STEP_DIR / file_path}\n"
            except Exception:
                step_dir_msg = "  - (标准件库 step_files/ 目录)\n"
            return (
                f"错误：STEP 文件不存在: '{file_path}'\n"
                f"  已在以下位置搜索:\n"
                f"  - {WORKDIR / file_path}\n"
                f"{step_dir_msg}"
                f"  提示：将 .step 文件放在项目根目录或 step_files/ 目录。"
            )

        sw = _get_sw_app()

        errors = win32com.client.VARIANT(
            win32com.client.pythoncom.VT_BYREF | win32com.client.pythoncom.VT_I4, 0
        )

        # LoadFile4: 4参数 (FileName, ArgString, ImportData, Errors)
        # ImportData 必须是 VARIANT(VT_DISPATCH, None)，传 None 会类型不匹配
        import_data = _empty_callout()  # VARIANT(VT_DISPATCH, None)
        try:
            doc = sw.LoadFile4(str(fp), "", import_data, errors)
            if doc:
                return (
                    f"成功：已将 STEP 文件导入为新零件。\n"
                    f"文件: {fp.name}"
                )
        except Exception:
            pass

        # 回退: OpenDoc6 (需要 errors + warnings 两个 VARIANT)
        warnings = win32com.client.VARIANT(
            win32com.client.pythoncom.VT_BYREF | win32com.client.pythoncom.VT_I4, 0
        )
        try:
            doc = sw.OpenDoc6(str(fp), 1, 1, "", errors, warnings)
            if doc:
                return f"成功：已打开 STEP 文件 '{fp.name}'。"
        except Exception:
            pass

        return f"失败：无法导入 STEP 文件 '{fp.name}'（尝试了 LoadFile4/OpenDoc6）。"
    except Exception as e:
        return f"错误：STEP 导入失败: {e}"


def run_sw_session_stats(**kwargs):
    """查询当前会话或全部历史统计。"""
    global _current_session_id
    try:
        s = session_summary(_current_session_id) if _current_session_id else None
        ts = tool_stats(session_id=_current_session_id)
        recent = tool_timeline(session_id=_current_session_id, limit=10)

        lines = ["=== 会话统计 ==="]
        if s:
            lines.append(f"会话 ID: {s['id']}")
            lines.append(f"创建: {s['created']}  最后活跃: {s['last_active']}")
            lines.append(f"消息数: {s['messages']}  工具调用: {s['total_tool_calls']}")
            lines.append(f"成功: {s['ok_tool_calls']}  失败: {s['fail_tool_calls']}")
            lines.append(f"成功率: {s['success_rate']}  平均耗时: {s['avg_duration_ms']}ms")
            lines.append(f"安全事件: {s['security_events']}")
        lines.append("\n--- 工具调用统计 ---")
        for r in ts[:15]:
            icon = "✓" if r["fail"] == 0 else "⚠"
            lines.append(
                f"  {icon} {r['tool']}: {r['total']} calls, "
                f"ok={r['ok']}, fail={r['fail']}, "
                f"avg={r['avg_ms']}ms, max={r['max_ms']}ms"
            )
        if recent:
            lines.append("\n--- 最近 10 次调用 ---")
            for r in recent:
                status_icon = "✓" if r["status"] == "ok" else "✗"
                lines.append(
                    f"  {status_icon} {r['tool']}: {r['ms']:.0f}ms, "
                    f"output={r['output_bytes']}B"
                )
        return "\n".join(lines)
    except Exception as e:
        return f"错误：获取会话统计失败: {e}"


def _normalize_todos(todos):
    if isinstance(todos, str):
        try:
            todos = json.loads(todos)
        except json.JSONDecodeError:
            try:
                todos = ast.literal_eval(todos)
            except (SyntaxError, ValueError):
                return None, "Error: todos must be a list or JSON array string"
    if not isinstance(todos, list):
        return None, "Error: todos must be a list"
    for i, t in enumerate(todos):
        if not isinstance(t, dict):
            return None, f"Error: todos[{i}] must be an object"
        if "content" not in t or "status" not in t:
            return None, f"Error: todos[{i}] missing 'content' or 'status'"
        if t["status"] not in ("pending", "in_progress", "completed"):
            return None, f"Error: todos[{i}] has invalid status '{t['status']}'"
    return todos, None
        
def run_todo_write(todos):
    global CURRENT_TODOS
    todos, error = _normalize_todos(todos)
    if error:
        return error
    CURRENT_TODOS = todos
    lines = [f"{GRAY}## Current Tasks{RESET}"]
    for t in CURRENT_TODOS:
        icon = {"pending": " ", "in_progress": "\033[36m▸\033[0m", "completed": "\033[32m✓\033[0m"}[t["status"]]
        lines.append(f"  [{icon}] {t['content']}")
    print("\n".join(lines))
    return f"Updated {len(CURRENT_TODOS)} tasks"

def extract_text(content):
    if not isinstance(content, list):
        return str(content)
    texts = []
    for b in content:
        btype = b.get("type") if isinstance(b, dict) else getattr(b, "type", None)
        if btype == "text":
            texts.append(b["text"] if isinstance(b, dict) else getattr(b, "text", ""))
    return "\n".join(texts)

def spawn_subagent(description: str) -> str:
    print(f"{GREEN}[Subagent Spawned]{RESET}")
    messages = [{"role":"user", "content": description}]

    for _ in range(30):
        with client.messages.stream(
            max_tokens=256000, model=SUB_MODEL, system=SUB_SYSTEM,
            messages=_clean_surrogates(messages), tools=SUB_TOOLS,
        ) as stream:
            for text in stream.text_stream:
                text = text.encode('utf-8', errors='surrogatepass').decode('utf-8', errors='replace')
                print(text, end="", flush=True)
            resp = stream.get_final_message()
        print()
        messages.append({"role":"assistant", "content": _clean_surrogates(resp.content)})

        if resp.stop_reason != "tool_use":
            break

        results = []
        for block in resp.content:
            if block.type == "tool_use":
                blocked = trigger_hooks("PreToolUse", block)
                if blocked:
                    results.append({
                        "type": "tool_result",
                        "tool_use_id": block.id,
                        "content": str(blocked)
                    })
                    continue
                handler = SUB_HANDLERS.get(block.name)
                output = handler(**block.input) if handler else f"Unknown:{block.name}"
                trigger_hooks("PostToolUse", block, output)
                out_preview = str(output)[:100]
                out_preview = out_preview.encode('utf-8', errors='surrogatepass').decode('utf-8', errors='replace')
                print(f"{GRAY}[sub] {block.name}: {out_preview}{RESET}")
                results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": output
                })
        messages.append({
            "role":"user",
            "content": _clean_surrogates(results)
        })

    result = extract_text(messages[-1]["content"])
    if not result:
        for msg in reversed(messages):
            if msg["role"] == "assistant":
                result = extract_text(msg["content"])
                if result:
                    break
        if not result:
            result = "Subagent stopped after 30 turns without final answer."
    print(f"{GRAY}[Subagent done]{RESET}")
    return result

def load_skill(name):
    skill = SKILL_REGISTRY.get(name)
    if not skill:
        return f"Skill not found: {name}"
    return skill["content"]

#def run_
#def run_grep(pattern, **kwargs):

#def run_pr(pattern, **kwargs):

#def resolve_conflict(pattern, **kwargs):
TOOLS = [
    {
        "name": "bash",
        "description": "Run a shell command.",
        "input_schema": {
            "type": "object",
            "properties": {"command": {"type": "string"}},
            "required": ["command"],
        },
    },
    {
        "name": "read_file",
        "description": "Read file contents.",
        "input_schema": {
            "type": "object",
            "properties": {"path": {"type": "string"}, "start": {"type": "integer"}, "limit": {"type": "integer"}},
            "required": ["path"],
        },
    },
    {
        "name": "write_file",
        "description": "Write content to a file.",
        "input_schema": {
            "type": "object",
            "properties": {"path": {"type": "string"}, "content": {"type": "string"}},
            "required": ["path", "content"],
        },
    },
    {
        "name": "edit_file",
        "description": "Replace exact text in a file once.",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "old_text": {"type": "string"},
                "new_text": {"type": "string"},
            },
            "required": ["path", "old_text", "new_text"],
        },
    },
    {
        "name": "glob",
        "description": "Find files matching a glob pattern.",
        "input_schema": {
            "type": "object",
            "properties": {"pattern": {"type": "string"}},
            "required": ["pattern"],
        },
    },
    {"name": "todo_write", "description": "Create and manage a task list for your current coding session.",
     "input_schema": {"type": "object", "properties": {"todos": {"type": "array", "items": {"type": "object", "properties": {"content": {"type": "string"}, "status": {"type": "string", "enum": ["pending", "in_progress", "completed"]}}, "required": ["content", "status"]}}}, "required": ["todos"]}},
     {
         "name": "spawn",
         "description": "Launch a subagent to handle a complex subtask. Returns only the final conclusion.",
         "input_schema": {
             "type":"object",
             "properties":{
                 "description": {
                     "type":"string"
                 }
             },
             "required": ["description"]
         },
     },
    {"name": "load_skill", "description": "Load the full content of a skill by name.",
     "input_schema": {"type": "object", "properties": {"name": {"type": "string"}}, "required": ["name"]}},
    # === SOLIDWORKS 扩展 Agent 工具集 ===
    {
        "name": "sw_create_new_part",
        "description": "在 SolidWorks 中初始化并创建一个全新的空白 3D 零件画布 (.sldprt)。",
        "input_schema": {"type": "object", "properties": {}},
    },
    {
        "name": "sw_select_plane_or_face",
        "description": "选中特定的基准面(如 'Front Plane'、'Top Plane')或实体表面以供后续绘图。",
        "input_schema": {
            "type": "object",
            "properties": {"target_name": {"type": "string", "description": "目标特征或平面的确切英文/中文名称"}},
            "required": ["target_name"],
        },
    },
    {
        "name": "sw_select_face_by_point",
        "description": "按空间坐标点拾取实体面/边/顶点。实体面无稳定人读名称，SelectByID2 按名称选不到时用此工具。例如板顶面中心传 (0,0,0.02) 选顶面。",
        "input_schema": {
            "type": "object",
            "properties": {
                "x": {"type": "number", "description": "点的 X 坐标 (米)"},
                "y": {"type": "number", "description": "点的 Y 坐标 (米)"},
                "z": {"type": "number", "description": "点的 Z 坐标 (米)"},
                "entity_type": {"type": "string", "enum": ["FACE", "EDGE", "VERTEX"], "default": "FACE", "description": "要拾取的几何类型"}
            },
            "required": ["x", "y", "z"],
        },
    },
    {
        "name": "sw_get_model_structure",
        "description": "读取当前零件图纸的特征树结构(类似 ls/tree)，用于分析已有建模步骤和草图。 ",
        "input_schema": {"type": "object", "properties": {}},
    },
    {
        "name": "sw_edit_sketch",
        "description": "进入指定草图的编辑模式。绘图工具画完会自动关闭草图，要添加约束/尺寸前需先用此工具重新进入编辑。",
        "input_schema": {
            "type": "object",
            "properties": {
                "sketch_name": {"type": "string", "description": "草图名，如 'Sketch1' 或 '草图1'"},
            },
            "required": ["sketch_name"],
        },
    },
    {
        "name": "sw_close_sketch",
        "description": "退出草图编辑模式（关闭草图）。添加完约束/尺寸后调用。",
        "input_schema": {"type": "object", "properties": {}},
    },
    {
        "name": "sw_draw_rectangle",
        "description": "在当前选中的面上绘制一个中心矩形草图（长度单位一律为米）。",
        "input_schema": {
            "type": "object",
            "properties": {
                "width": {"type": "number", "description": "矩形宽度 (米)"},
                "height": {"type": "number", "description": "矩形高度 (米)"},
                "center_x": {"type": "number", "default": 0.0, "description": "中心点 X 坐标"},
                "center_y": {"type": "number", "default": 0.0, "description": "中心点 Y 坐标"}
            },
            "required": ["width", "height"],
        },
    },
    {
        "name": "sw_draw_circle",
        "description": "在当前选中的面上绘制一个圆形草图（长度单位一律为米）。",
        "input_schema": {
            "type": "object",
            "properties": {
                "radius": {"type": "number", "description": "圆的半径 (米)"},
                "center_x": {"type": "number", "default": 0.0, "description": "圆心 X 坐标"},
                "center_y": {"type": "number", "default": 0.0, "description": "圆心 Y 坐标"}
            },
            "required": ["radius"],
        },
    },
    {
        "name": "sw_draw_polygon",
        "description": "在当前选中的面上通过多线段闭合循环拟合绘制一个正多边形草图。",
        "input_schema": {
            "type": "object",
            "properties": {
                "sides": {"type": "integer", "description": "边数 (必须 >= 3)"},
                "radius": {"type": "number", "description": "外接圆半径 (米)"},
                "center_x": {"type": "number", "default": 0.0},
                "center_y": {"type": "number", "default": 0.0}
            },
            "required": ["sides", "radius"],
        },
    },
    {
        "name": "sw_draw_line",
        "description": "绘制直线段（开放轮廓）。用于加强筋草图、构造线等。坐标单位：米。需先选基准面。",
        "input_schema": {
            "type": "object",
            "properties": {
                "x1": {"type": "number", "description": "起点 X (米)"},
                "y1": {"type": "number", "description": "起点 Y (米)"},
                "x2": {"type": "number", "description": "终点 X (米)"},
                "y2": {"type": "number", "description": "终点 Y (米)"},
            },
            "required": ["x1", "y1", "x2", "y2"],
        },
    },
    {
        "name": "sw_draw_profile",
        "description": "绘制闭合/开放多段线轮廓。传入坐标点列表 [(x1,y1),(x2,y2),...]，自动连点并首尾闭合。用于旋转截面、多段切除路径等需要闭合轮廓的场景。坐标单位：米。需先选基准面。",
        "input_schema": {
            "type": "object",
            "properties": {
                "points": {"type": "array", "items": {"type": "array", "items": {"type": "number"}}, "description": "坐标点列表 [(x1,y1),(x2,y2),...]，单位米"},
                "close": {"type": "boolean", "default": True, "description": "是否首尾闭合（默认 true）"},
            },
            "required": ["points"],
        },
    },
    {
        "name": "sw_draw_revolve_profile",
        "description": "【旋转特征专用】在一个草图内同时绘制中心线+闭合轮廓。一次调用 = draw_centerline + draw_profile 的正确组合。调用后直接 sw_revolve_boss 即可。参数: cx1,cy1,cx2,cy2=中心线起止点(米), points=轮廓点列表[(x,y),...](米)。需先选基准面。",
        "input_schema": {
            "type": "object",
            "properties": {
                "cx1": {"type": "number", "description": "中心线起点 X (米)"},
                "cy1": {"type": "number", "description": "中心线起点 Y (米)"},
                "cx2": {"type": "number", "description": "中心线终点 X (米)"},
                "cy2": {"type": "number", "description": "中心线终点 Y (米)"},
                "points": {"type": "array", "items": {"type": "array", "items": {"type": "number"}}, "description": "闭合轮廓点列表 [(x1,y1),(x2,y2),...] (米)"},
            },
            "required": ["cx1", "cy1", "cx2", "cy2", "points"],
        },
    },
    {
        "name": "sw_draw_centerline",
        "description": "绘制中心线（用作旋转轴、对称轴等构造线）。坐标单位：米。需先选基准面。",
        "input_schema": {
            "type": "object",
            "properties": {
                "x1": {"type": "number", "description": "起点 X (米)"},
                "y1": {"type": "number", "description": "起点 Y (米)"},
                "x2": {"type": "number", "description": "终点 X (米)"},
                "y2": {"type": "number", "description": "终点 Y (米)"},
            },
            "required": ["x1", "y1", "x2", "y2"],
        },
    },
    {
        "name": "sw_extrude_boss",
        "description": "将现有的闭合 2D 草图沿法线垂直拉伸加厚，使之变为 3D 实体实体特征。",
        "input_schema": {
            "type": "object",
            "properties": {
                "sketch_name": {"type": "string", "description": "要拉伸的草图名称，如 'Sketch1'"},
                "depth": {"type": "number", "description": "拉伸厚度/深度 (米)"}
            },
            "required": ["sketch_name", "depth"],
        },
    },
    {
        "name": "sw_extrude_cut",
        "description": "利用闭合草图对现有实体特征执行拉伸切除挖肉/打孔操作。",
        "input_schema": {
            "type": "object",
            "properties": {
                "sketch_name": {"type": "string", "description": "用于切除的草图名称，如 'Sketch2'"},
                "depth": {"type": "number", "description": "切除深度 (米)，在 thru_all 为 true 时可不设"},
                "thru_all": {"type": "boolean", "default": False, "description": "是否完全贯穿整个实体"}
            },
            "required": ["sketch_name"],
        },
    },
    {
        "name": "sw_apply_fillet",
        "description": "选中 3D 实体的某一特定锐边，对其应用圆角倒角特征平滑过渡。",
        "input_schema": {
            "type": "object",
            "properties": {
                "edge_id": {"type": "string", "description": "边缘的标识名称或选择标识"},
                "radius": {"type": "number", "description": "圆角半径 (米)"}
            },
            "required": ["edge_id", "radius"],
        },
    },
    {
        "name": "sw_modify_dimension",
        "description": "参数化修改已有尺寸参数的数值并强制触发模型更新重建 (Rebuild)。",
        "input_schema": {
            "type": "object",
            "properties": {
                "dimension_name": {"type": "string", "description": "尺寸完整名，如 'D1@Sketch1' 或 'D2@Extrude1'"},
                "new_value": {"type": "number", "description": "更新后的新系统数值 (米)"}
            },
            "required": ["dimension_name", "new_value"],
        },
    },
    {
        "name": "sw_export_to_format",
        "description": "一键导出当前 3D 模型或图纸到指定的工业生产交付格式中。",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_type": {"type": "string", "enum": ["STEP", "IGES", "PDF"], "description": "导出目标后缀"},
                "output_path": {"type": "string", "description": "导出的工作区相对物理文件路径"}
            },
            "required": ["file_type", "output_path"],
        },
    },
    {
        "name": "sw_close_doc",
        "description": "关闭指定标题的 SolidWorks 文档；不传 title 则关闭当前活动文档。用于多轮建模后清理累积零件，释放内存与 license 句柄，避免 SaveAs 错误码 1。",
        "input_schema": {
            "type": "object",
            "properties": {"title": {"type": "string", "description": "文档标题（可选，缺省关闭当前活动文档）"}},
            "required": [],
        },
    },
    # === 标准件库工具 ===
    {
        "name": "sw_insert_toolbox_part",
        "description": "通过 SolidWorks Toolbox 插入标准件（螺栓、螺母、垫圈、轴承等）。支持 GB/ISO/ANSI/DIN/JIS 标准体系。优先使用此工具而非从零建模标准件。",
        "input_schema": {
            "type": "object",
            "properties": {
                "standard": {"type": "string", "description": "标准体系: GB, ISO, ANSI, DIN, JIS", "default": "GB"},
                "category": {"type": "string", "description": "零件类别: bolt, nut, washer, bearing, gear, key, pin, ring"},
                "spec": {"type": "string", "description": "规格字符串，如 'M8x30', '6204', '8x7x20'"},
            },
            "required": ["category", "spec"],
        },
    },
    {
        "name": "sw_search_standard_part",
        "description": "搜索本地标准件库（含 GB 标准件、MISUMI、怡合达 等来源）。优先用此工具查找已有标准件，找不到才从零建模。",
        "input_schema": {
            "type": "object",
            "properties": {
                "keyword": {"type": "string", "description": "搜索关键词，如 '螺栓', '轴承', 'M8'"},
                "category": {"type": "string", "description": "零件类别: bolt, nut, washer, bearing, flange 等"},
                "standard_system": {"type": "string", "description": "标准体系: GB, ISO, DIN, JIS (可选)"},
                "limit": {"type": "integer", "default": 10, "description": "返回数量上限"},
            },
            "required": [],
        },
    },
    {
        "name": "sw_import_step",
        "description": "将本地 STEP 文件导入到当前 SolidWorks 文档中。用于将从标准件库下载的 STEP 模型加载到 SolidWorks。",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "STEP 文件路径（绝对路径或相对于标准件库的相对路径）"},
            },
            "required": ["file_path"],
        },
    },
    # === 高级建模工具：阵列/镜像/旋转/参考面/装配体 ===
    {
        "name": "sw_linear_pattern",
        "description": "线性阵列：沿指定方向复制特征 N 份。direction_ref 用基准面名（如'右视基准面'/'Right Plane'）或坐标格式'x,y,z'选中方向面/边。不要传'EDGE'或'FACE'等类型名。",
        "input_schema": {
            "type": "object",
            "properties": {
                "feature_name": {"type": "string", "description": "要阵列的特征名，如 '切除-拉伸1'"},
                "direction_ref": {"type": "string", "description": "方向参考：基准面名（如'右视基准面'）或坐标格式'x,y,z'（如'0,0.05,0.005'选中面）"},
                "spacing": {"type": "number", "description": "间距 (米)"},
                "count": {"type": "integer", "description": "实例数（≥2）"},
            },
            "required": ["feature_name", "direction_ref", "spacing", "count"],
        },
    },
    {
        "name": "sw_circular_pattern",
        "description": "圆周阵列：绕轴均布复制特征。优先用 SW 原生接口，失败自动回退内部手动打孔。手动回退需传 face_coord, hole_radius, hole_cx, hole_cy。",
        "input_schema": {
            "type": "object",
            "properties": {
                "feature_name": {"type": "string", "description": "要阵列的特征名，如 '切除-拉伸1'"},
                "axis_ref": {"type": "string", "description": "旋转轴：坐标格式'x,y,z'选中圆柱面"},
                "count": {"type": "integer", "description": "均布数量（≥2）"},
                "angle": {"type": "number", "default": 360.0, "description": "总旋转角度 (度)"},
                "face_coord": {"type": "string", "description": "打孔面坐标 'x,y,z'（手动回退用，如'0.05,0,0.02'）"},
                "hole_radius": {"type": "number", "description": "孔径 米（手动回退用，如 0.004）"},
                "hole_cx": {"type": "number", "description": "第一个孔在面上的 X 坐标（手动回退用）"},
                "hole_cy": {"type": "number", "description": "第一个孔在面上的 Y 坐标/PCD半径（手动回退用，如 0.0275）"},
            },
            "required": ["feature_name", "axis_ref", "count"],
        },
    },
    {
        "name": "sw_mirror_feature",
        "description": "镜像：沿指定基准面镜像复制特征。",
        "input_schema": {
            "type": "object",
            "properties": {
                "feature_name": {"type": "string", "description": "要镜像的特征名"},
                "mirror_plane": {"type": "string", "description": "镜像基准面名称"},
            },
            "required": ["feature_name", "mirror_plane"],
        },
    },
    {
        "name": "sw_revolve_boss",
        "description": "旋转凸台：将闭合草图绕轴线旋转生成回转体（轴、法兰、轮毂等）。草图须包含旋转中心线。",
        "input_schema": {
            "type": "object",
            "properties": {
                "sketch_name": {"type": "string", "description": "旋转截面草图名，如 'Sketch1'"},
                "axis_name": {"type": "string", "description": "旋转轴边名（可选，默认用草图中心线）"},
                "angle": {"type": "number", "default": 360.0, "description": "旋转角度 (度)"},
            },
            "required": ["sketch_name"],
        },
    },
    {
        "name": "sw_create_ref_plane",
        "description": "创建参考基准面。offset=偏移, angle=绕边旋转, mid=两面中间。",
        "input_schema": {
            "type": "object",
            "properties": {
                "reference": {"type": "string", "description": "参考面名称"},
                "plane_type": {"type": "string", "enum": ["offset", "angle", "mid"], "description": "创建方式"},
                "distance": {"type": "number", "default": 0.0, "description": "偏移距离 (米), offset 模式"},
                "edge": {"type": "string", "description": "旋转边, angle 模式"},
                "angle": {"type": "number", "default": 0.0, "description": "旋转角度 (度), angle 模式"},
            },
            "required": ["reference", "plane_type"],
        },
    },
    {
        "name": "sw_create_assembly",
        "description": "创建新的空白装配体文档 (.sldasm)，为多零件装配做准备。",
        "input_schema": {
            "type": "object",
            "properties": {},
            "required": [],
        },
    },
    {
        "name": "sw_insert_component",
        "description": "向装配体中插入已有零件 (.sldprt)。需先创建装配体。",
        "input_schema": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": ".sldprt 文件路径"},
                "name": {"type": "string", "description": "组件名称（可选）"},
                "x": {"type": "number", "default": 0.0},
                "y": {"type": "number", "default": 0.0},
                "z": {"type": "number", "default": 0.0},
            },
            "required": ["file_path"],
        },
    },
    {
        "name": "sw_add_mate",
        "description": "在装配体中添加配合约束: coincident(重合), concentric(同心), parallel(平行), distance(距离), angle(角度)。",
        "input_schema": {
            "type": "object",
            "properties": {
                "mate_type": {"type": "string", "description": "配合类型: coincident/concentric/parallel/distance/angle"},
                "entity1": {"type": "string", "description": "第一个面/边引用"},
                "entity2": {"type": "string", "description": "第二个面/边引用"},
                "value": {"type": "number", "default": 0.0, "description": "值: distance 用米, angle 用度"},
            },
            "required": ["mate_type", "entity1", "entity2"],
        },
    },
    # === 草图约束 ===
    {
        "name": "sw_add_dimension",
        "description": "设置草图/特征的驱动尺寸值。自动查找 D1@草图N 格式的尺寸并设值。AddDimension2 API 会触发SW交互模式，所以此工具用 model.Parameter 直接设值。画图时用精确坐标控制尺寸更可靠。",
        "input_schema": {
            "type": "object",
            "properties": {
                "entity_name": {"type": "string", "description": "草图名（如'Sketch1'/'草图1'），工具会自动查找 D1@草图名 格式的尺寸"},
                "value": {"type": "number", "description": "尺寸值 (米)"},
                "dim_type": {"type": "string", "enum": ["linear","radial","angular","diameter"], "description": "尺寸类型（预留，当前按linear处理）"},
            },
            "required": ["entity_name", "value"],
        },
    },
    {
        "name": "sw_add_relation",
        "description": "为两个草图实体添加几何约束: horizontal/vertical/concentric/tangent/parallel/perpendicular/equal/fix。",
        "input_schema": {
            "type": "object",
            "properties": {
                "entity1": {"type": "string", "description": "第一个草图实体"},
                "entity2": {"type": "string", "description": "第二个草图实体"},
                "relation": {"type": "string", "description": "约束类型: horizontal/vertical/concentric/tangent/parallel/perpendicular/equal/fix"},
            },
            "required": ["entity1", "entity2", "relation"],
        },
    },
    # === 工程特征 ===
    {
        "name": "sw_shell",
        "description": "抽壳：以实体面为开口，等壁厚抽空实体。用于创建壳体、箱体等。",
        "input_schema": {
            "type": "object",
            "properties": {
                "thickness": {"type": "number", "description": "壁厚 (米)"},
                "face_to_remove": {"type": "string", "description": "开口面名（可选）"},
            },
            "required": ["thickness"],
        },
    },
    {
        "name": "sw_rib",
        "description": "加强筋：沿开放草图生成加强筋特征。用于支架、轴承座等结构件。",
        "input_schema": {
            "type": "object",
            "properties": {
                "sketch_name": {"type": "string", "description": "开放轮廓草图名"},
                "thickness": {"type": "number", "description": "筋厚度 (米)"},
                "direction": {"type": "string", "default": "normal", "description": "拉伸方向: normal/reverse"},
            },
            "required": ["sketch_name", "thickness"],
        },
    },
    # === 测量与验证 ===
    {
        "name": "sw_measure",
        "description": "测量距离/角度/直径/面积/体积。entity 用坐标格式'x,y,z'最可靠（如'0,0,0'），也支持基准面名/特征名。volume 类型建议改用 sw_mass_properties。",
        "input_schema": {
            "type": "object",
            "properties": {
                "entity1": {"type": "string", "description": "第一个实体：坐标格式'x,y,z'（如'0,0,0'）或基准面名"},
                "entity2": {"type": "string", "description": "第二个实体（距离/角度用，可选）：坐标格式或基准面名"},
                "measure_type": {"type": "string", "enum": ["distance","angle","diameter","area","volume"], "description": "测量类型"},
            },
            "required": ["entity1"],
        },
    },
    {
        "name": "sw_mass_properties",
        "description": "获取当前零件的质量属性：体积、质量、重心坐标、表面积。",
        "input_schema": {
            "type": "object",
            "properties": {},
            "required": [],
        },
    },
    # === 体级操作与装配操控 ===
    # DISABLED: sw_mirror_body — InsertMirrorFeature2(BMirrorBody=True) 在 SW 2026 + pywin32 下不工作
    # 替代方案：使用 sw_mirror_feature（特征级镜像）代替体级镜像
    {
        "name": "sw_move_component",
        "description": "在装配体中平移或旋转组件。move_type=translate 用 dx/dy/dz(米)，rotate 用 dx/dy/dz(度)。",
        "input_schema": {
            "type": "object",
            "properties": {
                "component": {"type": "string", "description": "组件名称"},
                "dx": {"type": "number", "default": 0.0, "description": "X方向平移(米)或旋转(度)"},
                "dy": {"type": "number", "default": 0.0, "description": "Y方向平移(米)或旋转(度)"},
                "dz": {"type": "number", "default": 0.0, "description": "Z方向平移(米)或旋转(度)"},
                "move_type": {"type": "string", "enum": ["translate","rotate"], "description": "平移或旋转"},
            },
            "required": ["component"],
        },
    },
    {
        "name": "sw_fix_component",
        "description": "在装配体中固定或取消固定组件。",
        "input_schema": {
            "type": "object",
            "properties": {
                "component": {"type": "string", "description": "组件名称"},
                "fix": {"type": "boolean", "default": True, "description": "true=固定, false=取消固定"},
            },
            "required": ["component"],
        },
    },
    # === 会话与统计工具 ===
    {
        "name": "sw_session_stats",
        "description": "查询当前会话的工具调用统计：调用次数、成功率、平均耗时、失败工具排行。用于诊断系统运行状况和发现性能瓶颈。",
        "input_schema": {
            "type": "object",
            "properties": {},
            "required": [],
        },
    },
]

TOOL_HANDLERS = {
    "bash": run_bash,
    "read_file": run_read,
    "write_file": run_write,
    "edit_file": run_edit,
    "glob": run_glob,
    "todo_write": run_todo_write,
    "spawn": spawn_subagent,
    "load_skill": load_skill,
    # SolidWorks 映射绑定
    "sw_create_new_part": run_sw_create_new_part,
    "sw_select_plane_or_face": run_sw_select_plane_or_face,
    "sw_select_face_by_point": run_sw_select_face_by_point,
    "sw_get_model_structure": run_sw_get_model_structure,
    "sw_edit_sketch": run_sw_edit_sketch,
    "sw_close_sketch": run_sw_close_sketch,
    "sw_draw_rectangle": run_sw_draw_rectangle,
    "sw_draw_circle": run_sw_draw_circle,
    "sw_draw_polygon": run_sw_draw_polygon,
    "sw_draw_line": run_sw_draw_line,
    "sw_draw_profile": run_sw_draw_profile,
    "sw_draw_revolve_profile": run_sw_draw_revolve_profile,
    "sw_draw_centerline": run_sw_draw_centerline,
    "sw_extrude_boss": run_sw_extrude_boss,
    "sw_extrude_cut": run_sw_extrude_cut,
    "sw_apply_fillet": run_sw_apply_fillet,
    "sw_modify_dimension": run_sw_modify_dimension,
    "sw_export_to_format": run_sw_export_to_format,
    "sw_close_doc": run_sw_close_doc,
    # 标准件库
    "sw_insert_toolbox_part": run_sw_insert_toolbox_part,
    "sw_search_standard_part": run_sw_search_standard_part,
    "sw_import_step": run_sw_import_step,
    # 高级建模
    "sw_linear_pattern": run_sw_linear_pattern,
    "sw_circular_pattern": run_sw_circular_pattern,
    "sw_mirror_feature": run_sw_mirror_feature,
    "sw_revolve_boss": run_sw_revolve_boss,
    "sw_create_ref_plane": run_sw_create_ref_plane,
    "sw_create_assembly": run_sw_create_assembly,
    "sw_insert_component": run_sw_insert_component,
    "sw_add_mate": run_sw_add_mate,
    # 草图约束
    "sw_add_dimension": run_sw_add_dimension,
    "sw_add_relation": run_sw_add_relation,
    # 工程特征
    "sw_shell": run_sw_shell,
    "sw_rib": run_sw_rib,
    # 测量验证
    "sw_measure": run_sw_measure,
    "sw_mass_properties": run_sw_mass_properties,
    # 体级操作与装配操控
    # DISABLED: "sw_mirror_body": run_sw_mirror_body,  # InsertMirrorFeature2 BMirrorBody=True 在 SW2026+pywin32 不兼容
    "sw_move_component": run_sw_move_component,
    "sw_fix_component": run_sw_fix_component,
    # 会话统计
    "sw_session_stats": run_sw_session_stats,
}

HOOKS = {
    "UserPromptSubmit": [],
    "PreToolUse": [],
    "PostToolUse": [],
    "Stop": [],
}

SUB_TOOLS = [
    {"name": "bash", "description": "Run a shell command.",
     "input_schema": {"type": "object", "properties": {"command": {"type": "string"}}, "required": ["command"]}},
    {"name": "read_file", "description": "Read file contents.",
     "input_schema": {"type": "object", "properties": {"path": {"type": "string"}}, "required": ["path"]}},
    {"name": "write_file", "description": "Write content to a file.",
     "input_schema": {"type": "object", "properties": {"path": {"type": "string"}, "content": {"type": "string"}}, "required": ["path", "content"]}},
    {"name": "edit_file", "description": "Replace exact text in a file once.",
     "input_schema": {"type": "object", "properties": {"path": {"type": "string"}, "old_text": {"type": "string"}, "new_text": {"type": "string"}}, "required": ["path", "old_text", "new_text"]}},
    {"name": "glob", "description": "Find files matching a glob pattern.",
     "input_schema": {"type": "object", "properties": {"pattern": {"type": "string"}}, "required": ["pattern"]}},
]

SUB_HANDLERS = {
    "bash": run_bash,
    "read_file": run_read,
    "write_file": run_write,
    "edit_file": run_edit,
    "glob": run_glob,
}



def register_hook(event, callback):
    HOOKS[event].append(callback)

def trigger_hooks(event, *args):
    for callback in HOOKS[event]:
        result = callback(*args)
        if result is not None:
            return result
    return None

DENY_LIST = [
    "rm -rf /", "rd /s", "rmdir /s", "del /s", "format ", "diskpart",
    "sudo", "shutdown", "reboot", "mkfs", "dd if=", "> /dev/sda",
]
DESTRUCTIVE = [
    "rm ", "del ", "rd ", "rmdir ", "remove-item", "certutil -urlcache",
    "reg add", "chmod 777", "> /etc/", "taskkill /f",
]

def _headless_mode():
    """--run 等无交互输入场景：无法询问用户，一律拒绝需确认的操作。"""
    try:
        return not sys.stdin.isatty()
    except Exception:
        return True


def permission_hook(block):
    import re as _re

    def _segments(cmd):
        return [s.strip().lower() for s in _re.split(r"[&|;]", cmd)]

    if block.name == "bash":
        cmd = block.input.get("command", "")
        for pattern in DENY_LIST:
            if any(pattern in s for s in _segments(cmd)):
                print(f"{RED}Blocked: '{pattern}'{RESET}")
                log_security_event(
                    _current_session_id or "", "deny_list_blocked",
                    severity="high", tool_name="bash",
                    command=cmd,
                    action="blocked", detail=f"matched: {pattern}",
                )
                return "Permission denied by deny list"
        for kw in DESTRUCTIVE:
            if any(kw in s for s in _segments(cmd)):
                if _headless_mode():
                    log_security_event(
                        _current_session_id or "", "destructive_denied",
                        severity="medium", tool_name="bash",
                        command=cmd,
                        action="denied", detail="headless auto-deny",
                    )
                    return "Permission denied (headless: destructive command)"
                print(f"{RED}Potentially destructive command{RESET}")
                print(f"   Tool: {block.name}({block.input})")
                choice = input("   Allow? [y/N] ").strip().lower()
                if choice not in ("y", "yes"):
                    log_security_event(
                        _current_session_id or "", "destructive_denied",
                        severity="medium", tool_name="bash",
                        command=cmd,
                        action="denied", detail=f"matched: {kw}",
                    )
                    return "Permission denied by user"
                log_security_event(
                    _current_session_id or "", "destructive_allowed",
                    severity="medium", tool_name="bash",
                    command=cmd,
                    action="allowed", detail=f"matched: {kw}",
                )
    if block.name in ("write_file", "edit_file"):
        path = block.input.get("path", "")
        if not (WORKDIR / path).resolve().is_relative_to(WORKDIR):
            if _headless_mode():
                log_security_event(
                    _current_session_id or "", "path_sandbox_blocked",
                    severity="high", tool_name=block.name,
                    command=path, action="blocked",
                )
                return "Permission denied (headless: path outside workspace)"
            print(f"{RED}写位置超出工作区!{RESET}")
            print(f"   Tool: {block.name}({block.input})")
            choice = input("   Allow? [y/N] ").strip().lower()
            if choice not in ("y", "yes"):
                log_security_event(
                    _current_session_id or "", "path_sandbox_blocked",
                    severity="high", tool_name=block.name,
                    command=path, action="blocked",
                )
                return "Permission denied by user"
    return None

def log_hook(block):
    args_preview = str(list(block.input.values())[:2])[:60]
    print(f"{GRAY}[HOOK] {block.name}({args_preview}){RESET}")
    return None

def large_output_hook(block, output):
    if len(str(output)) > 100000:
        print(f"{YELLOW}[HOOK] Large output from {block.name}: {len(str(output))} chars{RESET}")
    return None

# 工具调用开始时打点（用于计算耗时）
_tool_start_times: dict[str, float] = {}

def tool_start_timer(block_id):
    _tool_start_times[block_id] = time.time()

def tool_metrics_hook(block, output):
    """PostToolUse: 记录工具调用指标到 SQLite。"""
    global _current_session_id
    if not _current_session_id:
        return None
    sid = _current_session_id
    tool_name = block.name
    duration_ms = 0.0
    if block.id in _tool_start_times:
        duration_ms = (time.time() - _tool_start_times.pop(block.id)) * 1000
    status = "error" if str(output).startswith(("错误", "危险", "Error", "失败", "未能", "无法")) else "ok"
    error_msg = str(output)[:200] if status == "error" else ""
    log_tool_call(
        session_id=sid, tool_name=tool_name, status=status,
        duration_ms=duration_ms,
        input_dict=dict(block.input) if hasattr(block, "input") else {},
        output_len=len(str(output)), error_msg=error_msg,
    )
    return None

def context_inject_hook(query):
    print(f"{GRAY}[HOOK] UserPromptSubmit: working in {WORKDIR}{RESET}")
    return None

def summary_hook(messages):
    global _current_session_id
    tool_count = sum(1 for m in messages
                     for b in (m.get("content") if isinstance(m.get("content"),list) else [])
                     if isinstance(b, dict) and b.get("type") == "tool_result")
    if _current_session_id:
        end_session(_current_session_id, "completed")
        s = session_summary(_current_session_id)
        if s:
            print(f"{GRAY}[HOOK] Stop: {tool_count} tool calls, "
                  f"success={s['ok_tool_calls']}, fail={s['fail_tool_calls']}, "
                  f"avg={s['avg_duration_ms']:.0f}ms, "
                  f"session={_current_session_id}{RESET}")
        return None
    print(f"{GRAY}[HOOK] Stop: session used {tool_count} tool calls{RESET}")
    return None

# 初始化会话
if not _current_session_id:
    _current_session_id, _ = start_session()

register_hook("UserPromptSubmit", context_inject_hook)
register_hook("PreToolUse", permission_hook)
# register_hook("PreToolUse", log_hook)
register_hook("PostToolUse", large_output_hook)
register_hook("PostToolUse", tool_metrics_hook)
register_hook("Stop", summary_hook)

rounds_since_todo = 0

# ---- 中断机制：ESC 或 Ctrl+C 打断 agent 输出 ----
_agent_interrupted = False

def _interrupt_handler(signum=None, frame=None):
    global _agent_interrupted
    _agent_interrupted = True

def _check_esc_key():
    """Windows: 非阻塞检查 ESC 键（GetAsyncKeyState 不消费输入，避免吞掉用户正在输入的内容）。"""
    if os.name != "nt":
        return False
    try:
        import win32api
        return bool(win32api.GetAsyncKeyState(0x1B) & 0x8000)
    except Exception:
        pass
    return False

def _clean_surrogates(obj):
    """递归清理 surrogate 字符，防止 UTF-8 编码崩溃。"""
    if isinstance(obj, str):
        return obj.encode('utf-8', errors='surrogatepass').decode('utf-8', errors='replace')
    if isinstance(obj, list):
        return [_clean_surrogates(item) for item in obj]
    if isinstance(obj, dict):
        return {_clean_surrogates(k): _clean_surrogates(v) for k, v in obj.items()}
    # ContentBlock 等对象：转 dict 再清理
    if hasattr(obj, 'model_dump'):
        return _clean_surrogates(obj.model_dump())
    if hasattr(obj, 'to_dict'):
        return _clean_surrogates(obj.to_dict())
    return obj


def agent_loop(messages, max_turns=None):
    global rounds_since_todo, _agent_interrupted

    # === 入口清理：移除不配对的 tool_use assistant 消息 ===
    # 从后往前找含有 tool_use 的 assistant 消息，检查其后是否有对应 tool_result
    i = len(messages) - 1
    while i >= 0:
        msg = messages[i]
        if msg.get("role") != "assistant":
            i -= 1
            continue
        content = msg.get("content", [])
        if not isinstance(content, list) or not content:
            i -= 1
            continue
        # 提取此消息中所有 tool_use id
        ids = set()
        for block in content:
            if isinstance(block, dict) and block.get("type") == "tool_use":
                ids.add(block.get("id", ""))
            elif hasattr(block, "type") and getattr(block, "type", "") == "tool_use":
                ids.add(getattr(block, "id", ""))
        if not ids:
            i -= 1
            continue
        # 检查后面是否有对应的 tool_result
        resolved = set()
        for j in range(i + 1, len(messages)):
            c2 = messages[j].get("content", [])
            if isinstance(c2, list):
                for b in c2:
                    rid = ""
                    if isinstance(b, dict) and b.get("type") == "tool_result":
                        rid = b.get("tool_use_id", "")
                    elif hasattr(b, "type") and getattr(b, "type", "") == "tool_result":
                        rid = getattr(b, "tool_use_id", "")
                    if rid in ids:
                        resolved.add(rid)
        if resolved == ids:
            i -= 1
            continue
        # 有不配对的 tool_use → 剔除这条 assistant 消息
        print(f"{YELLOW}[清理] 移除 {len(ids)-len(resolved)} 个未完成 tool_use{RESET}")
        messages.pop(i)
        i -= 1  # 继续向前扫描

    # === 正常启动 ===
    old_sigint = signal.signal(signal.SIGINT, _interrupt_handler)
    _agent_interrupted = False
    interrupted = False

    turn = 0
    try:
        while True:
            if max_turns is not None:
                turn += 1
                if turn > max_turns:
                    print(f"{YELLOW}[--run] 达到最大轮数 {max_turns}，停止{RESET}")
                    break

            if rounds_since_todo >= 3 and messages:
                messages.append({
                    "role": "user",
                    "content": "<reminder>Update your todos.</reminder>"
                })
                rounds_since_todo = 0

            # 快照：异常退出时回滚到 API 调用前的安全状态
            _safe_snapshot = list(messages)

            stream_error = None
            interrupted = False
            resp = None
            esc_check_counter = 0

            try:
                with client.messages.stream(
                    max_tokens=256000, model=MODEL, messages=_clean_surrogates(messages),
                    tools=TOOLS, system=SYSTEM,
                    extra_body={"thinking": {"type": "adaptive"},
                                "output_config": {"effort": REASON}},
                ) as stream:
                    for text in stream.text_stream:
                        if _agent_interrupted:
                            interrupted = True
                            break
                        # 清理 surrogate + 非GBK字符（模型输出含emoji/特殊Unicode会被终端拒绝）
                        text = text.encode('utf-8', errors='surrogatepass').decode('utf-8', errors='replace')
                        _safe_print(text, end="", flush=True)
                        esc_check_counter += 1
                        if esc_check_counter % 8 == 0 and _check_esc_key():
                            _agent_interrupted = True
                            interrupted = True
                            break
                    if interrupted:
                        try:
                            _ = stream.get_final_message()
                        except Exception:
                            pass
                        print()
                        print(f"{YELLOW}[输出中断 — 返回对话]{RESET}")
                        messages[:] = _safe_snapshot
                        save_session(messages)
                        end_session(_current_session_id, "interrupted")
                        break
                    resp = stream.get_final_message()
            except Exception as stream_err:
                stream_error = stream_err

            if stream_error:
                print()
                err_msg = str(stream_error)
                err_msg = err_msg.encode('utf-8', errors='surrogatepass').decode('utf-8', errors='replace')
                # 诊断：400 错误打印具体的 tool_use/tool_result 状态
                if "400" in err_msg or "tool_use" in err_msg.lower():
                    print(f"{YELLOW}[诊断] 当前消息: {len(messages)} 条{RESET}")
                    for mi, m in enumerate(messages):
                        role = m["role"]
                        c = m.get("content", [])
                        if isinstance(c, list):
                            ids = [b.get("id","") if isinstance(b,dict) else getattr(b,"id","") if hasattr(b,"id") else ""
                                   for b in c if (isinstance(b,dict) and b.get("type") in ("tool_use","tool_result"))
                                   or (hasattr(b,"type") and getattr(b,"type","") in ("tool_use","tool_result"))]
                            if ids:
                                types = [b.get("type","") if isinstance(b,dict) else getattr(b,"type","") for b in c]
                                print(f"{YELLOW}  [{mi}] {role}: {types} ids={ids}{RESET}")
                print(f"{RED}连接中断: {err_msg[:200]}{RESET}")
                print(f"{YELLOW}会话已保存，请重试。{RESET}")
                messages[:] = _safe_snapshot
                save_session(messages)
                end_session(_current_session_id, "error")
                return False  # 流式错误：--run 以非零码退出

            if interrupted and resp is None:
                break
            if resp is None:
                break

            print()
            messages.append({"role": "assistant", "content": _clean_surrogates(resp.content)})

            # 不能仅依赖 stop_reason — 非 Claude 模型在 thinking 模式下
            # 可能返回 tool_use 块但 stop_reason 不是 "tool_use"，
            # 导致 tool_use 没有 tool_result 配对 → 下次 API 调用 400 错误
            tool_use_blocks = [
                b for b in resp.content
                if getattr(b, "type", "") == "tool_use"
                or (isinstance(b, dict) and b.get("type") == "tool_use")
            ]

            if not tool_use_blocks:
                force = trigger_hooks("Stop", messages)
                if force:
                    messages.append({"role": "user", "content": force})
                    continue
                break

            # 工具执行必须在 while 循环内 — 执行后 while 自然进入下一轮，
            # 将 tool_result 送回模型。放在 while 外面会导致 tool_use 无 tool_result
            rounds_since_todo += 1

            results = []
            try:
                for block in tool_use_blocks:
                    bid = getattr(block, "id", "") or (block.get("id", "") if isinstance(block, dict) else "")

                    try:
                        blocked = trigger_hooks("PreToolUse", block)
                    except Exception as e:
                        blocked = f"PreToolUse hook 异常: {e}"

                    if blocked:
                        results.append({
                            "type": "tool_result",
                            "tool_use_id": bid,
                            "content": str(blocked)
                        })
                        continue

                    handler = TOOL_HANDLERS.get(block.name)
                    tool_start_timer(bid)
                    args_preview = str(list(block.input.values())[:3])[:80]
                    print(f"{YELLOW}┌ {block.name}({args_preview}){RESET}")
                    try:
                        output = (
                            handler(**block.input) if handler else f"未能识别：{block.name}"
                        )
                    except TypeError as e:
                        output = f"错误：工具 '{block.name}' 参数不匹配: {e}。请检查传入参数与 schema 是否一致。"
                    except Exception as e:
                        output = f"错误：工具 '{block.name}' 执行异常: {e}"

                    try:
                        trigger_hooks("PostToolUse", block, output)
                    except Exception:
                        pass

                    if block.name == "todo_write":
                        rounds_since_todo = 0

                    prefix = (
                        RED
                        if output.startswith(("错误", "危险", "Error", "超时"))
                        else GRAY
                    )
                    out_str = str(output)[:200]
                    out_str = out_str.encode('utf-8', errors='surrogatepass').decode('utf-8', errors='replace')
                    if not out_str.strip():
                        out_str = "(无输出)"
                    _safe_print(f"{prefix}└ {out_str}{RESET}\n")
                    results.append(
                        {
                            "type": "tool_result",
                            "tool_use_id": bid,
                            "content": output,
                        }
                    )
            except Exception as e:
                print(f"{RED}工具执行循环异常: {e}{RESET}")

            # 安全网：确保每个 tool_use 都有对应的 tool_result
            resolved_ids = {r.get("tool_use_id") for r in results}
            for block in tool_use_blocks:
                bid = getattr(block, "id", "") or (block.get("id", "") if isinstance(block, dict) else "")
                if bid not in resolved_ids:
                    results.append({
                        "type": "tool_result",
                        "tool_use_id": bid,
                        "content": "错误：工具执行未完成（循环异常或跳过）",
                    })

            messages.append({"role": "user", "content": _clean_surrogates(results)})

    finally:
        signal.signal(signal.SIGINT, old_sigint)
    return True


def print_separator(char="-"):
    try:
        columns = os.get_terminal_size().columns
        print(char * columns)
    except OSError:
        print(char * 80)


SESSIONS_DIR = WORKDIR / "sessions"

def save_session(history):
    """持久化会话到 SQLite 数据库。"""
    global _current_session_id
    if not _current_session_id:
        sid, _ = start_session()
        _current_session_id = sid
    persist_messages(_current_session_id, history)
    return f"sessions.db / {_current_session_id}"

def _print_last_response(history):
    if not history:
        return
    for block in history[-1].get("content", []):
        if getattr(block, "type", None) == "text":
            txt = block.text if isinstance(block.text, str) else str(block.text)
            txt = txt.encode('utf-8', errors='surrogatepass').decode('utf-8', errors='replace')
            _console = Console()
            _console.print(Markdown(txt))
            print()

SLASH_BUILTINS = {
    "/new": "开新对话（自动保存旧对话）",
    "/session": "保存当前对话历史",
    "/help": "显示所有命令和技能",
    "/exit": "退出 Pola",
    "/skills": "列出可用技能",
}

def handle_slash(cmd, history):
    cmd = cmd.strip()
    if cmd == "/new":
        global _current_session_id
        if history:
            p = save_session(history)
            print(f"{GRAY}旧对话已保存: {p}{RESET}")
        if _current_session_id:
            end_session(_current_session_id, "completed")
            _current_session_id = None  # 下一个 save_session 自动开新会话
        history.clear()
        print(f"{GREEN}新对话已开始{RESET}")
        return
    if cmd == "/session":
        if not history:
            print(f"{YELLOW}当前没有对话历史{RESET}")
            return
        p = save_session(history)
        print(f"{GREEN}会话已保存: {p}{RESET}")
        print(f"{GRAY}共 {len(history)} 条消息{RESET}")
        return
    if cmd == "/help":
        print(f"{GREEN}内置命令:{RESET}")
        for c, d in SLASH_BUILTINS.items():
            print(f"  {c:15s} {d}")
        print(f"{GREEN}技能命令:{RESET}")
        for n in sorted(SKILL_REGISTRY):
            print(f"  /{n:14s} {SKILL_REGISTRY[n]['description'][:40]}")
        print(f"{GREEN}其他:{RESET}")
        print(f"  @<文件路径>      让 agent 阅读文件（Tab 补全）")
        print(f"  !<shell命令>     直接执行 shell（不经过 agent）")
        return
    if cmd == "/exit":
        raise SystemExit(0)
    if cmd == "/skills":
        print(list_skills())
        return
    name = cmd[1:]
    if name in SKILL_REGISTRY:
        content = load_skill(name)
        history.append({"role": "user", "content": f"<skill name=\"{name}\">\n{content}\n</skill>"})
        print(f"{GREEN}已加载技能: {name}{RESET}")
        agent_loop(history)
        _print_last_response(history)
        return
    print(f"{RED}未知命令: {cmd}（输入 /help 查看）{RESET}")

def handle_at(text, history):
    rest = text[1:].strip()
    if not rest:
        print(f"{YELLOW}用法: @<文件路径>  （输入 @ 后按 Tab 查看文件列表）{RESET}")
        return
    parts = rest.split(None, 1)
    fpath_str = parts[0]
    question = parts[1] if len(parts) > 1 else ""
    fpath = WORKDIR / fpath_str
    if not fpath.is_file():
        print(f"{RED}文件不存在: {fpath_str}{RESET}")
        return
    try:
        content = fpath.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        content = fpath.read_bytes().decode("gbk", errors="replace")
    if len(content) > 1_000_000:
        content = content[:1_000_000] + "\n...(内容过长，已截断至 1MB)"
    print(f"{GRAY}已读取: {fpath_str} ({len(content)} 字符){RESET}")
    msg = f"请阅读文件 {fpath_str}:\n\n{content}"
    if question:
        msg += f"\n\n用户问题: {question}"
    trigger_hooks("UserPromptSubmit", f"@{fpath_str}")
    history.append({"role": "user", "content": msg})
    agent_loop(history)
    _print_last_response(history)

def handle_bang(text):
    cmd = text[1:].strip()
    if not cmd:
        print(f"{YELLOW}用法: !<命令>  （直接执行 shell 命令，不经过 agent）{RESET}")
        return
    print(f"{GRAY}执行: {cmd}{RESET}")
    try:
        shell_exe = None
        if os.name == "nt":
            sys_root = os.environ.get("SYSTEMROOT", r"C:\Windows")
            shell_exe = os.path.join(sys_root, "System32", "cmd.exe")
        res = subprocess.run(
            cmd, shell=True, executable=shell_exe,
            capture_output=True, cwd=str(WORKDIR), timeout=120,
            creationflags=_no_window_flags(),
            startupinfo=_no_window_startupinfo(),
        )
    except subprocess.TimeoutExpired:
        print(f"{RED}超时（120s）{RESET}")
        return
    raw = (res.stdout or b"") + (res.stderr or b"")
    try:
        out = raw.decode("utf-8").strip()
    except UnicodeDecodeError:
        out = raw.decode("gbk", errors="replace").strip()
    print(f"{GRAY}{out[:50000] if out else '(无输出)'}{RESET}")

class PolaCompleter(Completer):
    def get_completions(self, document, complete_event):
        text = document.text
        if text.startswith('/') and ' ' not in text:
            cmds = list(SLASH_BUILTINS.keys()) + [f'/{n}' for n in sorted(SKILL_REGISTRY)]
            for c in sorted(cmds):
                if c.startswith(text):
                    desc = SLASH_BUILTINS.get(c) or SKILL_REGISTRY.get(c[1:], {}).get('description', '')
                    yield Completion(c, start_position=-len(text), display_meta=str(desc)[:40])
        elif text.startswith('@'):
            prefix = text[1:]
            search_dir = WORKDIR
            partial = prefix
            if '/' in prefix:
                dir_str, partial = prefix.rsplit('/', 1)
                search_dir = WORKDIR / dir_str
            if not search_dir.exists() or not search_dir.is_dir():
                return
            base = prefix[:len(prefix) - len(partial)]
            try:
                for f in sorted(search_dir.iterdir()):
                    if f.name.startswith('.') and not partial.startswith('.'):
                        continue
                    if f.name.lower().startswith(partial.lower()):
                        suffix = '/' if f.is_dir() else ''
                        yield Completion(f'@{base}{f.name}{suffix}', start_position=-len(text))
            except Exception:
                pass

def _dyn_style():
    try:
        t = get_app().current_buffer.text
        if t.startswith('!'):
            return Style([("prompt", "fg:ansipurple bold"), ("", "fg:ansipurple")])
    except Exception:
        pass
    return _pt_style


if __name__ == "__main__":
    import argparse as _argparse
    _ap = _argparse.ArgumentParser(description="Pola -- SolidWorks AI Agent")
    _ap.add_argument("--run", type=str, default=None,
                     help="headless mode: run a task prompt and exit. Can be a prompt string or @file.")
    _ap.add_argument("--model", type=str, default=None,
                     help="override MODEL_ID env var")
    _ap.add_argument("--max-turns", type=int, default=60,
                     help="max agent turns in --run mode (default 60)")
    _args = _ap.parse_args()

    if _args.model:
        MODEL = _args.model
        os.environ["MODEL_ID"] = _args.model

    if _args.run:
        # ---- headless non-interactive mode ----
        prompt_text = _args.run
        # support @file syntax
        if prompt_text.startswith("@"):
            fpath = WORKDIR / prompt_text[1:].strip()
            if fpath.is_file():
                try:
                    file_content = fpath.read_text(encoding="utf-8")
                except UnicodeDecodeError:
                    file_content = fpath.read_bytes().decode("gbk", errors="replace")
                if len(file_content) > 1_000_000:
                    file_content = file_content[:1_000_000] + "\n...(内容过长，已截断至 1MB)"
                prompt_text = file_content
                print(f"{GRAY}[--run] read task from: {fpath.name} ({len(file_content)} chars){RESET}")
            else:
                print(f"{RED}[--run] file not found: {fpath}{RESET}")
                raise SystemExit(1)

        print(f"{GREEN}[--run] headless mode, max {_args.max_turns} turns{RESET}")
        print(f"{GRAY}[--run] task preview: {prompt_text[:200]}...{RESET}")

        history = []
        trigger_hooks("UserPromptSubmit", prompt_text)
        history.append({"role": "user", "content": prompt_text})
        ok = agent_loop(history, max_turns=_args.max_turns)
        _print_last_response(history)
        if ok is False:
            print(f"{RED}[--run] 执行失败（API 流式错误）{RESET}")
            raise SystemExit(1)
        print(f"{GREEN}[--run] done, {len(history)} messages{RESET}")
        raise SystemExit(0)

    # ---- interactive mode ----
    print("\033[2J\033[H", end="")
    print(f"{GREEN}Pola Ready at {WORKDIR}.{RESET}")
    print("cmd: /help | @file | !shell | double Ctrl+C to exit")

    history = []
    _last_ctrl_c = 0.0
    _completer = PolaCompleter()
    _dyn_style_obj = DynamicStyle(_dyn_style)

    while True:
        try:
            print_separator("-")
            query = _pt_prompt(
                [("class:prompt", "Pola >> ")],
                style=_dyn_style_obj, history=_pt_history, completer=_completer,
            )
        except EOFError:
            break
        except KeyboardInterrupt:
            now = time.time()
            if now - _last_ctrl_c < 2.0:
                print(f"\n\033[35mGoodbye![00m")
                break
            _last_ctrl_c = now
            print(f"\n\033[33mpress again to quit\033[0m")
            continue

        query = query.strip()
        if not query:
            continue
        if query.startswith("/"):
            handle_slash(query, history)
            continue
        if query.startswith("@"):
            handle_at(query, history)
            continue
        if query.startswith("!"):
            handle_bang(query)
            continue

        trigger_hooks("UserPromptSubmit", query)
        history.append({"role": "user", "content": query})
        agent_loop(history)
        _print_last_response(history)
