#!/usr/bin/env python3

import os, subprocess, ast, json, yaml, math
from pathlib import Path
import win32com.client
from rich.console import Console
from rich.markdown import Markdown

RED = "\033[31m"
YELLOW = "\033[33m"
GRAY = "\033[90m"
GREEN = "\033[36m"
RESET = "\033[0m"

from prompt_toolkit import prompt as _pt_prompt
from prompt_toolkit.history import InMemoryHistory
from prompt_toolkit.styles import Style

_pt_history = InMemoryHistory()
_pt_style = Style([("prompt", "fg:ansicyan")])

from dotenv import load_dotenv
from anthropic import Anthropic

load_dotenv(override=True)

if os.getenv("ANTHROPIC_BASE_URL"):
    os.environ.pop("ANTHROPIC_AUTH_TOKEN", None)


WORKDIR = Path.cwd()
SKILLS_DIR = WORKDIR / ".agents/skills"
client = Anthropic()
MODEL = os.environ["MODEL_ID"]
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
        f"You are pola, a coding agent at {WORKDIR}."
        f"""
        Use tools to solve tasks. Act, and explain what you did.
        ## About Me
        Name:Acid|Major:Mechanical Manufacturing|Language:Simplify Chinese
        I expect you to act in accordance with the following code of conduct

        ## 1. Think Before Coding

        **Don't assume. Don't hide confusion. Surface tradeoffs.**

        Before implementing:
        - State your assumptions explicitly. If uncertain, ask.
        - If multiple interpretations exist, present them - don't pick silently.
        - If a simpler approach exists, say so. Push back when warranted.
        - If something is unclear, stop. Name what's confusing. Ask.

        ## 2. Simplicity First

        **Minimum code that solves the problem. Nothing speculative.**

        - No features beyond what was asked.
        - No abstractions for single-use code.
        - No "flexibility" or "configurability" that wasn't requested.
        - No error handling for impossible scenarios.
        - If you write 200 lines and it could be 50, rewrite it.

        Ask yourself: "Would a senior engineer say this is overcomplicated?" If yes, simplify.

        ## 3. Surgical Changes

        **Touch only what you must. Clean up only your own mess.**

        When editing existing code:
        - Don't "improve" adjacent code, comments, or formatting.
        - Don't refactor things that aren't broken.
        - Match existing style, even if you'd do it differently.
        - If you notice unrelated dead code, mention it - don't delete it.

        When your changes create orphans:
        - Remove imports/variables/functions that YOUR changes made unused.
        - Don't remove pre-existing dead code unless asked.

        The test: Every changed line should trace directly to the user's request.

        ## 4. Goal-Driven Execution

        **Define success criteria. Loop until verified.**

        Transform tasks into verifiable goals:
        - "Add validation" → "Write tests for invalid inputs, then make them pass"
        - "Fix the bug" → "Write a test that reproduces it, then make it pass"
        - "Refactor X" → "Ensure tests pass before and after"

        For multi-step tasks, state a brief plan:
        ```
        1. [Step] → verify: [check]
        2. [Step] → verify: [check]
        3. [Step] → verify: [check]
        ```

        Strong success criteria let you loop independently. Weak criteria ("make it work") require constant clarification.

        ---

        **These guidelines are working if:** fewer unnecessary changes in diffs, fewer rewrites due to overcomplication, and clarifying questions come before implementation rather than after mistakes.
        """
        f"""
        ## SolidWorks 建模与绘图规范 (SolidWorks Modeling & Sketching Guidelines)

        当你受命使用 `sw_` 前缀的工具进行 3D 建模或绘图时，必须严格遵守以下工业级 CAD 规范：

        ### 5.1 严格的绝对单位制 (SI Units)
        - **所有工具的长度、半径、深度参数，其内部单位一律是“米 (meters)”！**
        - 如果用户说 `100mm`，你必须在输入参数时转换为 `0.1`。
        - 如果用户说 `5cm`，你必须输入 `0.05`。
        - 绝不允许直接把毫米数值作为参数传入工具。

        ### 5.2 “先选面，后画图”原则 (Select Before Sketching)
        - 任何 2D 绘图工具（如 `sw_draw_circle`, `sw_draw_rectangle`）在调用前，**必须先有明确选中的基准面或实体表面**。
        - 在新建零件后，你的第一步应当是调用 `sw_select_plane_or_face` 选中一个初始基准面（如 `"Front Plane"` 前视、`"Top Plane"` 上视 或 `"Right Plane"` 右视）。
        - 严禁在没有处于任何选中表面的状态下执行绘图工具。

        ### 5.3 实体生成的特征链 (Feature Chain)
        - 绘制完 2D 草图后，SolidWorks 会按照递增顺序自动为其命名（例如 `Sketch1`, `Sketch2`）。
        - 如果你无法确定当前草图的名称，请立刻调用 `sw_get_model_structure` 读取特征树，确保名称准确。
        - 在调用 `sw_extrude_boss`（拉伸）或 `sw_extrude_cut`（切除）时，必须准确传入你想操作的草图名称（如 `"Sketch1"`）。

        ### 5.4 闭环的交付流程 (Execution Workflow)
        - 一个完整的建模意图应当包含：创建画布 -> 选面 -> 绘图 -> 特征变换（拉伸/切除）-> 导出交付。
        - 除非用户明确要求停止，否则在建立完 3D 实体后，应当主动使用 `sw_export_to_format` 工具将其导出为用户指定的工业通用格式（如 `STEP`）。
        """
        f"Skills available:\n{catalog}\n"
        "Use load_skill to get full details when needed."
    )

SYSTEM = build_system()

SUB_SYSTEM = (
    f"You are a useful agent at {WORKDIR}. "
    "Complete the task you were given, then return a concise summary. "
    "Do not delegate further."
)

def run_bash(command, **kwargs):
    danger = ["rmdir", "rm -rf /", "sudo", "shutdown", "reboot", "> /dev/"]
    if any(d in command for d in danger):
        return "危险！请避免尝试执行高风险指令。"
    try:
        r = subprocess.run(
            command,
            shell=False,
            cwd=WORKDIR,
            capture_output=True,
            timeout=120,
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
                text, enc = raw.decode("utf-8"), "utf-8"
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

def run_sw_select_face_by_point(x, y, z, entity_type="FACE", **kwargs):
    """按空间坐标点拾取实体面/边/顶点。解决 SelectByID2 无法按名称选实体面的问题。
    例如板顶面中心可传 (0, 0, 0.02) 选顶面，传 (0.05, 0.05, 0) 选板底面角点附近的边。"""
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model:
            return "错误：当前没有打开任何活动文档。"

        t = str(entity_type).upper()
        if t not in ("FACE", "EDGE", "VERTEX"):
            return f"错误：entity_type 必须是 FACE/EDGE/VERTEX，收到 '{entity_type}'。"

        model.ClearSelection2(True)
        # SelectByID2 传空名称 + 坐标 + 类型，SolidWorks 会拾取该坐标处的实体几何
        success = model.Extension.SelectByID2(
            "", t, float(x), float(y), float(z),
            False, 0, _empty_callout(), 0
        )
        if success:
            return f"成功：已在坐标 ({x}, {y}, {z}) 选中 {t}。"
        return f"失败：坐标 ({x}, {y}, {z}) 处未命中 {t}。请确认坐标在实体表面/边线上。"
    except Exception as e:
        return f"错误：{e}"

def run_sw_get_model_structure(**kwargs):
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model:
            return "错误：当前没有活动文档。"

        structure = []
        feature = _sw_member(model, "FirstFeature")
        while feature:
            name = _sw_member(feature, "Name")
            type_name = _sw_member(feature, "GetTypeName2")
            structure.append(f"- {name} [{type_name}]")
            feature = _sw_member(feature, "GetNextFeature")

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

# --- 3. 3D 特征生成工具 ---

def run_sw_extrude_boss(sketch_name, depth, **kwargs):
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        _select_by_id(model.Extension, sketch_name, "SKETCH")

        feat = model.FeatureManager.FeatureExtrusion3(
            True, False, False, 0, 0, depth, 0.0,
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

        _select_by_id(model.Extension, sketch_name, "SKETCH")

        end_condition = 1 if thru_all else 0
        cut_depth = 0.01 if thru_all else depth

        # 与 skill sw_part.extrude_cut 对齐的 27 参数稳定签名
        # FeatureCut4 在草图与实体表面共面时方向判定会静默失败，按 troubleshooting.md
        # 的稳定写法：先 flip=False/dir=True，失败则翻 flip，再翻 direction，逐组合重试。
        # 重试上限由 4 降为 2：第 1 次失败先检查 license 服务是否仍运行，避免 license 挂起时
        # 无意义重试放大 COM 压力。
        attempts = [
            (True,  False, False),  # Sd, Flip, Dir
            (True,  True,  False),  # 翻转切除侧
        ]
        feat = None
        for idx, (sd, flip, direction) in enumerate(attempts):
            _select_by_id(model.Extension, sketch_name, "SKETCH")
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
        return f"失败：切除特征创建失败（已尝试 2 种 flip 组合）。请确认草图 '{sketch_name}' 闭合且与实体有交集，或尝试翻转草图平面。"
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
        model.EditRebuild3()
        return f"成功：已将尺寸 '{dimension_name}' 修改为 {new_value}，且已重建模型。"
    except Exception as e:
        return f"错误：{e}"

# --- 5. 出图与交付工具 ---

def run_sw_export_to_format(file_type, output_path, **kwargs):
    try:
        sw = _get_sw_app()
        model = sw.ActiveDoc
        if not model: return "错误：没有活动文档。"

        abs_path = str((WORKDIR / output_path).resolve())

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
    return "\n".join(getattr(b, "text", "") for b in content if getattr(b, "type", None) == "text")

def spawn_subagent(description: str) -> str:
    print(f"{GREEN}[Subagent Spawned]{RESET}")
    messages = [{"role":"user", "content": description}]

    for _ in range(30):
        with client.messages.stream(
            max_tokens=256000, model=SUB_MODEL, system=SUB_SYSTEM,
            messages=messages, tools=SUB_TOOLS,
        ) as stream:
            for text in stream.text_stream:
                print(text, end="", flush=True)
            resp = stream.get_final_message()
        print()

        messages.append({"role":"assistant", "content": resp.content})

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
                print(f"{GRAY}[sub] {block.name}: {str(output)[:100]}{RESET}")
                results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": output
                })
        messages.append({
            "role":"user",
            "content":results
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
    "sw_draw_rectangle": run_sw_draw_rectangle,
    "sw_draw_circle": run_sw_draw_circle,
    "sw_draw_polygon": run_sw_draw_polygon,
    "sw_extrude_boss": run_sw_extrude_boss,
    "sw_extrude_cut": run_sw_extrude_cut,
    "sw_apply_fillet": run_sw_apply_fillet,
    "sw_modify_dimension": run_sw_modify_dimension,
    "sw_export_to_format": run_sw_export_to_format,
    "sw_close_doc": run_sw_close_doc,
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
  "rm -rf /", "sudo", "shutdown", "reboot","mkfs", "dd if=", "> /dev/sda",
]
DESTRUCTIVE = [
    "rm ", "> /etc/", "chmod 777"
]

def permission_hook(block):
    if block.name == "bash":
        for pattern in DENY_LIST:
            if pattern in block.input.get("command", ""):
                print(f"{RED}Blocked: '{pattern}'{RESET}")
                return "Permission denied by deny list"
        for kw in DESTRUCTIVE:
            if kw in block.input.get("command", ""):
                print(f"{RED}Potentially destructive command{RESET}")
                print(f"   Tool: {block.name}({block.input})")
                choice = input("   Allow? [y/N] ").strip().lower()
                if choice not in ("y", "yes"):
                    return "Permission denied by user"
    if block.name in ("write_file", "edit_file"):
        path = block.input.get("path", "")
        if not (WORKDIR / path).resolve().is_relative_to(WORKDIR):
            print(f"{RED}写位置超出工作区!{RESET}")
            print(f"   Tool: {block.name}({block.input})")
            choice = input("   Allow? [y/N] ").strip().lower()
            if choice not in ("y", "yes"):
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

def context_inject_hook(query):
    print(f"{GRAY}[HOOK] UserPromptSubmit: working in {WORKDIR}{RESET}")
    return None

def summary_hook(messages):
    tool_count = sum(1 for m in messages
                     for b in (m.get("content") if isinstance(m.get("content"),list) else [])
                     if isinstance(b, dict) and b.get("type") == "tool_result")
    print(f"{GRAY}[HOOK] Stop: session used {tool_count} tool calls{RESET}")
    return None

register_hook("UserPromptSubmit", context_inject_hook)
register_hook("PreToolUse", permission_hook)
# register_hook("PreToolUse", log_hook)
register_hook("PostToolUse", large_output_hook)
register_hook("Stop", summary_hook)

rounds_since_todo = 0

def agent_loop(messages):
    global rounds_since_todo
    while True:

        if rounds_since_todo >=3 and messages:
            messages.append({
                "role": "user",
                "content": "<reminder>Update your todos.</reminder>"
            })
            rounds_since_todo = 0

        with client.messages.stream(
            max_tokens=256000, model=MODEL, messages=messages, tools=TOOLS, system=SYSTEM
        ) as stream:
            for text in stream.text_stream:
                print(text, end="", flush=True)
            resp = stream.get_final_message()
        print()

        messages.append({"role": "assistant", "content": resp.content})

        if resp.stop_reason != "tool_use":
            force = trigger_hooks("Stop", messages)
            if force:
                messages.append({"role": "user", "content": force})
                continue
            #print_separator("-")
            return

        rounds_since_todo += 1

        results = []
        for block in resp.content:
            if block.type != "tool_use":
                continue

            blocked = trigger_hooks("PreToolUse", block)
            if blocked:
                results.append({
                    "type": "tool_result",
                    "tool_use_id":block.id,
                    "content": str(blocked)
                })
                continue
                
            handler = TOOL_HANDLERS.get(block.name)
            try:
                output = (
                    handler(**block.input) if handler else f"未能识别：{block.name}"
                )
            except TypeError as e:
                output = f"错误：工具 '{block.name}' 参数不匹配: {e}。请检查传入参数与 schema 是否一致。"
            except Exception as e:
                output = f"错误：工具 '{block.name}' 执行异常: {e}"

            trigger_hooks("PostToolUse", block, output)

            if block.name == "todo_write":
                rounds_since_todo = 0

            # output = run_bash(block.input['command'])
            prefix = (
                RED
                if output.startswith(("错误", "危险", "Error", "超时"))
                else GRAY
            )
            print(f"{prefix}└ {str(output)[:200]}{RESET}\n")
            results.append(
                {
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": output,
                }
            )

        messages.append({"role": "user", "content": results})


def print_separator(char="-"):
    try:
        columns = os.get_terminal_size().columns
        print(char * columns)
    except OSError:
        print(char * 80)


if __name__ == "__main__":
    #os.system("cls" if os.name == "nt" else "clear")
    print("\033[2J\033[H", end="")
    print(f"{GREEN}Pola Ready at {WORKDIR}.{RESET}")
    print("回车发送消息，输入 q 退出.")

    history = []
    while True:
        try:
            print_separator("-")
            query = _pt_prompt([("class:prompt", "Pola >> ")], style=_pt_style, history=_pt_history)
        except (EOFError, KeyboardInterrupt):
            break
        if query.strip().lower() in ("q", "exit", "fuck"):
            print(f"\033[35mGoodbye![00m")
            break
        trigger_hooks("UserPromptSubmit", query)
        # if query.startswith("!"):
        #     # cmd = f"powershell -Command {query[1:].strip()}"
        #     cmd = query[1:].strip()
        #     res = subprocess.run(
        #         cmd, shell=True, capture_output=True, cwd=os.getcwd(), timeout=120
        #     )
        #     raw = (res.stdout or b"") + (res.stderr or b"")
        #     try:
        #         out = raw.decode("utf-8").strip()
        #     except UnicodeDecodeError:
        #         out = raw.decode("gbk", errors="replace").strip()
        #     print(f"{GRAY}{out[:50000] if out else '(无输出)'}{RESET}")
        #     history.append(
        #         {"role": "user", "content": out[:50000] if out else "(无输出)"}
        #     )
        # else:
        history.append({"role": "user", "content": query})
        agent_loop(history)

        for block in history[-1]["content"]:
            if getattr(block, "type", None) == "text":
                _console = Console()
                _console.print(Markdown(block.text))
                #print(block.text)
                print()
