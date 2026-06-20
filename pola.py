#!/usr/bin/env python3

import os, subprocess, ast, json, yaml
from pathlib import Path
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
            shell=True,
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
            output = (
                handler(**block.input) if handler else f"未能识别：{block.name}"
            )

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
