#!/usr/bin/env python3

import os, subprocess
from pathlib import Path

RED = "\033[31m"
YELLOW = "\033[33m"
GRAY = "\033[90m"
GREEN = "\033[36m"
RESET = "\033[0m"

try:
    import readline

    readline.parse_and_bind("set bind-tty-special-chars off")
    readline.parse_and_bind("set input-meta on")
    readline.parse_and_bind("set output-meta on")
    readline.parse_and_bind("set convert-meta off")
except ImportError:
    pass


from dotenv import load_dotenv
from anthropic import Anthropic

load_dotenv(override=True)

if os.getenv("ANTHROPIC_BASE_URL"):
    os.environ.pop("ANTHROPIC_AUTH_TOKEN", None)


WORKDIR = Path.cwd()
client = Anthropic()
MODEL = os.environ["MODEL_ID"]

SYSTEM = f"You are Pola, my personal assistant at {WORKDIR}. Use tools to solve tasks. Act, don't explain. You're in a terminal environment with no render engine for markdown,latex,etc. So avoid using markdown syntax."


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
]


def run_bash(command, **kwargs):
    danger = ["rmdir", "rm -rf /", "sudo", "shutdown", "reboot", "> /dev/"]
    if any(d in command for d in danger):
        return "危险！请避免尝试执行高风险指令。"
    try:
        r = subprocess.run(
            command,
            shell=True,
            cwd=os.getcwd(),
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
        lines = safe_path(path).read_text().splitlines()
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
        file_path.write_text(content)
        return f"Wrote {len(content)} bytes to {path}"
    except Exception as e:
        return f"错误：{e}"


def run_edit(path, old_text, new_text, **kwargs):
    try:
        file_path = safe_path(path)
        text = file_path.read_text()
        if old_text not in text:
            return f"{RED}错误：没有命中修改区域，可能要替换的文字不存在，或文件自上次查看已更新{RESET}"
        file_path.write_text(text.replace(old_text, new_text, 1))
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

#def run_grep(pattern, **kwargs):

#def run_pr(pattern, **kwargs):

#def resolve_conflict(pattern, **kwargs):



TOOL_HANDLERS = {
    "bash": run_bash,
    "read_file": run_read,
    "write_file": run_write,
    "edit_file": run_edit,
    "glob": run_glob,
}

DENY_LIST = [
  "rm -rf /", "sudo", "shutdown", "reboot","mkfs", "dd if=", "> /dev/sda",
]

def check_deny_list(command):
    for pattern in DENY_LIST:
        if pattern in command:
            return f"已阻止：{pattern} 在当前环境下被禁用"
    return None;

PERMISSION_RULES = [
    {
        "tools": ["write_file", "edit_file"],
        "check": lambda args: not (WORKDIR / args.get("path", "")).resolve().is_relative_to(WORKDIR),
        "messages": "写位置超出工作区!",
    },
    {
        "tools": ["bash"],
        "check": lambda args: any(kw in args.get("command", "") for kw in ["rm ", "> /etc/", "chmod 777"]),
        "message": "危险！请避免尝试执行高风险指令"},
    ]

def check_rules(tool_name, args):
    for rule in PERMISSION_RULES:
        if tool_name in rule["tools"] and rule["check"](args):
            return rule["message"]
    return None


def ask_user(tool_name, args, reason):
    print(f"\n{GREEN} {reason} {RESET}")
    print(f"   Tool: {tool_name}({args})")
    choice = input("   Allow? [y/N] ").strip().lower()
    return "allow" if choice in ("y", "yes") else "deny"


def check_permission(block) -> bool:
    if block.name == "bash":
        reason = check_deny_list(block.input.get("command", ""))
        if reason:
            print(f"\n\033[31m {reason}\033[0m")
            return False
    reason = check_rules(block.name, block.input)
    if reason:
        decision = ask_user(block.name, block.input, reason)
        if decision == "deny":
            return False
    return True

def agent_loop(messages):
    while True:
        resp = client.messages.create(
            max_tokens=8000, model=MODEL, messages=messages, tools=TOOLS, system=SYSTEM
        )

        messages.append({"role": "assistant", "content": resp.content})

        if resp.stop_reason != "tool_use":
            #print_separator("-")
            return

        results = []
        for block in resp.content:
            if block.type == "tool_use":
                print(f"{YELLOW}Tool({block.name}){RESET}")

                if not check_permission(block):
                    results.append({"type": "tool_result", "tool_use_id": block.id, "content": "Permission denied."})
                    continue
                
                handler = TOOL_HANDLERS.get(block.name)
                output = (
                    handler(**block.input) if handler else f"未能识别：{block.name}"
                )
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
    os.system("cls" if os.name == "nt" else "clear")
    print(f"{GREEN}Pola Ready at {WORKDIR}.{RESET}")
    print("回车发送消息，输入 q 退出.")

    history = []
    while True:
        try:
            print_separator("-")
            query = input(f"{GREEN}Pola >>{RESET}")
            print_separator("-")
        except (EOFError, KeyboardInterrupt):
            break
        if query.strip().lower() in ("q", "exit", "fuck"):
            break
        if query.startswith("!"):
            # cmd = f"powershell -Command {query[1:].strip()}"
            cmd = query[1:].strip()
            res = subprocess.run(
                cmd, shell=True, capture_output=True, cwd=os.getcwd(), timeout=120
            )
            raw = (res.stdout or b"") + (res.stderr or b"")
            try:
                out = raw.decode("utf-8").strip()
            except UnicodeDecodeError:
                out = raw.decode("gbk", errors="replace").strip()
            print(f"{GRAY}{out[:50000] if out else '(无输出)'}{RESET}")
            history.append(
                {"role": "user", "content": out[:50000] if out else "(无输出)"}
            )
        else:
            history.append({"role": "user", "content": query})
            agent_loop(history)

        resp_content = history[-1]["content"]
        if isinstance(resp_content, list):
            for block in resp_content:
                if getattr(block, "type", None) == "text":
                    print(block.text)
        print()
