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
    readline.parse_and_bind('set bind-tty-special-chars off')
    readline.parse_and_bind('set input-meta on')
    readline.parse_and_bind('set output-meta on')
    readline.parse_and_bind('set convert-meta off')
except ImportError:
    pass


from dotenv import load_dotenv
from anthropic import Anthropic

load_dotenv(override = True)

if os.getenv("ANTHROPIC_BASE_URL"):
  os.environ.pop("ANTHROPIC_AUTH_TOKEN", None)


WORKDIR = Path.cwd()
client = Anthropic();
MODEL = os.environ["MODEL_ID"]

SYSTEM = f"You are my personal assistant at {WORKDIR}. Use bash to solve tasks. Act, and explain what you did."


TOOLS = [
    {"name": "bash", "description": "Run a shell command.",
     "input_schema": {"type": "object", "properties": {"command": {"type": "string"}}, "required": ["command"]}},
    {"name": "read_file", "description": "Read file contents.",
     "input_schema": {"type": "object", "properties": {"path": {"type": "string"}, "limit": {"type": "integer"}}, "required": ["path"]}},
    {"name": "write_file", "description": "Write content to a file.",
     "input_schema": {"type": "object", "properties": {"path": {"type": "string"}, "content": {"type": "string"}}, "required": ["path", "content"]}},
    {"name": "edit_file", "description": "Replace exact text in a file once.",
     "input_schema": {"type": "object", "properties": {"path": {"type": "string"}, "old_text": {"type": "string"}, "new_text": {"type": "string"}}, "required": ["path", "old_text", "new_text"]}},
    {"name": "glob", "description": "Find files matching a glob pattern.",
     "input_schema": {"type": "object", "properties": {"pattern": {"type": "string"}}, "required": ["pattern"]}},
]

def run_bash(command):
  danger = ["rmdir","rm -rf /", "sudo", "shutdown", "reboot", "> /dev/"]
  if any(d in command for d in danger):
    return "危险！请避免尝试执行高风险指令。"
  try:
    r = subprocess.run(
      command,
      shell = True,
      cwd=os.getcwd(),
      capture_output = True,
      timeout = 120,
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
    raise ValueError(f"Path escapes workspace: {p}")
  return path


def run_read(path, limit: int | None = None):
  try:
    lines = safe_path(path).read_text().splitlines()
    if limit and limit < len(lines):
        lines = lines[:limit] + [f"... ({len(lines) - limit} more lines)"]
    return "\n".join(lines)
  except Exception as e:
    return f"错误：{e}"

def run_write(path, content):
  try:
    file_path = safe_path(path)
    file_path.parent.mkdir(parents=True, exist_ok=True)
    file_path.write_text(content)
    return f"Wrote {len(content)} bytes to {path}"
  except Exception as e:
    return f"错误：{e}"

def run_edit(path, old_text, new_text):
  try:
    file_path = safe_path(path)
    text = file_path.read_text()
    if old_text not in text:
        return f"{RED}错误：没有命中修改区域，可能要替换的文字不存在，或文件自上次查看已更新{RESET}"
    file_path.write_text(text.replace(old_text, new_text, 1))
    return f"{GREEN}{path} 已更新{RESET}"
  except Exception as e:
    return f"错误：{e}"

def run_glob(pattern):
  import glob as g
  try:
    results =[]
    for match in g.glob(pattern, root_dir = WORKDIR):
      if (WORKDIR / match).resolve().is_relative_to(WORKDIR):
        results.append(match)
    return "\n".join(g.glob(pattern, root_dir=WORKDIR))
  except Exception as e:
    return f"错误：{e}"

TOOL_HANDLERS = {
    "bash": run_bash,
    "read_file": run_read,
    "write_file": run_write,
    "edit_file": run_edit,
    "glob": run_glob,
}

def agent_loop(messages):
  while True:
    resp = client.messages.create(
      max_tokens=8000,
      model=MODEL,
      messages=messages,
      tools = TOOLS,
      system = SYSTEM
    )

    messages.append({"role": "assistant", "content": resp.content})

    if resp.stop_reason == "end_turn":
      return;

    results = []
    for block in resp.content:
      if block.type == "tool_use":
        print(f"{YELLOW}Tool_Call({block.name}){RESET}")
        handler = TOOL_HANDLERS.get(block.name)
        output = handler(**block.input) if handler else f"未能识别：{block.name}"
        #output = run_bash(block.input['command'])
        prefix = RED if output.startswith(("错误","危险", "Error", "超时")) else GRAY
        print(f"{prefix}└ {str(output)[:200]}{RESET}\n")
        results.append({
          "type": "tool_result",
          "tool_use_id": block.id,
          "content": output,
        })

    messages.append({"role":"user","content":results})

if __name__ == "__main__":
  print(f"{GREEN}Pola Ready On Call.{RESET}")
  print("回车发送消息，输入 q 退出. \n")

  history = []
  while True:
    try:
      query = input(f"{GREEN}Pola >> {RESET}")
    except (EOFError, KeyboardInterrupt):
      break
    if query.strip().lower() in ("q", "exit", "fuck"):
      break
    if query.startswith("!"):
      # cmd = f"powershell -Command {query[1:].strip()}"
      cmd = query[1:].strip()
      res = subprocess.run(
        cmd,
        shell=True,
        capture_output=True,
        cwd = os.getcwd(),
        timeout=120)
      raw = (res.stdout or b"") + (res.stderr or b"")
      try:
        out = raw.decode("utf-8").strip()
      except UnicodeDecodeError:
        out = raw.decode("gbk", errors="replace").strip()
      print(f"{GRAY}{out[:50000] if out else '(无输出)'}{RESET}")
      history.append({"role": "user", "content": out[:50000] if out else "(无输出)"})
    else:
      history.append({"role": "user", "content": query})
      agent_loop(history)

    resp_content = history[-1]["content"]
    if isinstance(resp_content, list):
      for block in resp_content:
        if getattr(block, "type",
          None) == "text":
          print(block.text)
    print()


