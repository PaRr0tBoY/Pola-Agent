import anthropic
import os
import subprocess

RED = "\033[31m"
YELLOW = "\033[33m"
GRAY = "\033[90m"
GREEN = "\033[36m"
RESET = "\033[0m"

from dotenv import load_dotenv

load_dotenv(override = True)

if os.getenv("ANTHROPIC_BASE_URL"):
  os.environ.pop("ANTHROPIC_AUTH_TOKEN", None)

client = anthropic.Anthropic();
MODEL = os.environ["MODEL_ID"]

SYSTEM = f"You are my personal assistant at {os.getcwd()}. Use bash to solve tasks. Act, and explain what you did."

TOOLS = [{
  "name": "bash",
  "description": "Run a shell command.",
  "input_schema": {
    "type": "object",
    "properties": {"command":{"type":"string"}},
    "required": ["command"],
  },
}]

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
    return e

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
        print(f"{YELLOW}Bash({block.input['command']}){RESET}")
        output = run_bash(block.input['command'])
        prefix = RED if output.startswith(("危险", "Error", "超时")) else GRAY
        print(f"{prefix}└ {output[:200]}{RESET}")
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
    history.append({"role": "user", "content": query})
    agent_loop(history)

    resp_content = history[-1]["content"]
    if isinstance(resp_content, list):
      for block in resp_content:
        if getattr(block, "type",
          None) == "text":
          print(block.text)
    print()


