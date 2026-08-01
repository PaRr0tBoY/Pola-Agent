# Pola — SolidWorks AI Agent

基于 Anthropic API 的 SolidWorks 自动化建模智能体。Pola 通过 COM 接口驱动 SolidWorks，支持 AI 自主建模、特征操作、装配、STEP 导出，以及本地标准件库检索。

## 环境要求

- Windows 10/11（SolidWorks COM 自动化仅支持 Windows）
- SolidWorks 2026 已安装并可启动
- Python 3.10+

## 安装

```bash
pip install -r requirements.txt
cp .env.example .env    # 填写 API Key 与模型 ID
```

`.env` 必填项：

| 变量 | 说明 |
|---|---|
| `ANTHROPIC_API_KEY` | API Key（Anthropic 或兼容服务商） |
| `MODEL_ID` | 主模型，如 `deepseek-v4-pro` |
| `SUB_MODEL_ID` | 子任务模型，如 `deepseek-v4-flash` |
| `ANTHROPIC_BASE_URL` | 可选，Anthropic 兼容服务商地址 |

## 使用

交互模式：

```bash
python pola.py
```

无头模式（跑任务文件或提示词后退出）：

```bash
python pola.py --run "@task_demo.md" --max-turns 60
python pola.py --run "创建一个圆柱体并导出 STEP" --max-turns 15
```

交互模式内置命令：`/help`、`/new`（开新对话）、`/session`（保存对话）、`/exit`；`@<文件>` 让 agent 阅读文件；`!<命令>` 直接执行 shell。

## 功能

- 30+ 个 `sw_*` 工具：建模、草图、特征、装配、测量、导出（详见 `pola.py` 中 `TOOLS`）
- 标准件库：本地 SQLite 索引 + GB 标准件种子数据（`.agents/parts_library`），工具 `sw_search_standard_part`
- 会话持久化：每轮对话自动保存到 `sessions.db`（含工具调用统计与安全事件日志）
- 技能系统：`.agents/skills/` 下按需加载技能（当前内置 agentic-engineer 机械工程文档工具箱）
- GBK 安全输出：Windows 中文终端不会因特殊字符崩溃

## 安全说明

Pola 的 `bash` 工具以当前用户权限执行命令。内置拦截清单（递归删除、格式化、关机等）与交互式确认门，但请勿在不受信任的任务文件上运行 `--run` 模式 —— 提示词注入可能诱导模型执行命令。

## 开发

开发循环（详见 `CLAUDE.md`）：

1. 编写任务描述（尺寸单位：米）
2. `python pola.py --run @task.md --max-turns 60`
3. 观察失败并修复根因

## 许可证

MIT
