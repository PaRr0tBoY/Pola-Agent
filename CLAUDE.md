# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Pola is a SolidWorks AI agent. It wraps an LLM agent loop with COM-based SolidWorks automation tools, enabling AI-driven 3D modeling, verification, and STEP export.

**Goal:** Build a reliable AI agent that can autonomously model mechanical parts in SolidWorks and export production-ready STEP files.

## Development Paradigm

The core development loop:

1. **Write a task** (plain text or Markdown) describing what to model, with exact dimensions in meters
2. **Run Pola headless:** `python pola.py --run @task_file.md --max-turns 60`
3. **Observe output** — look for failures (circular pattern, assembly insertion, sketch name drift, revolve degeneracy)
4. **Fix root causes** in `pola.py` (tool implementations, COM parameter fixes) or the system prompt (`build_system()`)

Principle: **subtraction over addition.** Fix code by simplifying, not by adding fallback layers. If a COM API is fundamentally broken in SW 2026 + pywin32 (e.g., `FeatureCircularPattern4`, `AddComponent5`), replace it with a simpler internal implementation rather than adding retry logic.

## Common Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run a task headless
python pola.py --run "@task_demo.md" --max-turns 60

# Run a quick inline prompt
python pola.py --run "create a simple cylinder and export as STEP" --max-turns 15

# Interactive mode
python pola.py

# Verify syntax after edits
python -c "import ast; ast.parse(open('pola.py',encoding='utf-8').read()); print('OK')"
```

## Architecture

Everything lives in `pola.py` (4106 lines). It is a single-file agent with these layers:

### 1. Agent Loop (`agent_loop()`, line ~3598)
Standard LLM agent: `while True` → stream API response → execute tool calls → append results → repeat. Handles interruption (ESC/Ctrl+C), surrogate character cleanup, and unpaired tool_use removal.

### 2. System Prompt (`build_system()`, line ~129)
Assembles the agent's behavioral instructions at runtime. Contains:
- Compact rule lines (units, sketch chains, pitfalls) inline in the prompt
- SolidWorks modeling rules and known pitfalls are inline rule lines, not numbered sections
- Dynamically loaded skill catalog from `.agents/skills/`

### 3. Tool Definitions + Handlers
Each SolidWorks tool has three parts (search for the tool name to find all three):
- **Handler function** (`run_sw_*`): the Python implementation using `win32com.client`
- **Schema** in `TOOLS` list (line 2760): JSON Schema for the LLM
- **Binding** in `TOOL_HANDLERS` dict (line 3311): maps name → function

### 4. SW COM Utilities
- `_get_sw_app()` — cached COM connection to SldWorks.Application
- `_sw_member(obj, attr)` — safe property/method access (handles pywin32 property-vs-method ambiguity)
- `_select_by_id()` — wrapped SelectByID2 with empty callout variant
- `_find_part_template()` / `_find_assembly_template()` — template discovery

### 5. Subsystems
- `.agents/skills/agentic-engineer/` — mechanical engineering doc toolkit (16 templates + 4 scripts); the old `solidworks-automation` skill was removed
- `.agents/parts_library/` — SQLite standard parts database with GB seed data (tracked; runtime `sessions.db` is gitignored)

## Key Code Locations

| Area | Approximate Line | What |
|------|-----------------|------|
| System prompt assembly | 129–156 | `build_system()` — edit for behavioral rules |
| SW modeling rules | inline in prompt | No separate 5.x sections; compact rule lines |
| Revolve profile handler | 780–906 | Degenerate closure detection |
| Extrude cut handler | 908–990 | Sketch name diagnostics |
| Circular pattern handler | 1187–1333 | Internal manual fallback |
| Assembly insertion | 1562–1679 | AddComponent4/5 with path resolution |
| Feature tree traversal | 523–600 | `_traverse_features()` |
| Tool schemas | 2760–3310 | JSON Schema for all sw_* tools |
| Tool handler bindings | 3311 | `TOOL_HANDLERS` dict |
## Known SW 2026 + pywin32 COM Limitations

These APIs **exist** on the COM object (confirmed via `getattr`) but **always return None**:

- `FeatureCircularPattern4/5` — **Workaround:** internal manual fallback in `run_sw_circular_pattern` (calculates hole positions, draws circles, extrudes cuts)
- `AddComponent5/4` — **No reliable fix.** ConfigOption=0 is correct per API docs (0=CurrentSelectedConfig for parts, 1/2 for assemblies). Assembly insertion is marked optional in tasks.

These work reliably:
- `FeatureExtrusion3`, `FeatureRevolve2`, `FeatureCut4`
- `InsertRefPlane`, `FeatureFillet`
- `SelectByID2`, `CreateCircleByRadius`, `CreateLine`, `CreateCenterRectangle`
- `SaveAs` (STEP export), `GetMassProperties`

## CHM API Documentation

Three extracted CHM directories provide SW API reference:

- **`chm_api/`** (17538 files) — .NET API reference (`sldworksapi.chm`). Covers `IFeatureManager::FeatureCircularPattern4` signature, `IAssemblyDoc` members, etc.
- **`chm_vb6/`** (16027 files) — VBA examples (`sldworksapivb6.chm`). Shows real VBA signatures (e.g., `AddComponent5` has 8 params: CompName, ConfigOption, NewConfigName, UseConfigForPartReferences, ExistingConfigName, X, Y, Z).
- **`chm_const/`** (1295 files) — Enum constants (`swconst.chm`). Contains `swAddComponentConfigOptions_e` (0=CurrentSelectedConfig, 1=NewConfigWithAllRefModels, 2=NewConfigWithAsmStructure).

To search: `grep -r "APIName" chm_api/ chm_vb6/`

## Task Design Rules

When writing modeling tasks for Pola:

1. **All dimensions in meters.** 100mm = 0.1, φ20mm → r=0.01.
2. **Avoid unreliable tools:** No `sw_circular_pattern` dependency (use individual draw_circle + extrude_cut instead). No assembly insertion dependency.
3. **Revolve profiles:** Always use ≤N points where the closing segment doesn't backtrack. If p[-2]→p[-1] and p[-1]→p[0] are collinear and opposite direction, the code will reject it. Remove the last point.
4. **Sketch names:** Don't assume Sketch1 exists — new parts may have an empty Sketch1; the first user sketch may be Sketch2 (草图2). Use `sw_get_model_structure` to verify.
5. **Face selection for cuts:** Drawing on selected faces may not persist the sketch. If cuts fail, try drawing on a standard plane (Right Plane) instead.
