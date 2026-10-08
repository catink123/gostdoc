# AGENTS.md

Guidance for AI coding agents (Claude Code, Codex, Cursor, Copilot, …) working in this repository.

## If you are asked to write a document with gostdoc

Read [`gostdoc/AI_INSTRUCTIONS.md`](gostdoc/AI_INSTRUCTIONS.md) first, all of it. It is the
complete Markdown syntax and the GOST rules the tool expects. When gostdoc is installed, the same
text is printed by `gostdoc --guide`. [`examples/example.md`](examples/example.md) uses every feature.

Then convert and fix every warning:

```bash
gostdoc report.md -o report.docx          # exit code 0 and no "warning:" lines = done
```

## If you are asked to change gostdoc itself

```bash
python -m venv .venv
.venv/Scripts/python -m pip install -e ".[dev]"   # Linux/macOS: .venv/bin/python
.venv/Scripts/python -m pytest -q tests
```

- Pipeline: `mdparse.py` (Markdown → block IR) → `numbering.py` (labels and numbers) →
  `render.py` (IR → python-docx) + `frame.py` (ЕСКД frame / title block) + `latex2omml.py`
  (LaTeX → OMML) → optional `wordpost.py` (Word COM: update fields, PDF).
- Settings and profiles live in `config.py`. A new setting needs a default in `_BASE`, a row in
  the settings table of `gostdoc/AI_INSTRUCTIONS.md`, and a test.
- If you change Markdown syntax or output, update `gostdoc/AI_INSTRUCTIONS.md` and
  `examples/example.md` in the same change. The guide is what other agents rely on.
- Tests must pass without Microsoft Word (`use_word="no"`); CI runs on Linux and Windows.
