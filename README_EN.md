<p align="center">
  <a href="README.md">🇨🇳 中文</a> &nbsp;|&nbsp;
  <a href="README_EN.md">🇺🇸 English</a> &nbsp;|&nbsp;
  <a href="README_JA.md">🇯🇵 日本語</a>
</p>

---

# Paradigm Report Auto-Generator

> Excel template extraction & auto-fill tool based on diff analysis

## What problem does it solve?

Enterprise report Excel files are 90% identical across versions — only a few fields change (author, version number, dates, etc.). Manually finding and updating these fields is inefficient and error-prone.

**This tool**: Feed in historical Excel files → auto-identify variable fields → fill in current values → auto-generate new report.

## Quick Start

```bash
pip install -r requirements.txt
```

```bash
# 1. Analyze historical reports
python3 main.py analyze --input ./data/history_reports --output ./project_config

# 2. Fill in variables (or interactive mode)
python3 main.py ask

# 3. Generate new report
python3 main.py generate -t ./project_config/template.xlsx -v ./data/current_values/current_values.json -o ./output/report.xlsx
```

## Commands

| Command | Description |
|---------|-------------|
| `analyze` | Analyze historical Excel, auto-identify variable fields |
| `generate` | Template + variables → new report (with global text replacement) |
| `validate` | Validate report integrity |
| `diff` | Compare two reports |
| `ask` | Interactive variable input |

## Project Structure

```
├── main.py              # CLI entry point
├── src/
│   ├── excel_reader.py      # Excel reader
│   ├── diff_analyzer.py     # Diff analysis
│   ├── variable_detector.py # Variable detection
│   ├── template_builder.py  # Template generation
│   ├── report_generator.py  # Report generation (XML-level)
│   ├── validator.py         # Validation
│   └── change_logger.py     # Change logging
├── data/
├── project_config/
├── tests/
└── requirements.txt
```

## Tech Stack

Python 3.10+ · Direct XML manipulation · Preserves original formatting · 17 tests passing
