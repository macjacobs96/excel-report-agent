<p align="center">
  <a href="README.md">🇨🇳 中文</a> &nbsp;|&nbsp;
  <a href="README_EN.md">🇺🇸 English</a> &nbsp;|&nbsp;
  <a href="README_JA.md">🇯🇵 日本語</a>
</p>

---

# 范式报告自动生成系统

> 基于差异分析的 Excel 模板抽取与自动填充工具

## 解决什么问题？

企业报告 Excel 每个版本之间 90% 内容相同，只有少量字段变化（编制人、版本号、日期等）。人工查找修改低效且容易遗漏。

**本工具**：喂入历史 Excel → 自动识别变动字段 → 填写本次变量 → 自动输出新报告。

## 快速开始

```bash
pip install -r requirements.txt
```

```bash
# 1. 分析历史报告
python3 main.py analyze --input ./data/history_reports --output ./project_config

# 2. 填写变量（或交互式）
python3 main.py ask

# 3. 生成新报告
python3 main.py generate -t ./project_config/template.xlsx -v ./data/current_values/current_values.json -o ./output/report.xlsx
```

## 命令一览

| 命令 | 功能 |
|------|------|
| `analyze` | 分析历史 Excel，自动识别可变字段 |
| `generate` | 模板 + 变量 → 新报告（含全局文本替换） |
| `validate` | 校验报告完整性 |
| `diff` | 新旧报告对比 |
| `ask` | 交互式问答填写变量 |

## 项目结构

```
├── main.py              # CLI 入口
├── src/
│   ├── excel_reader.py      # Excel 读取
│   ├── diff_analyzer.py     # 差异分析
│   ├── variable_detector.py # 变量识别
│   ├── template_builder.py  # 模板生成
│   ├── report_generator.py  # 报告生成 (XML级操作)
│   ├── validator.py         # 校验
│   └── change_logger.py     # 变更日志
├── data/
├── project_config/
├── tests/
└── requirements.txt
```

## 技术栈

Python 3.10+ · XML 直接操作 · 保留原始格式 · 17 测试通过
