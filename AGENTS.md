# AGENTS.md — Excel Report Agent

> 给未来接手这个项目的 Agent（包括未来的你）的指南。

## 项目本质

**测试报告 Excel 自动生成系统**：读取多个历史 Excel 测试报告 → 自动识别差异字段 → 用户填写本次变量 → 输出新报告。

核心原则：**不要用 LLM 生成 Excel，用代码读写 Excel。LLM 只负责识别变量和生成配置。**

## 架构速览

```
用户提供历史 Excel
       ↓
  ExcelReader  →  解析所有 Sheet / 单元格 / 合并单元格 / 格式
       ↓
  DiffAnalyzer  →  对比多个版本的差异
       ↓
  VariableDetector  →  识别变量字段、类型、置信度
       ↓
  TemplateBuilder  →  生成 template.xlsx + 配置文件
       ↓
  [用户填写 current_values.json]
       ↓
  ReportGenerator  →  填充变量到模板，输出新 Excel
       ↓
  Validator + ChangeLogger  →  校验 + 变更日志
```

## Harness 五子系统

| 子系统 | 职责 | 对应文件 |
|--------|------|----------|
| **Harness** | CLI 入口，编排流程 | `main.py` |
| **State** | 变量 Schema、cell mapping、配置 | `project_config/*.json` |
| **Tools** | Excel 读写、差异分析、填充 | `src/*.py` |
| **Memory** | 历史分析结果、学习记录 | `project_config/analysis_result.json` |
| **Evaluation** | 校验报告、变更日志 | `src/validator.py` + `src/change_logger.py` |

## 关键约束

1. **模板和变量分离**：Excel 模板负责任何格式，JSON 负责变量映射，用户输入负责本次内容。
2. **人工修正 > 模型推断**：`cell_mapping.json` 和 `variable_schema.json` 可手工编辑，模型推断只做辅助。
3. **必须输出变更日志**：每次生成必须告诉用户改了什么。
4. **格式保护优先**：尽量不破坏合并单元格、公式、字体、边框、行高列宽。
5. **第一版本地 CLI**：不上 Web，不接数据库。

## 开发阶段

| Phase | 内容 | 状态 |
|-------|------|------|
| P0 | Excel 读取 + 手工 mapping 填充 + 输出 + change log + 校验 | 当前 |
| P1 | 多历史 Excel 差异分析 + 自动识别变量 + 自动生成 schema | 待开始 |
| P2 | 交互式填写 (CLI 问答)、Jira 导入、多模板、PDF 输出、GUI | 待开始 |

## 技术栈

- Python 3.10+
- `openpyxl`：Excel 读写（唯一直接操作 Excel 的库）
- `pydantic`：数据结构校验
- `typer` + `rich`：CLI 工具
- `pytest`：测试

## 测试数据

`data/history_reports/` 放历史 Excel，`data/current_values/` 放本次变量 JSON。

## 坑与教训

- 合并单元格只能写入左上角，openpyxl 的 `merged_cells.ranges` 返回的 range 可能为空
- `cell.value` 可能返回 `None`、`str`、`int`、`float`、`datetime`，类型判断要全面
- 公式单元格默认不要覆盖（`cell.data_type == 'f'` 表示公式）
- openpyxl 不保留 .xlsm 宏，第一版只处理 .xlsx
