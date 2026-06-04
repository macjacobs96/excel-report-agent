"""
Harness: CLI 入口 —— 任务运行与编排系统

Usage:
    python main.py analyze --input ./data/history_reports --output ./project_config
    python main.py generate --template ./project_config/template.xlsx --values ./current_values.json --output ./output/report.xlsx
    python main.py validate --file ./output/report.xlsx --schema ./project_config/variable_schema.json
    python main.py diff --old ./history/report_v1.xlsx --new ./output/report_v2.xlsx
    python main.py ask
"""

import json
import sys
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

# 将项目根目录加入 sys.path
sys.path.insert(0, str(Path(__file__).parent))

from src.excel_reader import ExcelReader
from src.diff_analyzer import DiffAnalyzer
from src.variable_detector import VariableDetector
from src.template_builder import TemplateBuilder
from src.report_generator import ReportGenerator
from src.validator import Validator
from src.change_logger import ChangeLogger

app = typer.Typer(
    name="excel-report-agent",
    help="测试报告 Excel 自动生成系统",
    add_completion=False,
)

console = Console()


@app.command()
def analyze(
    input_dir: str = typer.Option("./data/history_reports", "--input", "-i", help="历史 Excel 文件夹路径"),
    output_dir: str = typer.Option("./project_config", "--output", "-o", help="输出配置文件夹路径"),
):
    """
    分析历史 Excel 报告，生成模板和变量配置。

    读取 input_dir 下的所有 .xlsx 文件，比较各版本差异，
    自动识别可变字段，生成 template.xlsx、variable_schema.json、
    cell_mapping.json 和 analysis_result.json。
    """
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    xlsx_files = sorted(input_path.glob("*.xlsx"))
    if not xlsx_files:
        console.print(f"[red]错误:[/red] {input_dir} 中没有找到 .xlsx 文件")
        raise typer.Exit(1)

    console.print(f"[bold]分析 {len(xlsx_files)} 个历史 Excel 文件...[/bold]\n")

    # Step 1: 读取所有历史 Excel
    reader = ExcelReader()
    all_data = {}
    for f in xlsx_files:
        console.print(f"  读取: {f.name}")
        all_data[f.name] = reader.read(f)
    console.print()

    # Step 2: 差异分析
    console.print("[bold]差异分析...[/bold]")
    analyzer = DiffAnalyzer()
    diff_result = analyzer.analyze(all_data)
    console.print(f"  发现 {len(diff_result.get('candidate_variables', []))} 个候选变量\n")

    # Step 3: 变量识别与归类
    console.print("[bold]变量识别...[/bold]")
    detector = VariableDetector()
    analysis = detector.detect(diff_result)
    console.print(f"  识别 {len(analysis.get('variables', {}))} 个业务字段\n")

    # Step 4: 保存分析结果
    with open(output_path / "analysis_result.json", "w", encoding="utf-8") as f:
        json.dump(analysis, f, ensure_ascii=False, indent=2)

    # Step 5: 生成 variable_schema.json 和 cell_mapping.json
    schema = detector.build_schema(analysis)
    with open(output_path / "variable_schema.json", "w", encoding="utf-8") as f:
        json.dump(schema, f, ensure_ascii=False, indent=2)

    mapping = detector.build_cell_mapping(analysis)
    with open(output_path / "cell_mapping.json", "w", encoding="utf-8") as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)

    # Step 6: 生成模板
    console.print("[bold]生成模板...[/bold]")
    builder = TemplateBuilder()
    template_path = output_path / "template.xlsx"
    latest_file = xlsx_files[-1]  # 最新版本作为模板
    builder.build(latest_file, template_path)
    console.print(f"  模板: {template_path}\n")

    # Step 7: 保存 template_meta.json
    meta = {
        "template_file": "template.xlsx",
        "source_files": [f.name for f in xlsx_files],
        "created_at": __import__("datetime").datetime.now().isoformat(),
        "sheet_names": list(all_data[xlsx_files[0].name].keys()),
    }
    with open(output_path / "template_meta.json", "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

    # 展示结果
    table = Table(title="识别到的变量字段")
    table.add_column("字段 Key", style="cyan")
    table.add_column("标签", style="green")
    table.add_column("类型")
    table.add_column("必填", justify="center")
    table.add_column("映射单元格数", justify="right")

    for key, var_info in schema.items():
        cell_count = sum(
            1 for m in mapping.get(key, [])
        )
        table.add_row(
            key,
            var_info.get("label", ""),
            var_info.get("type", "string"),
            "是" if var_info.get("required") else "否",
            str(cell_count),
        )

    console.print(table)
    console.print(f"\n[bold green]分析完成![/bold green] 配置已保存到 {output_path}/")


# === 全局替换辅助函数 ===

def _build_global_replacements(template_path, current_values, cell_mapping):
    """从模板中提取旧值，构建全局替换字典（旧→新版本号、旧→新人名）"""
    import zipfile, re
    from pathlib import Path

    template = Path(template_path)
    if not template.exists():
        return {}

    global_repl = {}

    with zipfile.ZipFile(template) as zf:
        names = set(zf.namelist())

        if 'xl/sharedStrings.xml' not in names:
            return {}

        ss = zf.read('xl/sharedStrings.xml').decode()

        def get_ss(idx):
            matches = list(re.finditer(r'<si>', ss))
            if idx >= len(matches):
                return ''
            start = matches[idx].end()
            end = ss.find('</si>', start)
            texts = re.findall(r'<t[^>]*>(.*?)</t>', ss[start:end], re.DOTALL)
            from xml.sax.saxutils import unescape
            return unescape(''.join(texts)).strip()

        # Find sheet files
        import xml.etree.ElementTree as ET
        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        wb_root = ET.fromstring(zf.read('xl/workbook.xml'))
        rels_root = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
        rid_to_target = {rel.get('Id'): rel.get('Target') for rel in rels_root}

        sheet_files = {}
        for sh in wb_root.findall(f'{{{ns}}}sheets/{{{ns}}}sheet'):
            rid = sh.get(f'{{{rns}}}id')
            if rid in rid_to_target:
                sheet_files[sh.get('name')] = f'xl/{rid_to_target[rid]}'

        def read_old_cell(sheet_name, cell_ref, new_key):
            nonlocal global_repl
            if sheet_name not in sheet_files:
                return
            xml = zf.read(sheet_files[sheet_name]).decode()
            m = re.search(r'<c\s+r="' + re.escape(cell_ref) + r'"[^>]*><v>(\d+)</v>', xml)
            if not m:
                return
            val = get_ss(int(m.group(1)))
            new_val = current_values.get(new_key, current_values.get('test_version', ''))
            if val and new_val and str(val) != str(new_val):
                global_repl[val] = str(new_val)
                trimmed = val.rstrip()
                if trimmed != val:
                    global_repl[trimmed] = str(new_val)

        # Version: read from Test Info C12
        read_old_cell('测试信息Test Info', 'C12', 'test_software_version')

        # Cover names: 编制/审核/批准 → replace old name with new
        for ref, key in [('F18', 'author_name'), ('F19', 'reviewer_name'), ('F20', 'approver_name')]:
            if '封面Cover' in sheet_files:
                xml = zf.read(sheet_files['封面Cover']).decode()
                m = re.search(r'<c\s+r="' + re.escape(ref) + r'"[^>]*><v>(\d+)</v>', xml)
                if m:
                    old_full = get_ss(int(m.group(1)))
                    old_name = old_full.split('/')[0].strip() if old_full else ''
                    new_full = current_values.get(key, '')
                    new_name = new_full.split('/')[0].strip() if isinstance(new_full, str) and '/' in str(new_full) else str(new_full) if new_full else ''
                    if old_name and new_name and old_name != new_name:
                        global_repl[old_name] = new_name

    return global_repl


@app.command()
def generate(
    template: str = typer.Option(..., "--template", "-t", help="模板 Excel 文件路径"),
    values: str = typer.Option(..., "--values", "-v", help="当前版本变量 JSON 文件路径"),
    output: str = typer.Option(..., "--output", "-o", help="输出 Excel 报告路径"),
    mapping: Optional[str] = typer.Option(None, "--mapping", "-m", help="cell_mapping.json 路径（默认同目录）"),
):
    """
    根据模板和变量值生成新版测试报告 Excel。
    """
    template_path = Path(template)
    values_path = Path(values)
    output_path = Path(output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if not template_path.exists():
        console.print(f"[red]错误:[/red] 模板文件不存在: {template}")
        raise typer.Exit(1)
    if not values_path.exists():
        console.print(f"[red]错误:[/red] 变量文件不存在: {values}")
        raise typer.Exit(1)

    # 加载配置
    if mapping is None:
        mapping = str(template_path.parent / "cell_mapping.json")
    mapping_path = Path(mapping)

    with open(values_path, encoding="utf-8") as f:
        current_values = json.load(f)

    cell_mapping = {}
    schema = {}
    if mapping_path.exists():
        with open(mapping_path, encoding="utf-8") as f:
            cell_mapping = json.load(f)

    schema_path = template_path.parent / "variable_schema.json"
    if schema_path.exists():
        with open(schema_path, encoding="utf-8") as f:
            schema = json.load(f)

    console.print(f"[bold]生成测试报告...[/bold]")
    console.print(f"  模板: {template_path.name}")
    console.print(f"  变量: {len(current_values)} 个字段")
    console.print(f"  映射: {len(cell_mapping)} 个字段映射\n")

    # Step 1: 填充变量
    generator = ReportGenerator()

    # 自动构建全局替换规则（旧版本号、旧人名 → 新值）
    global_repl = _build_global_replacements(template_path, current_values, cell_mapping)

    changes = generator.generate(
        template_path, current_values, cell_mapping, output_path,
        global_replacements=global_repl,
    )
    console.print(f"  [green]报告已生成:[/green] {output_path}")

    # Step 2: 输出变更日志
    logger = ChangeLogger()
    change_log_path = output_path.parent / "change_log.xlsx"
    logger.export(changes, change_log_path)
    console.print(f"  [green]变更日志:[/green] {change_log_path}")

    # Step 3: 校验
    if schema:
        validator = Validator()
        vresult = validator.validate(output_path, schema, cell_mapping)
        vreport_path = output_path.parent / "validation_report.json"
        with open(vreport_path, "w", encoding="utf-8") as f:
            json.dump(vresult, f, ensure_ascii=False, indent=2)

        if vresult.get("errors"):
            console.print(f"\n[yellow]校验警告 {len(vresult['errors'])} 条:[/yellow]")
            for err in vresult["errors"]:
                console.print(f"  - {err}")
        else:
            console.print(f"\n[green]校验通过![/green]")

        console.print(f"  [green]校验报告:[/green] {vreport_path}")

    console.print(f"\n[bold green]生成完成![/bold green] 共修改 {len(changes)} 个单元格")


@app.command()
def validate(
    file: str = typer.Option(..., "--file", "-f", help="要校验的 Excel 文件路径"),
    schema: str = typer.Option(..., "--schema", "-s", help="variable_schema.json 路径"),
    mapping: Optional[str] = typer.Option(None, "--mapping", "-m", help="cell_mapping.json 路径"),
):
    """
    校验生成的测试报告是否符合预期。
    """
    file_path = Path(file)
    schema_path = Path(schema)

    if not file_path.exists():
        console.print(f"[red]错误:[/red] 文件不存在: {file}")
        raise typer.Exit(1)

    with open(schema_path, encoding="utf-8") as f:
        schema_data = json.load(f)

    cell_mapping = {}
    if mapping:
        mapping_path = Path(mapping)
        if mapping_path.exists():
            with open(mapping_path, encoding="utf-8") as f:
                cell_mapping = json.load(f)

    validator = Validator()
    result = validator.validate(file_path, schema_data, cell_mapping)

    console.print(f"\n[bold]校验报告:[/bold] {file_path.name}\n")

    if result.get("errors"):
        console.print(f"[red]错误 ({len(result['errors'])}):[/red]")
        for e in result["errors"]:
            console.print(f"  - {e}")
    else:
        console.print("[green]校验通过，无错误[/green]")

    if result.get("warnings"):
        console.print(f"\n[yellow]警告 ({len(result['warnings'])}):[/yellow]")
        for w in result["warnings"]:
            console.print(f"  - {w}")

    console.print(f"\n[bold]统计:[/bold]")
    console.print(f"  总字段: {result.get('total_fields', 0)}")
    console.print(f"  已填充: {result.get('filled_fields', 0)}")
    console.print(f"  未填充: {result.get('empty_fields', 0)}")


@app.command()
def diff(
    old: str = typer.Option(..., "--old", help="旧版 Excel 文件路径"),
    new: str = typer.Option(..., "--new", help="新版 Excel 文件路径"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="输出变更日志路径"),
):
    """
    对比两个版本的测试报告差异。
    """
    old_path = Path(old)
    new_path = Path(new)

    if not old_path.exists():
        console.print(f"[red]错误:[/red] 文件不存在: {old}")
        raise typer.Exit(1)
    if not new_path.exists():
        console.print(f"[red]错误:[/red] 文件不存在: {new}")
        raise typer.Exit(1)

    reader = ExcelReader()
    old_data = reader.read(old_path)
    new_data = reader.read(new_path)

    analyzer = DiffAnalyzer()
    diffs = analyzer.compare_two(old_data, new_data)

    if diffs:
        table = Table(title=f"差异对比: {old_path.name} vs {new_path.name}")
        table.add_column("Sheet", style="cyan")
        table.add_column("Cell", style="yellow")
        table.add_column("旧值")
        table.add_column("新值", style="green")

        for d in diffs:
            table.add_row(d["sheet"], d["cell"], str(d["old_value"]), str(d["new_value"]))

        console.print(table)
        console.print(f"\n共 {len(diffs)} 处不同")
    else:
        console.print("[green]两个文件完全相同[/green]")

    if output:
        output_path = Path(output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        logger = ChangeLogger()
        logger.export(diffs, output_path)
        console.print(f"[green]变更日志已保存:[/green] {output_path}")


@app.command()
def ask(
    schema: Optional[str] = typer.Option(None, "--schema", "-s", help="variable_schema.json 路径"),
    output: str = typer.Option("./data/current_values/current_values.json", "--output", "-o", help="输出 JSON 路径"),
):
    """
    交互式问答填写本次版本变量。
    """
    schema_data = {}
    schema_path = Path(schema) if schema else None
    if schema_path and schema_path.exists():
        with open(schema_path, encoding="utf-8") as f:
            schema_data = json.load(f)

    if not schema_data:
        console.print("[yellow]未提供 schema，使用自由问答模式[/yellow]\n")

    values = {}
    console.print("[bold]请输入本次版本信息:[/bold]\n")

    for key, info in schema_data.items():
        required = info.get("required", False)
        label = info.get("label", key)
        desc = info.get("description", "")
        default = info.get("default", "")

        prompt = f"  {label}"
        if required:
            prompt += " [red]*[/red]"
        if desc:
            prompt += f" ({desc})"
        prompt += ": "

        if default:
            user_input = typer.prompt(prompt.strip(), default=default)
        else:
            user_input = typer.prompt(prompt.strip(), default="")

        if user_input:
            values[key] = user_input
        elif required:
            console.print(f"  [red]必填字段，请重新输入[/red]")
            user_input = typer.prompt(prompt.strip(), default="")
            values[key] = user_input
        else:
            values[key] = ""

    # 自由模式：追加更多字段
    console.print("\n[dim](输入空行结束)[/dim]")
    while True:
        extra_key = typer.prompt("  额外字段名（回车跳过）", default="", show_default=False)
        if not extra_key:
            break
        extra_val = typer.prompt(f"  {extra_key} 的值", default="")
        values[extra_key] = extra_val

    output_path = Path(output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(values, f, ensure_ascii=False, indent=2)

    console.print(f"\n[green]变量已保存:[/green] {output_path}")
    console.print(json.dumps(values, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    app()
