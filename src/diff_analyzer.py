"""
DiffAnalyzer —— Excel 差异分析工具

比较多个历史 Excel 的差异，识别候选变量。
"""

from typing import Any, Dict, List
from collections import defaultdict


class DiffAnalyzer:
    """差异分析器"""

    def analyze(self, all_data: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
        """
        分析多个 Excel 文件的差异。

        Args:
            all_data: {filename: {sheet_name: SheetData}}

        Returns:
            {
                "candidate_variables": [...],
                "file_count": int,
                "files": [...],
                "common_sheets": [...]
            }
        """
        filenames = list(all_data.keys())
        if len(filenames) < 2:
            return {
                "candidate_variables": [],
                "file_count": len(filenames),
                "files": filenames,
                "common_sheets": [],
                "note": "需要至少 2 个文件进行差异分析",
            }

        # 找出所有文件共有的 Sheet
        sheet_sets = [
            set(data.keys()) for data in all_data.values()
        ]
        common_sheets = list(sheet_sets[0].intersection(*sheet_sets[1:]))
        all_sheets = list(set.union(*sheet_sets))

        # 按 Sheet + Cell 对齐，比较值
        candidate_variables = []

        for sheet_name in common_sheets:
            # 收集所有文件中该 Sheet 的单元格
            sheet_cells = defaultdict(list)  # cell_ref -> [(filename, value)]
            for fname in filenames:
                sheet_data = all_data[fname].get(sheet_name)
                if sheet_data is None:
                    continue
                cells_dict = sheet_data.cells if hasattr(sheet_data, 'cells') else sheet_data
                for cell_ref, cell_info in cells_dict.items():
                    value = cell_info.value if hasattr(cell_info, 'value') else cell_info
                    sheet_cells[cell_ref].append((fname, value))

            # 找出值发生变化的单元格
            for cell_ref, entries in sheet_cells.items():
                if len(entries) < 2:
                    continue

                # 如果所有文件的同一个 cell 值都相同，跳过（固定内容）
                values = [str(v) if v is not None else "" for _, v in entries]
                unique_values = list(set(values))

                if len(unique_values) <= 1:
                    continue  # 完全没有变化

                candidate_variables.append({
                    "sheet": sheet_name,
                    "cell": cell_ref,
                    "values": {fname: val for fname, val in entries if val is not None},
                    "unique_values": unique_values,
                    "changed_in_files": len([1 for _, v in entries if v]),  # 非空有效变化次数
                })

        # 按 Sheet + Cell 排序
        candidate_variables.sort(key=lambda x: (x["sheet"], x["cell"]))

        return {
            "candidate_variables": candidate_variables,
            "file_count": len(filenames),
            "files": filenames,
            "common_sheets": common_sheets,
            "all_sheets": all_sheets,
        }

    def compare_two(
        self,
        old_data: Dict[str, Any],
        new_data: Dict[str, Any],
    ) -> List[Dict[str, Any]]:
        """
        对比两个 Excel 的差异（用于 diff 命令）。

        Returns:
            [{"sheet": "...", "cell": "B4", "old_value": "张三", "new_value": "李四"}]
        """
        diffs = []

        all_sheets = set(list(old_data.keys()) + list(new_data.keys()))

        for sheet_name in all_sheets:
            old_sheet = old_data.get(sheet_name)
            new_sheet = new_data.get(sheet_name)

            if old_sheet is None:
                new_cells = new_sheet.cells if hasattr(new_sheet, 'cells') else new_sheet
                for cell_ref, info in new_cells.items():
                    diffs.append({
                        "sheet": sheet_name,
                        "cell": cell_ref,
                        "old_value": None,
                        "new_value": info.value if hasattr(info, 'value') else info,
                    })
                continue

            if new_sheet is None:
                old_cells = old_sheet.cells if hasattr(old_sheet, 'cells') else old_sheet
                for cell_ref, info in old_cells.items():
                    diffs.append({
                        "sheet": sheet_name,
                        "cell": cell_ref,
                        "old_value": info.value if hasattr(info, 'value') else info,
                        "new_value": None,
                    })
                continue

            old_cells = old_sheet.cells if hasattr(old_sheet, 'cells') else old_sheet
            new_cells = new_sheet.cells if hasattr(new_sheet, 'cells') else new_sheet

            all_cells = set(old_cells.keys()) | set(new_cells.keys())

            for cell_ref in sorted(all_cells):
                old_val = None
                new_val = None

                if cell_ref in old_cells:
                    old_val = old_cells[cell_ref].value if hasattr(old_cells[cell_ref], 'value') else old_cells[cell_ref]
                if cell_ref in new_cells:
                    new_val = new_cells[cell_ref].value if hasattr(new_cells[cell_ref], 'value') else new_cells[cell_ref]

                if str(old_val) != str(new_val):
                    diffs.append({
                        "sheet": sheet_name,
                        "cell": cell_ref,
                        "old_value": old_val,
                        "new_value": new_val,
                    })

        return diffs
