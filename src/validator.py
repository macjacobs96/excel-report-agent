"""
Validator —— 校验工具

校验生成的测试报告是否符合预期：
1. required 字段是否填写
2. Sheet 是否存在
3. Cell 是否存在
4. 是否存在旧版本号残留
5. 文件是否能正常打开
"""

import json
from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl


class Validator:
    """校验器"""

    def validate(
        self,
        file_path: Path,
        schema: Dict[str, Any],
        cell_mapping: Optional[Dict[str, Any]] = None,
        old_values: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """
        校验 Excel 文件。

        Args:
            file_path: 待校验的 Excel 文件
            schema: variable_schema.json 结构
            cell_mapping: cell_mapping.json 结构（可选）
            old_values: 旧版变量值，用于检查残留（可选）

        Returns:
            {
                "errors": [...],
                "warnings": [...],
                "total_fields": int,
                "filled_fields": int,
                "empty_fields": int,
            }
        """
        result = {
            "errors": [],
            "warnings": [],
            "total_fields": 0,
            "filled_fields": 0,
            "empty_fields": 0,
        }

        # 1. 检查文件是否能正常打开
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
        except Exception as e:
            result["errors"].append(f"文件无法打开: {e}")
            return result

        sheet_names = wb.sheetnames

        # 2. 检查 cell_mapping 中的 Sheet 和 Cell 是否存在
        if cell_mapping:
            for var_key, mappings in cell_mapping.items():
                if not isinstance(mappings, list):
                    mappings = [mappings]

                for mapping in mappings:
                    sheet = mapping.get("sheet", "")
                    cell = mapping.get("cell", "")

                    if sheet not in sheet_names:
                        result["errors"].append(
                            f"cell_mapping 中引用的 Sheet '{sheet}' 不存在于生成的文件中"
                        )
                        continue

                    try:
                        cell_value = wb[sheet][cell].value
                    except (ValueError, KeyError):
                        result["errors"].append(
                            f"cell_mapping 中引用的 Cell '{cell}' 在 Sheet '{sheet}' 中不存在"
                        )
                        continue

                    result["total_fields"] += 1

                    if cell_value is not None and str(cell_value).strip():
                        result["filled_fields"] += 1
                    else:
                        result["empty_fields"] += 1

                        # 检查是否为 required 字段
                        field_schema = schema.get(var_key, {})
                        if field_schema.get("required", False):
                            result["errors"].append(
                                f"必填字段 '{var_key}' ({field_schema.get('label', var_key)}) "
                                f"在 {sheet}!{cell} 中为空"
                            )

        # 3. 检查是否有旧版本号残留
        if old_values and cell_mapping:
            for var_key, mappings in cell_mapping.items():
                if var_key == "test_version" and var_key in old_values:
                    old_ver = str(old_values[var_key])
                    if not isinstance(mappings, list):
                        mappings = [mappings]
                    for mapping in mappings:
                        sheet = mapping.get("sheet", "")
                        cell = mapping.get("cell", "")
                        if sheet in sheet_names:
                            try:
                                current_val = str(wb[sheet][cell].value or "")
                                if current_val == old_ver:
                                    result["warnings"].append(
                                        f"旧版本号 '{old_ver}' 可能残留在 {sheet}!{cell}"
                                    )
                            except (ValueError, KeyError):
                                pass

        # 4. 检查 Sheet 数量（粗略的质量检查）
        if len(sheet_names) == 0:
            result["errors"].append("文件中没有任何 Sheet")
        if isinstance(cell_mapping, dict) and len(cell_mapping) > 0:
            # 检查是否有 Sheet 名不匹配
            mapping_sheets = set()
            for _, mappings in cell_mapping.items():
                ms = mappings if isinstance(mappings, list) else [mappings]
                for m in ms:
                    mapping_sheets.add(m.get("sheet", ""))

            missing_sheets = mapping_sheets - set(sheet_names)
            for s in missing_sheets:
                result["errors"].append(f"配置中引用的 Sheet '{s}' 在文件中找不到")

        wb.close()
        return result
