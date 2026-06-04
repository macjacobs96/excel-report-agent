"""
ReportGenerator v6 —— Excel 报告生成工具

支持：
1. 单元格级别精确替换（cell_mapping）
2. 全局文本搜索替换（global_replacements）—— 处理嵌入在长文本中的版本号、人名等
"""

import os
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from xml.sax.saxutils import escape as xml_escape, unescape


class ReportGenerator:
    """报告生成器 v6 — 支持全局文本替换"""

    def generate(
        self,
        template_path: Path,
        current_values: Dict[str, Any],
        cell_mapping: Dict[str, List[Dict[str, str]]],
        output_path: Path,
        global_replacements: Optional[Dict[str, str]] = None,
    ) -> List[Dict[str, Any]]:
        """
        Args:
            template_path: 模板 xlsx
            current_values: 当前变量值
            cell_mapping: 变量映射
            output_path: 输出路径
            global_replacements: {旧文本: 新文本} 全局替换
        """
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # Flatten modifications
        cell_mods = []
        for var_key, value in current_values.items():
            if var_key not in cell_mapping:
                continue
            mappings = cell_mapping[var_key]
            if not isinstance(mappings, list):
                mappings = [mappings]
            for mapping in mappings:
                cell_mods.append((
                    mapping.get("sheet", ""),
                    mapping.get("cell", ""),
                    str(value) if value else "",
                    var_key,
                    mapping.get("overwrite_formula", False),
                ))

        changes = []
        sheets_cache = {}  # sheet_name → xml_string

        with tempfile.TemporaryDirectory() as tmpdir:
            extract_dir = os.path.join(tmpdir, "extracted")
            with zipfile.ZipFile(template_path, 'r') as zf:
                zf.extractall(extract_dir)

            sheet_map = self._build_sheet_map(extract_dir)

            # Phase 1: Cell-level modifications
            for sheet_name, cell_ref, new_value, var_key, overwrite_formula in cell_mods:
                change = self._modify_one_cell(
                    extract_dir, sheet_map, sheets_cache,
                    sheet_name, cell_ref, new_value, var_key, overwrite_formula
                )
                changes.append(change)

            # Phase 2: Global text replace in sharedStrings.xml (fast, covers ALL shared string cells)
            if global_replacements:
                gr_changes = self._global_replace_ss(extract_dir, global_replacements)
                changes.extend(gr_changes)

            # Write back modified sheets
            for sheet_name, xml_data in sheets_cache.items():
                sheet_path = os.path.join(extract_dir, sheet_map[sheet_name])
                with open(sheet_path, 'w', encoding='utf-8') as f:
                    f.write(xml_data)

            # Repack
            with zipfile.ZipFile(str(output_path), 'w', zipfile.ZIP_DEFLATED) as out_zf:
                for root, dirs, files in os.walk(extract_dir):
                    for fname in sorted(files):
                        full_path = os.path.join(root, fname)
                        arcname = os.path.relpath(full_path, extract_dir)
                        out_zf.write(full_path, arcname)

        return changes

    def _global_replace_ss(
        self, extract_dir, replacements: Dict[str, str]
    ) -> List[Dict]:
        """
        Global text replace in sharedStrings.xml only.
        This covers ALL cells that use shared strings (t="s") — which is most text cells.
        Fast single-pass operation.
        """
        ss_path = os.path.join(extract_dir, "xl", "sharedStrings.xml")
        if not os.path.exists(ss_path):
            return []

        with open(ss_path, 'r', encoding='utf-8') as f:
            ss_xml = f.read()

        changes = []
        for old_text, new_text in replacements.items():
            if old_text == new_text:
                continue
            count_before = ss_xml.count(old_text)
            if count_before == 0:
                continue

            ss_xml = ss_xml.replace(old_text, xml_escape(new_text))
            changes.append({
                "sheet": "sharedStrings",
                "cell": "-",
                "field": "全局替换",
                "old_value": f"'{old_text}' ×{count_before}",
                "new_value": new_text,
                "status": "SUCCESS",
            })

        if changes:
            with open(ss_path, 'w', encoding='utf-8') as f:
                f.write(ss_xml)

        return changes

    def _build_sheet_map(self, extract_dir: str) -> Dict[str, str]:
        import xml.etree.ElementTree as ET
        wb_xml = os.path.join(extract_dir, "xl", "workbook.xml")
        rels_path = os.path.join(extract_dir, "xl", "_rels", "workbook.xml.rels")
        if not os.path.exists(wb_xml) or not os.path.exists(rels_path):
            return {}

        rels_tree = ET.parse(rels_path)
        rels_map = {rel.get("Id", ""): rel.get("Target", "") for rel in rels_tree.getroot()}

        wb_tree = ET.parse(wb_xml)
        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        sheet_map = {}
        for sheet_el in wb_tree.getroot().findall(f"{{{ns}}}sheets/{{{ns}}}sheet"):
            name = sheet_el.get("name", "")
            rid = sheet_el.get(f"{{{rns}}}id", "")
            if name and rid and rid in rels_map:
                target = rels_map[rid]
                # Handle both relative (worksheets/sheet1.xml) and absolute (/xl/worksheets/sheet1.xml) paths
                if target.startswith('/'):
                    sheet_map[name] = target.lstrip('/')
                else:
                    sheet_map[name] = os.path.normpath(os.path.join("xl", target))

        return sheet_map

    def _modify_one_cell(
        self, extract_dir, sheet_map, sheets_cache,
        sheet_name, cell_ref, new_value, var_key, overwrite_formula
    ) -> Dict[str, Any]:
        if sheet_name not in sheet_map:
            return {"sheet": sheet_name, "cell": cell_ref, "field": var_key,
                    "old_value": "?", "new_value": new_value,
                    "status": f"FAIL: Sheet不存在"}

        sheet_path = os.path.join(extract_dir, sheet_map[sheet_name])
        if not os.path.exists(sheet_path):
            return {"sheet": sheet_name, "cell": cell_ref, "field": var_key,
                    "old_value": "?", "new_value": new_value,
                    "status": "FAIL: 文件不存在"}

        if sheet_name not in sheets_cache:
            with open(sheet_path, 'r', encoding='utf-8') as f:
                sheets_cache[sheet_name] = f.read()
        sheet_xml = sheets_cache[sheet_name]

        cell_re = re.compile(
            r'<c\s+r="' + re.escape(cell_ref) + r'"[^>]*>(.*?)</c>',
            re.DOTALL
        )
        cmatch = cell_re.search(sheet_xml)
        if not cmatch:
            sc_re = re.compile(r'<c\s+r="' + re.escape(cell_ref) + r'"[^>/]*/>')
            scmatch = sc_re.search(sheet_xml)
            if not scmatch:
                return {"sheet": sheet_name, "cell": cell_ref, "field": var_key,
                        "old_value": "<空>", "new_value": new_value,
                        "status": "SKIPPED: 单元格不存在"}
            old_val = "<空>"
        else:
            old_val = self._extract_text(cmatch.group(1))

        if old_val == new_value:
            return {"sheet": sheet_name, "cell": cell_ref, "field": var_key,
                    "old_value": old_val, "new_value": new_value,
                    "status": "UNCHANGED"}

        if cmatch:
            cell_content = cmatch.group(1)
            if '<f>' in cell_content or '<f ' in cell_content:
                if not overwrite_formula:
                    return {"sheet": sheet_name, "cell": cell_ref, "field": var_key,
                            "old_value": old_val, "new_value": new_value,
                            "status": "SKIPPED (公式保护)"}

        style_attr = ""
        if cmatch:
            style_m = re.search(r'<c\s+r="' + re.escape(cell_ref) + r'"([^>]*)>', cmatch.group())
            if style_m:
                sm = re.search(r'\ss="\d+"', style_m.group(1))
                if sm:
                    style_attr = sm.group()

        new_cell = (
            f'<c r="{cell_ref}"{style_attr} t="inlineStr">'
            f'<is><t>{xml_escape(new_value)}</t></is>'
            f'</c>'
        )

        if cmatch:
            sheet_xml = sheet_xml[:cmatch.start()] + new_cell + sheet_xml[cmatch.end():]
        else:
            sheet_xml = sheet_xml[:scmatch.start()] + new_cell + sheet_xml[scmatch.end():]

        sheets_cache[sheet_name] = sheet_xml

        return {"sheet": sheet_name, "cell": cell_ref, "field": var_key,
                "old_value": old_val, "new_value": new_value,
                "status": "SUCCESS"}

    def _extract_text(self, cell_content: str) -> str:
        vm = re.search(r'<v>(.*?)</v>', cell_content, re.DOTALL)
        if vm:
            val = unescape(vm.group(1).strip())
            if val:
                return val
        tms = re.findall(r'<t[^>]*>(.*?)</t>', cell_content, re.DOTALL)
        if tms:
            return unescape("".join(t for t in tms if t))
        fm = re.search(r'<f[^>]*>(.*?)</f>', cell_content, re.DOTALL)
        if fm:
            return "=" + unescape(fm.group(1).strip())
        return ""
