"""
ExcelReader —— Excel 文件读取工具

读取 .xlsx 文件，输出结构化数据：
- 所有 Sheet 名称
- 每个 Sheet 的非空单元格（值、数据类型、格式信息）
- 合并单元格信息
- 公式信息
"""

from pathlib import Path
from typing import Any, Dict, List, Optional
from dataclasses import dataclass, field

import openpyxl
from openpyxl.utils import get_column_letter


@dataclass
class CellInfo:
    """单个单元格信息"""
    value: Any = None
    data_type: str = "s"  # s=string, n=number, d=date, f=formula, b=bool, e=error, z=empty
    font_name: Optional[str] = None
    font_size: Optional[int] = None
    font_bold: bool = False
    fill_color: Optional[str] = None
    alignment_h: Optional[str] = None
    alignment_v: Optional[str] = None
    border_style: Optional[str] = None
    number_format: Optional[str] = None


@dataclass
class MergedRange:
    """合并单元格信息"""
    start_cell: str
    end_cell: str


@dataclass
class SheetData:
    """单个 Sheet 的数据"""
    name: str
    cells: Dict[str, CellInfo] = field(default_factory=dict)
    merged_ranges: List[MergedRange] = field(default_factory=list)
    frozen_panes: Optional[str] = None
    row_count: int = 0
    col_count: int = 0


class ExcelReader:
    """Excel 文件读取器"""

    def read(self, filepath: Path) -> Dict[str, SheetData]:
        """
        读取 Excel 文件，返回 {sheet_name: SheetData} 结构。

        Args:
            filepath: .xlsx 文件路径

        Returns:
            字典，键为 Sheet 名称，值为该 Sheet 的结构化数据
        """
        wb = openpyxl.load_workbook(filepath, data_only=False)

        result = {}
        for ws in wb.worksheets:
            sheet_data = self._read_sheet(ws)
            result[ws.title] = sheet_data

        wb.close()
        return result

    def _read_sheet(self, ws) -> SheetData:
        """读取单个 Sheet"""
        data = SheetData(name=ws.title)

        # 读取冻结窗格
        data.frozen_panes = ws.freeze_panes

        # 读取合并单元格
        for merged_range in ws.merged_cells.ranges:
            data.merged_ranges.append(
                MergedRange(
                    start_cell=str(merged_range.min_col) + str(merged_range.min_row),
                    end_cell=str(merged_range.max_col) + str(merged_range.max_row),
                )
            )

        # 读取非空单元格
        max_row = ws.max_row or 0
        max_col = ws.max_column or 0
        data.row_count = max_row
        data.col_count = max_col

        for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
            for cell in row:
                if cell.value is None:
                    continue

                cell_ref = f"{get_column_letter(cell.column)}{cell.row}"
                info = CellInfo()

                info.value = cell.value
                info.data_type = self._get_data_type(cell)

                # 字体信息
                if cell.font:
                    info.font_name = cell.font.name
                    info.font_size = cell.font.size
                    info.font_bold = cell.font.bold or False

                # 填充色
                if cell.fill and cell.fill.fgColor and cell.fill.fgColor.rgb:
                    info.fill_color = str(cell.fill.fgColor.rgb)

                # 对齐方式
                if cell.alignment:
                    info.alignment_h = cell.alignment.horizontal
                    info.alignment_v = cell.alignment.vertical

                # 数字格式
                info.number_format = cell.number_format

                data.cells[cell_ref] = info

        return data

    def read_values_only(self, filepath: Path) -> Dict[str, Dict[str, Any]]:
        """
        简化读取：只取单元格值。

        Returns:
            {sheet_name: {"B4": "张三", "B5": "V1.0"}}
        """
        full_data = self.read(filepath)
        return {
            sheet_name: {cell_ref: info.value for cell_ref, info in sheet_data.cells.items()}
            for sheet_name, sheet_data in full_data.items()
        }

    @staticmethod
    def _get_data_type(cell) -> str:
        """判断单元格数据类型"""
        if cell.data_type == 'f':
            return 'f'  # formula
        if cell.data_type == 'n':
            # 数字类型：可能是真数字或日期
            if cell.number_format and 'yy' in str(cell.number_format).lower():
                return 'd'  # date
            return 'n'  # number
        if cell.data_type == 'b':
            return 'b'  # boolean
        if cell.data_type == 'e':
            return 'e'  # error
        if cell.data_type == 's':
            return 's'  # string
        if cell.data_type == 'str':
            return 's'  # string (openpyxl variant)
        return 'z'  # unknown
