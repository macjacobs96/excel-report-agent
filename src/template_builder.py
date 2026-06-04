"""
TemplateBuilder —— 模板生成工具

从历史 Excel 中选择一个作为模板，复制到项目配置目录。
"""

import shutil
from pathlib import Path


class TemplateBuilder:
    """模板构建器"""

    def build(self, source_file: Path, target_path: Path) -> None:
        """
        将源 Excel 复制为模板文件。

        第一版策略：直接复制最新版本 Excel 作为模板。
        cell_mapping.json 控制哪些单元格被覆盖，
        而非在 Excel 中插入占位符。

        Args:
            source_file: 源 Excel 文件（通常是最新版本）
            target_path: 模板目标路径
        """
        target_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_file, target_path)

    def build_from_earliest(self, source_file: Path, target_path: Path) -> None:
        """
        使用最早版本作为模板（某些场景下更干净）。

        Args:
            source_file: 源 Excel 文件（通常是最早版本）
            target_path: 模板目标路径
        """
        self.build(source_file, target_path)
