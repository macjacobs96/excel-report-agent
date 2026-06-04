"""
VariableDetector —— 变量识别与归并工具 v2

针对车企测试报告真实格式优化：
- 版本号: ER9.1.07, ER8.1.34, DMS_2.1, V1.0.3
- 人名/日期: 田佳昊/2026-03-30, 李瑞鹏/龙航
- 标签-值对: B列=标签, C列=值 的 Sheet 布局
- 测试结果: 视觉OMS/DMS 中的 PASS/FAIL/NG
"""

import re
from typing import Any, Dict, List, Optional
from datetime import datetime
from collections import Counter


class VariableDetector:
    """变量检测器 v2"""

    # ================================================================
    # 字段类型定义
    # ================================================================
    FIELD_TYPES = {
        "test_software_version": {"label": "软件版本", "required": True, "description": "被测软件版本号"},
        "test_hardware_version": {"label": "硬件版本", "required": True, "description": "被测硬件版本号"},
        "tester_name": {"label": "测试人", "required": True, "description": "测试执行人员"},
        "dre_name": {"label": "零件工程师", "required": False, "description": "DRE 姓名"},
        "test_date": {"label": "测试日期", "required": True, "description": "测试日期"},
        "test_start_date": {"label": "测试开始日期", "required": True, "description": "开始日期"},
        "test_end_date": {"label": "测试结束日期", "required": True, "description": "结束日期"},
        "project_name": {"label": "项目名称", "required": True, "description": "项目名称/平台"},
        "dut_name": {"label": "被测件名称", "required": False, "description": "被测零部件名称"},
        "dut_config": {"label": "零件配置", "required": False, "description": "零件配置信息"},
        "test_device": {"label": "测试设备", "required": False, "description": "测试设备信息"},
        "test_environment": {"label": "测试环境", "required": False, "description": "实车/台架/模拟"},
        "test_conclusion": {"label": "测试结论", "required": False, "description": "整体结论"},
        "test_round": {"label": "测试轮次", "required": False, "description": "第几轮"},
        "jira_summary": {"label": "Jira/Bug汇总", "required": False, "description": "缺陷汇总"},
        "test_result_items": {"label": "测试项结果", "required": False, "description": "具体测试用例结果"},
        "cover_info": {"label": "封面信息", "required": False, "description": "封面页上的版本/人员/日期信息"},
        "unknown": {"label": "未知变量", "required": False, "description": "未自动分类的变更字段"},
    }

    # ================================================================
    # 版本号模式（车企常见格式）
    # ================================================================
    VERSION_PATTERNS = [
        re.compile(r'^[A-Z]{2,}\d+\.\d+(\.\d+)?(-\w+)?$'),  # ER9.1.07, DMS_2.1.0-beta
        re.compile(r'^[Vv]\d+\.\d+(\.\d+)?(-\w+)?$'),         # V1.0, v2.3.1-rc1
        re.compile(r'^\d+\.\d+\.\d+$'),                        # 9.1.07
        re.compile(r'^[A-Z]+\d+\.\d+$'),                       # ER9.1
    ]

    # ================================================================
    # Excel DateTime 检测
    # ================================================================
    @staticmethod
    def is_datetime(value: Any) -> bool:
        """判断值是否为 datetime"""
        if isinstance(value, datetime):
            return True
        # openpyxl 有时读取为字符串 '2026-05-30 00:00:00'
        if isinstance(value, str):
            dt_patterns = [
                re.compile(r'^\d{4}-\d{2}-\d{2}(\s\d{2}:\d{2}:\d{2})?$'),
                re.compile(r'^\d{4}/\d{1,2}/\d{1,2}$'),
                re.compile(r'^\d{4}年\d{1,2}月\d{1,2}日$'),
            ]
            return any(p.match(value.strip()) for p in dt_patterns)
        return False

    # ================================================================
    # 标签-值对关键词映射（用于匹配 B列标签→C列值 的布局）
    # ================================================================
    LABEL_KEYWORD_MAP = [
        (re.compile(r"测试软件版本|software.*version|SW.*[Vv]er", re.I), "test_software_version"),
        (re.compile(r"测试硬件版本|hardware.*version|HW.*[Vv]er|零件号", re.I), "test_hardware_version"),
        (re.compile(r"测试工程师|测试人员|测试人|tester.*name|executor", re.I), "tester_name"),
        (re.compile(r"零件工程师|DRE.*name", re.I), "dre_name"),
        (re.compile(r"测试开始日期|start.*date|开始日期", re.I), "test_start_date"),
        (re.compile(r"测试结束日期|end.*date|结束日期", re.I), "test_end_date"),
        (re.compile(r"项目名称|project.*name", re.I), "project_name"),
        (re.compile(r"被测件名称|DUT.*name|被测零件", re.I), "dut_name"),
        (re.compile(r"零件配置|configuration|配置", re.I), "dut_config"),
        (re.compile(r"测试设备|test.*device|设备", re.I), "test_device"),
        (re.compile(r"试验环境|specific.*envir|环境", re.I), "test_environment"),
        (re.compile(r"供应商名称|supplier.*name", re.I), "supplier_name"),
        (re.compile(r"供应商代码|supplier.*code", re.I), "supplier_code"),
        (re.compile(r"电源型号|power.*supply", re.I), "power_supply"),
        (re.compile(r"线缆|cable", re.I), "cable"),
        (re.compile(r"结论|conclusion|总结|判定", re.I), "test_conclusion"),
        (re.compile(r"版本|revision|release.*version", re.I), "test_software_version"),
    ]

    # ================================================================
    # 核心方法
    # ================================================================

    def detect(self, diff_result: Dict[str, Any]) -> Dict[str, Any]:
        """
        从差异分析结果中识别业务变量。

        Analysis process:
        1. Classify each candidate variable by: (a) label keyword near the cell,
           (b) value pattern matching, (c) sheet context
        2. Merge cells that belong to the same variable
        3. Separate "test result items" from "field variables"
        """
        candidates = diff_result.get("candidate_variables", [])
        variables = {}

        for candidate in candidates:
            sheet = candidate["sheet"]
            cell = candidate["cell"]
            values = candidate.get("unique_values", [])
            value_dict = candidate.get("values", {})

            if not values:
                continue

            # Step 1: Try to classify by sheet context and cell context
            key, confidence = self._classify_cell(sheet, cell, values, value_dict)

            if key not in variables:
                variables[key] = {
                    "cells": [],
                    "sample_values": values[:5],
                    "field_type": key,
                    "confidence": confidence,
                }

            # Don't add duplicate cells
            existing = variables[key]["cells"]
            if not any(e["sheet"] == sheet and e["cell"] == cell for e in existing):
                existing.append({"sheet": sheet, "cell": cell})

        return {"variables": variables}

    def _classify_cell(
        self, sheet: str, cell: str, values: List[Any], value_dict: Dict[str, Any]
    ) -> tuple:
        """对单个候选变量分类，返回 (key, confidence)"""

        # ---- Rule 1: Value pattern matching ----
        # Is it a version number?
        version_count = 0
        for v in values:
            vs = str(v).strip() if v else ""
            for p in self.VERSION_PATTERNS:
                if p.match(vs):
                    version_count += 1
                    break
        if version_count >= max(1, len(values) * 0.5):
            return "test_software_version", 0.9

        # Is it a datetime?
        date_count = sum(1 for v in values if self.is_datetime(v))
        if date_count >= max(1, len(values) * 0.5):
            return "test_date", 0.85

        # Is it a tester name? (format: "Name" or "Name1/Name2")
        # Exclude common non-name Chinese text (equipment, environment, etc.)
        NON_NAME_KEYWORDS = {"台架", "实车", "线刷", "模拟", "环境", "设备", "电源", "线缆", "稳压"}
        name_count = 0
        for v in values:
            vs = str(v).strip() if v else ""
            # Chinese names: 2-3 chars, or / separated
            if re.match(r'^[\u4e00-\u9fff]{2,4}(/[\u4e00-\u9fff]{2,4})*$', vs):
                # Exclude non-name content
                if not any(kw in vs for kw in NON_NAME_KEYWORDS):
                    name_count += 1
        if name_count >= max(1, len(values) * 0.5):
            return "tester_name", 0.7

        # Is it a Jira/defect list? (multi-line with Jira pattern)
        jira_count = sum(1 for v in values if re.search(r'[A-Z]+-\d+', str(v)) if v)
        if jira_count >= max(1, len(values) * 0.5):
            return "jira_summary", 0.85

        # Is it PASS/FAIL/NG test results?
        result_values = {"pass", "fail", "ng", "ok", "通过", "不通过", "未测", "N/A", "n/a", ""}
        result_count = sum(1 for v in values if str(v).strip().lower() in result_values)
        if result_count >= max(1, len(values) * 0.3):
            return "test_result_items", 0.6

        # ---- Rule 2: Sheet-based context ----
        sheet_lower = sheet.lower()

        if any(w in sheet_lower for w in ["oms", "dms", "视觉"]):
            # Test result sheets - likely test case results
            return "test_result_items", 0.5

        if any(w in sheet_lower for w in ["总结", "summary", "结论"]):
            # Only classify text cells (not formulas, not pure numbers)
            for v in values:
                vs = str(v).strip() if v else ""
                # Skip if it looks like a formula result (numbers only)
                if vs and (vs.isdigit() or re.match(r'^\d+\.?\d*$', vs)):
                    continue
                # Skip if it's very short (like "0")
                if len(vs) <= 1:
                    continue
                # This is a text description cell
                return "test_conclusion", 0.5
            return "unknown", 0.2

        if "cover" in sheet_lower or "封面" in sheet:
            return "cover_info", 0.4

        if "test info" in sheet_lower or "测试信息" in sheet:
            # These are the key variable fields
            return self._classify_by_value_content(values)

        return "unknown", 0.0

    def _classify_by_value_content(self, values: List[Any]) -> tuple:
        """基于值的内容推断类型（用于标签-值布局的Sheet）"""
        if not values:
            return "unknown", 0.0

        # Try each version pattern
        for v in values:
            vs = str(v).strip() if v else ""
            for p in self.VERSION_PATTERNS:
                if p.match(vs):
                    return "test_software_version", 0.85

        # Check for Chinese names (exclude equipment terms)
        NON_NAME = {"台架", "实车", "线刷", "模拟", "环境", "设备", "电源", "线缆", "稳压", "车机", "无"}
        for v in values:
            vs = str(v).strip() if v else ""
            if re.match(r'^[\u4e00-\u9fff]{2,4}(/[\u4e00-\u9fff]{2,4})?$', vs):
                if not any(kw in vs for kw in NON_NAME):
                    return "tester_name", 0.65

        # Check for dates
        if any(self.is_datetime(v) for v in values):
            return "test_date", 0.8

        return "unknown", 0.0

    # ================================================================
    # Build schema and mapping
    # ================================================================

    def build_schema(self, analysis: Dict[str, Any]) -> Dict[str, Any]:
        """生成 variable_schema.json"""
        variables = analysis.get("variables", {})
        schema = {}

        for key, info in variables.items():
            field_def = self.FIELD_TYPES.get(key, self.FIELD_TYPES["unknown"])

            # Determine data type
            if "date" in key:
                py_type = "date"
            elif "version" in key:
                py_type = "string"
            else:
                py_type = "string"

            schema[key] = {
                "label": field_def["label"],
                "type": py_type,
                "required": field_def["required"],
                "description": field_def["description"],
                "confidence": round(info.get("confidence", 0), 2),
                "cell_count": len(info.get("cells", [])),
            }

        return schema

    def build_cell_mapping(self, analysis: Dict[str, Any]) -> Dict[str, Any]:
        """生成 cell_mapping.json"""
        variables = analysis.get("variables", {})
        mapping = {}

        for key, info in variables.items():
            mapping[key] = info.get("cells", [])

        return mapping
