#!/usr/bin/env python3
"""
测试报告智能生成 —— 一站式脚本

自动检测模板中的旧值，执行单元格替换 + 全局文本替换。
"""

import json, re, zipfile, sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from src.report_generator import ReportGenerator

REMINDER = """
============================================================
  📋 测试报告生成 —— 请提供本次版本信息
============================================================
  1. 版本编号（如 ER9.22.11）：
  2. 编制（测试人姓名）：
  3. 审核（审核人 / DRE 姓名）：
  4. 日期（如 2088-01-01）：
  5. 测试开始日期：
  6. 测试结束日期：
============================================================
"""

def extract_old_values():
    """从模板中提取旧值"""
    ss = zipfile.ZipFile('project_config/template.xlsx').read('xl/sharedStrings.xml').decode()
    s2 = zipfile.ZipFile('project_config/template.xlsx').read('xl/worksheets/sheet2.xml').decode()

    def get_ss(idx):
        matches = list(re.finditer(r'<si>', ss))
        if idx >= len(matches): return ''
        start = matches[idx].end()
        end = ss.find('</si>', start)
        texts = re.findall(r'<t[^>]*>(.*?)</t>', ss[start:end], re.DOTALL)
        from xml.sax.saxutils import unescape
        return unescape(''.join(texts))

    old = {}
    for ref, key in [('C12', 'version'), ('C14', 'tester'), ('C9', 'dre'), ('C17', 'start_date'), ('C18', 'end_date')]:
        m = re.search(r'<c\s+r="' + ref + r'"[^>]*><v>(\d+)</v>', s2)
        if m:
            val = get_ss(int(m.group(1))).strip()
            old[key] = val
    
    return old


if __name__ == "__main__":
    print(REMINDER)
    old = extract_old_values()
    print(f"  ℹ️  模板旧值参考: 版本={old.get('version','?')}, 测试人={old.get('tester','?')}, DRE={old.get('dre','?')}")
    print("============================================================\n")
