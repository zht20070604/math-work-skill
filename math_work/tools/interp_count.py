# -*- coding: utf-8 -*-
"""统计 docx 中的解读短语数 vs 图/表提及数(图注+表注)
解读短语: 观察图X可知 / 由图X可知 / 由图X可见 / 由表X可知 / 分析表X可知 / 由表X可见
"""
import sys, re
from docx import Document

PATS = [
    r'观察图\s*\d+\s*可知',
    r'由图\s*\d+\s*(?:的\S{0,8})?(?:可知|可见)',
    r'由表\s*\d+\s*(?:的\S{0,8})?(?:可知|可见|可得)',
    r'分析表\s*\d+\s*可知',
    r'观察表\s*\d+\s*可知',
    r'结合图\s*\d+',
]
RE = re.compile('|'.join(PATS))

def main(path):
    doc = Document(path)
    interp = 0
    used = set()  # 被解读过的图/表号
    caps = 0
    cap_labels = []  # (kind, num)
    for p in doc.paragraphs:
        txt = ''.join(r.text for r in p.runs).strip()
        for m in RE.finditer(txt):
            interp += 1
            mm = re.search(r'(图|表)\s*(\d+)', m.group(0))
            if mm:
                used.add((mm.group(1), int(mm.group(2))))
        m = re.match(r'^(图|表)\s*(\d+)\s{2,}', txt)
        if m:
            caps += 1
            cap_labels.append((m.group(1), int(m.group(2)), txt))
    print(f"=== {path} ===")
    print(f"图注+表注数: {caps}")
    print(f"解读短语数: {interp}")
    print(f"被解读覆盖的图/表号: {sorted(used)}")
    # 符号说明表豁免(表1 等非结果表)
    def is_symbol(cap):
        return cap[0] == '表' and '符号' in cap[2]
    missing = [c for c in cap_labels if (c[0], c[1]) not in used and not is_symbol(c)]
    exempt = [c for c in cap_labels if is_symbol(c)]
    result_caps = caps - len(exempt)
    print(f"未覆盖的结果图/表: {missing}")
    print(f"豁免(符号说明表): {[(c[0], c[1]) for c in exempt]}")
    print(f"结果图表数={result_caps}, 解读数={interp}, 判定: {'✓ 解读数≥结果图表数' if interp >= result_caps else '✗ 解读数<结果图表数'}")

if __name__ == '__main__':
    main(sys.argv[1])
