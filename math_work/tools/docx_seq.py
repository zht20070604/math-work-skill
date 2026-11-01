# -*- coding: utf-8 -*-
"""读 docx 段落序列: 区分 图注F / 表注T / 公式M / 正文P / 标题H / 空E
输出每段的类型+文本摘要, 以及"图注之间无正文"的连排统计
"""
import sys
import re
from docx import Document
from docx.oxml.ns import qn

def classify(p):
    txt = ''.join(r.text for r in p.runs).strip()
    has_math = p._p.findall('.//' + qn('m:oMath')) or p._p.findall('.//' + qn('m:oMathPara'))
    if has_math:
        # 含内联公式且有叙述文字 = 正文叙述; 纯公式段 = M
        if txt and not re.match(r'^[\s(]*[A-Za-z\\$=]', txt):
            return 'P', txt[:40]
        return 'M', txt[:40] or '<formula>'
    if not txt:
        return 'E', ''
    # 图注/表注: "图N  标题"(数字后≥2空格)或"图N-1"; 解读句"图 X 以…/图 X 展示…"是正文
    if re.match(r'^图\s*\d+\s{2,}', txt) or re.match(r'^图\s*[A-Z]\s*\d', txt):
        return 'F', txt[:50]
    if re.match(r'^表\s*\d+\s{2,}', txt):
        return 'T', txt[:50]
    if re.match(r'^附录', txt) or re.match(r'^\d+\.\d', txt) or re.match(r'^\d+\s', txt):
        return 'H', txt[:40]
    if re.match(r'^(摘要|关键词|目录|参考文献)', txt):
        return 'H', txt[:40]
    return 'P', txt[:60]

def main(path):
    doc = Document(path)
    seq = []
    lines = []
    for p in doc.paragraphs:
        c, s = classify(p)
        seq.append(c)
        lines.append((c, s))
    full = ''.join(seq)
    print(f"=== {path} ===")
    print(f"段落总数: {len(seq)}")
    # 图/表注连排统计: 相邻图注/表注之间无 P(正文) 即连排
    ft = [i for i, c in enumerate(seq) if c in ('F', 'T')]
    print(f"图注+表注数: {len(ft)}")
    gaps = []
    for k in range(len(ft) - 1):
        a, b = ft[k], ft[k + 1]
        between = seq[a + 1:b]
        has_p = 'P' in between
        n_p = between.count('P')
        gaps.append((a, b, n_p, has_p))
    print("相邻图/表注间正文段数(n_p=0 表示连排/无正文):")
    for a, b, n_p, has_p in gaps:
        fa, fb = lines[a][1], lines[b][1]
        flag = "  ⚠️无正文" if n_p == 0 else ""
        print(f"  '{fa}' → '{fb}': 中间正文{n_p}段{flag}")
    return seq, lines

if __name__ == '__main__':
    main(sys.argv[1])
