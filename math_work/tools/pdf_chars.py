# -*- coding: utf-8 -*-
"""pypdf 逐页字符统计: 识别封面/目录/参考文献/附录, 统计正文页均字符数"""
import sys
from pypdf import PdfReader

def main(path):
    r = PdfReader(path)
    n = len(r.pages)
    print(f"=== {path} ===")
    print(f"总页数: {n}")
    per = []
    for i, pg in enumerate(r.pages):
        try:
            t = pg.extract_text() or ''
        except Exception:
            t = ''
        c = len(t.strip())
        per.append(c)
        tag = ''
        if i == 0:
            tag = ' [封面/首页]'
        if i < n - 1:
            nxt = (r.pages[i + 1].extract_text() or '')
            if '附录' in (t + nxt) and '附录' not in (nxt[:30]):
                pass
        first = t[:24].replace('\n', ' ')
        print(f"p{i+1:2d} {c:5d}  {first}")
    body = [c for i, c in enumerate(per) if c > 0]
    print(f"非空页均: {sum(body)//len(body)}")

if __name__ == '__main__':
    main(sys.argv[1])
