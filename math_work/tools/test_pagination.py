# -*- coding: utf-8 -*-
"""
分页软门回归测试：验证 verify_paper.check_pagination 能识别"几乎空白页"，并正确豁免
参考文献首行页 / 附录起始页 / 末页。
用法： python test_pagination.py（无参数）
  - 用 PyMuPDF(fitz) 直接构造 PDF（不依赖 LibreOffice/Word），只测"检测器"本身，
    文档 docx→pdf 的真实渲染路径已由对成品论文跑 verify_paper 覆盖。
  - 干净样例：三页全部写满 → 断言 0 条分页 WARN。
  - 症状样例：第 2 页只写一行（几乎空白）→ 断言第 2 页被命中（≥1 条）。
  - 豁免样例：几乎空白页但注明「参考文献」/「附录 1」 → 断言未被标记（验证白名单）。
  退出码 0=全 PASS，1=有断言失败。缺 PyMuPDF 则打印原因后退出 0（软跳过）。
"""
import os, sys, re
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import fitz
from verify_paper import check_pagination

W, H = 595, 842  # A4（pt）


def _fill(page):
    """把一页写满：内容底边 ~800pt，正常页。"""
    y = 60
    while y < 790:
        page.insert_text((50, y), "这是一段用于撑满页面的正文内容。", fontname="china-s")
        y += 22


def _mk(path, pages):
    """pages: list of dict。每页按 spec 生成：{} 空页；{'fill':True} 写满；{'line':s} 仅顶部一行。"""
    doc = fitz.open()
    for spec in pages:
        page = doc.new_page(width=W, height=H)
        if spec.get("fill"):
            _fill(page)
        elif spec.get("line"):
            page.insert_text((50, 60), spec["line"], fontname="china-s")
    doc.save(path)
    doc.close()


def main():
    if not fitz:
        print("[SKIP] PyMuPDF 不可用，跳过（退出 0）")
        return 0
    fails = []
    here = os.path.dirname(os.path.abspath(__file__))
    toks = {"fill": {"fill": True}}

    # 1) 干净样例：三页全满 → 0 WARN
    p_clean = os.path.join(here, "_pg_clean.pdf")
    _mk(p_clean, [toks["fill"], toks["fill"], toks["fill"]])
    issues, skip = check_pagination(p_clean)
    if skip:
        fails.append("干净样例返回跳过: %s" % skip)
    elif issues:
        fails.append("干净样例不应有 WARN: %s" % issues)
    else:
        print("[PASS] 干净样例（三页全满）→ 0 条分页 WARN")

    # 2) 症状样例：第 2 页几乎空白 → 应命中第 2 页
    p_sym = os.path.join(here, "_pg_sym.pdf")
    _mk(p_sym, [toks["fill"], {"line": "占位顶行"}, toks["fill"], toks["fill"]])
    issues, skip = check_pagination(p_sym)
    if skip:
        fails.append("症状样例返回跳过: %s" % skip)
    elif not any("第 2 页" in s for s in issues):
        fails.append("症状样例未在第 2 页报 WARN: %s" % issues)
    else:
        print("[PASS] 症状样例（第 2 页空页）→ 命中第 2 页")

    # 3) 豁免样例：几乎空白页但含「参考文献」/「附录 1」 → 不应标记
    p_wl = os.path.join(here, "_pg_wl.pdf")
    _mk(p_wl, [toks["fill"], {"line": "参考文献"}, {"line": "附录 1"}, toks["fill"]])
    issues, skip = check_pagination(p_wl)
    if skip:
        fails.append("豁免样例返回跳过: %s" % skip)
    elif any(re.search(r"第 [23] 页", s) for s in issues):
        fails.append("豁免样例不应标记参考文献/附录页: %s" % issues)
    else:
        print("[PASS] 豁免样例（空页但含参考文献/附录标记）→ 0 条 WARN")

    for f in (p_clean, p_sym, p_wl):
        try:
            os.remove(f)
        except OSError:
            pass

    if fails:
        print("\n结果：FAIL（分页软门回归）")
        for f in fails:
            print("  - " + f)
        return 1
    print("\n结果：全部 PASS（退出码 0）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
