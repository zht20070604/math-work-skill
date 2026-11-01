# -*- coding: utf-8 -*-
"""
nary缺sup 检测回归测试：验证 verify_paper.check_omml 用"直接子节点"判定，能识破嵌套 ∑ 盲区。
用法： python test_nary_sup.py（无参数）
背景：旧版用"非贪婪正则 <m:nary>.*?</m:nary> + 段内含 <m:sup>"，在"内层 ∑ 套在外层 <m:e> 里、
      而外层自身有 <m:sup>（或 e 深处 W^{1/2} 的上标）"时，会把"段内含 sup"当作"该 ∑ 有上标"而漏报，
      真缺 sup 的 ∑ 在 WPS 里渲染成虚线占位框（见 memory wps-nary-missing-sup）。
本测试用 ElementTree 逐个 <m:nary> 查"直接子节点"是否含 <m:sup>：
  - 正宗嵌套（内外层都有 sup）→ 断言 0 条；
  - 盲区样例（内层无 sup、外围有 sup K 与深层 W^{1/2}）→ 断言新检测报 ≥1（并复现旧正则漏报为 0）。
退出码 0=全 PASS，1=有断言失败。缺 python-docx 则打印原因后退出 0（软跳过）。
"""
import os, re, sys, zipfile
import xml.etree.ElementTree as ET
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.oxml import parse_xml
from verify_paper import check_omml

MNS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
# 一个最小可解析的 m:oMath 片段（xmlns:m 就地声明，python-docx 序列化时并入 document.xml）
_NS = 'xmlns:m="%s"' % MNS


def _good_nested_text():
    """J=Σ_{k=1}^{K} Σ_{b=l_k}^{r_k} x —— 内外层都有上标。"""
    return (
        '<m:oMath %s>'
        '<m:nary><m:naryPr/><m:sub>k=1</m:sub><m:sup>K</m:sup>'
        '<m:e><m:nary><m:naryPr/><m:sub>b=l_k</m:sub><m:sup>r_k</m:sup><m:e>x</m:e></m:nary></m:e>'
        '</m:nary></m:oMath>' % _NS)


def _blind_spots_text():
    """真缺 sup：内层 Σ_{b=l_k}（无上标），外圈有 <m:sup>K</m:sup> 且 <m:e> 内含 W^{1/2} 上标。
    旧正则从开标截到内层关，段内"找到 sup K"，误判内层有上标 → 漏报；新检测按直接子节点 → 命中。"""
    return (
        '<m:oMath %s>'
        '<m:nary><m:naryPr/><m:sub>k=1</m:sub><m:sup>K</m:sup>'
        '<m:e><m:nary><m:naryPr/><m:sub>b=l_k</m:sub>'
        '<m:e>W<m:sup>1/2</m:sup></m:e></m:nary></m:e>'
        '</m:nary></m:oMath>' % _NS)


def _mk(path, omml):
    doc = Document()
    p = doc.add_paragraph()
    p._p.append(parse_xml(omml))
    doc.save(path)


def _old_regex_count(path):
    """复现旧版漏报逻辑，用于证明盲区真实存在（应为 0）。"""
    xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8")
    narys = re.findall(r"<m:nary\b[^>]*>.*?</m:nary>", xml, re.S)
    return len([n for n in narys if "<m:sup>" not in n])


def main():
    if not ET or not Document:
        print("[SKIP] 缺 python-docx / ElementTree，跳过（退出 0）")
        return 0
    fails = []
    here = os.path.dirname(os.path.abspath(__file__))

    # 1) 正宗嵌套：内外层都有 sup → 新检测 0 条
    p_good = os.path.join(here, "_nary_good.docx")
    _mk(p_good, _good_nested_text())
    n_math, n_para, no_sup, res = check_omml(p_good)
    if no_sup != 0:
        fails.append("正宗嵌套不应报缺 sup，实报 %s: %s" % (no_sup, res))
    else:
        print("[PASS] 正宗嵌套（内外层均有上标）→ 新检测 0 条")

    # 2) 盲区样例：内层无 sup、外圈有 K 与深层 W^{1/2} → 新检测应命中外层 1 条
    p_bad = os.path.join(here, "_nary_bad.docx")
    _mk(p_bad, _blind_spots_text())
    n_math, n_para, no_sup, res = check_omml(p_bad)
    if no_sup < 1:
        fails.append("盲区样例应报缺 sup ≥1，实报 %s" % no_sup)
    else:
        print("[PASS] 盲区样例（内层无上标）→ 新检测报 %s 条" % no_sup)

    # 3) 证明旧正则确实漏报（对比）
    old = _old_regex_count(p_bad)
    if old == 0:
        print("       （旧正则对盲区样例漏报为 0 —— 证明这是真实修复，非多余）")
    else:
        fails.append("旧正则对盲区样例应漏报为 0（以确认盲区真实），实为 %s" % old)

    for f in (p_good, p_bad):
        try:
            os.remove(f)
        except OSError:
            pass

    if fails:
        print("\n结果：FAIL（nary缺sup 回归）")
        for f in fails:
            print("  - " + f)
        return 1
    print("\n结果：全部 PASS（退出码 0）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
