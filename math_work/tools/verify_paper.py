# -*- coding: utf-8 -*-
"""
论文一键结构校验（skill 检查清单的单入口）
用法： python verify_paper.py <论文.docx> [论文.pdf]
  PASS/FAIL 项：
    1) 连排检查    —— 相邻图/表注之间必须有 ≥1 正文段（FAIL，n_p=0 即连排）
    2) 解读短语≥图表 —— 全文"观察图X可知"等解读短语总数 ≥ 结果图/表数（FAIL）
    3) 摘要结构    —— 段数=题量+2（开头+每问一段+结尾）、"针对问题N"编号连续 1..N、
                      无"针对问题 X、Y"合并段（FAIL）；中文字符数 700~800（INFO）
    4) 页均字数     —— 正文页平均字符数（pypdf 文本层，仅密度诊断，不 FAIL）
    5) OMML 结构    —— <m:oMath> 数量、每个 <m:nary> 必有 <m:sup>（按"直接子节点"判，嵌套感知）、
                       剥离公式后无 "\" 残留（FAIL）
    6) 整数图号     —— 图/表编号必须是整数，禁止「第1.0图」「图4.0」小数图号（FAIL）
    7) 引用落点     —— 参考文献每条 [k] 必须在正文至少被引用一次（FAIL；无参考文献节/无条目则跳过）
    8) 截尾术语    —— 截断前正态位置/尺度参数（h0/σz）禁称"截尾分布的均值/标准差"（FAIL）
    9) 可复现性    —— 论文目录脚本无机器绝对路径、omml_builder 随产物捆绑（FAIL；缺第三参/目录不存在则 INFO 跳过）
    10) 文风建议    —— 裸图导语(R1)/解读开句同形(R2)/段尾拔高宣传句(R3)
                      （软门 WARN/INFO，不进 FAIL 不影响退出码；命中按 SKILL 第十节改写后复扫，0 条为佳）
    11) 产物新鲜度  —— 源码 vs docx mtime 提示（软 INFO，不 FAIL；mtime 不可靠，真门禁是
                      SKILL 规则 10 旧句回归核对：旧表述 0 命中 / 新表述已出现）
    12) 分页美观    —— PyMuPDF 逐页内容底边检测疑似空页/整页大留白（软门 WARN，不 FAIL；
                      豁免参考文献首行页/附录起始页/末页；fitz 不可用则 INFO 跳过）
  退出码 0 = 全 PASS，1 = 有 FAIL
配套：docx_seq.py（连排）、interp_count.py（解读短语）、pdf_chars.py（页均字数）、omml_builder.py（公式）
"""
import os, re, sys, zipfile
import xml.etree.ElementTree as ET
sys.stdout.reconfigure(encoding="utf-8")
from docx import Document
import docx_seq
import interp_count

F = "✓"  # ✓
X = "✗"  # ✗


def check_sequence(doc):
    """连排：相邻图注/表注之间无正文段"""
    seq, lines = [], []
    for p in doc.paragraphs:
        c, s = docx_seq.classify(p)
        seq.append(c)
        lines.append((c, s))
    ft = [i for i, c in enumerate(seq) if c in ("F", "T")]
    bad = []
    for k in range(len(ft) - 1):
        a, b = ft[k], ft[k + 1]
        n_p = seq[a + 1:b].count("P")
        if n_p == 0:
            bad.append((lines[a][1][:26], lines[b][1][:26]))
    return bad


def check_interp(doc):
    """解读短语数 vs 结果图/表数（符号说明表豁免）"""
    caps, interp, used, cap_labels = 0, 0, set(), []
    for p in doc.paragraphs:
        txt = "".join(r.text for r in p.runs).strip()
        for m in interp_count.RE.finditer(txt):
            interp += 1
            mm = re.search(r"(图|表)\s*(\d+)", m.group(0))
            if mm:
                used.add((mm.group(1), int(mm.group(2))))
        m = re.match(r"^(图|表)\s*(\d+)\s{2,}", txt)
        if m:
            caps += 1
            cap_labels.append((m.group(1), int(m.group(2)), txt))
    exempt = [c for c in cap_labels if c[0] == "表" and "符号" in c[2]]
    missing = [c for c in cap_labels
               if (c[0], c[1]) not in used and not (c[0] == "表" and "符号" in c[2])]
    return interp, caps - len(exempt), missing


def check_omml(docx_path):
    """OMML 结构：<m:oMath> 数量、每个 <m:nary> 含 <m:sup>、公式段内无 "\" 残留。
    每个 <m:nary> 的 <m:sup> 用"直接子节点"判定（ET 解析逐个 <m:nary> 查其子级有无 sup），
    以避免嵌套 ∑ 时被外层/内部上标误判（2026-09 修复，见 test_nary_sup.py）。
    残留只查"含 oMath 的段落"（解析器没吃完的 \命令 才会留在公式段 w:t 里）；
    附录代码清单等无 oMath 的段落天然排除。"""
    xml = zipfile.ZipFile(docx_path).read("word/document.xml").decode("utf-8")
    n_math = len(re.findall(r"<m:oMath\b", xml))
    n_para = len(re.findall(r"<m:oMathPara\b", xml))
    mns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    no_sup = 0
    try:
        # 嵌套感知：逐个 <m:nary> 查"直接子节点"是否有 <m:sup>。旧版"非贪婪正则 + 段内含 sup"
        # 在嵌套 ∑ 时（内层 ∑ 在外层 <m:e> 里、或 W^{1/2} 的 1/2 上标在 <m:e> 深处）会被外层/
        # 内部上标误判成"有上标"而漏报（真缺 sup 的 ∑ 在 WPS 里渲染虚线占位框）。ET 的 iter()
        # 下探嵌套元素，故内外层 ∑ 分别判定；只查直接子节点则不受 e 内深层上标干扰。
        root = ET.fromstring(xml)
        m_nary, m_sup = "{%s}nary" % mns, "{%s}sup" % mns
        for el in root.iter(m_nary):
            if not any(ch.tag == m_sup for ch in el):
                no_sup += 1
    except Exception:
        narys = re.findall(r"<m:nary\b[^>]*>.*?</m:nary>", xml, re.S)
        no_sup = len([n for n in narys if "<m:sup>" not in n])
    residue = []
    for para in re.findall(r"<w:p[ >].*?</w:p>", xml, re.S):
        if "<m:oMath" not in para:
            continue
        stripped = re.sub(r"<m:oMath>.*?</m:oMath>", "", para, flags=re.S)
        for t in re.findall(r"<w:t[^>]*>([^<]*)</w:t>", stripped):
            if "\\" in t:
                residue.append(t)
    return n_math, n_para, no_sup, residue


def check_abstract(doc):
    """摘要结构（2026-09 增补）：段数 = 题量 + 2（开头 + 每问各一段 + 结尾）；
    每个"针对问题 N"段编号连续 1..N；无"针对问题 X、Y"合并段。
    中文字符数 700~800 为 INFO（同密度诊断，只提示不 FAIL）。"""
    paras = []
    in_abs = False
    for p in doc.paragraphs:
        t = "".join(r.text for r in p.runs).strip()
        if re.match(r"^摘\s*要", t):
            in_abs = True
            continue
        if t.startswith("关键词"):
            break
        if in_abs and t:
            paras.append(t)
    if not paras:
        return None, ["未找到摘要段（缺少「摘要」标题或「关键词」段）"]
    qnums = [int(m.group(1)) for p in paras
             if (m := re.match(r"针对问题\s*(\d+)", p))]
    issues = []
    n = len(qnums)
    if n < 1:
        issues.append("没有任何以「针对问题 N」开头的段")
    if len(paras) != n + 2:
        issues.append("摘要段数 %d ≠ 题量(%d)+2(开头+结尾)" % (len(paras), n))
    if qnums != list(range(1, n + 1)):
        issues.append("针对问题编号非连续 1..N: %s" % qnums)
    for i, p in enumerate(paras):
        if re.search(r"针对问题\s*\d+\s*[、,，]\s*\d+", p):
            issues.append("第 %d 段为合并段（针对问题 X、Y）: %s" % (i + 1, p[:24]))
    cn = len(re.findall(r"[一-鿿]", "".join(paras)))
    return cn, issues


def check_density(pdf_path):
    """正文区页均字数（pypdf 文本层，密度诊断）。
    只统计"首个附录标题行（^附录\\s*\\d）之前"的页：附录代码清单/支撑材料页刻意稀疏，
    混入会让公式密集的正文误报"过简"。OMML 公式不在 pypdf 文本层内，公式页提取字数天然偏低，
    因此该诊断仅供人工参考，正文以内容要素门为准。"""
    from pypdf import PdfReader
    r = PdfReader(pdf_path)
    per = []
    for pg in r.pages:
        t = (pg.extract_text() or "").strip()
        per.append(len(t))
    # 找正文区终止页：首个以"附录"+数字开头的行（正文里的"见附录1"非行首，不误判）
    cutoff = len(r.pages)
    for i, t in enumerate((pg.extract_text() or "") for pg in r.pages):
        if any(re.match(r"^附录\s*\d", ln) for ln in t.splitlines()):
            cutoff = i
            break
    body = [c for c in per[:cutoff] if c > 0]
    avg = sum(body) // len(body) if body else 0
    n = len(r.pages)
    first = per[0] if per else 0
    last = per[-1] if per else 0
    return n, avg, first, last, cutoff


# ============ 分页软门（疑似空页/整页大留白；WARN/INFO，不 FAIL） ============
PAG_EMPTY_RATIO = 0.30  # 内容底边 < 页高该比例 → 判"整页大留白/疑似空页"（调此处即可松紧）


def check_pagination(pdf_path):
    """分页美观【软门，不 FAIL】——检测"几乎空白页"：内容底边 y 远低于页底。
    根因：表格单元格 keep_with_next 整表链式推挤、或正文手插 add_break(PAGE) 串出空白页。
    豁免：参考文献首行页、附录起始页、末页（章节/清单收尾页天然稀疏）。
    依赖 PyMuPDF(fitz)：不可用则返回 (None, [原因]) 软跳过，绝不影响退出码。"""
    try:
        import fitz
    except Exception:
        return None, ["PyMuPDF(fitz) 未安装，跳过分页检查（软门，不 FAIL）"]
    doc = fitz.open(pdf_path)
    pages = []
    for i, page in enumerate(doc):
        h = page.rect.height
        yb = 0.0
        for b in page.get_text("blocks"):
            yb = max(yb, b[3])
        for b in page.get_image_info():
            try:
                yb = max(yb, b["bbox"][3])
            except Exception:
                continue
        for d in page.get_drawings():
            try:
                yb = max(yb, d["rect"].y1)
            except Exception:
                continue
        pages.append({"i": i, "h": h, "yb": yb,
                      "txt": page.get_text("text") or ""})
    n = len(pages)
    first_ref = next((pg["i"] for pg in pages
                      if any(re.match(r"^\s*参考文献", ln) for ln in pg["txt"].splitlines())), None)
    first_app = next((pg["i"] for pg in pages
                      if any(re.match(r"^\s*附录\s*\d", ln) for ln in pg["txt"].splitlines())), None)
    issues = []
    for pg in pages:
        if pg["i"] == n - 1:          # 末页豁免
            continue
        if first_ref is not None and pg["i"] == first_ref:
            continue
        if first_app is not None and pg["i"] == first_app:
            continue
        if pg["h"] <= 0 or pg["yb"] >= pg["h"] * PAG_EMPTY_RATIO:
            continue
        issues.append("第 %d 页：内容底边 %.0fpt 仅占页高 %.0fpt 的 %.0f%%（>%d%% 留白）→ 疑似空页/整页大留白"
                      % (pg["i"] + 1, pg["yb"], pg["h"],
                         100 * pg["yb"] / pg["h"], int(100 * (1 - PAG_EMPTY_RATIO))))
    doc.close()
    return issues, []


def check_float_figref(doc):
    """图/表编号必须为整数：禁止「第1.0图」「图4.0」「表10.0」这类小数图号（2026-09-03 审计发现）。
    命中即 FAIL。扫描正文段 + 所有表格单元格。"""
    # 两种顺序都抓：图/表在数字前（「图4.0」）或数字在词语前（「第1.0图」）
    pat = re.compile(r"(?:第\s*\d+\.\d+\s*(?:图|表))|((?:图|表)\s*\d+\.\d+)")
    hits = []
    def _scan(text):
        for m in pat.finditer(text):
            s = max(0, m.start() - 8); e = min(len(text), m.end() + 4)
            hits.append(text[s:e])
    for p in doc.paragraphs:
        _scan("".join(r.text for r in p.runs))
    for tb in doc.tables:
        for r in tb.rows:
            for c in r.cells:
                _scan(c.text)
    return hits


def check_citation(doc):
    """正文引用落点：参考文献每条 [k] 必须在正文（关键词段之后、参考文献节之前）至少被引用一次。
    识别不到「参考文献」节，或该节无 [k] 条目时跳过（返回 (None, [原因])），不 FAIL。
    命中（正文章节没有任何 [k] 落点）即 FAIL。"""
    paras = doc.paragraphs
    ref_start = None
    for i, p in enumerate(paras):
        t = "".join(r.text for r in p.runs).strip()
        if re.match(r"^\s*\d*[\.、]?\s*参考文献", t) or t.replace(" ", "") == "参考文献":
            ref_start = i; break
    if ref_start is None:
        return None, ["未识别到「参考文献」节，跳过引用落点检查"]
    refs = set()
    for p in paras[ref_start:]:
        for m in re.finditer(r"\[\s*(\d+)\s*\]", "".join(r.text for r in p.runs)):
            refs.add(int(m.group(1)))
    if not refs:
        return None, ["参考文献节无 [k] 条目，跳过引用落点检查"]
    body = "\n".join("".join(r.text for r in p.runs) for p in paras[:ref_start])
    body_refs = set(int(x) for x in re.findall(r"\[\s*(\d+)\s*\]", body))
    missing = sorted(refs - body_refs)
    return refs, missing


def check_terminology(doc):
    """截尾分布术语核对（2026-09-03 审计发现）：截断前正态的 h0/σz 是位置/尺度参数，
    不得称为"截尾分布的均值/标准差/方差"。扫正文段 + 表格单元格。命中（错误说法）即 FAIL。"""
    pat = re.compile(r"截尾[^，。；]{0,8}(均值|标准差|方差|期望)")
    bad = []
    def _scan(text):
        for m in pat.finditer(text):
            s = max(0, m.start() - 12); e = min(len(text), m.end() + 6)
            bad.append(text[s:e])
    for p in doc.paragraphs:
        _scan("".join(r.text for r in p.runs))
    for tb in doc.tables:
        for r in tb.rows:
            for c in r.cells:
                _scan(c.text)
    return bad


def check_repro(paper_dir):
    """可复现性：论文目录下 .py 脚本不得含机器专属绝对路径；omml_builder 必须随产物捆绑。
    扫描 paper_dir 里所有 .py（含 BUILD 脚本），命中以下任一条即返回违规清单：
      - 硬编码本机绝对路径（C:\\Users\\、C:/Users/、D:\\tool、D:/tool、/home/、/Users/ 等）
      - sys.path.insert 指向技能仓库 tools 目录（越层 import omml_builder → 换机 ModuleNotFoundError）
    未提供目录 / 目录不存在 → 返回 ([], 原因) 跳过，不 FAIL。"""
    if not paper_dir or not os.path.isdir(paper_dir):
        return [], ["未提供有效论文目录，跳过可复现性检查"]
    bad = []
    abs_pat = re.compile(r"([A-Z]:[\\/]|/[Uu]sers/|/home/)")
    skill_tools = re.compile(r"math_work[\\/]tools|search[\\/]math")
    for fn in sorted(os.listdir(paper_dir)):
        if not fn.lower().endswith(".py"):
            continue
        fp = os.path.join(paper_dir, fn)
        try:
            txt = open(fp, encoding="utf-8", errors="ignore").read()
        except Exception:
            continue
        for i, ln in enumerate(txt.splitlines(), 1):
            if "sys.path.insert" in ln and skill_tools.search(ln):
                bad.append("%s:%d: sys.path.insert 指向技能仓库 tools → omml_builder 未随产物捆绑" % (fn, i))
        m = abs_pat.search(txt)
        if m:
            bad.append("%s: 含机器绝对路径 %s → 换机无法运行" % (fn, m.group(0)))
    return bad, []


def check_freshness(docx_path, paper_dir):
    """产物新鲜度【软 INFO，不 FAIL】——提示"这份 docx 是否由当前源码重建"。
    ⚠️ mtime 不可靠：2026-09 实测 docx 01:27 > 源码 01:24-01:26 却内容更旧，
    mtime 判断不出那次漂移。故本函数不判 PASS/FAIL，只在"源码比产物新"或
    "无法确认"时打一句提醒，真正的门禁是 SKILL 规则 10 的旧句回归核对
    （旧表述在重建后 docx 0 命中 / 新表述已出现）。"""
    newest_src = None
    src_pat = re.compile(r"(?:build_paper_.*\.py|make_figs\.py|solve\.py|paper_results\.py|omml_builder\.py)$")
    if paper_dir and os.path.isdir(paper_dir):
        for fn in os.listdir(paper_dir):
            if src_pat.search(fn):
                p = os.path.join(paper_dir, fn)
                try:
                    mt = os.path.getmtime(p)
                except OSError:
                    continue
                if newest_src is None or mt > newest_src[1]:
                    newest_src = (fn, mt)
    if newest_src is None:
        return ["未找到生成/求解源码（build_paper_*/make_figs/solve/paper_results/omml_builder），无法核对产物新鲜度"]
    try:
        doc_mt = os.path.getmtime(docx_path)
    except OSError:
        return []
    if newest_src[1] > doc_mt:
        return ["源码 %s 比产物新 → 可能忘了重建，先重跑整链再校验" % newest_src[0]]
    return ["mtime 显示产物已是最新，但 mtime 不可靠（2026-09 实例 docx 比源码新却内容旧）；仍请按规则 10 用旧句回归核对确认产物确由当前源码重建"]


# ================= 文风软门（降 AI 味；WARN/INFO，不 FAIL） =================
# 首批只收高置信、低误报的模板痕迹（先保守后扩）。阈值/名单改此处常量即可，不动逻辑。
R1_BARE_FIG_PAT = re.compile(r"如下图所示|如上图所示|如图所示|见下图|见上图")
R2_EMPH_MIN = 3           # 强调式开句（观察图/分析表/观察表/结合图）同形态≥N 个段首才提示
R2_EMPH_RATIO = 0.4       # 且占段首锚点解读段比例 > 该值（机械排比）；中性"由图/由表"不计
R3_TAIL_PATS = [          # 段尾拔高宣传句（只在段末句查，避免误报正文中部合法权衡句）
    re.compile(r"体现(?:了)?[^。]{0,24}的(?:合理性|有效性|实用性|科学性|可行性|适用性)"),
    re.compile(r"(?:具备|具有|展现|表现出)[^。]{0,16}(?:较强的|良好的|优异的)[^。]{0,12}(?:能力|鲁棒性|迁移性|扩展性|适应性|泛化性)"),
    re.compile(r"兼顾(?:了)?(?:安全|可行|高效|简洁|准确|实用|稳定|经济)性"),
    re.compile(r"构成(?:了)?(?:完整|统一|严密|系统的?)[^。]{0,10}(?:链条|体系|框架|闭环)"),
    re.compile(r"(?:充分|有力的?)[^。]{0,4}(?:验证|体现|说明|证明)了[^。]{0,16}(?:合理性|可行性|有效性|正确性|实用性)"),
    re.compile(r"显著[^。]{0,10}(?:提升|提高|改善|优于|减少|下降|加快)"),
]


def _ctx(text, m, radius=16):
    s = max(0, m.start() - radius)
    e = min(len(text), m.end() + radius)
    return ("…" if s else "") + text[s:e] + ("…" if e < len(text) else "")


def _style_scan(paras):
    """文风软门扫描。paras：正文段文本列表 → [(level, tag, msg)]。
    解读锚点口径与 check_interp 一致；OMML 内文不在 p.runs 里，天然不扫公式。"""
    issues = []

    # R1 裸图导语：无图号的图引用（正例"如图 3 所示"中间有数字，不命中）
    for text in paras:
        m = R1_BARE_FIG_PAT.search(text)
        if m:
            issues.append(("WARN", "R1 裸图导语",
                           _ctx(text, m) + " → 应写具体图号并配解读锚点（如“由图 3 可见…”）"))
            continue  # 每段报一处

    # R3 段尾拔高宣传句：取段末"最后一个句号所在的整句"再查（不含尾部标点前的上一句）
    for text in paras:
        idx = text.rfind("。")
        if idx < 0:
            tail = text
        else:
            prev = text.rfind("。", 0, idx)
            tail = text[prev + 1:]
        for pat in R3_TAIL_PATS:
            m = pat.search(tail)
            if m:
                issues.append(("WARN", "R3 段尾拔高",
                               _ctx(text, m) + " → 删或换成可核验的机制/数字限定（SKILL 第十节 T2）"))
                break

    # R2 解读段首开句同形：强调式开句（观察图/分析表/观察表/结合图）机械堆叠才提示。
    # 中性"由图/由表X可知"是正常技术语体（模型论文常用），不计入，避免误报。
    anchored = re.compile(r"^\s*(?:" + "|".join(interp_count.PATS) + r")")
    emph = [("观察图", r"观察图"), ("分析表", r"分析表"), ("观察表", r"观察表"), ("结合图", r"结合图")]
    counts, n_start = {}, 0
    for text in paras:
        if anchored.match(text):
            n_start += 1
            for bk, p in emph:
                if re.match(p, text):
                    counts[bk] = counts.get(bk, 0) + 1
                    break
    if n_start >= 6:
        bk, mx = max(counts.items(), key=lambda kv: kv[1]) if counts else ("", 0)
        if mx >= R2_EMPH_MIN and mx > n_start * R2_EMPH_RATIO:
            issues.append(("INFO", "R2 开句同形",
                           "段首锚点解读段 %d 段中，「%s」开句占 %d 段，强调式开句机械堆叠"
                           " → 轮换用「由图 X 可知/由表 X 可知/由图 X 可见」并把部分锚点移出段首"
                           "（SKILL 第十节 T5）" % (n_start, bk, mx)))
    return issues


def check_style(doc):
    """复用 docx_seq.classify，只扫正文叙述段（P）；图注/表注/标题/公式/空段不扫。"""
    paras = []
    for p in doc.paragraphs:
        c, _ = docx_seq.classify(p)
        if c == "P":
            txt = "".join(r.text for r in p.runs).strip()
            if txt:
                paras.append(txt)
    return _style_scan(paras)


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        sys.exit(1)
    docx_path = args[0]
    pdf_path = args[1] if len(args) > 1 else None
    # 第三参：论文程序目录（可复现性扫描）；缺省取 docx 同目录
    paper_dir = args[2] if len(args) > 2 else os.path.dirname(os.path.abspath(docx_path))
    doc = Document(docx_path)

    fails = []
    print("=== 论文结构校验：%s ===" % os.path.basename(docx_path))

    # 先算 OMML（密度诊断要引用公式数判断是否低估）
    n_math, n_para, n_no_sup, residue = check_omml(docx_path)

    # 1) 连排
    bad = check_sequence(doc)
    ok = len(bad) == 0
    print(f"[{'PASS' if ok else 'FAIL'}] 连排检查：相邻图/表注间无正文段 %d 处" % len(bad))
    if bad:
        fails.append("连排")
        for a, b in bad:
            print("       '%s' -> '%s'" % (a, b))

    # 2) 解读 ≥ 图表
    interp, result_caps, missing = check_interp(doc)
    ok = interp >= result_caps
    print(f"[{'PASS' if ok else 'FAIL'}] 解读短语 %d ≥ 结果图/表 %d" % (interp, result_caps))
    if not ok:
        fails.append("解读短语")
        print("       未覆盖: %s" % missing)

    # 3) 摘要结构（段数 = 题量+2 / 编号连续 1..N / 无合并段；字数 INFO）
    abs_cn, abs_issues = check_abstract(doc)
    ok = not abs_issues
    print(f"[{'PASS' if ok else 'FAIL'}] 摘要结构：段数=题量+2、编号连续、无「针对问题X、Y」合并段")
    if abs_issues:
        fails.append("摘要结构")
        for s in abs_issues:
            print("       " + s)
    if abs_cn is not None:
        hint = "" if 700 <= abs_cn <= 800 else "  ⚠ 目标 700~800 字（仅供精简参考，不 FAIL）"
        print(f"[ INFO] 摘要中文字符数 ≈ {abs_cn}{hint}")

    # 4) 页均字数（密度诊断，不 FAIL）
    if pdf_path:
        n, avg, first, last, cutoff = check_density(pdf_path)
        hint = ""
        if avg < 1800:
            if n_math >= 50:
                # 公式密集论文：OMML 不在 pypdf 文本层，密度必然低估，不误报
                hint = (f"  公式密集（<m:oMath> {n_math}）不在 pypdf 文本层内，密度低估属正常，"
                        "以内容要素门为准")
            else:
                hint = "  ⚠ 可能过简，检查是否有图无解读（密度诊断仅供参考，以内容要素门为准）"
        print(f"[ INFO] 正文区页均字数（pypdf 文本层，附录前 {cutoff} 页）≈ {avg} 字/页（总 {n} 页，首页 {first}，末页 {last}）{hint}")
    else:
        print("[ INFO] 未提供 PDF，跳过页均字数诊断（仅供密度参考，不影响判定）")

    # 4b) 分页美观（软门 WARN/INFO：疑似空页/整页大留白；不进 fails，不影响退出码）
    if pdf_path:
        pag_issues, pag_skip = check_pagination(pdf_path)
        if pag_skip:
            print("[ INFO] 分页：%s" % pag_skip[0])
        elif pag_issues:
            print("\n[ WARN] 分页：疑似空页/整页大留白 %d 处（软门，不 FAIL；多为表格单元格 "
                  "keep_with_next 链式推挤或正文手插分页符，见 SKILL 三 表格/分页禁则）" % len(pag_issues))
            for s in pag_issues:
                print("       " + s)
        else:
            print("[ PASS] 分页软门：无空页/异常大留白")
    else:
        print("[ INFO] 未提供 PDF，跳过分页检查（软门，不 FAIL）")

    # 5) OMML 结构
    ok = n_no_sup == 0 and not residue
    print(f"[{'PASS' if ok else 'FAIL'}] OMML：<m:oMath> {n_math}（块级 {n_para}），<m:nary> 缺 <m:sup> {n_no_sup} 个，公式外 '\\' 残留 {len(residue)} 处")
    if n_no_sup:
        fails.append("nary缺sup")
        print("       有 <m:nary> 缺上标（∑ 会渲染虚线框）：见检查清单 20f")
    if residue:
        fails.append("反斜杠残留")
        for t in residue[:5]:
            print("       w:t 含 '\\': %r" % t[:60])

    # 6) 图/表编号整数（小数图号「第1.0图」）
    figref = check_float_figref(doc)
    ok = len(figref) == 0
    print(f"[{'PASS' if ok else 'FAIL'}] 图/表编号整数：小数图号（如「第1.0图」「图4.0」）{len(figref)} 处")
    if figref:
        fails.append("图号小数")
        for h in figref[:5]:
            print("       " + repr(h))

    # 7) 正文引用落点（参考文献 [k] 须在正文被引用）
    refs, cit_missing = check_citation(doc)
    if refs is None:
        print(f"[ INFO] 引用落点：{cit_missing[0]}")
    else:
        ok = len(cit_missing) == 0
        print(f"[{'PASS' if ok else 'FAIL'}] 引用落点：参考文献 [{', '.join(map(str, sorted(refs)))}] 必须全部在正文被引用；正文未引用 {len(cit_missing)} 条")
        if cit_missing:
            fails.append("引用落点")
            print("       正文未出现: [%s]" % ", ".join(map(str, cit_missing)))

    # 7b) 截尾分布术语（h0/σz 是截断前位置/尺度参数，禁称"截尾分布的均值/标准差"）
    term = check_terminology(doc)
    ok = len(term) == 0
    print(f"[{'PASS' if ok else 'FAIL'}] 截尾术语：h0/σz 称「截断前位置/尺度参数」（不称「截尾分布的均值/标准差」）→ 错误 {len(term)} 处")
    if term:
        fails.append("截尾术语")
        for t in term[:5]:
            print("       " + repr(t))

    # 8) 可复现性（脚本含机器绝对路径 / omml_builder 未捆绑）
    repro, repro_skip = check_repro(paper_dir)
    if repro_skip:
        print(f"[ INFO] 可复现性：{repro_skip[0]}")
    else:
        ok = len(repro) == 0
        print(f"[{'PASS' if ok else 'FAIL'}] 可复现性：脚本无机器绝对路径、omml_builder 随产物捆绑 → 违规 {len(repro)} 处")
        if repro:
            fails.append("可复现性")
            for r in repro[:6]:
                print("       " + r)

    # 8b) 产物新鲜度（软 INFO：不进 fails，不影响退出码；真正的门禁是 SKILL 规则 10 旧句回归核对）
    for fres in check_freshness(docx_path, paper_dir):
        print("[ INFO] 新鲜度：%s" % fres)

    # 9) 文风建议（软门 WARN/INFO：不进 fails，不影响退出码）
    style_issues = check_style(doc)
    if style_issues:
        print("\n=== 文风建议（软门，不 FAIL；命中按 SKILL 第十节改写后复扫，0 条为佳） ===")
        for level, tag, msg in style_issues:
            print("[%s] %s: %s" % (level, tag, msg))
        print("共 %d 条建议（供人工/作者判断是否采纳，不影响本工具退出码）" % len(style_issues))
    else:
        print("\n[ PASS] 文风软门：0 条建议（裸图导语/解读开句同形/段尾拔高均未命中）")

    if fails:
        print("\n结果：FAIL（%s）" % ", ".join(fails))
        sys.exit(1)
    print("\n结果：全部 PASS（退出码 0）")
    sys.exit(0)


if __name__ == "__main__":
    main()
