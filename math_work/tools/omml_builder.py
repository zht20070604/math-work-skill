# -*- coding: utf-8 -*-
r"""
OMML 公式构造器 —— 把公式字符串直接构建为 Word 原生公式 XML（OMML）。
不依赖 Word/WPS COM，python-docx 直接写入，Word/WPS/LibreOffice 均可原生渲染。

这是 math_work 技能的配套公式工具：任何使用该技能生成论文的 AI，
一律调用本文件把公式转成 OMML，不要再走 Word/WPS COM BuildUp（本机 WPS 下不生效，见 SKILL.md）。

【导入方式（可复现性强制，2026-09-10 增补）】论文生成脚本引用本文件时，必须先把
omml_builder.py 复制到论文程序目录，再用 sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
本地导入。禁止在脚本里写技能仓库路径或本机绝对路径——那会同时触发 verify_paper.py 的
可复现性门两条判据（机器绝对路径 + 越层 sys.path.insert），且换机后 ModuleNotFoundError。

==================================== 支持语法 ====================================

【A. 简单写法（线性规划等报告一直在用，向后兼容）】
  - 下标：x_{11}、c_{ij}、p_{i}（花括号必须）
  - 上标：x^{2}、m^{3}
  - 求和：∑_{i=1}^{4}（操作数为后续到顶层运算符或结束的表达式）
  - 括号分组：( ... ) → 自动括号；{ ... } → 集合花括号；[ ... ] → 方括号
  - 运算符：+ - = ≤ ≥ ∈ , 以及中文逗号、·、× 等字符直接输出
  - 空格：两个符号间产生一个可见空格（如 max z、c_{ij} x_{ij}）
  - 中文/字母直接输出（正体）

【B. LaTeX 写法（复杂公式，支持命令 + 花括号参数）】
  分数    \frac{a}{b}、\dfrac、\tfrac、\binom{n}{k}（组合数，隐藏分数线）
  根号    \sqrt{x}、\sqrt[3]{x}、\sqrt x（无花括号参数也可）
  累加累乘 \sum_{i=1}^{n}、\prod、\int_{0}^{1}、\iint、\iiint、\oint、\oiint(∯)、
         \iiiint、\idotsint（无跨引擎独立字形，用 ∫∫∫∫/∫∫∫ 文本近似）、\bigcap、\bigcup、
         \sum_{\substack{i<j\\j<k}}（\substack 多行上下限，行间用 \\）
  极限    \lim_{x \to 0}（\limits / \nolimits 自动跳过）
  矩阵    \begin{matrix}…\end{matrix}
         \begin{pmatrix}…\end{pmatrix}   → 圆括号
         \begin{bmatrix}…\end{bmatrix}   → 方括号
         \begin{Bmatrix}…\end{Bmatrix}   → 花括号
         \begin{vmatrix}…\end{vmatrix}   → 竖线
         \begin{Vmatrix}…\end{Vmatrix}   → 双竖线
         \begin{array}{ccc}…\end{array}  → 普通矩阵（列格式 {ccc}、位置 [t] 自动跳过）
         \begin{smallmatrix}…\end{smallmatrix} → 内联小矩阵（无括号）
         矩阵可嵌套（内层矩阵的 \\ 由嵌套计数自动识别）
  分段    \begin{cases} x, & x>0 \\ -x, & x\le0 \end{cases}（\begin{dcases} 同）
  多行对齐 \begin{aligned} a &= b \\ c &= d \end{aligned}
         （\begin{split} \begin{align} \begin{align*} \begin{gathered} 均支持）
  重音    \bar{x}、\hat{x}、\dot{x}、\ddot{x}、\tilde{x}、\vec{x}、\overrightarrow{AB}、
         \widehat{AB}、\widetilde{AB}（宽帽/宽波浪）
  上下横线 \overline{AB}、\underline{AB}
  括号    \left( ... \right)、\left\{ ... \right\}、\left[ ... \right]、
         \left\langle ... \right\rangle、\left\lVert x \right\rVert、
         \left. ... \right|_{x=0}（'.' 为不可见定界符）、\left(a\middle|b\right)
         （定界符组按 \left...\right 配对成 <m:d>；\bigl( \Big[ 等单侧修饰命令跳过）
  省略号  \dots、\ldots、\cdots、\vdots、\ddots
  叠放    \overset{a}{b}、\underset{a}{b}（上下放置）；
         \overbrace{…}、\underbrace{…}（上下花括号）；
         \xrightarrow{label}、\xleftarrow{label}（带标签箭头）
  否定    \not= → ≠、\not\leq → ≰（基字符后加组合长斜线）；
         预组合符号 \nleq、\ngeq、\nless、\ngtr、\nmid、\nparallel 直接映射
  符号    \alpha、\beta、… 希腊字母；\to、\infty、\cdot、\times、\pm、\leq、\geq、\neq、
         \approx、\in、\notin、\subseteq、\cup、\cap、\partial、\nabla、\emptyset、
         \mid、\div、\top、\perp、\parallel、\rightarrow、\Leftrightarrow、
         \leqslant、\geqslant、\implies、\iff、\leadsto、\lVert、\rVert 等
  低频符号 \lfloor\lceil\rfloor\rceil（取整/四舍五入）、\varnothing、\aleph、\beth、
         \gimel、\daleth、\measuredangle、\preceq、\succeq、\nsim、\nexists、
         \nearrow、\searrow、\swarrow、\nwarrow、\rightleftharpoons、\oslash、
         \dagger、\ddagger、\blacksquare、\varkappa；
         \Re→ℜ、\Im→ℑ（复数实部/虚部标准记号，非正体函数 Re/Im）
  标准命令 \cong、\simeq、\asymp、\doteq、\prec、\succ、\sqsubseteq、\sqsupseteq、
         \models、\vdash、\bowtie、\ni、\owns、\vee、\wedge、\lnot、
         \Downarrow、\Uparrow、\Updownarrow、\Longleftarrow、\Longrightarrow、
         \Longleftrightarrow、\hookleftarrow、\hookrightarrow、harpoon 系列、
         \backslash、\amalg、\uplus、\sqcap、\sqcup、\bigodot、\bigvee、\bigwedge、
         \triangleleft、\triangleright、\bigtriangleup、\square、\checkmark、
         牌花 ♣♦♥♠、音符 ♭♮♯、\imath、\jmath、\pounds、\mathdollar 等
         （与 pandoc 内核 texmath 的标准 LaTeX 符号表对齐，见 test_omml.py）
  重音补 \acute、\grave、\breve、\check、\mathring（组合重音）
  字体补 \mathsf{...} → 无衬线体（scr=sans-serif）；\hspace{len} → 近似一个空格；
         \underbar{x} → 下横线（同 \underline）
  函数    \sin、\cos、\log、\ln、\max、\min、\limsup、\argmin、\argmax 等（正体）
  字体样式 \text{...}、\mathrm{...}（正体，不额外设置）；
         \mathbf{...} → 粗体；\mathit{...} → 斜体；\boldsymbol{...}/\bm{...} → 粗斜；
         \mathbb{...} → 双线体；\mathcal{...}/\mathscr{...} → 花体；
         \mathfrak{...} → 哥特体；\mathtt{...} → 等宽体（样式写入 <m:rPr> 原生生效）
  占位    \phantom{x}、\hphantom、\vphantom（占位，输出为空）；
         \boxed{x} → 公式外框（borderBox）
  脚标    x_{ij}（花括号）、x_{i<j}（组内任意表达式）；
         x_\alpha、\sum_\alpha（2026-08-26：无花括号的命令脚标也支持，
         \alpha 直接作脚标，不再误判为空参数而写 '?' 占位）
  矩阵内  行分隔用 \\，列分隔用 &

【C. 限制说明】
  - 本工具覆盖数学建模报告常见公式形状。若遇到不认识的 \命令，按 SKILL.md 的
    "技能自我完善规则"：扩展本文件（补 LATEX_SYMBOLS 或新增 handler），
    并在文件头注释里登记新语法 + 补一条 test_omml.py 测试用例。
  - 复杂结构（如 \begin{aligned} 内部精确对齐）仅做近似排版，视觉正确即可。
  - 空参数防护（2026-08-26 加固）：公式参数为空时（如 \frac{a}{}、x_{}、\sqrt{}、
    \sum^{}），该槽位写入占位符 "?"，不生成空 <m:e/>/<m:sub/>/<m:den/> 等空元素
    （空元素会在 Word/WPS/LibreOffice 里渲染成虚线占位框）。正常公式不会出现 "?"；
    若渲染结果里出现 "?"，说明公式写错了，须修正。
  - $$ 包裹符兜底（2026-08-26）：build_omath() 会剥掉首尾的 $$...$$（正常流程已由
    split_formula_marks() 剥离，此处是防御——若调用方直接把带 $$ 的完整串传进来，
    不会被当成普通 '$' 字符混进公式）。
"""

import re
from docx.oxml import parse_xml

MNS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'


def _esc(text):
    return (text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))


def _me(tag):
    return parse_xml('<m:%s xmlns:m="%s"/>' % (tag, MNS))


def _t(text):
    return parse_xml('<m:t xmlns:m="%s" xml:space="preserve">%s</m:t>' % (MNS, _esc(text)))


def om_r(text):
    """正体文本 run"""
    r = _me('r')
    rpr = parse_xml('<m:rPr xmlns:m="%s"><m:sty m:val="p"/></m:rPr>' % MNS)
    r.append(rpr)
    r.append(_t(text))
    return r


def om_sSub(base_el, sub_elems):
    el = _me('sSub')
    e = _me('e'); e.append(base_el)
    sub = _me('sub')
    for x in sub_elems: sub.append(x)
    el.append(e); el.append(sub)
    return el


def om_sSup(base_el, sup_elems):
    el = _me('sSup')
    e = _me('e'); e.append(base_el)
    sup = _me('sup')
    for x in sup_elems: sup.append(x)
    el.append(e); el.append(sup)
    return el


def om_nary(sub_elems, sup_elems, expr_elems, chr_val='∑'):
    """累加/积分/连乘等 n 元运算符。

    **必须给全上下限**：`sup_elems`/`sub_elems` 不能为空，LaTeX 侧要写
    `\\sum_{j=1}^{n}`，**禁止写 `\\sum_{j}`**。经实测（2026-08-27 钢管
    下料问题）：OMML `<m:nary>` 缺 `<m:sup>` 上标元素时，交互式 WPS 会在
    运算符上方渲染一个**虚线占位框**（WPS COM 导出 PDF 反而正常，极具迷惑
    性）。此前一度误归因于 `<m:limLoc m:val="undOvr"/>`，证据被"缺 sup"
    混淆，勿再犯。

    `naryPr` 只写 `<m:chr>`，不写 `<m:limLoc>`——这与《线性规划案例分
    析报告》已验证的干净结构一致。
    """
    el = _me('nary')
    pr = parse_xml('<m:naryPr xmlns:m="%s"><m:chr m:val="%s"/></m:naryPr>'
                   % (MNS, chr_val))
    el.append(pr)
    if sub_elems:
        sub = _me('sub')
        for x in sub_elems: sub.append(x)
        el.append(sub)
    if sup_elems:
        sup = _me('sup')
        for x in sup_elems: sup.append(x)
        el.append(sup)
    e = _me('e')
    for x in expr_elems: e.append(x)
    el.append(e)
    return el


def om_delim(inner_elems, beg='(', end=')'):
    el = _me('d')
    pr = parse_xml('<m:dPr xmlns:m="%s"><m:begChr m:val="%s"/><m:endChr m:val="%s"/></m:dPr>'
                   % (MNS, _esc(beg), _esc(end)))
    el.append(pr)
    e = _me('e')
    for x in inner_elems: e.append(x)
    el.append(e)
    return el


def om_frac(num_elems, den_elems):
    el = _me('f')
    num = _me('num')
    for x in num_elems: num.append(x)
    den = _me('den')
    for x in den_elems: den.append(x)
    el.append(num); el.append(den)
    return el


def om_binom(num_elems, den_elems):
    """组合数：隐藏分数线的分数结构（\binom{n}{k}，Word 原生 barHide）。"""
    el = _me('f')
    pr = parse_xml('<m:fPr xmlns:m="%s"><m:barHide m:val="1"/></m:fPr>' % MNS)
    el.append(pr)
    num = _me('num')
    for x in num_elems: num.append(x)
    den = _me('den')
    for x in den_elems: den.append(x)
    el.append(num); el.append(den)
    return el


def om_rad(body_elems, deg_elems=None):
    """根号。deg_elems 为 None 表示平方根（隐藏次数）；否则为 n 次根。"""
    el = _me('rad')
    if deg_elems is None:
        pr = parse_xml('<m:radPr xmlns:m="%s"><m:degHide m:val="1"/></m:radPr>' % MNS)
        el.append(pr)
        deg = _me('deg')
    else:
        deg = _me('deg')
        for x in deg_elems: deg.append(x)
    e = _me('e')
    for x in body_elems: e.append(x)
    el.append(deg); el.append(e)
    return el


def om_acc(chr_val, body_elems):
    """重音（帽子/均值/向量箭头等），字符置于基式上方。"""
    el = _me('acc')
    pr = parse_xml('<m:accPr xmlns:m="%s"><m:chr m:val="%s"/></m:accPr>' % (MNS, _esc(chr_val)))
    el.append(pr)
    e = _me('e')
    for x in body_elems: e.append(x)
    el.append(e)
    return el


def om_bar(pos, body_elems):
    r"""上下横线：pos='top' 为上横线（\overline），'bottom' 为下横线（\underline）。"""
    el = _me('bar')
    pr = parse_xml('<m:barPr xmlns:m="%s"><m:pos m:val="%s"/></m:barPr>' % (MNS, pos))
    el.append(pr)
    e = _me('e')
    for x in body_elems: e.append(x)
    el.append(e)
    return el


def om_matrix(rows):
    """矩阵：rows 为 [[elems,...], ...]。"""
    # 空矩阵防护：至少保留 1 行 1 个空单元格，避免 m:count m:val="0"
    if not rows:
        rows = [[]]
    ncols = max((len(r) for r in rows), default=1)
    if ncols == 0:
        ncols = 1
        rows = [[[]] if not r else r for r in rows]
    el = _me('m')
    pr = parse_xml(
        '<m:mPr xmlns:m="%s"><m:baseJc m:val="center"/><m:plcHide m:val="1"/><m:mcs>'
        '<m:mc><m:mcPr><m:count m:val="%d"/><m:mcJc m:val="center"/></m:mcPr>'
        '</m:mc></m:mcs></m:mPr>' % (MNS, ncols))
    el.append(pr)
    for row in rows:
        mr = _me('mr')
        for cell in row:
            e = _me('e')
            for x in cell: e.append(x)
            mr.append(e)
        for _ in range(ncols - len(row)):
            mr.append(_me('e'))
        el.append(mr)
    return el


def om_eqarr(lines):
    """多行对齐：lines 为每行 [elems,...]。"""
    el = _me('eqArr')
    for line in lines:
        e = _me('e')
        for x in line: e.append(x)
        el.append(e)
    return el


def om_borderbox(content_elems):
    """\boxed{...}：公式外加边框（borderBox 元素）。"""
    el = _me('borderBox')
    e = _me('e')
    for x in content_elems: e.append(x)
    el.append(e)
    return el


# ----------------------------------------------------------------------------
# 解析器
# ----------------------------------------------------------------------------

LATEX_SYMBOLS = {
    'to': '→', 'infty': '∞', 'cdot': '·', 'times': '×', 'pm': '±', 'mp': '∓',
    'leq': '≤', 'le': '≤', 'geq': '≥', 'ge': '≥', 'neq': '≠', 'ne': '≠',
    'approx': '≈', 'equiv': '≡', 'sim': '∼', 'propto': '∝',
    'in': '∈', 'notin': '∉', 'subset': '⊂', 'subseteq': '⊆',
    'supset': '⊃', 'supseteq': '⊇',
    'cup': '∪', 'cap': '∩', 'setminus': '∖', 'emptyset': '∅',
    'forall': '∀', 'exists': '∃', 'neg': '¬', 'land': '∧', 'lor': '∨',
    'partial': '∂', 'nabla': '∇',
    'cdotp': '·', 'ast': '∗', 'star': '⋆', 'circ': '∘', 'bullet': '∙',
    'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ', 'epsilon': 'ϵ',
    'varepsilon': 'ε', 'zeta': 'ζ', 'eta': 'η', 'theta': 'θ', 'vartheta': 'ϑ',
    'iota': 'ι', 'kappa': 'κ', 'lambda': 'λ', 'mu': 'μ', 'nu': 'ν',
    'xi': 'ξ', 'omicron': 'ο', 'pi': 'π', 'varpi': 'ϖ', 'rho': 'ρ',
    'varrho': 'ϱ', 'sigma': 'σ', 'varsigma': 'ς', 'tau': 'τ', 'upsilon': 'υ',
    'phi': 'φ', 'varphi': 'ϕ', 'chi': 'χ', 'psi': 'ψ', 'omega': 'ω',
    'Gamma': 'Γ', 'Delta': 'Δ', 'Theta': 'Θ', 'Lambda': 'Λ', 'Xi': 'Ξ',
    'Pi': 'Π', 'Sigma': 'Σ', 'Upsilon': 'Υ', 'Phi': 'Φ', 'Psi': 'Ψ', 'Omega': 'Ω',
    'lbrace': '{', 'rbrace': '}', 'langle': '⟨', 'rangle': '⟩',
    'vert': '|', 'Vert': '‖', 'because': '∵', 'therefore': '∴',
    'degree': '°', 'prime': '′', 'ell': 'ℓ',
    # 箭头
    'rightarrow': '→', 'leftarrow': '←', 'leftrightarrow': '↔',
    'uparrow': '↑', 'downarrow': '↓', 'updownarrow': '↕',
    'Rightarrow': '⇒', 'Leftarrow': '⇐', 'Leftrightarrow': '⇔',
    'longrightarrow': '→', 'longleftarrow': '←', 'mapsto': '↦', 'longmapsto': '⟼',
    # 省略号点
    'dots': '…', 'ldots': '…', 'cdots': '⋯', 'vdots': '⋮', 'ddots': '⋱',
    # 关系/运算符
    'mid': '|', 'div': '÷', 'top': '⊤', 'perp': '⊥', 'parallel': '∥',
    'angle': '∠', 'll': '≪', 'gg': '≫', 'oplus': '⊕', 'otimes': '⊗',
    'ominus': '⊖', 'odot': '⊙', 'triangle': '△', 'complement': '∁', 'hbar': 'ℏ',
    'leqslant': '≤', 'geqslant': '≥', 'nless': '≮', 'ngtr': '≯',
    'nleq': '≰', 'ngeq': '≱', 'nmid': '∤', 'nparallel': '∦',
    'implies': '⇒', 'iff': '⇔', 'leadsto': '↝',
    'lvert': '|', 'rvert': '|', 'lVert': '‖', 'rVert': '‖',
    # 取整（地板/天花板）
    'lfloor': '⌊', 'rfloor': '⌋', 'lceil': '⌈', 'rceil': '⌉',
    # 数集 / 空集变体 / 基数（希伯来字母）
    'varnothing': '∅', 'aleph': 'ℵ', 'beth': 'ℶ', 'gimel': 'ℷ', 'daleth': 'ℸ',
    # 角度 / 序关系 / 否定
    'measuredangle': '∡', 'preceq': '≼', 'succeq': '≽', 'nsim': '≁', 'nexists': '∄',
    # 斜向箭头 / 可逆箭头
    'nearrow': '↗', 'searrow': '↘', 'swarrow': '↙', 'nwarrow': '↖',
    'rightleftharpoons': '⇌',
    # 运算符 / 其他
    'oslash': '⊘', 'dagger': '†', 'ddagger': '‡',
    'blacksquare': '■', 'varkappa': 'ϰ',
    # 实部 / 虚部（标准记号 ℜ/ℑ，非正体函数 Re/Im）
    'Re': 'ℜ', 'Im': 'ℑ',
    # 补齐 texmath 标准 LaTeX 命令（2026-08-26：对比 pandoc 内核后补全）
    # 箭头（双线 / 长箭头 / 带钩 / harpoon）
    'Downarrow': '⇓', 'Uparrow': '⇑', 'Updownarrow': '⇕',
    'Longleftarrow': '⇐', 'Longrightarrow': '⇒', 'Longleftrightarrow': '⇔',
    'longleftrightarrow': '↔',
    'gets': '←', 'hookleftarrow': '↩', 'hookrightarrow': '↪',
    'leftharpoonup': '↼', 'leftharpoondown': '↽',
    'rightharpoonup': '⇀', 'rightharpoondown': '⇁',
    # 关系 / 序 / 逻辑
    'cong': '≅', 'simeq': '≃', 'asymp': '≍', 'doteq': '≐', 'thickapprox': '≈',
    'iiiint': '∫∫∫∫', 'idotsint': '∫∫∫',  # 四/多积分：无跨引擎独立字形，用多个积分号文本近似
    'prec': '≺', 'succ': '≻', 'sqsubseteq': '⊑', 'sqsupseteq': '⊒',
    'models': '⊨', 'vdash': '⊢', 'bowtie': '⋈', 'ni': '∋', 'owns': '∋',
    'vee': '∨', 'wedge': '∧', 'lnot': '¬', 'bot': '⊥',
    # 集合 / 运算符
    'backslash': '∖', 'amalg': '⨿', 'uplus': '⊎', 'sqcap': '⊓', 'sqcup': '⊔',
    'bigodot': '⨀', 'bigsqcup': '⨆', 'biguplus': '⨄', 'bigvee': '⋁', 'bigwedge': '⋀',
    'triangleleft': '◁', 'triangleright': '▷', 'bigtriangleup': '△', 'bigtriangledown': '▽',
    'slash': '/', 'lbrack': '[', 'rbrack': ']', 'lgroup': '⦃', 'rgroup': '⦄',
    # 字形 / 其他
    'square': '□', 'checkmark': '✓', 'diamond': '◊',
    'clubsuit': '♣', 'diamondsuit': '♦', 'heartsuit': '♥', 'spadesuit': '♠',
    'smile': '⌣', 'frown': '⌢', 'smallsmile': '⌣', 'smallfrown': '⌢',
    'flat': '♭', 'natural': '♮', 'sharp': '♯',
    'imath': 'ı', 'jmath': 'ȷ', 'pounds': '£', 'mathdollar': '$',
}

# 函数名（渲染为正体，如 sin、cos、max）
FUNCTION_NAMES = {'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
                  'sinh', 'cosh', 'tanh', 'arcsin', 'arccos', 'arctan',
                  'log', 'ln', 'lg', 'exp', 'max', 'min', 'sup', 'inf',
                  'limsup', 'liminf', 'det', 'dim', 'ker', 'deg', 'gcd',
                  'arg', 'mod', 'bmod', 'Pr',
                  'argmin', 'argmax', 'arginf', 'argsup'}

# \命令名 → (类别, 参数)
NARY_CHR = {'sum': '∑', 'prod': '∏', 'int': '∫', 'iint': '∬', 'iiint': '∭',
            'oint': '∮', 'oiint': '∯',
            'bigcap': '⋂', 'bigcup': '⋃', 'bigoplus': '⊕',
            'bigotimes': '⊗', 'coprod': '∐'}
ACCENT_CHR = {'bar': '̄', 'hat': '^', 'dot': '˙', 'ddot': '̈', 'tilde': '~',
              'vec': '⃗', 'overrightarrow': '⃗', 'overleftarrow': '⃖',
              'overleftrightarrow': '⃡',
              'widehat': '̂', 'widetilde': '̃',
              'acute': '́', 'grave': '̀', 'breve': '̆', 'check': '̌', 'mathring': '̊'}
STYLE_CMDS = ('text', 'mathrm', 'mathbf', 'mathbb', 'mathit', 'boldsymbol',
              'operatorname', 'mbox', 'mathcal', 'mathfrak', 'mathscr', 'bm',
              'mathtt', 'mathsf')

# 字体样式命令 → (sty, scr)。sty: p/b/i/bi；scr: script/double-struck/fraktur/monospace。
# (None, None) 表示保持正体默认（mathrm/text/mbox/operatorname 不额外设置）。
STYLE_MAP = {
    'mathbf': ('b', None),
    'boldsymbol': ('bi', None),
    'bm': ('bi', None),
    'mathit': ('i', None),
    'mathbb': (None, 'double-struck'),
    'mathcal': (None, 'script'),
    'mathscr': (None, 'script'),
    'mathfrak': (None, 'fraktur'),
    'mathtt': (None, 'monospace'),
    'mathsf': (None, 'sans-serif'),
    'mathrm': (None, None),
    'text': (None, None),
    'mbox': (None, None),
    'operatorname': (None, None),
}

# 定界符大小命令（\big 系列）：只跳过命令本身，括号字符照常处理
SIZING_CMDS = {'big', 'Big', 'bigg', 'Bigg', 'bigl', 'bigr', 'Bigl', 'Bigr',
               'biggl', 'biggr', 'Biggl', 'Biggr', 'bigm', 'Bigm'}

# \left / \right 定界符映射（'.' 为不可见定界符）
DELIM_CHR = {
    '(': '(', ')': ')', '[': '[', ']': ']',
    '{': '{', '}': '}', '|': '|',
    '\\{': '{', '\\}': '}', '\\|': '‖',
    '\\langle': '⟨', '\\rangle': '⟩',
    '\\lvert': '|', '\\rvert': '|',
    '\\lVert': '‖', '\\rVert': '‖',
    '.': '',
}


class MiniMath:
    """受限公式解析器（简单写法 + LaTeX 子集）→ OMML 元素列表"""

    OPS = '+-=≤≥∈,'

    def __init__(self, s):
        self.s = s
        self.n = len(s)
        self.pos = 0
        self.matrix_depth = 0  # 矩阵/多行环境嵌套深度（>0 时 \\ 表示换行、\end 结束环境）
        self.delim_depth = 0   # \left...\right 嵌套深度（>0 时 expr 遇 \right 即停，交给 latex_left 收尾）

    def parse(self):
        return self.expr('')

    def peek(self):
        return self.s[self.pos] if self.pos < self.n else ''

    def expr(self, stop_chars):
        """解析表达式，遇到 stop_chars 中的字符时停止（不消费）"""
        elems = []
        while self.pos < self.n:
            c = self.peek()
            if c in stop_chars:
                break
            if c == ' ':
                self.pos += 1
                j = self.pos
                while j < self.n and self.s[j] == ' ':
                    j += 1
                if j < self.n and self.s[j] not in self.OPS + ')},]^_&' and self.s[j] not in '·×':
                    if elems:
                        elems.append(om_r(' '))
                continue
            if c == '\\':
                # 在矩阵环境中：\\ 是换行，\end 是环境结束 → 都让外层处理
                if self.matrix_depth > 0:
                    if self.pos + 1 < self.n and self.s[self.pos + 1] == '\\':
                        break
                    if self.s[self.pos:self.pos + 4] == '\\end':
                        break
                # 在 \left...\right 内部：遇 \right 即停，交给 latex_left 收尾
                if self.delim_depth > 0 and self.s[self.pos:self.pos + 6] == '\\right':
                    break
                res = self.latex()
                if res is None:
                    continue
                if isinstance(res, list):
                    elems.extend(res)
                else:
                    elems.append(res)
                continue
            if c in self.OPS or c in '·×':
                elems.append(om_r(c))
                self.pos += 1
                continue
            if c in '({[':
                elems.append(self.group())
                continue
            if c in '_^':
                # 括号/结构/命令产物之后紧跟的上标下标：附到前一个元素上。
                # （ident/数字/符号等已在 attach_subsup 内消费掉自己的 _ ^，到不了这里）
                if elems:
                    op = c
                    self.pos += 1
                    inner = self._read_script()
                    base = elems[-1]
                    elems[-1] = om_sSub(base, inner) if op == '_' else om_sSup(base, inner)
                else:
                    # 孤立的 _ / ^（无基式）：跳过
                    self.pos += 1
                    self._read_script()
                continue
            if c == '∑':
                elems.append(self.sum_expr())
                continue
            if c == '&':
                # 不应出现在矩阵外；作为文本兜底
                elems.append(om_r('&'))
                self.pos += 1
                continue
            if c.isdigit():
                elems.append(self.number())
                continue
            if c.isalpha():
                elems.append(self.ident())
                continue
            # 其他字符（中文逗号、～ 等）作为文本
            elems.append(om_r(c))
            self.pos += 1
        return elems

    def group(self):
        c = self.peek()
        close = {'(': ')', '{': '}', '[': ']'}[c]
        self.pos += 1
        inner = self.expr(close)
        if self.pos < self.n and self.peek() == close:
            self.pos += 1
        if not inner:
            inner = [om_r('?')]  # 空括号/空集合：写占位符，避免空 <m:e/> 占位框
        return om_delim(inner, beg=c, end=close)

    def sum_expr(self):
        # 当前字符是 ∑（简单写法）。无 _/^ 时默认不带上下限（不再硬塞 i=1/n）。
        self.pos += 1
        sub_elems = []
        sup_elems = []
        self._skip_limits()
        if self.peek() == '_':
            self.pos += 1
            sub_elems = self._read_script()
        if self.peek() == '^':
            self.pos += 1
            sup_elems = self._read_script()
        operands = self.expr(self.OPS)
        if not operands:
            operands = [om_r('?')]
        return om_nary(sub_elems, sup_elems, operands)

    def _read_script(self):
        r"""读取 _ 或 ^ 后的内容：{...} 组 或 一个 run。空内容写 '?' 占位符。
        组内容以 \substack 开头 → 解析为 om_eqarr 多行堆叠（如 \sum_{\substack{i<j\\j<k}}）。"""
        while self.pos < self.n and self.s[self.pos] == ' ':
            self.pos += 1
        if self.peek() == '\\':
            # 无花括号的命令脚标：x_\alpha → 下标为 α（保持 LaTeX 语义，
            # 不能落回 '?' 占位——那会把正确公式误判成"空参数"）
            res = self.latex()
            if res is None:
                return [om_r('?')]
            if isinstance(res, list):
                res = [e for e in res if e is not None]
                return res or [om_r('?')]
            return [res]
        if self.peek() != '{':
            start = self.pos
            while self.pos < self.n and self.s[self.pos].isalnum():
                self.pos += 1
            if self.pos == start:
                return [om_r('?')]
            return [om_r(self.s[start:self.pos])]
        raw = self._read_braced_raw()
        if raw.lstrip().startswith('\\substack'):
            return [om_eqarr(self._parse_substack(raw))]
        inner = MiniMath(raw).parse()
        return inner or [om_r('?')]

    def _read_braced_raw(self):
        """读取平衡的 { ... } 参数组，返回其内部原始文本（不含外层花括号）。
        self.pos 前进到右花括号之后。未闭合时读到字符串末尾。"""
        while self.pos < self.n and self.s[self.pos] == ' ':
            self.pos += 1
        if self.peek() != '{':
            return ''
        self.pos += 1
        start = self.pos
        depth = 1
        while self.pos < self.n and depth > 0:
            c = self.s[self.pos]
            if c == '{':
                depth += 1
            elif c == '}':
                depth -= 1
            self.pos += 1
        return self.s[start:self.pos - 1] if self.pos > start else ''

    def _split_top_level(self, s):
        r"""按顶层 \\ 切分原始文本（花括号组内的 \\ 忽略），返回各行文本（去首尾空白）。"""
        lines = []
        cur = []
        i, n = 0, len(s)
        while i < n:
            c = s[i]
            if c == '{':
                cur.append(c)
                i += 1
                depth = 1
                while i < n and depth > 0:
                    ch = s[i]
                    cur.append(ch)
                    if ch == '{':
                        depth += 1
                    elif ch == '}':
                        depth -= 1
                    i += 1
            elif c == '\\' and i + 1 < n and s[i + 1] == '\\':
                lines.append(''.join(cur))
                cur = []
                i += 2
            else:
                cur.append(c)
                i += 1
        lines.append(''.join(cur))
        return [ln.strip() for ln in lines]

    def _parse_substack(self, content):
        r"""解析 \substack{行1 \\ 行2 \\ ...} → 每行一个 <m:e> 的多行 eqArr。
        content 为 \substack{...} 的完整参数文本（含 \substack 与花括号）。"""
        content = content.strip()
        if content.startswith('\\substack'):
            content = content[len('\\substack'):].strip()
        if content.startswith('{'):
            depth = 0
            end = 0
            for i, ch in enumerate(content):
                if ch == '{':
                    depth += 1
                elif ch == '}':
                    depth -= 1
                    if depth == 0:
                        end = i
                        break
            content = content[1:end]
        lines = []
        for raw_line in self._split_top_level(content):
            lines.append(MiniMath(raw_line).parse())
        if not lines:
            lines = [[om_r('?')]]
        return lines

    def _skip_limits(self):
        r"""跳过 \limits / \nolimits（及其前导空格）。"""
        while self.peek() == ' ':
            self.pos += 1
        if self.s[self.pos:self.pos + 7] == '\\limits':
            self.pos += 7
        elif self.s[self.pos:self.pos + 9] == '\\nolimits':
            self.pos += 9

    # ---------------- LaTeX 命令 ----------------

    def latex(self):
        """当前位置是 '\\'，解析一个 LaTeX 命令。返回元素/元素列表/None。"""
        self.pos += 1  # 消费反斜杠
        start = self.pos
        while self.pos < self.n and self.s[self.pos].isalpha():
            self.pos += 1
        name = self.s[start:self.pos]
        if not name:
            # 转义特殊字符：\{ \} \| \_ \^ \% \# \& \$ 等
            c = self.peek()
            if c in "{}|%_#&$~^'":
                self.pos += 1
                return om_r(c)
            if c in ',;:':
                # 薄/中等空格：\, \; \: → 一个可见空格（近似）
                self.pos += 1
                return om_r(' ')
            if c == '!':
                # \! 负空格：OMML 无原生负空格，输出空（让两侧字符紧贴）
                self.pos += 1
                return None
            if c == ' ':
                # \ 后跟空格：一个空格
                self.pos += 1
                return om_r(' ')
            return om_r('\\')
        if name in ('frac', 'dfrac', 'tfrac'):
            return self.latex_frac()
        if name in ('binom', 'dbinom', 'tbinom'):
            return self.latex_binom()
        if name == 'sqrt':
            return self.latex_sqrt()
        if name in FUNCTION_NAMES:
            return self.attach_subsup(om_r(name))
        if name in NARY_CHR:
            return self.latex_nary(NARY_CHR[name])
        if name == 'lim':
            el = om_r('lim')
            self._skip_limits()
            while self.peek() == ' ':
                self.pos += 1
            if self.peek() == '_':
                self.pos += 1
                el = om_sSub(el, self._read_script())
            return el
        if name in ACCENT_CHR:
            return self.latex_accent(ACCENT_CHR[name])
        if name in ('overline', 'underline'):
            return self.latex_bar('top' if name == 'overline' else 'bottom')
        if name in ('overset', 'underset'):
            return self.latex_stack_lim('limUpp' if name == 'overset' else 'limLow')
        if name in ('overbrace', 'underbrace'):
            return self.latex_groupchr('⏞' if name == 'overbrace' else '⏟',
                                       'top' if name == 'overbrace' else 'bottom')
        if name == 'pmod':
            return self.latex_pmod()
        if name == 'not':
            return self.latex_not()
        if name == 'boxed':
            return om_borderbox(self.read_braced_group())
        if name in ('xrightarrow', 'xleftarrow'):
            return self.latex_xarrow('→' if name == 'xrightarrow' else '←')
        if name == 'begin':
            return self.latex_begin()
        if name == 'left':
            return self.latex_left()
        if name == 'right':
            # 裸 \right（无配对 \left）：仅跳过（向后兼容）
            return self.latex_lr()
        if name in SIZING_CMDS:
            return self.latex_lr()
        if name in STYLE_CMDS:
            # 可选的 *（如 \operatorname*{arg\,min}）跳过，避免 '?' 垃圾
            while self.peek() == ' ':
                self.pos += 1
            if self.peek() == '*':
                self.pos += 1
            body = self.read_braced_group()
            sty, scr = STYLE_MAP.get(name, (None, None))
            if sty or scr:
                self._apply_style(body, sty, scr)
            return body
        if name in ('phantom', 'hphantom', 'vphantom'):
            # 占位符：消费参数组，输出空（OMML 无 phantom，对齐近似省略）
            self.read_braced_group()
            return None
        if name == 'hspace':
            # \hspace{len}：OMML 无 em 间距控制，近似为一个空格
            self.read_braced_group()
            return om_r(' ')
        if name == 'underbar':
            # \underbar{x}：下横线（同 \underline）
            return self.latex_bar('bottom')
        if name == 'middle':
            # \middle 定界符修饰：跳过，| 等字符照常处理
            return None
        if name in ('limits', 'nolimits'):
            # \limits / \nolimits：跳过（累加符上下限已由 nary 结构表达）
            return None
        if name in ('quad', 'qquad'):
            return om_r('    ' if name == 'qquad' else '  ')
        if name in LATEX_SYMBOLS:
            return self.attach_subsup(om_r(LATEX_SYMBOLS[name]))
        # 未知命令：原样输出文本
        return self.attach_subsup(om_r('\\' + name))

    # ---------------- 字体样式 ----------------

    def _apply_style(self, elems, sty=None, scr=None):
        """把字体样式（sty / scr）应用到 elems 内所有后代 <m:r>（find-or-create rPr）。
        sty 用 _merge_sty 合并（p+b→b、p+i→i、i+b→bi），绝不重复插入。"""
        for el in elems:
            for r in el.iter('{%s}r' % MNS):
                rPr = r.find('{%s}rPr' % MNS)
                if rPr is None:
                    rPr = _me('rPr')
                    r.insert(0, rPr)
                if sty:
                    self._merge_sty(rPr, sty)
                if scr:
                    self._set_scr(rPr, scr)
        return elems

    def _merge_sty(self, rPr, want):
        """合并粗斜样式：p+b→b、p+i→i、i+b→bi、p+bi→bi。已有样式则更新，绝不重复插入。"""
        cur_el = rPr.find('{%s}sty' % MNS)
        cur = cur_el.get('{%s}val' % MNS) if cur_el is not None else 'p'
        merged = cur + want
        if 'b' in merged and 'i' in merged:
            final = 'bi'
        elif 'b' in merged:
            final = 'b'
        elif 'i' in merged:
            final = 'i'
        else:
            final = 'p'
        if cur_el is None:
            cur_el = parse_xml('<m:sty xmlns:m="%s" m:val="%s"/>' % (MNS, final))
            rPr.append(cur_el)
        else:
            cur_el.set('{%s}val' % MNS, final)
        return rPr

    def _set_scr(self, rPr, scr):
        """设置脚本字体（double-struck/script/fraktur/monospace）。"""
        scr_el = rPr.find('{%s}scr' % MNS)
        if scr_el is None:
            scr_el = parse_xml('<m:scr xmlns:m="%s" m:val="%s"/>' % (MNS, scr))
            rPr.append(scr_el)
        else:
            scr_el.set('{%s}val' % MNS, scr)

    def read_braced_group(self):
        r"""读取 { ... } 参数组，返回内部解析出的元素列表。空组写 '?' 占位符。
        无花括号时按 LaTeX 语义读取单个参数（一个字符），如 \frac12、\frac ab、\sqrt x。"""
        while self.pos < self.n and self.s[self.pos] == ' ':
            self.pos += 1
        if self.peek() == '{':
            self.pos += 1
            inner = self.expr('}')
            if self.pos < self.n and self.peek() == '}':
                self.pos += 1
            return inner or [om_r('?')]
        if self.pos < self.n:
            c = self.s[self.pos]
            self.pos += 1
            return [om_r(c)]
        return [om_r('?')]

    def latex_frac(self):
        num = self.read_braced_group()
        den = self.read_braced_group()
        return om_frac(num, den)

    def latex_binom(self):
        top = self.read_braced_group()
        bottom = self.read_braced_group()
        return om_binom(top, bottom)

    def latex_stack_lim(self, kind):
        r"""\overset{上}{下} / \underset{下}{上}：上下叠放（limUpp/limLow）。"""
        first = self.read_braced_group()
        second = self.read_braced_group()
        el = _me('limLow' if kind == 'limLow' else 'limUpp')
        e = _me('e')
        for x in second: e.append(x)
        lim = _me('lim')
        for x in first: lim.append(x)
        el.append(e); el.append(lim)
        return el

    def latex_groupchr(self, chr_val, pos):
        r"""\overbrace{...} / \underbrace{...}：上下花括号（groupChr 元素）。"""
        body = self.read_braced_group()
        el = _me('groupChr')
        pr = parse_xml('<m:groupChrPr xmlns:m="%s"><m:chr m:val="%s"/><m:pos m:val="%s"/></m:groupChrPr>'
                       % (MNS, chr_val, pos))
        el.append(pr)
        e = _me('e')
        for x in body: e.append(x)
        el.append(e)
        return el

    def latex_pmod(self):
        r"""\pmod{n} → (mod n)。"""
        while self.peek() == ' ':
            self.pos += 1
        inner = self.read_braced_group()
        return [om_r(' '), om_delim([om_r('mod'), om_r(' ')] + inner, '(', ')')]

    def latex_not(self):
        r"""\not 否定前缀：\not= → ≠，\not\leq → ≰（基字符后追加组合长斜线 U+0338）。
        优先使用预组合字符（如 \nleq 直接给 ≰）；\not 是兜底。"""
        while self.peek() == ' ':
            self.pos += 1
        if self.peek() == '\\':
            # 下一个是命令：解析取字符文本
            start = self.pos
            inner = self.latex()
            txt = ''
            for el in (inner if isinstance(inner, list) else [inner]):
                if el is None:
                    continue
                t = el.find('{%s}t' % MNS)
                if t is not None and t.text:
                    txt += t.text
            if txt:
                return om_r(txt + '̸')
            return om_r('̸')
        if self.pos < self.n:
            c = self.s[self.pos]
            self.pos += 1
            return om_r(c + '̸')
        return om_r('?')

    def latex_xarrow(self, arrow):
        r"""\xrightarrow{label} → 箭头（下方/上方带标签，limUpp 结构近似）。"""
        label = self.read_braced_group()
        el = _me('limUpp')
        e = _me('e')
        e.append(om_r(arrow))
        lim = _me('lim')
        for x in label:
            lim.append(x)
        el.append(e)
        el.append(lim)
        return el

    def latex_sqrt(self):
        deg_elems = None
        while self.peek() == ' ':
            self.pos += 1
        if self.peek() == '[':
            self.pos += 1
            deg_elems = self.expr(']')
            if self.pos < self.n and self.peek() == ']':
                self.pos += 1
            if not deg_elems:
                deg_elems = [om_r('?')]
        body = self.read_braced_group()
        return om_rad(body, deg_elems)

    def latex_nary(self, chr_val):
        sub_elems, sup_elems = [], []
        self._skip_limits()
        if self.peek() == '_':
            self.pos += 1
            sub_elems = self._read_script()
            while self.peek() == ' ':
                self.pos += 1
        if self.peek() == '^':
            self.pos += 1
            sup_elems = self._read_script()
        operands = self.expr(self.OPS)
        if not operands:
            operands = [om_r('?')]
        return om_nary(sub_elems, sup_elems, operands, chr_val)

    def latex_accent(self, chr_val):
        body = self.read_braced_group()
        return om_acc(chr_val, body)

    def latex_bar(self, pos):
        body = self.read_braced_group()
        return om_bar(pos, body)

    def latex_lr(self):
        r"""\right / \big 系列（单侧修饰）：只跳过命令本身（和 \right. 的不可见点 '.'）。
        \left 走 latex_left（需配对收尾），此处仅处理不需配对的情况。"""
        while self.peek() == ' ':
            self.pos += 1
        if self.peek() == '.':
            # \left. / \right. 的不可见占位点
            self.pos += 1
        return None

    def latex_left(self):
        r"""\left( ... \right) 定界符组：读开定界符 → 解析内部 → 消费 \right + 闭定界符。
        支持 \{ \langle \| \lVert 与不可见定界符 .（\left. ... \right|）；嵌套由 delim_depth 计数。"""
        beg = self._read_delimiter()
        self.delim_depth += 1
        try:
            inner = self.expr('')
        finally:
            self.delim_depth -= 1
        end = ''
        while self.pos < self.n and self.s[self.pos] == ' ':
            self.pos += 1
        if self.s[self.pos:self.pos + 6] == '\\right':
            self.pos += 6
            end = self._read_delimiter()
        if not inner:
            inner = [om_r('?')]
        return om_delim(inner, beg, end)

    def _read_delimiter(self):
        r"""读取 \left / \right 后的定界符，返回映射字符（'.' → '' 不可见）。"""
        while self.pos < self.n and self.s[self.pos] == ' ':
            self.pos += 1
        if self.pos >= self.n:
            return ''
        if self.s[self.pos] == '\\':
            self.pos += 1
            start = self.pos
            while self.pos < self.n and self.s[self.pos].isalpha():
                self.pos += 1
            key = '\\' + self.s[start:self.pos]
            if key in DELIM_CHR:
                return DELIM_CHR[key]
            # 单字符转义定界符：\{ \} \|
            if self.pos < self.n and self.s[self.pos] in '{}|':
                self.pos += 1
                return self.s[self.pos - 1]
            return ''
        c = self.s[self.pos]
        if c in '()[]{}|.':
            self.pos += 1
            return DELIM_CHR.get(c, c)
        return ''

    def latex_begin(self):
        env = self.read_env_name()
        if env in ('array', 'tabular'):
            self.skip_column_spec()
        rows = self.parse_matrix_rows()
        if env in ('aligned', 'split', 'align', 'align*', 'gathered'):
            lines = []
            for row in rows:
                merged = []
                for cell in row:
                    if merged:
                        merged.append(om_r(' '))
                    merged.extend(cell)
                lines.append(merged)
            return om_eqarr(lines)
        mat = om_matrix(rows)
        if env in ('matrix', 'smallmatrix'):
            return mat
        if env == 'pmatrix':
            return om_delim([mat], '(', ')')
        if env == 'bmatrix':
            return om_delim([mat], '[', ']')
        if env == 'Bmatrix':
            return om_delim([mat], '{', '}')
        if env == 'vmatrix':
            return om_delim([mat], '|', '|')
        if env == 'Vmatrix':
            return om_delim([mat], '‖', '‖')
        if env in ('cases', 'dcases'):
            return om_delim([mat], '{', '')
        # 未知环境：按普通矩阵输出
        return mat

    def skip_column_spec(self):
        r"""\begin{array}[t]{ccc} 的位置/列格式参数：跳过 [t] 与 {ccc}，不进入矩阵内容。"""
        while self.peek() == ' ':
            self.pos += 1
        if self.peek() == '[':
            self.pos += 1
            while self.pos < self.n and self.s[self.pos] != ']':
                self.pos += 1
            if self.pos < self.n:
                self.pos += 1
            while self.peek() == ' ':
                self.pos += 1
        if self.peek() == '{':
            self.pos += 1
            while self.pos < self.n and self.s[self.pos] != '}':
                self.pos += 1
            if self.pos < self.n:
                self.pos += 1

    def read_env_name(self):
        while self.peek() == ' ':
            self.pos += 1
        if self.peek() == '{':
            self.pos += 1
            start = self.pos
            while self.pos < self.n and self.s[self.pos] != '}':
                self.pos += 1
            name = self.s[start:self.pos]
            if self.pos < self.n:
                self.pos += 1
            return name
        return ''

    def parse_matrix_rows(self):
        r"""从 \begin{env} 后开始，按 \\ 分行、& 分列，直到 \end{env}。
        嵌套环境（如矩阵单元格里再放 \begin{pmatrix}）通过 matrix_depth 计数，
        内层返回后外层仍能正确识别 \\ 换行。"""
        self.matrix_depth += 1
        rows = []
        try:
            while self.pos < self.n:
                cells = [self.expr('&')]
                while self.peek() == '&':
                    self.pos += 1
                    cells.append(self.expr('&'))
                rows.append(cells)
                if self.s[self.pos:self.pos + 2] == '\\\\':
                    self.pos += 2
                    continue
                if self.s[self.pos:self.pos + 4] == '\\end':
                    self.pos += 4
                    # 读 {env}（find 而非 index：未闭合时正常终止，不抛 ValueError）
                    if self.peek() == '{':
                        self.pos += 1
                        endi = self.s.find('}', self.pos)
                        if endi == -1:
                            break
                        self.pos = endi + 1
                    break
                # 既不是换行也不是 \end：终止
                break
        finally:
            self.matrix_depth -= 1
        return rows

    # ---------------- 简单写法的 token ----------------

    def ident(self):
        start = self.pos
        while self.pos < self.n and self.s[self.pos].isalpha():
            self.pos += 1
        base_el = om_r(self.s[start:self.pos])
        return self.attach_subsup(base_el)

    def number(self):
        start = self.pos
        while self.pos < self.n and (self.s[self.pos].isdigit() or self.s[self.pos] == '.'):
            self.pos += 1
        return self.attach_subsup(om_r(self.s[start:self.pos]))

    def attach_subsup(self, base_el):
        while self.pos < self.n and self.peek() in '_^':
            op = self.peek()
            self.pos += 1
            inner = self._read_script()
            if op == '_':
                base_el = om_sSub(base_el, inner)
            else:
                base_el = om_sSup(base_el, inner)
        return base_el


def build_omath(expr):
    """返回 <m:oMath> 元素，内含解析出的公式子元素"""
    expr = expr.strip()
    # 防御：若调用方误把 $$...$$ 包裹的完整公式直接传进来（如 FI() 的结果），
    # 剥掉首尾 $$ 再解析，避免 $$ 作为普通 '$' 字符混进公式。
    # 正常流程中 $$ 已由 split_formula_marks() 剥离，此处仅为兜底。
    if len(expr) >= 4 and expr.startswith('$$') and expr.endswith('$$'):
        expr = expr[2:-2].strip()
    om = _me('oMath')
    for el in MiniMath(expr).parse():
        om.append(el)
    return om


def split_formula_marks(text):
    """把含 $$...$$ 标记的文本拆成 [('t', 文本), ('m', 公式), ...]"""
    parts = re.split(r'\$\$(.*?)\$\$', text)
    out = []
    for k, seg in enumerate(parts):
        if not seg:
            continue
        if k % 2 == 1:
            out.append(('m', seg))
        else:
            out.append(('t', seg))
    return out
