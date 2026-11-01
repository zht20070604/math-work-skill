# -*- coding: utf-8 -*-
"""
omml_builder.py 的测试用例：断言每类公式生成的 OMML 结构正确 + 向后兼容。
运行：python test_omml.py
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lxml import etree
from omml_builder import build_omath

# Windows 默认控制台是 GBK，ℜ/∅ 等 Unicode 符号会在打印时触发 UnicodeEncodeError。
# 强制 stdout 用 UTF-8，保证任意编码控制台都能直接运行本测试。
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass  # 旧版 Python 无 reconfigure，忽略

PASS = 0
FAIL = 0


def check_count(name, formula, frag, want):
    """断言公式 XML 中 frag 出现次数恰为 want（嵌套结构计数用）"""
    global PASS, FAIL
    try:
        om = build_omath(formula)
        xml = etree.tostring(om, encoding='unicode')
    except Exception as e:
        FAIL += 1
        print('[FAIL] %-28s 公式解析出错: %s' % (name, e))
        return
    got = xml.count(frag)
    if got == want:
        PASS += 1
        print('[ OK ] %-28s  %s' % (name, formula))
    else:
        FAIL += 1
        print('[FAIL] %-28s  %s 出现次数=%d，期望=%d' % (name, frag, got, want))


def check(name, formula, must_contain, must_not_contain=()):
    global PASS, FAIL
    try:
        om = build_omath(formula)
        xml = etree.tostring(om, encoding='unicode')
    except Exception as e:
        FAIL += 1
        print('[FAIL] %-28s 公式解析出错: %s' % (name, e))
        return
    missing = [m for m in must_contain if m not in xml]
    unexpected = [m for m in must_not_contain if m in xml]
    if not missing and not unexpected:
        PASS += 1
        print('[ OK ] %-28s  %s' % (name, formula))
    else:
        FAIL += 1
        print('[FAIL] %-28s 缺=%s 多=%s' % (name, missing, unexpected))


# ===== LaTeX 新增语法 =====
check('分数 frac', r'\frac{a}{b}', ['<m:f>', '<m:num>', '<m:den>'])
check('分数 frac(嵌套)', r'\frac{x_{1}+x_{2}}{2}', ['<m:f>', '<m:sSub>', '<m:num>', '<m:den>'])
check('平方根 sqrt', r'\sqrt{x+1}', ['<m:rad>', '<m:degHide m:val="1"/>', '<m:e>'])
check('n次根 sqrt[3]', r'\sqrt[3]{x}', ['<m:rad>', '<m:deg>'], ['m:degHide'])
check('求和 sum', r'\sum_{i=1}^{n}x_{i}',
      ['<m:nary>', 'm:chr m:val="∑"', '<m:sub>', '<m:sup>'], ['m:limLoc'])
check('积分 int', r'\int_{0}^{1}f(x)dx', ['<m:nary>', 'm:chr m:val="∫"'], ['m:limLoc'])
check('连乘 prod', r'\prod_{i=1}^{n}i', ['<m:nary>', 'm:chr m:val="∏"'], ['m:limLoc'])
check('极限 lim', r'\lim_{x \to 0}f(x)', ['<m:sSub>', 'lim', '→'])
check('矩阵 pmatrix', r'\begin{pmatrix}a&b\\c&d\end{pmatrix}',
      ['<m:d>', 'm:begChr m:val="("', 'm:endChr m:val=")"', '<m:m>', '<m:mr>', 'm:count m:val="2"'])
check('矩阵 bmatrix', r'\begin{bmatrix}1&2\\3&4\end{bmatrix}',
      ['m:begChr m:val="["', 'm:endChr m:val="]"'])
check('矩阵 vmatrix', r'\begin{vmatrix}1&2\\3&4\end{vmatrix}',
      ['m:begChr m:val="|"'])
check('矩阵 3x3', r'\begin{matrix}a&b&c\\d&e&f\\g&h&i\end{matrix}',
      ['m:count m:val="3"', '<m:mr>'], ['m:begChr'])
check('分段 cases', r'\begin{cases}x,&x>0\\-x,&x\le0\end{cases}',
      ['m:begChr m:val="{"', 'm:endChr m:val=""', '<m:m>'])
check('多行对齐 aligned', r'\begin{aligned}a&=b\\c&=d\end{aligned}', ['<m:eqArr>'])
check('均值 bar', r'\bar{x}', ['<m:acc>', 'm:chr m:val="̄"'])
check('帽子 hat', r'\hat{x}', ['<m:acc>', 'm:chr m:val="^"'])
check('向量箭头 vec', r'\vec{x}', ['<m:acc>', 'm:chr m:val="⃗"'])
check('向量 overrightarrow', r'\overrightarrow{AB}', ['<m:acc>', 'm:chr m:val="⃗"'])
check('上横线 overline', r'\overline{AB}', ['<m:bar>', 'm:pos m:val="top"'])
check('下横线 underline', r'\underline{x}', ['<m:bar>', 'm:pos m:val="bottom"'])
check('希腊字母', r'\alpha+\beta+\sigma^{2}', ['α', 'β', 'σ', '<m:sSup>'])
check('函数名 sin', r'\sin x + \cos y + \ln z', ['sin', 'cos', 'ln'])
check('薄空格', r'\int_{0}^{1}f(x)\,dx', ['dx'])
check('符号映射', r'x \to \infty \cdot y \le 1 \in \mathbb{R}',
      ['→', '∞', '·', '≤', '∈', 'R'])
check('符号后跟上标', r'\sigma^{2}', ['σ', '<m:sSup>'])
check('括号 left-right', r'\left(x+\frac{1}{2}\right)', ['<m:d>', '<m:f>'])
check('正体 text', r'\text{min } z', ['min', 'z'])

# ===== 空参数防护（2026-08-26：空组写 '?' 占位，避免空元素占位框）=====
check('空分母防护', r'\frac{a}{}', ['?'], ['<m:den/>'])
check('空根号防护', r'\sqrt{}', ['?'], ['<m:e/>'])
check('空下标防护', 'x_{}', ['?'], ['<m:sub/>'])
check('空集合防护', r'x_{ij}\in{}', ['?', '<m:d>'], ['<m:e/>'])
check('空上下标防护', r'\sum_{}^{}', ['?'], ['<m:sub/>', '<m:sup/>', '<m:e/>'])

# ===== 已实现但此前无测试的算符（2026-08-26 补齐）=====
check('二重积分 iint', r'\iint_{D}f(x,y)dxdy', ['<m:nary>', 'm:chr m:val="∬"'])
check('环路积分 oint', r'\oint_{C}Pdx+Qdy', ['<m:nary>', 'm:chr m:val="∮"'])
check('大交 bigcap', r'\bigcap_{i=1}^{n}A_{i}', ['<m:nary>', 'm:chr m:val="⋂"'])
check('大并 bigcup', r'\bigcup_{i=1}^{n}B_{i}', ['<m:nary>', 'm:chr m:val="⋃"'])
check('余积 coprod', r'\coprod_{i\in I}X_{i}', ['<m:nary>', 'm:chr m:val="∐"'])

# ===== $$ 包裹符兜底（2026-08-26：build_omath 自动剥离首尾 $$）=====
check('$$剥离防护', '$$x_{11}$$', ['<m:sSub>'], ['$'])

# ===== 建模常见类别补齐（2026-08-26：括号后上标、省略号、binom、array 等）=====
check('括号后上标', r'(a+b)^2', ['<m:sSup>', '<m:d>'])
check('left-right后上标', r'\left(\frac{1}{2}\right)^{3}', ['<m:sSup>', '<m:f>'], ['^'])
check('sqrt后上标', r'\sqrt{x}^2', ['<m:sSup>', '<m:rad>'])
check('bar后上标', r'\bar{x}^2', ['<m:sSup>', '<m:acc>'])
check('矩阵转置', r'\begin{pmatrix}a&b\\c&d\end{pmatrix}^{T}', ['<m:sSup>'], ['^'])
check('正态分布PDF', r'f(x)=\frac{1}{\sqrt{2\pi}\sigma}e^{-\frac{(x-\mu)^2}{2\sigma^2}}',
      ['<m:rad>', '<m:f>', '<m:sSup>', 'π', 'σ', 'μ'], ['^'])
check('省略号点', r'x_{1}+\cdots+x_{n}', ['⋯'])
check('矩阵省略号', r'\begin{pmatrix}a&\cdots&b\\\vdots&\ddots&\vdots\end{pmatrix}', ['⋯', '⋮', '⋱'])
check('集合mid', r'\{x\mid x>0\}', ['|', '{'])
check('除号div', r'a\div b', ['÷'])
check('转置top', r'A^{\top}', ['⊤'])
check('导数码prime', r'f\'(x)', ["'"])
check('二项式binom', r'\binom{n}{k}', ['<m:barHide m:val="1"/>'])
check('array列格式', r'\begin{array}{ll}x,&x>0\end{array}', ['<m:m>'], ['{ll}'])
check('big定界符', r'\bigl(\frac{a}{b}\bigr)', ['<m:d>'], ['big'])
check('mathcal字体', r'\mathcal{L}', ['L'], ['\\mathcal'])
check('overset叠放', r'\overset{\sim}{X}', ['<m:limUpp>', '∼', 'X'])
check('underbrace花括号', r'\underbrace{a+b}_{=c}', ['<m:groupChr>', '⏟'], ['underbrace'])
check('pmod取模', r'x\equiv a\pmod{m}', ['mod', '≡', '<m:d>'])
check('Bmatrix花括号矩阵', r'\begin{Bmatrix}a&b\\c&d\end{Bmatrix}', ['m:begChr m:val="{"'])
check('Vmatrix双竖线矩阵', r'\begin{Vmatrix}a&b\\c&d\end{Vmatrix}', ['m:begChr m:val="‖"'])

# ===== 简单写法向后兼容 =====
check('简单下标', 'x_{11}', ['<m:sSub>'])
check('简单求和', '∑_{i=1}^{4}c_{ij} x_{ij}', ['<m:nary>', 'm:chr m:val="∑"'])
check('简单上标', 'm^{3}', ['<m:sSup>'])
check('集合花括号', 'x_{ij}∈{0,1}', ['<m:sSub>', '<m:d>'])
check('数字运算', '1.25-0.25=1.00', ['1.25', '0.25'])
check('中文+公式混排', '满负荷费用÷有效台时', ['满负荷费用', '÷'])
check('线性规划整段', 'max z=1.00(x_{11}+x_{12})+0.05x_{21}', ['<m:sSub>', 'max', 'z'])

# ===== 2026-08-26 增强：健壮性修复（Phase 1）=====
check('无花括号参数 frac12', r'\frac12', ['<m:f>'], ['?'])
check('无花括号参数 frac ab', r'\frac ab', ['<m:f>', 'a', 'b'], ['?'])
check('无花括号参数 sqrt x', r'\sqrt x', ['<m:rad>', 'x'], ['?'])
check('operatorname星号', r'\operatorname*{arg\,min}', ['arg', 'min'], ['?'])
check('裸∑无上下限', r'\sum x_i', ['<m:nary>'],
      ['i=1', '<m:sup>', 'm:limLoc'])
check('limits跳过', r'\sum\limits_{i=1}^{n}x_i', ['<m:nary>'], ['limits'])
check('phantom消费', r'\phantom{x}+y', ['y'], ['phantom'])
check('middle跳过', r'\left(a\middle|b\right)', ['<m:d>', '|'], ['middle'])
check('空矩阵防护', r'\begin{matrix}\end{matrix}', ['<m:m>'], ['m:count m:val="0"'])

# ===== 2026-08-26 增强：新语法覆盖（Phase 2）=====
check('argmin函数', r'\argmin_{i} f(i)', ['argmin', '<m:sSub>'], ['\\argmin'])
check('not否定', r'a\not= b', ['a', 'b', '̸'], ['\\not'])
check('not否定符号', r'\not\leq', ['≤', '̸'], ['\\not'])
check('nleq预组合', r'\nleq', ['≰'])
check('leqslant', r'\leqslant', ['≤'])
check('implies', r'p\implies q', ['⇒'])
check('宽帽widehat', r'\widehat{AB}', ['<m:acc>', '̂'], ['widehat'])
check('boxed边框', r'\boxed{x=1}', ['<m:borderBox>', 'x'], ['boxed'])
check('dcases环境', r'\begin{dcases}x,&x>0\\-x,&x\le0\end{dcases}', ['m:begChr m:val="{'], ['dcases'])
check('smallmatrix环境', r'\begin{smallmatrix}a&b\\c&d\end{smallmatrix}', ['<m:m>'], ['begChr'])
check('align环境', r'\begin{align}a&=b\\c&=d\end{align}', ['<m:eqArr>'], ['align'])
check('xrightarrow', r'\xrightarrow{label}', ['<m:limUpp>', '→', 'label'], ['xrightarrow'])
check('lVert范数', r'\lVert x \rVert', ['‖'], ['lVert'])
check('array位置参数', r'\begin{array}[t]{ll}x,&x>0\end{array}', ['<m:m>'], ['[t]'])

# ===== 2026-08-26 增强：嵌套矩阵（Phase 3）=====
NESTED = r'\begin{pmatrix}a&\begin{pmatrix}b&c\\d&e\end{pmatrix}\\f&g\end{pmatrix}'
check_count('嵌套矩阵层数', NESTED, '<m:m>', 2)
check_count('嵌套矩阵总行数', NESTED, '<m:mr>', 4)
check_count('嵌套矩阵列数', NESTED, 'm:count m:val="2"', 2)
check('嵌套矩阵无字面反斜杠', NESTED, [], ['\\'])

# ===== 2026-08-26 增强：脚本参数重写 + \substack（Phase 4）=====
check('substack堆叠', r'\sum_{\substack{i<j\\j<k}} x_{ij}',
      ['<m:nary>', '<m:eqArr>', 'i', 'j'], ['substack', '\\', '?'])
check('脚本组含小于号', r'\sum_{i<j} x_{ij}', ['<m:nary>', 'i', '<', 'j'], ['?'])

# ===== 2026-08-26 增强：\left\{ 等转义定界符（Phase 5）=====
check('left花括号定界符', r'\left\{\frac{a}{b}\right\}', ['<m:d>', 'm:begChr m:val="{"', '<m:f>'], ['\\', 'left', 'right'])
check('left不可见点竖线', r'\left.\frac{dy}{dx}\right|_{x=0}', ['<m:d>', 'm:endChr m:val="|"', '<m:sSub>'], ['left', 'right'])
check('left尖括号定界符', r'\left\langle x \right\rangle', ['<m:d>', '⟨', '⟩'], ['langle', 'rangle'])
check('left双竖线范数', r'\left\lVert x \right\rVert', ['<m:d>', '‖'], ['lVert', 'rVert'])

# ===== 2026-08-26 增强：字体样式落地（Phase 6）=====
check('mathbf粗体', r'\mathbf{X}', ['m:sty m:val="b"'], ['m:sty m:val="p"', '\\mathbf'])
check('mathbb双线', r'\mathbb{R}', ['m:scr m:val="double-struck"', 'R'], ['\\mathbb'])
check('mathcal花体', r'\mathcal{L}', ['m:scr m:val="script"', 'L'], ['\\mathcal'])
check('mathfrak哥特', r'\mathfrak{g}', ['m:scr m:val="fraktur"', 'g'], ['\\mathfrak'])
check('boldsymbol粗斜', r'\boldsymbol{\beta}', ['m:sty m:val="bi"', 'β'], ['\\boldsymbol'])
check('mathtt等宽', r'\mathtt{T}', ['m:scr m:val="monospace"', 'T'], ['\\mathtt'])
check('mathrm保持正体', r'\mathrm{d}x', ['d'], ['m:sty m:val="b"', '\\mathrm'])

# ===== 2026-08-26 补充：低频符号补齐（opencode 报告 18+2 项，全数实现）=====
check('取整 lfloor/rfloor', r'\lfloor x \rfloor', ['⌊', '⌋'], ['lfloor', 'rfloor'])
check('上取整 lceil/rceil', r'\lceil \frac{a}{b} \rceil', ['⌈', '⌉', '<m:f>'], ['lceil', 'rceil'])
check('空集变体 varnothing', r'A\varnothing B', ['∅'], ['varnothing'])
check('基数 aleph下标', r'\aleph_{0}', ['ℵ', '<m:sSub>'], ['aleph'])
check('希伯来基数三连', r'\beth \gimel \daleth', ['ℶ', 'ℷ', 'ℸ'], ['beth', 'gimel', 'daleth'])
check('测角 measuredangle', r'\measuredangle ABC', ['∡'], ['measuredangle'])
check('偏序 preceq/succeq', r'a\preceq b \succeq c', ['≼', '≽'], ['preceq', 'succeq'])
check('非同构 nsim', r'a\nsim b', ['≁'], ['nsim'])
check('不存在 nexists', r'\nexists x', ['∄'], ['nexists'])
check('斜向箭头四连', r'\nearrow \searrow \swarrow \nwarrow', ['↗', '↘', '↙', '↖'], ['nearrow'])
check('可逆箭头', r'\rightleftharpoons', ['⇌'], ['rightleftharpoons'])
check('斜除 oslash', r'a\oslash b', ['⊘'], ['oslash'])
check('剑标 dagger/ddagger', r'\dagger \ddagger', ['†', '‡'], ['dagger'])
check('证毕方块 blacksquare', r'\blacksquare', ['■'], ['blacksquare'])
check('变体kappa varkappa', r'\varkappa', ['ϰ'], ['varkappa'])
check('实部符号 Re→ℜ', r'\Re(z)', ['ℜ'], ['Re', '\\Re'])
check('虚部符号 Im→ℑ', r'\Im(z)', ['ℑ'], ['Im', '\\Im'])
check('取整+基数组合', r'|\lfloor x \rfloor|=\aleph_0', ['⌊', '⌋', 'ℵ'], ['\\'])

# ===== 2026-08-26 补充：对齐 texmath 标准 LaTeX 命令（74 个缺项补齐）=====
check('双线箭头', r'\Downarrow \Uparrow \Updownarrow', ['⇓', '⇑', '⇕'], ['Down'])
check('长箭头', r'\Longleftarrow \Longrightarrow \Longleftrightarrow \longleftrightarrow', ['⇐', '⇒', '⇔', '↔'], ['Long'])
check('带钩箭头', r'\hookleftarrow \hookrightarrow', ['↩', '↪'], ['hook'])
check('harpoon箭头', r'\leftharpoonup \rightharpoondown', ['↼', '⇁'], ['harpoon'])
check('近似关系', r'a\cong b \simeq c \asymp d \doteq e', ['≅', '≃', '≍', '≐'], ['cong'])
check('偏序prec/succ', r'a\prec b \succ c \sqsubseteq d \sqsupseteq e', ['≺', '≻', '⊑', '⊒'], ['prec'])
check('逻辑关系', r'A\models B \vdash C \bowtie D \ni E', ['⊨', '⊢', '⋈', '∋'], ['models'])
check('同义符号', r'\vee \wedge \lnot \bot \gets', ['∨', '∧', '¬', '⊥', '←'], ['vee'])
check('集合运算符', r'\backslash \amalg \uplus \sqcap \sqcup', ['∖', '⨿', '⊎', '⊓', '⊔'], ['amalg'])
check('大算符', r'\bigodot \bigvee \bigwedge', ['⨀', '⋁', '⋀'], ['bigvee'])
check('三角符号', r'\triangleleft \triangleright \bigtriangleup', ['◁', '▷', '△'], ['triangle'])
check('字形补全', r'\square \checkmark \clubsuit \spadesuit \sharp', ['□', '✓', '♣', '♠', '♯'], ['square'])
check('无点字母', r'\imath \jmath', ['ı', 'ȷ'], ['imath'])
check('组合重音acute/check/mathring', r'\acute{a} \check{b} \mathring{A}', ['<m:acc>', '́', '̌', '̊'], ['acute'])
check('mathsf无衬线', r'\mathsf{A}', ['m:scr m:val="sans-serif"', 'A'], ['mathsf'])
check('hspace间距', r'a\hspace{2em}b', ['a', 'b'], ['hspace'])
check('underbar下横线', r'\underbar{x}', ['<m:bar>', 'm:pos m:val="bottom"'], ['underbar'])
check('标准命令组合', r'\cong A,\ \Downarrow B,\ \mathsf{C}', ['≅', '⇓', 'C'], ['\\'])

# ===== 2026-08-26 补充：极冷门积分/近似符号（opencode 报告建议补齐）=====
check('曲面积分 oiint', r'\oiint_{S} F\cdot dS', ['<m:nary>', '∯', '<m:sub>'], ['oiint'])
check('四重积分 iiiint', r'\iiiint_{V} f', ['∫∫∫∫', '<m:sSub>'], ['iiiint'])
check('多积分 idotsint', r'\idotsint_{V} f', ['∫∫∫', '<m:sSub>'], ['idotsint'])
check('粗近似 thickapprox', r'x \thickapprox y', ['≈'], ['thickapprox'])
check('冷门符号组合', r'\oiint_S F,\ \thickapprox x', ['∯', '≈'], ['\\'])

# ===== 2026-08-26 复查补测：无花括号命令脚标（x_\alpha 不得落回 '?' 占位）=====
check('无花括号命令脚本', r'x_\alpha', ['α', '<m:sSub>'], ['?', '\\'])
check('无花括号符号脚本', r'\sum_\alpha x', ['<m:nary>', 'α'], ['?', '\\'])
check('命令脚本后跟上标', r'x_\beta^{2}', ['β', '<m:sSup>', '<m:sSub>'], ['?', '\\'])

print('\n结果: %d 通过, %d 失败' % (PASS, FAIL))
sys.exit(1 if FAIL else 0)
