# math_work — 数学建模竞赛论文写作技能

> 一个给 AI Agent 用的技能包：把「数学建模竞赛」从**跑出结果**推进到**交得出去的论文**。
> 给一道赛题，AI 按固定结构、固定格式、固定校验流程，产出一份可直接提交的 `.docx` / `.pdf`。

适用于 Claude Code、Cursor、Cline、Roo Code 等支持 `SKILL.md` 约定的 AI 编程工具，也可以被任何能读文件的 AI 直接加载。

---

## 它解决什么问题

数学建模论文最容易栽的地方，从来不是「算不出来」，而是这几类：

| 问题 | 技能里的对策 |
|---|---|
| **AI 爱偷懒跳步** —— 跳过公式渲染、跳过预览、跳过校验就宣布完成 | 第一节就是 **10 条零容忍规则**，违反条款直接写明后果 |
| **公式变纯文本** —— Word 里显示成 `nct+1=Lc\cdot nct` 这种乱码 | `tools/omml_builder.py` 直接把公式字符串构建成 Word 原生 `<m:oMath>`，不经过 Word/WPS COM |
| **排版扣分** —— 首页被摘要挤爆、附录结构不全、三线表画错、表格跨页 | `tools/verify_paper.py` 一键跑 8 个硬门 |
| **源码改对了、交出去的还是旧稿** —— 产物和源码漂移 | 强制「单一事实源」：docx/pdf 只是构建产物，改任何 `.py` 必须重跑整链 |
| **数字凭印象写** —— 正文里的结果和代码算出来的对不上 | `check_numbers` 模板做漂移扫描 + 反杜撰，退出码非 0 即不通过 |

---

## 目录结构

```
math_work/
├── SKILL.md                  # 技能主体：11 节、约 1600 行规范
└── tools/
    ├── omml_builder.py       # 公式 → Word 原生 OMML（核心）
    ├── verify_paper.py       # 论文一键结构校验（8 个硬门）
    ├── check_numbers_template.py  # 数字单一来源 / 反杜撰校验模板
    ├── docx_seq.py           # 解析 docx 段落序列（图注/表注/公式/正文）
    ├── interp_count.py       # 解读短语数 vs 图表数统计
    ├── pdf_chars.py          # PDF 逐页字符统计
    ├── test_omml.py          # OMML 生成器回归测试
    ├── test_nary_sup.py      # ∑/∏ 缺上标检测回归测试
    ├── test_pagination.py    # 分页软门回归测试
    └── test_style.py         # 文风软门 + 硬门回归测试
```

`SKILL.md` 的 11 节：

1. 论文整体结构（固定顺序）
2. 各部分内容要求
3. 格式规范
4. 关键禁止项
5. 最终检查清单
6. 评阅标准（来自竞赛章程）
7. 实战补充规范
8. 工具路径与最终工作流（强制）
9. 既有文档公式修复规范
10. 文风与「降 AI 味」写作规范
11. 正文与实现一致：「说与做」对照

---

## 快速开始

### 方式 A：装成 Claude Code 技能（推荐）

```bash
git clone https://github.com/<你的用户名>/math-work-skill.git
cp -r math-work-skill/math_work ~/.claude/skills/math_work
```

然后在对话里输入 `/math_work` 激活。

### 方式 B：不安装，直接让 AI 读

把 `math_work` 放在任意位置，对 AI 说：

> 读 `<路径>/math_work/SKILL.md`，严格按其中规范执行。题目如下：……

技能内部所有工具引用都是**相对路径**（相对 `SKILL.md` 所在目录），所以放在哪里都能用。

---

## 环境要求

| 项 | 要求 |
|---|---|
| Python | 3.9+ |
| 依赖 | `python-docx`、`pypdf`、`PyMuPDF`、`lxml` |
| 办公软件 | **LibreOffice**（用于转 PDF 与预览；Word / WPS 不参与生成） |

安装依赖：

```bash
pip install python-docx pypdf pymupdf lxml
```

> ⚠️ 用装了上述依赖的那个解释器。若系统里有多个 Python，建议在 `SKILL.md` 第八节的「工具路径」表里把解释器写成绝对路径。

跑一遍回归测试确认环境正常：

```bash
cd math_work
python tools/test_omml.py
python tools/test_nary_sup.py
python tools/test_pagination.py
python tools/test_style.py
```

四项全 PASS（退出码 0）即环境就绪。

---

## 标准工作流

```
求解脚本 solve.py
      ↓  （产出 data/*.pkl）
     make_figs.py        ← 画全部图
      ↓
  build_paper_*.py       ← python-docx 生成 docx，公式走 $$...$$ → OMML
      ↓
   产物 docx
      ↓  （LibreOffice 转档）
   产物 pdf
      ↓
  verify_paper.py 论文.docx 论文.pdf 论文程序目录     ← 8 个硬门
  check_numbers.py                                  ← 数字漂移 / 反杜撰
  LibreOffice 打开预览                               ← 人眼或程序化抽查
```

**硬约束**：改了上面任何一个 `.py`，必须从 `solve.py` 开始重跑整条链，校验只对**本次刚生成**的产物做。只改源码不重建 → 校验全绿也是假绿灯。

---

## 工具速查

### `omml_builder.py`

```python
from omml_builder import build_omath

# 段落内联公式
p = doc.add_paragraph()
p._p.append(build_omath(r"\sum_{i=1}^{n} x_i^2"))
```

不依赖 Word/WPS COM，任何机器都能跑；产出的仍是 Word 原生公式，双击可编辑。

### `verify_paper.py`

```bash
python tools/verify_paper.py 论文.docx 论文.pdf 论文程序目录
```

覆盖：连排检查、解读短语数、摘要结构、OMML 结构、整数图号、引用落点、可复现性、截尾术语。全部 PASS 才算合格。

### `check_numbers_template.py`

复制到论文程序目录、改名 `check_numbers.py`，改写 `derived()` 枚举本题所有结果值：

- Step B 漂移扫描：结果数字被写成字面量 → 退出码 1
- Step C 反杜撰：正文数字在结果集合里找不到 → 退出码 2

---

## 适用范围

✅ 适用：全国大学生数学建模竞赛（CUMCM）、美国大学生数学建模竞赛（MCM/ICM）、各类校赛 / 省赛 / 研赛；要求提交 Word 论文的赛道。

❌ 不适用：要求 LaTeX 排版的赛道；纯代码题；需要赛方专用模板（如特定 `.dotx`）且模板结构差异过大的情况——这时应把模板要求补进 `SKILL.md` 第二节。

**技能不是死的**：`SKILL.md` 开头有「技能自我完善规则」，要求 AI 遇到未覆盖情况时先自行解决、再向用户反馈建议、经同意后才改技能文件。

---

## 贡献

欢迎提 Issue 补充你踩过的坑。改动 `tools/*.py` 后请一并补回归测试（`test_*.py`），否则下次改动会静默退化。

---

## 许可

[MIT](LICENSE) © 2026 zht20070604

本项目仅供学习与技术研究使用，参赛请遵守各竞赛的学术规范。
