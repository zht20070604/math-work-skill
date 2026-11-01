# -*- coding: utf-8 -*-
"""test_style.py —— 文风软门 + 新增硬门回归测试（SKILL 第十节/第九节配套验证代码）

用法： python test_style.py      # 无参数，全 PASS 退出码 0，任一 FAIL 退出码 1

判定目标（"好论文零误报、模板痕迹全命中"）：
  A) 文风软门（check_style）
    1) 干净样例（已清理 D 论文语体摘录，含"如图 3 所示"带图号正例）→ 0 条建议
    2) 症状 R1 裸图导语（"其效果如图所示"）→ ≥1 命中
    3) 症状 R2 强调式开句堆叠（观察图×6）→ ≥1 命中
    4) 症状 R3 段尾拔高宣传句 → ≥1 命中
  B) 硬门机检
    5) check_float_figref：干净无小数图号 → 0 处；症状"第1.0图/图4.0"→ ≥1 命中
    6) check_citation：干净正文引用全部 [k] → 0 missing；症状正文未引用任何 [k] → ≥1 missing
    7) check_terminology：干净"截断前位置/尺度参数"→ 0 处；症状"截尾分布的均值"→ ≥1 命中
    8) check_freshness：源码比产物新 → 提示"忘了重建"；产物比源码新 → 提示"mtime 不可靠，
       请旧句回归核对"（软 INFO，确保改动 SKILL 规则 10 后此提示不被误改坏）

任何 Claude 改动 verify_paper / SKILL 第十节、第九节后都应重跑本文件，
保证既不误伤已清理论文、又能抓住模板痕迹与审计发现的缺陷。
"""
import sys, os
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx import Document
import verify_paper


def _doc(*texts):
    d = Document()
    for t in texts:
        d.add_paragraph(t)
    return d


def scan(*texts):
    """把文本列表当正文段落喂给与 verify_paper 同路径的 check_style。"""
    return verify_paper.check_style(_doc(*texts))


def expect(cond, msg):
    print(("[PASS] " if cond else "[FAIL] ") + msg)
    return 0 if cond else 1


# ---------- 干净样例：应 0 建议（D 论文批改后语体摘录） ----------
CLEAN = [
    "矿井突水一旦发生，高压水流沿巷道网络迅速推进，能否在通道被水阻断前把矿工撤出，"
    "取决于水流何时到达各巷道、矿工应沿哪条路线撤离。",
    "针对问题 1，把水流漫延看成巷道网络上的传播过程：先在所在巷道打断并引入节点，"
    "再按水流仅沿水平或下行方向传播把网络转为有向，用最短路递推求得各端点的到达时刻。",
    "由图 3 可见，附件 1 的水平网络可漫延至全部端点，远端巷道到达时刻明显滞后，"
    "源于节点均分流量使流量沿传播路径指数级衰减。",
    "到达时刻方面，附件 2 受高程约束仅 41 个低位端点有水流到达；"
    "双源合并后两源分处不同高程盆地、互不相通，其中 B 源独有 13 个。",
    "由表 3 可知，附件 1 在调整时刻 5 min 时，矿工 1、3 转走 E2 出口，"
    "到达时刻由 15.88 min、18.10 min 分别延至 16.45 min、18.50 min；"
    "矿工 2 仍走 E1，由 12.17 min 延至 15.23 min。",
    "推进速度由过水断面把分支流量换算而来，水流到达时刻如图 3 所示。",
    "从图 12 的分布看，双源合并后水流仍覆盖全部端点、整体提前；"
    "同一巷道的水深不能按两源流量直接求和，须按各自进水时刻分时累计。",
    "灵敏度分析显示，到达时刻对突水流量与涉水通行阈值的变化不敏感，"
    "对通知延迟则近似等比例推迟。",
]

# ---------- 症状 R1：裸图导语（无图号） ----------
SYM_R1 = [
    "将两套网络的水流到达结果可视化，其效果如图所示。",
    "水流推进的整体形态见下图。",
]

# ---------- 症状 R2：强调式开句机械堆叠（观察图×6） ----------
SYM_R2 = [
    "观察图 4 可知，曲线在 5 min 后趋于平缓。",
    "观察图 6 可知，附件 1 中矿工 2 的路径被水流截断。",
    "观察图 8 可知，末端巷道到达时刻明显晚于近端。",
    "观察图 10 可知，双源叠加后波前整体提前。",
    "观察图 12 可知，B 源独有端点集中在低位。",
    "观察图 14 可知，绕行路径从涉水位置重新接入干巷道。",
]

# ---------- 症状 R3：段尾拔高宣传句 ----------
SYM_R3 = [
    "该重规划策略兼顾了安全性与方案的可执行性。",
    "该框架具备较强的跨场景迁移能力。",
    "整体方案充分验证了模型的合理性与实用性。",
]


# ---------- 硬门 5) 整数图号 ----------
# 干净：正文与表格含正常整数图号/表号引用（"如图 3 所示""见表 2"）→ 0 命中
CLEAN_FIGREF = [
    "如图 3 所示，曲线在 5 min 后趋于平缓。",
    "结果汇总见表 2。",
]
# 症状：审计在实论文中抓到的小数图号（"第 1.0 图"数字在前 / "图 4.0"数字在后两种次序）
SYM_FIGREF = [
    "求解的总体结构见第1.0 图。",
    "命中概率分布见第4.0 图。",
    "变化见第10.0 图与表4。",
]


# ---------- 硬门 6) 引用落点 ----------
# 干净：正文实际引用了每条参考文献 [k]
CLEAN_CITE = [
    "摘要内容",
    "正文引用文献[1]的方法与文献[2]的结论",
]
REF_HEAD = ["参考文献"]   # 定位参考文献节的标题段
REF_FULL = ["[1] 张三，书名，出版社，2024。", "[2] 李四，论文名，杂志，2024。"]
# 症状：参考文献节有 [k] 但正文完全没引用 → missing 应报出全部
SYM_CITE_BODY = ["摘要内容", "正文没有出现任何方括号引用。"]


def main():
    bad = 0

    clean = scan(*CLEAN)
    bad += expect(len(clean) == 0, "干净样例 0 条建议（实际 %d 条: %s）" % (len(clean), clean))

    sym_r1 = scan(*SYM_R1)
    n1 = sum(1 for i in sym_r1 if i[1].startswith("R1"))
    bad += expect(n1 >= 1, "症状 R1 裸图导语 ≥1 命中（实际 %d，建议 %s）" % (n1, sym_r1))

    sym_r2 = scan(*SYM_R2)
    n2 = sum(1 for i in sym_r2 if i[1].startswith("R2"))
    bad += expect(n2 >= 1, "症状 R2 强调式开句堆叠 ≥1 命中（实际 %d，建议 %s）" % (n2, sym_r2))

    sym_r3 = scan(*SYM_R3)
    n3 = sum(1 for i in sym_r3 if i[1].startswith("R3"))
    bad += expect(n3 >= 1, "症状 R3 段尾拔高 ≥1 命中（实际 %d，建议 %s）" % (n3, sym_r3))

    # 交叉确认：症状句子里不允许互相串规则（R1/R2 各段本身不该误伤 R3 等）
    extra_r1 = sum(1 for i in sym_r3 if i[1].startswith("R1"))
    extra_r3 = sum(1 for i in sym_r2 if i[1].startswith("R3"))
    bad += expect(extra_r1 == 0, "症状 R3 文本不误报 R1（实际 %d）" % extra_r1)
    bad += expect(extra_r3 == 0, "症状 R2 文本不误报 R3（实际 %d）" % extra_r3)

    # ---------- 硬门 5) 整数图号 ----------
    f_clean = verify_paper.check_float_figref(_doc(*CLEAN_FIGREF))
    bad += expect(len(f_clean) == 0, "整数图号：干净样例 0 处（实际 %r）" % f_clean)
    f_sym = verify_paper.check_float_figref(_doc(*SYM_FIGREF))
    bad += expect(len(f_sym) >= 1, "整数图号：症状样例 ≥1 命中（实际 %d: %r）" % (len(f_sym), f_sym))

    # ---------- 硬门 6) 引用落点 ----------
    # 干净：正文段 + 参考文献节（标题段 + 条目段），正文已引用 [1][3] → 应为 (refs, [])
    d_clean = _doc(*CLEAN_CITE, *REF_HEAD, *REF_FULL)
    refs_ok, miss_ok = verify_paper.check_citation(d_clean)
    bad += expect(len(miss_ok) == 0, "引用落点：干净样例正文已引用全部（missing=%r）" % miss_ok)
    # 症状：正文未引用任何 [k] → missing 应等于全部 refs
    d_sym = _doc(*SYM_CITE_BODY, *REF_HEAD, *REF_FULL)
    refs_sym, miss_sym = verify_paper.check_citation(d_sym)
    bad += expect(len(miss_sym) >= 1 and len(miss_sym) == len(refs_sym),
                  "引用落点：症状样例正文未引用全部（refs=%s missing=%s）" % (sorted(refs_sym), miss_sym))

    # ---------- 硬门 7) 截尾术语 ----------
    # 干净：正文/表格用"截断前位置/尺度参数" → 0 命中
    TERM_CLEAN = ["h0 为截断前正态分布的位置参数，σz 为截断前尺度参数。"]
    # 症状：误称"截尾分布的均值/标准差" → ≥1 命中
    TERM_SYM = ["h0 为截尾分布的均值，σz 为截尾分布的标准差。"]
    t_clean = verify_paper.check_terminology(_doc(*TERM_CLEAN))
    bad += expect(len(t_clean) == 0, "截尾术语：干净样例 0 处（实际 %r）" % t_clean)
    t_sym = verify_paper.check_terminology(_doc(*TERM_SYM))
    bad += expect(len(t_sym) >= 1, "截尾术语：症状样例 ≥1 命中（实际 %d: %r）" % (len(t_sym), t_sym))

    # ---------- 硬门 11) 产物新鲜度（软 INFO） ----------
    # 场景 A：源码比产物新 → 应提示"可能忘了重建"
    # 场景 B：产物比源码新 → mtime 看似新，但仍提示"mtime 不可靠，请旧句回归核对"
    import tempfile, time
    tmp = tempfile.mkdtemp()
    docx_p = os.path.join(tmp, "论文.docx")
    open(docx_p, "w").close()
    src_p = os.path.join(tmp, "build_paper_D.py")
    open(src_p, "w").close()
    # 让源码 mtime 比 docx 新
    doc_t = time.time() - 200
    src_t = time.time() - 100
    os.utime(docx_p, (doc_t, doc_t))
    os.utime(src_p, (src_t, src_t))
    fres_a = verify_paper.check_freshness(docx_p, tmp)
    bad += expect(any("忘了重建" in s for s in fres_a),
                  "新鲜度：源码比产物新 → 提示先重建（实际 %r）" % fres_a)
    # 反传：docx 比源码新 → 不报"忘了重建"，改报 mtime 不可靠
    os.utime(docx_p, (src_t + 100, src_t + 100))
    fres_b = verify_paper.check_freshness(docx_p, tmp)
    bad += expect(any("mtime 不可靠" in s for s in fres_b),
                  "新鲜度：产物比源码新 → 提示 mtime 不可靠、回归核对（实际 %r）" % fres_b)

    if bad:
        print("\n结果：FAIL（%d 项未通过）" % bad)
        return 1
    print("\n结果：全部 PASS（退出码 0）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
