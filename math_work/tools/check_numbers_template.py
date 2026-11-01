# -*- coding: utf-8 -*-
"""
check_numbers 模板 —— "结果数字单一来源 + 反杜撰"自动校验（替代人肉筛查）

【用法（每篇论文都做一次）】
  1. 把本文件 复制到论文程序目录（与你的 build_paper_X.py、门面 同目录），改名 check_numbers.py；
  2. 改写下方 derived()：把本题"正文会出现的结果数字"全部从门面取出来 _add 进去（见函数内注释与示例）；
  3. 运行：python check_numbers.py [build_paper_X.py 路径]
       退出码 0 = 通过
       退出码 1 = 漂移：正文脚本里出现结果数字的字面量（应由门面取数，禁止写回）
       退出码 2 = 疑似杜撰：结果语境数字不在 结果集合∪允许集合（先查结果文件，查不到反馈用户）

三段式：
  Step A  结果集合 —— 全部从门面实时调用生成，绝不手抄
  Step B  漂移扫描 —— 结果值被写成字面量 → 退出码 1
  Step C  反杜撰扫描 —— 带单位/结果名词的数字不在白名单 → 退出码 2

本文件是 skill 的唯一权威扫描引擎模板：扫描逻辑（Step B/C/main）照抄不动，
只允许改 derived() 与 允许集合/结果名词（见文内"每篇论文按本题调整"标注）。
禁止从旧任务目录复制旧版 check_numbers.py —— 旧版枚举的是旧题结果集合，会漏检本题。
"""
import os, re, sys
sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

# ================= SRC：被扫描的正文生成脚本 =================
_candidates = [f for f in os.listdir(BASE)
               if re.fullmatch(r"build_paper_.*\.py", f)]
SRC = (sys.argv[1] if len(sys.argv) > 1
       else os.path.join(BASE, _candidates[0]) if len(_candidates) == 1
       else sys.exit("用法：python check_numbers.py build_paper_X.py（论文目录应恰好一个 build_paper_*.py）"))

import paper_results as R          # ← 你的门面模块名（结果单一数据源，与本文档同目录）

# ================= Step A：结果集合 =================
DERIVED_STR = {}   # 显示串 -> {tag,...}；tag="" 表示该值可裸匹配
DERIVED_F = set()  # 浮点白名单（Step C）
DERIVED_I = set()  # 整数白名单（Step C）


def _add(v, spec="%.2f", tag=None):
    """v: 结果原始值（必须来自门面访问器）；spec: 正文显示格式；tag: 该值后必须紧随的单位正则。
    小整数/易冲突值（如覆盖数、延迟分钟数、13/28/41 之类）务必给 tag（如 r"\\s*个"、r"\\s*min"），
    按"值+单位"精确匹配，避免误伤 `图13`、`（13）`、`4 m/s` 这类合法文本。"""
    if v is None:
        return
    s = spec % v
    f = float(s)
    if "." in s:
        DERIVED_F.add(f)
    else:
        DERIVED_I.add(int(f))
    DERIVED_STR.setdefault(s, set()).add(tag or "")


def derived():
    """Step A 结果集合枚举 —— 每篇论文按本题改写此处（其余照抄）。
    把"正文里会出现的结果数字"逐项 _add；显示单位写进正文的是什么格式，spec 就怎么写。
    示例（换成本题真实访问器，数值一律 R.xxx 取，禁止手写）：
        for net in ("附件1", "附件2"):
            _add(R.n_pts(net), "%d", r"\s*个")            # 端点计数
            _add(R.n_tunnels(net), "%d", r"\s*条")        # 巷道计数
            _add(R.q1_reached(net), "%d", r"\s*个")       # 覆盖端点
            for i in (1, 2, 3):
                _add(R.q2_arrival(net, i), "%.2f")        # 第 i 个矿工到达时刻(min)
        # 易与图/表号冲突的小整数按"值+单位"匹配：
        _add(R.delay("附件1"), "%.0f", r"\s*min")         # 延迟分钟
    """
    raise NotImplementedError(
        "改写 derived()：把本题结果集合从门面枚举进 _add（模板注释里有示例）。")


# ============ 允许集合（Step C 白名单，每篇论文按本题"问题给定常量+叙述圆整"增删） ============
ALLOW_F = {0.1, 0.2, 0.3, 0.4, 0.5, 1.0, 1.2, 2.0, 3.0, 4.0, 5.0, 6.0, 10.0,
           12.0, 20.0, 24.0, 30.0, 36.0, 60.0, 120.0, 240.0, 1440.0}
ALLOW_I = {1, 2, 3, 4, 5, 6, 10, 12, 20, 24, 30, 36, 60, 120, 240, 1440}

# ============ 结果名词（Step C 语境判定；每篇论文按本题"结果单位/名词"增删） ============
# min 需词边界（避开 m^3/min、变量名 q1_min( 等里的 min）
# m   仅当其后随逗号/引号（避免误触 min/max 等变量名）
NOUN = re.compile(r'(?:'
                  r'(?<![A-Za-z_/])min(?=[\s,，\'`)}一-鿿])|分钟'
                  r'|\s*天|\s*小时'
                  r'|名矿工|个矿工|个端点|个低位端点|条巷道'
                  r'|%'
                  r'|\s*m/s|\s*m\^{3}|\s*m\^{2}|\s*m\^3|\s*m\^2'
                  r'|\s*m(?=[,，\'`]))')

# ================= Step B：漂移扫描 =================
def _b_patterns():
    pats = []
    for v, tags in DERIVED_STR.items():
        esc = re.escape(v)
        if "" in tags:
            pats.append((re.compile(r'(?<![\d.])%s(?![\d.])' % esc), v))  # 数值唯一，裸匹配
        for tag in tags:
            if tag:
                # 小整数/允许值：要求后随其单位，避免误伤 图13、（13）、4 m/s、2 小时 等
                pats.append((re.compile(r'(?<![\d.])%s(?![\d.])%s' % (esc, tag)), v))
    return pats


def scan_drift(src):
    hits = []
    for pat, v in _b_patterns():
        for m in pat.finditer(src):
            line = src.count("\n", 0, m.start()) + 1
            hits.append((line, v, src[m.start():m.end()]))
    return hits


# ================= Step C：反杜撰扫描 =================
TOKEN = re.compile(r'(?<![\d.])(\d+(?:\.\d+)?)(?![\d.])')
# 数字紧邻这些前缀即视为结构标识（节点号 P0xxx、图/表/附录编号、清单号、变量名），非结果字面量
_EXEMPT_PRE = set("_^{") | set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz") \
    | set("图表附录第问题件（([【")


def _is_exempt(src, m):
    if m.start() == 0:
        return False
    return src[m.start() - 1] in _EXEMPT_PRE


def _result_window(src, m):
    """数字后紧跟的内容，截断到下一个数字处（单位必紧邻数字；不跨数抓后面句子的单位）"""
    win = src[m.end():m.end() + 18]
    nxt = re.search(r'\d', win)
    if nxt:
        win = win[:nxt.start()]
    return win


def scan_fabrication(src):
    hits = []
    for m in TOKEN.finditer(src):
        v = m.group(1)
        if _is_exempt(src, m):
            continue
        if not NOUN.search(_result_window(src, m)):
            continue
        f = float(v)
        if "." in v:
            ok = f in DERIVED_F or f in ALLOW_F
        else:
            ok = int(f) in DERIVED_I or int(f) in ALLOW_I
        if not ok:
            line = src.count("\n", 0, m.start()) + 1
            hits.append((line, v, src[m.start():m.end() + 14]))
    return hits


# ================= 主流程 =================
def main():
    derived()  # Step A：先由门面生成结果集合；改写漏了会在这里直接报 NotImplementedError
    R.check()  # 先确认结果文件键齐全，避免"拿旧结果校验新脚本"
    src = open(SRC, encoding="utf-8").read()
    print("[check_numbers] 目标:", os.path.basename(SRC))
    print("[check_numbers] 结果集合 %d 项（%d 浮点 / %d 整数）"
          % (len(DERIVED_STR), len(DERIVED_F), len(DERIVED_I)))

    drift = scan_drift(src)
    fab = scan_fabrication(src)

    if drift:
        print("\n[Step B] 漂移扫描 FAIL —— 以下结果数字被写成了字面量（应改用门面取数）：")
        for line, v, ctx in drift:
            print("  L%-4d %-10s ...%r..." % (line, v, ctx))
        sys.exit(1)
    print("[Step B] 漂移扫描通过：结果集合无字面量残留")

    if fab:
        print("\n[Step C] 反杜撰扫描 FAIL —— 以下数字带结果单位却不在白名单"
              "（先查结果文件，查不到请反馈用户，禁止编造/改白名单放行）：")
        for line, v, ctx in fab:
            print("  L%-4d %-10s ...%r..." % (line, v, ctx))
        sys.exit(2)
    print("[Step C] 反杜撰扫描通过：结果语境数字均在白名单")

    print("\n[check_numbers] PASS（退出码 0）")


if __name__ == "__main__":
    main()
