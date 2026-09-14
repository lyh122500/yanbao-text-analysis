# -*- coding: utf-8 -*-
"""
问题1：从研报文本中提取分析师对被分析公司的评级
需求来源：《需要对研报文本进行的进一步分析.docx》第 1 条
数据来源：综合分析结果.xlsx 的“研报文本”列

提取思路（详见 README.md“问题1”一节）：
1. 以“评级”为锚点，在锚点前后窗口内查找评级词；
2. 处理评级词位于锚点之后的句式（“下调评级至中性-B”“评级由买入上调至增持”“评级：买入”）；
3. 识别并跳过券商评级标准说明段（连续出现多个评级档位）；
4. 无“评级”锚点时使用保守兜底句式（维持/继续/重申 + 评级词）；
5. 一条文本出现多处评级时，优先取带动作词、靠近“公司简称”、位置靠后的候选。

输出：output/问题1_评级提取结果.xlsx
    列：fordate, 序号, stkcd, 公司简称, 证券公司, 分析师姓名,
        评级_原始, 评级_标准化, 评级动作, 提取证据, 提取状态
"""

import os
import re
from collections import Counter
from pathlib import Path

import openpyxl

# 项目根目录（本脚本位于 problem1_评级提取/ 子文件夹内）
PROJECT_ROOT = Path(__file__).resolve().parent.parent
INPUT_XLSX = PROJECT_ROOT / "综合分析结果.xlsx"
OUTPUT_DIR = PROJECT_ROOT / "problem1_评级提取" / "output"
OUTPUT_XLSX = OUTPUT_DIR / "问题1_评级提取结果.xlsx"

# ---------------------------------------------------------------------------
# 评级词表与正则
# ---------------------------------------------------------------------------
CORE = ["买入", "增持", "推荐", "中性", "持有", "减持", "卖出", "回避", "观望"]
MOD = ["强烈", "谨慎", "审慎"]
SUFFIX = r"(?:-[ABC])?"

RATING_SRC = r"(?:强烈|谨慎|审慎)?" + r"(?:买入|增持|推荐|强推|中性|持有|减持|卖出|回避|观望)" + SUFFIX
# 评级词之后不能紧跟这些词，排除“推荐逻辑”“推荐顺序”等干扰
RATING_TERM = re.compile(RATING_SRC + r"(?!逻辑|顺序)")
ANCHOR = re.compile(r"评级")

ACTION_WORDS = ["维持", "继续", "仍维持", "重申", "上调", "下调",
                "给予", "给与", "给以", "首次", "初次", "暂给", "调整"]

# 兜底句式（文本中无有效“评级”锚点时）：维持/继续/重申 + 评级词
# 增持单独处理：排除股东增持（增持股份/至XX%等）
FALLBACK_PATTERNS = [
    # 排除“继续推荐机构投资者‘买入’”——此处“推荐”是动词，真实评级在后文
    re.compile(r"(维持|继续|仍维持|重申)[“”\"']*(?P<r>强烈推荐|审慎推荐|谨慎推荐|推荐|买入|中性|持有|减持|卖出|回避|观望)(?!机构|投资者|客户|给)"),
    re.compile(r"(维持|继续|仍维持|重申)[“”\"']*(?P<r>增持)(?!至|股份|计划|比例|了|约|达|超|数|万股)"),
    # “强推”是“强烈推荐”的缩写（如“维持强推评级”“继续强推——水井坊”），排除“强推广/强推品种”等非评级用法
    re.compile(r"(维持|继续|仍维持|重申|给定|给予)[“”\"']*(?P<r>强推)(?!广|品种|渠道)"),
    # “给予X”但后面没有“评级”（如“首次给予“推荐”。”“给予“买入”的投资建议”）
    re.compile(r"(继续给予|再次给予|首次给予|给予|给与|给以)[“”\"'\s]*(?P<r>" + RATING_SRC + r")(?!逻辑|顺序|机会|时点|股份)"),
    # “继续推荐机构投资者‘买入’”：推荐/建议 + 机构/投资者 时，推荐是动词，真实评级在其后
    re.compile(r"(?P<a>继续|维持|重申|仍然)?(?:推荐|建议)(?:机构|投资者|客户)+[“”\"'\s]*(?P<r>" + RATING_SRC + r")"),
]

# 文中明确“未给出评级”的句式（暂不给予评级 / 暂无投资评级 / 无投资评级 / 未有评级 等）
NO_RATING_PATTERN = re.compile(r"(没有|未有|暂无|无|不予|不给予|暂不给予|暂不给|暂不调整|不调整|暂不)(投资)?评级")

# 截断句式：原文换列/换行把“评级”一词截断（如“维持“谨慎推荐”评2013上半年……”），
# 要求“评”字后面不是“级/估/价/论/述/审/选/议/注/点”
TRUNCATED_RATING = re.compile(
    r"(维持|继续|重申|给予|给与|给以|首次)[^。；\n]{0,12}?[“”\"']*(?P<r>" + RATING_SRC + r")[”\"' ]*评(?!级|估|价|论|述|审|选|议|注|点)")
# 建议句式：建议(投资者/现价/积极/逢低/可/继续/强烈) + 评级词
SUGGEST_RATING = re.compile(
    r"建议[“”\"'\s投资者现价积极逢低可继续强烈]*?(?P<r>" + RATING_SRC + r")(?!逻辑|顺序)")
# 裸评级句式：句尾的“推荐/买入/强烈买入/坚定推荐……”等非正式评级（如“……，维持强烈推荐。”“强烈买入！”）
BARE_RATING = re.compile(
    r"(?<!不)(?<!难以)(?<!组合)(?<!板块)(?<!重点)"
    r"(?P<r>强烈推荐|审慎推荐|谨慎推荐|强烈买入|强烈增持|坚定推荐|持续推荐|继续推荐|推荐|买入)"
    r"(?=[。！!；\s]|$)")

# 标准化映射：去掉前缀（强烈/谨慎/审慎）与后缀（-A/-B/-C）后的基准档位
def normalize_rating(r: str) -> str:
    r = re.sub(r"-[ABC]$", "", r)
    r = re.sub(r"^(?:强烈|谨慎|审慎)", "", r) or r
    return "推荐" if r in ("强推", "坚定推荐", "持续推荐", "继续推荐") else r

# ---------------------------------------------------------------------------
# 候选提取
# ---------------------------------------------------------------------------
def _clean(text: str) -> str:
    """清洗文本：去掉软连字符等隐形字符，统一引号，合并被换行拆散的“评 级”。"""
    for ch in ["\xad", "﻿", "​", "‌", "‍", "\xa0"]:
        text = text.replace(ch, "")
    text = text.replace("''", "”").replace("``", "“")
    text = re.sub(r"评[ \t\r\n]+级", "评级", text)
    return text.strip()


def _is_scale_description(rest: str) -> bool:
    """评级说明段：评级锚点之后连续出现多个评级档位（如“强烈推荐——……；推荐——……；中性——……”）。"""
    hits = list(RATING_TERM.finditer(rest[:45]))
    return len(hits) >= 2


def _detect_action(ctx: str) -> str:
    """在上下文中找动作词，取离锚点最近的一个。"""
    found = [(m.group(0), m.start()) for w in ACTION_WORDS if (m := re.search(re.escape(w), ctx))]
    if not found:
        return ""
    found.sort(key=lambda x: -x[1])
    return found[0][0]


def _bonus_for_company(company: str, text: str, pos: int) -> bool:
    """评级候选附近是否出现被分析公司的简称（用于多股票研报的归属判定）。"""
    if not company:
        return False
    return company in text[max(0, pos - 40):pos + 15]


def find_candidates(text: str, company: str):
    """返回 [(rating, action, pos, method, evidence, company_bonus), ...]"""
    cands = []
    for m in ANCHOR.finditer(text):
        s = m.start()
        after = text[s + 2:s + 18]
        after_long = text[s + 2:s + 26]   # 至-句式可能跨更长的距离
        before = text[max(0, s - 30):s]

        # 跳过券商评级标准说明的标题行（如“平安证券综合研究所投资评级：”）
        if "综合研究所" in before:
            continue

        # 句式1：评级之后出现“至/到/为 + X”（下调评级至中性-B / 下调评级到“增持” / 评级由买入下调为持有）
        # “为”需排除“原为/认为/作为/成为/因为/视为”等非评级变动用法，
        # 否则“下调至“中性”评级（原为“买入”）”会把被替换掉的旧评级当成新评级
        mm = re.search(
            r"(?:至|到|(?<![原认作成因以视称定行变更尤极较甚本])为)\s*[“”\"'\s]*(?P<r>" + RATING_SRC + r")(?!逻辑|顺序)",
            after_long)
        if mm and not _is_scale_description(after_long[mm.end():]):
            r = mm.group("r")
            pos = s + 2 + mm.start()
            act = _detect_action(text[max(0, s - 25):s]) or "调整"
            ev = text[max(0, pos - 25):pos + len(r) + 12]
            cands.append((r, act, pos, "至-句式", ev, _bonus_for_company(company, text, pos)))
            # 注意：此处不能 continue，同一锚点还要让“锚点前句式”也产出候选，
            # 否则“给予‘买入’评级，继续强烈推荐”会只剩错误的“推荐”候选，
            # 优先级打分无从比较（曾导致误判）。孰优由 METHOD_PRIORITY 裁决。

        # 句式2：评级之后紧跟“：/为/仍为/动作词 + 评级词”（评级：买入 / 评级仍为买入）
        mm = re.match(
            r"^[，,：:为\s]*(?:维持|继续|仍维持|重申|上调|下调|给予|给与|给以|首次|暂给|仍为)?[“”\"'\s]*(?P<r>" + RATING_SRC + r")(?!逻辑|顺序)",
            after)
        if mm and not _is_scale_description(after[mm.end():]):
            r = mm.group("r")
            pos = s + 2 + mm.start()
            # 动作词优先取锚点与评级词之间（“评级继续推荐”→继续），
            # 而非锚点之前（“给予‘买入’评级，继续推荐”→不应取“给予”）
            act = _detect_action(after[:mm.start()]) or _detect_action(text[max(0, s - 25):s])
            ev = text[max(0, pos - 25):pos + len(r) + 12]
            cands.append((r, act, pos, "锚点后句式", ev, _bonus_for_company(company, text, pos)))
            # 同上，不能 continue，需让“锚点前句式”也能产出候选

        # 句式3：评级词位于锚点之前（维持“买入”评级 / 买入-A 的投资评级）
        # 注意：窗口内**所有**评级词都要产出候选，不能只取最后一个——
        # “维持对中行A 股谨慎买入和H 股买入评级不变”里，只取最后一个会丢掉
        # A 股的“谨慎买入”，导致误取 H 股档位。孰优由打分裁决。
        before_start = max(0, s - 30)
        for mm3 in RATING_TERM.finditer(before):
            # 评级词与锚点之间不能有句号/分号/换行（跨句无效）
            if re.search(r"[。；\n]", before[mm3.end():]):
                continue
            r = mm3.group(0)
            pos = before_start + mm3.start()
            act = _detect_action(before[max(0, mm3.start() - 25):mm3.start()])
            ev = text[max(0, pos - 25):pos + len(r) + 12]
            cands.append((r, act, pos, "锚点前句式", ev, _bonus_for_company(company, text, pos)))

    # 句式4：全文扫描“加入X名单”（如“我们将其加入买入名单”）
    for m in re.finditer(r"加入[“”\"']*(?P<r>" + RATING_SRC + r")(?!逻辑|顺序)名单", text):
        r = m.group("r")
        pos = m.start() + 2
        ev = text[max(0, pos - 25):pos + len(r) + 12]
        cands.append((r, "纳入", pos, "加入名单句式", ev, _bonus_for_company(company, text, pos)))

    # 句式5：全文扫描截断句式（“评级”一词被换列截断）
    for m in TRUNCATED_RATING.finditer(text):
        r = m.group("r")
        pos = m.start()
        ev = text[max(0, pos - 25):pos + len(m.group(0)) + 12]
        cands.append((r, m.group(1), pos, "截断评级句式", ev, _bonus_for_company(company, text, pos)))

    # 句式6：全文扫描“建议X”（建议买入 / 建议增持 / 建议投资者买入）
    for m in SUGGEST_RATING.finditer(text):
        r = m.group("r")
        pos = m.start()
        ev = text[max(0, pos - 25):pos + len(m.group(0)) + 12]
        cands.append((r, "建议", pos, "建议句式", ev, _bonus_for_company(company, text, pos)))

    # 句式7：全文扫描句尾裸评级词（“……，维持强烈推荐。”“强烈买入！”）
    for m in BARE_RATING.finditer(text):
        r = m.group("r")
        pos = m.start()
        ev = text[max(0, pos - 25):pos + len(r) + 12]
        cands.append((r, "", pos, "裸评级句式", ev, _bonus_for_company(company, text, pos)))
    return cands


def find_fallback(text: str, company: str):
    """无“评级”锚点时的兜底：维持/继续/重申/给予 + 评级词。返回 [(rating, action, pos, method, evidence), ...]"""
    cands = []
    for pat in FALLBACK_PATTERNS:
        for m in pat.finditer(text):
            r = m.group("r")
            pos = m.start()
            ev = text[max(0, pos - 25):pos + len(m.group(0)) + 12]
            cands.append((r, m.group(1), pos, "兜底句式", ev, _bonus_for_company(company, text, pos)))
    return cands


# 句式可信度分层：显式带“评级”锚点的句式最可靠，软句式（建议/裸评级/兜底）次之。
# 注意：分层优先级必须高于公司简称加分，否则会出现“积极推荐”这类软句式
# 因靠近公司名而压过正文中明确写的“维持‘买入’评级”的情况（曾导致误判）。
METHOD_PRIORITY = {
    "至-句式": 60,        # 下调评级至X / 评级由买入下调为持有 —— 最明确
    "锚点前句式": 50,      # 维持“买入”评级 —— 最常见最可靠
    # 锚点后必须低于锚点前：“给予‘买入’评级，继续强烈推荐”里，评级是买入，
    # “，继续强烈推荐”只是分析师表热情的新分句，不能当评级（曾导致误判）
    "锚点后句式": 45,      # 投资评级：增持
    "截断评级句式": 44,    # 维持“谨慎推荐”评[被截断]
    "加入名单句式": 40,    # 加入买入名单
    "建议句式": 30,        # 建议买入
    "兜底句式": 25,        # 维持推荐。
    "裸评级句式": 20,      # ……，强烈推荐。
}


def _market_preference(ev: str, rating_len: int) -> int:
    """A/H 股双评级时偏向 A 股（本数据集为 A 股研报，stkcd 是 A 股代码）。

    证据串以评级词为中心：前 25 字 + 评级词 + 后 12 字。取**离评级词最近**的
    市场标记判定归属——不能简单地“窗口里有A股就加分”，因为
    “维持对A 股买入评级和H 股持有评级”里两个候选的窗口都含 A/H 股。
    """
    pre, post = ev[:25], ev[25 + rating_len:25 + rating_len + 12]
    best = None  # (距评级词的距离, 'A'/'B'/'H')
    for m in re.finditer(r"([ABH])\s*股", pre):       # 往前找，越靠右越近
        d = len(pre) - m.end()
        if best is None or d < best[0]:
            best = (d, m.group(1))
    for m in re.finditer(r"([ABH])\s*股", post):      # 往后找，越靠左越近
        if best is None or m.start() < best[0]:
            best = (m.start(), m.group(1))
    if best is None:
        return 0
    return 8 if best[1] == "A" else -8


def pick_best(cands):
    """打分选最优候选：句式可信度 + 动作词 +10 + 公司简称邻近 +15 + A股偏好 ±8 + 位置靠后。"""
    if not cands:
        return None
    best = max(
        cands,
        key=lambda c: METHOD_PRIORITY.get(c[3], 10) + (10 if c[1] else 0) + (15 if c[5] else 0)
        + _market_preference(c[4], len(c[0])) + c[2] / 1000,
    )
    return best


def extract_rating(text: str, company: str):
    """返回 (原始评级, 标准化评级, 动作, 证据, 状态)"""
    if not text or not text.strip():
        return ("", "", "", "", "文本为空")
    text = _clean(text)
    cands = find_candidates(text, company)
    method = "锚点"
    if not cands:
        cands = find_fallback(text, company)
        method = "兜底"
    best = pick_best(cands)
    if best is None:
        if NO_RATING_PATTERN.search(text):
            return ("", "", "", "", "未给出评级（文中明确未给）")
        return ("", "", "", "", "未提取到评级")
    rating, action, pos, m, ev, _ = best
    return (rating, normalize_rating(rating), action, ev, f"已提取（{method}/{m}）")


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def main():
    if not INPUT_XLSX.exists():
        raise FileNotFoundError(f"未找到输入文件 {INPUT_XLSX}")

    wb = openpyxl.load_workbook(INPUT_XLSX, read_only=True)
    ws = wb["Sheet1"]

    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "评级提取结果"
    headers = ["fordate", "序号", "stkcd", "公司简称", "证券公司", "分析师姓名",
               "评级_原始", "评级_标准化", "评级动作", "提取证据", "提取状态"]
    out_ws.append(headers)

    dist = Counter()
    status = Counter()
    n = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        n += 1
        fordate, seq, stkcd, company, broker, text = row[0], row[1], row[2], row[3], row[4], row[5]
        analyst = row[6]
        raw, norm, action, ev, st = extract_rating(text or "", company or "")
        out_ws.append([fordate, seq, stkcd, company, broker, analyst, raw, norm, action, ev, st])
        if norm:
            dist[norm] += 1
        status[st] += 1
        if n % 5000 == 0:
            print(f"  已处理 {n} 行 ...")

    wb.close()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out_wb.save(OUTPUT_XLSX)

    # 汇总
    print(f"\n处理完成：共 {n} 行，结果写入 {OUTPUT_XLSX}\n")
    print("提取状态分布：")
    for k, v in status.most_common():
        print(f"  {k}: {v} ({v / n * 100:.2f}%)")
    print("\n标准化评级分布（已提取部分）：")
    total_extracted = sum(dist.values())
    for k, v in dist.most_common():
        print(f"  {k}: {v} ({v / total_extracted * 100:.2f}%)")


if __name__ == "__main__":
    main()
