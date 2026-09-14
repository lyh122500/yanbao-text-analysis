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
    re.compile(r"(维持|继续|仍维持|重申)[“”\"']*(?P<r>强烈推荐|审慎推荐|谨慎推荐|推荐|买入|中性|持有|减持|卖出|回避|观望)"),
    re.compile(r"(维持|继续|仍维持|重申)[“”\"']*(?P<r>增持)(?!至|股份|计划|比例|了|约|达|超|数|万股)"),
    # “强推”是“强烈推荐”的缩写（如“维持强推评级”“继续强推——水井坊”），排除“强推广/强推品种”等非评级用法
    re.compile(r"(维持|继续|仍维持|重申|给定|给予)[“”\"']*(?P<r>强推)(?!广|品种|渠道)"),
    # “给予X”但后面没有“评级”（如“首次给予“推荐”。”“给予“买入”的投资建议”）
    re.compile(r"(继续给予|再次给予|首次给予|给予|给与|给以)[“”\"'\s]*(?P<r>" + RATING_SRC + r")(?!逻辑|顺序|机会|时点|股份)"),
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
        mm = re.search(r"(?:至|到|为)\s*[“”\"'\s]*(?P<r>" + RATING_SRC + r")(?!逻辑|顺序)", after_long)
        if mm and not _is_scale_description(after_long[mm.end():]):
            r = mm.group("r")
            pos = s + 2 + mm.start()
            act = _detect_action(text[max(0, s - 25):s]) or "调整"
            ev = text[max(0, pos - 25):pos + len(r) + 12]
            cands.append((r, act, pos, "至-句式", ev, _bonus_for_company(company, text, pos)))
            continue

        # 句式2：评级之后紧跟“：/为/仍为/动作词 + 评级词”（评级：买入 / 评级仍为买入）
        mm = re.match(
            r"^[，,：:为\s]*(?:维持|继续|仍维持|重申|上调|下调|给予|给与|给以|首次|暂给|仍为)?[“”\"'\s]*(?P<r>" + RATING_SRC + r")(?!逻辑|顺序)",
            after)
        if mm and not _is_scale_description(after[mm.end():]):
            r = mm.group("r")
            pos = s + 2 + mm.start()
            act = _detect_action(text[max(0, s - 25):s]) or _detect_action(after[:mm.start()])
            ev = text[max(0, pos - 25):pos + len(r) + 12]
            cands.append((r, act, pos, "锚点后句式", ev, _bonus_for_company(company, text, pos)))
            continue

        # 句式3：评级词位于锚点之前（维持“买入”评级 / 买入-A 的投资评级）
        matches = list(RATING_TERM.finditer(before))
        if matches:
            last = matches[-1]
            between = before[last.end():]
            # 评级词与锚点之间不能有句号/分号/换行（跨句无效）
            if not re.search(r"[。；\n]", between):
                r = last.group(0)
                pos = s - 30 + last.start()
                act = _detect_action(before[max(0, last.start() - 25):last.start()])
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


def pick_best(cands):
    """打分选最优候选：动作 +30，至-句式 +20，公司简称 +40，位置靠后 + pos/1000。"""
    if not cands:
        return None
    best = max(
        cands,
        key=lambda c: (30 if c[1] else 0) + (20 if c[3] == "至-句式" else 0) + (40 if c[5] else 0) + c[2] / 1000,
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
