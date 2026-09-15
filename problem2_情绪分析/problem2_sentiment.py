# -*- coding: utf-8 -*-
"""
问题2：研报文本的情绪分析
需求来源：《需要对研报文本进行的进一步分析.docx》第 2 条
数据来源：综合分析结果.xlsx 的“研报文本”列
词典来源：中文金融情感词典（姜富伟等，2020），见 dict/ 目录与 README

方法：jieba 分词（把词典词加入用户词典，避免被切碎）→ 逐词匹配词典 →
      统计正面/负面词频 → 计算净值、比例、密度等指标。

同时输出三套口径，由使用者按研究设计决定用哪套：
  1. 原始计数——忠于词典的词袋计数，与姜富伟等原论文口径一致；
  2. 否定调整——“不理想”“未实现”等否定式会把极性翻过来。
     语料中“否定词+正面词典词”共 3132 次、覆盖 13.8% 的研报，
     不做处理会把“不理想”记成正向；
  3. 正文口径——剔除例行的“风险提示”段。语料中“风险”二字的出现
     有 93.3% 落在该段内（15823/16964 次），是强制披露而非分析师观点。
     按 <?> 分段删除，不从头截断——风险段并非总在末尾（见 remove_risk_section）。

三套口径都剔除了落在“公司简称出现位置”内的命中（如“沧州明珠”的“明珠”），
详见 _count。词典中的正负重叠词（21 个）也已剔除。

输出：output/问题2_情绪分析结果.xlsx
"""

import os
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import jieba
import openpyxl

PROJECT_ROOT = Path(__file__).resolve().parent.parent
INPUT_XLSX = PROJECT_ROOT / "综合分析结果.xlsx"
DICT_XLSX = Path(__file__).resolve().parent / "dict" / "中文金融情感词典_姜富伟等(2020).xlsx"
OUTPUT_DIR = Path(__file__).resolve().parent / "output"
OUTPUT_XLSX = OUTPUT_DIR / "问题2_情绪分析结果.xlsx"

# 否定词：出现在情绪词之前（最多 NEG_WINDOW 个词以内）时翻转极性。
# 采用最长匹配，故“难以/无法”等双字词要排在单字前。
NEGATORS = [
    "没有", "难以", "无法", "并非", "并没", "并没有", "未能", "未曾", "不予",
    "不足", "不太", "不够", "不曾", "未见", "未现", "不复", "不宜", "不良",
    "不", "未", "无", "非", "没", "莫", "勿",
]
NEG_WINDOW = 3          # 向前看几个 token
MAX_HIT_TERMS = 8       # 输出里保留的命中词个数


# ---------------------------------------------------------------------------
# 词典
# ---------------------------------------------------------------------------
def load_dictionary(path):
    """返回 (正面词集合, 负面词集合, 正负重叠词集合)。

    词典 xlsx 有两张表：positive / negative，条目带尾部换行需清洗。
    正负表有 21 个词重叠（如“激进”“出乎意料”），语义歧义，
    计入任一侧都会引入噪声，故从两侧同时剔除。
    """
    if not path.exists():
        raise FileNotFoundError(
            f"未找到词典 {path}\n"
            f"请从 https://github.com/MengLingchao/Chinese_financial_sentiment_dictionary "
            f"下载《中文金融情感词典_姜富伟等(2020).xlsx》放入 dict/ 目录"
        )
    wb = openpyxl.load_workbook(path, read_only=True)
    if "positive" not in wb.sheetnames or "negative" not in wb.sheetnames:
        raise ValueError(f"词典格式不符，期望含 positive/negative 两张表，实际 {wb.sheetnames}")

    def _load(sheet):
        out = set()
        for row in wb[sheet].iter_rows(min_row=2, values_only=True):
            if row[0]:
                w = str(row[0]).strip().replace("\n", "").replace(" ", "")
                if w:
                    out.add(w)
        return out

    pos, neg = _load("positive"), _load("negative")
    ambiguous = pos & neg
    wb.close()
    return pos - ambiguous, neg - ambiguous, ambiguous


def register_to_jieba(words):
    """把词典词加入 jieba 用户词典。

    不注册时“现金流充沛”可能被切成“充沛”以外的形态，注册后词典词优先成词，
    匹配率显著提高（实测“下滑/下降/充沛”等本就能切对，注册主要保证长词如
    “不及预期”不被切碎）。
    """
    for w in words:
        jieba.add_word(w, freq=200_000)


# ---------------------------------------------------------------------------
# 文本清洗（与问题1保持一致的归一化规则）
# ---------------------------------------------------------------------------
def clean_text(text: str) -> str:
    """清洗：破折号归一化、去隐形字符、去段落分隔符 <?>。

    与 problem1_rating.py 的 _clean 采用同一套归一化，保证两个问题的
    文本口径一致（便于问题2的情绪与问题1的评级做交叉校验）。
    """
    text = re.sub(r"[\xad‐‑‒–—−－]", "-", text)
    for ch in ["﻿", "​", "‌", "‍", "\xa0"]:
        text = text.replace(ch, "")
    text = text.replace("''", "”").replace("``", "“")
    text = re.sub(r"[‘’‚‛‟「」]", '"', text)
    text = text.replace("<?>", " ")
    return text.strip()


# ---------------------------------------------------------------------------
# 情绪计算
# ---------------------------------------------------------------------------
def tokenize(text: str):
    """返回 [(词, 起始位置, 结束位置)]。

    用 jieba.tokenize 而非 lcut，是为了拿到字符位置——据此把落在
    “公司简称出现位置”内的词判为误匹配（见 _count 的 blocked 参数）。
    """
    return [(w, s, e) for w, s, e in jieba.tokenize(text) if w.strip()]


def company_spans(text: str, company: str):
    """公司简称在正文中出现的所有字符区间。"""
    if not company:
        return []
    return [(m.start(), m.end()) for m in re.finditer(re.escape(company), text)]


def _has_negator(tokens, idx):
    """情绪词之前 NEG_WINDOW 个 token 内是否有否定词。

    返回否定词个数，奇数个表示极性翻转（处理“并非不理想”这类双重否定）。
    """
    cnt = 0
    for j in range(idx - 1, max(-1, idx - 1 - NEG_WINDOW), -1):
        t = tokens[j][0]
        # 遇到标点视为跨句，不再向前看
        if re.match(r"^[，。；：、！？,.;:!?）)】”\"']+$", t):
            break
        if t in NEGATORS:
            cnt += 1
    return cnt


def _count(tokens, pos_set, neg_set, blocked=None):
    """统计一段 token 序列的情绪词。返回 (raw_pos, raw_neg, adj_pos, adj_neg, n_flip, terms)。

    blocked 为需跳过的字符区间（公司简称出现的位置）。词典里有不少词会
    命中公司名——如“沧州**明珠**”里的“明珠”（正面）、“中国**平安**”里的“平安”，
    这些是公司名而非情绪表达。实测影响 513 行、1199 次命中，
    受影响行的净值平均偏差 2.13（约 8%），故按字符位置剔除。
    """
    raw_pos = raw_neg = adj_pos = adj_neg = n_flip = 0
    terms = {"pos": Counter(), "neg": Counter()}
    for i, (w, s, e) in enumerate(tokens):
        if blocked and any(s < se and e > ss for ss, se in blocked):
            continue
        is_pos, is_neg = w in pos_set, w in neg_set
        if not (is_pos or is_neg):
            continue
        flipped = _has_negator(tokens, i) % 2 == 1
        if flipped:
            n_flip += 1
        if is_pos:
            raw_pos += 1
            terms["pos"][w] += 1
            if flipped:
                adj_neg += 1
            else:
                adj_pos += 1
        else:
            raw_neg += 1
            terms["neg"][w] += 1
            if flipped:
                adj_pos += 1
            else:
                adj_neg += 1
    return raw_pos, raw_neg, adj_pos, adj_neg, n_flip, terms


def _ratio(p, n):
    tot = p + n
    return round(p / tot, 4) if tot else None


# 研报的例行风险披露段。语料中“风险”二字的出现有 93.3% 落在这一段内
# （15,823/16,964 次），是强制性披露而非分析师观点，会系统性拉低负面净值。
RISK_SECTION = re.compile(r"(风险提示|风险因素|投资风险|风险分析|主要风险|风险及对策|风险揭示)")
SECTION_SEP = "<?>"          # 源数据里的段落分隔符


def remove_risk_section(raw_text: str) -> str:
    """按段落删除“风险提示”所在的那一段，返回剔除后的原文。

    注意不能简单地“从头截断到风险提示处”：风险段并不总在末尾。
    实测 10,113 篇可定位风险段的研报中，有 581 篇（5.7%）风险段位于 50%~80% 处、
    76 篇（0.8%）位于 50% 之前——它们后面还有真正的正文
    （如双鹭药业的风险段在 65% 处，其后紧跟“下半年增速有望加快，维持‘推荐’评级…”）。
    故按 <?> 分段，只删掉含风险标题的那一段，保留其余各段。
    """
    if RISK_SECTION.search(raw_text) is None:
        return raw_text
    kept, removed = [], False
    for seg in raw_text.split(SECTION_SEP):
        if not removed:
            m = RISK_SECTION.search(seg)
            if m:
                # 标题之前若还有正文，保留；标题及其后的风险条目丢弃
                if seg[:m.start()].strip():
                    kept.append(seg[:m.start()])
                removed = True
                continue
        kept.append(seg)
    return SECTION_SEP.join(kept)


def score_text(raw_text: str, pos_set, neg_set, company: str = ""):
    """返回该文本的情绪指标 dict（三套口径，见 README）。参数为**未清洗**的原文。"""
    full = clean_text(raw_text)
    tokens = tokenize(full)
    n_tokens = len(tokens)
    rp, rn, ap, an, n_flip, terms = _count(
        tokens, pos_set, neg_set, company_spans(full, company))

    # 正文口径：剔除风险提示段后重算（正文单独清洗，故公司名区间需按正文重算）
    body = clean_text(remove_risk_section(raw_text))
    if body != full:
        body_tokens = tokenize(body)
        bp, bn, _, _, _, _ = _count(
            body_tokens, pos_set, neg_set, company_spans(body, company))
    else:
        bp, bn = rp, rn

    return {
        "分词数": n_tokens,
        "正面词数": rp,
        "负面词数": rn,
        "情绪词数": rp + rn,
        "情绪净值": rp - rn,
        "情绪比例": _ratio(rp, rn),
        "情绪密度": round((rp + rn) / n_tokens * 100, 3) if n_tokens else None,
        "正面词数_否定调整": ap,
        "负面词数_否定调整": an,
        "情绪净值_否定调整": ap - an,
        "情绪比例_否定调整": _ratio(ap, an),
        "否定翻转词数": n_flip,
        "正面词数_正文": bp,
        "负面词数_正文": bn,
        "情绪净值_正文": bp - bn,
        "情绪比例_正文": _ratio(bp, bn),
        "正面命中词": ";".join(w for w, _ in terms["pos"].most_common(MAX_HIT_TERMS)),
        "负面命中词": ";".join(w for w, _ in terms["neg"].most_common(MAX_HIT_TERMS)),
    }


# 极性分类阈值：以“正面词占比”划分。研报整体偏正面，故中性带按比例而非绝对值设。
POLARITY_POS, POLARITY_NEG = 0.75, 0.45


def polarity(ratio):
    if ratio is None:
        return "无情绪词"
    if ratio >= POLARITY_POS:
        return "正面"
    if ratio <= POLARITY_NEG:
        return "负面"
    return "中性"


# ---------------------------------------------------------------------------
# 与问题1结果对齐（用于交叉校验：买入评级的研报是否真的更正面）
# ---------------------------------------------------------------------------
PROBLEM1_XLSX = (PROJECT_ROOT / "problem1_评级提取" / "output"
                 / "问题1_评级提取结果.xlsx")


def load_problem1_ratings():
    """按**行号**读取问题1的标准化评级。

    注意：源数据的“序号”列不唯一（17,712 行只有 1,496 个不同序号），
    不能作主键，因此这里按行序对齐——问题1输出与本脚本都保持源数据行序。
    """
    if not PROBLEM1_XLSX.exists():
        print("  提示：未找到问题1结果，评级列留空（不影响情绪指标）")
        return []
    ws = openpyxl.load_workbook(PROBLEM1_XLSX, read_only=True).active
    return [r[7] or "" for r in ws.iter_rows(min_row=2, values_only=True)]


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def main():
    if not INPUT_XLSX.exists():
        raise FileNotFoundError(f"未找到输入文件 {INPUT_XLSX}")

    print("加载词典 ...")
    pos_set, neg_set, ambiguous = load_dictionary(DICT_XLSX)
    print(f"  正面词 {len(pos_set)} 个，负面词 {len(neg_set)} 个，"
          f"剔除正负重叠 {len(ambiguous)} 个")
    register_to_jieba(pos_set | neg_set)

    wb = openpyxl.load_workbook(INPUT_XLSX, read_only=True)
    ws = wb["Sheet1"]

    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "情绪分析结果"
    # metric_cols[:16] 均为 score_text 直接返回的指标；后 5 列另行计算/拼接
    metric_cols = [
        "分词数", "正面词数", "负面词数", "情绪词数", "情绪净值", "情绪比例",
        "情绪密度", "正面词数_否定调整", "负面词数_否定调整", "情绪净值_否定调整",
        "情绪比例_否定调整", "否定翻转词数",
        "正面词数_正文", "负面词数_正文", "情绪净值_正文", "情绪比例_正文",
        "情绪极性", "情绪极性_否定调整", "情绪极性_正文",
        "正面命中词", "负面命中词",
    ]
    out_ws.append(["fordate", "序号", "stkcd", "公司简称", "证券公司",
                   "评级_标准化", "研报字数"] + metric_cols)

    ratings = load_problem1_ratings()

    dist_raw, dist_adj, dist_body = Counter(), Counter(), Counter()
    tot_pos, tot_neg, tot_flip = 0, 0, 0
    # 评级 × 情绪极性 交叉表（用净值均值的差来检验区分度）
    crosstab = defaultdict(lambda: {"n": 0, "net": 0, "net_adj": 0, "net_body": 0})
    n = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        n += 1
        fordate, seq, stkcd, company, broker = row[0], row[1], row[2], row[3], row[4]
        text, wordcount = row[5] or "", row[7]
        m = score_text(text, pos_set, neg_set, company or "")
        pol = polarity(m["情绪比例"])
        pol_adj = polarity(m["情绪比例_否定调整"])
        pol_body = polarity(m["情绪比例_正文"])
        dist_raw[pol] += 1
        dist_adj[pol_adj] += 1
        dist_body[pol_body] += 1
        tot_pos += m["正面词数"]
        tot_neg += m["负面词数"]
        tot_flip += m["否定翻转词数"]
        rating = ratings[n - 1] if n - 1 < len(ratings) else ""
        if rating:
            cell = crosstab[rating]
            cell["n"] += 1
            cell["net"] += m["情绪净值"]
            cell["net_adj"] += m["情绪净值_否定调整"]
            cell["net_body"] += m["情绪净值_正文"]
        out_ws.append([fordate, seq, stkcd, company, broker, rating, wordcount]
                      + [m[c] for c in metric_cols[:16]] + [pol, pol_adj, pol_body]
                      + [m["正面命中词"], m["负面命中词"]])
        if n % 5000 == 0:
            print(f"  已处理 {n} 行 ...")

    wb.close()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out_wb.save(OUTPUT_XLSX)

    print(f"\n处理完成：共 {n} 行，结果写入 {OUTPUT_XLSX}\n")
    print(f"词典命中总量：正面词 {tot_pos} 次，负面词 {tot_neg} 次，"
          f"否定翻转 {tot_flip} 次（占命中 "
          f"{tot_flip / max(1, tot_pos + tot_neg) * 100:.1f}%）")
    print("\n情绪极性分布（原始计数）：")
    for k, v in dist_raw.most_common():
        print(f"  {k}: {v} ({v / n * 100:.2f}%)")
    print("\n情绪极性分布（否定调整后）：")
    for k, v in dist_adj.most_common():
        print(f"  {k}: {v} ({v / n * 100:.2f}%)")
    print("\n情绪极性分布（正文口径，剔除风险提示段）：")
    for k, v in dist_body.most_common():
        print(f"  {k}: {v} ({v / n * 100:.2f}%)")

    if crosstab:
        order = ["买入", "推荐", "增持", "中性", "持有", "减持", "卖出", "回避", "观望"]
        print("\n交叉校验：各评级下的平均情绪净值（越高表示文本越正面）")
        print(f"  {'评级':<6}{'样本数':>8}{'均值_原始':>11}{'均值_否定调整':>13}{'均值_正文':>11}")
        for r in order:
            if r not in crosstab:
                continue
            c = crosstab[r]
            print(f"  {r:<6}{c['n']:>8}{c['net'] / c['n']:>11.2f}"
                  f"{c['net_adj'] / c['n']:>13.2f}{c['net_body'] / c['n']:>11.2f}")


if __name__ == "__main__":
    main()
