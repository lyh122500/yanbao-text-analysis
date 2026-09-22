#!/usr/bin/env python3
"""为稀缺维度生成 PPMI 候选池：逐种子扩展 + 显式频次过滤。

对应问题5 流程的「阶段一：构建词表」。目标维度由语料覆盖（而非种子数）选出，
见 README「词表扩展（稀缺维度）」：夸张修辞度 3.7%、情绪极端度 14.4%。

两点设计（沿用 problem5/expand_scarce_seeds.py，理由见该文件）：

1. **逐种子扩展，不用"类别中心"**。类别中心法在种子稀少时中心退化成该词自身的
   向量，扩出来全是它的搭配词。逐个种子各自扩 Top-N 再合并，避免被单个高频种子主导。
2. **低频次阈值**（--min-count 默认 3）。min_count=20 会杀掉定义该维度的稀缺词。

本任务新增一处显式化：原问题5 的「只保留来自频次≥10 的种子的候选」是手工做的
（README 只有结论没有代码）。这里做成参数 `--seed-min-count`，让这一步可复现。

输出仅为待审核候选池，供独立判断与人工审核过滤；**不参与正式计分**，
也不修改 resources/lexicon.json。
"""
import argparse
import re
from collections import Counter, defaultdict
from pathlib import Path

import jieba
import numpy as np
import openpyxl
import scipy
import scipy.sparse as sp
from sklearn.preprocessing import normalize

from score_extremity import (BOUNDARY, DIMS, HERE, RES, Scorer, dump, normalized,
                             sha, spans)

# 目标维度：按语料覆盖最低的两个维度选（夸张修辞 3.7%、情绪极端 14.4%）
TARGET = ("hyperbole", "emotion")

# 停用词：与 problem5/expand_candidates.py 同表（同一语料，同一停用依据）
STOP = set("公司 行业 市场 目前 未来 我们 以及 由于 其中 同时 主要 此外 因此 方面 情况 "
           "通过 进行 实现 相关 业务 产品 年 月 日 万元 亿元 有限公司 股份有限公司 "
           "的 了 是 在 将 会 也 和 与 为 有 对 但 等 及 于 中 而 以 其 之一 作为 一家 "
           "今年 明年".split())

LABELS = {k: label for k, label in DIMS}
# 情绪维度含两个子类，候选要带子类，否则审核后的词不知道进哪个
SUBCLASSES = {"hyperbole": ["夸张修辞"],
              "emotion": ["极端积极", "极端消极"]}


def build_freq_and_examples(input_path, seeds):
    """返回 (句子序列, 词频, 词->例句)。整篇去重、种子预注册到分词词典。"""
    tokenizer = jieba.Tokenizer()
    tokenizer.initialize()
    for term in sorted(seeds):
        tokenizer.add_word(term, freq=200000)
    wb = openpyxl.load_workbook(input_path, read_only=True, data_only=True)
    it = wb.active.values
    h = list(next(it))
    ti = h.index("研报文本")
    seen, sentences, examples = set(), [], defaultdict(list)
    for source_row, row in enumerate(it, 2):
        text = normalized(str(row[ti] or ""))[0]
        key = sha_text(text)
        if key in seen:                       # 完全重复的研报只算一次
            continue
        seen.add(key)
        for a, b in spans(text, BOUNDARY):
            for segment in re.findall(r"[一-鿿]+", text[a:b]):
                tokens = list(tokenizer.cut(segment, HMM=False))
                if not tokens:
                    continue
                sentences.append(tokens)
                for t in set(tokens):
                    if len(examples[t]) < 3:
                        examples[t].append(f"源行{source_row}：{text[a:b][:150]}")
    wb.close()
    return sentences, Counter(t for s in sentences for t in s), examples


def sha_text(text):
    import hashlib
    return hashlib.sha256(text.encode()).digest()


def per_seed_candidates(sentences, freq, seeds, min_count, max_vocab, max_context,
                        window, top_n):
    """每个种子各自扩 Top-N，返回 ({seed: [(word, score, freq), ...]}, 统计)。"""
    eligible = [w for w, n in freq.items() if n >= min_count and w not in STOP]
    vocab = sorted(eligible, key=lambda w: (-freq[w], w))[:max_vocab]
    vocab = sorted(set(vocab) | {w for w in seeds if w in freq})
    contexts = sorted(eligible, key=lambda w: (-freq[w], w))[:max_context]
    wi = {w: i for i, w in enumerate(vocab)}
    ci = {w: i for i, w in enumerate(contexts)}
    counts = Counter()
    for sentence in sentences:
        for i, word in enumerate(sentence):
            if word not in wi:
                continue
            for j in range(max(0, i - window), min(len(sentence), i + window + 1)):
                ctx = sentence[j]
                if j != i and ctx in ci:
                    counts[(wi[word], ci[ctx])] += 1
    if not counts:
        raise ValueError("语料不足以构建共现矩阵")
    keys = sorted(counts)
    m = sp.coo_matrix(([counts[k] for k in keys], ([k[0] for k in keys], [k[1] for k in keys])),
                      shape=(len(vocab), len(contexts)), dtype=np.float64)
    row_sum = np.asarray(m.sum(axis=1)).ravel()
    col_sum = np.asarray(m.sum(axis=0)).ravel()
    # PPMI：正点互信息，逐元素 max(0, log(pmi))
    m.data = np.maximum(0, np.log(m.data * m.data.sum() / (row_sum[m.row] * col_sum[m.col])))
    vec = normalize(m.tocsr())

    out = {}
    for seed in seeds:
        if seed not in wi or not vec[wi[seed]].nnz:
            out[seed] = []
            continue
        scores = np.asarray((vec @ vec[wi[seed]].T).todense()).ravel()
        ranked = sorted((w for w in vocab if w not in seeds and w not in STOP
                         and freq[w] >= min_count and vec[wi[w]].nnz),
                        key=lambda w: (-scores[wi[w]], w))[:top_n]
        out[seed] = [(w, float(scores[wi[w]]), freq[w]) for w in ranked]
    return out, {"vocabulary": len(vocab), "contexts": len(contexts),
                 "cooccurrence_nonzero": len(counts)}


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--input", type=Path, default=HERE.parent / "综合分析结果.xlsx")
    p.add_argument("--output", type=Path, default=HERE / "output" / "稀缺维度候选池.json")
    p.add_argument("--min-count", type=int, default=3,
                   help="候选词与种子的最低语料频数（低阈值以保留稀缺种子）")
    p.add_argument("--seed-min-count", type=int, default=10,
                   help="种子的最低语料频数；低频种子扩出的近邻没有统计意义，整条剔除")
    p.add_argument("--top-n", type=int, default=15)
    p.add_argument("--max-vocab", type=int, default=20000)
    p.add_argument("--max-context", type=int, default=3000)
    p.add_argument("--window", type=int, default=3)
    args = p.parse_args()

    scorer = Scorer()
    # 只用偏激词作种子：参照词是分母与对照，语义上不属于该维度
    seeds_by_cat = {c: {w for w, m in scorer.dims[c].items() if m["role"] == "偏激词"}
                    for c in TARGET}
    all_seeds = {w for ws in seeds_by_cat.values() for w in ws}
    print("构建语料共现 ...")
    sentences, freq, examples = build_freq_and_examples(args.input, all_seeds)
    print(f"  分句 {len(sentences):,} 段，词表 {len(freq):,} 词")

    # 频次过滤：低频种子的共现窗口是噪声（见 README 的「绝无」例），整条剔除
    dropped_seeds = sorted(w for w in all_seeds if freq[w] < args.seed_min_count)
    for c in TARGET:
        seeds_by_cat[c] = {w for w in seeds_by_cat[c] if freq[w] >= args.seed_min_count}
    used_seeds = {w for ws in seeds_by_cat.values() for w in ws}
    print(f"  种子 {len(all_seeds)} 个，频次≥{args.seed_min_count} 的 {len(used_seeds)} 个"
          f"，剔除 {len(dropped_seeds)} 个：{' '.join(dropped_seeds) or '无'}")

    per_seed, detail = per_seed_candidates(
        sentences, freq, used_seeds, args.min_count, args.max_vocab,
        args.max_context, args.window, args.top_n)

    # 合并去重：记录每个候选被哪些种子扩出来、最高分
    merged = {}
    for cat, seeds in seeds_by_cat.items():
        for seed in seeds:
            for word, score, f in per_seed.get(seed, []):
                rec = merged.setdefault(word, {"freq": f, "max_score": 0.0,
                                               "from_seeds": [], "categories": set()})
                rec["max_score"] = max(rec["max_score"], score)
                rec["from_seeds"].append(f"{seed}({score:.3f})")
                rec["categories"].add(cat)
    rows = []
    for w, r in merged.items():
        if w in scorer.lookup:                 # 已在词表中
            continue
        if len(w) < 2:                         # 单字噪声大
            continue
        if not re.fullmatch(r"[一-鿿]+", w):
            continue
        rows.append({"候选词": w, "语料频数": r["freq"],
                     "最高相似度": round(r["max_score"], 6),
                     "来源维度": "+".join(sorted(LABELS[c] for c in r["categories"])),
                     "来源种子": "、".join(sorted(r["from_seeds"])),
                     "语料例句": "\n".join(examples.get(w, [])[:2]),
                     "独立判断A": "", "独立判断B": "", "独立判断C": "",
                     "三判一致": "", "人工是否收入": "", "审核后子类": "",
                     "审核人": "", "备注": ""})
    rows.sort(key=lambda x: (-x["语料频数"], x["候选词"]))
    dump(args.output, {
        "method": "逐种子PPMI上下文向量扩展（非类别中心）",
        "target_categories": [LABELS[c] for c in TARGET],
        "target_rule": "按语料覆盖最低的两个维度选（夸张修辞 3.7%、情绪极端 14.4%），"
                       "非按种子数——覆盖度才对应扩展的目的",
        "subclasses": {LABELS[c]: SUBCLASSES[c] for c in TARGET},
        "min_count": args.min_count, "seed_min_count": args.seed_min_count,
        "top_n": args.top_n, "window": args.window,
        "max_vocabulary": args.max_vocab, "max_contexts": args.max_context,
        "input_sha256": sha(args.input), "script_sha256": sha(Path(__file__)),
        "lexicon_sha256": sha(RES / "lexicon.json"),
        # PPMI 结果依赖 jieba 切分，缺了依赖版本与分词词典哈希就无法复现
        "jieba": jieba.__version__, "numpy": np.__version__, "scipy": scipy.__version__,
        "jieba_dict_sha256": sha(Path(jieba.__file__).parent / "dict.txt"), "HMM": False,
        "stopwords": sorted(STOP),
        "seeds_used": {k: v for k, v in per_seed.items()},
        "seeds_dropped_low_freq": {w: freq[w] for w in dropped_seeds},
        "candidates": len(rows), "status": "待独立判断与人工审核；不参与正式计分",
        **detail})
    dump(args.output.with_suffix(".rows.json"), rows)
    print(f"候选池: {len(rows)} 条 → {args.output}")


if __name__ == "__main__":
    main()
