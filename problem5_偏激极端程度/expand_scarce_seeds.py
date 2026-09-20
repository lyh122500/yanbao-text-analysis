#!/usr/bin/env python3
"""为“绝对化/夸张修辞”两个稀缺维度生成候选池：逐种子扩展 + 低频次阈值。

与 expand_candidates.py 的两点区别，针对其已暴露的缺陷：

1. **逐种子扩展**，不用“类别中心”。原方法把某类全部种子的向量取平均作为中心；
   绝对化只有 1 个种子通过筛选时，中心就等于该词自身，扩出来的全是它的搭配词
   （实测候选为“国内/具有/领域”）。逐个种子各自扩 Top-N 再合并，避免被单个
   高频种子主导。
2. **低频次阈值**（--min-count 默认 3）。原方法 min_count=20 会杀掉定义该维度
   的稀缺词：绝对化的 6 个种子只有“绝对”通过，夸张修辞 21 个只有 3 个通过。

输出仅为待审核候选池，供后续独立判断过滤；**不参与正式计分**。
"""
import argparse
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

import jieba
import numpy as np
import openpyxl
import scipy
import scipy.sparse as sp
from sklearn.preprocessing import normalize

from problem5_extremity import HERE, Analyzer, BOUNDARY, LABELS, dump, normalized, sha, spans
from expand_candidates import STOP

TARGET = ("absolute", "hyperbole")


def build_freq_and_examples(input_path, analyzer):
    """返回 (句子序列, 词频, 词->例句)。整篇去重、种子预注册。"""
    tokenizer = jieba.Tokenizer()
    tokenizer.initialize()
    for term in sorted(analyzer.lookup):
        tokenizer.add_word(term, freq=200000)
    wb = openpyxl.load_workbook(input_path, read_only=True, data_only=True)
    it = wb.active.values
    h = list(next(it))
    ti = h.index("研报文本")
    seen, sentences, examples = set(), [], defaultdict(list)
    for source_row, row in enumerate(it, 2):
        text = normalized(str(row[ti] or ""))[0]
        key = sha_text(text)
        if key in seen:
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
    freq = Counter(t for s in sentences for t in s)
    return sentences, freq, examples


def sha_text(text):
    import hashlib
    return hashlib.sha256(text.encode()).digest()


def per_seed_candidates(sentences, freq, seeds, min_count, max_vocab, max_context,
                        window, top_n):
    """每个种子各自扩 Top-N，返回 {seed: [(word, score, freq), ...]}。"""
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
    p.add_argument("--min-count", type=int, default=3)
    p.add_argument("--top-n", type=int, default=15)
    p.add_argument("--max-vocab", type=int, default=20000)
    p.add_argument("--max-context", type=int, default=3000)
    p.add_argument("--window", type=int, default=3)
    args = p.parse_args()

    analyzer = Analyzer()
    seeds_by_cat = {c: analyzer.categories[c] for c in TARGET}
    all_seeds = {w for ws in seeds_by_cat.values() for w in ws}
    print("构建语料共现 ...")
    sentences, freq, examples = build_freq_and_examples(args.input, analyzer)
    print(f"  分句 {len(sentences):,} 段，词表 {len(freq):,} 词")
    per_seed, detail = per_seed_candidates(
        sentences, freq, all_seeds, args.min_count, args.max_vocab,
        args.max_context, args.window, args.top_n)

    # 合并去重：记录每个候选被哪些种子扩出来、最高分
    merged = {}
    for cat, seeds in seeds_by_cat.items():
        for seed in seeds:
            for word, score, f in per_seed.get(seed, []):
                rec = merged.setdefault(word, {"candidate": word, "freq": f,
                                               "max_score": 0.0, "from_seeds": [],
                                               "categories": set()})
                rec["max_score"] = max(rec["max_score"], score)
                rec["from_seeds"].append(f"{seed}({score:.3f})")
                rec["categories"].add(cat)
    rows = []
    for w, r in merged.items():
        if w in analyzer.lookup:      # 已在词表中
            continue
        if len(w) < 2:                # 单字噪声大
            continue
        if not re.fullmatch(r"[一-鿿]+", w):
            continue
        rows.append({"候选词": w, "语料频数": r["freq"],
                     "最高相似度": round(r["max_score"], 6),
                     "来源维度": "+".join(sorted(LABELS[c] for c in r["categories"])),
                     "来源种子": "、".join(sorted(r["from_seeds"])),
                     "语料例句": "\n".join(examples.get(w, [])[:2]),
                     "独立判断A": "", "独立判断B": "", "独立判断C": "",
                     "三判一致": "", "人工是否收入": "", "审核后维度": "", "审核人": "", "备注": ""})
    rows.sort(key=lambda x: (-x["语料频数"], x["候选词"]))
    dump(args.output, {
        "method": "逐种子PPMI上下文向量扩展（非类别中心）；低阈值以保留稀缺种子",
        "target_categories": [LABELS[c] for c in TARGET],
        "min_count": args.min_count, "top_n": args.top_n, "window": args.window,
        "max_vocabulary": args.max_vocab, "max_contexts": args.max_context,
        "input_sha256": sha(args.input), "script_sha256": sha(Path(__file__)),
        "lexicon_sha256": sha(analyzer.lexicon_path),
        # 依赖版本与分词词典哈希：PPMI 结果依赖 jieba 切分，缺了这项无法复现
        "jieba": jieba.__version__, "numpy": np.__version__, "scipy": scipy.__version__,
        "jieba_dict_sha256": sha(Path(jieba.__file__).parent / "dict.txt"), "HMM": False,
        "stopwords": sorted(STOP),
        "seeds_used": {k: v for k, v in per_seed.items()},
        "candidates": len(rows), "status": "待独立判断与人工审核；不参与正式计分",
        **detail})
    dump(args.output.with_suffix(".rows.json"), rows)
    print(f"候选池: {len(rows)} 条 → {args.output}")


if __name__ == "__main__":
    main()
