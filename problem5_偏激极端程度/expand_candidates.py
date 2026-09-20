#!/usr/bin/env python3
"""Deterministic PPMI context-vector expansion; writes review candidates ONLY."""
import argparse
import csv
import hashlib
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

STOP = set("公司 行业 市场 目前 未来 我们 以及 由于 其中 同时 主要 此外 因此 方面 情况 通过 进行 实现 相关 业务 产品 年 月 日 万元 亿元 有限公司 股份有限公司 的 了 是 在 将 会 也 和 与 为 有 对 但 等 及 于 中 而 以 其 之一 作为 一家 今年 明年".split())


def candidates(sentences, categories, min_count=20, max_vocab=10000, max_context=3000, window=3, top_n=30):
    freq = Counter(t for sentence in sentences for t in sentence)
    seeds = {w for words in categories.values() for w in words}
    eligible = [w for w, n in freq.items() if n >= min_count and w not in STOP]
    vocab = sorted(eligible, key=lambda w: (-freq[w], w))[:max_vocab]
    vocab = sorted(set(vocab) | (seeds & set(freq)))
    contexts = sorted(eligible, key=lambda w: (-freq[w], w))[:max_context]
    wi, ci = {w: i for i, w in enumerate(vocab)}, {w: i for i, w in enumerate(contexts)}
    counts = Counter()
    for sentence in sentences:
        for i, word in enumerate(sentence):
            if word not in wi: continue
            for j in range(max(0, i-window), min(len(sentence), i+window+1)):
                context = sentence[j]
                if j != i and context in ci:
                    counts[(wi[word], ci[context])] += 1
    if not counts: raise ValueError("语料不足以构建共现矩阵")
    keys = sorted(counts)
    matrix = sp.coo_matrix(([counts[k] for k in keys], ([k[0] for k in keys], [k[1] for k in keys])),
                           shape=(len(vocab), len(contexts)), dtype=np.float64)
    row_sum = np.asarray(matrix.sum(axis=1)).ravel()
    col_sum = np.asarray(matrix.sum(axis=0)).ravel()
    matrix.data = np.maximum(0, np.log(matrix.data * matrix.data.sum() / (row_sum[matrix.row]*col_sum[matrix.col])))
    vectors = normalize(matrix.tocsr())
    result, used_seeds = [], {}
    for category, words in sorted(categories.items()):
        used = sorted(w for w in words if w in wi and freq[w] >= min_count and vectors[wi[w]].nnz)
        used_seeds[category] = used
        if not used: continue
        center = normalize(np.asarray(vectors[[wi[w] for w in used]].mean(axis=0))).ravel()
        scores = np.asarray(vectors @ center).ravel()
        ranked = sorted((w for w in vocab if w not in seeds and w not in STOP and freq[w] >= min_count),
                        key=lambda w: (-scores[wi[w]], w))[:top_n]
        for rank, word in enumerate(ranked, 1):
            result.append({"建议维度": LABELS[category], "排名": rank, "候选词": word,
                           "语料频数": freq[word], "上下文余弦相似度": round(float(scores[wi[word]]), 10),
                           "是否收入词典": "", "审核后维度": "", "审核人": "", "备注": ""})
    return result, {"vocabulary": len(vocab), "contexts": len(contexts), "used_seeds": used_seeds,
                    "cooccurrence_nonzero": len(counts)}


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--input", type=Path, default=HERE.parent / "综合分析结果.xlsx")
    p.add_argument("--output-dir", type=Path, default=HERE / "output")
    p.add_argument("--min-count", type=int, default=20)
    p.add_argument("--top-n", type=int, default=30)
    args = p.parse_args()
    analyzer = Analyzer()
    tokenizer = jieba.Tokenizer()
    tokenizer.initialize()
    for term in sorted(analyzer.lookup): tokenizer.add_word(term, freq=200000)
    wb = openpyxl.load_workbook(args.input, read_only=True, data_only=True)
    it = wb.active.values
    h = list(next(it)); ti = h.index("研报文本")
    seen, sentences, examples = set(), [], defaultdict(list)
    duplicates = 0
    for source_row, row in enumerate(it, 2):
        text = normalized(str(row[ti] or ""))[0]
        key = hashlib.sha256(text.encode()).digest()
        if key in seen:
            duplicates += 1
            continue
        seen.add(key)
        for a, b in spans(text, BOUNDARY):
            # Keep punctuation/numeric positions as context breaks, not artificial neighbors.
            for segment in re.findall(r"[\u4e00-\u9fff]+", text[a:b]):
                tokens = list(tokenizer.cut(segment, HMM=False))
                sentences.append(tokens)
                for t in set(tokens):
                    if len(examples[t]) < 2:
                        examples[t].append("源行{}：{}".format(source_row, text[a:b][:160]))
    wb.close()
    result, details = candidates(sentences, analyzer.categories, min_count=args.min_count, top_n=args.top_n)
    for r in result: r["语料例句"] = "\n".join(examples[r["候选词"]])
    args.output_dir.mkdir(parents=True, exist_ok=True)
    path = args.output_dir / "词典扩展候选_待审核.csv"
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(result[0]) if result else ["候选词"])
        w.writeheader(); w.writerows(result)
    dump(args.output_dir / "expansion_manifest.json", {"method": "PPMI上下文向量余弦相似度；不训练随机神经词向量",
        "input_sha256": sha(args.input), "script_sha256": sha(__file__),
        "analyzer_sha256": sha(HERE / "problem5_extremity.py"), "lexicon_sha256": sha(analyzer.lexicon_path),
        "jieba": jieba.__version__, "numpy": np.__version__, "scipy": scipy.__version__,
        "jieba_dict_sha256": sha(Path(jieba.__file__).parent / "dict.txt"), "HMM": False,
        "window": 3, "min_count": args.min_count, "top_n": args.top_n,
        "max_vocabulary": 10000, "max_contexts": 3000, "stopwords": sorted(STOP),
        "duplicate_documents_skipped": duplicates, "candidates": len(result),
        "candidate_sha256": sha(path), "status": "未经人工审核，不参与正式计分", **details})
    print(json.dumps({"candidates": len(result), "duplicates_skipped": duplicates}, ensure_ascii=False))


if __name__ == "__main__": main()
