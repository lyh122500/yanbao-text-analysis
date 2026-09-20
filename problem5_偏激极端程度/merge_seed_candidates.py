#!/usr/bin/env python3
"""合并"语料候选 + 多方独立判断"，生成稀缺维度的待审核种子表。

输入（三者都是冻结文件，改了要重新核验）：
  - output/稀缺维度候选池.rows.json   语料候选池（expand_scarce_seeds.py 产出）
  - resources/独立判断结果.json        三次独立判断的原始分类与新增候选
  - resources/种子审核结论.json        人工审核裁定（是否收入词典／维度／理由）

规则：
  - **只保留来自语料候选池的词**。模型“凭知识提出”的词没有语料溯源
    （说不出它由哪个种子、多少相似度扩出来），不可复现也不可审计，
    默认剔除；确需保留时用 --include-model-proposed。
  - 只有被 ≥2 个独立判断判为同一维度的词才进入待审核表；
  - 未裁定的词若已在 lexicon.json 中则剔除（Analyzer 不允许跨类别重复）；
    已裁定的候选始终保留，以便保存完整审核记录——该词可能已按裁定进入词表；
  - 每条保留“来源种子／最高相似度／语料频数”，可回溯到 PPMI 计算过程。

输出的审核栏由 种子审核结论.json 填入，不手改本 CSV——手改会在重跑时丢失。

输出：output/稀缺维度种子候选_待审核.csv
"""
import argparse
import csv
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

from problem5_extremity import HERE, LABELS, dump, normalized, sha

DIM = {"绝对化": "absolute", "夸张修辞": "hyperbole"}


def corpus_freq_and_examples(input_path, words, limit=2):
    """统计候选词在语料中的出现次数，并各取 limit 条真实语境。"""
    src = openpyxl.load_workbook(input_path, read_only=True, data_only=True)
    it = src.active.values
    h = list(next(it))
    ti = h.index("研报文本")
    freq, examples = Counter(), defaultdict(list)
    for source_row, row in enumerate(it, 2):
        text = normalized(str(row[ti] or ""))[0]
        for w in words:
            for m in re.finditer(re.escape(w), text):
                freq[w] += 1
                if len(examples[w]) < limit:
                    examples[w].append(f"源行{source_row}：{text[max(0,m.start()-45):m.end()+55]}")
    src.close()
    return freq, examples


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--input", type=Path, default=HERE.parent / "综合分析结果.xlsx")
    p.add_argument("--pool", type=Path, default=HERE / "output" / "稀缺维度候选池.rows.json")
    p.add_argument("--judgments", type=Path, default=HERE / "resources" / "独立判断结果.json")
    p.add_argument("--output", type=Path, default=HERE / "output" / "稀缺维度种子候选_待审核.csv")
    p.add_argument("--include-model-proposed", action="store_true",
                   help="同时保留模型凭知识提出的词（默认剔除：无语料溯源，不可复现）")
    p.add_argument("--decisions", type=Path, default=HERE / "resources" / "种子审核结论.json",
                   help="人工审核结论（冻结输入）。已裁定的候选即使已进入词表也保留在表中，"
                        "以保存完整的审核记录；未裁定的词若已在词表中则剔除。")
    args = p.parse_args()

    spec = json.loads((HERE / "resources" / "lexicon.json").read_text(encoding="utf-8"))
    existing = {t: c for c, v in spec["categories"].items() for t in v}
    pool = {r["候选词"]: r for r in json.loads(args.pool.read_text(encoding="utf-8"))}
    judges = json.loads(args.judgments.read_text(encoding="utf-8"))

    # 1) 汇总三方对候选池的分类：{词: {标签: [判断名, ...]}}
    votes = defaultdict(lambda: defaultdict(list))
    for jname, jdata in judges.items():
        for word, label in jdata.get("classify", {}).items():
            base = re.sub(r"[（(].*", "", label).strip()
            votes[word][base].append(jname)

    # 2) 汇总三方新增候选：{词: {维度: [判断名, ...]}}
    props = defaultdict(lambda: defaultdict(list))
    for jname, jdata in judges.items():
        for dim_label, words in jdata.get("propose", {}).items():
            for w in words:
                props[w][dim_label].append(jname)

    # 人工审核结论是冻结输入：已裁定的候选保留在表中（保存完整审核记录，
    # 即使该词已进入词表）；未裁定的词若已在词表则不再列为候选。
    decisions = json.loads(args.decisions.read_text(encoding="utf-8")) if args.decisions.exists() else {}
    rows = []

    def add(word, dim_label, agreement, judges_list, source):
        if word in existing and word not in decisions:
            return
        p = pool.get(word, {})
        rows.append({
            "建议维度": dim_label, "候选词": word, "独立判断一致度": agreement,
            "各判断标签": "；".join(judges_list),
            # 溯源：这条候选是怎么从语料扩出来的，可回查 PPMI 计算过程
            "来源种子": p.get("来源种子", ""), "最高相似度": p.get("最高相似度", ""),
            "语料频数": "", "语料例句": "",
            "是否收入词典": "", "审核后维度": "", "审核人": "", "备注": ""})

    for word, labs in votes.items():
        for lab, names in labs.items():
            # “条件性”单独保留：这类词单独不是断言，只有组成特定搭配才是
            # （如“优势”需“绝对优势”），正对应 lexicon.json 的 exclude_patterns 机制。
            if lab not in DIM and lab != "条件性":
                continue
            if len(names) < 2:
                continue
            detail = [f'{n}={judges[n]["classify"].get(word, "")}' for n in sorted(names)]
            add(word, lab, f"{len(names)}/3", detail, "语料候选池")
    if args.include_model_proposed:
        for word, dims in props.items():
            for dim_label, names in dims.items():
                if len(names) < 2:
                    continue
                add(word, dim_label, f"{len(names)}/3",
                    [f"{n}=模型知识提出" for n in sorted(names)], "模型知识提出")
    else:
        skipped = sum(1 for dims in props.values()
                      for names in dims.values() if len(names) >= 2)
        print(f"  已剔除模型知识提出的候选 {skipped} 条（--include-model-proposed 可保留）")

    # 3) 应用审核结论 + 补语料频数与例句
    for r in rows:
        d = decisions.get(r["候选词"], {})
        r["是否收入词典"] = d.get("是否收入词典", "")
        r["审核后维度"] = d.get("审核后维度", "")
        r["备注"] = d.get("备注", "")
    freq, examples = corpus_freq_and_examples(args.input, {r["候选词"] for r in rows})
    for r in rows:
        r["语料频数"] = freq[r["候选词"]]
        r["语料例句"] = "\n".join(examples.get(r["候选词"], [])[:2])
    rows.sort(key=lambda r: (r["建议维度"], -int(r["独立判断一致度"].split("/")[0]), -r["语料频数"]))

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)
    dump(args.output.with_suffix(".manifest.json"), {
        "method": "语料候选（逐种子PPMI）+ 三方独立判断，≥2 方一致才入选",
        "judges": sorted(judges), "existing_lexicon_excluded": True,
        "pool_sha256": sha(args.pool), "judgments_sha256": sha(args.judgments),
        "decisions_sha256": sha(args.decisions) if args.decisions.exists() else None,
        "input_sha256": sha(args.input), "script_sha256": sha(Path(__file__)),
        "candidates": len(rows), "by_dimension": dict(Counter(r["建议维度"] for r in rows)),
        "agreement": dict(Counter(r["独立判断一致度"] for r in rows)),
        "status": "未经人工审核，不参与正式计分"})
    print(f"待审核候选 {len(rows)} 条 → {args.output}")
    for r in rows:
        print(f'   [{r["建议维度"]}] {r["候选词"]:<8} {r["独立判断一致度"]}'
              f'  频数{r["语料频数"]:>5}  相似度{r["最高相似度"]}  ← {r["来源种子"]}')


if __name__ == "__main__":
    main()
