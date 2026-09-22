#!/usr/bin/env python3
"""合并两份词表，生成 5 维度的种子词表（lexicon.json）。

来源：
  1. resources/任务5词表.xlsx —— 主词表，5 维度 185 词条（含参照词）
     列：序号/维度/子类/词语/测度角色/来源/英文对应/备注
  2. resources/LM词表中文校准对照表.xlsx —— 补充「确定性强度」
     只取「汇总」表第 15–16 行：确定性词(强断言) 与 模糊词(不确定)

处置（依据作者裁定，见 README「词表合并」）：
  - 「绝对」同时出现在两个维度，**归入绝对化**，从确定性词中移除；
  - 全库零命中的词保留在词表中并标注，不影响计分但会随产出记录；
  - 维度间不重复计分：同一词只归属一个维度；
  - LM 补充的确定性/模糊词按 resources/确定性词表校准.json 逐词校准
    （109 条 drop、5 条 reclassify，每条均有语料实测依据，见该文件与 README）；
  - 阶段一 PPMI 扩展中经人工审核「收入」的词，按 resources/种子审核结论.json
    指定的子类并入（见 README 第六节）。

输出：resources/lexicon.json
"""
import argparse
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
RES = HERE / "resources"

DIMS = [
    ("certainty", "确定性强度"),
    ("absolute", "绝对化/最高级"),
    ("emotion", "情绪极端度"),
    ("hyperbole", "夸张修辞度"),
    ("intensity", "程度加强度"),
]
LABEL2KEY = {label: key for key, label in DIMS}


def load_main(path):
    """读主词表，返回 [(维度标签, 子类, 词语, 测度角色, 来源)]。"""
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb["词表"]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or not r[3]:
            continue
        rows.append((str(r[1]).strip(), str(r[2]).strip(), str(r[3]).strip(),
                     str(r[4]).strip(), str(r[5]).strip()))
    wb.close()
    return rows


def load_lm(path):
    """读 LM 对照表的「汇总」表 15–16 行，返回 {子类: [词]}。

    第 15 行 = 确定性词(强断言)，第 16 行 = 模糊词(不确定)。
    单元格内以顿号分隔；条目里可能含 "/"（如「一定/确定」），拆成并列词。
    """
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb["汇总"]
    out = {}
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i not in (15, 16):
            continue
        label = str(row[0]).strip()
        cell = str(row[1] or "")
        words = []
        for part in cell.split("、"):
            part = part.strip()
            if not part:
                continue
            # 「一定/确定」这类斜杠并列拆开；但保留「不确定的/无限期的」原样含义时
            # 也按并列处理——LM 原文即把同义译法并列
            for w in re.split(r"[/／]", part):
                w = w.strip()
                if w:
                    words.append(w)
        out[label] = words
    wb.close()
    return out


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--main", type=Path, default=RES / "任务5词表.xlsx")
    p.add_argument("--lm", type=Path, default=RES / "LM词表中文校准对照表.xlsx")
    p.add_argument("--calibration", type=Path, default=RES / "确定性词表校准.json")
    p.add_argument("--decisions", type=Path, default=RES / "种子审核结论.json",
                   help="PPMI 扩展的人工审核结论；「收入」的词按指定子类并入")
    p.add_argument("--output", type=Path, default=RES / "lexicon.json")
    args = p.parse_args()

    main_rows = load_main(args.main)
    lm = load_lm(args.lm)
    cal = json.loads(args.calibration.read_text(encoding="utf-8"))
    drop, reclass = cal["drop"], cal["reclassify"]
    role_override = cal.get("role_override", {})
    decisions = json.loads(args.decisions.read_text(encoding="utf-8")) \
        if args.decisions.exists() else {}

    # 先收集全部词条（同一词可能在多处出现），再按裁定解决归属，
    # 这样跨维度重复不会在放置阶段被静默丢弃。
    entries = []          # (维度key, 子类, 词, 角色, 来源)
    for dim_label, subclass, word, role, source in main_rows:
        key = LABEL2KEY.get(dim_label)
        if key is None:
            raise ValueError(f"未知维度：{dim_label}")
        entries.append((key, subclass, word, role, source))
    n_main = len(entries)

    lm_added = Counter()
    lm_entries = []
    for label, words in lm.items():
        role = "偏激词" if "确定性" in label else "参照词"
        for w in words:
            lm_entries.append(("certainty", label, w, role, "LM汇总"))
    entries += lm_entries

    # 阶段一 PPMI 扩展：人工审核「收入」的词，按裁定子类并入。
    # 子类 → 维度 由主词表反查，不在代码里硬编码第二份映射。
    sub2cat = {subclass: LABEL2KEY[label]
               for label, subclass, *_ in main_rows if label in LABEL2KEY}
    accepted = []
    for word, d in decisions.items():
        if d.get("是否收入词典") != "是":
            continue
        sub = d.get("审核后子类")
        if sub not in sub2cat:
            raise ValueError(f"审核结论里 {word} 的子类「{sub}」不在词表维度中")
        accepted.append((sub2cat[sub], sub, word, "偏激词", "PPMI扩展+人工审核"))
    entries += accepted

    # 作者裁定：「绝对」归入绝对化（覆盖其在确定性词中的归属）
    RULING = {"绝对": "absolute"}
    RULING.update(reclass)
    # 校准只在 certainty 维度生效，避免误伤主词表里同形的其它维度词条。
    # 按「不同词」计数，同一词在多处出现只算一次。
    n_dropped = len({w for k, _, w, _, _ in entries if k == "certainty" and w in drop})
    entries = [e for e in entries if not (e[0] == "certainty" and e[2] in drop)]
    for i, (key, subclass, word, role, source) in enumerate(entries):
        if word in RULING:
            new_key = RULING[word]
            # 改判到绝对化的是最高级表达，角色随之改为偏激词，避免混入参照词分母
            new_sub = ("最高级" if word.startswith("无") else "绝对化") \
                if new_key == "absolute" else subclass
            new_role = "偏激词" if new_key == "absolute" else role
            entries[i] = (new_key, new_sub, word, new_role, source)
        elif key == "certainty" and word in role_override:
            entries[i] = (key, subclass, word, role_override[word], source)

    dims = {k: {} for k, _ in DIMS}
    seen = {}
    conflicts = []
    for key, subclass, word, role, source in entries:
        prev = seen.get(word)
        if prev is not None and prev != key:
            # 主词表先处理，其归属为准；LM 只补充新词，不覆盖已有归属
            conflicts.append((word, prev, key))
            continue
        if word in dims[key]:
            continue
        seen[word] = key
        dims[key][word] = {"subclass": subclass, "role": role, "source": source}
        if key == "certainty" and source == "LM汇总":
            lm_added[subclass] += 1

    doc = {
        "version": "1.0.0",
        "provenance": {
            "主词表": "resources/任务5词表.xlsx（AI 生成，含 5 维度 185 词条）",
            "补充": "resources/LM词表中文校准对照表.xlsx 的「汇总」表 15–16 行",
            "说明": "本词表为项目自建中文种子表，未经独立人工效度验证；"
                    "来源标注见各词的 source 字段。",
            "裁定": ["「绝对」跨维度重复，归入绝对化，从确定性词移除"],
            "PPMI扩展审核": {
                "文件": "resources/种子审核结论.json",
                "收入词数": len(accepted),
                "收入明细": {w: sub for _, sub, w, _, _ in accepted},
            },
            "确定性词表校准": {
                "文件": "resources/确定性词表校准.json",
                "依据": cal["_依据"],
                "移除词数": n_dropped,
                "改判词数": sum(1 for w in reclass if w in RULING),
                "移除明细": {w: r for w, r in sorted(drop.items())},
                "改判明细": dict(sorted(reclass.items())),
            },
        },
        "dimensions": {k: {"label": label, "words": dims[k]} for k, label in DIMS},
    }
    args.output.write_text(json.dumps(doc, ensure_ascii=False, indent=2) + "\n",
                           encoding="utf-8")

    print(f"词表已生成 → {args.output}")
    total = 0
    for k, label in DIMS:
        n = len(dims[k])
        total += n
        roles = Counter(v["role"] for v in dims[k].values())
        print(f"  {label:<12} {n:>3} 词   偏激词 {roles['偏激词']} / 参照词 {roles['参照词']}")
    print(f"  合计 {total} 词")
    print(f"  LM 补充进确定性强度：{dict(lm_added)}")
    print(f"  确定性词表校准：移除 {n_dropped} 词、改判 {len(reclass)} 词"
          f"（依据 {args.calibration.name}）")
    if accepted:
        by_sub = Counter(sub for _, sub, _, _, _ in accepted)
        print(f"  PPMI 扩展人工审核：收入 {len(accepted)} 词 {dict(by_sub)}")
    if conflicts:
        print(f"  ⚠ 跨维度重复（主词表归属优先，LM 未覆盖）：{conflicts}")


if __name__ == "__main__":
    main()
