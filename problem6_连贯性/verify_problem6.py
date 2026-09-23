#!/usr/bin/env python3
"""Verify row alignment, invariants, hashes and frozen-reference replay for problem 6."""
import argparse
import json
import subprocess
import sys
import tempfile
from collections import Counter
from pathlib import Path

import openpyxl

from problem6_coherence import HERE, sha256, dump_json


def workbook_rows(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    values = list(wb.active.iter_rows(values_only=True))
    wb.close()
    return values


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=HERE.parent / "综合分析结果.xlsx")
    parser.add_argument("--output-dir", type=Path, default=HERE / "output")
    parser.add_argument("--skip-replay", action="store_true")
    args = parser.parse_args()
    xlsx = args.output_dir / "问题6_连贯性结果.xlsx"
    rows = workbook_rows(xlsx)
    header, data = rows[0], rows[1:]
    index = {name: i for i, name in enumerate(header)}
    required = {"源Excel行号", "计算状态", "句数", "相邻转移数", "等权连贯分", "PCA连贯分"}
    checks = {
        "required_columns": required.issubset(index),
        "row_count_17712": len(data) == 17712,
        "source_row_alignment": all(row[index["源Excel行号"]] == i for i, row in enumerate(data, 2)),
        "transition_count": all(row[index["相邻转移数"]] == max(0, row[index["句数"]] - 1) for row in data),
        "score_range": all(row[index["PCA连贯分"]] is None or 0 <= row[index["PCA连贯分"]] <= 100 for row in data),
        "noneligible_score_blank": all(row[index["计算状态"]] == "可比较" or
                                       (row[index["PCA连贯分"]] is None and row[index["等权连贯分"]] is None)
                                       for row in data),
    }
    manifest = json.loads((args.output_dir / "manifest.json").read_text(encoding="utf-8"))
    checks["source_hash"] = manifest["source_sha256"] == sha256(args.input)
    checks["reference_hash"] = manifest["reference_sha256"] == sha256(args.output_dir / "reference.joblib")
    checks["output_hashes"] = all(sha256(args.output_dir / name) == digest
                                  for name, digest in manifest["outputs"].items())
    checks["status_total"] = sum(Counter(row[index["计算状态"]] for row in data).values()) == 17712
    replay_equal = None
    if not args.skip_replay:
        with tempfile.TemporaryDirectory(prefix="problem6_replay_") as tmp:
            command = [sys.executable, str(HERE / "problem6_coherence.py"), "--input", str(args.input),
                       "--output-dir", tmp, "--reference", str(args.output_dir / "reference.joblib")]
            subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            replay_equal = workbook_rows(Path(tmp) / xlsx.name) == rows
            replay_summary = json.loads((Path(tmp) / "summary.json").read_text(encoding="utf-8"))
            current_summary = json.loads((args.output_dir / "summary.json").read_text(encoding="utf-8"))
            checks["frozen_reference_workbook_exact"] = replay_equal
            checks["frozen_reference_summary_exact"] = replay_summary == current_summary
    result = {"checks": checks, "all_passed": all(checks.values()),
              "rows": len(data), "columns": len(header),
              "status_counts": dict(Counter(row[index["计算状态"]] for row in data)),
              "replay_performed": not args.skip_replay}
    dump_json(args.output_dir / "verification.json", result)
    if not result["all_passed"]:
        raise SystemExit(json.dumps(result, ensure_ascii=False, indent=2))
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
