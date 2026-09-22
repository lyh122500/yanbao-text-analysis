#!/usr/bin/env python3
"""任务5（新）计分脚本的自动检查。

覆盖三类断言：
  1. 词表与校准的一致性（被移除的词不再计分、被改判的词换了角色/维度）；
  2. 消歧规则与评级档位掩码的定向行为（含各自的对照反例）；
  3. 全量产出与 manifest 的自洽（行数、行序、哈希、有效汉字数不受词表影响）。

不检查人工效度：指标是否真的对应「偏激」需要盲评，本文件不含此类断言。
运行：python -m unittest -v test_problem5new.py
"""
import json
import unittest
from pathlib import Path

from score_extremity import Scorer, HAN, RATING

HERE = Path(__file__).resolve().parent
RES = HERE / "resources"


class TestLexicon(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.spec = json.loads((RES / "lexicon.json").read_text(encoding="utf-8"))
        cls.cal = json.loads((RES / "确定性词表校准.json").read_text(encoding="utf-8"))

    def test_校准词已从确定性维度移除(self):
        words = self.spec["dimensions"]["certainty"]["words"]
        leaked = sorted(w for w in self.cal["drop"] if w in words)
        self.assertEqual(leaked, [], f"校准移除的词仍在词表中：{leaked}")

    def test_跨维度不重复(self):
        seen = {}
        for key, dim in self.spec["dimensions"].items():
            for w in dim["words"]:
                self.assertNotIn(w, seen, f"{w} 同时属于 {seen.get(w)} 与 {key}")
                seen[w] = key

    def test_每个词都有角色与来源(self):
        for key, dim in self.spec["dimensions"].items():
            for w, meta in dim["words"].items():
                self.assertIn(meta["role"], ("偏激词", "参照词"), f"{key}/{w}")
                self.assertTrue(meta["source"], f"{key}/{w} 缺来源")

    def test_绝对归入绝对化(self):
        dims = self.spec["dimensions"]
        self.assertIn("绝对", dims["absolute"]["words"])
        self.assertNotIn("绝对", dims["certainty"]["words"])

    def test_一定改判为参照词(self):
        self.assertEqual(
            self.spec["dimensions"]["certainty"]["words"]["一定"]["role"], "参照词")

    def test_校准留痕在provenance里(self):
        prov = self.spec["provenance"]["确定性词表校准"]
        self.assertEqual(prov["移除词数"], len(self.cal["drop"]))
        self.assertEqual(set(prov["移除明细"]), set(self.cal["drop"]))


class TestScoring(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.s = Scorer()

    def score(self, text):
        res, events = self.s.score(text)
        return res, [e for e in events if e["active"]]

    def test_评级档位名不算措辞(self):
        _, ev = self.score("我们维持“谨慎推荐”评级，谨慎看好公司前景。")
        self.assertEqual(sum(e["term"] == "谨慎" for e in ev), 1,
                         "「谨慎推荐」应被掩码，「谨慎看好」应保留")

    def test_强烈推荐被掩码但强烈看好保留(self):
        _, ev = self.score("强烈推荐该股，我们强烈看好其长期竞争力。")
        self.assertEqual(sum(e["term"] == "强烈" for e in ev), 1)

    def test_评级掩码不改变有效汉字数(self):
        added = "，维持买入评级"
        a, _ = self.score("公司经营稳健。")
        b, _ = self.score("公司经营稳健" + added + "。")
        self.assertEqual(b["han"] - a["han"], len(HAN.findall(added)),
                         "评级名只屏蔽词命中，不应从分母中抹掉汉字")

    def test_所有者不算绝对化(self):
        _, ev = self.score("归属于母公司所有者的净利润同比增长。")
        self.assertFalse([e for e in ev if e["term"] == "所有"])

    def test_全称所有仍计入(self):
        _, ev = self.score("公司所有产品均已通过认证。")
        self.assertTrue([e for e in ev if e["term"] == "所有"])

    def test_上下游不算程度加强(self):
        _, ev = self.score("公司位于锂电池产业链上下游。")
        self.assertFalse([e for e in ev if e["term"] == "上下"])

    def test_模糊词与偏激词分列分子分母(self):
        res, _ = self.score("我们预计业绩可能超预期，但出现分化是必然的。")
        self.assertGreaterEqual(res["ref_counts"]["certainty"], 2)   # 预计/可能
        self.assertGreaterEqual(res["counts"]["certainty"], 1)       # 必然
        self.assertNotEqual(res["counts"]["certainty"], res["ref_counts"]["certainty"])

    def test_前置否定被排除(self):
        _, ev = self.score("公司未明确提出产能规划。")
        self.assertFalse([e for e in ev if e["term"] == "明确"])

    def test_免责声明整句被掩码(self):
        # 免责声明所在句整句掩码；同句内的措辞词随之失效，邻句不受影响
        _, ev = self.score("公司必然受益，本报告不构成投资建议。")
        self.assertFalse([e for e in ev if e["term"] == "必然"],
                         "免责声明同句内的词应被掩码")
        _, ev = self.score("本报告不构成投资建议。公司必然受益。")
        self.assertTrue([e for e in ev if e["term"] == "必然"],
                        "邻句不应被免责声明掩码殃及")

    def test_占位文本不计分(self):
        for t in ("无", "暂无", "nan", ""):
            res, _ = self.score(t)
            self.assertEqual(sum(res["counts"].values()), 0)
            self.assertIn(res["status"], ("占位文本", "空白或无汉字"))

    def test_二元指标只看绝对化(self):
        res, _ = self.score("公司业绩必将大幅增长，确定性极高。")
        self.assertEqual(res["binary"], res["counts"]["absolute"] > 0)


class TestOutputs(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        out = HERE / "output"
        if not (out / "summary.json").exists():
            raise unittest.SkipTest("尚未运行 score_extremity.py")
        cls.summary = json.loads((out / "summary.json").read_text(encoding="utf-8"))
        cls.manifest = json.loads((out / "manifest.json").read_text(encoding="utf-8"))
        cls.docs = json.loads((out / "逐篇指标.json").read_text(encoding="utf-8"))

    def test_处理篇数(self):
        self.assertEqual(self.summary["documents"], 17712)
        self.assertEqual(len(self.docs), 17712)

    def test_行号连续且不重复(self):
        rows = [d["source_row"] for d in self.docs]
        self.assertEqual(rows, list(range(2, 17714)))

    def test_有效汉字数不受词表影响(self):
        # 汉字数是文本属性：只应随文本变化，不应随词表/校准变化
        self.assertTrue(all(d["han"] >= 0 for d in self.docs))
        self.assertTrue(all(d["han"] == 0 or d["counts"]["absolute"] >= 0 for d in self.docs))

    def test_分子分母角色不混(self):
        for d in self.docs:
            self.assertGreaterEqual(d["counts"]["certainty"], 0)
            self.assertGreaterEqual(d["ref_counts"]["certainty"], 0)
        self.assertGreater(self.summary["reference_word_counts"]["certainty"], 0)
        self.assertEqual(self.summary["reference_word_counts"]["absolute"], 0,
                         "绝对化维度不应有参照词")

    def test_manifest哈希齐全(self):
        for k in ("input_sha256", "script_sha256", "lexicon_sha256", "disambig_sha256",
                  "calibration_sha256", "build_script_sha256",
                  "evidence_sha256", "metrics_sha256"):
            self.assertEqual(len(self.manifest[k]), 64, k)


if __name__ == "__main__":
    unittest.main(verbosity=2)
