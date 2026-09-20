import unittest
from problem5_extremity import Analyzer, normalized


class TestExtremity(unittest.TestCase):
    @classmethod
    def setUpClass(cls): cls.a = Analyzer()

    def test_domain_ambiguities(self):
        r, e = self.a.analyze("存在一定程度的压力，绝对估值不贵，收益有不确定性。")
        self.assertEqual(r["strong_count"], 0)
        self.assertEqual(r["counts"]["hedge"], 1)

    def test_rating_and_company_masks(self):
        r, e = self.a.analyze("唯一公司维持强烈推荐评级。我们非常看好公司。", ["唯一公司"])
        self.assertEqual(r["strong_count"], 1)
        self.assertEqual(sum(x["excluded_reason"] == "固定评级名称" for x in e), 1)

    def test_longest_match_and_negation(self):
        r, e = self.a.analyze("毫无疑问，不一定盈利，未必增长，不是极其乐观，不可能下跌。")
        self.assertEqual(r["counts"]["certainty"], 2)
        self.assertEqual(r["counts"]["hedge"], 2)
        self.assertEqual(r["counts"]["intensity"], 0)

    def test_context_and_no_cancellation(self):
        r, e = self.a.analyze("如果业绩超预期会带来前所未有的机会。利润可能大幅增长。公司必将成长。")
        self.assertEqual(r["extreme_count"], 1)
        self.assertEqual(r["strict_extreme_count"], 0)
        self.assertEqual(r["counts"]["magnitude"], 1)
        self.assertEqual(r["counts"]["certainty"], 1)

    def test_numeric_changes_not_extremism(self):
        r, e = self.a.analyze("利润大幅增长300%，销量达到历史最高，EPS为1.23元。")
        self.assertEqual(r["strong_count"], 0)
        self.assertEqual(r["counts"]["magnitude"], 1)
        self.assertEqual(r["counts"]["superlative"], 1)

    def test_offsets_soft_wrap(self):
        text = "公司极\n其看好<?>毫无\u00ad疑问，前所未有。"
        r, e = self.a.analyze(text)
        for hit in e:
            self.assertEqual(text[hit["start"]:hit["end"]], hit["raw_match"])
            self.assertEqual(normalized(hit["raw_match"])[0], hit["term"])
        self.assertEqual(r["counts"]["certainty"], 1)

    def test_attribution_question(self):
        r, e = self.a.analyze("管理层表示公司必将增长。公司真的必将增长？")
        self.assertEqual(r["strong_count"], 2)
        self.assertEqual(r["strict_strong_count"], 0)

    def test_disclaimer(self):
        r, e = self.a.analyze("本报告不保证信息绝对准确。公司必将增长。")
        self.assertEqual(r["counts"]["absolute"], 0)
        self.assertEqual(r["counts"]["certainty"], 1)
        self.assertEqual(r["sentences"], 1)

    def test_empty_vs_zero(self):
        r, _ = self.a.analyze("无")
        self.assertEqual(r["han"], 0)
        r, _ = self.a.analyze("公司发布三季度财务报告。")
        self.assertGreater(r["han"], 0)
        self.assertEqual(r["strong_count"], 0)

    def test_sentence_union(self):
        r, _ = self.a.analyze("公司必将取得前所未有的成绩，非常乐观。另一项业务保持稳定。")
        self.assertEqual(r["strong_count"], 3)
        self.assertEqual(r["strong_sentences"], 1)
        self.assertEqual(r["sentences"], 2)

    def test_new_corpus_false_positive_cases(self):
        r, _ = self.a.analyze("唯一一家企业受到异常气候影响，收入符合预期，占绝对比重。")
        self.assertEqual(r["strong_count"], 0)
        self.assertEqual(r["counts"]["hedge"], 0)
        r, _ = self.a.analyze("若干项目必将增长，公司绝对不会退出。")
        self.assertEqual(r["strict_strong_count"], 2)

    def test_cross_word_and_technical_terms(self):
        r, _ = self.a.analyze("西地那非预计上市，南北极高端旅游，积极大力开拓，非常规油气，十分之一。")
        self.assertEqual(r["strong_count"], 0)
        self.assertEqual(r["counts"]["hedge"], 1)
        r, _ = self.a.analyze("详见《大潮势不可挡》。")
        self.assertEqual(r["extreme_count"], 1)
        self.assertEqual(r["strict_extreme_count"], 0)

    def test_expansion_is_deterministic_and_leaves_review_blank(self):
        from expand_candidates import candidates
        sentences = [["非常", "乐观", "增长"], ["十分", "乐观", "增长"],
                     ["较为", "乐观", "增长"], ["较为", "乐观", "增长"],
                     ["可能", "下滑", "利润"], ["或许", "下滑", "利润"]] * 3
        categories = {"intensity": ["非常", "十分"], "hedge": ["可能", "或许"]}
        a, _ = candidates(sentences, categories, min_count=2, top_n=2)
        b, _ = candidates(sentences, categories, min_count=2, top_n=2)
        self.assertEqual(a, b)
        self.assertEqual(len(a), 4)
        for row in a:
            self.assertEqual(row["审核人"], "")
            self.assertEqual(row["是否收入词典"], "")
            self.assertNotIn(row["候选词"], ["非常", "十分", "可能", "或许"])


if __name__ == "__main__": unittest.main()
