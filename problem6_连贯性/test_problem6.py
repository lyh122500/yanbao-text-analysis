import unittest

import numpy as np
from scipy.sparse import csr_matrix

from problem6_coherence import analyze_document, content_tokens, percentile, split_report


class TestCoherence(unittest.TestCase):
    def test_sections_and_soft_wrap(self):
        s, sec, duplicates = split_report("第一句换\n行。第二句<?>第三句！")
        self.assertEqual(s, ["第一句换行", "第二句", "第三句"])
        self.assertEqual(sec, [1, 1, 2])
        self.assertEqual(duplicates, 0)

    def test_placeholder(self):
        self.assertEqual(split_report("无"), ([], [], 0))

    def test_exact_duplicate_sentence_removed(self):
        s, sec, duplicates = split_report("同一句内容。另一句内容<?>同一句内容。")
        self.assertEqual(s, ["同一句内容", "另一句内容"])
        self.assertEqual(sec, [1, 1])
        self.assertEqual(duplicates, 1)

    def test_local_order_has_positive_excess(self):
        # Consecutive sentences form two close pairs; distant pairs are unrelated.
        matrix = np.array([[1, .8, .0, .0], [.8, 1, .7, .0], [.0, .7, 1, .8], [.0, .0, .8, 1]])
        tokens = [{"甲", "共同"}, {"共同", "乙"}, {"乙", "连接"}, {"连接", "丙"}]
        m, _, _ = analyze_document(["a", "b", "c", "d"], [1]*4, matrix, tokens)
        self.assertGreater(m["语义局部超额"], 0)
        self.assertGreater(m["词汇局部超额"], 0)
        self.assertEqual(m["相邻共享实词比例"], 1)

    def test_connective_and_reference(self):
        s = ["行业景气回升", "因此公司盈利改善", "该业务仍有空间"]
        m, _, _ = analyze_document(s, [1, 1, 1], np.eye(3), [content_tokens(x) for x in s])
        self.assertEqual(m["显式连接覆盖率"], .5)
        self.assertEqual(m["指代承接覆盖率"], .5)

    def test_tie_aware_percentile(self):
        actual = percentile(np.array([1., 2., 2., 3.]), np.array([1., 2., 2., 3.]))
        np.testing.assert_allclose(actual, [0, 50, 50, 100])

    def test_single_sentence(self):
        m, _, _ = analyze_document(["公司盈利"], [1], np.ones((1, 1)), [{"盈利"}])
        self.assertIsNone(m["相邻句语义相似度"])
        self.assertIsNone(m["首末句语义相似度"])


if __name__ == "__main__":
    unittest.main()
