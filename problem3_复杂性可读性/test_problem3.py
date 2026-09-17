import json
import tempfile
import unittest
from pathlib import Path
import problem3_complexity as m

class MeasurementTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.primary=m.load_chars(m.HERE/'dict/现代汉语常用字表_一级常用字2500.txt',2500)
        cls.secondary=m.load_chars(m.HERE/'dict/现代汉语常用字表_次常用字1000.txt',1000)

    def test_real_character_tables(self):
        self.assertFalse(self.primary & self.secondary)
        self.assertEqual(len(self.primary | self.secondary),3500)

    def test_numbers_do_not_create_boundaries(self):
        text='EPS为0.64元，收入1,234,567.89万元。净利增长31.2%。'
        self.assertEqual(len(m.split_sentences(text)),2)
        self.assertEqual(len(m.split_clauses(m.split_sentences(text)[0])),2)

    def test_wrap_and_explicit_boundary(self):
        a,_=m.clean('公司利润\n持续提高。产品销量增长。')
        b,_=m.clean('公司利润持续提高。产品销量增长。')
        self.assertEqual(a,b)
        c,_=m.clean('公司利润持续提高<?>产品销量增长')
        self.assertEqual(len(m.split_sentences(c)),2)

    def test_full_width_decimal(self):
        text,_=m.clean('收入１，２３４．５万元，利润增长。')
        self.assertEqual(len(m.split_clauses(text)),2)

    def test_missing_and_placeholder_are_not_zero_scores(self):
        for text,state in [(None,'文本为空'),('无','占位文本'),('','文本为空')]:
            r=m.extract(text,'',self.primary,self.secondary)
            self.assertEqual(r['状态'],state)
            self.assertIsNone(r['复杂性_代理'])

    def test_company_exclusion_only_affects_lexicon(self):
        text='龘龘公司利润增长。'
        a=m.extract(text,'龘龘',self.primary,self.secondary)
        b=m.extract(text,'',self.primary,self.secondary)
        self.assertEqual(a['汉字数'],b['汉字数'])
        self.assertEqual(a['字表统计汉字数'],b['字表统计汉字数']-2)
        self.assertEqual(a['表外字数'],b['表外字数']-2)

    def test_familiarity_conservation(self):
        r=m.extract('一乙匕龘公司营业收入增加。','',self.primary,self.secondary)
        self.assertEqual(r['一级常用字数']+r['次常用字数']+r['表外字数'],r['字表统计汉字数'])

    def test_tie_aware_frozen_percentiles(self):
        ref={'features':{k:m.distribution([1,2,2,3]) for k in m.FEATURES}}
        dist=m.prepare_reference(ref)[m.FEATURES[0]]
        self.assertEqual(m.rank(2,dist),50)
        self.assertEqual(m.rank(0,dist),0)
        self.assertEqual(m.rank(4,dist),100)
        restored=m.prepare_reference(json.loads(json.dumps(ref)))
        self.assertEqual(m.rank(2,restored[m.FEATURES[0]]),50)

    def test_insufficient_text_not_scored(self):
        r=m.extract('公司利润增长。','',self.primary,self.secondary)
        self.assertEqual(r['状态'],'短文本')
        self.assertIsNone(m.score(r,{})['复杂性_代理'])

    def test_added_clause_increases_structure(self):
        simple='公司今年通过改进生产流程提高生产效率并持续推进新的产品开发工作。'*5
        complex_text=simple.replace('。','，公司未来将继续扩大市场份额并完善销售体系。')
        a=m.extract(simple,'',self.primary,self.secondary)
        b=m.extract(complex_text,'',self.primary,self.secondary)
        self.assertGreater(b['句均汉字数'],a['句均汉字数'])
        self.assertGreater(b['句均分段数'],a['句均分段数'])

if __name__=='__main__':unittest.main()
