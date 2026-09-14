# 问题6：连贯性 coherence

**需求**（需求文档第 6 条）：衡量研报文本的连贯性。

## 怎么做（计划）

需求文档给出两条路线：

1. **方法一（自算衔接指标 + PCA）**：计算研报的指代衔接、连接词衔接、词汇衔接、
   首句与末句 embedding 相似度等多个维度，再 PCA 降维到单一连贯性指标。
   参考 Coh-Metrix（Graesser et al., 2004）的思路，但商业版价格高
   （约 2 万篇 12600 元），优先找开源实现（如
   [mkclairhong/cohmetrix](https://github.com/mkclairhong/cohmetrix)）；
2. **方法二（l2c-cohesion）**：若方法一找不到开源库，尝试
   [arXiv:2402.13583](https://arxiv.org/abs/2402.13583) 的
   [l2c-cohesion](https://github.com/mybluue/l2c-cohesion)（需确认是否付费）。

## 做了什么（待开发）

⏳ 尚未开始。

## 结果如何（待开发）

⏳ 尚无结果。
