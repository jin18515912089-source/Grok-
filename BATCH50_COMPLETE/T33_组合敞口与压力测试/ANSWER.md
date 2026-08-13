===== BEGIN T33 =====
STATUS: COMPLETE_TEXT_PROXY
ORIGINAL_DELIVERABLE: 中文风险快照（敞口、压力P&L、两项对冲）
PROXY_DELIVERABLE: 完整风险快照

## ANSWER

NAV = 5,000 万美元 = 50 百万美元。头寸单位百万美元。

| Position | MV | Beta | Sector | |MV| | beta×MV |
|---|---:|---:|---|---:|---:|
| Long A | +20 | 1.3 | Tech | 20 | 26.0 |
| Long B | +15 | 0.8 | Health | 15 | 12.0 |
| Long C | +10 | 1.1 | Tech | 10 | 11.0 |
| Short D | −12 | 1.2 | Tech | 12 | −14.4 |
| Short E | −8 | 0.6 | Consumer | 8 | −4.8 |
| 合计 | **+25** | | | **65** | **29.8** |

**总敞口 Gross** = 20+15+10+12+8 = **65**（绝对值之和）  
Gross/NAV = 65/50 = **130%**  
**净敞口 Net** = 20+15+10−12−8 = **+25**  
Net/NAV = 25/50 = **50%**  
**Beta 调整净敞口** = 29.8 / 50 = **0.596**（59.6%）

**行业敞口（带符号）**  
- Tech：20+10−12 = **+18**（占 NAV 36%）  
- Health：+15（30%）  
- Consumer：−8（−16%）  
行业净敞口之和 = 18+15−8=25，与 Net 一致。

**压力情景 P&L**（Tech −10%，Health −5%，Consumer +4%，其余 0）  
A: 20×(−10%) = −2.00  
B: 15×(−5%) = −0.75  
C: 10×(−10%) = −1.00  
D: −12×(−10%) = +1.20  
E: −8×(+4%) = −0.32  
**合计 −2.87 百万美元 = −5.74% NAV**

Tech 净 18×(−10%)=−1.80；Health −0.75；Consumer 空头在板块上涨时亏损 −0.32。

**两项对冲建议**
1. **减 Tech 净多：** 再空 8–10 百万美元高 beta Tech（或买入保护性看跌），使 Tech 净敞口接近 0。粗算：若再空 10、beta 1.2，Tech 净 8，压力 Tech 腿约 −0.80 而非 −1.80。须控制 Gross 不要再大幅上升，或先减 Long A 一部分替代加空。  
2. **覆盖 Consumer 空头或降低其在“消费反弹”情景的杀伤：** Short E 在 Consumer +4% 时亏 0.32；若该空头是对冲消费而非独立观点，可用 beta 更匹配的空头替换，或买回部分 E、用期权定义最大损失。同时 Health 净多 15 在 −5% 情景亏 0.75，若不想加总风险，优先降 Tech 而非再空 Health。

## ASSUMPTIONS_AND_UNKNOWNS
压力为板块平行冲击，无特异质。Beta 为给定。对冲工具流动性未知。

## UNMET_ARTIFACT_REQUIREMENTS
无。

## SELF_CHECK
- Gross 65、Net 25、beta 29.8/50=0.596、压力 −2.87 复核。

===== END T33 =====
