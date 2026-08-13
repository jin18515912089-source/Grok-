===== BEGIN T39 =====
STATUS: COMPLETE_TEXT_PROXY
ORIGINAL_DELIVERABLE: 供应商评分表+采购建议+谈判要点
PROXY_DELIVERABLE: 完整评分与建议

## ANSWER

权重：Cost 30%，Quality 25%，Delivery 20%，Security 15%，Sustainability 10%。  
Cost Score = 最低**合格**报价 / 该供应商报价 × 100。  
Security <80 **不可豁免淘汰**。预算 1.25m。委员会可谈价格，不能改 Security 门槛。

**资格：** Y Security=70 <80 → **淘汰**，不参与最低价、不参与总分排名。合格：X（95）、Z（90）。最低合格价 = X 的 1.00。

| | X | Y | Z |
|---|---:|---:|---:|
| Cost $m | 1.00 | 0.85 | 1.20 |
| Cost Score | 1.00/1.00×100=**100.00** | 淘汰 | 1.00/1.20×100=**83.33** |
| Quality | 85 | 75 | 95 |
| Delivery | 70 | 90 | 80 |
| Security | 95 | 70 | 90 |
| Sustainability | 60 | 80 | 90 |
| 加权总分 | 100×0.30+85×0.25+70×0.20+95×0.15+60×0.10 = **85.50** | DQ | 83.33×0.30+95×0.25+80×0.20+90×0.15+90×0.10 = **87.25** |

Y 即使报价最低、若强行计分也会因 Security 出局；不得因任何总分绕过门槛。

**推荐：Vendor Z**（合格者中总分 87.25 > X 85.50）。两者均 ≤1.25 预算。X Delivery 70、Sustainability 60 明显拖后腿；Z 成本分低但质量/交付/可持续更高。

**谈判要点（对 Z）**
1. 价格：1.20 → 向 1.00–1.10 靠拢。若谈到 1.00，Cost Score=100，总分升至 30+23.75+16+13.5+9=**92.25**。  
2. 实施里程碑与 Delivery 80 的短板：把延期扣款写入合同。  
3. Security 90 已过线，维持门槛，不换条款降安全。  
4. 不重新纳入 Y：除非 Y 使 Security≥80（本次分数已定，不能改门槛）。

不推荐 X 作首选；若 Z 价格不让且 Delivery 合同保护不足，可把 X 作备选，重点谈 Delivery。

## ASSUMPTIONS_AND_UNKNOWNS
非成本分为委员会给定。Y 无补测机会（题未给）。Cost 用合格最低价，不用已淘汰的 0.85。

## UNMET_ARTIFACT_REQUIREMENTS
无。

## SELF_CHECK
- Y 淘汰。Z 87.25、X 85.50。未用 Y 当最低价（否则 X 成本分=85，会扭曲）。

===== END T39 =====
