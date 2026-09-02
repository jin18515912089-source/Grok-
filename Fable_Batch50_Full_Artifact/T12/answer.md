===== BEGIN T12 =====
STATUS: COMPLETE
ORIGINAL_DELIVERABLE: 可编辑 XLSX（明细、部门汇总、管理图表、三条结论）
PROXY_DELIVERABLE: 真实可编辑文件：T12_部门预算差异分析.xlsx（附本说明）

## ANSWER

### FILE: T12_部门预算差异分析.xlsx
四个工作表：**Detail**（月度预算/实际 + 月差异/季度差异/差异率/Fav-Unfav/Flag 公式，条件格式红绿）、**Dept Summary**（部门季度汇总 + 三个月累计差异）、**Charts**（部门差异柱形图 + 累计差异趋势折线图，数据源为公式）、**Conclusions**（三条结论 + 公式引用核对）。

核心公式：差异 `=L3-K3`（Actual−Budget）；差异率 `=M3/K3`（总预算分母）；状态 `=IF(M3>0,"Unfavorable",IF(M3<0,"Favorable","On budget"))`；标记 `=IF(AND(M3>0,OR(M3>20,N3>0.05)),"RED",IF(M3<0,"GREEN","WATCH"))`。

预期结果（千美元）：
| Department | Q1 Budget | Q1 Actual | Var | Var % | Flag |
|---|---|---|---|---|---|
| Marketing | 330 | 353 | +23 | +7.0% | RED |
| R&D | 630 | 670 | +40 | +6.3% | RED |
| G&A | 245 | 240 | −5 | −2.0% | GREEN |
| Sales | 480 | 505 | +25 | +5.2% | RED |
| Total | 1,685 | 1,768 | +83 | +4.9% | — |

## ASSUMPTIONS_AND_UNKNOWNS

- Sales 超支是否与收入增长相关：题内无收入数据，标为待确认。
- “WATCH”标记用于不利但未越阈值的情形（本题无部门落入）。

## UNMET_ARTIFACT_REQUIREMENTS

无

## SELF_CHECK

- Marketing 330/353/+23/7.0%；R&D 630/670/+40/6.3%；G&A 245/240/−5/−2.0%；Sales 480/505/+25/5.2%；合计 1685/1768/+83/4.9%（83/1685=4.93%）。
- 阈值判定：23>20 且 7.0%>5% → RED；40>20 → RED；25>20 且 5.2%>5% → RED；−5 → GREEN。
- 月度明细保留；差异全部为公式；工作簿内置核对：Dept Summary!D8 = 0、Conclusions!B10 = 3。

===== END T12 =====
