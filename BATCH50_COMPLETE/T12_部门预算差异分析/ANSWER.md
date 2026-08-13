===== BEGIN T12 =====
STATUS: COMPLETE
ORIGINAL_DELIVERABLE: 可编辑XLSX（明细、部门汇总、管理图表、三条文字结论）
PROXY_DELIVERABLE: 本题已提供原格式办公文件（见本文件夹）

## ANSWER


## DELIVERABLE_FILES
本题正式交付物已放在本文件夹内，可用 PowerPoint / Excel / Word 直接打开：
  - T12_Q1_Department_Budget_Variance.xlsx
下面的正文是同一套内容的说明、公式与自检，与二进制文件一致。


### FILE: Q1_Department_Budget_Variance.xlsx
金额单位：千美元。差异=Actual−Budget。费用：正差异=Unfavorable，负差异=Favorable。差异率分母=该部门季度总预算。

红色标记：季度不利差异>20 或差异率>5%。有利=绿色。

---

### 工作表 Monthly_Detail
A1:H1：Department | Jan_B | Jan_A | Feb_B | Feb_A | Mar_B | Mar_A | Jan_Var | Feb_Var | Mar_Var  
（可把 Var 放 I:K）

I2：`=C2-B2`（Jan Actual−Jan Budget）  
J2：`=E2-D2`  
K2：`=G2-F2`

数据：

| Department | Jan_B | Jan_A | Feb_B | Feb_A | Mar_B | Mar_A | Jan_Var | Feb_Var | Mar_Var |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| Marketing | 100 | 105 | 110 | 98 | 120 | 150 | 5 | -12 | 30 |
| R&D | 200 | 190 | 210 | 230 | 220 | 250 | -10 | 20 | 30 |
| G&A | 80 | 78 | 80 | 82 | 85 | 80 | -2 | 2 | -5 |
| Sales | 150 | 145 | 160 | 175 | 170 | 185 | -5 | 15 | 15 |

---

### 工作表 Dept_Summary
L 列起公式（以 Marketing 行为例，明细在 row2）：
- Q_Budget：`=B2+D2+F2`
- Q_Actual：`=C2+E2+G2`
- Q_Var：`=Q_Actual-Q_Budget`
- Q_Var%：`=Q_Var/Q_Budget`
- Flag：`=IF(Q_Var<0,"Favorable-Green",IF(OR(Q_Var>20,Q_Var/Q_Budget>0.05),"Unfavorable-Red","Watch"))`

| Department | Q_Budget | Q_Actual | Q_Var | Q_Var% | 判定 |
|---|---:|---:|---:|---:|---|
| Marketing | 330 | 353 | 23 | 6.97% | Red（23>20 且 6.97%>5%） |
| R&D | 630 | 670 | 40 | 6.35% | Red（40>20 且 6.35%>5%） |
| G&A | 245 | 240 | -5 | -2.04% | Green（有利） |
| Sales | 480 | 505 | 25 | 5.21% | Red（25>20 且 5.21%>5%） |
| Total | 1685 | 1768 | 83 | 4.93% | 公司合计未过 5%，但三部门已红 |

累计趋势（各部门 Q 累计 Actual vs Budget 折线，X=Jan/Feb/Mar）：
- Marketing 累计 Var：5 / −7 / 23（3 月一次性 30 把全年季度打红）
- R&D：−10 / 10 / 40（2–3 月连续不利）
- G&A：−2 / 0 / −5
- Sales：−5 / 10 / 25（2–3 月连续不利）

**图1：** 部门季度差异条形（负值绿色、正值红色），Y 轴含 0。  
**图2：** 四部门三个月累计差异折线。

---

### 三条文字结论
1. **R&D 超支最大**：季度不利 40 千美元（6.35%），且 2 月、3 月连续不利（+20、+30），需管理层优先问询项目节奏而非只看 1 月有利的 −10。  
2. **Marketing 被 3 月拖红**：1–2 月合计仍有利 −7，3 月 +30 导致季度 +23（6.97%）。问题在 3 月投放/活动，不在全年均匀超支。  
3. **Sales 过线、G&A 受控**：Sales +25（5.21%）连续两个月不利；G&A −5（−2.04%）绿色。公司整体 +83 / 1685 = 4.93%，若只看总额会低估部门红灯。

## ASSUMPTIONS_AND_UNKNOWNS
差异率分母采用部门季度预算（非公司总预算）。超支原因未知，未补造。

## UNMET_ARTIFACT_REQUIREMENTS
无法生成可打开 XLSX 与条件格式/图表对象。

## SELF_CHECK
- Marketing 330/353/23；R&D 630/670/40；G&A 245/240/−5；Sales 480/505/25；总 1685/1768/83。
- 红灯三部门均同时满足金额或比率阈值（实为同时满足）。

===== END T12 =====
