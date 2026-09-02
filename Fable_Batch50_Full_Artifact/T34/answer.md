===== BEGIN T34 =====
STATUS: COMPLETE
ORIGINAL_DELIVERABLE: 可编辑 XLSX（24 个月租赁负债、利息、本金、ROU 折旧、年度汇总）
PROXY_DELIVERABLE: 真实可编辑文件：T34_IFRS16租赁摊销表.xlsx（附本说明）

## ANSWER

### FILE: T34_IFRS16租赁摊销表.xlsx
工作表：**Assumptions**（全部假设与初始计量公式）、**Schedule**（24 行：Month、Period End Date、Opening Liability、Interest、Payment、Principal、Closing Liability、ROU Depreciation、ROU Carrying Amount、Year）、**Annual Summary**（2027/2028 利息、现金支付、折旧、期末负债）、**Checks**（7 项勾稽）。

关键公式：月率 `=(1+B6)^(1/12)-1`；初始负债 `=B4*(1-(1+B7)^(-B5))/B7`；利息 `=C4*Assumptions!$B$7`；本金 `=E4-D4`；期末 `=C4-F4`；折旧 `=Assumptions!$B$10`；年度汇总 `=SUMIFS(Schedule!$D$4:$D$27,Schedule!$J$4:$J$27,A3)`。

预期结果（元，显示到元）：
| 项目 | 数值 |
|---|---|
| 月折现率 | 0.486755% |
| 初始租赁负债 = ROU 初始值 | 2,259,937 |
| 月折旧 | 94,164 |
| 第 1 月利息 / 本金 / 期末负债 | 11,000 / 89,000 / 2,170,937 |
| 2027 利息 / 现金支付 / 折旧 / 期末负债 | 102,943 / 1,200,000 / 1,129,968 / 1,162,880 |
| 2028 利息 / 现金支付 / 折旧 / 期末负债 | 37,120 / 1,200,000 / 1,129,968 / -0（≈0） |
| 利息合计 | 140,063（= 2,400,000 − 2,259,937） |

## ASSUMPTIONS_AND_UNKNOWNS

- 无增值税、初始直接成本、预付款、激励或复原义务，按题内简化。
- 期末日期以 EOMONTH 计算，仅用于年度归类。

## UNMET_ARTIFACT_REQUIREMENTS

无

## SELF_CHECK

- (1+0.00486755)^12 = 1.06；PV = 100,000×(1−1.06^−2)/0.00486755 = 2,259,936.67；折旧 94,164.03/月。
- 勾稽：第 24 月期末负债 -0.0000（≈0）；付款 2,400,000 = 初始负债 2,259,937 + 利息 140,063；工作簿 Checks 表 7 项公式检查。
- 已用 LibreOffice 重算核对（见批次验证记录）。

===== END T34 =====
