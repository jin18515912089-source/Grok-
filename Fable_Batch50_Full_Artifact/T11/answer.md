===== BEGIN T11 =====
STATUS: COMPLETE
ORIGINAL_DELIVERABLE: 可编辑 XLSX（Raw/Clean/Summary/Exceptions）
PROXY_DELIVERABLE: 真实可编辑文件：T11_销售数据清洗与汇总.xlsx（附本说明）

## ANSWER

### FILE: T11_销售数据清洗与汇总.xlsx
四个工作表：**Raw**（15 行原始数据 + 重复检测公式 Dup_Key/Dup_Count/Keep_Flag）、**Clean**（14 行，Region/Rep/Net_Revenue 均为引用 Raw 的公式；Rep 映射表在 N:O）、**Summary**（Region/Rep/Product 三张 SUMIFS/COUNTIFS 汇总 + 柱形图 + 勾稽校验行）、**Exceptions**（17 条异常记录 + 2 条公式核对）。

关键公式（可复制）：
- 重复标记：`=IF(COUNTIF($I$2:I2,I2)=1,"Keep","Drop-exact duplicate")`
- Region：`=IF(TRIM(Raw!C2)="","Unknown",PROPER(TRIM(Raw!C2)))`
- Rep：`=IFERROR(VLOOKUP(TRIM(Raw!D2),$N$2:$O$9,2,FALSE),TRIM(Raw!D2))`
- 净收入：`=IF(I2="Completed",G2*H2,IF(I2="Refund",-ABS(G2)*H2,0))`
- 区域汇总：`=SUMIFS(Clean!$J$2:$J$15,Clean!$D$2:$D$15,A3)`；订单数 `=COUNTIFS(Clean!$D$2:$D$15,A3)`

预期结果（手工核算，供对照）：
| Region | Net_Revenue | Orders |
|---|---|---|
| North | 3,400 | 4 |
| South | 1,960 | 3 |
| East | 1,060 | 3 |
| West | 1,360 | 3 |
| Unknown | 1,400 | 1 |
| **Total** | **9,180** | **14** |

## ASSUMPTIONS_AND_UNKNOWNS

- 日期 7/1/2026 按 M/D/YYYY 解读为 2026-07-01（列入 Exceptions 待确认）。
- 1009 缺失区域的真实归属未知，计入 Unknown。
- 1001 与 1010 视为不同订单（Order_ID 与日期不同），未删除。

## UNMET_ARTIFACT_REQUIREMENTS

无

## SELF_CHECK

- 手工核算：North 1200+1000+1200+0=3400；South 960+0+1000=1960；East −240+800+500=1060；West 1000+720−360=1360；Unknown 1400；合计 9180；订单 14（15−1 重复）。
- 工作簿内置勾稽：Summary!B9/C9 = 区域合计 − Clean 合计，应为 0；Exceptions 底部两条公式核对重复行数=1、Unknown 行数=1。
- 工作表数量 = 4（Raw/Clean/Summary/Exceptions），Rep 映射表放在 Clean!N:O 未新增工作表；Summary 含 1 张柱形图。
- 公式值已用 LibreOffice 重算核对（见批次验证记录）。

===== END T11 =====
