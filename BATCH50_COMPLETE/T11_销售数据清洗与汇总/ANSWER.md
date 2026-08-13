===== BEGIN T11 =====
STATUS: COMPLETE
ORIGINAL_DELIVERABLE: 可编辑XLSX（Raw/Clean/Summary/Exceptions 四表）
PROXY_DELIVERABLE: 本题已提供原格式办公文件（见本文件夹）

## ANSWER


## DELIVERABLE_FILES
本题正式交付物已放在本文件夹内，可用 PowerPoint / Excel / Word 直接打开：
  - T11_Sales_Orders_Clean_Summary.xlsx
下面的正文是同一套内容的说明、公式与自检，与二进制文件一致。


### FILE: Sales_Orders_Clean_Summary.xlsx

**清洗规则落地**
- 完全重复行：Order 1007 出现两次，保留 1 条，删除 1 条。
- Rep：A. Chen / A Chen / Alex Chen → Alex Chen；Nina K. → Nina Ko。其余保持。
- Region：首字母大写（North/South/East/West）；缺失 → Unknown，并进 Exceptions。
- 日期统一为 ISO `YYYY-MM-DD`。
- Net_Revenue：Completed = Units*Unit_Price；Refund = Units*Unit_Price（题内 Units 已为负，净额为负）；Cancelled/Pending = 0。

---

### 工作表 Raw（15 行，含重复 1007）
字段 A1:H1：Order_ID | Date | Region | Rep | Product | Units | Unit_Price | Status  
数据按题面原样录入 1001–1014（1007 两行）。

---

### 工作表 Clean（14 行）
字段 A1:J1：Order_ID | Date | Region_Raw | Region | Rep_Raw | Rep | Product | Units | Unit_Price | Status | Net_Revenue  
（为可追踪，建议保留 Raw 对照列；最少需要标准化后字段+Net）

**推荐公式（数据从第2行起，假设标准化后 Region 在 D，Rep 在 F，Units H，Price I，Status J）：**

Region 标准化（若从 Raw!C2）：
`=IF(TRIM(Raw!C2)="","Unknown",UPPER(LEFT(TRIM(Raw!C2),1))&LOWER(MID(TRIM(Raw!C2),2,99)))`

Rep 标准化：
`=IFS(OR(Raw!D2="A. Chen",Raw!D2="A Chen",Raw!D2="Alex Chen"),"Alex Chen",OR(Raw!D2="Nina K.",Raw!D2="Nina Ko"),"Nina Ko",TRUE,Raw!D2)`

Net_Revenue（Completed/Refund/其他）：
`=IF(J2="Completed",H2*I2,IF(J2="Refund",H2*I2,0))`

Date 标准化示例：
`=IF(ISNUMBER(Raw!B2),TEXT(Raw!B2,"yyyy-mm-dd"),Raw!B2)`（需按单元格日期解析；文本日期手工统一如下）

**Clean 结果表（计算后值）**

| Order_ID | Date | Region | Rep | Product | Units | Unit_Price | Status | Net_Revenue |
|---:|---|---|---|---|---:|---:|---|---:|
| 1001 | 2026-07-01 | North | Alex Chen | Alpha | 10 | 120 | Completed | 1200 |
| 1002 | 2026-07-01 | North | Alex Chen | Beta | 5 | 200 | Completed | 1000 |
| 1003 | 2026-07-02 | South | Maria Li | Alpha | 8 | 120 | Completed | 960 |
| 1004 | 2026-07-02 | South | Maria Li | Gamma | 3 | 500 | Cancelled | 0 |
| 1005 | 2026-07-03 | East | Sam Wu | Alpha | -2 | 120 | Refund | -240 |
| 1006 | 2026-07-03 | East | Sam Wu | Beta | 4 | 200 | Completed | 800 |
| 1007 | 2026-07-04 | West | Nina Ko | Gamma | 2 | 500 | Completed | 1000 |
| 1008 | 2026-07-05 | West | Nina Ko | Alpha | 6 | 120 | Completed | 720 |
| 1009 | 2026-07-05 | Unknown | Jordan | Beta | 7 | 200 | Completed | 1400 |
| 1010 | 2026-07-06 | North | Alex Chen | Alpha | 10 | 120 | Completed | 1200 |
| 1011 | 2026-07-06 | North | Alex Chen | Beta | 5 | 200 | Pending | 0 |
| 1012 | 2026-07-07 | East | Sam Wu | Gamma | 1 | 500 | Completed | 500 |
| 1013 | 2026-07-07 | South | Maria Li | Beta | 5 | 200 | Completed | 1000 |
| 1014 | 2026-07-08 | West | Nina Ko | Alpha | -3 | 120 | Refund | -360 |

未删除 1001 与 1010：相似但 Date 不同，非完全重复。

---

### 工作表 Summary
按 Region 汇总（公式，Clean 表名为 `Clean`，Region 在 D，Net 在 K，Order_ID 在 A）：

| Region | Orders | Net_Revenue | Excel |
|---|---:|---:|---|
| East | 3 | 1060 | `=COUNTIF(Clean!D:D,"East")` / `=SUMIF(Clean!D:D,"East",Clean!K:K)` |
| North | 4 | 3400 | 同上 North |
| South | 3 | 1960 | 同上 South |
| West | 3 | 1360 | 同上 West |
| Unknown | 1 | 1400 | 同上 Unknown |
| Total | 14 | 9180 | `=COUNTA(Clean!A2:A15)` / `=SUM(Clean!K2:K15)` |

**图：** 簇状柱形图，X=Region（含 Unknown），Y=Net_Revenue；坐标从 0 起。结论：North 净收入最高（3400），Unknown 单笔 1400 需补区域，否则区域排名失真。

---

### 工作表 Exceptions
| Issue | Order_ID | Field | Raw_Value | Action |
|---|---|---|---|---|
| Missing Region | 1009 | Region | (blank) | 记 Unknown；需销售补全 |
| Duplicate dropped | 1007 | entire row | identical second row | 保留一条 |

可选：Refund 负 Units 不作为错误，因规则需要负净收入。

## ASSUMPTIONS_AND_UNKNOWNS
日期解析在纯文本代理中按字面规则转换为 ISO。1001 与 1010 是否业务重复未知，按规则保留。

## UNMET_ARTIFACT_REQUIREMENTS
无法生成可打开 XLSX，无法在 Excel 中验证公式与图表对象。

## SELF_CHECK
- 15→14 行；1007 只留 1。
- 净收入勾稽 1200+1000+960+0-240+800+1000+720+1400+1200+0+500+1000-360=9180。
- 区域合计 1060+3400+1960+1360+1400=9180。

===== END T11 =====
