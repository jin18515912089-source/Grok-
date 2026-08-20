===== BEGIN T38 =====
STATUS: COMPLETE_TEXT_PROXY
ORIGINAL_DELIVERABLE: 一周排班表+约束检查+备选
PROXY_DELIVERABLE: 主方案、验证、备选

## ANSWER

每班 2 人；周一至周五 AM/PM 共 10 班、20 人次。每班 ≥1 名 Bilingual；周二 PM、周四 PM ≥1 名 Senior。不可用时段不排；不超过 Max Shifts。尽量避免同一人同一天 AM+PM。周五 PM 希望至少一名本周累计（不含该班前）≤4 班者。

Bilingual：A,B,D。Senior：A,C。C 仅 PM。A 周三 PM 不可用。B 不可 Wed–Fri PM。D 仅周三至周五。

**主方案（无同日双班；均衡次优但全约束满足）**

| | AM | PM |
|---|---|---|
| 周一 | A（Bi,Sr）, E（En,Jr） | B（Bi,Jr）, C（En,Sr） |
| 周二 | A, E | B, C（含 Senior） |
| 周三 | A, B（双 Bi） | C, D（Bi；A 不可用） |
| 周四 | A, B | C, D（含 Senior） |
| 周五 | B, E | C, D |

班次统计：A4 / B5 / C5 / D3 / E3。Max：A6 B7 C5 D6 E8，均未超。C=5 触及其上限。

**约束检查**
- 人数：每班 2 人，共 20。  
- Bilingual：周一 AM A；PM B。周二同。周三 AM A/B；PM D。周四 AM A/B；PM D。周五 AM B；PM D。均 ≥1。  
- Senior：Tue PM C；Thu PM C。  
- 可用性：A 无 Wed PM；B 无 Wed–Fri PM；C 无 AM；D 无 Mon–Tue；E 全周。均未违规。  
- 同日双班：**无**。  
- 周五 PM：C 在该班前已排 Mon–Thu PM 共 4 班，D 共 2 班；**至少一人 ≤4**（两人均 ≤4）。C 上完周五 PM 后周总 5。

**无法完全均衡的原因：** 20/5=4，但 C 只能 PM（5 个 PM 中 Tue/Thu 必须有 Senior：A 若用于 AM 以避免双班，则 PM Senior 主要由 C 承担）。D 周一周二不可用。故 D/E 偏低、C 偏高，属可行域边界，不是漏排。

**备选（仍无双班）：** 周五 AM 改为 A+E，周五 PM 仍 C+D → 统计 A5 B4 C5 D3 E3。同样合法；主方案让 A=4 更接近均值。若必须给 D 更多班，可把周三 PM 改为 D+E、周四 AM 改为 D+E，但须重检每班 Bilingual（Wed PM DE 有 D；Thu AM DE 有 D）。

## ASSUMPTIONS_AND_UNKNOWNS
“本周累计不超过 4 班”按周五 PM **之前**的班次数。未要求最小化最高班次以外的目标。

## UNMET_ARTIFACT_REQUIREMENTS
无。

## SELF_CHECK
- 手检 10 班 Bilingual 与 Tue/Thu PM Senior 通过。无人超 Max。无同日双班。

===== END T38 =====
