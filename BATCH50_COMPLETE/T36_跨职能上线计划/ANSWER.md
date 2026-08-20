===== BEGIN T36 =====
STATUS: COMPLETE
ORIGINAL_DELIVERABLE: 项目计划表+甘特+RACI+风险登记册
PROXY_DELIVERABLE: 本题已提供原格式办公文件（见本文件夹）

## ANSWER


## DELIVERABLE_FILES
本题正式交付物已放在本文件夹内，可用 PowerPoint / Excel / Word 直接打开：
  - T36_Launch_Plan.xlsx
下面的正文是同一套内容的说明、公式与自检，与二进制文件一致。


### FILE: Launch_Plan.xlsx
项目开始 2026-09-01（周二）。工作日排程，不考虑公共假期。后置任务在所有前置完成后的下一工作日开始。周末不工作。

**工作日推演**  
A Requirements 3d：09-01 至 09-03  
B Design 4d（后 A）：09-04, 09-07, 09-08, 09-09（跨周末）  
C Legal 5d（后 A）：09-04 至 09-10  
D Build 8d（后 B）：09-10,11,14,15,16,17,18,21 → 09-21  
E Content 6d（后 B）：09-10,11,14,15,16,17 → 09-17  
F UAT 4d（后 D,E → 等 D 的 09-21）：09-22 至 09-25  
G Training 3d（后 C,F → 等 F 的 09-25）：09-28 至 09-30（09-26/27 周末）  
H Launch 1d（后 G）：**2026-10-01**  
上线日 2026-10-01 刚好命中，无缓冲。

**关键路径：** A–B–D–F–G–H = 3+4+8+4+3+1 = **23 个工作日**。E、C 不在关键路径（C 早于 F 完成；E 早于 D 完成）。

### 计划表

| ID | 任务 | 工期 | 前置 | 开始 | 结束 | Owner |
|---|---|---:|---|---|---|---|
| A | Requirements | 3 | — | 2026-09-01 | 2026-09-03 | Product |
| B | Solution design | 4 | A | 2026-09-04 | 2026-09-09 | Engineering |
| C | Legal review | 5 | A | 2026-09-04 | 2026-09-10 | Legal |
| D | Build | 8 | B | 2026-09-10 | 2026-09-21 | Engineering |
| E | Content | 6 | B | 2026-09-10 | 2026-09-17 | Marketing |
| F | UAT | 4 | D,E | 2026-09-22 | 2026-09-25 | Product |
| G | Training | 3 | C,F | 2026-09-28 | 2026-09-30 | Enablement |
| H | Launch | 1 | G | 2026-10-01 | 2026-10-01 | Program Lead |

### 简易甘特（工作日；█=占用）
```
         1  2  3  4  7  8  9 10 11 14 15 16 17 18 21 22 23 24 25 28 29 30 10/1
A        █  █  █
B                 █  █  █  █
C                 █  █  █  █  █
D                            █  █  █  █  █  █  █        █
E                            █  █  █  █  █  █
F                                                          █  █  █  █
G                                                                      █  █  █
H                                                                               █
```

### RACI（每项仅一个 A）
R=执行，A=问责（唯一），C=咨询，I=知情。Sponsor=COO。信息安全在 UAT 期间 Consulted。

| 任务 | Product | Eng | Legal | Mkt | Enablement | Program Lead | COO | InfoSec |
|---|---|---|---|---|---|---|---|---|
| A Requirements | A/R | C | I | I | I | I | I | I |
| B Design | C | A/R | I | C | I | I | I | C |
| C Legal | I | I | A/R | I | I | I | C | I |
| D Build | C | A/R | I | I | I | I | I | C |
| E Content | C | I | C | A/R | I | I | I | I |
| F UAT | A/R | R | I | C | C | C | I | **C** |
| G Training | C | I | I | C | A/R | I | I | I |
| H Launch | C | C | C | C | C | **A/R** | C | I |

Marketing 9/14–15 休假：E 正在进行，副手覆盖，R 仍属 Marketing（A 不转移）。

### 风险登记册

| ID | 风险 | 影响 | 缓解 |
|---|---|---|---|
| R1 | Legal 可能延迟 3 天 | C 结束 09-15 仍早于 F 结束 09-25，**不碰关键路径**；若再延迟超过 F，G 将推过 10-01 | 09-10 检查 C；超期则 Training 预读非法律内容 |
| R2 | UAT 环境只有一个 | D 与 E 并行无冲突；F 不能再拆并行 → 4 天无备份环境 | 09-21 前完成环境独占日历；缺陷修复不得另开第二环境 |
| R3 | Marketing 9/14–15 休假 | E 落在该两天 | 副手覆盖；09-13 前完成需本人签字的素材 |
| R4 | 10-01 零缓冲 | 关键路径任何 +1 工作日即错过上线 | 对 D/F 设每日站会；范围冻结在 B 结束时 |
| R5 | 10-01 周四，若被当假期（本题不考虑） | — | 题设忽略公共假期 |

## ASSUMPTIONS_AND_UNKNOWNS
工期含开始日。未考虑时区。InfoSec 仅在 UAT Consulted，其他阶段未规定。

## UNMET_ARTIFACT_REQUIREMENTS
无法生成可打开 XLSX/DOCX 甘特对象。

## SELF_CHECK
- 依赖未违反；H=2026-10-01。每项一个 A。关键路径 ABD FGH。

===== END T36 =====
