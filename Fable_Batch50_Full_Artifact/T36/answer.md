===== BEGIN T36 =====
STATUS: COMPLETE
ORIGINAL_DELIVERABLE: 项目计划表 + 简易甘特图 + RACI + 风险登记册（XLSX）
PROXY_DELIVERABLE: 真实可编辑文件：T36_跨职能上线计划.xlsx（附本说明）

## ANSWER

### FILE: T36_跨职能上线计划.xlsx
工作表：**Schedule**（WORKDAY 正推 ES/EF、倒推 LS/LF、NETWORKDAYS 浮动、关键路径标记、上线日校验）、**Gantt**（8/31–10/3 日历格公式甘特图，红=关键、蓝=非关键、灰=周末）、**RACI**（每项唯一 A，含 A 计数校验；InfoSec 在 UAT 为 C）、**Risk Register**（6 项风险，含基于排程浮动的分析）。

预期排程（工作日，不含周末）：
| ID | 任务 | 工期 | 前置 | 开始 | 完成 | 总浮动 | 关键 |
|---|---|---|---|---|---|---|---|
| A | Requirements | 3 | — | 09-01（二） | 09-03（四） | 0 | YES |
| B | Solution design | 4 | A | 09-04 | 09-09 | 0 | YES |
| C | Legal review | 5 | A | 09-04 | 09-10 | 11 | |
| D | Build | 8 | B | 09-10 | 09-21 | 0 | YES |
| E | Content | 6 | B | 09-10 | 09-17 | 2 | |
| F | UAT | 4 | D,E | 09-22 | 09-25 | 0 | YES |
| G | Training | 3 | C,F | 09-28 | 09-30 | 0 | YES |
| H | Launch | 1 | G | **10-01（四）** | 10-01 | 0 | YES |

关键路径 A→B→D→F→G→H，共 23 个工作日；计划上线日恰为 2026-10-01，零余量。Legal 延迟 3 天可被 C 的 11 天浮动吸收；Marketing 休假 2 天在 E 的 2 天浮动内，需副手覆盖。

## ASSUMPTIONS_AND_UNKNOWNS

- 不考虑公共假期，按题内规定。
- RACI 中 Sponsor 全程 I、InfoSec 在 B/D 为 C 为合理补充；题内仅规定 InfoSec 在 UAT 为 C。

## UNMET_ARTIFACT_REQUIREMENTS

无

## SELF_CHECK

- 手工日历核对：A 9/1–3；B 9/4,7,8,9；C 9/4,7,8,9,10；D 9/10,11,14,15,16,17,18,21；E 9/10,11,14,15,16,17；F 9/22–25；G 9/28–30；H 10/1。
- 依赖：F 于 max(D 9/21, E 9/17) 后一个工作日 9/22 开始；G 于 max(C 9/10, F 9/25) 后 9/28 开始。
- 浮动：C 9/11–9/25 共 11 个工作日；E 9/18、9/21 共 2 个工作日；其余 0。RACI 每行 A 计数 = 1（工作簿公式校验）。已用 LibreOffice 重算核对。

===== END T36 =====
