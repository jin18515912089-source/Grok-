# 客服一周排班表（周一至周五 AM/PM，每班 2 人）

## 一、主方案

| 班次 | 周一 | 周二 | 周三 | 周四 | 周五 |
|---|---|---|---|---|---|
| **AM** | B、E | B、E | A、B | A、E | A、B |
| **PM** | A、C | A、C | D、C | D、C | D、E |

## 二、员工班次统计（共 20 班）

| Agent | 技能 | 班次数 | Max | 是否超限 |
|---|---|---|---|---|
| A | Bilingual, Senior | 5 | 6 | 否 |
| B | Bilingual, Junior | 4 | 7 | 否 |
| C | English, Senior | 4 | 5 | 否 |
| D | Bilingual, Junior | 3 | 6 | 否 |
| E | English, Junior | 4 | 8 | 否 |

**班次均衡：** 5/4/4/3/4，基本均衡。

## 三、约束检查

### 1. 每班至少 1 名 Bilingual

| 班次 | Bilingual 覆盖 |
|---|---|
| 周一 AM | B ✓ |
| 周一 PM | A ✓ |
| 周二 AM | B ✓ |
| 周二 PM | A ✓ |
| 周三 AM | A ✓ |
| 周三 PM | D ✓ |
| 周四 AM | A ✓ |
| 周四 PM | D ✓ |
| 周五 AM | A ✓ |
| 周五 PM | D ✓ |

### 2. 周二 PM 与周四 PM 至少 1 名 Senior

- 周二 PM：A（Senior）✓
- 周四 PM：C（Senior）✓

### 3. 可用性

| Agent | 可用性 | 排班 | 检查 |
|---|---|---|---|
| A | All except Wed PM | 未排 Wed PM | ✓ |
| B | All AM + Mon/Tue PM | 仅 AM | ✓ |
| C | All PM | 仅 PM | ✓ |
| D | Wed–Fri AM/PM | 仅 Wed/Thu/Fri PM | ✓ |
| E | All | 任意 | ✓ |

### 4. 偏好检查

- **同一天 AM+PM**：本方案无人同一天双班 ✓。
- **周五 PM 疲劳控制**：周五 PM = D（3 班）、E（4 班），二人本周累计均 ≤4 班 ✓。

## 四、结论

主方案满足全部硬约束（每班 ≥1 Bilingual、周二/周四 PM ≥1 Senior、可用性、Max Shifts），并满足两项运营偏好（无双班、周五 PM 低负荷人员）。备选方案（供员工请假时替换）：

**备选方案**（若 D 不可用）：周三 PM 无 Bilingual 可替换，需从外部借调或调整；可将周三 PM 改为 A（若 A 例外可排 Wed PM）并复核其 6 班上限。此备选会违反 A 的“不可用 Wed PM”，需人工批准并记录。
