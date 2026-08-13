===== BEGIN T14 =====
STATUS: COMPLETE
ORIGINAL_DELIVERABLE: 可编辑XLSX（补货计算、异常检查、采购清单）
PROXY_DELIVERABLE: 本题已提供原格式办公文件（见本文件夹）

## ANSWER


## DELIVERABLE_FILES
本题正式交付物已放在本文件夹内，可用 PowerPoint / Excel / Word 直接打开：
  - T14_Reorder_Calculator.xlsx
下面的正文是同一套内容的说明、公式与自检，与二进制文件一致。


### FILE: Reorder_Calculator.xlsx

**公式（行2 起）**  
B OnHand, C AvgDailyDemand, D LeadTime, E Safety, F OpenPO, G CasePack  
H ROP：`=C2*D2+E2`  
I Available：`=B2+F2`  
J Trigger：`=IF(I2<=H2,"YES","NO")`  
K Target：`=C2*(D2+14)+E2`  
L RawQty：`=MAX(0,K2-B2-F2)`  
M OrderQty：`=IF(J2="NO",0,IF(L2=0,0,CEILING(L2,G2)))`  
（仅触发才下单；向上取整到 Case Pack 整数倍）

| SKU | OH | ADD | LT | SS | OpenPO | Pack | ROP | Avail | Trig | Target | Raw | Order |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---|---:|---:|---:|
| A | 120 | 8 | 10 | 40 | 0 | 20 | 120 | 120 | YES | 232 | 112 | **120** |
| B | 500 | 12 | 7 | 60 | 100 | 50 | 144 | 600 | NO | 312 | 0 | 0 |
| C | 40 | 5 | 14 | 25 | 0 | 10 | 95 | 40 | YES | 165 | 125 | **130** |
| D | 200 | 20 | 5 | 50 | 0 | 25 | 150 | 200 | NO | 430 | 230 | 0 |
| E | 75 | 4 | 21 | 30 | 50 | 10 | 114 | 125 | NO | 170 | 45 | 0 |
| F | 0 | 3 | 30 | 20 | 30 | 25 | 110 | 30 | YES | 152 | 122 | **125** |
| G | 90 | 6 | 12 | 36 | 0 | 12 | 108 | 90 | YES | 192 | 102 | **108** |
| H | 1000 | 25 | 14 | 100 | 500 | 100 | 450 | 1500 | NO | 800 | 0 | 0 |

**异常检查**
- F：On Hand=0，Avail 30≤ROP 110，且 Open PO 30 仍远低于 ROP → 缺货风险最高。
- A：Avail 恰好等于 ROP（120=120），按“≤”触发。
- D/E：Target>Avail 但未触及 ROP，按规则不下单（避免把 min-max 误用成每日补到目标）。
- B/H：库存远高于 ROP，Open PO 已大。
- F 建议订 125，Pack=25，122→125；订到后 Avail=30+125=155≥Target 152。
- C：125→130；G：102→108。

**采购清单（仅 Trigger=YES）**

| SKU | OrderQty | 备注 |
|---|---:|---|
| A | 120 | 6 cases |
| C | 130 | 13 cases |
| F | 125 | 5 cases；优先 |
| G | 108 | 9 cases |
| 合计 SKU 数 | 4 | 件数 483 |

## ASSUMPTIONS_AND_UNKNOWNS
需求在交期+14 天内视为恒定日均。未给在途以外的供应约束。CEILING 对已整除数保持原值。

## UNMET_ARTIFACT_REQUIREMENTS
无法生成可打开 XLSX。

## SELF_CHECK
- ROP：A 80+40=120；B 84+60=144；C 70+25=95；D 100+50=150；E 84+30=114；F 90+20=110；G 72+36=108；H 350+100=450。
- A 触发因 ≤；D 200>150 不触发。
- 取整：112→120，125→130，122→125，102→108。

===== END T14 =====
