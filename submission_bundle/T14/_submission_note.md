# T14 提交说明

## 交付文件
- T14_库存补货计算器.xlsx（补货计算、异常检查、采购清单）

## 关键假设
- Reorder Point = Avg Daily Demand × Lead Time + Safety Stock。
- 触发条件：On Hand + Open PO ≤ Reorder Point。
- 目标库存 = 需求 × (Lead Time + 14天) + Safety Stock。
- 建议下单量 = max(0, 目标库存 − On Hand − Open PO)，向上取整到 Case Pack 整数倍。

## 未满足项
- 无缺失。

## 自检结果
- 触发 SKU：A（下单 120）、C（130）、F（125）、G（108）。
- 未触发：B、D、E、H（On Hand+PO > ROP）。
- 计算：A ROP=120 触发；F 目标 152、下单 122→取整 125；G 目标 192、下单 102→取整 108。
