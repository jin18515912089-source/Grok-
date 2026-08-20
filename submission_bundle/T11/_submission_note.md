# T11 提交说明

## 交付文件
- T11_销售数据清洗与汇总.xlsx（Raw、Clean、Summary、Exceptions 四表）

## 关键假设
- 净收入规则：Completed=Units×Price；Refund 用负 Units×Price；Cancelled/Pending=0。
- 完全重复行（1007 出现两次）保留一条；相似订单（如 1010 与 1001）不删除。

## 未满足项
- 无缺失；含 4 个工作表与 1 张区域净收入图。

## 自检结果
- Rep 标准化：A. Chen/A Chen/Alex Chen → Alex Chen；Nina K. → Nina Ko。
- Region 首字母大写，缺失写 Unknown（订单 1009）并进 Exceptions。
- Summary 按 Region 汇总：North 3400、South 1960、East 1060、West 1360、Unknown 1400，合计 9180；订单数 14。
