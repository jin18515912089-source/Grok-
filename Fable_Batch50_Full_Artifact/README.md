# Fable Batch-50 Full-Artifact Submission

本目录为 OpenRouter / Claude Fable 5 Batch-50 白领工作能力基准测试的**全量交付包**（T01–T50）。

## 目录结构

- `ALL_ANSWERS_T01-T50.md`：50 题全部答案（含边界标记、ASSUMPTIONS、UNMET、SELF_CHECK），可直接交给评分模型。
- `submission_manifest.csv`：50 行提交清单。
- `Txx/answer.md`：单题答案；`Txx/*.docx|*.pptx|*.xlsx|*.md`：该题的真实交付文件。

## 交付文件清单

| Task | Files |
|---|---|
| T01 | T01_管理周报.docx |
| T02 | T02_CRM采购决策备忘录.docx |
| T03 | T03_董事会月度经营简报.docx |
| T04 | T04_会议纪要与行动项.docx |
| T05 | T05_生产事故初步复盘.docx |
| T06 | T06_区域销售经营汇报.pptx、T06_管理层摘要.md |
| T07 | T07_增长战略选项董事会.pptx |
| T08 | T08_季度业绩投资者更新.pptx |
| T09 | T09_反钓鱼培训课件.pptx |
| T10 | T10_新产品销售赋能.pptx |
| T11 | T11_销售数据清洗与汇总.xlsx |
| T12 | T12_部门预算差异分析.xlsx |
| T13 | T13_订阅用户Cohort留存分析.xlsx |
| T14 | T14_库存补货计算器.xlsx |
| T15 | T15_六个月现金流情景模型.xlsx |
| T16 | T16_B2B整合营销文案包.docx |
| T17 | T17_品牌语气本地化包.docx |
| T18 | T18_客户投诉分级与回复.docx |
| T19 | T19_企业客户RFP响应.docx |
| T20 | T20_线索评分与外联.docx |
| T21 | T21_双向NDA条款审阅.docx |
| T22 | T22_MSA_Working_Draft.docx |
| T23 | T23_员工纪律调查通知.docx |
| T24 | T24_隐私事件通知判断.docx |
| T25 | T25_营销主张合规审查.docx |
| T26 | T26_竞争格局与供应商短名单.docx |
| T27 | T27_市场规模TAM_SAM_SOM.docx |
| T28 | T28_云供应商风险尽调.docx |
| T29 | T29_混合办公政策备忘录.docx |
| T30 | T30_存储公司季度业绩研究.docx |
| T31 | T31_三张财务报表勾稽检查.docx |
| T32 | T32_多情景股票投资备忘录.docx |
| T33 | T33_组合敞口与压力测试.docx |
| T34 | T34_IFRS16租赁摊销表.xlsx |
| T35 | T35_采购三年TCO_NPV.xlsx、T35_采购建议.docx |
| T36 | T36_跨职能上线计划.xlsx |
| T37 | T37_客户上线SOP.docx |
| T38 | T38_客服排班优化.docx |
| T39 | T39_供应商加权评分.docx |
| T40 | T40_制造缺陷根因初析.docx |
| T41 | T41_岗位说明书与评分卡.docx |
| T42 | T42_面试证据综合与录用建议.docx |
| T43 | T43_绩效反馈与30天改进计划.docx |
| T44 | T44_办公室搬迁公告与FAQ.docx |
| T45 | T45_培训需求分析与30天计划.docx |
| T46 | T46_高管收件箱分流.docx |
| T47 | T47_跨城市高管差旅议程.docx |
| T48 | T48_客户活动方案.docx |
| T49 | T49_双语合同差异核对.docx |
| T50 | T50_董事会PPT.pptx、T50_董事会备忘录.docx、T50_董事会指标附录.xlsx |

## 生成与验证说明

- DOCX 由 python-docx 生成；PPTX 由 python-pptx 生成（原生可编辑图表）；XLSX 由 openpyxl 生成（真实 Excel 公式、条件格式、图表）。
- 全部 XLSX 已用 LibreOffice 无头重算并核对关键单元格结果与手工计算一致，未发现 #REF!/#VALUE!/#DIV/0! 等错误。
- 全部 PPTX 页数已按题目要求核对（T06=6、T07=7、T08=5、T09=8、T10=7、T50=6）；DOCX/PPTX 已渲染为 PDF 检查可打开性与版式。
- 颜色约定（XLSX）：黄色 = 输入，绿色 = 公式。
