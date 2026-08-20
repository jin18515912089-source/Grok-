#!/usr/bin/env python3
"""Build required PPTX deliverables for T06-T10 and T50."""
import sys
from pathlib import Path

from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).parent))
from pptx_util import (
    NAVY, TEAL, ORANGE, RED, GREEN, GOLD, WHITE, SLATE, LIGHT, MUTED,
    new_prs, add_blank, bar, textbox, bullets, header, footer, add_table, set_run_font,
)

OUT = Path("/workspace/office-deliverables")
OUT.mkdir(parents=True, exist_ok=True)


def add_chart(slide, chart_type, cats, series, l, t, w, h, title=None):
    data = CategoryChartData()
    data.categories = cats
    for name, vals in series:
        data.add_series(name, vals)
    chart = slide.shapes.add_chart(chart_type, l, t, w, h, data).chart
    chart.has_legend = len(series) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.paragraphs[0].text = title
    plot = chart.plots[0]
    plot.has_data_labels = False
    try:
        chart.value_axis.has_major_gridlines = True
        chart.value_axis.scaling.min = 0
    except Exception:
        pass
    return chart


def t06():
    prs = new_prs()
    total = 6
    label = "Q2 区域销售回顾 ｜ 可编辑 PPTX"

    s = add_blank(prs)
    bar(s, Inches(0), Inches(0), Inches(0.22), Inches(7.5), TEAL)
    bar(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), NAVY)
    bar(s, Inches(0), Inches(0), Inches(0.22), Inches(7.5), GOLD)
    textbox(s, Inches(0.7), Inches(1.6), Inches(12), Inches(0.4), "QUARTERLY BUSINESS REVIEW", 14, True, GOLD)
    textbox(s, Inches(0.7), Inches(2.05), Inches(12), Inches(1.2), "Q2 区域销售回顾：增长来自北美伙伴，质量卡在毛利与亚太", 28, True, WHITE)
    bullets(s, Inches(0.7), Inches(3.5), Inches(11.5), Inches(2.6), [
        "Q2 总收入 31.0 百万元，环比 +10.7%（Q1 28.0）",
        "增量 80% 来自北美 +2.4；亚太 −10.0%，日本换渠道流失 1.1",
        "Q2 加权毛利率测算 52.4%，仅高于 52% 底线 0.4 个百分点",
        "下季硬门槛：收入 ≥33.48 且整体毛利率 ≥52%",
    ], 18, WHITE)
    footer(s, 1, total, label)

    s = add_blank(prs)
    header(s, "区域趋势：北美贡献了增长的 4/5", "金额：百万元人民币；柱状图坐标从 0 起，不截断")
    add_table(s, Inches(0.4), Inches(1.2), Inches(6.3), Inches(2.4), [
        ["区域", "Q1", "Q2", "增量", "环比"],
        ["北美", "12.0", "14.4", "+2.4", "+20.0%"],
        ["欧洲", "8.0", "8.4", "+0.4", "+5.0%"],
        ["亚太", "6.0", "5.4", "−0.6", "−10.0%"],
        ["拉美", "2.0", "2.8", "+0.8", "+40.0%"],
        ["合计", "28.0", "31.0", "+3.0", "+10.7%"],
    ], col_w=[Inches(1.5), Inches(1.1), Inches(1.1), Inches(1.2), Inches(1.4)])
    add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, ["北美", "欧洲", "亚太", "拉美"],
              [("Q1 收入", (12.0, 8.0, 6.0, 2.0)), ("Q2 收入", (14.4, 8.4, 5.4, 2.8))],
              Inches(6.9), Inches(1.15), Inches(6.0), Inches(3.5), "区域收入（从 0 起）")
    add_chart(s, XL_CHART_TYPE.BAR_CLUSTERED, ["北美", "欧洲", "亚太", "拉美"],
              [("Q2−Q1 增量", (2.4, 0.4, -0.6, 0.8))],
              Inches(0.4), Inches(3.85), Inches(6.3), Inches(3.0), "增量拆解（+3.0）")
    textbox(s, Inches(6.9), Inches(4.75), Inches(6.0), Inches(2.1),
            "事实：北美增长主要来自两个渠道伙伴；拉美从较小基数增长；亚太在日本渠道更换期间流失 1.1。若无该流失，亚太应为 6.5，将高于 Q1。不要把 1.1 从 5.4 里再扣一次。",
            13, False, SLATE)
    footer(s, 2, total, label)

    s = add_blank(prs)
    header(s, "产品组合与毛利：整体毛利只剩 0.4 个百分点缓冲", "Q2 毛利测算 = Σ(收入×毛利率)；Q1 各产品毛利率未知，故不画 Q1 毛利对比")
    add_table(s, Inches(0.4), Inches(1.2), Inches(7.4), Inches(2.3), [
        ["产品", "Q1 收入", "Q2 收入", "环比", "Q2 毛利率", "Q2 毛利测算"],
        ["Core", "15.0", "18.0", "+20.0%", "62%", "11.16"],
        ["Growth", "7.0", "8.0", "+14.3%", "48%", "3.84"],
        ["Legacy", "6.0", "5.0", "−16.7%", "25%", "1.25"],
        ["合计", "28.0", "31.0", "+10.7%", "52.4%", "16.25"],
    ])
    add_chart(s, XL_CHART_TYPE.COLUMN_STACKED, ["Q1", "Q2"],
              [("Core", (15, 18)), ("Growth", (7, 8)), ("Legacy", (6, 5))],
              Inches(7.9), Inches(1.15), Inches(4.9), Inches(3.4), "产品收入结构")
    bullets(s, Inches(0.4), Inches(3.8), Inches(7.3), Inches(3.0), [
        "公式：(18×0.62 + 8×0.48 + 5×0.25) / 31 = 16.25 / 31 = 52.419% ≈ 52.4%",
        "Q2 结构：Core 58.1% / Growth 25.8% / Legacy 16.1%",
        "Legacy 计划在未来两个季度停止新增销售（计划，非已实现）",
        "区域合计与产品合计 Q1=28、Q2=31，勾稽通过",
    ], 15)
    footer(s, 3, total, label)

    s = add_blank(prs)
    header(s, "问题诊断：8% 增长与 52% 毛利会互相打架", "不得把渠道伙伴增长写成可无限复制的因果")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.4), [
        "增长过度集中北美两伙伴——集中度风险；题内无伙伴合同细节。",
        "亚太日本渠道更换是已知流失 1.1，是否已稳定未知。",
        "Legacy 停新增：收入可能继续掉；若用低毛利折扣填坑会再压 GM。",
        "测算：若 Q3=33.48 且 Legacy 维持 5.0、毛利率结构同 Q2，则非 Legacy 需 28.48；Q2 非 Legacy 为 26.0，还需 +2.48。",
        "欧洲仅 +5.0%，低于公司 8% 目标，列为第三优先级。",
        "拉美增速高但无分区域毛利，不能用低价换量。",
    ], 17)
    footer(s, 4, total, label)

    s = add_blank(prs)
    header(s, "下季优先级：亚太修复 > 毛利守线 > 北美伙伴扩量", "任何折扣申请先过毛利测算")
    add_table(s, Inches(0.4), Inches(1.25), Inches(12.5), Inches(4.6), [
        ["优先级", "动作", "现在", "30 天", "90 天"],
        ["1 亚太", "日本渠道替换后客户回流清单，先止住环比", "锁定流失客户名单", "一对一挽回", "区域止跌"],
        ["2 毛利", "停止 Legacy 新增；存量只续不扩", "冻结新售", "Core/Growth 配额对齐 33.48", "GM≥52%"],
        ["3 北美", "两伙伴打成可复制 playbook，设集中度上限", "复盘伙伴贡献", "上限待负责人定（题内无）", "降低单点依赖"],
        ["4 欧洲", "低于 8% 目标，第三优先级", "诊断缺口", "补管道", "追平公司增速"],
        ["5 拉美", "保持增速，不用低价换量", "毛利未知，先观察", "禁止无毛利测算折扣", "小基数扩张"],
    ])
    footer(s, 5, total, label)

    s = add_blank(prs)
    header(s, "附录：数据、公式与未知项", "图表坐标均从 0 起")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.4), [
        "数据源：I06-A 区域收入、I06-B Q2 产品组合、I06-C 经营背景。单位：百万元。",
        "测算毛利率公式：Σ(收入×毛利率)/Σ收入 = 16.25/31 = 52.419%。",
        "Q3 收入门槛：31.0 × 1.08 = 33.48。",
        "未知：分区域毛利、伙伴收入拆分、日本可恢复金额、Q1 产品毛利率。",
        "管理层摘要：Q2 收入 31.0 百万元，环比 +10.7%。增长 80% 来自北美（+2.4，渠道伙伴），拉美 +40% 基数小；亚太 −10.0%，日本换渠道流失 1.1 百万元。Q2 毛利率测算 52.4%，仅高于 52% 底线 0.4 个百分点；Legacy 毛利率 25% 且将停新增。下季要同时做到收入 ≥33.48 且毛利率 ≥52%，必须先修复亚太并压降 Legacy 混比。",
    ], 16)
    footer(s, 6, total, label)
    prs.save(OUT / "T06_Q2_Regional_Sales_Review.pptx")


def t07():
    prs = new_prs()
    total = 7
    label = "增长战略选项 ｜ 董事会"

    s = add_blank(prs)
    bar(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), NAVY)
    textbox(s, Inches(0.7), Inches(1.5), Inches(12), Inches(0.4), "BOARD DECISION REQUEST", 14, True, GOLD)
    textbox(s, Inches(0.7), Inches(2.0), Inches(12), Inches(1.4), "请批准 Beta（美国 B2B）为主方案，Alpha 为不超过 0.8 百万元的验证试点", 28, True, WHITE)
    bullets(s, Inches(0.7), Inches(3.7), Inches(12), Inches(2.8), [
        "主方案唯一：Beta 美国 B2B",
        "保留备选：Alpha 德国 D2C，试点预算 ≤0.8 百万元，不建运营团队",
        "不选 Gamma 作主方案：初始投资 12>8，必须另行融资",
        "约束：18 个月投资上限 8；希望 24 个月内盈亏平衡；团队无跨境消费者运营经验",
        "概率调整收入几乎打平，不是选型依据（不等于期望利润）",
    ], 18, WHITE)
    footer(s, 1, total, label)

    s = add_blank(prs)
    header(s, "三方案一览（百万元）", "硬约束：投资≤8、BE≤24 个月")
    add_table(s, Inches(0.35), Inches(1.2), Inches(12.6), Inches(5.5), [
        ["指标", "Alpha 德国 D2C", "Beta 美国 B2B", "Gamma 收购"],
        ["TAM", "900", "600", "450"],
        ["CAGR", "8%", "12%", "5%"],
        ["初始投资", "5", "3", "12"],
        ["盈亏平衡", "第 30 月", "第 18 月", "交割后即期"],
        ["第3年收入", "22", "18", "26"],
        ["成功概率", "65%", "80%", "55%"],
        ["监管 / 执行", "中 / 中", "低 / 低", "中 / 高"],
        ["客户渠道复用", "低", "高", "中"],
        ["18个月预算", "通过", "通过", "不通过（须融资）"],
        ["24个月BE", "不通过（30月）", "通过", "通过（即期）"],
    ])
    footer(s, 2, total, label)

    s = add_blank(prs)
    header(s, "定量比较：概率调整收入几乎相同", "概率调整收入 = 第3年收入 × 成功概率。这是比较工具，不等于期望利润，未扣投资/成本/资本成本，也不是 NPV。")
    add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, ["Alpha", "Beta", "Gamma"],
              [("第3年收入", (22, 18, 26)), ("概率调整收入", (14.3, 14.4, 14.3))],
              Inches(0.5), Inches(1.3), Inches(7.8), Inches(5.4))
    bullets(s, Inches(8.5), Inches(1.5), Inches(4.4), Inches(5.2), [
        "Alpha 22 × 65% = 14.3",
        "Beta 18 × 80% = 14.4",
        "Gamma 26 × 55% = 14.3",
        "三者几乎打平，所以用约束和风险分胜负。",
        "Beta 投资 3、BE 18 月，落在 24 月意愿内。",
        "Alpha BE 30 月超出。",
        "Gamma 即期 BE，但 12 的资金不在现有 8 的口袋里。",
    ], 15)
    footer(s, 3, total, label)

    s = add_blank(prs)
    header(s, "风险与能力匹配", "不要把 TAM 最大（Alpha 900）当成已赢市场")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "Beta：监管低、执行低、复用高——与“无跨境消费者经验”相容（B2B 可复用现有客户/渠道）。",
        "Alpha：D2C 消费者运营经验缺口大，执行中、复用低；即便试点成功，全面铺开仍缺能力。",
        "Gamma：执行高、须另融资；整合风险高。若要收购，需单独立融资讨论，本次主方案不纳入。",
        "成功概率为题内给定，非自估。试点 0.8 建议计入 8 以内以免超支（是否单列未知）。",
    ], 18)
    footer(s, 4, total, label)

    s = add_blank(prs)
    header(s, "主推荐逻辑：唯一同时满足三条硬约束", "否决 Alpha 作主方案是因为 BE 30 月与能力缺口，不是收入数字")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "唯一同时满足：投资 ≤8、BE ≤24 月、不依赖跨境消费者能力、不强制融资。",
        "第3年收入 Beta 低于 Gamma，但可执行性与成功概率最高（80%）。",
        "Gamma 因 12>8 不能主推。",
        "Alpha 因 BE 30>24 且团队无跨境消费者运营经验，不能主推。",
    ], 18)
    footer(s, 5, total, label)

    s = add_blank(prs)
    header(s, "实施路线图（Beta 主方案 + Alpha 试点）", "Gamma 不启动，除非董事会另批融资路径")
    add_table(s, Inches(0.4), Inches(1.25), Inches(12.5), Inches(4.8), [
        ["阶段", "动作", "资金/门槛"],
        ["0–3 月", "美国 B2B 管道验证、定价、合规清单", "投资节奏 ≤3 的首期"],
        ["3–12 月", "复用现有客户/渠道打样", "里程碑=签约与实施产能"],
        ["12–18 月", "冲刺盈亏平衡", "目标第 18 月"],
        ["18–24 月", "若未 BE，启动收缩而非加码 Alpha/Gamma", "不追加超预算投资"],
        ["并行试点", "Alpha 小规模 D2C 假设检验（获客成本、履约）", "≤0.8；18 个月复盘是否升格"],
        ["Gamma", "不启动", "须另行融资讨论"],
    ])
    footer(s, 6, total, label)

    s = add_blank(prs)
    header(s, "附录、待确认与表决票", "未知：各方案成本结构、税率、协同、收购倍数拆分、融资条件")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "公式：概率调整收入 = 第3年收入 × 成功概率。",
        "无折现、无成本，故不做 NPV。",
        "请表决：1) 批准 Beta 主方案；2) 批准 Alpha ≤0.8 试点；3) 搁置 Gamma。",
    ], 18)
    footer(s, 7, total, label)
    prs.save(OUT / "T07_Growth_Options_Board.pptx")


def t08():
    prs = new_prs()
    total = 5
    label = "Lumina Memory Q4 投资者更新 ｜ 测算数字已标注"

    s = add_blank(prs)
    bar(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), NAVY)
    textbox(s, Inches(0.7), Inches(1.3), Inches(12), Inches(0.35), "INVESTOR UPDATE ｜ 虚构公司 Lumina Memory", 14, True, GOLD)
    textbox(s, Inches(0.7), Inches(1.75), Inches(12), Inches(1.3), "表面超预期，质量不够支撑估值上修", 30, True, WHITE)
    bullets(s, Inches(0.7), Inches(3.3), Inches(12), Inches(3.5), [
        "收入 12.4 vs 12.0（测算 beat +3.3%）；毛利率 56.0% vs 55.0%（+1.0 个百分点）；营业利润 4.1 vs 3.8（测算 +7.9%）",
        "GAAP 净利润 5.6 vs 3.5，含处置子公司税后收益 2.2；归一化净利润测算 5.6−2.2=3.4，低于一致预期 3.5（测算 miss 0.1）",
        "出货同比 +4% vs 公司指引 +8% 至 +10%；ASP 同比 +10% vs 市场约 +7%",
        "业绩后股价 −7%。评级：持有；目标价测算 182 美元 vs 现价 178（测算上行 +2.2%）",
    ], 16, WHITE)
    footer(s, 1, total, label)

    s = add_blank(prs)
    header(s, "业绩质量：一次性收益必须从归一化盈利中剔除", "管理层称出货偏弱来自客户验收延后——题内注明尚未独立验证")
    add_table(s, Inches(0.4), Inches(1.2), Inches(6.6), Inches(3.3), [
        ["项目", "实际", "一致预期", "差额"],
        ["收入（十亿美元）", "12.4", "12.0", "+0.4 / +3.3%"],
        ["毛利率", "56.0%", "55.0%", "+1.0pp"],
        ["营业利润", "4.1", "3.8", "+0.3 / +7.9%"],
        ["GAAP 净利润", "5.6", "3.5", "含一次性"],
        ["归一化净利润（测算）", "3.4", "3.5", "miss 0.1"],
    ])
    add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED,
              ["收入", "营业利润", "GAAP NI", "归一化 NI"],
              [("实际", (12.4, 4.1, 5.6, 3.4)), ("一致预期", (12.0, 3.8, 3.5, 3.5))],
              Inches(7.2), Inches(1.15), Inches(5.7), Inches(4.4))
    textbox(s, Inches(0.4), Inches(4.7), Inches(6.6), Inches(2.0),
            "归一化 EPS 测算 3.4 美元（摊薄 1.0 十亿股）。GAAP EPS 5.6 不可用于盈利质量。出货 +4% 低于指引 4–6 个百分点。电话会叙事不能当订单恢复的证据。",
            14, False, SLATE)
    footer(s, 2, total, label)

    s = add_blank(prs)
    header(s, "指引：中点高于一致预期，但量的可信度被本季出货 miss 削弱")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "下季收入指引 13.2±0.3，即 12.9–13.5；一致预期 12.8。中点 13.2 相对一致预期测算 +0.4（+3.1%）。",
        "毛利率指引 58%±1 个百分点，即 57%–59%，中点高于 Q4 实际 56.0%。",
        "解读：指引中点高于一致预期，但量的指引可信度被本季出货 miss 削弱。",
    ], 18)
    add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, ["下季收入"],
              [("指引下限", (12.9,)), ("指引中点", (13.2,)), ("指引上限", (13.5,)), ("一致预期", (12.8,))],
              Inches(2.2), Inches(3.5), Inches(8.8), Inches(3.3))
    footer(s, 3, total, label)

    s = add_blank(prs)
    header(s, "估值框架（仅用给定输入）", "不得引用外部倍数")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "CY2027 营业利润 22 十亿美元（给定）。",
        "目标 EV/营业利润 8.0× → EV 测算 176。",
        "预计净现金 6 → 股权价值测算 182。",
        "摊薄 1.0 十亿股 → 目标价测算 182 美元。",
        "现价 178；目标回报测算 (182−178)/178 = +2.2%。",
    ], 18)
    footer(s, 4, total, label)

    s = add_blank(prs)
    header(s, "评级：持有 ｜ 风险", "规则：>15% 买入；−10% 至 +15% 持有；<−10% 卖出")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "+2.2% → 持有。",
        "为何下跌：市场给价的是量与归一化盈利，不是含 2.2 的 GAAP。",
        "风险：验收延后若为需求问题而非时点问题；出货持续低于指引；毛利率指引 58% 依赖未知组合。",
        "催化剂未知（题内无后续独立验证时点）。处置收益已是税后，直接从 GAAP 净利润扣除。",
    ], 18)
    footer(s, 5, total, label)
    prs.save(OUT / "T08_Lumina_Memory_Q4_Investor_Update.pptx")


def t09():
    prs = new_prs()
    total = 8
    label = "反钓鱼培训 15 分钟 ｜ 非技术员工"

    s = add_blank(prs)
    bar(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), NAVY)
    textbox(s, Inches(0.7), Inches(2.0), Inches(12), Inches(0.4), "15 分钟信息安全培训", 14, True, GOLD)
    textbox(s, Inches(0.7), Inches(2.5), Inches(12), Inches(1.2), "收到可疑邮件时，你只做三件事", 32, True, WHITE)
    bullets(s, Inches(0.7), Inches(4.0), Inches(12), Inches(2.5), [
        "不点链接、不打开附件、不回复",
        "用邮件客户端“报告钓鱼”按钮",
        "若已输入密码，立即致电 IT 服务台 7000 并重置密码",
    ], 20, WHITE)
    footer(s, 1, total, label)

    s = add_blank(prs)
    header(s, "公司政策原文，请照做", "IT 永远不会通过邮件索取密码或 MFA 验证码")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "可疑邮件：不要点击链接、不要打开附件、不要回复。",
        "举报：只使用“报告钓鱼”按钮。没有其他举报渠道。",
        "已输入密码：立即致电 7000 并重置密码。",
        "付款请求：电话回拨已知号码向请求人核实（不要用邮件里的电话）。",
        "禁止：自行转发可疑邮件；点击链接“验证一下”。",
    ], 18)
    footer(s, 2, total, label)

    s = add_blank(prs)
    header(s, "看起来像真的，仍然危险（示例2）", "正常语气 + 正常发件人，也可以是真威胁")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "长期合作供应商从正常邮箱发来新银行账户通知，要求今天付款，附件 PDF，语气正常。",
        "为何危险：账户变更 + 紧迫付款是商务邮件欺诈典型结构；正常发件箱也可能被盗用。",
        "正确动作：不打开附件、不回复；报告钓鱼；用已知号码回拨核实。",
        "高亮信号：新账户、今天、PDF 附件。",
    ], 18)
    footer(s, 3, total, label)

    s = add_blank(prs)
    header(s, "其他三例速判")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "示例1：it-support@novacIoud.com（大写 i），30 分钟停用，链接文字像公司域名、悬停为陌生短链 → 钓鱼。不点击，报告钓鱼。",
        "示例3：HR 系统通知看工资单，域名/链接/登录页与书签一致，未索取验证码 → 当前信息下像正规通知。仍建议走书签而非邮件链接。",
        "示例4：“CEO”出差要买礼品卡并把卡号回邮件、强调保密 → 典型权威+保密欺诈。不回复、报告钓鱼。",
        "示例1 的 novacIoud 是字形欺骗（I 与 l）。",
    ], 17)
    footer(s, 4, total, label)

    s = add_blank(prs)
    header(s, "互动题 1（先不要看答案页）")
    textbox(s, Inches(0.5), Inches(1.4), Inches(12.2), Inches(1.2),
            "你收到 IT 邮件要你 30 分钟内点链接确认密码，否则停用。你怎么做？", 20, True, NAVY)
    bullets(s, Inches(0.7), Inches(2.8), Inches(12), Inches(3.8), [
        "A. 点链接看看是不是真的",
        "B. 回复邮件问是不是 IT 发的",
        "C. 不点击、不回复，点“报告钓鱼”；若已输入密码，打 7000",
        "D. 把邮件转发给同事帮忙看",
    ], 20)
    footer(s, 5, total, label)

    s = add_blank(prs)
    header(s, "互动题 2（先不要看答案页）")
    textbox(s, Inches(0.5), Inches(1.4), Inches(12.2), Inches(1.2),
            "合作多年的供应商来信改银行账户，要今天付款，附件是 PDF。你怎么做？", 20, True, NAVY)
    bullets(s, Inches(0.7), Inches(2.8), Inches(12), Inches(3.8), [
        "A. 打开 PDF 核对账号",
        "B. 按邮件里的新账号付款以免断供",
        "C. 不打开附件；报告钓鱼；用已知号码回拨对方核实",
        "D. 回复“请再发一次公章文件”",
    ], 20)
    footer(s, 6, total, label)

    s = add_blank(prs)
    header(s, "答案页")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "题1 答案：C。A 违反“不要点击”；B 违反“不要回复”；D 不是政策渠道，且可能扩散风险。政策渠道只有“报告钓鱼”和 7000。",
        "题2 答案：C。付款类必须回拨已知号码；A/B/D 都会进入攻击者流程。",
        "补充：IT 不会用邮件要密码或 MFA。",
    ], 18)
    footer(s, 7, total, label)

    s = add_blank(prs)
    header(s, "90 秒行动卡（带走）")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "停：不点、不打开、不回。",
        "报：报告钓鱼。",
        "呼：已输密码 → 7000 + 重置；付款 → 回拨已知号码。",
        "没有其他举报渠道。",
    ], 22)
    footer(s, 8, total, label)
    prs.save(OUT / "T09_Anti_Phishing_Training.pptx")


def t10():
    prs = new_prs()
    total = 7
    label = "FocusFlow Enterprise 销售赋能 ｜ 可编辑"

    s = add_blank(prs)
    bar(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), NAVY)
    textbox(s, Inches(0.7), Inches(1.8), Inches(12), Inches(0.35), "PRODUCT MARKETING ｜ SALES ENABLEMENT", 14, True, GOLD)
    textbox(s, Inches(0.7), Inches(2.3), Inches(12), Inches(1.3), "FocusFlow Enterprise：让任务、会议行动和项目阻塞可见", 28, True, WHITE)
    bullets(s, Inches(0.7), Inches(4.0), Inches(12), Inches(2.6), [
        "目标客户：500 人以上专业服务和软件企业",
        "买方：COO、IT、部门负责人",
        "痛点：跨团队任务不可见、会议后行动丢失、管理者无法判断项目阻塞",
        "禁句：保证提升生产率、完全合规；无客户案例名称可引用",
    ], 18, WHITE)
    footer(s, 1, total, label)

    s = add_blank(prs)
    header(s, "功能与边界（如实）", "尚未取得 HIPAA 认证——医疗/PHI 场景不得承诺")
    add_table(s, Inches(0.5), Inches(1.4), Inches(12.3), Inches(4.5), [
        ["提供", "不提供 / 未认证"],
        ["会议纪要转行动项", "HIPAA 认证"],
        ["跨项目仪表板", "保证提升生产率"],
        ["SSO", "完全合规"],
        ["审计日志", "未提供的客户 logo / 案例"],
        ["数据驻留可选新加坡或法兰克福", "试用含生产数据迁移"],
    ])
    footer(s, 2, total, label)

    s = add_blank(prs)
    header(s, "已批准证据（必须带限定）")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "12 家试点客户中位数：逾期行动项下降 22%；统计窗口 8 周；样本未经独立审计。",
        "正确讲法：在 12 家试点、8 周窗口内，逾期行动项中位数下降 22%（未经独立审计）。",
        "错误讲法：保证生产率、扩大到所有客户、暗示审计通过。",
    ], 18)
    footer(s, 3, total, label)

    s = add_blank(prs)
    header(s, "发现问题清单（Discovery）")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "会议后行动目前如何追踪？丢失率如何衡量？",
        "管理者能否在一个仪表板看到跨项目阻塞？",
        "现有工具是否覆盖会议→行动闭环，还是只有任务列表？",
        "SSO、审计日志、数据驻留地是否为采购门槛？",
        "是否涉及 HIPAA/PHI？（若是，当前不能作为合规方案出售。）",
        "席位数是否 ≥500（标准版）或 ≥800（Plus）？",
        "是否接受 30 天试用最多 100 席、不含生产数据迁移？",
    ], 17)
    footer(s, 4, total, label)

    s = add_blank(prs)
    header(s, "异议处理")
    add_table(s, Inches(0.4), Inches(1.2), Inches(12.5), Inches(5.2), [
        ["异议", "回应边界"],
        ["已有项目管理工具", "承认不替代所有 PM 工具；聚焦会议行动丢失与跨项目阻塞；用试点中位数 22% 并加限定。"],
        ["担心数据迁移", "30 天试用不含生产数据迁移；驻留地可选新加坡或法兰克福；迁移范围需单独立项。"],
        ["员工不愿学新工具", "不保证采用率；建议先用试用 100 席验证工作流。"],
        ["价格高", "回到最低席位与年预付结构；折扣不得超过规则（见报价页）。"],
    ])
    footer(s, 5, total, label)

    s = add_blank(prs)
    header(s, "报价与商务规则（必须准确）", "销售可自行批准最高 8%；8%–15% 需销售副总裁；超过 15% 不允许")
    add_table(s, Inches(0.4), Inches(1.25), Inches(12.5), Inches(4.8), [
        ["版本", "单价", "最低席位", "含什么", "测算年费下限"],
        ["标准版", "18 美元/用户/月", "500，按年预付", "标准功能", "18×500×12 = 108,000 美元"],
        ["Enterprise Plus", "25 美元/用户/月", "800", "高级审计 + 专属成功经理", "25×800×12 = 240,000 美元"],
        ["试用", "—", "最多 100 席 / 30 天", "不含生产数据迁移", "—"],
    ])
    footer(s, 6, total, label)

    s = add_blank(prs)
    header(s, "赢单检查与禁句")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "禁句：保证提升生产率、完全合规、HIPAA 已认证、未提供的客户 logo/案例。",
        "出单前：买方角色、驻留地、是否 PHI、席位门槛、折扣级、试用范围均已记录。",
        "下一步 CTA：合格账户安排试用（≤100 席）或 Plus 评估。",
    ], 18)
    footer(s, 7, total, label)
    prs.save(OUT / "T10_FocusFlow_Sales_Enablement.pptx")


def t50_ppt():
    prs = new_prs()
    total = 6
    label = "董事会材料 ｜ 与备忘录、XLSX 数字一致"

    s = add_blank(prs)
    bar(s, Inches(0), Inches(0), Inches(13.333), Inches(7.5), NAVY)
    textbox(s, Inches(0.7), Inches(1.4), Inches(12), Inches(0.35), "CEO OFFICE ｜ BOARD PACK", 14, True, GOLD)
    textbox(s, Inches(0.7), Inches(1.9), Inches(12), Inches(1.2), "ARR 在涨，质量在变差：请批三项决定", 28, True, WHITE)
    bullets(s, Inches(0.7), Inches(3.4), Inches(12), Inches(3.4), [
        "ARR 120（测算 QoQ +9.1%）但 NRR 104%、流失 2.8%、跑道测算 9.3 个月",
        "D1 批准 2 名上线经理试点（年成本 0.36m），不把 10 周→7 周当已验证承诺",
        "D2 批准高级 SLA 工程 0.25m，最早 Q4；不向客户承诺 99.95% 直至产品与合同完成",
        "D3 否决全体客户立即提价 8%",
        "禁止对外承诺 AI 8/15 或 9/15；销售材料 8/15 产品未批准；离线准确率 88%，门槛 92%",
    ], 16, WHITE)
    footer(s, 1, total, label)

    s = add_blank(prs)
    header(s, "经营质量：不是只报 ARR", "事实；金额百万美元")
    add_table(s, Inches(0.4), Inches(1.2), Inches(7.2), Inches(4.0), [
        ["指标", "Q2", "Q1", "Q2 LY"],
        ["ARR ($m)", "120", "110", "95"],
        ["NRR", "104%", "110%", "112%"],
        ["Gross Margin", "68%", "72%", "70%"],
        ["Logo Churn", "2.8%", "1.9%", "1.7%"],
        ["New Logos", "40", "55", "48"],
        ["Cash ($m)", "28", "31", "36"],
        ["Monthly Burn ($m)", "3.0", "2.5", "2.2"],
    ])
    add_chart(s, XL_CHART_TYPE.LINE_MARKERS, ["Q2 LY", "Q1", "Q2"],
              [("ARR", (95, 110, 120)), ("Cash", (36, 31, 28))],
              Inches(7.8), Inches(1.2), Inches(5.1), Inches(4.5), "ARR 上升 vs 现金下降")
    textbox(s, Inches(0.4), Inches(5.4), Inches(12.5), Inches(1.4),
            "测算：ARR QoQ +9.1%、YoY +26.3%；NRR −6pp；毛利率 −4pp；流失 +0.9pp；新客 logo −27.3%；现金 −9.7%；跑道 Q2 28/3.0=9.3 月 vs Q1 12.4 月。",
            14, False, SLATE)
    footer(s, 2, total, label)

    s = add_blank(prs)
    header(s, "客户证据", "访谈 ≠ 普查；10 个流失访谈")
    add_chart(s, XL_CHART_TYPE.BAR_CLUSTERED, ["上线慢", "AI 摘要准确性", "价格"],
              [("提及次数", (6, 3, 1))],
              Inches(0.5), Inches(1.3), Inches(7.5), Inches(5.2))
    bullets(s, Inches(8.2), Inches(1.5), Inches(4.6), Inches(5.0), [
        "五家企业要 99.95% SLA，当前标准 99.9%。",
        "其中两家愿为高级 SLA 付费。",
        "销售建议全体提价 8%，未提供流失弹性。",
        "理论年化毛增量 0.08×120=9.6m，未考虑流失、折扣、时点。",
    ], 15)
    footer(s, 3, total, label)

    s = add_blank(prs)
    header(s, "产品与法务约束", "事实、管理层判断、建议分开；不得做未授权承诺")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "AI 模块：销售材料曾写 8 月 15 日，产品未批准；内部目标 9 月 15 日。",
        "离线测试准确率 88%，上线门槛 92% → 不能对外承诺日期。",
        "法务：自动续约通知仅提前 30 天，两地区内部政策要求至少 60 天，须在下轮续约前改流程。",
        "SLA 现状 99.9%。高级 SLA 最早 Q4，定价未定。",
    ], 18)
    footer(s, 4, total, label)

    s = add_blank(prs)
    header(s, "决策情景（与 XLSX Decision_Scenarios 一致）")
    add_table(s, Inches(0.35), Inches(1.2), Inches(12.6), Inches(5.4), [
        ["方案", "成本", "声称收益", "证据", "建议"],
        ["D1 上线经理×2", "0.36m/年", "10 周→7 周", "过去项目估计，未经试点", "批为试点"],
        ["D2 高级 SLA", "0.25m 一次性，最早 Q4", "产品化 99.95%", "定价未知；两家愿付费", "批工程，不批对外承诺"],
        ["D3 全员+8%", "0 直接成本", "理论 +9.6m ARR", "无弹性；与流失/NRR 冲突", "否决"],
    ])
    footer(s, 5, total, label)

    s = add_blank(prs)
    header(s, "现金、未知项与附件索引", "三件文件数字必须一致")
    bullets(s, Inches(0.5), Inches(1.3), Inches(12.2), Inches(5.5), [
        "跑道测算 28/3.0=9.3 个月。D1+D2 合计 0.61m，不显著改变跑道数量级。",
        "附录 XLSX：KPI、Calculations、Decision_Scenarios。公式可追踪。",
        "未知：提价弹性、AI 达 92% 的日期、SLA 定价。",
        "请表决 D1 / D2（附带不得对客户承诺日期与 SLA）/ D3 否决立即全量提价。另授权法务把续约通知改到 ≥60 天。",
    ], 18)
    footer(s, 6, total, label)
    prs.save(OUT / "T50_Board_Pack.pptx")


if __name__ == "__main__":
    t06(); print("T06")
    t07(); print("T07")
    t08(); print("T08")
    t09(); print("T09")
    t10(); print("T10")
    t50_ppt(); print("T50 ppt")
