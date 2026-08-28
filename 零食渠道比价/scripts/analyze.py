#!/usr/bin/env python3
"""按 SKU 汇总各渠道价格，计算单件价差与单位价格差（克价/毫升价）。

方法要点
--------
1. 每个 (SKU, 渠道类型) 只取「一条」代表观测，避免把不同连锁、不同规格的价格平均成
   一个没有实际含义的数字。代表观测按渠道优先级 + 价格类型优先级挑选。
2. 由于量贩渠道大量使用「渠道限定规格」，跨渠道直接比单件价往往无意义。因此主指标为
   单位价格（元/g 或 元/ml）；单件价差只在规格完全一致时才计算。
3. 对规格不一致的 SKU，额外给出「折算价」：把量贩单位价乘以对照渠道的规格，得到
   “同样分量下量贩要多少钱”，从而得到可加总的元金额差。
4. 促销价/券后价默认参与计算，但会单独标注，并输出仅含常规标价的敏感性结果。

用法:
    python3 scripts/analyze.py
"""

import csv
import os
from collections import defaultdict
from statistics import mean

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, "data")
OUT = os.path.join(BASE, "output")

TARGET_CHAINS = ("零食很忙", "赵一鸣零食")
GROUPS = ["量贩零食", "商超", "便利店", "电商"]
COMPARE_GROUPS = ["商超", "便利店", "电商"]

# 渠道代表性优先级：越靠前越优先被选为该渠道类型的代表观测
CHANNEL_PRIORITY = [
    "零食很忙", "赵一鸣零食", "好想来", "好想来/老婆大人", "老婆大人", "吖嘀吖嘀",
    "大润发", "大润发等连锁超市", "永辉超市", "盒马鲜生",
    "世纪联华", "世纪联华/物美", "物美超市", "济南华联/张小年生活超市",
    "淘宝", "天猫", "拼多多", "京东自营",
    "罗森", "7-Eleven", "美宜佳", "美宜佳等便利店", "京东便利店",
]
# 价格类型优先级：常规标价优先于促销/券后价
PRICE_TYPE_PRIORITY = ["标价", "整包折算", "日常标价", "页面价", "折后价",
                       "拼单价", "促销价", "券后价", "实付价"]
PROMO_TYPES = {"促销价", "券后价", "实付价", "折后价", "拼单价"}


def rank(value, table, default=999):
    try:
        return table.index(value)
    except ValueError:
        return default


def load_basket():
    with open(os.path.join(DATA, "sku_basket_50.csv"), encoding="utf-8") as f:
        return {r["sku_id"]: r for r in csv.DictReader(f)}


def load_observations():
    rows = []
    with open(os.path.join(DATA, "price_observations.csv"), encoding="utf-8") as f:
        for r in csv.DictReader(f):
            r["规格数值"] = float(r["规格数值"])
            r["价格元"] = float(r["价格元"])
            r["单位价"] = r["价格元"] / r["规格数值"]
            r["是促销"] = r["价格类型"] in PROMO_TYPES
            rows.append(r)
    return rows


def pick(rows, prefer_targets=False):
    """从候选观测中挑一条代表值。"""
    if not rows:
        return None

    def key(r):
        is_target = r["渠道"] in TARGET_CHAINS
        return (
            0 if (prefer_targets and is_target) else 1,
            rank(r["价格类型"], PRICE_TYPE_PRIORITY),
            rank(r["渠道"], CHANNEL_PRIORITY),
        )

    return sorted(rows, key=key)[0]


def main():
    basket = load_basket()
    obs = load_observations()

    by_sku = defaultdict(lambda: defaultdict(list))
    for r in obs:
        by_sku[r["sku_id"]][r["渠道类型"]].append(r)

    records = []
    for sku_id in sorted(by_sku):
        meta = basket.get(sku_id, {})
        lf = pick(by_sku[sku_id].get("量贩零食", []), prefer_targets=True)
        if not lf:
            continue
        rec = {
            "sku_id": sku_id,
            "品类": meta.get("品类", ""),
            "单品": (meta.get("品牌", "") + " " + meta.get("单品名称", "")).strip(),
            "量贩渠道": lf["渠道"],
            "是零食很忙或赵一鸣": "是" if lf["渠道"] in TARGET_CHAINS else "否",
            "量贩规格": f'{lf["规格数值"]:g}{lf["规格单位"]}',
            "量贩价": lf["价格元"],
            "量贩单位价": lf["单位价"],
            "量贩价格类型": lf["价格类型"],
            "含促销": lf["是促销"],
        }
        for grp in COMPARE_GROUPS:
            other = pick(by_sku[sku_id].get(grp, []))
            prefix = f"vs{grp}"
            if not other:
                for suffix in ("_渠道", "_规格", "_价", "_单位价", "_单件价差",
                               "_折算价差", "_单位价便宜%", "_同规格", "_含促销"):
                    rec[prefix + suffix] = None
                continue
            same_spec = (other["规格数值"] == lf["规格数值"]
                         and other["规格单位"] == lf["规格单位"])
            rec[prefix + "_渠道"] = other["渠道"]
            rec[prefix + "_规格"] = f'{other["规格数值"]:g}{other["规格单位"]}'
            rec[prefix + "_价"] = other["价格元"]
            rec[prefix + "_单位价"] = other["单位价"]
            rec[prefix + "_同规格"] = same_spec
            rec[prefix + "_含促销"] = other["是促销"] or lf["是促销"]
            rec[prefix + "_单件价差"] = round(other["价格元"] - lf["价格元"], 2) if same_spec else None
            # 折算价差：把量贩单位价折算到对照渠道的规格后的金额差
            rec[prefix + "_折算价差"] = round(
                other["价格元"] - lf["单位价"] * other["规格数值"], 2)
            rec[prefix + "_单位价便宜%"] = round(
                (other["单位价"] - lf["单位价"]) / other["单位价"] * 100, 1)
        records.append(rec)

    os.makedirs(OUT, exist_ok=True)
    with open(os.path.join(OUT, "comparison_table.csv"), "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(records[0].keys()))
        w.writeheader()
        w.writerows(records)

    L = []
    L.append("# 逐 SKU 价格对比表（由 scripts/analyze.py 自动生成，勿手工编辑）\n")
    L.append(f"- 已取得可核验价格的 SKU：**{len(records)} / 50**")
    L.append(f"- 其中含「零食很忙 / 赵一鸣」门店实价的 SKU："
             f"**{sum(1 for r in records if r['是零食很忙或赵一鸣'] == '是')}**")
    L.append("- 主指标为单位价格（元/g、元/ml）；单件价差仅在规格完全一致时给出。")
    L.append("- 「折算价差」= 对照渠道售价 − 量贩单位价 × 对照渠道规格，即“买同样分量时贵多少钱”。\n")

    L.append("\n## 一、量贩零食 vs 商超（大润发 / 永辉 / 世纪联华 / 物美 / 华联等）\n")
    L.append("| SKU | 品类 | 单品 | 量贩门店 | 量贩规格 | 量贩价 | 商超 | 商超规格 | 商超价 | 同规格单件价差 | 折算价差 | 量贩单位价便宜 |")
    L.append("|---|---|---|---|---|---|---|---|---|---|---|---|")
    for r in records:
        if r.get("vs商超_价") is None:
            continue
        L.append("| {} | {} | {} | {} | {} | {:.2f} | {} | {} | {:.2f} | {} | {:+.2f} | {:+.1f}% |".format(
            r["sku_id"], r["品类"], r["单品"], r["量贩渠道"], r["量贩规格"], r["量贩价"],
            r["vs商超_渠道"], r["vs商超_规格"], r["vs商超_价"],
            f'{r["vs商超_单件价差"]:+.2f}' if r["vs商超_单件价差"] is not None else "—",
            r["vs商超_折算价差"], r["vs商超_单位价便宜%"]))

    for idx, grp in zip(("二", "三"), ("便利店", "电商")):
        sub = [r for r in records if r.get(f"vs{grp}_价") is not None]
        if not sub:
            continue
        L.append(f"\n## {idx}、量贩零食 vs {grp}\n")
        L.append(f"| SKU | 单品 | 量贩规格 | 量贩价 | {grp} | {grp}规格 | {grp}价 | 折算价差 | 量贩单位价便宜 |")
        L.append("|---|---|---|---|---|---|---|---|---|")
        for r in sub:
            L.append("| {} | {} | {} | {:.2f} | {} | {} | {:.2f} | {:+.2f} | {:+.1f}% |".format(
                r["sku_id"], r["单品"], r["量贩规格"], r["量贩价"],
                r[f"vs{grp}_渠道"], r[f"vs{grp}_规格"], r[f"vs{grp}_价"],
                r[f"vs{grp}_折算价差"], r[f"vs{grp}_单位价便宜%"]))

    L.append("\n## 四、平均价格差汇总\n")
    L.append("| 对比 | 可比 SKU 数 | 平均单位价格差 | 平均折算价差(元) | 量贩更便宜的 SKU |")
    L.append("|---|---|---|---|---|")
    for grp in COMPARE_GROUPS:
        sub = [r for r in records if r.get(f"vs{grp}_价") is not None]
        if not sub:
            continue
        pcts = [r[f"vs{grp}_单位价便宜%"] for r in sub]
        amts = [r[f"vs{grp}_折算价差"] for r in sub]
        win = sum(1 for p in pcts if p > 0)
        L.append("| 量贩零食 vs {} | {} | 量贩便宜 {:.1f}% | {:+.2f} | {}/{} ({:.0f}%) |".format(
            grp, len(sub), mean(pcts), mean(amts), win, len(sub), win / len(sub) * 100))

    L.append("\n### 口径 A：规格完全一致的 SKU（最严格，可直接比标价）\n")
    same = [r for r in records if r.get("vs商超_同规格")]
    if same:
        d = [r["vs商超_单件价差"] for r in same]
        p = [(r["vs商超_价"] - r["量贩价"]) / r["vs商超_价"] * 100 for r in same]
        L.append(f"- 可比 SKU：{len(same)} 个（{', '.join(r['sku_id'] for r in same)}）")
        L.append(f"- 平均单件价差：**{mean(d):+.2f} 元**（商超价 − 量贩价）")
        L.append(f"- 平均价差比例：**量贩比商超便宜 {mean(p):.1f}%**")

    L.append("\n### 口径 B：仅含「零食很忙 / 赵一鸣」门店实价的 SKU\n")
    tgt = [r for r in records if r["是零食很忙或赵一鸣"] == "是" and r.get("vs商超_价") is not None]
    if tgt:
        p = [r["vs商超_单位价便宜%"] for r in tgt]
        a = [r["vs商超_折算价差"] for r in tgt]
        L.append(f"- 可比 SKU：{len(tgt)} 个（{', '.join(r['sku_id'] for r in tgt)}）")
        L.append(f"- 平均单位价格差：**量贩比商超便宜 {mean(p):.1f}%**")
        L.append(f"- 平均折算价差：**{mean(a):+.2f} 元**")

    L.append("\n### 口径 C：剔除所有促销价 / 券后价，仅用常规标价\n")
    reg = [r for r in records if r.get("vs商超_价") is not None and not r.get("vs商超_含促销")]
    if reg:
        p = [r["vs商超_单位价便宜%"] for r in reg]
        a = [r["vs商超_折算价差"] for r in reg]
        L.append(f"- 可比 SKU：{len(reg)} 个（{', '.join(r['sku_id'] for r in reg)}）")
        L.append(f"- 平均单位价格差：**量贩比商超便宜 {mean(p):.1f}%**")
        L.append(f"- 平均折算价差：**{mean(a):+.2f} 元**")

    L.append("\n## 五、按品类平均单位价格差（量贩 vs 商超）\n")
    L.append("| 品类 | 可比 SKU 数 | 量贩单位价平均便宜 |")
    L.append("|---|---|---|")
    bycat = defaultdict(list)
    for r in records:
        if r.get("vs商超_价") is not None:
            bycat[r["品类"]].append(r["vs商超_单位价便宜%"])
    for cat in sorted(bycat, key=lambda c: -mean(bycat[c])):
        L.append(f"| {cat} | {len(bycat[cat])} | {mean(bycat[cat]):+.1f}% |")

    L.append("\n## 六、量贩渠道单位价格反而更贵的 SKU（反例）\n")
    L.append("| SKU | 单品 | 量贩 | 商超 | 量贩单位价便宜 |")
    L.append("|---|---|---|---|---|")
    for r in records:
        if r.get("vs商超_单位价便宜%") is not None and r["vs商超_单位价便宜%"] <= 0:
            L.append("| {} | {} | {} {} {:.2f}元 | {} {} {:.2f}元 | {:+.1f}% |".format(
                r["sku_id"], r["单品"], r["量贩渠道"], r["量贩规格"], r["量贩价"],
                r["vs商超_渠道"], r["vs商超_规格"], r["vs商超_价"], r["vs商超_单位价便宜%"]))

    md = "\n".join(L) + "\n"
    with open(os.path.join(OUT, "comparison_table.md"), "w", encoding="utf-8") as f:
        f.write(md)
    print(md)


if __name__ == "__main__":
    main()
