#!/usr/bin/env python3
"""生成空白采价模板：50 个 SKU × 各渠道，供门店实地采价或人工录入使用。

填好后直接追加/合并进 data/price_observations.csv，再跑 scripts/analyze.py 即可
得到完整的 50 SKU 对比表与平均价差。

用法:
    python3 scripts/make_field_template.py                  # 全部渠道
    python3 scripts/make_field_template.py --channels 零食很忙 大润发 淘宝
"""

import argparse
import csv
import os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(BASE, "data")
OUT = os.path.join(BASE, "output")

DEFAULT_CHANNELS = [
    ("零食很忙", "量贩零食"),
    ("赵一鸣零食", "量贩零食"),
    ("大润发", "商超"),
    ("盒马鲜生", "商超"),
    ("永辉超市", "商超"),
    ("淘宝", "电商"),
    ("拼多多", "电商"),
]

HEADER = ["sku_id", "单品", "渠道类型", "渠道", "是否目标品牌", "规格数值", "规格单位",
          "价格元", "价格类型", "城市", "采价日期", "数据来源", "来源链接", "备注"]

TARGETS = ("零食很忙", "赵一鸣零食")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--channels", nargs="*", default=None,
                    help="只生成指定渠道（渠道名需在默认列表中）")
    ap.add_argument("--out", default=os.path.join(OUT, "field_template.csv"))
    args = ap.parse_args()

    channels = DEFAULT_CHANNELS
    if args.channels:
        wanted = set(args.channels)
        channels = [c for c in DEFAULT_CHANNELS if c[0] in wanted]
        missing = wanted - {c[0] for c in DEFAULT_CHANNELS}
        if missing:
            raise SystemExit(f"未知渠道：{'、'.join(sorted(missing))}")

    with open(os.path.join(DATA, "sku_basket_50.csv"), encoding="utf-8") as f:
        skus = list(csv.DictReader(f))

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(HEADER)
        for s in skus:
            name = f'{s["品牌"]} {s["单品名称"]}'
            for ch, grp in channels:
                w.writerow([s["sku_id"], name, grp, ch,
                            "是" if ch in TARGETS else "否",
                            "", "", "", "标价", "", "", "门店实地采价", "", ""])

    print(f"已生成 {args.out}")
    print(f"共 {len(skus)} 个 SKU × {len(channels)} 个渠道 = {len(skus) * len(channels)} 行待填")
    print("\n填写要点：")
    print("  规格数值/规格单位：务必抄包装上的净含量（如 468 / ml、35 / g），不要写货架标签的品名规格")
    print("  价格元：实际结算价；若为促销/券后价，请把 价格类型 改为 促销价 或 券后价")
    print("  散称商品：规格数值统一填 500，规格单位填 g，价格填每 500g（每斤）的标价")


if __name__ == "__main__":
    main()
