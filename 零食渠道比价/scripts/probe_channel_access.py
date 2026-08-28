#!/usr/bin/env python3
"""渠道线上可采价性审计：逐个访问各渠道的公开入口，记录是否能拿到商品价格。

这是本次调研得出「哪些渠道能远程采价、哪些必须到店或登录」结论的依据，
可随时重跑以验证结论是否仍然成立（各平台风控策略会变）。

依赖:
    pip install playwright   # 复用系统已安装的 Chrome，无需 playwright install

用法:
    python3 scripts/probe_channel_access.py
    python3 scripts/probe_channel_access.py --chrome /usr/bin/google-chrome
"""

import argparse
import asyncio
import json
import os

UA_DESKTOP = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36")
UA_MOBILE = ("Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 "
             "(KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1")

# (渠道, 目标, URL, 是否用移动端 UA)
PROBES = [
    ("淘宝", "商品搜索", "https://s.taobao.com/search?q=%E4%B9%90%E4%BA%8B%E8%96%AF%E7%89%87", False),
    ("拼多多", "商品搜索", "https://mobile.yangkeduo.com/search_result.html?search_key=%E4%B9%90%E4%BA%8B%E8%96%AF%E7%89%87", True),
    ("京东", "商品搜索", "https://search.jd.com/Search?keyword=%E4%B9%90%E4%BA%8B%E8%96%AF%E7%89%87&enc=utf-8", False),
    ("盒马", "商品搜索", "https://www.freshhema.com/searchList?keyword=%E4%B9%90%E4%BA%8B%E8%96%AF%E7%89%87", True),
    ("永辉超市", "官网商品目录", "https://www.yonghui.com.cn/", True),
    ("大润发", "官网商品目录", "https://www.rt-mart.com.cn/", True),
    ("京东到家", "门店商品搜索", "https://daojia.jd.com/html/main/search?keyword=%E9%9B%B6%E9%A3%9F%E5%BE%88%E5%BF%99", True),
    ("美团外卖", "门店商品列表", "https://h5.waimai.meituan.com/waimai/mindex/kingkong?navigateType=1", True),
    ("苏宁易购", "商品搜索", "https://search.suning.com/%E4%B9%90%E4%BA%8B%E8%96%AF%E7%89%87/", False),
]

BLOCK_MARKERS = ["请登录", "亲，请登录", "欢迎登录", "手机登录", "扫码登录",
                 "身份核实", "验证", "请下载", "打开APP", "立即打开", "页面不存在"]


async def probe(ctx_desktop, ctx_mobile, channel, target, url, mobile):
    ctx = ctx_mobile if mobile else ctx_desktop
    pg = await ctx.new_page()
    result = {"渠道": channel, "目标": target, "URL": url}
    try:
        await pg.goto(url, timeout=40000, wait_until="domcontentloaded")
        await pg.wait_for_timeout(6000)
        final = pg.url
        title = await pg.title()
        body = await pg.evaluate("document.body ? document.body.innerText : ''")
        prices = await pg.evaluate(
            r"""(() => {
                   const t = document.body ? document.body.innerText : '';
                   return (t.match(/[¥￥]\s?\d+(\.\d{1,2})?/g) || []).slice(0, 8);
                 })()""")
        hits = [m for m in BLOCK_MARKERS if m in body or m in title]
        result.update({
            "最终URL": final,
            "标题": title,
            "重定向到登录/验证": bool(hits) or "login" in final or "passport" in final or "verify" in final,
            "命中拦截关键词": hits,
            "页面上抓到的价格样本": prices,
            "结论": "可采价" if prices and not hits else "不可采价",
        })
    except Exception as e:
        result.update({"结论": "不可采价", "错误": f"{type(e).__name__}: {str(e)[:160]}"})
    await pg.close()
    return result


async def main_async(chrome):
    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        b = await p.chromium.launch(executable_path=chrome, headless=True,
                                    args=["--no-sandbox",
                                          "--disable-blink-features=AutomationControlled"])
        ctx_d = await b.new_context(user_agent=UA_DESKTOP, locale="zh-CN",
                                    viewport={"width": 1440, "height": 900})
        ctx_m = await b.new_context(user_agent=UA_MOBILE, locale="zh-CN",
                                    viewport={"width": 414, "height": 896},
                                    is_mobile=True, has_touch=True,
                                    geolocation={"latitude": 28.2278, "longitude": 112.9388},
                                    permissions=["geolocation"])
        results = []
        for channel, target, url, mobile in PROBES:
            r = await probe(ctx_d, ctx_m, channel, target, url, mobile)
            results.append(r)
            print(f'[{r["结论"]}] {channel} / {target} -> {r.get("标题", "")}'
                  f'{"  拦截:" + "、".join(r["命中拦截关键词"]) if r.get("命中拦截关键词") else ""}',
                  flush=True)
        await b.close()
        return results


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--chrome", default="/usr/local/bin/google-chrome",
                    help="Chrome/Chromium 可执行文件路径")
    ap.add_argument("--out", default=os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        "output", "channel_access_audit.json"))
    args = ap.parse_args()

    if not os.path.exists(args.chrome):
        raise SystemExit(f"找不到 Chrome：{args.chrome}，请用 --chrome 指定路径")

    results = asyncio.run(main_async(args.chrome))
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    ok = [r["渠道"] for r in results if r["结论"] == "可采价"]
    no = [r["渠道"] for r in results if r["结论"] != "可采价"]
    print(f"\n可远程采价：{'、'.join(ok) if ok else '无'}")
    print(f"不可远程采价：{'、'.join(no)}")
    print(f"明细已写入 {args.out}")


if __name__ == "__main__":
    main()
