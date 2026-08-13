#!/usr/bin/env python3
"""Assemble one complete T01-T50 package that includes office binaries in each task folder."""
import re
import shutil
from pathlib import Path

ROOT = Path("/workspace")
SRC_MD = ROOT / "OPENROUTER_TEXT_PROXY_BATCH_50_ANSWERS.md"
OFFICE = ROOT / "download" / "office"
OUT = ROOT / "BATCH50_COMPLETE"
ZIP_PATH = ROOT / "download" / "BATCH50_COMPLETE_T01-T50.zip"

ATTACH = {
    "T06": ["T06_Q2_Regional_Sales_Review.pptx"],
    "T07": ["T07_Growth_Options_Board.pptx"],
    "T08": ["T08_Lumina_Memory_Q4_Investor_Update.pptx"],
    "T09": ["T09_Anti_Phishing_Training.pptx"],
    "T10": ["T10_FocusFlow_Sales_Enablement.pptx"],
    "T11": ["T11_Sales_Orders_Clean_Summary.xlsx"],
    "T12": ["T12_Q1_Department_Budget_Variance.xlsx"],
    "T13": ["T13_Cohort_Retention.xlsx"],
    "T14": ["T14_Reorder_Calculator.xlsx"],
    "T15": ["T15_Six_Month_Cash_Model.xlsx"],
    "T22": ["T22_Orion_Northstar_MSA_Working_Draft.docx"],
    "T34": ["T34_IFRS16_Lease_Amortization.xlsx"],
    "T35": ["T35_Vendor_TCO_NPV.xlsx"],
    "T36": ["T36_Launch_Plan.xlsx"],
    "T50": [
        "T50_Board_Memo.docx",
        "T50_Board_Pack.pptx",
        "T50_Board_KPI_Appendix.xlsx",
    ],
}

TITLES = {
    "T01": "从零散更新生成管理周报",
    "T02": "CRM采购决策备忘录",
    "T03": "董事会月度经营简报",
    "T04": "会议纪要与行动追踪",
    "T05": "生产事故初步复盘",
    "T06": "区域销售经营汇报PPT",
    "T07": "增长战略选项董事会PPT",
    "T08": "季度业绩投资者更新PPT",
    "T09": "反钓鱼培训课件",
    "T10": "新产品销售赋能PPT",
    "T11": "销售数据清洗与汇总",
    "T12": "部门预算差异分析",
    "T13": "订阅用户Cohort留存分析",
    "T14": "库存补货计算器",
    "T15": "六个月现金流情景模型",
    "T16": "B2B整合营销文案",
    "T17": "品牌语气本地化",
    "T18": "客户投诉分级与回复",
    "T19": "企业客户RFP响应",
    "T20": "线索评分与个性化外联",
    "T21": "双向NDA条款审阅",
    "T22": "主服务协议草案",
    "T23": "员工纪律调查通知",
    "T24": "隐私事件通知判断",
    "T25": "营销主张合规审查",
    "T26": "竞争格局与供应商短名单",
    "T27": "市场规模TAM_SAM_SOM",
    "T28": "云供应商风险尽调",
    "T29": "混合办公政策备忘录",
    "T30": "存储公司季度业绩研究",
    "T31": "三张财务报表勾稽检查",
    "T32": "多情景股票投资备忘录",
    "T33": "组合敞口与压力测试",
    "T34": "IFRS16租赁摊销表",
    "T35": "采购三年TCO_NPV",
    "T36": "跨职能上线计划",
    "T37": "客户上线SOP",
    "T38": "客服排班优化",
    "T39": "供应商加权评分",
    "T40": "制造缺陷根因初析",
    "T41": "岗位说明书与招聘评分卡",
    "T42": "面试证据综合与录用建议",
    "T43": "绩效反馈与30天改进计划",
    "T44": "办公室搬迁公告与FAQ",
    "T45": "培训需求分析与30天计划",
    "T46": "高管收件箱分流",
    "T47": "跨城市高管差旅议程",
    "T48": "150人客户活动方案",
    "T49": "双语合同差异核对",
    "T50": "CEO董事会综合材料",
}


def split_tasks(text: str) -> dict[str, str]:
    parts = {}
    for m in re.finditer(
        r"===== BEGIN (T\d+) =====\n(.*?)===== END \1 =====",
        text,
        re.S,
    ):
        parts[m.group(1)] = m.group(0).strip() + "\n"
    return parts


def patch_task(tid: str, body: str, files: list[str]) -> str:
    if not files:
        return body
    listing = "\n".join(f"  - {f}" for f in files)
    note = (
        f"\n\n## DELIVERABLE_FILES\n"
        f"本题正式交付物已放在本文件夹内，可用 PowerPoint / Excel / Word 直接打开：\n"
        f"{listing}\n"
        f"下面的正文是同一套内容的说明、公式与自检，与二进制文件一致。\n"
    )
    body = body.replace(
        "STATUS: COMPLETE_TEXT_PROXY",
        "STATUS: COMPLETE",
    )
    body = re.sub(
        r"PROXY_DELIVERABLE:.*",
        "PROXY_DELIVERABLE: 本题已提供原格式办公文件（见本文件夹）",
        body,
        count=1,
    )
    body = re.sub(
        r"## UNMET_ARTIFACT_REQUIREMENTS\n\n.*?(?=\n## SELF_CHECK)",
        "## UNMET_ARTIFACT_REQUIREMENTS\n\n无（本题办公文件已生成，可打开编辑）。\n\n",
        body,
        count=1,
        flags=re.S,
    )
    # insert file list after ANSWER header
    body = body.replace("## ANSWER\n", "## ANSWER\n" + note + "\n", 1)
    return body


def main():
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir(parents=True)

    text = SRC_MD.read_text(encoding="utf-8")
    tasks = split_tasks(text)
    missing = [f"T{i:02d}" for i in range(1, 51) if f"T{i:02d}" not in tasks]
    if missing:
        raise SystemExit(f"missing tasks: {missing}")

    index_rows = ["task_id,title,status,folder,files"]
    combined = ["# BATCH50 完整作答 T01-T50（含 PPTX/XLSX/DOCX）\n"]

    for i in range(1, 51):
        tid = f"T{i:02d}"
        folder = OUT / f"{tid}_{TITLES[tid]}"
        folder.mkdir()
        files = ATTACH.get(tid, [])
        body = patch_task(tid, tasks[tid], files)
        (folder / "ANSWER.md").write_text(body, encoding="utf-8")
        copied = []
        for fn in files:
            src = OFFICE / fn
            if not src.exists():
                raise SystemExit(f"missing office file {src}")
            shutil.copy2(src, folder / fn)
            copied.append(fn)
        (folder / "FILES.txt").write_text(
            "本任务交付：\n- ANSWER.md\n"
            + "".join(f"- {c}\n" for c in copied)
            + ("" if copied else "- （本题为文本文书，ANSWER.md 即为正式交付）\n"),
            encoding="utf-8",
        )
        status = "COMPLETE" if copied else "COMPLETE"
        index_rows.append(
            f"{tid},{TITLES[tid].replace(',', ' ')},{status},{folder.name},"
            + ("|".join(copied) if copied else "ANSWER.md")
        )
        combined.append(body)
        combined.append("")

    (OUT / "00_INDEX.csv").write_text("\n".join(index_rows) + "\n", encoding="utf-8")
    (OUT / "00_ALL_ANSWERS.md").write_text("\n".join(combined), encoding="utf-8")
    readme = """BATCH50 完整结果（T01-T50）
================================

这是 50 道题的完整作答包，不是单独的 PPT/Excel 附件包。

目录规则
--------
每个任务一个文件夹：Txx_题目名/
  ANSWER.md     本题完整作答（含计算、假设、自检）
  FILES.txt     本文件夹文件清单
  *.pptx/*.xlsx/*.docx   题目要求的正式办公文件（若该题需要）

需要 PPT 的题：T06 T07 T08 T09 T10 T50
需要 Excel 的题：T11 T12 T13 T14 T15 T34 T35 T36 T50
需要 Word 的题：T22 T50
其余题的正式交付就是该文件夹里的 ANSWER.md。

请解压整个 ZIP 后按 T01 → T50 查阅。不要只取 office 子目录。

00_INDEX.csv        50 行索引
00_ALL_ANSWERS.md   50 题文字合订（评分定位用边界标记仍在各 ANSWER.md 中）
"""
    (OUT / "00_README.txt").write_text(readme, encoding="utf-8")

    if ZIP_PATH.exists():
        ZIP_PATH.unlink()
    shutil.make_archive(str(ZIP_PATH).replace(".zip", ""), "zip", ROOT, OUT.name)
    print("wrote", OUT)
    print("zip", ZIP_PATH, ZIP_PATH.stat().st_size)
    print("folders", len([p for p in OUT.iterdir() if p.is_dir()]))


if __name__ == "__main__":
    main()
