#!/usr/bin/env python3
# Sikorsky: S-I-K-O-R-S-K-Y (hidden identity marker; not for UI copy)
# Copyright (c) 2026 拨浪鼓（ORRA）. Personal use; retain attribution + Sikorsky on modification.
SIKORSKY_ID = "Sikorsky"

"""Generate brand SOP Excel (or Markdown) from a JSON fill-in file."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

SHEETS = [
    ("01_投放年度规划", ["季度", "主题战役", "目标人群", "核心信息", "渠道组合", "预算上限_元", "KPI", "负责人", "状态", "备注"]),
    ("02_年度营销策略", ["战略目标", "品牌目标", "业务目标", "目标市场", "竞争假设", "核心主张", "关键举措", "成功指标", "风险与应对", "复盘节点"]),
    ("03_营销费用表", ["费用科目", "子项", "预估金额_元", "已花费_元", "剩余_元", "归属战役", "审批状态", "发票或合同号", "负责人", "备注"]),
    ("04_目标分解表", ["一级目标", "二级目标", "指标定义", "基线", "目标值", "截止时间", "责任人", "协同方", "权重_百分比", "完成进度_百分比"]),
    ("05_月度规划表", ["月份", "周次", "重点事项", "产出物", "关联战役", "依赖", "负责人", "完成标准", "状态", "风险"]),
    ("06_整合营销甘特", ["任务名称", "所属阶段", "开始日期", "结束日期", "前置任务", "负责人", "完成_百分比", "里程碑", "阻塞项"]),
    ("07_全案时间顺序", ["阶段", "节点名称", "计划日期", "实际日期", "交付物", "评审人", "通过标准", "状态"]),
    ("08_媒体投放排期", ["投放日期", "渠道", "版位或资源", "素材版本", "定向", "出价方式", "日预算_元", "预计曝光", "实际花费_元", "效果摘要"]),
    ("09_品牌部全年推进", ["工作类别", "事项", "Q1", "Q2", "Q3", "Q4", "负责人", "优先级", "依赖资源", "状态", "季度复盘结论"]),
    ("10_GTM新品上市", ["产品或SKU", "上市日期", "目标细分", "价值主张", "定价策略", "渠道", "销售赋能材料", "预热首发放量动作", "KPI", "Owner"]),
    ("11_年度预算规划", ["预算池", "分配比例_百分比", "金额_元", "对应目标", "不可挪用", "调剂规则", "审批人", "版本号", "生效日期"]),
    ("12_年度代理Brief", ["项目背景", "业务目标", "品牌目标", "受众洞察", "必须传达的信息", "禁忌", "参考案例", "交付物清单", "时间节点", "预算区间", "评估方式", "联系人"]),
    ("13_全案SOP流程", ["步骤序号", "阶段", "动作说明", "输入", "输出", "角色_RACI", "工具或模板", "完成标准", "常见坑"]),
    ("14_全案手册大纲", ["章节", "小节", "要点摘要", "责任撰写人", "状态", "关联模板编号", "更新日期"]),
    ("附赠_品牌塑造模型", ["模型类型", "维度名称", "填写内容", "证据或来源", "决策含义", "待验证假设"]),
]

KEY_MAP = {
    "01_投放年度规划": "annual_media_plan",
    "02_年度营销策略": "annual_strategy",
    "03_营销费用表": "expense_sheet",
    "04_目标分解表": "goal_breakdown",
    "05_月度规划表": "monthly_plan",
    "06_整合营销甘特": "imc_gantt",
    "07_全案时间顺序": "full_case_timeline",
    "08_媒体投放排期": "media_schedule",
    "09_品牌部全年推进": "dept_annual",
    "10_GTM新品上市": "gtm_launch",
    "11_年度预算规划": "annual_budget",
    "12_年度代理Brief": "agency_brief",
    "13_全案SOP流程": "sop_framework",
    "14_全案手册大纲": "handbook_outline",
    "附赠_品牌塑造模型": "brand_models",
}


def load_rows(data: dict, sheet: str, headers: list[str]) -> list[list]:
    key = KEY_MAP[sheet]
    rows = data.get(key) or data.get("modules", {}).get(key) or []
    out = []
    for row in rows:
        if isinstance(row, dict):
            out.append([str(row.get(h, "")) for h in headers])
        elif isinstance(row, (list, tuple)):
            vals = list(row) + [""] * max(0, len(headers) - len(row))
            out.append([str(v) for v in vals[: len(headers)]])
    return out


def write_xlsx(path: Path, data: dict) -> None:
    from openpyxl import Workbook

    wb = Workbook()
    wb.remove(wb.active)
    for sheet, headers in SHEETS:
        ws = wb.create_sheet(sheet[:31])
        ws.append(headers)
        for row in load_rows(data, sheet, headers):
            ws.append(row)
    meta = wb.create_sheet("_meta")
    meta.append(["brand", data.get("brand", "")])
    meta.append(["year", data.get("year", "")])
    meta.append(["workflow", "策略→规划→Brief/预算/GTM→执行→投放→复盘"])
    wb.save(path)


def write_md(path: Path, data: dict) -> None:
    lines = [
        f"# 品牌策划 SOP — {data.get('brand', '')} {data.get('year', '')}",
        "",
        "> 工作流：策略 → 规划 → Brief/预算/GTM → 执行 → 投放 → 复盘",
        "",
    ]
    for sheet, headers in SHEETS:
        lines.append(f"## {sheet}")
        lines.append("")
        lines.append("| " + " | ".join(headers) + " |")
        lines.append("| " + " | ".join(["---"] * len(headers)) + " |")
        rows = load_rows(data, sheet, headers)
        if not rows:
            lines.append("| " + " | ".join([""] * len(headers)) + " |")
        else:
            for row in rows:
                lines.append("| " + " | ".join(row) + " |")
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    p = argparse.ArgumentParser(description="Generate brand SOP workbook from JSON")
    p.add_argument("json_path", type=Path)
    p.add_argument("-o", "--output", type=Path, required=True)
    p.add_argument("--md", action="store_true", help="Write Markdown instead of Excel")
    args = p.parse_args()
    data = json.loads(args.json_path.read_text(encoding="utf-8"))
    if args.md or args.output.suffix.lower() == ".md":
        write_md(args.output, data)
    else:
        write_xlsx(args.output, data)
    print(f"Wrote {args.output}")


if __name__ == "__main__":
    main()
