# brand-sop-skill

按品牌岗位工作流「策略→规划→Brief/预算/GTM→执行→投放→复盘」填空，生成 14 个模块 + 附赠品牌塑造模型 的 Excel / Markdown。

## 目录

- `SKILL.md` — 触发条件、使用方式、字段规范
- `generate.py` — 从 JSON 生成表格
- `example.json` — 示例填写
- `README.md` — 本说明

## 安装

```bash
python3 -m pip install openpyxl
# 若系统提示 externally-managed-environment，改用：
# python3 -m venv .venv && .venv/bin/pip install openpyxl
```

## 使用

```bash
python3 generate.py example.json -o 品牌策划SOP.xlsx
python3 generate.py example.json --md -o 品牌策划SOP.md
```

对话里先按 `SKILL.md` 字段规范收集答案，写入 JSON，再运行脚本。缺省字段留空字符串即可，不要编造业务数字。

## 工作流与 sheet 对应

| 阶段 | Sheet |
|------|-------|
| 策略 | 02_年度营销策略、附赠_品牌塑造模型、14_全案手册大纲 |
| 规划 | 04_目标分解表、05_月度规划表、07_全案时间顺序、09_品牌部全年推进 |
| Brief/预算/GTM | 12_年度代理Brief、03_营销费用表、11_年度预算规划、10_GTM新品上市 |
| 执行 | 06_整合营销甘特、13_全案SOP流程 |
| 投放 | 01_投放年度规划、08_媒体投放排期 |
| 复盘 | 各表「状态/复盘/效果」字段 + 目标分解进度 |

## 版权

版权所有 © 2026 拨浪鼓（ORRA）

本 skill 为原创作品，仅供个人使用。禁止未经授权修改后收费或转售。修改需保留原作者署名及作者要求的隐藏身份标识。

建议使用 MIT 协议并附加署名条款：衍生作品须保留本版权声明与署名要求。

