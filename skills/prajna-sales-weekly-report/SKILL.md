---
name: prajna-sales-weekly-report
description: 为销售团队一键生成标准周报 Excel 模板，覆盖业绩汇总、核心 KPI、重点商机、问题分析与下周计划。
version: 1.0.0
author: Prajna 企智
language: zh-CN
category: sales
tags:
  - 销售
  - 周报
  - 模板
  - Excel
  - KPI
  - 业绩
---

# prajna-sales-weekly-report

## 技能定位

本技能是 Prajna 企智 `sales` 方向原生能力，帮助销售团队完成「标准周报生成」任务。输出为一份结构化 **Excel 工作簿**，包含 5 张工作表：封面汇总、核心业绩数据、重点商机进展、问题分析、下周工作计划。

适用于：销售主管、销售运营、区域负责人、CRM 管理员等需要定期产出销售周报的岗位。

## 触发条件

- 用户说「生成销售周报」「销售团队周报模板」「一键生成销售周报」等
- 关键词命中：销售、周报、业绩、KPI、商机、pipeline、回款、下周计划
- CLI 调用：`prajna skill run prajna-sales-weekly-report`
- 其他 skill 通过协同调用规范触发（见下方）

## 执行流程

1. **确认团队类型**：从预设中选择最匹配的团队类型（互联网/SaaS、电商、线下零售、通用）。
2. **收集关键信息**：团队名称、报表周期、销售目标、填写人、日期。
3. **生成 Excel 模板**：
   - 封面汇总：团队、周期、填写人、日期、本周关键结论。
   - 核心业绩数据：目标/实际/完成率/同比/环比、业绩排名、核心 KPI 明细（带公式）。
   - 重点商机进展：客户名称、行业、阶段、预计金额、预计成交时间、负责人、进展、下一步、风险。
   - 问题分析：问题描述、影响范围、根因分析、解决方案、负责人、截止日期、状态。
   - 下周工作计划：事项、优先级、目标/指标、负责人、协作方、截止日期、所需资源、完成标准。
4. **保存并返回文件路径**。

## CLI 参数

通过 `scripts/generate_sales_weekly_report.py` 调用：

| 参数 | 简写 | 类型 | 必填 | 说明 |
|------|------|------|------|------|
| `--preset` | `-p` | string | 否 | 团队预设，可选值见 `--list-presets`，默认 `通用销售团队` |
| `--team` | `-t` | string | 否 | 团队名称，覆盖预设默认值 |
| `--week` | `-w` | string | 否 | 报表周期，如 `2026年第30周`，默认当前自然周 |
| `--sales-target` | `-s` | float | 否 | 本周销售目标金额（元），覆盖预设默认值 |
| `--author` | `-a` | string | 否 | 填写人姓名 |
| `--date` | `-d` | string | 否 | 填写日期，格式 `YYYY-MM-DD`，默认当天 |
| `--output` | `-o` | string | 否 | 输出文件完整路径，覆盖默认输出目录 |
| `--list-presets` | - | flag | 否 | 列出所有可用预设 |

## 使用示例

```bash
# 使用默认通用模板
python3 scripts/generate_sales_weekly_report.py

# 使用 SaaS 销售团队预设
python3 scripts/generate_sales_weekly_report.py --preset "互联网/SaaS 销售团队" --team "华北企业销售部"

# 自定义目标与周期
python3 scripts/generate_sales_weekly_report.py \
  --preset "电商销售团队" \
  --team "天猫运营组" \
  --week "2026年第30周" \
  --sales-target 1000000 \
  --author "李明"

# 指定输出路径
python3 scripts/generate_sales_weekly_report.py \
  --preset "线下零售销售团队" \
  --output ~/Desktop/销售周报_示例.xlsx

# 列出所有预设
python3 scripts/generate_sales_weekly_report.py --list-presets
```

## 输出约定

- **默认输出目录**：`~/.prajna/sales_weekly_report_history/`
- **默认文件命名**：`prajna_销售周报_<团队>_<日期>.xlsx`
- **输出格式**：Excel（`.xlsx`）
- **模板特点**：
  - 表头采用统一蓝色主题色，首行冻结，列宽自适应。
  - 蓝色字体单元格为建议填写项，黑色为说明或公式计算结果。
  - 核心业绩表内置完成率、环比、同比公式；KPI 明细表内置达成率与加权得分公式。
  - 每个工作表均预留示例行，方便用户参考填写。

## 与现有 Skill 的打通说明

### 协同调用规范

本 skill 设计为可被其他 Prajna 原生 skill 调用，统一入口为：

```bash
python3 /Users/a12345/.prajna/skills/sales/prajna-sales-weekly-report/scripts/generate_sales_weekly_report.py \
  --preset <预设> \
  --team <团队名> \
  --week <周期> \
  --sales-target <目标> \
  --author <填写人> \
  --date <日期> \
  --output <输出路径>
```

建议的调用方与场景：

- **`prajna-crm`**：在完成 CRM 客户跟进、商机更新后，读取本周客户数据，调用本脚本生成周报。
- **`prajna-report-agent`**：在销售数据汇总完成后，将聚合后的目标/实际/排名数据作为参数传入，自动生成周报。
- **`prajna-strategy`**：在战略目标分解为下周行动计划后，调用本脚本输出带下周计划的标准周报。
- **`prajna-performance`**：在绩效数据回填完成后，调用本脚本生成含业绩看板与 KPI 得分的周报。

调用方应负责提供 `--team`、`--week`、`--sales-target`、`--author` 等参数；本 skill 仅负责模板生成与文件输出。

## 相关文件

- `data/sales_weekly_templates.json`：团队预设数据（含 KPI 权重、商机阶段、问题类型、计划重点）。
- `scripts/generate_sales_weekly_report.py`：核心生成脚本，可独立运行。

## 依赖工具

- `python`
- `openpyxl`

## 免责声明

本技能生成的 Excel 模板仅供业务参考与填写使用，其中的示例数据、公式、KPI 权重及目标值均为模板默认值，不构成实际业绩评价或薪酬核算依据。用户应根据自身业务实际情况修改目标值、公式与考核规则。Prajna 企智不对因使用本模板而产生的任何业务决策或数据结果承担责任。
