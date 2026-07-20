---
name: prajna-budget-execution-ppt
description: 面向管理层一键生成本月预算执行情况汇报 PPTX，包含预算完成率对比图、超支项目分析、结余资金使用建议、下月预算调整方案，首页设计高级大气。
version: 1.0.0
author: Prajna 企智
language: zh-CN
category: finance
tags:
  - 预算
  - 财务汇报
  - PPT
  - 管理报表
  - 费用分析
  - 预算调整
---

# prajna-budget-execution-ppt

## 技能定位

本技能是 Prajna 企智 `finance` 方向原生能力，帮助财务、运营及管理层快速产出「本月预算执行情况汇报」演示文稿。输出为一份结构化 **PPTX**，共 9 页，覆盖总体执行概览、预算完成率对比图、超支项目分析、偏差根因、结余资金使用建议、下月预算调整方案及总结行动。

适用于：财务BP、预算经理、费用会计、运营负责人、高管助理等需要定期向管理层汇报预算进度的岗位。

## 触发条件

- 用户说「生成本月预算执行情况汇报」「预算执行PPT」「费用使用进度汇报」「预算偏差分析PPT」等
- 关键词命中：预算、执行率、超支、结余、费用、预算调整、管理层汇报
- CLI 调用：`prajna skill run prajna-budget-execution-ppt`
- 其他 skill 通过协同调用规范触发（见下方）

## 执行流程

1. **加载数据**：读取 `data/budget_execution_sample.json` 中的示例预算数据；若缺失则使用内置默认数据。
2. **可选覆盖参数**：根据用户输入覆盖月份、公司名称、汇报部门。
3. **生成 PPTX**：
   - 封面：深蓝底色 + 金色装饰线，高级大气，展示标题、公司、月份、汇报部门、日期。
   - 本月预算执行总览：总预算、实际支出、总体完成率、偏差额、超支/结余项目数。
   - 预算完成率对比图：各费用项目预算 vs 实际支出的簇状条形图。
   - 完成率分布与排名：柱状图 + 100% 基准线，直观识别超支/结余项目。
   - 超支项目分析：费用项目、预算、实际、偏差额、完成率、责任部门、偏差原因。
   - 偏差根因分析：外部市场、业务增长、费用集中投放、内部管控四类因素。
   - 结余资金使用建议：费用项目、结余金额、建议使用方向、预期收益、责任部门。
   - 下月预算调整方案：费用项目、原预算、调整金额、调整后预算、调整说明。
   - 总结与下一步行动：核心结论、风险预警、下月重点。
4. **保存并返回文件路径**。

## CLI 参数

通过 `scripts/generate_budget_execution_ppt.py` 调用：

| 参数 | 简写 | 类型 | 必填 | 说明 |
|------|------|------|------|------|
| `--month` | `-m` | string | 否 | 汇报月份，例如 `2026年7月` |
| `--company` | `-c` | string | 否 | 公司名称，覆盖默认值 |
| `--author` | `-a` | string | 否 | 汇报部门/作者 |
| `--data` | `-d` | string | 否 | 自定义数据 JSON 文件路径 |
| `--output` | `-o` | string | 否 | 输出文件完整路径，覆盖默认输出目录 |

## 使用示例

```bash
# 使用默认示例数据生成
python3 scripts/generate_budget_execution_ppt.py

# 指定公司、月份和汇报部门
python3 scripts/generate_budget_execution_ppt.py \
  --company "星河创新科技" \
  --month "2026年7月" \
  --author "财务部"

# 使用自定义预算数据
python3 scripts/generate_budget_execution_ppt.py \
  --data ./my_budget_data.json \
  --output ~/Desktop/预算执行汇报_星河创新.pptx
```

## 输出约定

- **默认输出目录**：`~/.prajna/prajna-budget-execution-ppt/samples/`
- **默认文件命名**：`本月预算执行情况汇报_<公司>_<月份>.pptx`
- **输出格式**：PowerPoint（`.pptx`），共 9 页
- **设计特点**：
  - 封面采用深蓝色背景 + 金色装饰线，标题居中，简洁大气。
  - 内容页统一蓝色标题栏，关键指标卡片化展示。
  - 超支项使用红色标注，结余项使用常规黑色/灰色，便于快速识别。
  - 内置簇状条形图、柱状图等可视化图表，无需额外配置。

## 数据文件格式

自定义 JSON 可参考 `data/budget_execution_sample.json`：

```json
{
  "month": "2026年7月",
  "company": "启明星科技有限公司",
  "author": "财务部",
  "categories": [
    {"name": "人力成本", "budget": 1200000, "actual": 1250000,
     "dept": "人力资源部/财务部", "reason": "招聘进度超预期...", "severity": "high"}
  ],
  "root_causes": [
    {"category": "外部市场因素", "detail": "..."}
  ],
  "surplus_suggestions": [
    {"name": "差旅费", "surplus": 15000, "usage": "...", "benefit": "...", "dept": "销售部"}
  ],
  "next_month_adjustments": [
    {"name": "市场推广费", "original_budget": 300000, "adjustment": 35000, "reason": "..."}
  ]
}
```

## 与现有 Skill 的打通说明

### 协同调用规范

本 skill 设计为可被其他 Prajna 原生 skill 调用，统一入口为：

```bash
python3 /Users/a12345/.prajna/skills/finance/prajna-budget-execution-ppt/scripts/generate_budget_execution_ppt.py \
  --month <月份> \
  --company <公司> \
  --author <汇报部门> \
  --data <数据JSON路径> \
  --output <输出路径>
```

建议的调用方与场景：

- **`prajna-finance-report`**：在月度费用数据汇总后，调用本脚本生成管理层汇报 PPT。
- **`prajna-budget-control`**：在预算预警触发后，传入超支/结余项目数据，自动生成偏差分析 PPT。
- **`prajna-erp-connector`**：从 ERP/费控系统拉取实际入账数据后，生成本月预算执行汇报。
- **`prajna-meeting-agent`**：在管理层会议前，自动调用本技能生成会议材料。

调用方应负责提供 `--month`、`--company`、`--author`、`--data` 等参数；本 skill 仅负责 PPT 生成与文件输出。

## 依赖工具

- `python`
- `python-pptx`

若环境未安装 `python-pptx`，脚本会优雅退出并提示安装命令，不会破坏其他任务。

## 免责声明

> 【人工智能生成-需人工核验】本技能生成的 PPT 及示例数据仅供业务汇报和内部管理参考，不构成财务审计、投资决策或法律意见。预算数字、偏差原因、调整方案均应根据企业真实财务数据和内部审批流程核实后使用。Prajna 企智不对因使用本材料而产生的任何业务决策承担责任。
