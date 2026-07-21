---
name: prajna-compensation-system
description: 为企业一键生成完整的薪酬体系 Excel 工作簿，包含岗位薪酬矩阵、薪酬结构、城市系数、调薪方案、年度预算、人力成本测算与绩效联动规则，与薪资模板、招聘助手、绩效体系助手数据打通。
version: 1.0.0
author: prajna 企智
language: zh-CN
category: hr
tags:
  - 人力资源
  - 薪酬体系
  - 薪酬矩阵
  - 职级带宽
  - 城市系数
  - 调薪方案
  - 人力成本
  - 薪酬预算
  - 绩效联动
  - Excel
---

# prajna-compensation-system

## 技能定位

本技能是 prajna 企智 `hr` 方向原生能力，面向企业 HR、财务 BP 与管理层，一键生成标准化的 **薪酬体系 Excel 工作簿**。

工作簿包含 7 大核心工作表，覆盖从薪酬框架设计、结构拆分、城市差异、调薪策略、预算编制到人力成本测算与绩效联动的完整闭环，帮助企业快速搭建可落地、可迭代、可与其他 prajna 技能协同的薪酬体系。

适用于：互联网企业、制造业、电商公司、零售企业以及任何需要规范化薪酬管理的中型企业。

## 触发条件

- 用户说「生成薪酬体系」「搭建薪酬体系」「做薪酬矩阵」「年度薪酬预算」「人力成本测算」等
- 关键词命中：薪酬、薪资、工资、职级、带宽、城市系数、调薪、预算、人力成本、绩效联动
- CLI 调用：`prajna skill run prajna-compensation-system`
- 其他 skill 通过协同调用规范触发（见下方）

## 执行流程

1. **确认企业信息**：企业名称、行业、总部城市、人员规模、岗位序列、职级序列、年度薪酬预算、普调幅度。
2. **加载基准数据**：优先复用 `prajna-salary-template` 中的职级带宽、城市薪酬系数与岗位类型模板。
3. **生成 7 大工作表**：
   - 岗位薪酬矩阵（岗位序列 × 职级，最小值/中位值/最大值）
   - 薪酬结构表（基本工资/岗位工资/绩效工资/津贴/奖金占比）
   - 城市薪酬系数
   - 调薪方案（普调/晋升/绩效/特殊保留）
   - 薪酬预算表（年度人力成本预算）
   - 人力成本测算（社保公积金/个税/年终奖/离职成本）
   - 薪酬与绩效联动规则
4. **附加使用说明与免责声明**。
5. **保存并返回文件路径**。

## CLI 参数

通过 `scripts/generate_compensation_system.py` 调用：

| 参数 | 简写 | 类型 | 必填 | 说明 |
|------|------|------|------|------|
| `--company` | `-c` | string | 否 | 企业名称，默认 `智云科技` |
| `--industry` | | string | 否 | 行业，默认 `互联网` |
| `--city` | | string | 否 | 总部城市，默认 `广州` |
| `--scale` | | integer | 否 | 企业规模（人数），默认 `200` |
| `--seqs` | | string | 否 | 岗位序列，逗号分隔 |
| `--levels` | | string | 否 | 职级序列，逗号分隔，默认 `P1,P2,P3,P4,P5,P6,P7` |
| `--budget` | | number | 否 | 年度薪酬总预算（元），默认 `30000000` |
| `--growth` | | number | 否 | 年度普调幅度（%），默认 `5.0` |
| `--output` | `-o` | string | 否 | 输出文件完整路径，覆盖默认路径 |

## 使用示例

```bash
# 使用默认参数生成
python3 scripts/generate_compensation_system.py

# 指定企业与预算
python3 scripts/generate_compensation_system.py \
  --company "智云科技" \
  --industry 互联网 \
  --city 广州 \
  --scale 200 \
  --budget 30000000

# 自定义岗位序列与职级
python3 scripts/generate_compensation_system.py \
  --company "锐行制造" \
  --industry 制造业 \
  --city 东莞 \
  --seqs "生产类,质量类,供应链类,行政类" \
  --levels "P1,P2,P3,P4,P5" \
  --budget 15000000 \
  --growth 4.5
```

## 输出约定

- **默认输出目录**：`~/.prajna/prajna-compensation-system/samples/`
- **默认文件命名**：`prajna_薪酬体系_<企业名>_<行业>_<城市>_<日期>.xlsx`
- **输出格式**：Excel（`.xlsx`）
- **工作表清单**：
  1. 岗位薪酬矩阵
  2. 薪酬结构表
  3. 城市薪酬系数
  4. 调薪方案
  5. 薪酬预算表
  6. 人力成本测算
  7. 薪酬与绩效联动规则
  8. 使用说明与免责声明

## 与现有 Skill 的打通说明

### 数据打通

- **`prajna-salary-template`**：本技能优先读取其 `data/city_salary_multipliers.json` 与 `data/job_level_bandwidth.json`，确保城市系数、职级带宽、岗位类型口径一致。
- **`prajna-recruitment-assistant`**：招聘助手的「Offer 薪资建议」可直接引用本技能生成的「岗位薪酬矩阵」与「薪酬结构表」，实现城市/职级/序列基准统一。
- **`prajna-performance-system`**：绩效体系助手的 KPI 指标库、绩效评分表可对照本技能「薪酬与绩效联动规则」设计绩效奖金、年终奖与调薪方案。

### 协同调用规范

其他 skill 可通过以下命令调用本技能生成薪酬基准：

```bash
python3 /Users/a12345/.prajna/skills/hr/prajna-compensation-system/scripts/generate_compensation_system.py \
  --company <企业名> \
  --industry <行业> \
  --city <城市> \
  --scale <人数> \
  --budget <年度预算> \
  --growth <普调幅度>
```

调用方负责提供企业基本信息，本技能负责生成标准化薪酬体系工作簿。

## 相关文件

- `scripts/generate_compensation_system.py`：核心生成脚本，可独立运行。

## 依赖工具

- `python`
- `openpyxl`
- 标准库：`argparse`、`datetime`、`json`、`pathlib`、`re`、`sys`

## 免责声明

> 【人工智能生成-需人工核验】本技能生成的薪酬体系模板仅供参考，不构成劳动用工、税务或法律建议。薪酬标准、社保公积金比例、个税计算、最低工资标准应以当地最新法规及企业实际制度为准，建议由 HR、财务、法务审核后执行。prajna 企智不对因使用本模板而产生的任何业务决策或制度执行结果承担责任。

## 版本记录

- v1.0.0（2026-07-20）：初始版本，支持生成 7 大薪酬体系工作表，打通 prajna-salary-template 职级带宽与城市系数。
