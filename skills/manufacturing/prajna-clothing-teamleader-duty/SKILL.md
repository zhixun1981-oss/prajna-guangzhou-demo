---
name: prajna-clothing-teamleader-duty
description: 为服装厂车间小组长一键生成岗位职责与 KPI 考核 Excel 模板，覆盖上游/下游部门、岗位职责、考核指标、量化标准五大模块及 8 项核心 KPI。
version: 1.0.0
author: prajna 企智
language: zh-CN
category: manufacturing
tags:
  - 服装厂
  - 小组长
  - 岗位职责
  - KPI
  - 考核指标
  - Excel
  - 制造业
---

# prajna-clothing-teamleader-duty

## 技能定位

本技能是 prajna 企智 `manufacturing` 方向原生能力，面向服装厂车间小组长岗位，一键生成结构完整的 **岗位职责与 KPI 考核 Excel 工作簿**。

输出包含：
- 服装厂小组长岗位职责矩阵（上游部门、下游部门、岗位职责、考核指标、量化标准）
- 8 项核心 KPI 考核指标（产量达成率、质量合格率、返修率、人员出勤率、员工流失率、设备故障停机时间、安全事故、物料损耗率）
- 绩效评分记录表，支持填写实际完成值后自动计算加权得分

适用于：缝纫组组长、裁剪组组长、包装组组长、车间班组长及生产一线管理人员。

## 触发条件

- 用户说「生成服装厂小组长岗位职责表」「服装厂 KPI 考核表」「车间班组长职责 Excel」
- 关键词命中：服装厂、小组长、班组长、岗位职责、KPI、考核指标、生产执行、质量管控、人员管理
- CLI 调用：`prajna skill run prajna-clothing-teamleader-duty`
- 其他 skill 通过协同调用规范触发（见下方）

## 执行流程

1. **识别班组类型**：从预设中选择最匹配的班组（缝纫组、裁剪组、包装组、通用小组）。
2. **收集关键信息**：工厂名称、班组名称、考核周期、编制人、生成日期。
3. **生成 Excel 工作簿**：
   - **使用说明**：模板用途、填写须知、免责声明。
   - **岗位职责矩阵**：按生产执行、质量管控、人员管理、设备与安全、成本与现场 5 大模块列出岗位职责及上下游关系。
   - **KPI 考核指标**：8 项核心 KPI 的权重、目标值、计算公式、数据来源、评分规则、考核周期。
   - **绩效评分记录**：输入实际完成值后自动计算加权得分，用于月度绩效评定。
4. **保存并返回文件路径**。

## CLI 参数

通过 `scripts/generate_clothing_teamleader_duty.py` 调用：

| 参数 | 简写 | 类型 | 必填 | 说明 |
|------|------|------|------|------|
| `--preset` | `-p` | string | 否 | 班组预设，可选值见 `--list-presets`，默认 `服装缝纫组` |
| `--team` | `-t` | string | 否 | 班组名称，覆盖预设默认值 |
| `--factory` | `-f` | string | 否 | 工厂名称，默认 `成衣A厂` |
| `--month` | `-m` | string | 否 | 考核周期，如 `2026-07`，默认当前月份 |
| `--author` | `-a` | string | 否 | 编制人姓名 |
| `--output` | `-o` | string | 否 | 输出文件完整路径，覆盖默认输出目录 |
| `--list-presets` | - | flag | 否 | 列出所有可用预设 |

## 使用示例

```bash
# 使用默认缝纫组模板
python3 scripts/generate_clothing_teamleader_duty.py

# 使用裁剪组预设
python3 scripts/generate_clothing_teamleader_duty.py --preset "服装裁剪组" --factory "精品女装厂"

# 自定义班组与周期
python3 scripts/generate_clothing_teamleader_duty.py \
  --preset "服装包装组" \
  --team "成品包装B组" \
  --factory "运动服饰厂" \
  --month "2026-08" \
  --author "王芳"

# 指定输出路径
python3 scripts/generate_clothing_teamleader_duty.py \
  --team "缝制一组" \
  --output ~/Desktop/服装厂小组长岗位职责与KPI.xlsx

# 列出所有预设
python3 scripts/generate_clothing_teamleader_duty.py --list-presets
```

## 输出约定

- **默认输出目录**：`~/.prajna/prajna-clothing-teamleader-duty/samples/`
- **默认文件命名**：`prajna_服装厂小组长岗位职责与KPI_<工厂>_<班组>_<周期>.xlsx`
- **输出格式**：Excel（`.xlsx`）
- **模板特点**：
  - 表头采用统一蓝色主题色，首行冻结，列宽自适应。
  - 岗位职责列自动换行，便于打印与分发。
  - KPI 表内置权重合计校验（合计 100%）。
  - 绩效评分记录表内置加权得分公式，输入实际完成值后自动汇总。

## 与现有 Skill 的打通说明

### 协同调用规范

本 skill 设计为可被其他 prajna 原生 skill 调用，统一入口为：

```bash
python3 /Users/a12345/.prajna/skills/manufacturing/prajna-clothing-teamleader-duty/scripts/generate_clothing_teamleader_duty.py \
  --preset <预设> \
  --team <班组名> \
  --factory <工厂名> \
  --month <考核周期> \
  --author <编制人> \
  --output <输出路径>
```

建议的调用方与场景：

- **`prajna-manufacturing-hr`**：在车间岗位编制、绩效考核方案设计完成后，调用本脚本输出岗位职责与 KPI 模板。
- **`prajna-production-planning`**：在排产计划确定后，调用本脚本将产量、质量、物料等指标固化到班组考核表。
- **`prajna-quality-control`**：在质量异常分析后，调用本脚本更新返修率、合格率等质量 KPI。
- **`prajna-training`**：在新员工培训计划制定后，调用本脚本补充人员培训与考核指标。

调用方应负责提供 `--team`、`--factory`、`--month`、`--author` 等参数；本 skill 仅负责模板生成与文件输出。

## 相关文件

- `scripts/generate_clothing_teamleader_duty.py`：核心生成脚本，可独立运行。

## 依赖工具

- `python`
- `openpyxl`

## 免责声明

> 【人工智能生成-需人工核验】本技能生成的岗位职责、KPI 权重、目标值、评分规则及量化标准仅供企业管理参考，不构成正式劳动用工、绩效考核或薪酬核算依据。用户应结合企业实际情况、行业规范及当地法律法规进行调整，并由 HR、生产管理及相关管理层审核确认。prajna 企智不对因使用本模板而产生的任何业务决策或数据结果承担责任。
