---
name: prajna-performance-system
description: Prajna 企智 — 绩效体系助手。为企业一键生成绩效管理体系套件，包括 KPI 指标库、绩效合同/目标责任书、绩效评分表、绩效面谈记录、绩效结果分布、绩效改进计划 PIP，以及绩效管理制度 Word 文档。
version: 1.0.0
author: Prajna 企智
language: zh-CN
category: hr
tags:
  - 绩效体系
  - KPI
  - OKR
  - 360评估
  - 绩效合同
  - 绩效面谈
  - PIP
  - 绩效管理
  - Excel
  - Word
  - HR
---

# prajna-performance-system

## 技能定位

本技能是 Prajna 企智 **人力资源方向原生能力**，面向企业 HR、部门负责人与管理者，一键生成结构完整、可直接落地使用的绩效管理体系套件。

无论是 KPI、OKR、360 还是 MBO 考核方法，也无论是月度、季度还是年度考核周期，均可通过本技能快速生成对应的 KPI 指标库、绩效合同、评分表、面谈记录、结果分布与 PIP 模板，并配套输出一份完整的《绩效管理制度》Word 文档。

## 触发条件

- “生成绩效管理制度 / 绩效体系”
- “帮我做一份 XX 岗位的 KPI 考核表”
- “绩效合同 / 绩效评分表 / 绩效面谈表 / PIP 模板”
- “季度绩效考核方案 / 年度绩效管理方案”
- CLI 调用：`prajna skill run prajna-performance-system`

## 执行流程

1. **识别岗位与部门**
   - 从用户输入提取：企业名称、部门、岗位、考核周期、考核方法、绩效等级。
   - 若用户未指定，使用默认示例（电商运营部 / 电商运营助理 / 季度 / KPI / A/B/C/D）。

2. **加载 KPI 指标库**
   - 内置多部门/多岗位 KPI 模板：电商运营、生产制造、服装厂小组长、销售、人力资源、财务、研发技术、产品、通用岗位等。
   - KPI 设计复用 `prajna-salary-template`、`prajna-clothing-teamleader-duty` 等现有技能的岗位 KPI 思路，确保指标一致性与 Skill 间打通。

3. **生成 Excel 工作簿（6 个工作表）**
   - **KPI 指标库**：按部门/岗位分类的 KPI 编号、名称、权重、目标值、计算方式、数据来源、评分规则、考核周期。
   - **绩效合同/目标责任书**：员工与上级共同确认的绩效目标、等级与结果应用、双方签字栏。
   - **绩效评分表**：自评、上级评、HR 复核三栏评分，自动计算最终得分与加权得分。
   - **绩效面谈记录**：绩效回顾、上级反馈、发展需求、下期目标、行动计划、申诉备注。
   - **绩效结果分布**：绩效等级、分数区间、建议比例、结果应用、团队分布示例。
   - **绩效改进计划 PIP**：待改进项、行动计划、辅导资源、结果约定、签字栏。

4. **生成 Word 管理制度**
   - 输出完整的《绩效管理制度》，包含目的、适用范围、考核原则、考核周期、考核方法、考核流程、绩效等级与结果应用、绩效申诉、附则及免责声明。

5. **保存示例文件**
   - 默认输出到 `~/.prajna/prajna-performance-system/samples/`，文件名含企业、部门、岗位与日期。

## CLI 参数

通过 `scripts/generate_performance_system.py` 调用：

| 参数 | 简写 | 类型 | 必填 | 说明 |
|------|------|------|------|------|
| `--department` | `-d` | string | 否 | 考核部门名称，默认 `电商运营部` |
| `--position` | `-p` | string | 否 | 考核岗位名称，默认 `电商运营助理` |
| `--cycle` | `-c` | string | 否 | 考核周期，如 `月度`、`季度`、`半年度`、`年度`，默认 `季度` |
| `--method` | `-m` | string | 否 | 考核方法，如 `KPI`、`OKR`、`360`、`MBO`，默认 `KPI` |
| `--levels` | `-l` | string | 否 | 绩效等级划分，如 `A/B/C/D`、`S/A/B/C/D`，默认 `A/B/C/D` |
| `--company` | `-co` | string | 否 | 企业名称，默认 `Prajna示范企业` |
| `--output` | `-o` | string | 否 | 输出目录或完整路径，覆盖默认输出目录 |
| `--format` | `-f` | string | 否 | 输出格式：`excel`、`word`、`all`，默认 `all` |

## 使用示例

```bash
# 生成默认示例（电商运营助理 / 季度 / KPI）
python3 /Users/a12345/.prajna/skills/hr/prajna-performance-system/scripts/generate_performance_system.py

# 指定岗位与周期
python3 /Users/a12345/.prajna/skills/hr/prajna-performance-system/scripts/generate_performance_system.py \
  --department="电商运营部" \
  --position="电商运营助理" \
  --cycle=季度 \
  --method=KPI

# 制造业生产主管，年度考核
python3 /Users/a12345/.prajna/skills/hr/prajna-performance-system/scripts/generate_performance_system.py \
  --department="生产制造部" \
  --position="生产主管" \
  --cycle=年度 \
  --method=KPI \
  --company="智云制造"

# 仅生成 Excel
python3 /Users/a12345/.prajna/skills/hr/prajna-performance-system/scripts/generate_performance_system.py \
  --position="软件工程师" \
  --format=excel

# 自定义输出路径
python3 /Users/a12345/.prajna/skills/hr/prajna-performance-system/scripts/generate_performance_system.py \
  --output ~/Desktop/绩效体系示例 \
  --company="智云科技"
```

## 输出约定

- **默认输出目录**：`~/.prajna/prajna-performance-system/samples/`
- **默认文件命名**：`prajna_绩效体系_<企业>_<部门>_<岗位>_<日期>.xlsx/.docx`
- **输出格式**：Excel（6 个工作表）+ Word（绩效管理制度）
- **所有涉及法律/合规/敏感决策的输出，需附加免责声明并建议人工复核**

## 与现有 Skill 的打通说明

本 skill 作为绩效体系中台，与以下 Prajna 原生 skill 在关键节点打通：

| 调用方 Skill | 调用节点 | 输入 | 输出物 |
|---|---|---|---|
| `prajna-salary-template` | 绩效方案设计完成后、奖金方案配套时 | 岗位、KPI 维度与权重 | 含绩效考核与奖金方案的完整薪资模板 |
| `prajna-compensation-system` | 薪酬体系搭建、绩效结果与薪酬联动设计时 | 岗位序列、绩效等级 | 薪酬-绩效联动规则、调薪方案 |
| `prajna-recruitment-package` | JD 生成后、试用期考核时 | 目标岗位 | KPI 考核指标库、绩效合同模板 |
| `prajna-hr-methodologies` | 绩效面谈、公平性审查、人才盘点时 | 绩效结果、员工数据 | 面谈提纲、公平性审查报告、九宫格分布 |
| `prajna-report-agent` | 人力成本与绩效分析报表时 | 部门/岗位范围 | 绩效结果分布、KPI 达成率分析 |

### KPI 复用说明

- `电商运营助理`、`软件工程师`、`产品经理` 等岗位的 KPI 设计，复用自 `prajna-salary-template` 的岗位 KPI 模板，确保绩效指标与薪资结构中的绩效考核维度一致。
- `服装厂小组长` 的 KPI 设计，复用自 `prajna-clothing-teamleader-duty` 的岗位职责与 KPI 考核思路，包含产量达成率、质量合格率、返修率、人员出勤率等核心制造指标。

### 调用约定

1. **参数传递**：调用方负责收集企业名称、部门、岗位、考核周期、考核方法等变量，作为 `--department`/`--position`/`--cycle`/`--method` 传入脚本。
2. **输出承接**：本 skill 输出 Excel 与 Word 文件路径，调用方可直接交付用户或进一步嵌入报告。
3. **免责声明**：所有输出必须附加 `[人工智能生成-需人工核验]` 声明，绩效管理制度与考核结果建议人工复核。
4. **失败回退**：若本 skill 调用失败，调用方应使用自身默认流程继续，不影响主任务完成。

## 内置岗位 KPI 模板

| 部门 | 岗位示例 |
|---|---|
| 电商运营部 | 电商运营助理、电商运营经理 |
| 生产制造部 | 服装厂小组长、生产主管 |
| 销售部 | 销售代表、销售经理 |
| 人力资源部 | 招聘专员、HRBP |
| 财务部 | 财务会计 |
| 研发技术部 | 软件工程师、产品经理 |
| 通用 | 通用岗位（默认回退） |

## 扩展路线

- **v1.1.0**：支持更多行业/岗位 KPI 模板（客服、物流、设计、市场等）。
- **v1.2.0**：支持用户上传现有绩效方案，自动识别字段并填充到模板。
- **v2.0.0**：与 `prajna-salary-template`、`prajna-compensation-system` 深度打通，基于绩效结果自动生成奖金方案与调薪建议。

## 免责声明

> 【人工智能生成-需人工核验】本技能生成的所有绩效管理制度、KPI 指标、绩效合同、评分表、面谈记录、PIP 模板及结果应用建议仅供辅助参考，不构成法律或劳动用工建议。最终绩效方案须由企业 HR、法务及管理层依据当地法律法规和公司政策审核确认。

## 版本记录

- v1.0.0（2026-07-20）：初始版本。支持 KPI/OKR/360/MBO 多种考核方法，生成 6 工作表 Excel + 绩效管理制度 Word，内置多部门/岗位 KPI 模板，与薪资、薪酬、招聘、HR 方法论、报表 Skill 完成调用打通。
