---
name: prajna-recruitment-assistant
description: Prajna 企智招聘助手原生技能。一键生成岗位说明书（JD）、任职资格与胜任力模型、结构化面试评估表、招聘漏斗与周报、人才画像卡、Offer 薪资建议等 6 张 Excel 工作表，以及完整招聘套件 Word 文档。与城市薪酬系数、职级带宽、薪资模板原生打通。
version: 1.0.0
author: Prajna 企智
language: zh-CN
category: hr
tags:
  - 招聘
  - 岗位说明书
  - JD
  - 面试评估
  - 招聘漏斗
  - 人才画像
  - Offer
  - 薪酬带宽
  - Excel
  - Word
  - HR
---

# prajna-recruitment-assistant

## 技能定位

本技能是 Prajna 企智 `hr` 方向原生能力，面向 HR、用人部门及招聘负责人，一键生成结构完整的 **招聘套件**。

输出包含：
- 岗位说明书（JD）Excel 工作表
- 任职资格与胜任力模型
- 结构化面试评估表（含评分维度、权重、评分标准）
- 招聘漏斗与周报
- 人才画像卡
- Offer 薪资建议（与城市薪酬系数、职级带宽一致）
- 完整招聘套件 Word 文档（JD + 面试提纲 + Offer 模板）

适用于：社会招聘、校园招聘、内部竞聘、岗位新增、编制扩张等场景。

## 触发条件

- 用户说「帮我写一份 JD」「生成招聘套件」「做一份面试评估表」
- 关键词命中：招聘、JD、岗位说明书、面试评估、招聘漏斗、人才画像、Offer、薪资建议
- CLI 调用：`prajna skill run prajna-recruitment-assistant`
- 其他 skill 通过协同调用规范触发（见下方）

## 执行流程

1. **识别岗位信息**：从命令行参数提取岗位、部门、城市、职级、薪资范围、汇报对象、编制、紧急程度等。
2. **加载薪酬基准**：读取内置城市薪酬系数与职级带宽，与 `prajna-salary-template`、`prajna-compensation-system` 保持一致。
3. **生成 Excel 工作簿**：
   - **岗位说明书（JD）**：岗位基本信息、岗位目的、核心职责、工作关系、绩效目标。
   - **任职资格与胜任力模型**：学历、经验、专业知识、技能要求、胜任力维度及行为描述。
   - **结构化面试评估表**：评分维度、权重、评分标准、面试官打分、加权得分、建议结论。
   - **招聘漏斗与周报**：渠道、投递数、初筛通过、初试、复试、offer、入职，转化率与阶段耗时。
   - **人才画像卡**：理想候选人特征、关键经历、核心能力、风险信号、吸引点。
   - **Offer 薪资建议**：城市、职级、薪资带宽、固定工资、绩效工资、补贴、年终奖、总包、市场竞争力评估。
4. **生成 Word 文档**：整合 JD、面试提纲、Offer 模板，便于打印、邮件发送或归档。
5. **保存并返回文件路径**。

## CLI 参数

通过 `scripts/generate_recruitment_kit.py` 调用：

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `--position` | string | 是 | 岗位名称，如「电商运营助理」 |
| `--department` | string | 否 | 所属部门，默认「人力资源部」 |
| `--city` | string | 否 | 工作城市，默认「广州」 |
| `--level` | string | 否 | 职级，如 P1-P9，默认 P2 |
| `--salary-min` | number | 否 | 月薪下限（元），默认按职级带宽 |
| `--salary-max` | number | 否 | 月薪上限（元），默认按职级带宽 |
| `--reports-to` | string | 否 | 汇报对象，默认「部门经理」 |
| `--headcount` | int | 否 | 招聘人数，默认 1 |
| `--urgency` | string | 否 | 紧急程度：高/中/低，默认「中」 |
| `--output` | string | 否 | 输出目录或文件路径，默认 `~/.prajna/prajna-recruitment-assistant/samples/` |
| `--format` | string | 否 | 输出格式：`excel`/`word`/`all`，默认 `all` |

## 使用示例

```bash
# 使用默认参数生成招聘套件
python3 ~/.prajna/skills/hr/prajna-recruitment-assistant/scripts/generate_recruitment_kit.py \
  --position="电商运营助理" --department="电商运营部" --city=广州 --level=P2

# 完整参数示例
python3 ~/.prajna/skills/hr/prajna-recruitment-assistant/scripts/generate_recruitment_kit.py \
  --position="高级产品经理" \
  --department="产品部" \
  --city=深圳 \
  --level=P4 \
  --salary-min=18000 \
  --salary-max=25000 \
  --reports-to="产品总监" \
  --headcount=1 \
  --urgency=高 \
  --format=all

# 只输出 Excel
python3 ~/.prajna/skills/hr/prajna-recruitment-assistant/scripts/generate_recruitment_kit.py \
  --position="软件工程师" --city=北京 --level=P3 --format=excel
```

## 输出约定

- **默认输出目录**：`~/.prajna/prajna-recruitment-assistant/samples/`
- **Excel 文件命名**：`prajna_招聘套件_<岗位>_<城市>_<日期>.xlsx`
- **Word 文件命名**：`prajna_招聘套件_<岗位>_<城市>_<日期>.docx`
- **输出格式**：Excel（6 张工作表）+ Word（完整招聘套件）
- **模板特点**：
  - 表头采用统一蓝色主题色，首行冻结，列宽自适应。
  - 面试评估表内置权重合计校验（合计 100%）与加权得分公式。
  - 招聘漏斗表内置转化率公式，输入实际数据后自动汇总。
  - Offer 薪资建议与城市薪酬系数、职级带宽联动。

## 与现有 Skill 的打通说明

### 协同调用规范

本 skill 设计为可被其他 Prajna 原生 skill 调用，统一入口为：

```bash
python3 /Users/a12345/.prajna/skills/hr/prajna-recruitment-assistant/scripts/generate_recruitment_kit.py \
  --position <岗位> \
  --department <部门> \
  --city <城市> \
  --level <职级> \
  --salary-min <月薪下限> \
  --salary-max <月薪上限> \
  --reports-to <汇报对象> \
  --headcount <招聘人数> \
  --urgency <紧急程度> \
  --output <输出路径> \
  --format <输出格式>
```

建议的调用方与场景：

- **`prajna-salary-template`**：在 Offer 薪资建议阶段，调用本脚本生成岗位招聘套件，再输出薪酬结构参考。
- **`prajna-compensation-system`**：薪酬体系设计完成后，调用本脚本校验岗位薪资带宽与市场对标。
- **`prajna-performance-system`**：绩效指标确定后，调用本脚本将 KPI 写入岗位说明书与面试评估表。
- **`prajna-hr-onboarding`**：候选人入职前，调用本脚本归档招聘材料。

### 数据打通

- **城市薪酬系数**：与 `prajna-salary-template/data/city_salary_multipliers.json` 保持一致。
- **职级带宽**：与 `prajna-salary-template/data/job_level_bandwidth.json` 保持一致。
- **岗位序列**：Offer 薪资建议中的岗位序列可与薪酬体系岗位序列对应。
- **KPI 思路**：绩效相关的胜任力维度可复用 `prajna-clothing-teamleader-duty` 及电商运营岗位的 KPI 设计思路。

## 相关文件

- `scripts/generate_recruitment_kit.py`：核心生成脚本，可独立运行。
- `data/recruitment_defaults.json`：默认参数与城市/职级基准。

## 依赖工具

- `python3`
- `openpyxl`
- `python-docx`

## 免责声明

> 【人工智能生成-需人工核验】本技能生成的岗位说明书、任职资格、面试评估表、招聘漏斗、人才画像、Offer 薪资建议仅供企业招聘管理参考，不构成正式劳动用工、薪酬决策或法律文件。最终招聘方案、录用决定及薪酬待遇须由企业 HR、用人部门及管理层依据当地法律法规和公司政策审核确认。Prajna 企智不对因使用本模板而产生的任何业务决策或数据结果承担责任。

## 版本记录

- v1.0.0（2026-07-20）：初始版本。支持岗位说明书、任职资格与胜任力模型、结构化面试评估表、招聘漏斗与周报、人才画像卡、Offer 薪资建议 6 张 Excel 工作表，以及完整招聘套件 Word 文档。与薪资模板、薪酬体系、绩效体系 Skill 完成数据打通。
