---
name: prajna-bidding-assistant
description: Prajna 企智 — 招投标助手。为销售/投标团队一键生成招投标套件：招标文件拆解清单、投标人资格自审表、技术偏离表、商务报价表、评分响应索引、时间节点与材料清单，以及投标文件 Word 大纲。支持项目信息自定义、行业适配、Excel + Word 双格式输出。
version: 1.0.0
author: Prajna（企智）
language: zh-CN
category: sales
tags:
  - 招投标
  - 投标文件
  - 招标文件拆解
  - 资格自审
  - 技术偏离
  - 商务报价
  - 评分响应
  - Excel
  - Word
---

# prajna-bidding-assistant

## 技能定位

本技能是 Prajna 企智的 **招投标 Agent 原生能力**，面向销售、投标、商务团队，针对具体招标项目快速生成一套结构化、可落地的投标准备材料。

无论是 IT 信息化项目、工程建设、服务外包、货物采购，均可通过本技能完成招标要求拆解、资格自审、技术与商务偏离分析、报价结构设计、评分响应索引以及投标文件大纲输出。

## 触发条件

- “帮我准备一份投标文件”
- “生成招投标套件 / 招标拆解表 / 投标资格自审表”
- “我要投 XX 项目，做一份技术偏离表和报价表”
- `prajna skill run prajna-bidding-assistant --project="智慧园区项目"`

## 执行流程

1. **识别招标项目信息**
   - 从用户输入提取：项目名称、投标人、招标人、投标金额、工期/服务期、行业领域、输出格式等。
   - 若用户未指定，使用通用招投标模板。

2. **生成 Excel 投标准备工作簿（6 个工作表）**
   - **招标文件拆解清单**：招标要求、责任部门、截止时间、重要性。
   - **投标人资格自审表**：资质项、招标要求、我方情况、满足状态、证明文件。
   - **技术偏离表**：序号、招标要求、我方响应、偏离情况、说明。
   - **商务报价表**：分项名称、单位、数量、单价、总价、税率。
   - **评分响应索引**：评分项、分值、我方响应要点、证明材料位置。
   - **时间节点与材料清单**：阶段、事项、责任人、截止日期、交付物、状态。

3. **生成 Word 投标文件大纲**
   - 封面、目录、公司介绍、技术方案、商务方案、报价说明、资质证明、附件清单。
   - 自动填入项目信息，并预留待填写段落。

4. **输出与命名**
   - 默认输出到 `~/.prajna/prajna-bidding-assistant/samples/`。
   - Excel：`prajna_招投标套件_<项目简称>_<日期>.xlsx`
   - Word：`prajna_投标文件大纲_<项目简称>_<日期>.docx`

## CLI 参数

| 参数 | 说明 |
|---|---|
| `--project` | 项目名称（必填） |
| `--bidder` | 投标人/公司名称 |
| `--tenderer` | 招标人/采购人 |
| `--amount` | 投标总金额（元） |
| `--duration` | 工期/服务期，如 "180天" |
| `--industry` | 行业领域，如 IT、建筑、服务、制造 |
| `--output` | 输出目录或文件路径前缀 |
| `--format` | `excel` / `word` / `all`，默认 `all` |

## 使用示例

```bash
# 生成完整招投标套件（Excel + Word）
python3 /Users/a12345/.prajna/skills/sales/prajna-bidding-assistant/scripts/generate_bidding_kit.py \
  --project="智慧园区智能化建设项目" \
  --bidder="智讯科技有限公司" \
  --tenderer="广州高新区管委会" \
  --amount=5800000 \
  --duration="180天" \
  --industry=IT

# 仅生成 Excel
python3 /Users/a12345/.prajna/skills/sales/prajna-bidding-assistant/scripts/generate_bidding_kit.py \
  --project="劳务派遣服务采购" \
  --format=excel

# 指定输出目录
python3 /Users/a12345/.prajna/skills/sales/prajna-bidding-assistant/scripts/generate_bidding_kit.py \
  --project="档案数字化项目" \
  --output=/tmp/bidding_kit
```

## 输出约定

- 默认输出目录：`~/.prajna/prajna-bidding-assistant/samples/`
- Excel 文件包含 6 个工作表，表头清晰、字段完整、带示例数据。
- Word 文件包含 8 个主要章节，可直接作为投标文件结构参考。
- 所有输出均附加 `[人工智能生成-需人工核验]` 免责声明。

## 协同调用规范

本 skill 可与 `prajna-bid-research` 协同使用：

| 调用方 Skill | 调用节点 | 输入 | 输出物 |
|---|---|---|---|
| `prajna-bid-research` | 完成招标信息抓取与评分拆解后 | 项目名称、招标人、预算、评分项 | 完整投标准备套件 |

### 调用约定

1. **参数传递**：调用方负责收集项目基本信息，作为 `--project`、`--tenderer`、`--amount` 等传入脚本。
2. **输出承接**：本 skill 输出 Excel + Word 文件路径，调用方可直接交付用户或嵌入投标项目管理流程。
3. **免责声明**：所有输出必须附加 `[人工智能生成-需人工核验]` 声明，正式投标前建议人工复核。
4. **失败回退**：若本 skill 调用失败，调用方应使用自身默认流程继续，不影响主任务完成。

## 扩展路线

- **v1.1.0**：支持读取招标文件 PDF/Word 自动提取关键信息填充模板。
- **v1.2.0**：接入企业资质库、案例库、人员库，自动匹配资格自审表。
- **v2.0.0**：与 `prajna-bid-research` 深度打通，基于评分项自动生成响应话术与证明材料清单。

## 免责声明

> 【人工智能生成-需人工核验】本技能生成的所有招投标文件、报价方案、资格材料、技术响应内容仅供辅助参考，不构成法律或商务承诺。正式投标文件须由企业投标负责人、法务、财务及业务负责人依据招标文件要求和当地法律法规审核确认。

## 版本记录

- v1.0.0（2026-07-20）：初始版本，支持项目信息自定义、6 工作表 Excel 投标准备套件、Word 投标文件大纲、行业适配与双格式输出。
