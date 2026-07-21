---
name: prajna-procurement-assistant
description: 一键生成采购管理 Excel 套件，覆盖采购申请单、供应商评估表、询价比价单、采购合同审查清单、采购台账，帮助企业规范采购流程。
version: 1.0.0
author: Prajna 企智
language: zh-CN
category: procurement
tags:
  - 采购
  - 供应商管理
  - 询价比价
  - 合同审查
  - Excel
---

# prajna-procurement-assistant

## 技能定位

本技能是 Prajna 企智 `procurement` 方向原生能力，面向采购部门、供应链管理人员，一键生成结构完整的 **采购管理 Excel 套件**。

输出包含：
- 采购申请单
- 供应商评估表（含综合得分与评级公式）
- 询价比价单（含推荐供应商与节省金额公式）
- 采购合同审查清单
- 采购台账（含总价公式与状态跟踪）

## CLI 参数

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `--company` | string | 否 | 公司名称，默认「智采科技」 |
| `--department` | string | 否 | 申请部门，默认「采购部」 |
| `--applicant` | string | 否 | 申请人，默认「张采购」 |
| `--date` | string | 否 | 申请日期，默认当天 |
| `--delivery-date` | string | 否 | 期望到货日期，默认当天 |
| `--output` | string | 否 | 输出文件路径 |

## 使用示例

```bash
python3 skills/procurement/prajna-procurement-assistant/scripts/generate_procurement_kit.py \
  --company 美的集团 --department 采购部 --applicant 张采购
```

## 免责声明

> 【人工智能生成-需人工核验】本技能生成的采购模板、供应商评分、比价结果仅供参考，不构成正式采购决策或合同审查结论。最终方案由企业采购、法务、财务部门审核确认。
