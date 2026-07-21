---
name: prajna-contract-review-assistant
description: 一键生成合同审查意见书 Word 文档，覆盖主体资格、商务条款、付款财务、违约责任、知识产权、争议解决六大维度，附风险登记与修改建议表。
version: 1.0.0
author: Prajna 企智
language: zh-CN
category: legal
tags:
  - 法务
  - 合同审查
  - 风险控制
  - Word
---

# prajna-contract-review-assistant

## 技能定位

本技能是 Prajna 企智 `legal` 方向原生能力，面向法务、合规、商务人员，一键生成 **合同审查意见书**。

输出包含：
- 审查结论
- 主体资格、商务条款、付款财务、违约解除、知识产权与保密、争议解决六大审查维度
- 风险登记与修改建议表
- 最终签署建议

## CLI 参数

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `--company` | string | 否 | 公司名称，默认「智信科技」 |
| `--contract-name` | string | 否 | 合同名称 |
| `--contract-type` | string | 否 | 合同类型，默认「采购合同」 |
| `--amount` | float | 否 | 合同金额（元） |
| `--term` | string | 否 | 合作期限 |
| `--risk-level` | string | 否 | 风险等级：高/中/低 |
| `--reviewer` | string | 否 | 审查人 |
| `--date` | string | 否 | 审查日期 |
| `--output` | string | 否 | 输出文件路径 |

## 使用示例

```bash
python3 skills/legal/prajna-contract-review-assistant/scripts/generate_contract_review.py \
  --company 京东集团 --contract-name 仓储服务合同 --amount 1200000 --term 24个月
```

## 免责声明

> 【人工智能生成-需人工核验】本技能生成的合同审查意见仅供参考，不构成正式法律意见或律师建议。重大合同应由执业律师或企业法务负责人审核确认。
