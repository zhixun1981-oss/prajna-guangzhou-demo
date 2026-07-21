---
name: prajna-customer-service-sop
description: 一键生成客服标准作业程序 Excel 套件，覆盖话术库、工单处理流程、客诉分类与升级规则、FAQ，帮助企业建立标准化客服体系。
version: 1.0.0
author: Prajna 企智
language: zh-CN
category: customer-service
tags:
  - 客服
  - SOP
  - 话术库
  - 工单
  - 客诉升级
  - Excel
---

# prajna-customer-service-sop

## 技能定位

本技能是 Prajna 企智 `customer-service` 方向原生能力，面向客服中心、运营团队，一键生成 **客服 SOP Excel 套件**。

输出包含：
- 客服话术库（场景、情绪、话术、禁止用语）
- 工单处理流程（6 步标准流程与 SLA）
- 客诉分类与升级规则
- 常见问题 FAQ

## CLI 参数

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `--company` | string | 否 | 公司名称，默认「智服电商」 |
| `--date` | string | 否 | 更新日期，默认当天 |
| `--output` | string | 否 | 输出文件路径 |

## 使用示例

```bash
python3 skills/customer-service/prajna-customer-service-sop/scripts/generate_cs_sop.py \
  --company 京东集团
```

## 免责声明

> 【人工智能生成-需人工核验】本技能生成的话术、流程、规则仅供参考，企业应结合自身业务特点、品牌调性与合规要求调整完善。
