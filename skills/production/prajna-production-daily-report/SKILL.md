---
name: prajna-production-daily-report
description: 一键生成生产日报 Excel 套件，覆盖生产日报、产量统计、设备运行、质量检验、人员出勤五大工作表，支持完成率、OEE、合格率公式联动。
version: 1.0.0
author: prajna 企智
language: zh-CN
category: production
tags:
  - 生产
  - 日报
  - 设备 OEE
  - 质量检验
  - 人员出勤
  - Excel
---

# prajna-production-daily-report

## 技能定位

本技能是 prajna 企智 `production` 方向原生能力，面向生产制造企业的车间管理、生产计划、质量管理，一键生成 **生产日报 Excel 套件**。

输出包含：
- 生产日报（计划/实际/完成率/差异原因）
- 产量统计（近 7 日趋势）
- 设备运行（运行时长、故障时长、OEE）
- 质量检验（合格率、不良项、处理措施）
- 人员出勤（出勤率、缺勤原因、加班人数）

## CLI 参数

| 参数 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `--factory` | string | 否 | 工厂名称，默认「智造工厂」 |
| `--date` | string | 否 | 日报日期，默认当天 |
| `--output` | string | 否 | 输出文件路径 |

## 使用示例

```bash
python3 skills/production/prajna-production-daily-report/scripts/generate_production_daily.py \
  --factory 美的集团武汉工厂 --date 2026-07-21
```

## 免责声明

> 【人工智能生成-需人工核验】本技能生成的生产数据、OEE、合格率等仅为示例，企业应接入真实 MES/ERP 数据进行替换与校验。
