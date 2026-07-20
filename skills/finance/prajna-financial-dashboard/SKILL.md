---
name: prajna-financial-dashboard
description: 为管理层一键生成「财务核心指标分析看板」HTML 页面，覆盖资产负债率、净利润率、现金流量比率、营业收入、毛利率、ROE、流动比率、经营现金流净额等关键指标的趋势图表与同比环比分析。
version: 1.0.0
author: Prajna 企智
language: zh-CN
category: finance
tags:
  - 财务
  - 看板
  - 核心指标
  - 资产负债率
  - 净利润率
  - 现金流量比率
  - Chart.js
  - HTML
  - 经营决策
---

# prajna-financial-dashboard

## 技能定位

本技能是 Prajna 企智 `finance` 方向原生能力，帮助财务、管理层快速产出「财务核心指标分析看板」。输出为一份自包含、可离线打开的 **HTML 页面**，内置交互式 Chart.js 图表，支持资产负债率、净利润率、现金流量比率、营业收入、毛利率、净资产收益率（年化）、流动比率、经营现金流净额等关键指标的趋势展示与同比环比分析。

适用于：月度经营分析会、董事会/管理层决策、财务健康度监控、银行/投资人汇报、内部风控复盘等场景。

## 触发条件

- 用户说「生成财务看板」「财务核心指标分析」「月度经营决策看板」「财务健康度监控」等
- 关键词命中：财务、看板、资产负债率、净利润率、现金流量比率、同比、环比、Chart.js、HTML
- CLI 调用：`prajna skill run prajna-financial-dashboard`
- 其他 skill 通过协同调用规范触发（见下方）

## 执行流程

1. **确认企业/行业预设**：从内置预设中选择最匹配的企业类型（通用制造集团、互联网/SaaS、零售连锁、新能源科技、通用企业）。
2. **收集关键信息**：企业名称、报表周期、输出路径。
3. **加载 24 个月示例财务数据**：营业收入、毛利润、净利润、总资产、总负债、股东权益、流动资产、流动负债、经营现金流净额。
4. **计算核心指标**：
   - 资产负债率 = 总负债 / 总资产
   - 净利润率 = 净利润 / 营业收入
   - 现金流量比率 = 经营现金流净额 / 流动负债
   - 毛利率、净资产收益率（年化）、流动比率
5. **生成 HTML 看板**：
   - 顶部企业信息与健康评级
   - 8 张核心指标卡片（含环比、同比与阈值状态）
   - 4 个交互图表：资产负债率趋势、营业收入与净利润率、现金流量比率与经营现金流净额、营业收入同比增速
   - 近 12 个月指标明细表（含同比/环比）
   - 经营洞察与免责声明
6. **保存并返回文件路径**。

## CLI 参数

通过 `scripts/generate_financial_dashboard.py` 调用：

| 参数 | 简写 | 类型 | 必填 | 说明 |
|------|------|------|------|------|
| `--preset` | `-p` | string | 否 | 财务预设类型，可选值见 `--list-presets`，默认 `通用企业` |
| `--company` | `-c` | string | 否 | 企业名称，覆盖预设默认值 |
| `--period` | `-r` | string | 否 | 报表周期描述，例如 `2026年6月` |
| `--output` | `-o` | string | 否 | 输出 HTML 文件完整路径，覆盖默认输出目录 |
| `--list-presets` | - | flag | 否 | 列出所有可用预设 |

## 内置预设

| 预设名称 | 企业类型 | 特点 |
|---|---|---|
| 通用制造集团 | 制造业 | 资产重、负债适中、毛利率稳健 |
| 互联网/SaaS企业 | 科技/互联网 | 轻资产、高毛利、低负债、高增长 |
| 零售连锁企业 | 零售/连锁 | 营收规模大、周转快、季节性强 |
| 新能源科技企业 | 新能源/科技 | 高成长、利润率与现金流双优 |
| 通用企业 | 综合 | 中等规模、指标均衡 |

## 使用示例

```bash
# 使用默认通用企业预设
python3 scripts/generate_financial_dashboard.py

# 使用新能源科技企业预设
python3 scripts/generate_financial_dashboard.py --preset "新能源科技企业" --period "2026年6月"

# 自定义企业与周期
python3 scripts/generate_financial_dashboard.py \
  --preset "互联网/SaaS企业" \
  --company "星河科技" \
  --period "2026年第二季度"

# 指定输出路径
python3 scripts/generate_financial_dashboard.py \
  --preset "通用制造集团" \
  --output ~/Desktop/财务核心指标看板_华智制造.html

# 列出所有预设
python3 scripts/generate_financial_dashboard.py --list-presets
```

## 输出约定

- **默认输出目录**：`~/.prajna/prajna-financial-dashboard/samples/`
- **默认文件命名**：`prajna_财务核心指标看板_<企业>_<月份>.html`
- **输出格式**：单文件 HTML（`.html`），离线可用，通过 CDN 加载 Chart.js
- **看板特点**：
  - 深蓝渐变头部，卡片式 KPI 布局，专业大气
  - 指标卡片带环比/同比与阈值状态色
  - 4 组交互式图表，支持鼠标悬停查看明细
  - 近 12 个月明细表，含营收同比/环比
  - 财务健康评级与经营洞察摘要

## 与现有 Skill 的打通说明

### 协同调用规范

本 skill 设计为可被其他 Prajna 原生 skill 调用，统一入口为：

```bash
python3 /Users/a12345/.prajna/skills/finance/prajna-financial-dashboard/scripts/generate_financial_dashboard.py \
  --preset <预设> \
  --company <企业名称> \
  --period <报表周期> \
  --output <输出路径>
```

建议的调用方与场景：

- **`prajna-report-agent`**：在月度/季度经营报告汇总完成后，调用本脚本生成可视化财务看板页面。
- **`prajna-finance`**：在财务数据校验、合并报表完成后，调用本脚本输出管理层监控看板。
- **`prajna-strategy`**：在战略复盘或预算执行分析时，调用本脚本展示核心财务趋势。
- **`prajna-ceo-assistant`**：在高管会议前，自动生成一页式财务健康度看板作为会议材料。

调用方应负责提供 `--company`、`--period` 等参数；本 skill 仅负责数据计算、页面渲染与文件输出。

## 相关文件

- `data/financial_presets.json`：内置企业预设与 24 个月示例财务数据。
- `scripts/generate_financial_dashboard.py`：核心生成脚本，可独立运行。

## 依赖工具

- `python`
- `jinja2`（推荐；未安装时使用内置占位符替换降级）

## 免责声明

本技能生成的 HTML 看板及示例数据仅供管理层经营分析与财务健康度监控参考，不构成投资、融资、信贷或任何商业决策依据。实际财务数据应以企业 ERP/财务系统为准，关键决策建议由财务负责人、审计机构及管理层共同审核确认。Prajna 企智不对因使用本看板而产生的任何业务决策或数据结果承担责任。
