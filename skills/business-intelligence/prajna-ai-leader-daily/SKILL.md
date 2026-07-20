---
name: prajna-ai-leader-daily
description: 每日追踪李开复、Sam Altman/OpenAI、Anthropic、Google DeepMind、Meta 等全球 AI 领袖的最新技术与战略动态，生成卡片式、响应式、适合邮件阅读的 HTML 日报。
version: 1.0.0
author: Prajna 企智
language: zh-CN
category: business-intelligence
tags:
  - AI领袖
  - 日报
  - HTML
  - 邮件
  - 行业动态
  - 竞争情报
---

# prajna-ai-leader-daily

## 技能定位

本技能是 Prajna 企智 `business-intelligence` 方向原生能力，面向关注全球 AI 产业趋势的高管、投资人、产品经理与研究员，一键生成「AI 领袖动态日报」HTML 页面。输出采用卡片式布局、响应式设计、邮件客户端友好，可直接嵌入企业邮件、内部 Wiki 或移动端阅读。

## 触发条件

- 用户说「生成 AI 领袖日报」「今日 AI 领袖动态」「AI 行业日报」
- 关键词命中：AI 领袖、李开复、OpenAI、Anthropic、DeepMind、Meta、日报、动态
- CLI 调用：`prajna skill run prajna-ai-leader-daily`
- 其他 skill 通过协同调用规范触发（见下方）

## 执行流程

1. **确认日报范围**：默认覆盖李开复/零一万物、Sam Altman/OpenAI、Anthropic、Google DeepMind、Meta。
2. **获取最新动态**：调用 WebSearch 等联网工具检索上述领袖的最新技术与战略新闻。
3. **精选与摘要**：每家精选 1-3 条高价值动态，撰写 50-80 字中文摘要，标注来源链接。
4. **生成 HTML 日报**：使用 Jinja2 模板渲染响应式卡片页面，适配桌面与邮件阅读。
5. **保存示例文件**：默认输出到 `~/.prajna/prajna-ai-leader-daily/samples/`，文件名含日期。
6. **返回文件路径与摘要**：向用户展示生成结果与核心内容摘要。

## CLI 参数

通过 `scripts/generate_ai_leader_daily.py` 调用：

| 参数 | 简写 | 类型 | 必填 | 说明 |
|------|------|------|------|------|
| `--date` | `-d` | string | 否 | 日报日期，格式 `YYYY-MM-DD`，默认当天 |
| `--output` | `-o` | string | 否 | 输出文件完整路径，覆盖默认输出目录 |
| `--data` | - | string | 否 | 自定义新闻数据 JSON 路径 |
| `--list-leaders` | - | flag | 否 | 列出默认覆盖的 AI 领袖 |

## 使用示例

```bash
# 生成今日 AI 领袖日报
python3 /Users/a12345/.prajna/skills/business-intelligence/prajna-ai-leader-daily/scripts/generate_ai_leader_daily.py

# 指定日期
python3 /Users/a12345/.prajna/skills/business-intelligence/prajna-ai-leader-daily/scripts/generate_ai_leader_daily.py \
  --date 2026-07-20

# 指定输出路径
python3 /Users/a12345/.prajna/skills/business-intelligence/prajna-ai-leader-daily/scripts/generate_ai_leader_daily.py \
  --output ~/Desktop/AI领袖日报_2026-07-20.html

# 使用自定义数据
python3 /Users/a12345/.prajna/skills/business-intelligence/prajna-ai-leader-daily/scripts/generate_ai_leader_daily.py \
  --data data/ai_leader_curated.json

# 列出默认覆盖的 AI 领袖
python3 /Users/a12345/.prajna/skills/business-intelligence/prajna-ai-leader-daily/scripts/generate_ai_leader_daily.py \
  --list-leaders
```

## 输出约定

- **默认输出目录**：`~/.prajna/prajna-ai-leader-daily/samples/`
- **默认文件命名**：`prajna_ai_leader_daily_<YYYY-MM-DD>.html`
- **输出格式**：HTML（内嵌 CSS，响应式，邮件友好）
- **内容约定**：
  - 每家领袖 1-3 条动态，摘要 50-80 字。
  - 每条动态标注来源名称与可点击链接。
  - 不修改已有 skill，仅创建新目录与示例。

## 与现有 Skill 的打通说明

### 协同调用规范

本 skill 可被其他 Prajna 原生 skill 调用，统一入口为：

```bash
python3 /Users/a12345/.prajna/skills/business-intelligence/prajna-ai-leader-daily/scripts/generate_ai_leader_daily.py \
  --date <YYYY-MM-DD> \
  --output <输出路径> \
  --data <自定义数据JSON>
```

建议调用方与场景：

- **`prajna-industry-intelligence`**：在行业研究完成后生成 AI 领袖动态日报，作为竞争情报附件。
- **`prajna-report-agent`**：在晨报/晚报中嵌入本日报，或作为邮件正文/附件发送给订阅者。
- **`prajna-strategy`**：在战略规划会议前调用，提供外部 AI 产业风向标。
- **`prajna-marketing-content`**：基于日报热点生成社交媒体推文、公众号选题或内部简报。

调用方负责提供 `--date`、`--output` 等参数；本 skill 负责模板渲染与文件输出。

## 相关文件

- `data/ai_leader_curated.json`：默认精选的 AI 领袖动态样本数据（含来源链接）。
- `scripts/generate_ai_leader_daily.py`：核心生成脚本，可独立运行。

## 依赖工具

- `python`
- `jinja2`（可选，缺失时使用内置 `string.Template` 降级渲染）

## 免责声明

本技能生成的日报内容基于公开网络搜索与示例数据，仅供信息参考，不构成投资建议或商业决策依据。新闻摘要与来源链接可能随时间变化，用户应自行核实原始报道。Prajna 企智不对因使用本日报而产生的任何决策承担责任。
