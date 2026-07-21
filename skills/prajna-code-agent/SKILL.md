# prajna-code-agent

## 定位
prajna 的**企业业务代码智能体**。接收自然语言任务，生成可执行 Python 代码，在沙箱中运行，返回结果文件，并把代码与执行记录沉淀到记忆。

## 核心能力
- **业务语义识别**：自动判断任务属于销售分析、薪酬核算、简历筛选、库存分析、KPI 评分等场景
- **代码生成**：基于业务模板生成 Python + pandas + openpyxl 代码，公式与逻辑可二次编辑
- **沙箱执行**：在隔离临时目录中运行代码，带超时保护
- **结果沉淀**：每次任务自动生成代码、输出文件、执行日志，存入 `~/.prajna/code_agent_memory`
- **LLM-ready**：当前为规则模板模式，后续可一键切换为 DeepSeek/Claude/GPT 生成代码

## 使用方式

### CLI
```bash
python3 skills/prajna-code-agent/scripts/code_agent.py \
  "读取 sales_data.xlsx，按区域统计销售额和完成率，生成 HTML 报告" \
  --input sales_data.xlsx
```

### Python
```python
from skills.prajna_code_agent.scripts.code_agent import CodeAgent
agent = CodeAgent()
result = agent.run("核算本月员工工资，输出 payroll_result.xlsx", input_files=["payroll_data.xlsx"])
```

## 当前支持场景
| 意图 | 触发关键词 | 输出 |
|---|---|---|
| sales_analysis | 销售、业绩、完成率、环比、图表 | HTML 报告（含 Chart.js 图表） |
| payroll_calculation | 工资、薪酬、个税、社保、公积金 | Excel 薪酬核算表 |
| resume_screening | 简历、招聘、筛选、候选人 | Excel 筛选结果 |
| inventory_analysis | 库存、周转、滞销、SKU | Excel 库存分析 |
| kpi_scoring | 绩效、KPI、考核、评分 | Excel 绩效评分 |

## 记忆结构
```
~/.prajna/code_agent_memory/
└── code_sales_analysis_20260721_143022/
    ├── metadata.json      # 任务元数据
    ├── generated_script.py # 生成的源代码
    └── sales_report.html   # 输出产物
```

## 扩展计划
1. 接入 LLM 后端，支持开放式代码生成
2. 增加更多垂直场景：生产排程、物流路径、财务对账、合同条款提取
3. 与现有 15 个 Skill 联动：招聘结果 → 薪酬核算 → 绩效评分
4. 接入记忆核心：自动复用历史数据格式与业务口径
