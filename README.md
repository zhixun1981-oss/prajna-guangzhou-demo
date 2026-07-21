# prajna 企业智能体 · 多场景模板中台 + 代码智能体

本项目是 **prajna** 的企业级智能体能力演示：既是一个覆盖 HR、销售、财务、生产、供应链、法务、客服、招投标、情报等多业务场景的**模板中台**，也是一个能**理解业务语义、生成 Python 代码、沙箱执行并沉淀记忆**的代码智能体。用户只需输入参数或一句话自然语言描述，prajna 即可自动识别意图、生成文档或执行数据分析任务。

## 在线 Demo

**https://prajna-guangzhou-demo.onrender.com**

## 核心能力

**🤖 自然语言智能体模式**：输入一句话，prajna 自动识别意图、匹配技能、填充参数并生成文档。

例如：
- "帮我做一份深圳互联网产品经理 P5 的薪资模板"
- "生成电商销售团队本周周报，目标 80 万"
- "搭建智云电商财务知识库目录"
- "帮我做一份智慧园区建设项目的投标书"
- "生成招聘电商运营助理的 JD 和面试评估表"

## 支持的 16 个企业场景

| 业务域 | 技能 | 输出 | 说明 |
|---|---|---|---|
| **智能体中台** | 🤖 代码智能体 | Python + Excel/HTML | 自然语言生成 Python 代码，沙箱执行，自动沉淀代码与结果到记忆 |
| **人力资源** | 💰 薪资模板 | Excel（8 表） | 薪资结构、绩效考核、调薪机制、奖金方案、福利明细、计算示例、合规校验、绩效系数对照 |
| **人力资源** | 🤝 招聘助手 | Excel + Word | 岗位说明书 JD、任职资格、结构化面试评估表、招聘漏斗、人才画像、Offer 模板 |
| **人力资源** | 💵 薪酬体系 | Excel（8 表） | 岗位薪酬矩阵、薪酬结构、城市系数、调薪方案、薪酬预算、人力成本测算、薪酬绩效联动 |
| **人力资源** | ⭐ 绩效体系 | Excel + Word | KPI 指标库、绩效合同、评分表、绩效面谈、结果分布、PIP 改进计划、绩效管理制度 |
| **销售商务** | 📊 销售团队标准周报 | Excel（5 表） | 封面汇总、核心业绩数据、重点商机进展、问题分析、下周工作计划 |
| **销售商务** | 🎯 招投标助手 | Excel + Word | 招标文件拆解、资格自审、技术偏离、商务报价、评分响应索引、投标文件大纲 |
| **财务经营** | 📁 财务知识库目录 | Markdown | 核算规范、成本利润、库存管理、资金税务、经营分析、内控审计、工具模板七大模块 |
| **财务经营** | 📈 财务核心指标看板 | HTML | 资产负债率、净利润率、现金流量比率等 KPI 卡片 + Chart.js 趋势图 + 同比环比 |
| **财务经营** | 📑 预算执行汇报 PPT | PPTX | 预算完成率对比、超支分析、结余建议、下月调整方案 |
| **生产运营** | 🏭 服装厂岗位职责 | Excel（4 表） | 服装厂小组长岗位职责矩阵、8 项 KPI、绩效评分记录 |
| **生产运营** | 🏭 生产日报 | Excel（5 表） | 生产日报、产量统计、设备运行 OEE、质量检验、人员出勤 |
| **供应链** | 🛒 采购管理套件 | Excel（5 表） | 采购申请单、供应商评估、询价比价、合同审查清单、采购台账 |
| **法务合规** | ⚖️ 合同审查助手 | Word | 主体资格、商务、财务、违约、知产、争议解决六维度审查意见书 |
| **客户服务** | 🎧 客服 SOP | Excel（4 表） | 话术库、工单流程、客诉分类与升级规则、FAQ |
| **情报资讯** | 📰 AI 领袖日报 | HTML | 李开复、OpenAI、Anthropic、DeepMind、Meta 最新动态卡片式日报 |

## Demo 亮点

- **🤖 自然语言智能体**：一句话识别意图，自动调用对应技能。
- **🧩 多场景模板中台**：16 个技能覆盖 HR、销售、财务、生产、供应链、法务、客服、招投标、情报、代码智能体 9 大业务域。
- **💻 企业业务代码智能体**：销售分析、薪酬核算、简历筛选、库存分析、KPI 评分等业务数据任务，自动生成 Python 代码并执行，输出可下载的 Excel/HTML 报告。
- **🧠 记忆沉淀**：每次代码执行任务自动保存生成的代码、执行日志与产物到 `~/.prajna/code_agent_memory`，支持复用与审计。
- **🔌 LLM-ready**：内置 DeepSeek / OpenRouter / OpenAI / Anthropic 适配器，配置 API key 后即可从规则模板升级到开放式代码生成。
- **🎯 多岗位智能适配**：内置 10+ 行业预设、城市薪酬系数、职级带宽体系。
- **🌍 城市薪酬基准**：根据城市等级自动调整薪资、社保公积金基数。
- **📈 职级带宽体系**：P1-P9 职级对应薪酬、绩效与奖金带宽。
- **🔗 公式联动**：所有汇总、绩效、个税、完成率、KPI 加权得分均保留 Excel 公式，可二次编辑。

## 本地运行

```bash
# 1. 克隆仓库
git clone https://github.com/zhixun1981-oss/prajna-guangzhou-demo.git
cd prajna-guangzhou-demo

# 2. 安装依赖
pip install -r requirements.txt

# 3. 启动 Streamlit
streamlit run streamlit_app.py
```

浏览器访问 `http://localhost:8501`。

## 启用 LLM 开放式代码生成（可选）

prajna-code-agent 默认使用规则模板引擎，无需 API key。如需启用 LLM 生成开放式业务代码，配置以下环境变量：

```bash
export PRAJNA_LLM_PROVIDER=deepseek      # deepseek | openrouter | openai | anthropic
export PRAJNA_LLM_API_KEY=sk-xxxxxxxx
export PRAJNA_LLM_MODEL=deepseek-chat    # 可选，不填使用默认模型
export PRAJNA_LLM_BASE_URL=              # 可选，自定义 API 地址
```

配置后，代码智能体会优先调用 LLM；若 LLM 未配置或调用失败，自动回退到规则模板。

### 无 API key 测试 LLM 路径

设置 mock 模式即可在不调用真实 LLM 的情况下测试 LLM 生成路径：

```bash
export PRAJNA_LLM_MOCK=1
# 或
export PRAJNA_LLM_PROVIDER=mock
```

然后勾选「使用 LLM 生成代码」即可体验 LLM 代码生成与执行的完整流程。

## 部署到 Render

1. 打开 [render.com](https://render.com/)，用 GitHub 登录。
2. 点击 **New +** → **Web Service**。
3. 选择仓库 `zhixun1981-oss/prajna-guangzhou-demo`，分支 `main`。
4. 填写：
   - **Name**：`prajna-guangzhou-demo`
   - **Runtime**：Python 3
   - **Build Command**：`pip install -r requirements.txt`
   - **Start Command**：`streamlit run streamlit_app.py --server.port $PORT --server.address 0.0.0.0`
5. 选择 **Free** 实例，点击 **Deploy Web Service**。
6. 等待 2-5 分钟，Render 会生成在线链接。

## 仓库结构

```
prajna-guangzhou-demo/
├── streamlit_app.py                  # Demo 主程序
├── README.md                         # 项目说明
├── requirements.txt                  # Python 依赖
├── .streamlit/config.toml            # Streamlit 配置
├── assets/                           # 示例文件
└── skills/                           # 16 个智能体技能
    ├── prajna-code-agent/              # 代码生成 + 沙箱执行 + 记忆沉淀
    ├── prajna-salary-template/
    ├── prajna-sales-weekly-report/
    ├── hr/
    │   ├── prajna-recruitment-assistant/
    │   ├── prajna-compensation-system/
    │   └── prajna-performance-system/
    ├── sales/
    │   └── prajna-bidding-assistant/
    ├── finance/
    │   ├── prajna-ecommerce-finance-kb-catalog/
    │   ├── prajna-financial-dashboard/
    │   └── prajna-budget-execution-ppt/
    ├── manufacturing/
    │   └── prajna-clothing-teamleader-duty/
    ├── procurement/
    │   └── prajna-procurement-assistant/
    ├── legal/
    │   └── prajna-contract-review-assistant/
    ├── customer-service/
    │   └── prajna-customer-service-sop/
    ├── production/
    │   └── prajna-production-daily-report/
    └── business-intelligence/
        └── prajna-ai-leader-daily/
```

## 免责声明

【人工智能生成-需人工核验】本 Demo 生成的所有薪酬结构、绩效考核、奖金福利、合规校验、销售数据、周报内容、招投标文件、招聘信息、财务指标、预算方案等仅供辅助参考，不构成法律、税务、劳动用工、招投标或商业决策建议。最终方案须由企业 HR、法务、财务、销售、采购及管理层依据当地法律法规和公司政策审核确认。
