# Prajna 企业智能体 · 多场景模板中台 Demo

本项目是 **Prajna** 参加广州「广智能」超级智能体大赛的可运行成果，定位为企业级多场景模板中台。用户只需输入参数或一句话自然语言描述，Prajna 即可自动识别意图并生成对应的企业级文档：Excel、Word、PPT、HTML。

## 在线 Demo

**https://prajna-guangzhou-demo.onrender.com**

## 核心能力

**🤖 自然语言智能体模式**：输入一句话，Prajna 自动识别意图、匹配技能、填充参数并生成文档。

例如：
- "帮我做一份深圳互联网产品经理 P5 的薪资模板"
- "生成电商销售团队本周周报，目标 80 万"
- "搭建智云电商财务知识库目录"
- "帮我做一份智慧园区建设项目的投标书"
- "生成招聘电商运营助理的 JD 和面试评估表"

## 支持的 11 个企业场景

| 业务域 | 技能 | 输出 | 说明 |
|---|---|---|---|
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
| **情报资讯** | 📰 AI 领袖日报 | HTML | 李开复、OpenAI、Anthropic、DeepMind、Meta 最新动态卡片式日报 |

## Demo 亮点

- **🤖 自然语言智能体**：一句话识别意图，自动调用对应技能。
- **🧩 多场景模板中台**：11 个技能覆盖 HR、销售、财务、生产、招投标、情报 6 大业务域。
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
└── skills/                           # 11 个模板生成技能
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
    └── business-intelligence/
        └── prajna-ai-leader-daily/
```

## 大赛报名表填写参考

| 字段 | 建议填写内容 |
|---|---|
| 项目简介 | Prajna 是面向企业服务的多智能体平台，本次 Demo 展示了其「企业模板中台」能力：通过自然语言或结构化参数，一键生成 HR、销售、财务、生产、招投标等多场景企业文档。 |
| 核心问题与目标用户 | 中小企业缺乏标准化人力资源、销售管理、财务经营与招投标体系，岗位多、城市差异大、文档制作成本高。目标用户为企业 HR、销售主管、财务经理、采购/投标专员。 |
| 核心功能与应用场景 | 1）自然语言识别意图；2）城市/职级薪酬基准；3）自动生成 11 类企业模板；4）Excel 公式联动；5）多行业/多岗位适配。 |
| 技术实现路径 | 规则引擎 + JSON 模板 + openpyxl/python-docx/python-pptx 输出；Streamlit 提供 Web 交互；技能体系支持横向扩展。 |
| 技术创新点 | 自然语言一键切换 11 个企业场景；城市薪酬系数与职级带宽解耦；所有输出保留公式支持二次编辑；一个 Demo 覆盖 HR、销售、财务、生产、招投标五大业务域。 |
| 可运行成果链接 | https://prajna-guangzhou-demo.onrender.com |
| 代码仓库链接 | https://github.com/zhixun1981-oss/prajna-guangzhou-demo |

## 免责声明

【人工智能生成-需人工核验】本 Demo 生成的所有薪酬结构、绩效考核、奖金福利、合规校验、销售数据、周报内容、招投标文件、招聘信息、财务指标、预算方案等仅供辅助参考，不构成法律、税务、劳动用工、招投标或商业决策建议。最终方案须由企业 HR、法务、财务、销售、采购及管理层依据当地法律法规和公司政策审核确认。
