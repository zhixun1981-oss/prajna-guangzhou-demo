# Prajna 企业智能体 · 多场景模板中台 Demo

本项目是 **Prajna** 的企业级模板中台 Demo，聚焦两个高频企业场景：

- **HR 薪酬场景**：输入行业、岗位、城市、职级，一键生成结构完整、公式联动的 Excel 薪资模板。
- **销售管理场景**：输入团队、周期、销售目标，一键生成销售团队标准周报。

所有输出均保留 Excel 公式，支持二次编辑。

## Demo 功能

- **💰 薪资模板**：生成 8 个工作表（薪资结构、绩效考核、调薪机制、奖金方案、福利明细、计算示例、合规校验、绩效系数对照表）。
- **📊 销售团队标准周报**：生成 5 个工作表（封面汇总、核心业绩数据、重点商机进展、问题分析、下周工作计划）。
- **多行业/多岗位适配**：薪资模板内置 10 套预设，支持自然语言岗位描述自动推断；周报模板内置 4 种销售团队预设。
- **城市薪酬基准 + 职级带宽**：自动调整薪资、社保公积金基数与奖金带宽。
- **公式联动**：所有汇总、绩效加权、个税、完成率、KPI 加权得分均使用 Excel 公式。

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
├── assets/                           # 示例 Excel 文件
└── skills/                           # 模板生成技能
    ├── prajna-salary-template/
    └── prajna-sales-weekly-report/
```

## 免责声明

【人工智能生成-需人工核验】本 Demo 生成的所有薪酬结构、绩效考核、奖金福利、合规校验、销售数据、周报内容仅供辅助参考，不构成法律、税务或商业决策建议。最终方案须由企业 HR、法务、财务、销售管理层依据当地法律法规和公司政策审核确认。
