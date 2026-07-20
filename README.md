# Prajna 企业智能体 · 广州“广智能”超级智能体大赛 Demo

本项目是 **Prajna** 参加[首届超级智能体大赛](http://121.8.227.238/cygxdj/gzn/#/dasai)的可运行成果 Demo，定位 **企业多场景模板中台**：输入简单参数，即可一键生成结构完整、公式联动的 **Excel 薪资模板** 与 **销售团队标准周报**，覆盖 HR 薪酬与销售管理两大高频企业场景。

## 在线 Demo

👉 **部署到 Streamlit Cloud 后，把下面的占位链接替换成你的真实链接**

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/你的用户名/prajna-guangzhou-demo/main/streamlit_app.py)

- 临时本地公网演示（仅作调试）：`streamlit run streamlit_app.py` 后，`npx localtunnel --port 8501`
- 大赛报名表要求持久可访问链接，请使用 Streamlit Cloud（见下方部署步骤）。

## 代码仓库

✅ 已 push 到 GitHub：

**https://github.com/zhixun1981-oss/prajna-guangzhou-demo**

（代码仓库链接填这个）

## 在线演示视频

👉 **上传到 Bilibili / YouTube 后，把下面的占位链接替换成你的真实链接**

例如：`https://www.bilibili.com/video/BVxxxxxxxxxx`

- 视频需自行录制并上传，参考下文「录制建议」
- 上传时标题建议：`Prajna 企业智能体 Demo｜一键生成薪资模板+销售周报｜广智能大赛`
- Bilibili 上传后，在「创作中心 → 稿件管理」复制 `BV` 号即可。

### 录制建议

建议总时长 60-90 秒：

| 时间 | 画面内容 |
|---|---|
| 0-5s | Demo 首页，展示「薪资模板 + 销售周报」双入口 |
| 5-25s | 选择薪资模板，输入行业/岗位/城市/职级，点击生成并下载 |
| 25-45s | 打开下载的 Excel，快速翻一下 8 个工作表 |
| 45-65s | 切换销售周报，输入团队/周期/目标，生成并下载 |
| 65-85s | 打开周报 Excel，翻一下 5 个工作表 |
| 85-90s | 结尾页：Prajna 企业模板中台 · 广智能大赛 |

## 如何获得报名表里的三个链接（10 分钟清单）

| 步骤 | 操作 | 产出 |
|---|---|---|
| 1 | 在 GitHub 新建公开仓库 `prajna-guangzhou-demo` | 仓库首页链接 |
| 2 | `git remote add origin https://github.com/你的用户名/prajna-guangzhou-demo.git` 并 `git push -u origin main` | 代码仓库链接 |
| 3 | 打开 [Streamlit Cloud](https://share.streamlit.io/)，用 GitHub 登录，新建 App，选择该仓库，主文件路径填 `streamlit_app.py` | 在线 Demo 链接 |
| 4 | 把 `assets/demo_video.mp4` 上传到 [Bilibili](https://member.bilibili.com/platform/upload) 或 YouTube | 在线演示视频链接 |
| 5 | 把三个真实链接填回 `README.md` 和大奖报名页 | 报名表完整提交 |

## Demo 功能

- **💰 薪资模板**：输入行业、岗位、城市、职级，生成包含 8 个工作表的完整 Excel（薪资结构、绩效考核、调薪机制、奖金方案、福利明细、计算示例、合规校验、绩效系数对照表）。
- **📊 销售团队标准周报**：输入团队名称、类型、周期、销售目标，生成包含 5 个工作表的 Excel（封面汇总、核心业绩数据、重点商机进展、问题分析、下周工作计划）。
- **多行业/多岗位适配**：薪资模板内置 10 套预设，支持自然语言岗位描述自动推断；周报模板内置 4 种销售团队预设。
- **城市薪酬基准 + 职级带宽**：自动调整薪资、社保公积金基数与奖金带宽。
- **公式联动**：所有汇总、绩效加权、个税、完成率、KPI 加权得分均使用 Excel 公式，支持二次编辑。

## 本地运行

```bash
# 1. 克隆仓库
git clone https://github.com/你的用户名/prajna-guangzhou-demo.git
cd prajna-guangzhou-demo

# 2. 安装依赖
pip install -r requirements.txt

# 3. 启动 Streamlit
streamlit run streamlit_app.py
```

浏览器访问 `http://localhost:8501`。

## 部署到 Streamlit Cloud（推荐，免费）

1. 将本仓库 push 到你的 GitHub 公开仓库。
2. 打开 [Streamlit Cloud](https://share.streamlit.io/)。
3. 点击 **New app** → 选择仓库 → 主文件路径填 `streamlit_app.py`。
4. 点击 **Deploy**，等待 1-2 分钟即可获得在线链接。

## 部署到 Render（备选）

1. 将仓库 push 到 GitHub。
2. 在 [Render](https://render.com/) 创建 **Web Service**。
3. Build command：`pip install -r requirements.txt`
4. Start command：`streamlit run streamlit_app.py --server.port $PORT --server.address 0.0.0.0`
5. 部署后获得 `https://xxx.onrender.com`。

## 大赛报名表填写参考

| 字段 | 建议填写内容 |
|---|---|
| 报名赛道 | 开放赛道 → 垂类行业智能体 |
| 项目所属主题方向 | 垂类行业智能体 |
| 项目简介 | Prajna 是面向企业服务的多智能体平台，本 Demo 展示其作为「企业模板中台」的能力：通过自然语言与结构化参数，一键生成 HR 薪资模板与销售团队标准周报，覆盖企业薪酬管理与销售运营两大高频场景。 |
| 核心问题与目标用户 | 中小企业缺乏标准化薪酬体系与周报机制，岗位多、城市差异大、销售数据分散。目标用户为企业 HR、销售主管、HR SaaS 厂商、人力资源外包公司。 |
| 核心功能与应用场景 | 1）自然语言识别岗位；2）城市/职级薪酬基准；3）自动生成 8 工作表薪资模板；4）自动生成 5 工作表销售周报；5）合规校验与业绩追踪。 |
| 技术实现路径 | 规则引擎 + JSON 模板 + openpyxl 公式化输出；FastAPI/Streamlit 提供 Web 交互；Skills 体系支持横向扩展。 |
| 技术创新点 | 无需 LLM 即可多岗位/多团队推断；城市薪酬系数与职级带宽解耦；所有输出保留 Excel 公式，支持二次编辑；一个 Demo 同时覆盖 HR 与销售双场景。 |
| 项目当前阶段 | 已完成薪资模板与销售周报核心技能开发，具备双场景 Demo 演示能力。 |
| 可运行成果链接 | 你的 Streamlit / Render 链接 |

## 免责声明

【人工智能生成-需人工核验】本 Demo 生成的所有薪酬结构、绩效考核、奖金福利、合规校验、销售数据、周报内容仅供辅助参考，不构成法律、税务或商业决策建议。最终方案须由企业 HR、法务、财务、销售管理层依据当地法律法规和公司政策审核确认。
