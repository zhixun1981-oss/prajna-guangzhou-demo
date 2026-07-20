# Prajna 企业智能体 · 广州“广智能”超级智能体大赛 Demo

本项目是 **Prajna** 参加[首届超级智能体大赛](http://121.8.227.238/cygxdj/gzn/#/dasai)的可运行成果 Demo，聚焦 **HR 薪酬场景**：输入行业、岗位、城市、职级，即可一键生成结构完整、公式联动的 Excel 薪资模板。

## 在线 Demo

👉 **部署后替换为你的 Streamlit Cloud / Render / Vercel 链接**

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/你的用户名/prajna-guangzhou-demo/main/streamlit_app.py)

## 代码仓库

👉 **部署后替换为你的 GitHub 仓库链接**

例如：`https://github.com/你的用户名/prajna-guangzhou-demo`

## 在线演示视频

👉 **部署后替换为你的 Bilibili / YouTube / 腾讯视频链接**

例如：`https://www.bilibili.com/video/BVxxxxxxxxxx`

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
| 项目简介 | Prajna 是面向企业服务的多智能体平台，覆盖 HR、招聘、薪酬、绩效、合规等场景。本 Demo 展示其在 HR 薪酬核算场景下的自动化能力。 |
| 核心问题与目标用户 | 中小企业 HR 缺乏标准化薪酬体系，岗位多、城市差异大、合规要求高。目标用户为企业 HR、HR SaaS 厂商、人力资源外包公司。 |
| 核心功能与应用场景 | 1）自然语言识别岗位；2）城市/职级薪酬基准；3）自动生成 8 工作表 Excel；4）合规校验；5）历史版本对比。 |
| 技术实现路径 | 规则引擎 + JSON 模板 + openpyxl 公式化输出；FastAPI/Streamlit 提供 Web 交互；Skills 体系支持横向扩展。 |
| 技术创新点 | 无需 LLM 即可多岗位推断；城市薪酬系数与职级带宽解耦；所有输出保留 Excel 公式，支持二次编辑。 |
| 项目当前阶段 | 已完成核心技能开发，具备多岗位 Demo 演示能力。 |
| 可运行成果链接 | 你的 Streamlit / Render 链接 |

## 免责声明

【人工智能生成-需人工核验】本 Demo 生成的所有薪酬结构、绩效考核、奖金福利、合规校验结果仅供辅助参考，不构成法律或税务建议。最终薪酬方案须由企业 HR、法务、财务及管理层依据当地法律法规和公司政策审核确认。
