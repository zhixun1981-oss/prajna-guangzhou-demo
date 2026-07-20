import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import streamlit as st

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
APP_DIR = Path(__file__).resolve().parent
SKILLS_DIR = APP_DIR / "skills"

# 11 skills
SALARY_SCRIPT = SKILLS_DIR / "prajna-salary-template" / "scripts" / "generate_salary_template.py"
SALES_SCRIPT = SKILLS_DIR / "prajna-sales-weekly-report" / "scripts" / "generate_sales_weekly_report.py"
FINANCE_KB_SCRIPT = SKILLS_DIR / "finance" / "prajna-ecommerce-finance-kb-catalog" / "scripts" / "generate_catalog.py"
CLOTHING_SCRIPT = SKILLS_DIR / "manufacturing" / "prajna-clothing-teamleader-duty" / "scripts" / "generate_clothing_teamleader_duty.py"
AI_DAILY_SCRIPT = SKILLS_DIR / "business-intelligence" / "prajna-ai-leader-daily" / "scripts" / "generate_ai_leader_daily.py"
FIN_DASHBOARD_SCRIPT = SKILLS_DIR / "finance" / "prajna-financial-dashboard" / "scripts" / "generate_financial_dashboard.py"
BUDGET_PPT_SCRIPT = SKILLS_DIR / "finance" / "prajna-budget-execution-ppt" / "scripts" / "generate_budget_execution_ppt.py"
BIDDING_SCRIPT = SKILLS_DIR / "sales" / "prajna-bidding-assistant" / "scripts" / "generate_bidding_kit.py"
RECRUIT_SCRIPT = SKILLS_DIR / "hr" / "prajna-recruitment-assistant" / "scripts" / "generate_recruitment_kit.py"
COMPENSATION_SCRIPT = SKILLS_DIR / "hr" / "prajna-compensation-system" / "scripts" / "generate_compensation_system.py"
PERFORMANCE_SCRIPT = SKILLS_DIR / "hr" / "prajna-performance-system" / "scripts" / "generate_performance_system.py"

SALARY_SAMPLE = APP_DIR / "assets" / "sample_薪资模板_广州_电商运营助理.xlsx"
SALES_SAMPLE = APP_DIR / "assets" / "sample_销售周报_示例.xlsx"

HISTORY_DIR = Path.home() / ".prajna" / "demo_history"

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Prajna 企业智能体 · 多场景模板中台",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .hero {
        background: linear-gradient(135deg, #1F4E78 0%, #2E86C1 100%);
        padding: 2.5rem 2rem;
        border-radius: 16px;
        color: white;
        margin-bottom: 2rem;
    }
    .hero h1 {
        margin: 0 0 0.5rem 0;
        font-size: 2.4rem;
    }
    .hero p {
        margin: 0;
        font-size: 1.1rem;
        opacity: 0.95;
    }
    .feature-card {
        background: #f8f9fa;
        border-radius: 12px;
        padding: 1.2rem;
        border-left: 5px solid #1F4E78;
        height: 100%;
    }
    .preview-card {
        background: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 12px;
        padding: 1.2rem;
        height: 100%;
    }
    .thinking-card {
        background: #f0f7ff;
        border: 1px solid #b3d7ff;
        border-radius: 12px;
        padding: 1.2rem;
        height: 100%;
    }
    .metric-box {
        background: #ffffff;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    }
    .footer {
        text-align: center;
        color: #888;
        font-size: 0.85rem;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #eee;
    }
    .stButton>button {
        border-radius: 8px;
        padding: 0.6rem 1.2rem;
    }
    .skill-tab {
        background: #ffffff;
        border-radius: 12px;
        padding: 1rem;
        border: 1px solid #e0e0e0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Skill registry
# ---------------------------------------------------------------------------
SKILL_REGISTRY = {
    "salary": {
        "icon": "💰",
        "name": "薪资模板",
        "script": SALARY_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成包含 8 个工作表的完整薪资模板",
    },
    "sales": {
        "icon": "📊",
        "name": "销售周报",
        "script": SALES_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成包含 5 个工作表的销售团队标准周报",
    },
    "finance_kb": {
        "icon": "📁",
        "name": "财务知识库目录",
        "script": FINANCE_KB_SCRIPT,
        "ext": "md",
        "mime": "text/markdown",
        "desc": "生成电商企业财务知识库钉钉文档目录",
    },
    "clothing_duty": {
        "icon": "🏭",
        "name": "服装厂岗位职责",
        "script": CLOTHING_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成服装厂小组长岗位职责与 KPI 考核表",
    },
    "ai_daily": {
        "icon": "📰",
        "name": "AI 领袖日报",
        "script": AI_DAILY_SCRIPT,
        "ext": "html",
        "mime": "text/html",
        "desc": "生成 AI 领袖动态日报 HTML 页面",
    },
    "finance_dashboard": {
        "icon": "📈",
        "name": "财务核心指标看板",
        "script": FIN_DASHBOARD_SCRIPT,
        "ext": "html",
        "mime": "text/html",
        "desc": "生成含 Chart.js 交互图表的财务看板",
    },
    "budget_ppt": {
        "icon": "📑",
        "name": "预算执行汇报 PPT",
        "script": BUDGET_PPT_SCRIPT,
        "ext": "pptx",
        "mime": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "desc": "生成本月预算执行情况汇报 PPT",
    },
    "bidding": {
        "icon": "🎯",
        "name": "招投标助手",
        "script": BIDDING_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成招投标套件 Excel + 投标文件 Word 大纲",
    },
    "recruitment": {
        "icon": "🤝",
        "name": "招聘助手",
        "script": RECRUIT_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成招聘套件：JD、面试评估表、Offer 模板",
    },
    "compensation": {
        "icon": "💵",
        "name": "薪酬体系",
        "script": COMPENSATION_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成企业岗位薪酬矩阵、薪酬结构与预算",
    },
    "performance": {
        "icon": "⭐",
        "name": "绩效体系",
        "script": PERFORMANCE_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成 KPI 指标库、绩效合同、评分表与 PIP",
    },
}

# ---------------------------------------------------------------------------
# Natural language parser for Prajna Agent Mode
# ---------------------------------------------------------------------------
CITIES = [
    "北京", "上海", "广州", "深圳", "杭州", "成都", "武汉", "西安", "南京", "苏州",
    "重庆", "天津", "长沙", "郑州", "东莞", "宁波", "佛山", "合肥", "青岛", "厦门",
    "无锡", "福州", "济南", "沈阳", "大连", "昆明", "南昌", "哈尔滨", "长春", "石家庄",
    "贵阳", "南宁", "温州", "珠海", "惠州", "中山", "南通", "绍兴", "烟台", "潍坊",
    "唐山", "保定", "洛阳", "常州", "徐州", "泉州", "金华", "嘉兴", "台州", "临沂",
]

INDUSTRIES = {
    "互联网": ["互联网", "科技", "IT", "软件"],
    "制造业": ["制造业", "工厂", "工业"],
    "零售": ["零售", "连锁", "门店"],
    "金融": ["金融", "银行", "证券", "保险"],
    "餐饮": ["餐饮", "餐厅", "酒店"],
    "新能源": ["新能源", "光伏", "储能"],
    "房地产": ["房地产", "地产", "建筑"],
    "教育": ["教育", "培训", "学校"],
    "医疗": ["医疗", "医药", "医院"],
    "物流": ["物流", "快递", "供应链"],
    "汽车": ["汽车", "车联网", "自动驾驶"],
    "综合": ["综合", "行政", "通用"],
}

INTENT_KEYWORDS = {
    "salary": ["薪资", "工资", "薪酬", "salary", "收入", "待遇", "薪水"],
    "sales": ["周报", "销售周报", "weekly report", "销售报告", "销售团队", "销售报表"],
    "finance_kb": ["财务知识库", "知识库目录", "钉钉文档目录", "财务目录"],
    "clothing_duty": ["服装厂", "小组长", "岗位职责", "服装", "班组长", "缝纫"],
    "ai_daily": ["ai日报", "ai领袖", "李开复", "openai", "altman", "anthropic", "deepmind", "meta"],
    "finance_dashboard": ["财务看板", "财务指标", "资产负债率", "净利润率", "现金流量比率", "dashboard"],
    "budget_ppt": ["预算", "预算执行", "预算汇报", "ppt"],
    "bidding": ["投标", "招标", "招投标", "标书", "bid"],
    "recruitment": ["招聘", "jd", "岗位说明书", "面试", "offer", "recruitment"],
    "compensation": ["薪酬体系", "薪酬矩阵", "职级薪酬", "compensation", "薪酬带宽"],
    "performance": ["绩效", "kpi", "绩效考核", "绩效体系", "performance", "okr", "pip"],
}


def detect_intent(text):
    text_lower = text.lower()
    scores = {}
    for intent, kws in INTENT_KEYWORDS.items():
        scores[intent] = sum(1 for kw in kws if kw.lower() in text_lower)
    # bidding should not be triggered by "招聘"
    if "招聘" in text and scores.get("bidding", 0) == 1 and "招标" not in text:
        scores["bidding"] = 0
    best = max(scores, key=scores.get)
    if scores[best] == 0:
        # fallback: P-level -> salary, 销售/团队 -> sales
        if re.search(r"P\d", text.upper()):
            return "salary"
        if "团队" in text or "销售" in text:
            return "sales"
        return "salary"
    return best


def extract_city(text):
    for city in CITIES:
        if city in text:
            return city
    return "广州"


def extract_level(text):
    m = re.search(r"P\d", text.upper())
    return m.group(0) if m else "P2"


def extract_industry(text):
    text_lower = text.lower()
    for industry, kws in INDUSTRIES.items():
        for kw in kws:
            if kw in text_lower:
                return industry
    return "互联网"


def extract_position(text):
    cleaned = text
    cleaned = re.sub(r"P\d", "", cleaned, flags=re.I)
    for city in CITIES:
        cleaned = cleaned.replace(city, "")
    for kws in INDUSTRIES.values():
        for kw in kws:
            cleaned = cleaned.replace(kw, "")
    generic_words = [
        "模板", "生成", "做", "一份", "的", "和", "与", "给我", "帮我", "请", "需要", "表",
        "薪资", "工资", "薪酬", "薪水", "周报", "报告", "报表", "岗位说明书", "招聘", "面试",
        "绩效", "kpi", "okr", "pip", "投标", "招标", "标书", "预算", "ppt", "知识库", "目录",
        "财务", "看板", "dashboard", "ai", "日报", "领袖", "服装厂", "小组长", "班组长", "缝纫",
    ]
    for kw in generic_words:
        cleaned = cleaned.replace(kw, "")
    cleaned = re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9]", "", cleaned)
    cleaned = re.sub(r"(表|模板|岗位|报告|助手|套件)+$", "", cleaned)
    return cleaned.strip()[:10] or "岗位"


def extract_team(text):
    team = text.split("周报")[0].strip() or "销售团队"
    team = re.sub(r"^(帮我|给我|请|做|一份|的|生成)", "", team).strip() or "销售团队"
    team = re.sub(r"(本周|第\s*\d+\s*周|的|销售)+$", "", team).strip()
    team = re.sub(r"^销售", "", team).strip() or "销售团队"
    return team


def extract_week(text):
    m = re.search(r"第\s*(\d+)\s*周", text)
    if m:
        return f"{datetime.now().year}年第{int(m.group(1))}周"
    return f"{datetime.now().year}年第{datetime.now().isocalendar().week}周"


def extract_target(text):
    m = re.search(r"(\d+(?:\.\d+)?)\s*[万亿]", text)
    if m:
        num = float(m.group(1))
        if "万" in text:
            return int(num * 10000)
        if "亿" in text:
            return int(num * 100000000)
    m2 = re.search(r"(\d+)\s*元", text)
    if m2:
        return int(m2.group(1))
    return 1200000


def extract_amount(text):
    m = re.search(r"(\d+(?:\.\d+)?)\s*[万亿]", text)
    if m:
        num = float(m.group(1))
        if "万" in text:
            return int(num * 10000)
        if "亿" in text:
            return int(num * 100000000)
    m2 = re.search(r"(\d+)\s*元", text)
    if m2:
        return int(m2.group(1))
    return 1000000


def infer_sales_preset(text):
    text = text.lower()
    if "saas" in text or "互联网" in text or "科技" in text:
        return "互联网/SaaS 销售团队"
    if "电商" in text:
        return "电商销售团队"
    if "零售" in text or "门店" in text or "线下" in text:
        return "线下零售销售团队"
    return "通用销售团队"


def infer_finance_preset(text):
    text = text.lower()
    if "零售" in text or "连锁" in text:
        return "零售连锁企业"
    if "制造" in text:
        return "通用制造集团"
    if "saas" in text or "互联网" in text or "科技" in text:
        return "互联网/SaaS企业"
    if "新能源" in text:
        return "新能源科技企业"
    return "通用企业"


def parse_agent_input(text):
    text = text.strip()
    if not text:
        return None

    intent = detect_intent(text)
    city = extract_city(text)
    level = extract_level(text)
    industry = extract_industry(text)

    result = {
        "intent": intent,
        "raw": text,
        "city": city,
        "level": level,
        "industry": industry,
    }

    if intent == "salary":
        position = extract_position(text)
        if position in ["", "岗位"] and industry:
            position = f"{industry}专员"
        result.update({
            "position": position,
            "description": text,
        })
    elif intent == "sales":
        result.update({
            "team": extract_team(text),
            "preset": infer_sales_preset(text),
            "week": extract_week(text),
            "target": extract_target(text),
            "author": "Prajna",
            "date": datetime.now().strftime("%Y-%m-%d"),
        })
    elif intent == "finance_kb":
        result.update({
            "company": re.sub(r"^(帮我|给我|请|做|一份|的|生成|搭建|电商企业财务知识库|财务知识库|钉钉文档目录)", "", text).strip()[:20] or "智云电商",
            "author": "财务部",
            "date": datetime.now().strftime("%Y-%m-%d"),
        })
    elif intent == "clothing_duty":
        result.update({
            "team": "缝纫一组",
            "factory": "成衣A厂",
            "month": datetime.now().strftime("%Y-%m"),
            "author": "Prajna",
        })
    elif intent == "ai_daily":
        result.update({
            "date": datetime.now().strftime("%Y-%m-%d"),
        })
    elif intent == "finance_dashboard":
        result.update({
            "preset": infer_finance_preset(text),
            "company": re.sub(r"^(帮我|给我|请|做|一份|的|生成|搭建|财务核心指标看板|财务看板)", "", text).strip()[:20] or "示范企业股份",
            "period": datetime.now().strftime("%Y年%m月"),
        })
    elif intent == "budget_ppt":
        result.update({
            "month": datetime.now().strftime("%Y年%m月"),
            "company": re.sub(r"^(帮我|给我|请|做|一份|的|生成|本月预算执行情况汇报|预算执行汇报)", "", text).strip()[:20] or "启明星科技",
            "author": "财务部",
        })
    elif intent == "bidding":
        result.update({
            "project": re.sub(r"^(帮我|给我|请|做|一份|的|生成|投标|招标|招投标)", "", text).strip()[:30] or "智慧园区建设项目",
            "bidder": "智讯科技有限公司",
            "tenderer": "广州高新区管委会",
            "amount": extract_amount(text),
            "duration": "180天",
            "industry": industry,
        })
    elif intent == "recruitment":
        position = extract_position(text)
        if position in ["", "岗位"]:
            position = "电商运营助理"
        result.update({
            "position": position,
            "department": f"{industry}运营部" if industry != "综合" else "综合管理部",
            "salary_min": 6000,
            "salary_max": 9000,
            "reports_to": "部门经理",
            "headcount": 1,
            "urgency": "中",
        })
    elif intent == "compensation":
        result.update({
            "company": re.sub(r"^(帮我|给我|请|做|一份|的|生成|搭建|薪酬体系|企业薪酬)", "", text).strip()[:20] or "智云科技",
            "industry": industry,
            "scale": 200,
            "budget": 30000000,
            "growth": 5.0,
        })
    elif intent == "performance":
        position = extract_position(text)
        if position in ["", "岗位"]:
            position = "电商运营助理"
        result.update({
            "department": f"{industry}运营部" if industry != "综合" else "综合管理部",
            "position": position,
            "cycle": "季度",
            "method": "KPI",
            "levels": "A/B/C/D",
            "company": "Prajna示范企业",
        })

    return result


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("🧠 Prajna")
    st.markdown("**企业多场景模板中台**")
    st.divider()
    st.markdown(
        """
        ### 关于本 Demo
        展示 Prajna 通过结构化参数与自然语言描述，
        一键生成企业级文档的能力。

        ### 当前支持 11 个场景
        - 💰 薪资模板
        - 📊 销售周报
        - 📁 财务知识库目录
        - 🏭 服装厂岗位职责
        - 📰 AI 领袖日报
        - 📈 财务核心指标看板
        - 📑 预算执行汇报 PPT
        - 🎯 招投标助手
        - 🤝 招聘助手
        - 💵 薪酬体系
        - ⭐ 绩效体系
        """
    )
    st.divider()
    st.markdown(
        """
        ### 快速链接
        - [GitHub 仓库](https://github.com/zhixun1981-oss/prajna-guangzhou-demo)
        - [在线 Demo](https://prajna-guangzhou-demo.onrender.com)
        """
    )
    st.divider()
    st.caption("【人工智能生成-需人工核验】")

# ---------------------------------------------------------------------------
# Hero
# ---------------------------------------------------------------------------
st.markdown(
    """
    <div class="hero">
        <h1>Prajna 企业智能体 · 多场景模板中台</h1>
        <p>输入参数或自然语言，一键生成 HR、销售、财务、生产、招投标等企业级模板</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Quick stats
# ---------------------------------------------------------------------------
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown('<div class="metric-box"><h3>🧠</h3><p>自然语言理解</p><h5>智能体模式</h5></div>', unsafe_allow_html=True)
with c2:
    st.markdown('<div class="metric-box"><h3>📦</h3><p>企业场景</p><h5>11 个技能</h5></div>', unsafe_allow_html=True)
with c3:
    st.markdown('<div class="metric-box"><h3>🧩</h3><p>多岗位适配</h3><h5>10+ 行业</h5></div>', unsafe_allow_html=True)
with c4:
    st.markdown('<div class="metric-box"><h3>🔗</h3><p>可二次编辑</p><h5>Excel 公式</h5></div>', unsafe_allow_html=True)

st.divider()

# ---------------------------------------------------------------------------
# Sample downloads
# ---------------------------------------------------------------------------
st.subheader("📎 快速体验：下载示例模板")
sample_cols = st.columns(2)
with sample_cols[0]:
    if SALARY_SAMPLE.exists():
        with open(SALARY_SAMPLE, "rb") as f:
            st.download_button(
                label="⬇️ 下载：广州 · 电商运营助理薪资模板",
                data=f,
                file_name=SALARY_SAMPLE.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.warning("薪资示例文件未找到。")

with sample_cols[1]:
    if SALES_SAMPLE.exists():
        with open(SALES_SAMPLE, "rb") as f:
            st.download_button(
                label="⬇️ 下载：销售团队标准周报",
                data=f,
                file_name=SALES_SAMPLE.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.warning("周报示例文件未找到。")

st.divider()

# ---------------------------------------------------------------------------
# Template selection
# ---------------------------------------------------------------------------
st.subheader("1️⃣ 选择交互模式")
template_type = st.radio(
    "交互模式",
    ["🤖 Prajna 智能体模式（自然语言）", "🛠️ 手动选择技能"],
    horizontal=True,
    label_visibility="collapsed",
)

# ---------------------------------------------------------------------------
# Prajna Agent Mode
# ---------------------------------------------------------------------------
if template_type == "🤖 Prajna 智能体模式（自然语言）":
    st.markdown("<br>", unsafe_allow_html=True)
    left, right = st.columns([2, 1])

    with left:
        st.markdown("#### 2️⃣ 告诉 Prajna 你要什么")
        agent_input = st.text_area(
            "自然语言输入",
            value="帮我做一份深圳互联网产品经理 P5 的薪资模板",
            height=120,
            placeholder="例如：帮我做一份深圳互联网产品经理 P5 的薪资模板 / 生成电商销售团队本周周报 / 搭建智云电商财务知识库目录",
            label_visibility="collapsed",
        )

        st.markdown(
            """
            <small>💡 你可以这样说：</small><br>
            <small>• "帮我做一份上海制造业生产主管 P4 的薪资模板"</small><br>
            <small>• "生成电商销售团队本周周报，目标 80 万"</small><br>
            <small>• "搭建智云电商财务知识库目录"</small><br>
            <small>• "生成服装厂缝纫一组小组长岗位职责 KPI"</small><br>
            <small>• "帮我做一份智慧园区建设项目的投标书"</small><br>
            <small>• "生成招聘电商运营助理的 JD 和面试评估表"</small>
            """,
            unsafe_allow_html=True,
        )

    with right:
        st.markdown('<div class="thinking-card">', unsafe_allow_html=True)
        st.markdown("#### 🧠 Prajna 思考中...")
        parsed = parse_agent_input(agent_input)
        if parsed:
            meta = SKILL_REGISTRY[parsed["intent"]]
            st.markdown(f"**识别意图：** {meta['icon']} {meta['name']}")
            # Show key params
            lines = []
            for k, v in parsed.items():
                if k in ("intent", "raw"):
                    continue
                if isinstance(v, str) and v:
                    lines.append(f"**{k}：** {v}")
            st.markdown("  \n".join(lines[:8]))
        else:
            st.markdown("请输入你想生成的内容...")
        st.markdown("</div>", unsafe_allow_html=True)

    execute_clicked = st.button("🚀 Prajna 执行", type="primary", width="stretch")

    if execute_clicked:
        if not parsed:
            st.error("无法识别输入内容，请尝试更清晰的描述。")
            st.stop()

        intent = parsed["intent"]
        meta = SKILL_REGISTRY[intent]

        # Build command based on intent
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        HISTORY_DIR.mkdir(parents=True, exist_ok=True)

        if intent == "salary":
            out_name = f"prajna_薪资模板_{parsed['industry']}_{parsed['position']}_{timestamp}.xlsx"
            out_path = HISTORY_DIR / out_name
            cmd = [
                sys.executable, str(meta["script"]),
                "--industry", parsed["industry"],
                "--position", parsed["position"],
                "--city", parsed["city"],
                "--level", parsed["level"],
                "--description", parsed["raw"],
                "--output", str(out_path),
            ]
        elif intent == "sales":
            safe_team = parsed["team"].replace(" ", "_")
            out_name = f"prajna_销售周报_{safe_team}_{timestamp}.xlsx"
            out_path = HISTORY_DIR / out_name
            cmd = [
                sys.executable, str(meta["script"]),
                "--preset", parsed["preset"],
                "--team", parsed["team"],
                "--week", parsed["week"],
                "--sales-target", str(parsed["target"]),
                "--author", parsed["author"],
                "--date", parsed["date"],
                "--output", str(out_path),
            ]
        elif intent == "finance_kb":
            out_name = f"prajna_财务知识库目录_{parsed['company']}_{timestamp}.md"
            out_path = HISTORY_DIR / out_name
            cmd = [
                sys.executable, str(meta["script"]),
                "--company", parsed["company"],
                "--author", parsed["author"],
                "--date", parsed["date"],
                "--output", str(out_path),
            ]
        elif intent == "clothing_duty":
            out_name = f"prajna_服装厂小组长岗位职责_{parsed['team']}_{timestamp}.xlsx"
            out_path = HISTORY_DIR / out_name
            cmd = [
                sys.executable, str(meta["script"]),
                "--team", parsed["team"],
                "--factory", parsed["factory"],
                "--month", parsed["month"],
                "--author", parsed["author"],
                "--output", str(out_path),
            ]
        elif intent == "ai_daily":
            out_name = f"prajna_AI领袖日报_{parsed['date']}_{timestamp}.html"
            out_path = HISTORY_DIR / out_name
            cmd = [
                sys.executable, str(meta["script"]),
                "--date", parsed["date"],
                "--output", str(out_path),
            ]
        elif intent == "finance_dashboard":
            out_name = f"prajna_财务核心指标看板_{parsed['company']}_{timestamp}.html"
            out_path = HISTORY_DIR / out_name
            cmd = [
                sys.executable, str(meta["script"]),
                "--preset", parsed["preset"],
                "--company", parsed["company"],
                "--period", parsed["period"],
                "--output", str(out_path),
            ]
        elif intent == "budget_ppt":
            out_name = f"prajna_预算执行汇报_{parsed['company']}_{timestamp}.pptx"
            out_path = HISTORY_DIR / out_name
            cmd = [
                sys.executable, str(meta["script"]),
                "--month", parsed["month"],
                "--company", parsed["company"],
                "--author", parsed["author"],
                "--output", str(out_path),
            ]
        elif intent == "bidding":
            safe_project = re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9_-]", "_", parsed["project"])[:20]
            out_prefix = HISTORY_DIR / f"prajna_招投标套件_{safe_project}_{timestamp}"
            out_path = Path(f"{out_prefix}.xlsx")
            cmd = [
                sys.executable, str(meta["script"]),
                "--project", parsed["project"],
                "--bidder", parsed["bidder"],
                "--tenderer", parsed["tenderer"],
                "--amount", str(parsed["amount"]),
                "--duration", parsed["duration"],
                "--industry", parsed["industry"],
                "--output", str(out_prefix),
                "--format", "all",
            ]
        elif intent == "recruitment":
            safe_pos = re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9_-]", "_", parsed["position"])[:10]
            out_prefix = HISTORY_DIR / f"prajna_招聘套件_{safe_pos}_{parsed['city']}_{timestamp}"
            out_path = Path(f"{out_prefix}.xlsx")
            cmd = [
                sys.executable, str(meta["script"]),
                "--position", parsed["position"],
                "--department", parsed["department"],
                "--city", parsed["city"],
                "--level", parsed["level"],
                "--salary-min", str(parsed["salary_min"]),
                "--salary-max", str(parsed["salary_max"]),
                "--reports-to", parsed["reports_to"],
                "--headcount", str(parsed["headcount"]),
                "--urgency", parsed["urgency"],
                "--output", str(out_prefix),
                "--format", "all",
            ]
        elif intent == "compensation":
            out_name = f"prajna_薪酬体系_{parsed['company']}_{parsed['industry']}_{timestamp}.xlsx"
            out_path = HISTORY_DIR / out_name
            cmd = [
                sys.executable, str(meta["script"]),
                "--company", parsed["company"],
                "--industry", parsed["industry"],
                "--city", parsed["city"],
                "--scale", str(parsed["scale"]),
                "--budget", str(parsed["budget"]),
                "--growth", str(parsed["growth"]),
                "--output", str(out_path),
            ]
        elif intent == "performance":
            safe_pos = re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9_-]", "_", parsed["position"])[:10]
            out_prefix = HISTORY_DIR / f"prajna_绩效体系_{parsed['company']}_{safe_pos}_{timestamp}"
            out_path = Path(f"{out_prefix}.xlsx")
            cmd = [
                sys.executable, str(meta["script"]),
                "--department", parsed["department"],
                "--position", parsed["position"],
                "--cycle", parsed["cycle"],
                "--method", parsed["method"],
                "--levels", parsed["levels"],
                "--company", parsed["company"],
                "--output", str(out_prefix),
                "--format", "all",
            ]

        with st.spinner(f"Prajna 正在生成 {meta['name']}..."):
            try:
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    check=True,
                    timeout=90,
                )
            except subprocess.CalledProcessError as e:
                st.error("生成失败，请检查输入或查看日志。")
                st.code(e.stderr or e.stdout)
                st.stop()
            except Exception as e:
                st.error(f"运行异常：{e}")
                st.stop()

        # For bidding/recruitment/performance, also offer the Word file
        extra_files = []
        if intent in ("bidding", "recruitment", "performance"):
            word_path = out_path.with_suffix(".docx")
            if word_path.exists():
                extra_files.append(word_path)

        if out_path.exists():
            file_size = out_path.stat().st_size
            st.success(f"✅ {meta['name']} 已生成！（{file_size / 1024:.1f} KB）")
            cols = st.columns([1] + [1] * len(extra_files))
            with cols[0]:
                with open(out_path, "rb") as f:
                    st.download_button(
                        label=f"📥 下载 {out_name}",
                        data=f,
                        file_name=out_name,
                        mime=meta["mime"],
                    )
            for idx, extra in enumerate(extra_files, start=1):
                with cols[idx]:
                    with open(extra, "rb") as f:
                        st.download_button(
                            label=f"📥 下载 {extra.name}",
                            data=f,
                            file_name=extra.name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
            st.info(meta["desc"])
        else:
            st.error("生成后未找到文件。")
            st.code(result.stdout)

# ---------------------------------------------------------------------------
# Manual skill mode
# ---------------------------------------------------------------------------
else:
    st.markdown("<br>", unsafe_allow_html=True)

    skill_options = {
        f"{meta['icon']} {meta['name']}": key
        for key, meta in SKILL_REGISTRY.items()
    }
    selected_label = st.selectbox("选择技能", list(skill_options.keys()))
    selected_skill = skill_options[selected_label]
    meta = SKILL_REGISTRY[selected_skill]

    left, right = st.columns([2, 1])

    with left:
        st.markdown(f"#### 2️⃣ 输入 {meta['name']} 参数")

        params = {}

        if selected_skill == "salary":
            c1, c2 = st.columns(2)
            with c1:
                params["industry"] = st.text_input("行业", value="互联网")
                params["position"] = st.text_input("岗位", value="电商运营助理")
                params["city"] = st.text_input("城市", value="广州")
            with c2:
                params["level"] = st.selectbox("职级", ["P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9"], index=1)
                params["description"] = st.text_area("自然语言描述（可选）", value="", height=80)

        elif selected_skill == "sales":
            c1, c2 = st.columns(2)
            with c1:
                params["team"] = st.text_input("团队名称", value="华南销售一部")
                params["preset"] = st.selectbox("团队类型", ["互联网/SaaS 销售团队", "电商销售团队", "线下零售销售团队", "通用销售团队"])
                params["week"] = st.text_input("报表周期", value=f"{datetime.now().year}年第{datetime.now().isocalendar().week}周")
            with c2:
                params["sales_target"] = st.number_input("本周销售目标（元）", min_value=0, value=1200000, step=10000)
                params["author"] = st.text_input("填写人", value="销售主管")
                params["date"] = st.date_input("填写日期", value=datetime.now()).strftime("%Y-%m-%d")

        elif selected_skill == "finance_kb":
            params["company"] = st.text_input("企业名称", value="智云电商")
            params["author"] = st.text_input("维护部门", value="财务部")
            params["date"] = st.date_input("更新日期", value=datetime.now()).strftime("%Y-%m-%d")

        elif selected_skill == "clothing_duty":
            c1, c2 = st.columns(2)
            with c1:
                params["preset"] = st.selectbox("班组预设", ["服装缝纫组", "服装裁剪组", "服装包装组", "通用"])
                params["team"] = st.text_input("班组名称", value="缝纫一组")
            with c2:
                params["factory"] = st.text_input("工厂名称", value="成衣A厂")
                params["month"] = st.text_input("考核周期", value=datetime.now().strftime("%Y-%m"))
                params["author"] = st.text_input("编制人", value="Prajna")

        elif selected_skill == "ai_daily":
            params["date"] = st.date_input("日报日期", value=datetime.now()).strftime("%Y-%m-%d")

        elif selected_skill == "finance_dashboard":
            c1, c2 = st.columns(2)
            with c1:
                params["preset"] = st.selectbox("企业预设", ["通用企业", "通用制造集团", "互联网/SaaS企业", "零售连锁企业", "新能源科技企业"])
                params["company"] = st.text_input("企业名称", value="示范企业股份")
            with c2:
                params["period"] = st.text_input("报表周期", value=datetime.now().strftime("%Y年%m月"))

        elif selected_skill == "budget_ppt":
            params["month"] = st.text_input("汇报月份", value=datetime.now().strftime("%Y年%m月"))
            params["company"] = st.text_input("公司名称", value="启明星科技")
            params["author"] = st.text_input("汇报部门", value="财务部")

        elif selected_skill == "bidding":
            c1, c2 = st.columns(2)
            with c1:
                params["project"] = st.text_input("项目名称", value="智慧园区智能化建设项目")
                params["bidder"] = st.text_input("投标人", value="智讯科技有限公司")
                params["tenderer"] = st.text_input("招标人", value="广州高新区管委会")
            with c2:
                params["amount"] = st.number_input("投标总金额（元）", min_value=0, value=5800000, step=100000)
                params["duration"] = st.text_input("工期/服务期", value="180天")
                params["industry"] = st.text_input("行业领域", value="IT")

        elif selected_skill == "recruitment":
            c1, c2 = st.columns(2)
            with c1:
                params["position"] = st.text_input("岗位名称", value="电商运营助理")
                params["department"] = st.text_input("所属部门", value="电商运营部")
                params["city"] = st.text_input("工作城市", value="广州")
                params["level"] = st.selectbox("职级", ["P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9"], index=1)
            with c2:
                params["salary_min"] = st.number_input("月薪下限", min_value=0, value=6000, step=500)
                params["salary_max"] = st.number_input("月薪上限", min_value=0, value=9000, step=500)
                params["reports_to"] = st.text_input("汇报对象", value="部门经理")
                params["headcount"] = st.number_input("招聘人数", min_value=1, value=1, step=1)
                params["urgency"] = st.selectbox("紧急程度", ["高", "中", "低"])

        elif selected_skill == "compensation":
            c1, c2 = st.columns(2)
            with c1:
                params["company"] = st.text_input("企业名称", value="智云科技")
                params["industry"] = st.text_input("行业", value="互联网")
                params["city"] = st.text_input("总部城市", value="广州")
                params["scale"] = st.number_input("企业规模（人）", min_value=1, value=200, step=10)
            with c2:
                params["budget"] = st.number_input("年度薪酬总预算（元）", min_value=0, value=30000000, step=1000000)
                params["growth"] = st.number_input("年度普调幅度（%）", min_value=0.0, value=5.0, step=0.5)

        elif selected_skill == "performance":
            c1, c2 = st.columns(2)
            with c1:
                params["department"] = st.text_input("考核部门", value="电商运营部")
                params["position"] = st.text_input("考核岗位", value="电商运营助理")
                params["cycle"] = st.selectbox("考核周期", ["月度", "季度", "半年度", "年度"])
            with c2:
                params["method"] = st.selectbox("考核方法", ["KPI", "OKR", "360", "MBO"])
                params["levels"] = st.text_input("绩效等级", value="A/B/C/D")
                params["company"] = st.text_input("企业名称", value="Prajna示范企业")

    with right:
        st.markdown('<div class="preview-card">', unsafe_allow_html=True)
        st.markdown(f"#### 🔍 生成预览")
        st.markdown(f"**技能：** {meta['icon']} {meta['name']}")
        st.markdown(f"**输出格式：** .{meta['ext']}")
        st.markdown(f"**说明：** {meta['desc']}")
        st.markdown("</div>", unsafe_allow_html=True)

    generate_clicked = st.button(f"🚀 生成 {meta['name']}", type="primary", width="stretch")

    if generate_clicked:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        HISTORY_DIR.mkdir(parents=True, exist_ok=True)

        if selected_skill == "salary":
            out_name = f"prajna_薪资模板_{params['industry']}_{params['position']}_{timestamp}.xlsx"
            out_path = HISTORY_DIR / out_name
            cmd = [sys.executable, str(meta["script"]), "--industry", params["industry"], "--position", params["position"], "--city", params["city"], "--level", params["level"], "--output", str(out_path)]
            if params["description"].strip():
                cmd += ["--description", params["description"].strip()]
        elif selected_skill == "sales":
            safe_team = params["team"].replace(" ", "_")
            out_name = f"prajna_销售周报_{safe_team}_{timestamp}.xlsx"
            out_path = HISTORY_DIR / out_name
            cmd = [sys.executable, str(meta["script"]), "--preset", params["preset"], "--team", params["team"], "--week", params["week"], "--sales-target", str(params["sales_target"]), "--author", params["author"], "--date", params["date"], "--output", str(out_path)]
        elif selected_skill == "finance_kb":
            out_name = f"prajna_财务知识库目录_{params['company']}_{timestamp}.md"
            out_path = HISTORY_DIR / out_name
            cmd = [sys.executable, str(meta["script"]), "--company", params["company"], "--author", params["author"], "--date", params["date"], "--output", str(out_path)]
        elif selected_skill == "clothing_duty":
            out_name = f"prajna_服装厂小组长岗位职责_{params['team']}_{timestamp}.xlsx"
            out_path = HISTORY_DIR / out_name
            cmd = [sys.executable, str(meta["script"]), "--preset", params["preset"], "--team", params["team"], "--factory", params["factory"], "--month", params["month"], "--author", params["author"], "--output", str(out_path)]
        elif selected_skill == "ai_daily":
            out_name = f"prajna_AI领袖日报_{params['date']}_{timestamp}.html"
            out_path = HISTORY_DIR / out_name
            cmd = [sys.executable, str(meta["script"]), "--date", params["date"], "--output", str(out_path)]
        elif selected_skill == "finance_dashboard":
            out_name = f"prajna_财务核心指标看板_{params['company']}_{timestamp}.html"
            out_path = HISTORY_DIR / out_name
            cmd = [sys.executable, str(meta["script"]), "--preset", params["preset"], "--company", params["company"], "--period", params["period"], "--output", str(out_path)]
        elif selected_skill == "budget_ppt":
            out_name = f"prajna_预算执行汇报_{params['company']}_{timestamp}.pptx"
            out_path = HISTORY_DIR / out_name
            cmd = [sys.executable, str(meta["script"]), "--month", params["month"], "--company", params["company"], "--author", params["author"], "--output", str(out_path)]
        elif selected_skill == "bidding":
            safe_project = re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9_-]", "_", params["project"])[:20]
            out_prefix = HISTORY_DIR / f"prajna_招投标套件_{safe_project}_{timestamp}"
            out_path = Path(f"{out_prefix}.xlsx")
            cmd = [sys.executable, str(meta["script"]), "--project", params["project"], "--bidder", params["bidder"], "--tenderer", params["tenderer"], "--amount", str(params["amount"]), "--duration", params["duration"], "--industry", params["industry"], "--output", str(out_prefix), "--format", "all"]
        elif selected_skill == "recruitment":
            safe_pos = re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9_-]", "_", params["position"])[:10]
            out_prefix = HISTORY_DIR / f"prajna_招聘套件_{safe_pos}_{params['city']}_{timestamp}"
            out_path = Path(f"{out_prefix}.xlsx")
            cmd = [sys.executable, str(meta["script"]), "--position", params["position"], "--department", params["department"], "--city", params["city"], "--level", params["level"], "--salary-min", str(params["salary_min"]), "--salary-max", str(params["salary_max"]), "--reports-to", params["reports_to"], "--headcount", str(params["headcount"]), "--urgency", params["urgency"], "--output", str(out_prefix), "--format", "all"]
        elif selected_skill == "compensation":
            out_name = f"prajna_薪酬体系_{params['company']}_{params['industry']}_{timestamp}.xlsx"
            out_path = HISTORY_DIR / out_name
            cmd = [sys.executable, str(meta["script"]), "--company", params["company"], "--industry", params["industry"], "--city", params["city"], "--scale", str(params["scale"]), "--budget", str(params["budget"]), "--growth", str(params["growth"]), "--output", str(out_path)]
        elif selected_skill == "performance":
            safe_pos = re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9_-]", "_", params["position"])[:10]
            out_prefix = HISTORY_DIR / f"prajna_绩效体系_{params['company']}_{safe_pos}_{timestamp}"
            out_path = Path(f"{out_prefix}.xlsx")
            cmd = [sys.executable, str(meta["script"]), "--department", params["department"], "--position", params["position"], "--cycle", params["cycle"], "--method", params["method"], "--levels", params["levels"], "--company", params["company"], "--output", str(out_prefix), "--format", "all"]

        with st.spinner(f"Prajna 正在生成 {meta['name']}..."):
            try:
                result = subprocess.run(cmd, capture_output=True, text=True, check=True, timeout=90)
            except subprocess.CalledProcessError as e:
                st.error("生成失败，请检查输入或查看日志。")
                st.code(e.stderr or e.stdout)
                st.stop()
            except Exception as e:
                st.error(f"运行异常：{e}")
                st.stop()

        extra_files = []
        if selected_skill in ("bidding", "recruitment", "performance"):
            word_path = out_path.with_suffix(".docx")
            if word_path.exists():
                extra_files.append(word_path)

        if out_path.exists():
            file_size = out_path.stat().st_size
            st.success(f"✅ {meta['name']} 已生成！（{file_size / 1024:.1f} KB）")
            cols = st.columns([1] + [1] * len(extra_files))
            with cols[0]:
                with open(out_path, "rb") as f:
                    st.download_button(label=f"📥 下载 {out_name}", data=f, file_name=out_name, mime=meta["mime"])
            for idx, extra in enumerate(extra_files, start=1):
                with cols[idx]:
                    with open(extra, "rb") as f:
                        st.download_button(label=f"📥 下载 {extra.name}", data=f, file_name=extra.name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.info(meta["desc"])
        else:
            st.error("生成后未找到文件。")
            st.code(result.stdout)

# ---------------------------------------------------------------------------
# Recent files
# ---------------------------------------------------------------------------
st.divider()
with st.expander("🕘 最近生成的文件（本次会话）"):
    all_files = sorted(HISTORY_DIR.glob("*"), key=lambda p: p.stat().st_mtime, reverse=True)[:15]
    if not all_files:
        st.info("暂无最近生成的文件。")
    else:
        recent = []
        for f in all_files:
            recent.append({
                "文件名": f.name,
                "大小(KB)": f"{f.stat().st_size / 1024:.1f}",
                "时间": datetime.fromtimestamp(f.stat().st_mtime).strftime("%m-%d %H:%M"),
            })
        st.dataframe(recent, width="stretch", hide_index=True)

# ---------------------------------------------------------------------------
# What it demonstrates
# ---------------------------------------------------------------------------
st.divider()
st.subheader("🌟 这个 Demo 展示了什么？")
feature_cols = st.columns(3)
features = [
    ("🤖", "自然语言智能体", "输入一句话，Prajna 自动识别意图并调用对应技能"),
    ("🧩", "多场景模板中台", "11 个企业技能，覆盖 HR、销售、财务、生产、招投标"),
    ("🎯", "多岗位智能适配", "内置 10+ 行业预设与城市薪酬系数"),
    ("🌍", "城市薪酬基准", "根据城市等级自动调整薪资与社保公积金基数"),
    ("📈", "职级带宽体系", "P1-P9 职级对应薪酬、绩效与奖金带宽"),
    ("🔗", "Excel 公式联动", "所有汇总、绩效、个税、完成率均保留公式"),
]
for i, (icon, title, desc) in enumerate(features):
    with feature_cols[i % 3]:
        st.markdown(
            f'<div class="feature-card"><h4>{icon} {title}</h4><p>{desc}</p></div>',
            unsafe_allow_html=True,
        )

# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.markdown(
    '<div class="footer">Prajna 企业智能体 · 广州「广智能」超级智能体大赛 · 开放赛道</div>',
    unsafe_allow_html=True,
)
st.caption(
    "【人工智能生成-需人工核验】本 Demo 输出内容仅供参考，具体薪酬、销售、财务、招投标、招聘等决策请以当地法律法规及公司政策为准。"
)
