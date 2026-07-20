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
    page_title="Prajna 企业智能体平台",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# Custom CSS - Modern, clean, card-based design
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    :root {
        --primary: #2563eb;
        --primary-dark: #1d4ed8;
        --secondary: #7c3aed;
        --accent: #06b6d4;
        --success: #10b981;
        --warning: #f59e0b;
        --danger: #ef4444;
        --bg: #f8fafc;
        --card: #ffffff;
        --text: #0f172a;
        --text-secondary: #475569;
        --border: #e2e8f0;
    }

    .stApp {
        background: linear-gradient(135deg, #f8fafc 0%, #eef2ff 50%, #f0f9ff 100%);
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif;
    }

    /* Hero Section */
    .hero-container {
        background: linear-gradient(135deg, #1e3a8a 0%, #3730a3 50%, #581c87 100%);
        border-radius: 24px;
        padding: 3.5rem 2.5rem;
        margin: 1rem 0 2rem 0;
        position: relative;
        overflow: hidden;
        box-shadow: 0 25px 50px -12px rgba(30, 58, 138, 0.25);
    }
    .hero-container::before {
        content: '';
        position: absolute;
        top: -50%;
        right: -10%;
        width: 500px;
        height: 500px;
        background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, transparent 70%);
        border-radius: 50%;
    }
    .hero-title {
        color: white;
        font-size: 3rem;
        font-weight: 700;
        margin: 0 0 1rem 0;
        letter-spacing: -0.02em;
    }
    .hero-subtitle {
        color: rgba(255,255,255,0.9);
        font-size: 1.25rem;
        margin: 0 0 2rem 0;
        max-width: 700px;
        line-height: 1.6;
    }
    .hero-badge {
        display: inline-flex;
        align-items: center;
        gap: 0.5rem;
        background: rgba(255,255,255,0.15);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(255,255,255,0.2);
        color: white;
        padding: 0.5rem 1rem;
        border-radius: 9999px;
        font-size: 0.875rem;
        font-weight: 500;
        margin-right: 0.75rem;
        margin-bottom: 0.75rem;
    }

    /* Section Headers */
    .section-title {
        font-size: 1.75rem;
        font-weight: 700;
        color: var(--text);
        margin: 2rem 0 1rem 0;
    }
    .section-subtitle {
        color: var(--text-secondary);
        font-size: 1rem;
        margin-bottom: 1.5rem;
        line-height: 1.6;
    }

    /* Cards */
    .capability-card {
        background: white;
        border-radius: 16px;
        padding: 1.75rem;
        border: 1px solid var(--border);
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05), 0 2px 4px -1px rgba(0,0,0,0.03);
        transition: all 0.2s ease;
        height: 100%;
    }
    .capability-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 20px 25px -5px rgba(0,0,0,0.08), 0 10px 10px -5px rgba(0,0,0,0.02);
        border-color: #c7d2fe;
    }
    .capability-icon {
        font-size: 2.5rem;
        margin-bottom: 1rem;
    }
    .capability-title {
        font-size: 1.25rem;
        font-weight: 700;
        color: var(--text);
        margin: 0 0 0.5rem 0;
    }
    .capability-desc {
        color: var(--text-secondary);
        font-size: 0.95rem;
        line-height: 1.6;
        margin: 0;
    }

    /* Skill Cards */
    .skill-card {
        background: white;
        border-radius: 12px;
        padding: 1.25rem;
        border: 2px solid transparent;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        transition: all 0.2s ease;
        cursor: pointer;
        height: 100%;
    }
    .skill-card:hover {
        border-color: var(--primary);
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0,0,0,0.08);
    }
    .skill-card-active {
        border-color: var(--primary);
        background: #eff6ff;
    }
    .skill-icon {
        font-size: 2rem;
        margin-bottom: 0.75rem;
    }
    .skill-name {
        font-size: 1.1rem;
        font-weight: 600;
        color: var(--text);
        margin: 0 0 0.25rem 0;
    }
    .skill-desc {
        color: var(--text-secondary);
        font-size: 0.85rem;
        line-height: 1.5;
        margin: 0;
    }

    /* Agent Flow */
    .flow-step {
        background: white;
        border-radius: 12px;
        padding: 1.25rem;
        border-left: 4px solid var(--primary);
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 0.75rem;
    }
    .flow-step-title {
        font-weight: 600;
        color: var(--text);
        margin-bottom: 0.25rem;
    }
    .flow-step-desc {
        color: var(--text-secondary);
        font-size: 0.9rem;
        line-height: 1.5;
    }
    .flow-arrow {
        text-align: center;
        color: var(--primary);
        font-size: 1.5rem;
        margin: 0.25rem 0;
    }

    /* Architecture Diagram */
    .arch-layer {
        border-radius: 12px;
        padding: 1.25rem;
        margin-bottom: 0.75rem;
        color: white;
        position: relative;
    }
    .arch-layer-1 { background: linear-gradient(90deg, #1e40af 0%, #3b82f6 100%); }
    .arch-layer-2 { background: linear-gradient(90deg, #5b21b6 0%, #8b5cf6 100%); }
    .arch-layer-3 { background: linear-gradient(90deg, #0e7490 0%, #06b6d4 100%); }
    .arch-layer-4 { background: linear-gradient(90deg, #047857 0%, #10b981 100%); }
    .arch-title {
        font-weight: 700;
        font-size: 1.1rem;
        margin-bottom: 0.5rem;
    }
    .arch-items {
        font-size: 0.9rem;
        opacity: 0.95;
        line-height: 1.5;
    }

    /* Memory Modules */
    .memory-module {
        background: white;
        border-radius: 12px;
        padding: 1.25rem;
        border-top: 4px solid var(--secondary);
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        height: 100%;
    }
    .memory-module-icon {
        font-size: 1.75rem;
        margin-bottom: 0.5rem;
    }
    .memory-module-title {
        font-weight: 700;
        font-size: 1.05rem;
        color: var(--text);
        margin-bottom: 0.5rem;
    }
    .memory-module-desc {
        color: var(--text-secondary);
        font-size: 0.85rem;
        line-height: 1.5;
    }

    /* Metrics */
    .metric-card {
        background: white;
        border-radius: 12px;
        padding: 1.25rem;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        border-bottom: 4px solid var(--primary);
    }
    .metric-value {
        font-size: 1.75rem;
        font-weight: 700;
        color: var(--primary);
        margin-bottom: 0.25rem;
    }
    .metric-label {
        color: var(--text-secondary);
        font-size: 0.85rem;
    }

    /* Input Area */
    .agent-input-container {
        background: white;
        border-radius: 16px;
        padding: 1.5rem;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
        border: 1px solid var(--border);
    }

    /* Result Area */
    .result-card {
        background: white;
        border-radius: 12px;
        padding: 1.25rem;
        border-left: 4px solid var(--success);
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-top: 1rem;
    }

    /* Footer */
    .footer {
        text-align: center;
        color: #94a3b8;
        font-size: 0.85rem;
        margin-top: 4rem;
        padding-top: 2rem;
        border-top: 1px solid var(--border);
    }

    /* Streamlit overrides */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem;
        background: white;
        padding: 0.5rem;
        border-radius: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .stTabs [data-baseweb="tab"] {
        padding: 0.75rem 1.25rem;
        border-radius: 8px;
        font-weight: 500;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%) !important;
        color: white !important;
    }
    .stButton>button {
        border-radius: 10px;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
        color: white;
        border: none;
        box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2);
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #1d4ed8 0%, #6d28d9 100%);
        box-shadow: 0 8px 12px -2px rgba(37, 99, 235, 0.3);
    }
    .stButton>button:active {
        transform: translateY(0);
    }
    div[data-testid="stExpander"] {
        background: white;
        border-radius: 12px;
        border: 1px solid var(--border);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Skill registry with categories
# ---------------------------------------------------------------------------
SKILL_REGISTRY = {
    "salary": {
        "icon": "💰",
        "name": "薪资模板",
        "category": "人力资源",
        "script": SALARY_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成包含 8 个工作表的完整薪资模板，含公式联动",
    },
    "recruitment": {
        "icon": "🤝",
        "name": "招聘助手",
        "category": "人力资源",
        "script": RECRUIT_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成 JD、面试评估表、Offer 模板等招聘套件",
    },
    "compensation": {
        "icon": "💵",
        "name": "薪酬体系",
        "category": "人力资源",
        "script": COMPENSATION_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成岗位薪酬矩阵、薪酬结构与预算",
    },
    "performance": {
        "icon": "⭐",
        "name": "绩效体系",
        "category": "人力资源",
        "script": PERFORMANCE_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成 KPI 指标库、绩效合同、评分表与 PIP",
    },
    "sales": {
        "icon": "📊",
        "name": "销售周报",
        "category": "销售商务",
        "script": SALES_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成包含 5 个工作表的销售团队标准周报",
    },
    "bidding": {
        "icon": "🎯",
        "name": "招投标助手",
        "category": "销售商务",
        "script": BIDDING_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成招投标套件与投标文件 Word 大纲",
    },
    "finance_kb": {
        "icon": "📁",
        "name": "财务知识库目录",
        "category": "财务经营",
        "script": FINANCE_KB_SCRIPT,
        "ext": "md",
        "mime": "text/markdown",
        "desc": "生成电商企业财务知识库钉钉文档目录",
    },
    "finance_dashboard": {
        "icon": "📈",
        "name": "财务核心指标看板",
        "category": "财务经营",
        "script": FIN_DASHBOARD_SCRIPT,
        "ext": "html",
        "mime": "text/html",
        "desc": "生成含 Chart.js 交互图表的财务看板",
    },
    "budget_ppt": {
        "icon": "📑",
        "name": "预算执行汇报 PPT",
        "category": "财务经营",
        "script": BUDGET_PPT_SCRIPT,
        "ext": "pptx",
        "mime": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "desc": "生成本月预算执行情况汇报 PPT",
    },
    "clothing_duty": {
        "icon": "🏭",
        "name": "服装厂岗位职责",
        "category": "生产运营",
        "script": CLOTHING_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成服装厂小组长岗位职责与 KPI 考核表",
    },
    "ai_daily": {
        "icon": "📰",
        "name": "AI 领袖日报",
        "category": "情报资讯",
        "script": AI_DAILY_SCRIPT,
        "ext": "html",
        "mime": "text/html",
        "desc": "生成 AI 领袖动态日报 HTML 页面",
    },
}

CATEGORIES = {
    "人力资源": ["salary", "recruitment", "compensation", "performance"],
    "销售商务": ["sales", "bidding"],
    "财务经营": ["finance_kb", "finance_dashboard", "budget_ppt"],
    "生产运营": ["clothing_duty"],
    "情报资讯": ["ai_daily"],
}

# ---------------------------------------------------------------------------
# Natural language parser
# ---------------------------------------------------------------------------
CITIES = [
    "北京", "上海", "广州", "深圳", "杭州", "成都", "武汉", "西安", "南京", "苏州",
    "重庆", "天津", "长沙", "郑州", "东莞", "宁波", "佛山", "合肥", "青岛", "厦门",
    "无锡", "福州", "济南", "沈阳", "大连", "昆明", "南昌", "哈尔滨", "长春", "石家庄",
    "贵阳", "南宁", "温州", "珠海", "惠州", "中山", "南通", "绍兴", "烟台", "潍坊",
]

INDUSTRIES = {
    "互联网": ["互联网", "科技", "IT", "软件"],
    "制造业": ["制造业", "工厂", "工业"],
    "零售": ["零售", "连锁", "门店"],
    "金融": ["金融", "银行", "证券", "保险"],
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
    scores = {intent: sum(1 for kw in kws if kw.lower() in text_lower) for intent, kws in INTENT_KEYWORDS.items()}
    if "招聘" in text and scores.get("bidding", 0) == 1 and "招标" not in text:
        scores["bidding"] = 0
    best = max(scores, key=scores.get)
    if scores[best] == 0:
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
    result = {"intent": intent, "raw": text, "city": city, "level": level, "industry": industry}

    if intent == "salary":
        position = extract_position(text)
        if position in ["", "岗位"] and industry:
            position = f"{industry}专员"
        result.update({"position": position, "description": text})
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
        result.update({"date": datetime.now().strftime("%Y-%m-%d")})
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
# Execution helper
# ---------------------------------------------------------------------------
def run_skill(intent, parsed, meta):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)

    if intent == "salary":
        out_name = f"prajna_薪资模板_{parsed['industry']}_{parsed['position']}_{timestamp}.xlsx"
        out_path = HISTORY_DIR / out_name
        cmd = [sys.executable, str(meta["script"]), "--industry", parsed["industry"], "--position", parsed["position"], "--city", parsed["city"], "--level", parsed["level"], "--description", parsed["raw"], "--output", str(out_path)]
    elif intent == "sales":
        safe_team = parsed["team"].replace(" ", "_")
        out_name = f"prajna_销售周报_{safe_team}_{timestamp}.xlsx"
        out_path = HISTORY_DIR / out_name
        cmd = [sys.executable, str(meta["script"]), "--preset", parsed["preset"], "--team", parsed["team"], "--week", parsed["week"], "--sales-target", str(parsed["target"]), "--author", parsed["author"], "--date", parsed["date"], "--output", str(out_path)]
    elif intent == "finance_kb":
        out_name = f"prajna_财务知识库目录_{parsed['company']}_{timestamp}.md"
        out_path = HISTORY_DIR / out_name
        cmd = [sys.executable, str(meta["script"]), "--company", parsed["company"], "--author", parsed["author"], "--date", parsed["date"], "--output", str(out_path)]
    elif intent == "clothing_duty":
        out_name = f"prajna_服装厂小组长岗位职责_{parsed['team']}_{timestamp}.xlsx"
        out_path = HISTORY_DIR / out_name
        cmd = [sys.executable, str(meta["script"]), "--team", parsed["team"], "--factory", parsed["factory"], "--month", parsed["month"], "--author", parsed["author"], "--output", str(out_path)]
    elif intent == "ai_daily":
        out_name = f"prajna_AI领袖日报_{parsed['date']}_{timestamp}.html"
        out_path = HISTORY_DIR / out_name
        cmd = [sys.executable, str(meta["script"]), "--date", parsed["date"], "--output", str(out_path)]
    elif intent == "finance_dashboard":
        out_name = f"prajna_财务核心指标看板_{parsed['company']}_{timestamp}.html"
        out_path = HISTORY_DIR / out_name
        cmd = [sys.executable, str(meta["script"]), "--preset", parsed["preset"], "--company", parsed["company"], "--period", parsed["period"], "--output", str(out_path)]
    elif intent == "budget_ppt":
        out_name = f"prajna_预算执行汇报_{parsed['company']}_{timestamp}.pptx"
        out_path = HISTORY_DIR / out_name
        cmd = [sys.executable, str(meta["script"]), "--month", parsed["month"], "--company", parsed["company"], "--author", parsed["author"], "--output", str(out_path)]
    elif intent == "bidding":
        safe_project = re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9_-]", "_", parsed["project"])[:20]
        out_prefix = HISTORY_DIR / f"prajna_招投标套件_{safe_project}_{timestamp}"
        out_path = Path(f"{out_prefix}.xlsx")
        cmd = [sys.executable, str(meta["script"]), "--project", parsed["project"], "--bidder", parsed["bidder"], "--tenderer", parsed["tenderer"], "--amount", str(parsed["amount"]), "--duration", parsed["duration"], "--industry", parsed["industry"], "--output", str(out_prefix), "--format", "all"]
    elif intent == "recruitment":
        safe_pos = re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9_-]", "_", parsed["position"])[:10]
        out_prefix = HISTORY_DIR / f"prajna_招聘套件_{safe_pos}_{parsed['city']}_{timestamp}"
        out_path = Path(f"{out_prefix}.xlsx")
        cmd = [sys.executable, str(meta["script"]), "--position", parsed["position"], "--department", parsed["department"], "--city", parsed["city"], "--level", parsed["level"], "--salary-min", str(parsed["salary_min"]), "--salary-max", str(parsed["salary_max"]), "--reports-to", parsed["reports_to"], "--headcount", str(parsed["headcount"]), "--urgency", parsed["urgency"], "--output", str(out_prefix), "--format", "all"]
    elif intent == "compensation":
        out_name = f"prajna_薪酬体系_{parsed['company']}_{parsed['industry']}_{timestamp}.xlsx"
        out_path = HISTORY_DIR / out_name
        cmd = [sys.executable, str(meta["script"]), "--company", parsed["company"], "--industry", parsed["industry"], "--city", parsed["city"], "--scale", str(parsed["scale"]), "--budget", str(parsed["budget"]), "--growth", str(parsed["growth"]), "--output", str(out_path)]
    elif intent == "performance":
        safe_pos = re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9_-]", "_", parsed["position"])[:10]
        out_prefix = HISTORY_DIR / f"prajna_绩效体系_{parsed['company']}_{safe_pos}_{timestamp}"
        out_path = Path(f"{out_prefix}.xlsx")
        cmd = [sys.executable, str(meta["script"]), "--department", parsed["department"], "--position", parsed["position"], "--cycle", parsed["cycle"], "--method", parsed["method"], "--levels", parsed["levels"], "--company", parsed["company"], "--output", str(out_prefix), "--format", "all"]

    result = subprocess.run(cmd, capture_output=True, text=True, check=True, timeout=90)

    extra_files = []
    if intent in ("bidding", "recruitment", "performance"):
        word_path = out_path.with_suffix(".docx")
        if word_path.exists():
            extra_files.append(word_path)

    return out_path, out_name, extra_files, result


# ---------------------------------------------------------------------------
# Sidebar - minimal
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("🧠 Prajna")
    st.markdown("**企业智能体平台**")
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
# Hero Section
# ---------------------------------------------------------------------------
st.markdown(
    """
    <div class="hero-container">
        <div class="hero-title">🧠 Prajna 企业智能体平台</div>
        <div class="hero-subtitle">
            一句话生成企业级文档 · 原生 Agent 架构底座 · 全模态记忆核心
        </div>
        <div>
            <span class="hero-badge">🤖 自然语言智能体</span>
            <span class="hero-badge">🧠 原生 Agent 架构</span>
            <span class="hero-badge">💾 全模态记忆核心</span>
            <span class="hero-badge">📦 11 个企业场景</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Main Tabs
# ---------------------------------------------------------------------------
tab_home, tab_templates, tab_architecture, tab_agents = st.tabs([
    "🏠 首页",
    "🛠️ 企业模板中台",
    "🧠 Prajna 核心架构",
    "🤖 Agent 联动示例",
])

# ---------------------------------------------------------------------------
# Tab 1: Home
# ---------------------------------------------------------------------------
with tab_home:
    st.markdown(
        """
        <div class="section-title">核心能力</div>
        <div class="section-subtitle">Prajna 不仅是模板生成器，更是一套面向企业的原生智能体平台</div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(
            """
            <div class="capability-card">
                <div class="capability-icon">🤖</div>
                <div class="capability-title">自然语言智能体</div>
                <div class="capability-desc">输入一句话即可识别意图、匹配技能、填充参数并生成文档。支持薪资、招聘、绩效、招投标、财务看板等 11 个企业场景。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            """
            <div class="capability-card">
                <div class="capability-icon">🧠</div>
                <div class="capability-title">原生 Agent 架构</div>
                <div class="capability-desc">定义应用层、能力层、模型层、数据层四层架构，规范 Agent 生命周期、元数据、协同协议与白盒追溯，支撑企业级多智能体协同。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            """
            <div class="capability-card">
                <div class="capability-icon">💾</div>
                <div class="capability-title">全模态记忆核心</div>
                <div class="capability-desc">时间记忆、语义网络、智能剪枝、自我反思四大模块，为所有 Agent 提供统一记忆服务，实现跨会话、跨 Agent、跨模态长期记忆联动。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">🚀 快速体验</div>', unsafe_allow_html=True)

    agent_input = st.text_area(
        "告诉 Prajna 你要生成什么",
        value="帮我做一份深圳互联网产品经理 P5 的薪资模板",
        height=100,
        placeholder="例如：生成电商销售团队本周周报，目标 80 万 / 搭建智云电商财务知识库目录 / 帮我做一份智慧园区建设项目的投标书",
        label_visibility="collapsed",
    )

    st.markdown(
        """
        <small style="color:#64748b">💡 试试这样说：</small><br>
        <small style="color:#64748b">"帮我做一份上海制造业生产主管 P4 的薪资模板" · "生成招聘电商运营助理的 JD 和面试评估表" · "生成本月预算执行汇报 PPT"</small>
        """,
        unsafe_allow_html=True,
    )

    execute_clicked = st.button("🚀 Prajna 立即生成", type="primary", use_container_width=True)

    if execute_clicked:
        parsed = parse_agent_input(agent_input)
        if not parsed:
            st.error("无法识别输入内容，请尝试更清晰的描述。")
        else:
            intent = parsed["intent"]
            meta = SKILL_REGISTRY[intent]
            with st.spinner(f"Prajna 正在生成 {meta['name']}..."):
                try:
                    out_path, out_name, extra_files, result = run_skill(intent, parsed, meta)
                except subprocess.CalledProcessError as e:
                    st.error("生成失败，请检查输入或查看日志。")
                    st.code(e.stderr or e.stdout)
                    st.stop()
                except Exception as e:
                    st.error(f"运行异常：{e}")
                    st.stop()

            if out_path.exists():
                file_size = out_path.stat().st_size
                st.markdown(
                    f"""
                    <div class="result-card">
                        <h4>✅ {meta['name']} 已生成</h4>
                        <p>{meta['desc']}（{file_size / 1024:.1f} KB）</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                cols = st.columns([1] + [1] * len(extra_files))
                with cols[0]:
                    with open(out_path, "rb") as f:
                        st.download_button(label=f"📥 下载 {out_name}", data=f, file_name=out_name, mime=meta["mime"], use_container_width=True)
                for idx, extra in enumerate(extra_files, start=1):
                    with cols[idx]:
                        with open(extra, "rb") as f:
                            st.download_button(label=f"📥 下载 {extra.name}", data=f, file_name=extra.name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            else:
                st.error("生成后未找到文件。")
                st.code(result.stdout)

    # Recent files
    st.markdown("<br>", unsafe_allow_html=True)
    with st.expander("🕘 最近生成的文件"):
        all_files = sorted(HISTORY_DIR.glob("*"), key=lambda p: p.stat().st_mtime, reverse=True)[:12]
        if not all_files:
            st.info("暂无最近生成的文件。")
        else:
            recent = [{"文件名": f.name, "大小(KB)": f"{f.stat().st_size / 1024:.1f}", "时间": datetime.fromtimestamp(f.stat().st_mtime).strftime("%m-%d %H:%M")} for f in all_files]
            st.dataframe(recent, width="stretch", hide_index=True)

# ---------------------------------------------------------------------------
# Tab 2: Template Center
# ---------------------------------------------------------------------------
with tab_templates:
    st.markdown(
        """
        <div class="section-title">企业模板中台</div>
        <div class="section-subtitle">按业务域选择技能，输入参数即可生成专业文档</div>
        """,
        unsafe_allow_html=True,
    )

    category = st.selectbox("选择业务域", list(CATEGORIES.keys()))
    skills_in_cat = CATEGORIES[category]

    # Skill selection cards
    cols = st.columns(len(skills_in_cat))
    selected_skill = None
    for idx, skill_key in enumerate(skills_in_cat):
        meta = SKILL_REGISTRY[skill_key]
        with cols[idx]:
            active_class = "skill-card-active" if st.session_state.get("selected_skill") == skill_key else ""
            st.markdown(
                f"""
                <div class="skill-card {active_class}" id="skill-card-{skill_key}">
                    <div class="skill-icon">{meta['icon']}</div>
                    <div class="skill-name">{meta['name']}</div>
                    <div class="skill-desc">{meta['desc']}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button(f"选择 {meta['name']}", key=f"select_{skill_key}", use_container_width=True):
                st.session_state.selected_skill = skill_key
                st.rerun()

    selected_skill = st.session_state.get("selected_skill")
    if selected_skill and selected_skill not in skills_in_cat:
        selected_skill = None

    if not selected_skill:
        selected_skill = skills_in_cat[0]
        st.session_state.selected_skill = selected_skill

    meta = SKILL_REGISTRY[selected_skill]
    st.markdown(f'<br><div class="section-title">{meta["icon"]} {meta["name"]}</div>', unsafe_allow_html=True)

    left, right = st.columns([2, 1])

    with left:
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

        generate_clicked = st.button(f"🚀 生成 {meta['name']}", type="primary", use_container_width=True)

    with right:
        st.markdown(
            f"""
            <div class="capability-card" style="background: linear-gradient(135deg, #eff6ff 0%, #f5f3ff 100%);">
                <div class="capability-title">🔍 生成预览</div>
                <div class="capability-desc" style="margin-bottom:1rem">
                    <strong>技能：</strong>{meta['icon']} {meta['name']}<br>
                    <strong>业务域：</strong>{meta['category']}<br>
                    <strong>输出格式：</strong>.{meta['ext']}
                </div>
                <div class="capability-desc">
                    {meta['desc']}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

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
            st.markdown(
                f"""
                <div class="result-card">
                    <h4>✅ {meta['name']} 已生成</h4>
                    <p>{meta['desc']}（{file_size / 1024:.1f} KB）</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            cols = st.columns([1] + [1] * len(extra_files))
            with cols[0]:
                with open(out_path, "rb") as f:
                    st.download_button(label=f"📥 下载 {out_name}", data=f, file_name=out_name, mime=meta["mime"], use_container_width=True)
            for idx, extra in enumerate(extra_files, start=1):
                with cols[idx]:
                    with open(extra, "rb") as f:
                        st.download_button(label=f"📥 下载 {extra.name}", data=f, file_name=extra.name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        else:
            st.error("生成后未找到文件。")
            st.code(result.stdout)

# ---------------------------------------------------------------------------
# Tab 3: Architecture
# ---------------------------------------------------------------------------
with tab_architecture:
    st.markdown(
        """
        <div class="section-title">🧠 Prajna 原生 Agent 架构</div>
        <div class="section-subtitle">四层架构 + 五大原生能力，构建企业级智能体底座</div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="arch-layer arch-layer-1">
            <div class="arch-title">🏢 应用层 Application Layer</div>
            <div class="arch-items">多模态对话 · AI 员工 · 记忆库 · 定时任务 · 企业门户</div>
        </div>
        <div class="arch-layer arch-layer-2">
            <div class="arch-title">⚡ 能力层 Capability Layer</div>
            <div class="arch-items">自主规划执行 · 多智能体协同 · 工具集成 · 白盒可追溯</div>
        </div>
        <div class="arch-layer arch-layer-3">
            <div class="arch-title">🧬 模型层 Model Layer</div>
            <div class="arch-items">全模态记忆大模型 · 记忆迭代 · 小样本适配</div>
        </div>
        <div class="arch-layer arch-layer-4">
            <div class="arch-title">💾 数据层 Data Layer</div>
            <div class="arch-items">时间记忆 · 语义网络 · 企业知识库 · 业务数据</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">📊 记忆核心关键指标</div>', unsafe_allow_html=True)

    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown('<div class="metric-card"><div class="metric-value">&lt; 1%</div><div class="metric-label">遗忘率</div></div>', unsafe_allow_html=True)
    with m2:
        st.markdown('<div class="metric-card"><div class="metric-value">25×</div><div class="metric-label">Token 消耗降低</div></div>', unsafe_allow_html=True)
    with m3:
        st.markdown('<div class="metric-card"><div class="metric-value">&lt; 0.2%</div><div class="metric-label">幻觉率</div></div>', unsafe_allow_html=True)
    with m4:
        st.markdown('<div class="metric-card"><div class="metric-value">&lt; 200ms</div><div class="metric-label">记忆召回延迟</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">🧩 记忆核心四大模块</div>', unsafe_allow_html=True)

    mc1, mc2, mc3, mc4 = st.columns(4)
    with mc1:
        st.markdown(
            """
            <div class="memory-module">
                <div class="memory-module-icon">⏱️</div>
                <div class="memory-module-title">时间记忆</div>
                <div class="memory-module-desc">按时间轴记录事件、状态变化与决策点，支持时间范围查询与因果关系链维护。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with mc2:
        st.markdown(
            """
            <div class="memory-module">
                <div class="memory-module-icon">🕸️</div>
                <div class="memory-module-title">语义网络</div>
                <div class="memory-module-desc">构建实体-关系-属性图谱，支持语义相似度检索，跨文档跨会话关联实体。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with mc3:
        st.markdown(
            """
            <div class="memory-module">
                <div class="memory-module-icon">✂️</div>
                <div class="memory-module-title">智能剪枝</div>
                <div class="memory-module-desc">基于重要性、时效性、访问频率自动压缩冗余记忆，保留关键决策依据。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with mc4:
        st.markdown(
            """
            <div class="memory-module">
                <div class="memory-module-icon">🔄</div>
                <div class="memory-module-title">自我反思</div>
                <div class="memory-module-desc">Agent 执行任务后自动记录结果与反思，识别成功与失败模式并更新策略记忆。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">🔄 Agent 标准生命周期</div>', unsafe_allow_html=True)

    lifecycle_steps = [
        ("1️⃣ 注册 Register", "向 Agent 调度中心上报名称、能力、依赖与资源配额"),
        ("2️⃣ 监听 Listen", "接收来自调度中心的任务或事件"),
        ("3️⃣ 感知 Perceive", "调用 memory_recall 加载相关上下文"),
        ("4️⃣ 规划 Plan", "拆解任务、选择工具、评估风险"),
        ("5️⃣ 执行 Act", "调用工具、调用其他 Agent、处理数据"),
        ("6️⃣ 反思 Reflect", "调用 memory_reflect 记录结果与改进点"),
        ("7️⃣ 上报 Report", "返回结果、状态与后续建议"),
    ]
    for title, desc in lifecycle_steps:
        st.markdown(
            f"""
            <div class="flow-step">
                <div class="flow-step-title">{title}</div>
                <div class="flow-step-desc">{desc}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">📋 原生架构迁移 Checklist</div>', unsafe_allow_html=True)

    checklist_items = [
        "SKILL.md frontmatter 包含 native_capabilities 与 requires 声明",
        "任务开始前调用 memory_recall 加载相关上下文",
        "任务结束后调用 memory_reflect 或 memory_remember 沉淀经验",
        "禁止各 Agent 私自维护独立记忆副本",
        "关键决策记录 reasoning_trace，支持白盒追溯",
    ]
    for item in checklist_items:
        st.checkbox(item, value=True, disabled=True)

# ---------------------------------------------------------------------------
# Tab 4: Agent Collaboration
# ---------------------------------------------------------------------------
with tab_agents:
    st.markdown(
        """
        <div class="section-title">🤖 Agent 联动示例</div>
        <div class="section-subtitle">Prajna 原生 Agent 如何通过记忆核心实现跨 Agent 协同</div>
        """,
        unsafe_allow_html=True,
    )

    agent_example = st.selectbox("选择 Agent 联动场景", ["客服 Agent", "标书 Agent", "招聘 Agent"])

    if agent_example == "客服 Agent":
        st.markdown(
            """
            <div class="flow-step" style="border-left-color:#10b981;">
                <div class="flow-step-title">👤 用户发起咨询</div>
                <div class="flow-step-desc">"我的报销单被退回了，请问是什么原因？"</div>
            </div>
            <div class="flow-arrow">⬇️</div>
            <div class="flow-step">
                <div class="flow-step-title">🧠 记忆召回 memory_recall</div>
                <div class="flow-step-desc">加载 user 记忆（偏好正式风格）、session 记忆（当前对话主题）、business 记忆（报销规则：需提前 3 天申请、发票须为专票）。</div>
            </div>
            <div class="flow-arrow">⬇️</div>
            <div class="flow-step">
                <div class="flow-step-title">⚡ 客服 Agent 执行</div>
                <div class="flow-step-desc">结合记忆与工单系统，定位退回原因：缺少专票信息 → 生成回复与补正指引。</div>
            </div>
            <div class="flow-arrow">⬇️</div>
            <div class="flow-step">
                <div class="flow-step-title">💾 记忆沉淀 memory_remember</div>
                <div class="flow-step-desc">将本次问题与解决方案写入 business 记忆；用户负面情绪与满意度写入 feedback 记忆；未结工单写入 project 记忆。</div>
            </div>
            <div class="flow-arrow">⬇️</div>
            <div class="flow-step" style="border-left-color:#f59e0b;">
                <div class="flow-step-title">🔄 自我反思 memory_reflect</div>
                <div class="flow-step-desc">复盘同类问题高频原因，建议财务部门优化报销单填写引导，更新 business 规则记忆。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    elif agent_example == "标书 Agent":
        st.markdown(
            """
            <div class="flow-step" style="border-left-color:#10b981;">
                <div class="flow-step-title">📄 招标公告输入</div>
                <div class="flow-step-desc">"智慧园区智能化建设项目"招标公告发布，要求 180 天交付、500 万以上业绩、IT 资质。</div>
            </div>
            <div class="flow-arrow">⬇️</div>
            <div class="flow-step">
                <div class="flow-step-title">🧠 记忆召回 memory_recall</div>
                <div class="flow-step-desc">加载 business 记忆（历史资质库、案例库）、project 记忆（同类项目复盘经验）、agent 记忆（上次投标得失）。</div>
            </div>
            <div class="flow-arrow">⬇️</div>
            <div class="flow-step">
                <div class="flow-step-title">⚡ 标书 Agent 执行</div>
                <div class="flow-step-desc">自动拆解招标文件，生成资格自审表、技术偏离表、商务报价表与评分响应索引；协同财务 Agent 核对报价，法务 Agent 审核条款。</div>
            </div>
            <div class="flow-arrow">⬇️</div>
            <div class="flow-step">
                <div class="flow-step-title">💾 记忆沉淀 memory_remember</div>
                <div class="flow-step-desc">招标解析结果写入 project 记忆；投标文档版本与关键决策写入 agent 记忆；中标/废标结果写回 business 记忆。</div>
            </div>
            <div class="flow-arrow">⬇️</div>
            <div class="flow-step" style="border-left-color:#f59e0b;">
                <div class="flow-step-title">🔄 自我反思 memory_reflect</div>
                <div class="flow-step-desc">投标后调用 memory_reflect 复盘：报价策略、技术方案亮点、竞争对手动向，沉淀为下次投标的策略记忆。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    else:  # 招聘 Agent
        st.markdown(
            """
            <div class="flow-step" style="border-left-color:#10b981;">
                <div class="flow-step-title">📋 招聘需求输入</div>
                <div class="flow-step-desc">"电商运营部招聘电商运营助理 1 名，P2，月薪 6-9K，广州。"</div>
            </div>
            <div class="flow-arrow">⬇️</div>
            <div class="flow-step">
                <div class="flow-step-title">🧠 记忆召回 memory_recall</div>
                <div class="flow-step-desc">加载 business 记忆（岗位胜任力模型、薪酬带宽）、agent 记忆（同类岗位历史录用特征）、user 记忆（用人部门偏好）。</div>
            </div>
            <div class="flow-arrow">⬇️</div>
            <div class="flow-step">
                <div class="flow-step-title">⚡ 招聘 Agent 执行</div>
                <div class="flow-step-desc">生成 JD、结构化面试评估表、Offer 薪资建议；沉淀岗位画像与渠道效果数据。</div>
            </div>
            <div class="flow-arrow">⬇️</div>
            <div class="flow-step">
                <div class="flow-step-title">💾 记忆沉淀 memory_remember</div>
                <div class="flow-step-desc">面试评价写入 feedback 记忆；录用决策与候选人特征写入 project 记忆；岗位画像更新写入 business 记忆。</div>
            </div>
            <div class="flow-arrow">⬇️</div>
            <div class="flow-step" style="border-left-color:#f59e0b;">
                <div class="flow-step-title">🤝 入职 Agent 接管</div>
                <div class="flow-step-desc">Offer 发出后触发入职 Agent，读取招聘 Agent 沉淀的候选人信息，自动生成分配计划、培训计划与薪酬开通流程。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.markdown(
    """
    <div class="footer">
        <strong>Prajna 企业智能体平台</strong><br>
        广州「广智能」超级智能体大赛 · 开放赛道 · 自然语言智能体 · 原生 Agent 架构 · 全模态记忆核心
    </div>
    """,
    unsafe_allow_html=True,
)
st.caption(
    "【人工智能生成-需人工核验】本 Demo 输出内容仅供参考，具体薪酬、销售、财务、招投标、招聘等决策请以当地法律法规及公司政策为准。"
)
