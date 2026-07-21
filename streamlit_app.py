import io
import re
import subprocess
import sys
import zipfile
from datetime import datetime
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

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
PROCUREMENT_SCRIPT = SKILLS_DIR / "procurement" / "prajna-procurement-assistant" / "scripts" / "generate_procurement_kit.py"
LEGAL_SCRIPT = SKILLS_DIR / "legal" / "prajna-contract-review-assistant" / "scripts" / "generate_contract_review.py"
CS_SOP_SCRIPT = SKILLS_DIR / "customer-service" / "prajna-customer-service-sop" / "scripts" / "generate_cs_sop.py"
PRODUCTION_SCRIPT = SKILLS_DIR / "production" / "prajna-production-daily-report" / "scripts" / "generate_production_daily.py"

SALARY_SAMPLE = APP_DIR / "assets" / "sample_薪资模板_广州_电商运营助理.xlsx"
SALES_SAMPLE = APP_DIR / "assets" / "sample_销售周报_示例.xlsx"

HISTORY_DIR = Path.home() / ".prajna" / "demo_history"

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="prajna 企业智能体平台",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# Custom CSS - Modern, clean, card-based design
# ---------------------------------------------------------------------------
st.markdown(
    f"""<style>{(APP_DIR / "assets" / "style.css").read_text(encoding="utf-8")}</style>""",
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
        "desc": "生成 AI 领袖动态日报 HTML 页面（已接入真实网络数据）",
    },
    "procurement": {
        "icon": "🛒",
        "name": "采购管理套件",
        "category": "供应链",
        "script": PROCUREMENT_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成采购申请、供应商评估、询价比价、合同审查、采购台账",
    },
    "legal": {
        "icon": "⚖️",
        "name": "合同审查助手",
        "category": "法务合规",
        "script": LEGAL_SCRIPT,
        "ext": "docx",
        "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "desc": "生成合同审查意见书 Word 文档",
    },
    "cs_sop": {
        "icon": "🎧",
        "name": "客服 SOP",
        "category": "客户服务",
        "script": CS_SOP_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成客服话术库、工单流程、客诉升级规则、FAQ",
    },
    "production": {
        "icon": "🏭",
        "name": "生产日报",
        "category": "生产运营",
        "script": PRODUCTION_SCRIPT,
        "ext": "xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "desc": "生成生产日报、产量统计、设备运行、质量检验、人员出勤",
    },
}

CATEGORIES = {
    "人力资源": ["salary", "recruitment", "compensation", "performance"],
    "销售商务": ["sales", "bidding"],
    "财务经营": ["finance_kb", "finance_dashboard", "budget_ppt"],
    "供应链": ["procurement"],
    "法务合规": ["legal"],
    "客户服务": ["cs_sop"],
    "生产运营": ["clothing_duty", "production"],
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
    "salary": ["薪资", "工资", "salary", "收入", "待遇", "薪水"],
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
    "procurement": ["采购", "供应商", "询价比价", "采购申请", "采购合同", "采购台账"],
    "legal": ["合同审查", "法务", "合同审核", "法律", "合规", "contract review"],
    "cs_sop": ["客服", "客服话术", "工单", "客诉", "sop", "faq", "客服流程"],
    "production": ["生产日报", "产量统计", "设备运行", "质量检验", "人员出勤", "车间日报"],
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
            "author": "prajna",
            "date": datetime.now().strftime("%Y-%m-%d"),
        })
    elif intent == "finance_kb":
        company_clean = re.sub(r"^(帮我|给我|请|做|一份|的|生成|搭建|为|设计)", "", text)
        company_clean = re.sub(r"(电商企业财务知识库|财务知识库|财务目录|知识库目录|钉钉文档目录).*$", "", company_clean)
        for kw in ["电商", "企业", "财务", "知识库", "目录", "搭建", "生成", "的"]:
            company_clean = company_clean.replace(kw, "")
        result.update({
            "company": company_clean.strip()[:20] or "京东集团",
            "author": "财务部",
            "date": datetime.now().strftime("%Y-%m-%d"),
        })
    elif intent == "clothing_duty":
        result.update({
            "team": "缝纫一组",
            "factory": "成衣A厂",
            "month": datetime.now().strftime("%Y-%m"),
            "author": "prajna",
        })
    elif intent == "ai_daily":
        result.update({"date": datetime.now().strftime("%Y-%m-%d"), "search": False})
    elif intent == "procurement":
        company_clean = re.sub(r"^(帮我|给我|请|做|一份|的|生成|为|设计|采购管理|采购)", "", text)
        for kw in ["套件", "模板", "Excel", "excel", "生成", "搭建", "的", "为"]:
            company_clean = company_clean.replace(kw, "")
        result.update({
            "company": company_clean.strip()[:20] or "美的集团",
            "department": "采购部",
            "applicant": "张采购",
            "date": datetime.now().strftime("%Y-%m-%d"),
            "delivery_date": datetime.now().strftime("%Y-%m-%d"),
        })
    elif intent == "legal":
        company_clean = re.sub(r"^(帮我|给我|请|做|一份|的|生成|为|设计|审查|审核)", "", text)
        company_clean = re.sub(r"(合同审查意见书|合同审查|合同审核|意见书|合同).*$", "", company_clean)
        for kw in ["法务", "法律", "的", "生成", "搭建", "为", "公司", "企业"]:
            company_clean = company_clean.replace(kw, "")
        result.update({
            "company": company_clean.strip()[:20] or "京东集团",
            "contract_name": "软件采购与服务合同",
            "contract_type": "采购合同",
            "amount": 580000,
            "term": "12个月",
            "risk_level": "中",
            "reviewer": "法务经理",
            "date": datetime.now().strftime("%Y-%m-%d"),
        })
    elif intent == "cs_sop":
        company_clean = re.sub(r"^(帮我|给我|请|做|一份|的|生成|搭建|为|设计)", "", text)
        company_clean = re.sub(r"(客服SOP|客服话术|客服标准作业程序|客服流程|FAQ|工单流程).*$", "", company_clean)
        for kw in ["客服", "SOP", "sop", "话术", "的", "生成", "搭建", "为", "公司", "企业"]:
            company_clean = company_clean.replace(kw, "")
        result.update({
            "company": company_clean.strip()[:20] or "京东集团",
            "date": datetime.now().strftime("%Y-%m-%d"),
        })
    elif intent == "production":
        factory_clean = re.sub(r"^(帮我|给我|请|做|一份|的|生成|搭建|为|设计)", "", text)
        factory_clean = re.sub(r"(生产日报|产量统计|设备运行|质量检验|人员出勤|日报).*$", "", factory_clean)
        for kw in ["生产", "车间", "工厂", "的", "生成", "搭建", "为"]:
            factory_clean = factory_clean.replace(kw, "")
        result.update({
            "factory": factory_clean.strip()[:20] or "美的集团武汉工厂",
            "date": datetime.now().strftime("%Y-%m-%d"),
        })
    elif intent == "finance_dashboard":
        company_clean = re.sub(r"^(帮我|给我|请|做|一份|的|生成|搭建|为|设计)", "", text)
        company_clean = re.sub(r"(财务核心指标看板|财务看板|财务指标).*$", "", company_clean)
        for kw in ["资产负债率", "净利润率", "现金流量比率", "核心指标", "关注", "和", "、", "互联网", "企业", "公司"]:
            company_clean = company_clean.replace(kw, "")
        result.update({
            "preset": infer_finance_preset(text),
            "company": company_clean.strip()[:20] or "美的集团",
            "period": datetime.now().strftime("%Y年%m月"),
        })
    elif intent == "budget_ppt":
        company_clean = re.sub(r"^(帮我|给我|请|做|一份|的|生成|为|设计)", "", text)
        company_clean = re.sub(r"(本月预算执行情况汇报|预算执行情况汇报|预算执行汇报|月度预算汇报|PPT|ppt).*$", "", company_clean)
        for kw in ["生成", "本月", "预算", "执行", "情况", "汇报", "的", "为"]:
            company_clean = company_clean.replace(kw, "")
        result.update({
            "month": datetime.now().strftime("%Y年%m月"),
            "company": company_clean.strip()[:20] or "美的集团",
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
        company_clean = re.sub(r"^(帮我|给我|请|做|一份|的|生成|搭建|为|设计)", "", text)
        company_clean = re.sub(r"(企业|公司)?\s*(薪酬体系|薪酬矩阵|企业薪酬|compensation).*$", "", company_clean)
        for kw in ["设计", "互联网", "做", "一份", "的", "生成", "搭建", "为", "公司", "企业"]:
            company_clean = company_clean.replace(kw, "")
        result.update({
            "company": company_clean.strip()[:20] or "智云科技",
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
            "company": "prajna示范企业",
        })
    return result


# ---------------------------------------------------------------------------
# Execution helper
# ---------------------------------------------------------------------------
def run_skill(intent, parsed, meta):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    out_name = None

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
        if parsed.get("search"):
            cmd.append("--search")
    elif intent == "procurement":
        out_name = f"prajna_采购管理套件_{parsed['company']}_{timestamp}.xlsx"
        out_path = HISTORY_DIR / out_name
        cmd = [sys.executable, str(meta["script"]), "--company", parsed["company"], "--department", parsed["department"], "--applicant", parsed["applicant"], "--date", parsed["date"], "--delivery-date", parsed["delivery_date"], "--output", str(out_path)]
    elif intent == "legal":
        out_name = f"prajna_合同审查意见书_{parsed['company']}_{timestamp}.docx"
        out_path = HISTORY_DIR / out_name
        cmd = [sys.executable, str(meta["script"]), "--company", parsed["company"], "--contract-name", parsed["contract_name"], "--contract-type", parsed["contract_type"], "--amount", str(parsed["amount"]), "--term", parsed["term"], "--risk-level", parsed["risk_level"], "--reviewer", parsed["reviewer"], "--date", parsed["date"], "--output", str(out_path)]
    elif intent == "cs_sop":
        out_name = f"prajna_客服SOP_{parsed['company']}_{timestamp}.xlsx"
        out_path = HISTORY_DIR / out_name
        cmd = [sys.executable, str(meta["script"]), "--company", parsed["company"], "--date", parsed["date"], "--output", str(out_path)]
    elif intent == "production":
        out_name = f"prajna_生产日报_{parsed['factory']}_{timestamp}.xlsx"
        out_path = HISTORY_DIR / out_name
        cmd = [sys.executable, str(meta["script"]), "--factory", parsed["factory"], "--date", parsed["date"], "--output", str(out_path)]
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
        out_prefix = HISTORY_DIR / f"prajna_招聘套件_{safe_pos}_{parsed['city']}_{timestamp}_output"
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
        # These scripts create a directory from out_prefix and put files inside it
        output_dir = Path(str(out_prefix))
        if output_dir.exists() and output_dir.is_dir():
            excel_files = sorted(output_dir.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
            word_files = sorted(output_dir.glob("*.docx"), key=lambda p: p.stat().st_mtime, reverse=True)
            if excel_files:
                out_path = excel_files[0]
                out_name = out_path.name
            if word_files:
                extra_files.append(word_files[0])

    if out_name is None:
        out_name = out_path.name

    return out_path, out_name, extra_files, result


# ---------------------------------------------------------------------------
# Batch / scenario helpers
# ---------------------------------------------------------------------------
SAMPLE_INPUTS = {
    "salary": "帮我做一份广州互联网电商运营助理 P2 的薪资模板",
    "sales": "生成华南销售一部本周销售周报，目标 120 万",
    "finance_kb": "搭建京东集团财务知识库目录",
    "clothing_duty": "生成服装厂缝纫一组小组长岗位职责",
    "ai_daily": "生成今日ai领袖日报",
    "finance_dashboard": "搭建美的集团财务看板，关注资产负债率和净利润率",
    "budget_ppt": "为美的集团生成本月预算执行汇报 PPT",
    "bidding": "帮我做一份智慧园区建设项目的投标书",
    "recruitment": "帮我招聘一名广州 P2 电商运营助理",
    "compensation": "搭建智云科技企业薪酬体系",
    "performance": "生成电商运营助理岗位的绩效体系",
    "procurement": "生成美的集团采购管理套件",
    "legal": "为京东集团生成合同审查意见书",
    "cs_sop": "生成京东集团客服 SOP",
    "production": "生成美的集团武汉工厂生产日报",
}


def run_single_skill_demo(skill_key, timestamp):
    text = SAMPLE_INPUTS[skill_key]
    parsed = parse_agent_input(text)
    meta = SKILL_REGISTRY[skill_key]
    out_path, out_name, extra_files, result = run_skill(skill_key, parsed, meta)
    return {
        "key": skill_key,
        "name": meta["name"],
        "category": meta["category"],
        "main": out_path,
        "extras": extra_files,
        "ok": True,
    }


def create_zip_from_results(results, zip_name):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in results:
            for f in [r["main"]] + r["extras"]:
                if f and f.exists():
                    zf.write(f, arcname=f"{r['category']}/{r['name']}/{f.name}")
    zip_buffer.seek(0)
    return zip_buffer.getvalue()


SCENARIOS = {
    "招聘入职包": {
        "icon": "🤝",
        "desc": "招聘 → 薪资 → 薪酬体系 → 绩效体系，覆盖人才选用育留完整闭环",
        "skills": ["recruitment", "salary", "compensation", "performance"],
        "color": "#2563eb",
    },
    "投标商务包": {
        "icon": "🎯",
        "desc": "招投标套件 → 财务看板 → 预算 PPT，支撑销售与经营决策",
        "skills": ["bidding", "finance_dashboard", "budget_ppt"],
        "color": "#7c3aed",
    },
    "HR 管理包": {
        "icon": "💼",
        "desc": "薪资模板 → 薪酬体系 → 绩效体系 → 岗位职责，服务人力资源日常管理",
        "skills": ["salary", "compensation", "performance", "clothing_duty"],
        "color": "#059669",
    },
    "经营决策包": {
        "icon": "📈",
        "desc": "财务知识库 → 财务看板 → 预算 PPT → 销售周报，管理层月度经营参考",
        "skills": ["finance_kb", "finance_dashboard", "budget_ppt", "sales"],
        "color": "#d97706",
    },
    "供应链法务包": {
        "icon": "🛒",
        "desc": "采购管理套件 → 合同审查意见书 → 生产日报，覆盖供应链与生产合规",
        "skills": ["procurement", "legal", "production"],
        "color": "#0891b2",
    },
    "客户服务包": {
        "icon": "🎧",
        "desc": "客服 SOP → 生产日报 → 销售周报，打通服务与运营数据",
        "skills": ["cs_sop", "production", "sales"],
        "color": "#db2777",
    },
}


def run_scenario(scenario_key):
    cfg = SCENARIOS[scenario_key]
    results = []
    for skill_key in cfg["skills"]:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        r = run_single_skill_demo(skill_key, timestamp)
        results.append(r)
    return results


# ---------------------------------------------------------------------------
# Sidebar - minimal
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("🧠 prajna")
    st.markdown("**企业智能体平台**")
    st.divider()
    st.markdown(
        """
        ### 关于 prajna
        prajna 由锦辉人力自主研发，以 HRO 为首个深度落地场景，内置 32 位数字员工 Agent，面向中小企业通用 Agent 平台。

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
        <div style="display:inline-block;background:rgba(255,255,255,0.15);backdrop-filter:blur(10px);border:1px solid rgba(255,255,255,0.2);color:white;padding:0.35rem 0.9rem;border-radius:999px;font-size:0.8rem;font-weight:600;margin-bottom:1rem;">
            🏆 广州「广智能」超级智能体大赛 · 可运行成果
        </div>
        <div class="hero-title">🧠 prajna 企业智能体平台</div>
        <div class="hero-subtitle">
            人力带着业务走 · 人力资源做导演，贴近业务，引领变化<br>
            <span style="opacity:0.85;font-size:1.05rem;">先做自己的 0 号客户，再做行业的基础设施 · 32 位数字员工 Agent 覆盖 HR 核心、业务增长与管理决策全链路</span>
        </div>
        <div>
            <span class="hero-badge">🤖 32 位数字员工 Agent</span>
            <span class="hero-badge">🧠 原生 Agent 架构</span>
            <span class="hero-badge">💾 全模态记忆核心</span>
            <span class="hero-badge">📦 15 个企业场景</span>
            <span class="hero-badge">🔄 多 Agent 协同</span>
            <span class="hero-badge">🖥️ 自主桌面自动化</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Dynamic stats row
skill_count = len(SKILL_REGISTRY)
cat_count = len(CATEGORIES)
st.markdown(
    f"""
    <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:1rem;margin-bottom:2rem;">
        <div class="metric-card"><div class="metric-value">32</div><div class="metric-label">数字员工 Agent</div></div>
        <div class="metric-card"><div class="metric-value">{skill_count}</div><div class="metric-label">可扩展技能</div></div>
        <div class="metric-card"><div class="metric-value">{cat_count}</div><div class="metric-label">业务领域</div></div>
        <div class="metric-card"><div class="metric-value">20+</div><div class="metric-label">协作平台</div></div>
        <div class="metric-card"><div class="metric-value">0</div><div class="metric-label">外部 API 依赖</div></div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Main Tabs
# ---------------------------------------------------------------------------
tab_home, tab_templates, tab_architecture, tab_agents, tab_showcase = st.tabs([
    "🏠 首页",
    "🛠️ 企业模板中台",
    "🧠 prajna 核心架构",
    "🤖 Agent 联动示例",
    "🎁 prajna 全能力",
])

# ---------------------------------------------------------------------------
# Tab 1: Home
# ---------------------------------------------------------------------------
with tab_home:
    st.markdown(
        """
        <div class="section-title">关于 prajna</div>
        <div style="background:white;border-radius:20px;padding:2rem;border:1px solid #e2e8f0;box-shadow:0 8px 32px -8px rgba(15,23,42,0.08);margin-bottom:1.5rem;">
            <p style="margin:0 0 1rem 0;color:#334155;line-height:1.8;font-size:1.05rem;">
                prajna 由锦辉人力自主研发，是以 <strong>HRO 为首个深度落地场景</strong>、面向中小企业的通用 Agent 平台。
                我们最早想做 prajna，不是因为看到了"AI+HR"这个风口，而是因为我们自己就是那个每月被算薪表、合规排查、标书截止日期反复折磨的人。
                <strong>先给自己治病，再谈别的。</strong>
            </p>
            <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:1rem;">
                <div style="text-align:center;padding:1rem;background:#f8fafc;border-radius:12px;">
                    <div style="font-size:1.75rem;font-weight:800;color:#2563eb;">200+</div>
                    <div style="font-size:0.85rem;color:#64748b;">服务客户企业</div>
                </div>
                <div style="text-align:center;padding:1rem;background:#f8fafc;border-radius:12px;">
                    <div style="font-size:1.75rem;font-weight:800;color:#2563eb;">1万+</div>
                    <div style="font-size:0.85rem;color:#64748b;">派遣/外包员工</div>
                </div>
                <div style="text-align:center;padding:1rem;background:#f8fafc;border-radius:12px;">
                    <div style="font-size:1.75rem;font-weight:800;color:#2563eb;">20+年</div>
                    <div style="font-size:0.85rem;color:#64748b;">一线 HRO 运营经验</div>
                </div>
                <div style="text-align:center;padding:1rem;background:#f8fafc;border-radius:12px;">
                    <div style="font-size:1.75rem;font-weight:800;color:#2563eb;">AAA</div>
                    <div style="font-size:0.85rem;color:#64748b;">和谐劳动关系企业</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="section-title">核心能力</div>
        <div class="section-subtitle">prajna 不仅是模板生成器，更是一套面向企业的原生智能体平台 · 听得懂 · 干得了 · 记得住 · 可合规</div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            """
            <div class="capability-card">
                <div class="capability-icon">🗣️</div>
                <div class="capability-title">听得懂</div>
                <div class="capability-desc">自然语言交互，自动识别招聘、薪酬、合规、招投标等 24+ 业务意图，模糊目标主动澄清。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            """
            <div class="capability-card">
                <div class="capability-icon">🛠️</div>
                <div class="capability-title">干得了</div>
                <div class="capability-desc">47+ 工具、6 种执行后端，可调用终端、浏览器、Office、数据库真实操作系统，子 Agent 并行隔离执行。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            """
            <div class="capability-card">
                <div class="capability-icon">🧠</div>
                <div class="capability-title">记得住</div>
                <div class="capability-desc">持久记忆 + 用户建模，跨会话保持上下文。时间记忆、语义网络、智能剪枝、自我反思四大模块联动。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            """
            <div class="capability-card">
                <div class="capability-icon">🛡️</div>
                <div class="capability-title">可合规</div>
                <div class="capability-desc">AIGC 隐式水印、算法审计日志、危险命令审批，关键操作保留人工复核，满足企业合规要求。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        """
        <div class="section-title">旗舰 Agent 实测效果</div>
        <div class="section-subtitle">三个月封闭试点验证：薪酬核算效率提升 95%，招聘方案生成效率提升 98%</div>
        """,
        unsafe_allow_html=True,
    )

    f1, f2, f3 = st.columns(3)
    with f1:
        st.markdown(
            """
            <div class="capability-card" style="border-top:4px solid #2563eb;">
                <div class="capability-title">💰 薪酬 Agent</div>
                <div class="capability-desc">拉取考勤 → 标记异常 → 计算应发 → 计算个税 → 三级核验 → 生成报盘</div>
                <div style="margin-top:1rem;display:flex;align-items:baseline;gap:0.5rem;">
                    <span style="font-size:2rem;font-weight:800;color:#2563eb;">3天→30分钟</span>
                </div>
                <div style="font-size:0.85rem;color:#64748b;">单轮核算效率提升</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with f2:
        st.markdown(
            """
            <div class="capability-card" style="border-top:4px solid #10b981;">
                <div class="capability-title">🛡️ 合规 Agent</div>
                <div class="capability-desc">实时扫描全量客户用工风险，按高/中/低自动分级预警；法理思辨式咨询，引用法条给出"要做/不要做"对照表</div>
                <div style="margin-top:1rem;display:flex;align-items:baseline;gap:0.5rem;">
                    <span style="font-size:2rem;font-weight:800;color:#10b981;">事后→实时</span>
                </div>
                <div style="font-size:0.85rem;color:#64748b;">风险发现模式转变</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with f3:
        st.markdown(
            """
            <div class="capability-card" style="border-top:4px solid #7c3aed;">
                <div class="capability-title">🎯 标书 Agent</div>
                <div class="capability-desc">招标文件解析 → 评分项提取 → 结构规划 → AI 撰写 → 资质关联 → 格式规范化 → 导出</div>
                <div style="margin-top:1rem;display:flex;align-items:baseline;gap:0.5rem;">
                    <span style="font-size:2rem;font-weight:800;color:#7c3aed;">数天→数小时</span>
                </div>
                <div style="font-size:0.85rem;color:#64748b;">完整标书（8 章节 / 约 40KB）</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">🚀 快速体验</div>', unsafe_allow_html=True)

    agent_input = st.text_area(
        "告诉 prajna 你要生成什么",
        value="帮我做一份深圳互联网产品经理 P5 的薪资模板",
        height=100,
        placeholder="例如：生成电商销售团队本周周报 / 搭建京东集团财务知识库目录 / 为美的集团生成本月预算执行汇报 PPT / 生成美的集团采购管理套件 / 审查一份采购合同",
        label_visibility="collapsed",
    )

    st.markdown(
        """
        <small style="color:#64748b">💡 试试这样说：</small><br>
        <small style="color:#64748b">"帮我做一份上海制造业生产主管 P4 的薪资模板" · "生成招聘电商运营助理的 JD 和面试评估表" · "为京东集团生成客服 SOP" · "生成美的集团武汉工厂生产日报"</small>
        """,
        unsafe_allow_html=True,
    )

    execute_clicked = st.button("🚀 prajna 立即生成", type="primary", use_container_width=True)

    if execute_clicked:
        parsed = parse_agent_input(agent_input)
        if not parsed:
            st.error("无法识别输入内容，请尝试更清晰的描述。")
        else:
            intent = parsed["intent"]
            meta = SKILL_REGISTRY[intent]
            with st.spinner(f"prajna 正在生成 {meta['name']}..."):
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

    # Full capability demo
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        """
        <div class="section-title">🎬 全能力一键演示</div>
        <div class="section-subtitle">点击按钮，prajna 会依次调用全部 15 个技能并打包成 zip，真实展示平台覆盖能力</div>
        """,
        unsafe_allow_html=True,
    )

    run_all_clicked = st.button("🚀 一键运行全部 15 个技能", type="primary", use_container_width=True)

    if run_all_clicked:
        results = []
        failed = []
        progress = st.progress(0.0)
        status_container = st.empty()
        total = len(SKILL_REGISTRY)
        for idx, (skill_key, meta) in enumerate(SKILL_REGISTRY.items()):
            status_container.markdown(
                f"<div class='result-card'><b>⏳ 正在生成 {meta['icon']} {meta['name']} ({idx+1}/{total})...</b></div>",
                unsafe_allow_html=True,
            )
            try:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                r = run_single_skill_demo(skill_key, timestamp)
                results.append(r)
            except Exception as e:
                failed.append((meta["name"], str(e)))
            progress.progress((idx + 1) / total)

        status_container.empty()
        progress.empty()

        if failed:
            st.error(f"{len(failed)} 个技能生成失败")
            for name, err in failed:
                st.markdown(f"- **{name}**：{err}")
        else:
            zip_name = f"prajna_full_demo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
            zip_bytes = create_zip_from_results(results, zip_name)
            st.success(f"✅ 全部 {len(results)} 个技能已生成，共 {len(results) + sum(len(r['extras']) for r in results)} 个文件")
            st.download_button(
                label=f"📦 下载全能力演示包（{len(zip_bytes)/1024:.1f} KB）",
                data=zip_bytes,
                file_name=zip_name,
                mime="application/zip",
                use_container_width=True,
            )
            with st.expander("查看本次生成的文件清单"):
                for r in results:
                    files = [r["main"].name] + [e.name for e in r["extras"]]
                    st.markdown(f"**{r['icon'] if False else r['name']}**：{', '.join(files)}")

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
            params["company"] = st.text_input("企业名称", value="京东集团")
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
                params["author"] = st.text_input("编制人", value="prajna")

        elif selected_skill == "ai_daily":
            c1, c2 = st.columns(2)
            with c1:
                params["date"] = st.date_input("日报日期", value=datetime.now()).strftime("%Y-%m-%d")
            with c2:
                params["search"] = st.toggle("🌐 联网搜索最新动态", value=False, help="开启后会调用 DuckDuckGo 实时搜索，可能受网络环境影响")

        elif selected_skill == "finance_dashboard":
            c1, c2 = st.columns(2)
            with c1:
                params["preset"] = st.selectbox("企业预设", ["通用企业", "通用制造集团", "互联网/SaaS企业", "零售连锁企业", "新能源科技企业"])
                params["company"] = st.text_input("企业名称", value="美的集团")
            with c2:
                params["period"] = st.text_input("报表周期", value=datetime.now().strftime("%Y年%m月"))

        elif selected_skill == "budget_ppt":
            params["month"] = st.text_input("汇报月份", value=datetime.now().strftime("%Y年%m月"))
            params["company"] = st.text_input("公司名称", value="美的集团")
            params["author"] = st.text_input("汇报部门", value="财务部")

        elif selected_skill == "procurement":
            c1, c2 = st.columns(2)
            with c1:
                params["company"] = st.text_input("公司名称", value="美的集团")
                params["department"] = st.text_input("申请部门", value="采购部")
                params["applicant"] = st.text_input("申请人", value="张采购")
            with c2:
                params["date"] = st.date_input("申请日期", value=datetime.now()).strftime("%Y-%m-%d")
                params["delivery_date"] = st.date_input("期望到货日期", value=datetime.now()).strftime("%Y-%m-%d")

        elif selected_skill == "legal":
            c1, c2 = st.columns(2)
            with c1:
                params["company"] = st.text_input("公司名称", value="京东集团")
                params["contract_name"] = st.text_input("合同名称", value="软件采购与服务合同")
                params["contract_type"] = st.text_input("合同类型", value="采购合同")
            with c2:
                params["amount"] = st.number_input("合同金额（元）", min_value=0, value=580000, step=10000)
                params["term"] = st.text_input("合作期限", value="12个月")
                params["risk_level"] = st.selectbox("风险等级", ["高", "中", "低"], index=1)
                params["reviewer"] = st.text_input("审查人", value="法务经理")
                params["date"] = st.date_input("审查日期", value=datetime.now()).strftime("%Y-%m-%d")

        elif selected_skill == "cs_sop":
            params["company"] = st.text_input("公司名称", value="京东集团")
            params["date"] = st.date_input("更新日期", value=datetime.now()).strftime("%Y-%m-%d")

        elif selected_skill == "production":
            c1, c2 = st.columns(2)
            with c1:
                params["factory"] = st.text_input("工厂名称", value="美的集团武汉工厂")
            with c2:
                params["date"] = st.date_input("日报日期", value=datetime.now()).strftime("%Y-%m-%d")

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
                params["company"] = st.text_input("企业名称", value="prajna示范企业")

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
            if params.get("search"):
                cmd.append("--search")
        elif selected_skill == "procurement":
            out_name = f"prajna_采购管理套件_{params['company']}_{timestamp}.xlsx"
            out_path = HISTORY_DIR / out_name
            cmd = [sys.executable, str(meta["script"]), "--company", params["company"], "--department", params["department"], "--applicant", params["applicant"], "--date", params["date"], "--delivery-date", params["delivery_date"], "--output", str(out_path)]
        elif selected_skill == "legal":
            out_name = f"prajna_合同审查意见书_{params['company']}_{timestamp}.docx"
            out_path = HISTORY_DIR / out_name
            cmd = [sys.executable, str(meta["script"]), "--company", params["company"], "--contract-name", params["contract_name"], "--contract-type", params["contract_type"], "--amount", str(params["amount"]), "--term", params["term"], "--risk-level", params["risk_level"], "--reviewer", params["reviewer"], "--date", params["date"], "--output", str(out_path)]
        elif selected_skill == "cs_sop":
            out_name = f"prajna_客服SOP_{params['company']}_{timestamp}.xlsx"
            out_path = HISTORY_DIR / out_name
            cmd = [sys.executable, str(meta["script"]), "--company", params["company"], "--date", params["date"], "--output", str(out_path)]
        elif selected_skill == "production":
            safe_factory = re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9_-]", "_", params["factory"])[:20]
            out_name = f"prajna_生产日报_{safe_factory}_{timestamp}.xlsx"
            out_path = HISTORY_DIR / out_name
            cmd = [sys.executable, str(meta["script"]), "--factory", params["factory"], "--date", params["date"], "--output", str(out_path)]
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

        with st.spinner(f"prajna 正在生成 {meta['name']}..."):
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
            # These scripts create a directory from out_prefix and put files inside it
            output_dir = Path(str(out_prefix))
            if output_dir.exists() and output_dir.is_dir():
                excel_files = sorted(output_dir.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
                word_files = sorted(output_dir.glob("*.docx"), key=lambda p: p.stat().st_mtime, reverse=True)
                if excel_files:
                    out_path = excel_files[0]
                    out_name = out_path.name
                if word_files:
                    extra_files.append(word_files[0])

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
        <div class="section-title">🧠 prajna 原生 Agent 架构</div>
        <div class="section-subtitle">四层架构 + 五大原生能力，构建企业级智能体底座</div>
        """,
        unsafe_allow_html=True,
    )

    components.html(
        """
        <div class="mermaid-container">
        <div class="mermaid">
        %%{init: {'theme': 'base', 'themeVariables': { 'primaryColor': '#dbeafe', 'primaryTextColor': '#1e3a8a', 'primaryBorderColor': '#2563eb', 'lineColor': '#64748b', 'secondaryColor': '#f3e8ff', 'tertiaryColor': '#ecfeff' }}}%%
        graph TB
            subgraph APP["🏢 应用层 Application Layer"]
                A1[多模态对话]
                A2[AI 员工门户]
                A3[企业模板中台]
                A4[定时任务]
            end
            subgraph CAP["⚡ 能力层 Capability Layer"]
                B1[自主规划执行]
                B2[多智能体协同]
                B3[工具集成]
                B4[白盒可追溯]
            end
            subgraph MODEL["🧬 模型层 Model Layer"]
                C1[全模态记忆大模型]
                C2[记忆迭代]
                C3[意图理解]
            end
            subgraph DATA["💾 数据层 Data Layer"]
                D1[时间记忆]
                D2[语义网络]
                D3[企业知识库]
                D4[业务数据]
            end
            APP --> CAP
            CAP --> MODEL
            MODEL --> DATA
            DATA -->|memory_recall| CAP
        </div>
        </div>
        <script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>
        <script>mermaid.initialize({startOnLoad:true, securityLevel:'loose'});</script>
        """,
        height=520,
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
    st.markdown(
        """
        <div class="section-title">🏗️ 三层系统架构</div>
        <div class="section-subtitle">数字员工 Agent 层 · AI Agent 框架层 · 基础设施层</div>
        """,
        unsafe_allow_html=True,
    )

    l3, l2, l1 = st.columns(3)
    with l3:
        st.markdown(
            """
            <div class="capability-card" style="border-top:4px solid #2563eb;">
                <div class="capability-title">L3 数字员工 Agent 层</div>
                <div class="capability-desc">招聘 / 薪酬 / 入职 / 客服 / 合规 / 标书 / 审批全链路；32 位专业 Agent 协同完成复杂 HR 业务流程。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with l2:
        st.markdown(
            """
            <div class="capability-card" style="border-top:4px solid #7c3aed;">
                <div class="capability-title">L2 AI Agent 框架层</div>
                <div class="capability-desc">智能对话引擎、工具调用系统、多平台消息网关、多 LLM 调度、持久记忆管理。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with l1:
        st.markdown(
            """
            <div class="capability-card" style="border-top:4px solid #10b981;">
                <div class="capability-title">L1 基础设施层</div>
                <div class="capability-desc">CLI 终端界面、消息平台适配器、安全与权限控制、AIGC 隐式水印与审计日志。</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(
        """
        <div class="section-title">👥 32 位数字员工 Agent 矩阵</div>
        <div class="section-subtitle">四大类 Agent 覆盖 HR 核心、业务增长、管理决策与平台支撑</div>
        """,
        unsafe_allow_html=True,
    )

    agent_matrix = [
        ("💼", "HR 核心", "15 位", "招聘 / 薪酬 / 入职 / 客服 / 合规 / 标书 / 审批等", "#2563eb"),
        ("📈", "业务增长", "5 位", "文案 / 市场 / 销售 / 客户成功等", "#7c3aed"),
        ("🎯", "管理决策", "8 位", "财务 / 会议 / 战略 / 采购 / 经营分析等", "#10b981"),
        ("⚙️", "平台支撑", "4 位", "意图 / 配置 / 调度 / 数据接入等", "#f59e0b"),
    ]
    am1, am2, am3, am4 = st.columns(4)
    for col, (icon, title, count, desc, color) in zip([am1, am2, am3, am4], agent_matrix):
        with col:
            st.markdown(
                f"""
                <div class="capability-card" style="border-top:4px solid {color};">
                    <div style="font-size:2rem;margin-bottom:0.5rem;">{icon}</div>
                    <div class="capability-title">{title} <span style="font-size:1.5rem;color:{color};">{count}</span></div>
                    <div class="capability-desc">{desc}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">🔄 Agent 标准生命周期</div>', unsafe_allow_html=True)

    components.html(
        """
        <div class="mermaid-container">
        <div class="mermaid">
        %%{init: {'theme': 'base', 'themeVariables': { 'primaryColor': '#dbeafe', 'primaryTextColor': '#1e3a8a', 'primaryBorderColor': '#2563eb', 'lineColor': '#64748b', 'secondaryColor': '#f3e8ff' }}}%%
        graph LR
            A["1️⃣ 注册 Register"] --> B["2️⃣ 监听 Listen"]
            B --> C["3️⃣ 感知 Perceive"]
            C --> D["4️⃣ 规划 Plan"]
            D --> E["5️⃣ 执行 Act"]
            E --> F["6️⃣ 反思 Reflect"]
            F --> G["7️⃣ 上报 Report"]
        </div>
        </div>
        <script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>
        <script>mermaid.initialize({startOnLoad:true, securityLevel:'loose'});</script>
        """,
        height=220,
    )

    lifecycle_steps = [
        ("1️⃣ 注册 Register", "向 Agent 调度中心上报名称、能力、依赖与资源配额"),
        ("2️⃣ 监听 Listen", "接收来自调度中心的任务或事件"),
        ("3️⃣ 感知 Perceive", "调用 memory_recall 加载相关上下文"),
        ("4️⃣ 规划 Plan", "拆解任务、选择工具、评估风险"),
        ("5️⃣ 执行 Act", "调用工具、调用其他 Agent、处理数据"),
        ("6️⃣ 反思 Reflect", "调用 memory_reflect 记录结果与改进点"),
        ("7️⃣ 上报 Report", "返回结果、状态与后续建议"),
    ]
    c1, c2 = st.columns(2)
    for i, (title, desc) in enumerate(lifecycle_steps):
        with (c1 if i < 4 else c2):
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
        <div class="section-subtitle">prajna 原生 Agent 如何通过记忆核心实现跨 Agent 协同</div>
        """,
        unsafe_allow_html=True,
    )

    # Interactive scheduler demo
    st.markdown(
        """
        <div style="background:linear-gradient(135deg,#eff6ff,#f5f3ff);border-radius:16px;padding:1.5rem;border:1px solid #e2e8f0;margin-bottom:1.5rem;">
            <h4 style="margin-top:0;color:#1e40af;">🎛️ Agent 调度中心：输入任务，查看 prajna 如何分配 Agent</h4>
            <p style="color:#64748b;margin:0;">调度中心会根据意图识别结果、Agent 能力匹配度、负载、历史成功率与记忆协同分，自动选择最优 Agent。</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    sched_input = st.text_input("输入任务描述", value="帮我生成一份广州 P2 电商运营助理的招聘套件", key="agent_sched_input")
    if st.button("🎯 执行调度", use_container_width=True, key="agent_sched_btn"):
        detected_intent = detect_intent(sched_input)
        intent_meta = SKILL_REGISTRY.get(detected_intent, {"name": "未知", "category": "其他"})

        agent_pool = [
            {"key": "recruitment", "name": "招聘Agent", "base_match": 0.95, "load": 0.35, "success": 0.96, "memory": 0.92},
            {"key": "salary", "name": "薪酬Agent", "base_match": 0.55, "load": 0.20, "success": 0.99, "memory": 0.70},
            {"key": "compensation", "name": "薪酬体系Agent", "base_match": 0.50, "load": 0.25, "success": 0.98, "memory": 0.65},
            {"key": "bidding", "name": "标书Agent", "base_match": 0.30, "load": 0.40, "success": 0.94, "memory": 0.60},
            {"key": "sales", "name": "销售Agent", "base_match": 0.25, "load": 0.60, "success": 0.88, "memory": 0.50},
            {"key": "customer_service", "name": "客服Agent", "base_match": 0.20, "load": 0.60, "success": 0.88, "memory": 0.50},
            {"key": "procurement", "name": "采购Agent", "base_match": 0.20, "load": 0.45, "success": 0.93, "memory": 0.55},
            {"key": "legal", "name": "法务Agent", "base_match": 0.20, "load": 0.30, "success": 0.95, "memory": 0.60},
            {"key": "production", "name": "生产Agent", "base_match": 0.20, "load": 0.50, "success": 0.91, "memory": 0.55},
        ]
        # Boost the agent whose key matches detected intent
        for a in agent_pool:
            if a["key"] == detected_intent:
                a["base_match"] = 0.98

        scored = []
        for a in agent_pool:
            total = 0.35 * a["base_match"] + 0.25 * (1 - a["load"]) + 0.20 * a["success"] + 0.20 * a["memory"]
            scored.append({**a, "score": total})
        scored.sort(key=lambda x: x["score"], reverse=True)

        st.markdown(
            f"""
            <div class="result-card" style="border-left-color:#7c3aed;">
                <b>🧠 意图识别：</b>{intent_meta['name']}（intent={detected_intent}）<br>
                <b>📂 业务域：</b>{intent_meta['category']}
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("**调度评分结果**")
        for idx, a in enumerate(scored):
            rank = "🥇" if idx == 0 else "🥈" if idx == 1 else "🥉" if idx == 2 else f"{idx+1}️⃣"
            st.markdown(
                f"""
                <div style="background:white;border-radius:10px;padding:1rem;border:1px solid #e2e8f0;margin-bottom:0.5rem;display:flex;justify-content:space-between;align-items:center;">
                    <div><b>{rank} {a['name']}</b><br><span style="color:#64748b;font-size:0.85rem;">能力匹配 {a['base_match']:.0%} · 负载 {a['load']:.0%} · 历史成功率 {a['success']:.0%} · 记忆协同 {a['memory']:.0%}</span></div>
                    <div style="font-size:1.25rem;font-weight:800;color:#2563eb;">{a['score']:.3f}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        st.success(f"✅ 调度决策：任务分配给 **{scored[0]['name']}**，原因：能力匹配度最高且记忆协同分最优。")
        st.info("📝 调度结果已写入 agent 类型记忆：任务 ID、目标 Agent、调度原因、预期 SLA。")

    st.divider()
    st.markdown('<div class="section-subtitle">典型 Agent 联动场景</div>', unsafe_allow_html=True)
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
# Tab 5: prajna Full Capability Showcase
# ---------------------------------------------------------------------------
with tab_showcase:
    st.markdown(
        """
        <div class="section-title">🎁 prajna 全能力 showcase</div>
        <div class="section-subtitle">从底层记忆核心到上层业务 Agent，一览 prajna 企业智能体平台的完整能力矩阵</div>
        """,
        unsafe_allow_html=True,
    )

    # Sub-tabs for organized showcase
    sub_home, sub_scenarios, sub_skills, sub_memory, sub_workflow, sub_spec, sub_system, sub_trace, sub_create = st.tabs([
        "🗺️ 能力全景图",
        "📦 业务场景包",
        "📦 技能目录",
        "💾 记忆核心控制台",
        "🔄 多 Agent 协同",
        "📋 原生规范",
        "🎛️ 系统级技能",
        "🔍 白盒追溯",
        "➕ 创建新 Skill",
    ])

    # ---------- Sub-tab 1: Capability Map ----------
    with sub_home:
        st.markdown(
            """
            <div style="background:white;border-radius:16px;padding:2rem;border:1px solid #e2e8f0;box-shadow:0 4px 6px -1px rgba(0,0,0,0.05);">
                <h3 style="margin-top:0;color:#1e3a8a;">prajna 能力分层全景</h3>
                <div style="display:flex;flex-direction:column;gap:0.75rem;margin-top:1.5rem;">
                    <div style="background:linear-gradient(90deg,#1e40af,#3b82f6);color:white;border-radius:12px;padding:1rem;">
                        <strong>🚀 应用层</strong> · AI 员工门户 · 多模态对话 · 定时任务 · 企业模板中台
                    </div>
                    <div style="background:linear-gradient(90deg,#5b21b6,#8b5cf6);color:white;border-radius:12px;padding:1rem;">
                        <strong>⚡ 能力层</strong> · 自然语言意图识别 · 自主规划执行 · 多智能体协同 · 工具集成 · 白盒可追溯
                    </div>
                    <div style="background:linear-gradient(90deg,#0e7490,#06b6d4);color:white;border-radius:12px;padding:1rem;">
                        <strong>🧠 模型层</strong> · 全模态记忆大模型 · 记忆迭代 · 小样本适配 · LLM 意图理解
                    </div>
                    <div style="background:linear-gradient(90deg,#047857,#10b981);color:white;border-radius:12px;padding:1rem;">
                        <strong>💾 数据层</strong> · 时间记忆 · 语义网络 · 企业知识库 · 业务数据 · 智能剪枝
                    </div>
                </div>
                <p style="color:#64748b;margin-top:1.5rem;line-height:1.6;">
                    prajna 不是单一 Agent，而是一套<b>企业智能体操作系统</b>：底层记忆核心让 Agent 拥有长期记忆，
                    原生架构规范让所有 Agent 可协同、可追溯，上层技能市场覆盖 HR、销售、财务、生产、情报等全业务场景。
                </p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # ---------- Sub-tab 2: Scenario Packages ----------
    with sub_scenarios:
        st.markdown(
            """
            <div class="section-title">📦 多 Agent 业务场景包</div>
            <div class="section-subtitle">一个真实业务场景往往需要多个 Agent 协同。点击卡片，prajna 会按工作流顺序生成完整文档包。</div>
            """,
            unsafe_allow_html=True,
        )

        scenario_cols = st.columns(2)
        for idx, (scenario_key, cfg) in enumerate(SCENARIOS.items()):
            with scenario_cols[idx % 2]:
                st.markdown(
                    f"""
                    <div class="scenario-card" style="border-top:4px solid {cfg['color']};">
                        <div class="scenario-title">{cfg['icon']} {scenario_key}</div>
                        <div class="scenario-desc">{cfg['desc']}</div>
                        <div class="scenario-tags">
                            {''.join(f'<span class="scenario-tag">{SKILL_REGISTRY[k]["name"]}</span>' for k in cfg['skills'])}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button(f"🚀 生成 {scenario_key}", key=f"scenario_{scenario_key}", use_container_width=True):
                    with st.spinner(f"prajna 正在生成 {scenario_key}..."):
                        try:
                            results = run_scenario(scenario_key)
                            zip_name = f"prajna_{scenario_key}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
                            zip_bytes = create_zip_from_results(results, zip_name)
                            total_files = sum(1 + len(r["extras"]) for r in results)
                            st.success(f"✅ {scenario_key} 已生成，共 {total_files} 个文件")
                            st.download_button(
                                label=f"📦 下载 {scenario_key}（{len(zip_bytes)/1024:.1f} KB）",
                                data=zip_bytes,
                                file_name=zip_name,
                                mime="application/zip",
                                use_container_width=True,
                                key=f"dl_{scenario_key}",
                            )
                            with st.expander("文件清单"):
                                for r in results:
                                    st.markdown(f"**{r['name']}**：{r['main'].name}" + (f"、{', '.join(e.name for e in r['extras'])}" if r['extras'] else ""))
                        except Exception as e:
                            st.error(f"生成失败：{e}")

    # ---------- Sub-tab 3: Skill Catalog ----------
    with sub_skills:
        @st.cache_data(ttl=300)
        def scan_skills(root_dirs):
            skills = []
            seen = set()
            for root_dir in root_dirs:
                root = Path(root_dir).expanduser()
                if not root.exists():
                    continue
                for skill_dir in root.rglob("SKILL.md"):
                    key = str(skill_dir.parent)
                    if key in seen:
                        continue
                    seen.add(key)
                    rel = skill_dir.relative_to(root)
                    parts = rel.parts
                    category = parts[0] if len(parts) > 1 else "其他"
                    name = skill_dir.parent.name
                    desc = ""
                    skill_id = ""
                    try:
                        with open(skill_dir, "r", encoding="utf-8") as f:
                            content = f.read(2000)
                        # parse frontmatter description
                        if "description:" in content:
                            m = re.search(r"description:\s*(\|?\s*\n?\s*)(.+?)(\n\w+:|\n# |\n---|\Z)", content, re.S)
                            if m:
                                desc = m.group(2).strip().replace("\n", " ")[:120]
                        if "skill_id:" in content:
                            m = re.search(r"skill_id:\s*(.+)", content)
                            if m:
                                skill_id = m.group(1).strip()
                    except Exception:
                        pass
                    skills.append({
                        "category": category,
                        "name": name,
                        "skill_id": skill_id or name,
                        "path": str(skill_dir.parent),
                        "desc": desc or "企业级智能体技能",
                    })
            return skills

        all_skills = scan_skills(["~/.prajna/skills", str(APP_DIR / "skills")])

        # Stats
        s1, s2, s3 = st.columns(3)
        with s1:
            st.markdown(f'<div class="metric-card"><div class="metric-value">{len(all_skills)}</div><div class="metric-label">已安装技能</div></div>', unsafe_allow_html=True)
        with s2:
            st.markdown(f'<div class="metric-card"><div class="metric-value">{len(set(s["category"] for s in all_skills))}</div><div class="metric-label">技能领域</div></div>', unsafe_allow_html=True)
        with s3:
            st.markdown('<div class="metric-card"><div class="metric-value">∞</div><div class="metric-label">可扩展</div></div>', unsafe_allow_html=True)

        # Highlighted skills (system + contest)
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div style="font-weight:600;color:#1e293b;margin-bottom:0.75rem;">🌟 核心系统技能 + 本次大赛技能</div>', unsafe_allow_html=True)

        highlight_keywords = [
            "prajna-memory-core", "prajna-native-agent-architecture", "prajna-native-agent-adoption",
            "Agent调度中心", "监控看板", "系统配置技能",
            "prajna-salary-template", "prajna-sales-weekly-report", "prajna-recruitment-assistant",
            "prajna-compensation-system", "prajna-performance-system", "prajna-bidding-assistant",
            "prajna-ecommerce-finance-kb-catalog", "prajna-financial-dashboard", "prajna-budget-execution-ppt",
            "prajna-clothing-teamleader-duty", "prajna-ai-leader-daily",
            "prajna-procurement-assistant", "prajna-contract-review-assistant",
            "prajna-customer-service-sop", "prajna-production-daily-report",
        ]
        highlighted = [s for s in all_skills if any(kw in s["path"] for kw in highlight_keywords)]

        cols = st.columns(3)
        for idx, skill in enumerate(highlighted[:18]):
            with cols[idx % 3]:
                st.markdown(
                    f"""
                    <div class="skill-card">
                        <div class="skill-name">{skill['name']}</div>
                        <div style="color:#64748b;font-size:0.75rem;margin-bottom:0.5rem;">{skill['category']}</div>
                        <div class="skill-desc">{skill['desc']}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        # Full catalog by category
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("📂 查看全部技能目录"):
            categories = sorted(set(s["category"] for s in all_skills))
            for cat in categories:
                st.markdown(f"**{cat}**")
                cat_skills = [s for s in all_skills if s["category"] == cat]
                for skill in cat_skills:
                    st.markdown(f"<small>• {skill['name']} — {skill['desc'][:80]}</small>", unsafe_allow_html=True)

    # ---------- Sub-tab 3: Memory Core Console ----------
    with sub_memory:
        st.markdown(
            """
            <div style="background:linear-gradient(135deg,#f5f3ff,#eff6ff);border-radius:16px;padding:1.5rem;border:1px solid #e2e8f0;margin-bottom:1rem;">
                <h4 style="margin-top:0;color:#5b21b6;">💾 记忆核心交互演示</h4>
                <p style="color:#64748b;margin:0;">模拟调用 memory_remember / memory_recall / memory_reflect / memory_forget 四个原生接口。数据仅存于当前会话，用于演示 prajna 全模态记忆核心的接口形态与智能剪枝效果。</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if "prajna_memory" not in st.session_state:
            st.session_state.prajna_memory = [
                {"name": "user_preference_formal", "type": "user", "content": "用户偏好正式汇报风格", "importance": 4},
                {"name": "business_reimburse_rule", "type": "business", "content": "报销需提前3天申请，发票须为增值税专用发票", "importance": 5},
                {"name": "project_q3_compensation", "type": "project", "content": "Q3目标：上线薪酬自动核算模块", "importance": 5},
                {"name": "agent_bidding_lesson", "type": "agent", "content": "上次智慧园区投标因报价偏高未中标，需加强成本核算", "importance": 4},
                {"name": "feedback_cs_long_report", "type": "feedback", "content": "客服主管反馈周报内容太长，希望只看关键指标", "importance": 3},
                {"name": "reference_salary_guangzhou_p2", "type": "reference", "content": "广州 P2 电商运营助理市场薪资带宽 6-9K", "importance": 4},
            ]

        # Stats
        mem_types = [m["type"] for m in st.session_state.prajna_memory]
        avg_importance = sum(m["importance"] for m in st.session_state.prajna_memory) / max(len(st.session_state.prajna_memory), 1)
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.markdown(f'<div class="metric-card"><div class="metric-value" style="color:#8b5cf6;">{len(st.session_state.prajna_memory)}</div><div class="metric-label">记忆条目</div></div>', unsafe_allow_html=True)
        with m2:
            st.markdown(f'<div class="metric-card"><div class="metric-value" style="color:#10b981;">{len(set(mem_types))}</div><div class="metric-label">记忆类型</div></div>', unsafe_allow_html=True)
        with m3:
            st.markdown(f'<div class="metric-card"><div class="metric-value" style="color:#f59e0b;">{avg_importance:.1f}</div><div class="metric-label">平均重要性</div></div>', unsafe_allow_html=True)
        with m4:
            st.markdown(f'<div class="metric-card"><div class="metric-value" style="color:#06b6d4;"><span class="pulse-dot"></span>正常</div><div class="metric-label">记忆核心状态</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        col_a, col_b = st.columns([1, 1])
        with col_a:
            st.markdown("**memory_remember — 写入记忆**")
            mem_name = st.text_input("记忆名称", value="feedback_report_concise", key="mem_name")
            mem_type = st.selectbox("记忆类型", ["user", "session", "project", "business", "agent", "feedback", "reference"], key="mem_type")
            mem_content = st.text_area("记忆内容", value="用户反馈生成的报告太长，希望更简洁", key="mem_content")
            mem_importance = st.slider("重要性", 1, 5, 3, key="mem_importance")
            if st.button("📝 写入记忆", use_container_width=True, key="mem_remember_btn"):
                st.session_state.prajna_memory.append({
                    "name": mem_name, "type": mem_type, "content": mem_content, "importance": mem_importance
                })
                st.success(f"✅ 已写入记忆：{mem_name}")

        with col_b:
            st.markdown("**memory_recall — 召回记忆**")
            recall_query = st.text_input("查询意图", value="用户喜欢什么风格", key="recall_query")
            recall_types = st.multiselect("记忆类型过滤", ["user", "session", "project", "business", "agent", "feedback", "reference"], default=["user", "business"], key="recall_types")
            if st.button("🔍 召回记忆", use_container_width=True, key="mem_recall_btn"):
                results = []
                for mem in st.session_state.prajna_memory:
                    if recall_types and mem["type"] not in recall_types:
                        continue
                    if any(kw in mem["content"] for kw in recall_query.split()):
                        results.append(mem)
                if results:
                    st.markdown(f"<div style='color:#64748b;font-size:0.85rem;margin-bottom:0.5rem;'>召回 {len(results)} 条相关记忆</div>", unsafe_allow_html=True)
                    for r in results:
                        st.markdown(f"<div style='background:white;border-radius:8px;padding:0.75rem;border-left:3px solid #8b5cf6;margin-bottom:0.5rem;'><b>{r['name']}</b> <span style='color:#64748b;font-size:0.8rem;'>[{r['type']}]</span><br><span style='font-size:0.9rem;'>{r['content']}</span></div>", unsafe_allow_html=True)
                else:
                    st.info("未召回相关记忆")

        st.markdown("<br>", unsafe_allow_html=True)

        # Memory reflect & prune
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**memory_reflect — 自我反思**")
            reflect_types = st.multiselect("选择要反思的记忆类型", ["user", "session", "project", "business", "agent", "feedback", "reference"], default=["feedback", "agent"], key="reflect_types")
            if st.button("🔄 生成反思洞察", use_container_width=True, key="mem_reflect_btn"):
                filtered = [m for m in st.session_state.prajna_memory if m["type"] in reflect_types]
                if filtered:
                    insights = []
                    feedbacks = [m for m in filtered if m["type"] == "feedback"]
                    agents = [m for m in filtered if m["type"] == "agent"]
                    if feedbacks:
                        topics = set()
                        for f in feedbacks:
                            if "长" in f["content"] or "简洁" in f["content"]:
                                topics.add("报告篇幅控制")
                            if "慢" in f["content"] or "快" in f["content"]:
                                topics.add("响应速度")
                        insights.append(f"发现 {len(feedbacks)} 条反馈记忆，主要关注点：{', '.join(topics) if topics else '用户体验'}")
                    if agents:
                        insights.append(f"发现 {len(agents)} 条 Agent 经验记忆，可沉淀为策略记忆优化后续任务")
                    st.success("反思完成")
                    for ins in insights:
                        st.markdown(f"<div style='background:#f0fdf4;border-radius:8px;padding:0.75rem;border-left:3px solid #10b981;margin-bottom:0.5rem;'>💡 {ins}</div>", unsafe_allow_html=True)
                else:
                    st.info("所选类型下暂无记忆")

        with c2:
            st.markdown("**智能剪枝 — 清理低重要性记忆**")
            prune_threshold = st.slider("重要性阈值", 1, 5, 2, key="prune_threshold")
            if st.button("✂️ 执行智能剪枝", use_container_width=True, key="mem_prune_btn"):
                before = len(st.session_state.prajna_memory)
                st.session_state.prajna_memory = [m for m in st.session_state.prajna_memory if m["importance"] >= prune_threshold]
                after = len(st.session_state.prajna_memory)
                st.success(f"已剪枝 {before - after} 条低重要性记忆，剩余 {after} 条")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("**记忆语义网络**")
        node_html = ""
        for mem in st.session_state.prajna_memory:
            node_html += f'<span class="memory-node memory-node-{mem["type"]}">{mem["type"]}: {mem["name"]} {"⭐"*mem["importance"]}</span>'
        st.markdown(f'<div style="background:white;border-radius:16px;padding:1rem;border:1px solid #e2e8f0;">{node_html}</div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("**当前记忆库清单**")
        for mem in st.session_state.prajna_memory[-10:]:
            badge_color = {"user":"#3b82f6","session":"#06b6d4","project":"#8b5cf6","business":"#10b981","agent":"#f59e0b","feedback":"#ef4444","reference":"#64748b"}.get(mem["type"], "#64748b")
            st.markdown(
                f"""
                <div style="display:flex;align-items:center;gap:0.75rem;background:white;border-radius:8px;padding:0.75rem;border:1px solid #e2e8f0;margin-bottom:0.5rem;">
                    <span style="background:{badge_color};color:white;font-size:0.7rem;padding:0.2rem 0.5rem;border-radius:999px;white-space:nowrap;">{mem['type']}</span>
                    <span style="font-weight:500;">{mem['name']}</span>
                    <span style="color:#64748b;font-size:0.85rem;flex:1;">{mem['content']}</span>
                    <span style="color:#f59e0b;font-size:0.8rem;">{'⭐' * mem['importance']}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # ---------- Sub-tab 4: Multi-Agent Workflow ----------
    with sub_workflow:
        st.markdown(
            """
            <div style="background:linear-gradient(135deg,#eff6ff,#f0fdf4);border-radius:16px;padding:1.5rem;border:1px solid #e2e8f0;margin-bottom:1rem;">
                <h4 style="margin-top:0;color:#1e40af;">🔄 端到端多 Agent 协同演示</h4>
                <p style="color:#64748b;margin:0;">选择一个业务场景，查看意图 Agent → 调度中心 → 业务 Agent → 记忆核心 的完整协作链路。</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        scenario = st.selectbox("选择演示场景", ["招聘一名电商运营助理", "投标智慧园区建设项目", "处理用户报销咨询", "采购一批服务器设备", "审查一份采购合同"])

        if scenario == "招聘一名电商运营助理":
            steps = [
                ("👤 用户输入", "\"我要招聘一名广州 P2 电商运营助理，月薪 6-9K\""),
                ("🎯 意图 Agent", "调用 memory_recall(user/session/business) → 识别为招聘意图 → 提取岗位、城市、职级、薪资"),
                ("🎛️ Agent 调度中心", "匹配招聘 Agent，评估其历史岗位画像经验 → 分发任务"),
                ("🤝 招聘 Agent", "生成 JD、面试评估表、Offer 薪资建议；沉淀岗位画像"),
                ("💰 薪酬 Agent", "根据城市系数与职级带宽校验 Offer 合理性"),
                ("⭐ 绩效 Agent", "预生成该岗位 KPI 指标库与绩效合同"),
                ("📝 入职 Agent", "Offer 接受后自动接管，生成分配计划、培训计划、薪酬开通"),
                ("💾 记忆沉淀", "招聘结果写入 project/business；成功经验写入 agent 记忆"),
            ]
        elif scenario == "投标智慧园区建设项目":
            steps = [
                ("📄 招标公告输入", "\"智慧园区智能化建设项目，预算 580 万，工期 180 天\""),
                ("🎯 意图 Agent", "识别为招投标意图，解析关键约束：金额、工期、资质"),
                ("🎛️ Agent 调度中心", "并行分发：标书 Agent + 财务 Agent + 法务 Agent"),
                ("🎯 标书 Agent", "生成资格自审表、技术偏离表、商务报价表、评分响应索引"),
                ("💵 财务 Agent", "核算成本与毛利率，提供报价建议"),
                ("⚖️ 法务 Agent", "审核合同条款与风险点"),
                ("📑 合成 Agent", "汇总生成投标文件 Word 大纲与 Excel 套件"),
                ("🔄 投标后反思", "memory_reflect 复盘报价策略，沉淀为 agent 记忆"),
            ]
        elif scenario == "处理用户报销咨询":
            steps = [
                ("👤 用户输入", "\"我的报销单被退回了，是什么原因？\""),
                ("🎯 意图 Agent", "识别为客服咨询，加载 user/session/business 记忆"),
                ("🎛️ Agent 调度中心", "匹配客服 Agent，读取用户历史工单"),
                ("🎧 客服 Agent", "查询退回原因：缺少增值税专用发票 → 生成补正指引"),
                ("📝 工单 Agent", "创建未结工单，写入 project 记忆"),
                ("📊 反馈 Agent", "识别负面情绪，写入 feedback 记忆"),
                ("🔄 反思优化", "memory_reflect 建议财务部门优化报销单填写引导"),
            ]
        elif scenario == "采购一批服务器设备":
            steps = [
                ("👤 用户输入", "\"采购部需要采购 5 台服务器，预算 10 万，请走采购流程\""),
                ("🎯 意图 Agent", "识别为采购意图，提取品类、数量、预算、期望到货时间"),
                ("🎛️ Agent 调度中心", "匹配采购 Agent，触发供应商评估与询价比价"),
                ("🛒 采购 Agent", "生成采购申请单、供应商评估表、询价比价单"),
                ("⚖️ 法务 Agent", "审查采购合同关键条款与风险点"),
                ("💵 财务 Agent", "核对预算科目与付款节点"),
                ("📦 入库 Agent", "到货验收后更新采购台账与库存记忆"),
                ("🔄 反思优化", "memory_reflect 沉淀供应商表现与采购周期数据"),
            ]
        else:
            steps = [
                ("👤 用户输入", "\"请法务帮忙审查这份软件采购合同，金额 58 万\""),
                ("🎯 意图 Agent", "识别为合同审查意图，解析合同类型与金额"),
                ("🎛️ Agent 调度中心", "匹配法务 Agent，调用历史合同审查经验"),
                ("⚖️ 法务 Agent", "从主体资格、商务、财务、违约、知产、争议解决六维度审查"),
                ("💵 财务 Agent", "评估付款条款与税务风险"),
                ("🎯 业务 Agent", "确认交付验收标准与 SLA"),
                ("📑 合成 Agent", "输出合同审查意见书 Word 文档"),
                ("💾 记忆沉淀", "审查结论写入 business 记忆，风险点写入 project 记忆"),
            ]

        for idx, (title, desc) in enumerate(steps):
            st.markdown(
                f"""
                <div class="flow-step">
                    <div class="flow-step-title">{title}</div>
                    <div class="flow-step-desc">{desc}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if idx < len(steps) - 1:
                st.markdown('<div class="flow-arrow">⬇️</div>', unsafe_allow_html=True)

    # ---------- Sub-tab 5: Native Spec ----------
    with sub_spec:
        st.markdown(
            """
            <div style="background:white;border-radius:16px;padding:1.5rem;border:1px solid #e2e8f0;box-shadow:0 4px 6px -1px rgba(0,0,0,0.05);">
                <h4 style="margin-top:0;color:#1e3a8a;">📋 SKILL.md Frontmatter 规范</h4>
                <pre style="background:#0f172a;color:#e2e8f0;padding:1rem;border-radius:8px;overflow-x:auto;font-size:0.85rem;">
---
name: 🎁 prajna
skill_id: recruitment-agent
description: 一句话描述 Agent 能力
version: 2.0.0
author: 锦辉人力·prajna
tags: [prajna, agent, hr]
model: deepseek-chat
native_capabilities:
  - autonomous_planning
  - multi_agent_collaboration
  - tool_integration
  - traceability
  - memory_linkage
requires:
  - prajna-memory-core
  - prajna-native-agent-architecture
memory_context:
  - user
  - business
  - agent
---
                </pre>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            """
            <div style="background:white;border-radius:16px;padding:1.5rem;border:1px solid #e2e8f0;box-shadow:0 4px 6px -1px rgba(0,0,0,0.05);">
                <h4 style="margin-top:0;color:#1e3a8a;">🔧 记忆核心统一接口</h4>
                <table style="width:100%;border-collapse:collapse;font-size:0.9rem;">
                    <tr style="background:#f1f5f9;"><th style="padding:0.75rem;text-align:left;border:1px solid #e2e8f0;">接口</th><th style="padding:0.75rem;text-align:left;border:1px solid #e2e8f0;">作用</th></tr>
                    <tr><td style="padding:0.75rem;border:1px solid #e2e8f0;"><code>memory_recall</code></td><td style="padding:0.75rem;border:1px solid #e2e8f0;">召回与当前任务相关的记忆</td></tr>
                    <tr><td style="padding:0.75rem;border:1px solid #e2e8f0;"><code>memory_remember</code></td><td style="padding:0.75rem;border:1px solid #e2e8f0;">将信息写入记忆核心</td></tr>
                    <tr><td style="padding:0.75rem;border:1px solid #e2e8f0;"><code>memory_reflect</code></td><td style="padding:0.75rem;border:1px solid #e2e8f0;">对执行过程反思并更新策略记忆</td></tr>
                    <tr><td style="padding:0.75rem;border:1px solid #e2e8f0;"><code>memory_forget</code></td><td style="padding:0.75rem;border:1px solid #e2e8f0;">删除或归档指定记忆</td></tr>
                </table>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # ---------- Sub-tab 6: System-Level Skills ----------
    with sub_system:
        sys_tab = st.radio("选择系统级技能", ["🎛️ Agent 调度中心", "📊 监控看板", "⚙️ 系统配置"], horizontal=True)

        if sys_tab == "🎛️ Agent 调度中心":
            st.markdown(
                """
                <div style="background:linear-gradient(135deg,#eff6ff,#f5f3ff);border-radius:16px;padding:1.5rem;border:1px solid #e2e8f0;margin-bottom:1rem;">
                    <h4 style="margin-top:0;color:#1e40af;">Agent 调度中心演示</h4>
                    <p style="color:#64748b;margin:0;">输入一个任务，调度中心会根据 Agent 能力匹配度、负载、历史成功率和记忆协同分，选择最优 Agent。</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            task_input = st.text_input("输入任务描述", value="帮我生成一份广州 P2 电商运营助理的招聘套件", key="system_sched_input")
            if st.button("🎯 执行调度", use_container_width=True, key="system_sched_btn"):
                # Simulate scoring
                agents = [
                    {"name": "招聘Agent", "match": 0.98, "load": 0.35, "success": 0.96, "memory": 0.92},
                    {"name": "薪酬Agent", "match": 0.45, "load": 0.20, "success": 0.99, "memory": 0.70},
                    {"name": "客服Agent", "match": 0.30, "load": 0.60, "success": 0.88, "memory": 0.50},
                    {"name": "法务Agent", "match": 0.55, "load": 0.30, "success": 0.95, "memory": 0.70},
                    {"name": "采购Agent", "match": 0.35, "load": 0.45, "success": 0.93, "memory": 0.55},
                    {"name": "生产Agent", "match": 0.25, "load": 0.50, "success": 0.91, "memory": 0.55},
                ]
                scored = []
                for a in agents:
                    base = 0.25 * (1 - a["load"]) + 0.25 * a["match"] + 0.15 * a["success"] + 0.15 * (1 - 0.1)
                    memory_score = 0.10 * a["success"] + 0.10 * a["memory"]
                    total = base + memory_score
                    scored.append({**a, "score": total})
                scored.sort(key=lambda x: x["score"], reverse=True)

                st.markdown("**调度评分结果**")
                for idx, a in enumerate(scored):
                    rank = "🥇" if idx == 0 else "🥈" if idx == 1 else "🥉" if idx == 2 else "4️⃣"
                    st.markdown(
                        f"""
                        <div style="background:white;border-radius:10px;padding:1rem;border:1px solid #e2e8f0;margin-bottom:0.5rem;display:flex;justify-content:space-between;align-items:center;">
                            <div><b>{rank} {a['name']}</b><br><span style="color:#64748b;font-size:0.85rem;">能力匹配 {a['match']:.0%} · 负载 {a['load']:.0%} · 历史成功率 {a['success']:.0%} · 记忆协同 {a['memory']:.0%}</span></div>
                            <div style="font-size:1.25rem;font-weight:700;color:#2563eb;">{a['score']:.3f}</div>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                winner = scored[0]["name"]
                st.success(f"✅ 调度决策：任务分配给 **{winner}**，原因：能力匹配度最高且记忆协同分最优。")
                st.info("📝 调度结果已写入 agent 类型记忆：任务 ID、目标 Agent、调度原因、预期 SLA。")

        elif sys_tab == "📊 监控看板":
            st.markdown(
                """
                <div style="background:linear-gradient(135deg,#f0fdf4,#eff6ff);border-radius:16px;padding:1.5rem;border:1px solid #e2e8f0;margin-bottom:1rem;">
                    <h4 style="margin-top:0;color:#047857;">prajna 系统运行监控</h4>
                    <p style="color:#64748b;margin:0;">实时展示各 Agent 健康状态、业务指标与记忆核心健康度（演示数据）。</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.markdown('<div class="metric-card" style="border-bottom-color:#10b981;"><div class="metric-value" style="color:#10b981;">98.5%</div><div class="metric-label">系统可用性</div></div>', unsafe_allow_html=True)
            with m2:
                st.markdown('<div class="metric-card" style="border-bottom-color:#3b82f6;"><div class="metric-value" style="color:#3b82f6;">42</div><div class="metric-label">活跃任务数</div></div>', unsafe_allow_html=True)
            with m3:
                st.markdown('<div class="metric-card" style="border-bottom-color:#f59e0b;"><div class="metric-value" style="color:#f59e0b;">3</div><div class="metric-label">排队任务数</div></div>', unsafe_allow_html=True)
            with m4:
                st.markdown('<div class="metric-card" style="border-bottom-color:#8b5cf6;"><div class="metric-value" style="color:#8b5cf6;">156ms</div><div class="metric-label">平均响应时间</div></div>', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            agent_status = [
                ("招聘Agent", "🟢 健康", "1.2s", "96%"),
                ("薪酬Agent", "🟢 健康", "0.8s", "99.9%"),
                ("客服Agent", "🟡 负载较高", "4.5s", "87%"),
                ("标书Agent", "🟢 健康", "2.1s", "94%"),
                ("意图Agent", "🟢 健康", "0.3s", "98%"),
                ("财务Agent", "🟢 健康", "1.5s", "99%"),
                ("法务Agent", "🟢 健康", "1.8s", "95%"),
                ("采购Agent", "🟢 健康", "1.6s", "93%"),
                ("生产Agent", "🟢 健康", "2.0s", "91%"),
            ]
            st.markdown("**Agent 运行状态**")
            for name, status, rt, success in agent_status:
                st.markdown(
                    f"""
                    <div style="display:flex;justify-content:space-between;background:white;border-radius:8px;padding:0.75rem;border:1px solid #e2e8f0;margin-bottom:0.5rem;">
                        <span style="font-weight:500;">{name}</span>
                        <span style="color:#64748b;font-size:0.85rem;">{status} · 响应 {rt} · 成功率 {success}</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**记忆核心健康度**")
            mem_metrics = [
                ("遗忘率", "0.6%", "< 1%", "🟢"),
                ("幻觉率", "0.12%", "< 0.2%", "🟢"),
                ("召回延迟", "142ms", "< 200ms", "🟢"),
                ("写入成功率", "99.98%", "> 99.9%", "🟢"),
                ("跨 Agent 共享率", "100%", "100%", "🟢"),
            ]
            for name, val, target, status in mem_metrics:
                st.markdown(
                    f"""
                    <div style="display:flex;justify-content:space-between;background:white;border-radius:8px;padding:0.75rem;border:1px solid #e2e8f0;margin-bottom:0.5rem;">
                        <span>{name}</span>
                        <span><b>{val}</b> <span style="color:#94a3b8;">/ 目标 {target}</span> {status}</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        else:  # 系统配置
            st.markdown(
                """
                <div style="background:linear-gradient(135deg,#fff7ed,#eff6ff);border-radius:16px;padding:1.5rem;border:1px solid #e2e8f0;margin-bottom:1rem;">
                    <h4 style="margin-top:0;color:#c2410c;">系统全局配置</h4>
                    <p style="color:#64748b;margin:0;">管理全局参数、特性开关、记忆核心与原生 Agent 架构开关（演示用途，变更仅在当前会话生效）。</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if "prajna_config" not in st.session_state:
                st.session_state.prajna_config = {
                    "memory_core_enabled": True,
                    "memory_pruning_enabled": True,
                    "memory_reflection_enabled": True,
                    "native_agent_arch_enabled": True,
                    "multi_agent_collaboration_enabled": True,
                    "memory_linkage_required": True,
                    "audit_enabled": True,
                }

            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**记忆核心开关**")
                st.session_state.prajna_config["memory_core_enabled"] = st.toggle("启用记忆核心", value=st.session_state.prajna_config["memory_core_enabled"])
                st.session_state.prajna_config["memory_pruning_enabled"] = st.toggle("启用智能剪枝", value=st.session_state.prajna_config["memory_pruning_enabled"])
                st.session_state.prajna_config["memory_reflection_enabled"] = st.toggle("启用自我反思", value=st.session_state.prajna_config["memory_reflection_enabled"])
            with c2:
                st.markdown("**原生 Agent 架构开关**")
                st.session_state.prajna_config["native_agent_arch_enabled"] = st.toggle("强制遵循原生架构", value=st.session_state.prajna_config["native_agent_arch_enabled"])
                st.session_state.prajna_config["multi_agent_collaboration_enabled"] = st.toggle("允许多 Agent 协同", value=st.session_state.prajna_config["multi_agent_collaboration_enabled"])
                st.session_state.prajna_config["memory_linkage_required"] = st.toggle("强制 Agent 调用记忆核心", value=st.session_state.prajna_config["memory_linkage_required"])

            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("💾 保存配置（演示）", use_container_width=True):
                st.success("配置已热加载，10 秒内通知所有 Agent 订阅者。")
                st.json(st.session_state.prajna_config)

    # ---------- Sub-tab 7: White-box Trace ----------
    with sub_trace:
        st.markdown(
            """
            <div style="background:linear-gradient(135deg,#fef2f2,#eff6ff);border-radius:16px;padding:1.5rem;border:1px solid #e2e8f0;margin-bottom:1rem;">
                <h4 style="margin-top:0;color:#b91c1c;">🔍 白盒追溯：招聘任务完整链路</h4>
                <p style="color:#64748b;margin:0;">每个决策步骤记录思考过程、依据、结果，支持按任务 ID 回放完整执行链路。</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        trace_steps = [
            {
                "time": "T+0ms",
                "step": "1. 用户输入",
                "reasoning": "接收到用户请求：'帮我招聘一名广州 P2 电商运营助理'",
                "memory_id": "—",
                "output": "原始请求文本",
            },
            {
                "time": "T+120ms",
                "step": "2. 意图识别",
                "reasoning": "关键词匹配：招聘、岗位、城市、职级。置信度 0.97。",
                "memory_id": "memory://business/recruitment_patterns",
                "output": "意图=招聘，实体={岗位:电商运营助理, 城市:广州, 职级:P2}",
            },
            {
                "time": "T+250ms",
                "step": "3. 记忆召回",
                "reasoning": "调用 memory_recall(query='电商运营助理 广州 P2', types=['business','agent'])，命中 3 条历史记忆。",
                "memory_id": "memory://agent/recruitment_agent_experience",
                "output": "历史岗位画像 + 广州 P2 薪资带宽 + 既往面试评估维度",
            },
            {
                "time": "T+380ms",
                "step": "4. 调度决策",
                "reasoning": "招聘Agent 能力匹配度 0.98，历史成功率 96%，记忆协同分 0.93，评分最高。",
                "memory_id": "memory://agent/scheduler_decisions",
                "output": "任务分配给 招聘Agent，预期 SLA 120s",
            },
            {
                "time": "T+950ms",
                "step": "5. Agent 执行",
                "reasoning": "根据记忆加载的岗位画像，生成 JD、面试评估表、Offer 薪资建议。",
                "memory_id": "memory://business/job_profile_ecommerce_assistant",
                "output": "招聘套件 Excel + Word 文档",
            },
            {
                "time": "T+1200ms",
                "step": "6. 结果校验",
                "reasoning": "检查输出完整性：JD 6 项职责、面试评估 7 个维度、Offer 薪资在广州 P2 带宽内。",
                "memory_id": "—",
                "output": "校验通过",
            },
            {
                "time": "T+1500ms",
                "step": "7. 记忆沉淀",
                "reasoning": "将岗位画像更新写入 business 记忆，任务结果写入 project 记忆。",
                "memory_id": "memory://business/job_profile_ecommerce_assistant",
                "output": "记忆写入成功",
            },
            {
                "time": "T+1800ms",
                "step": "8. 自我反思",
                "reasoning": "本次生成耗时 1.8s，输出完整。可优化点：JD 中可补充直播运营经验要求。",
                "memory_id": "memory://agent/recruitment_agent_experience",
                "output": "反思记录已更新",
            },
        ]

        for step in trace_steps:
            st.markdown(
                f"""
                <div style="background:white;border-radius:12px;padding:1.25rem;border-left:4px solid #2563eb;box-shadow:0 1px 3px rgba(0,0,0,0.05);margin-bottom:0.75rem;">
                    <div style="display:flex;justify-content:space-between;margin-bottom:0.5rem;">
                        <span style="font-weight:700;color:#1e3a8a;">{step['step']}</span>
                        <span style="color:#94a3b8;font-size:0.85rem;">{step['time']}</span>
                    </div>
                    <div style="color:#475569;font-size:0.9rem;line-height:1.6;">
                        <b>思考：</b>{step['reasoning']}<br>
                        <b>记忆来源：</b><code style="background:#f1f5f9;padding:0.1rem 0.3rem;border-radius:4px;">{step['memory_id']}</code><br>
                        <b>输出：</b>{step['output']}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # ---------- Sub-tab 8: Create New Skill ----------
    with sub_create:
        st.markdown(
            """
            <div style="background:linear-gradient(135deg,#f0f9ff,#f5f3ff);border-radius:16px;padding:1.5rem;border:1px solid #e2e8f0;margin-bottom:1rem;">
                <h4 style="margin-top:0;color:#0369a1;">➕ 一句话创建新 Skill</h4>
                <p style="color:#64748b;margin:0;">用自然语言描述你想创建的 Skill，prajna 自动生成 SKILL.md 框架与 manifest.json。</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        skill_desc = st.text_area(
            "描述你的 Skill",
            value="创建一个合同审查助手，能够自动识别合同中的风险条款、付款条件和违约责任，并生成审查报告。",
            height=100,
        )
        skill_author = st.text_input("作者", value="prajna")
        skill_category = st.selectbox("分类", ["业务", "工具", "系统", "分析", "内容", "其他"])

        if st.button("✨ 生成 Skill 框架", use_container_width=True):
            # Simple heuristic to generate skill_id and name
            desc_lower = skill_desc.lower()
            if "合同" in desc_lower:
                skill_id = "contract-review-assistant"
                skill_name = "合同审查助手"
            elif "报表" in desc_lower or "报告" in desc_lower:
                skill_id = "report-generator"
                skill_name = "报告生成助手"
            elif "客服" in desc_lower or "问答" in desc_lower:
                skill_id = "customer-service-agent"
                skill_name = "智能客服助手"
            elif "数据" in desc_lower or "分析" in desc_lower:
                skill_id = "data-analytics-agent"
                skill_name = "数据分析助手"
            elif "营销" in desc_lower or "文案" in desc_lower:
                skill_id = "marketing-copy-agent"
                skill_name = "营销文案助手"
            else:
                skill_id = "custom-skill"
                skill_name = "自定义技能"

            skill_md = f"""---
name: {skill_name}
skill_id: {skill_id}
description: |
  {skill_desc.strip()}
version: 1.0.0
author: {skill_author}
category: {skill_category}
tags: [skill, {skill_category}]
requires:
  - prajna-memory-core
  - prajna-native-agent-architecture
memory_context:
  - business
  - project
---

# {skill_name}

## 1. 角色定位

{skill_name} 是 prajna 企业智能体平台中的 {skill_category} 类 Skill，{skill_desc.strip()}。

## 2. 核心功能

- 功能 1：接收输入并解析关键参数
- 功能 2：调用记忆核心加载相关上下文
- 功能 3：执行业务逻辑并生成输出
- 功能 4：沉淀执行结果与反思到记忆核心

## 3. 输入输出

| 输入 | 类型 | 说明 |
|------|------|------|
| input_text | string | 用户自然语言输入 |

| 输出 | 类型 | 说明 |
|------|------|------|
| output_file | file | 生成的文档或报告 |

## 4. 记忆联动

- 执行前调用 `memory_recall` 加载 `business` 和 `project` 记忆
- 执行后调用 `memory_reflect` 记录反思

## 5. 免责声明

【人工智能生成-需人工核验】本 Skill 生成的内容仅供参考，不构成法律、财务或商业决策建议。
"""

            manifest_json = f"""{{
  "id": "{skill_id}",
  "name": "{skill_name}",
  "version": "1.0.0",
  "description": "{skill_desc.strip()}",
  "author": "{skill_author}",
  "category": "{skill_category}",
  "tags": ["skill", "{skill_category}"],
  "requires": ["prajna-memory-core", "prajna-native-agent-architecture"],
  "memory_context": ["business", "project"]
}}"""

            st.markdown("**生成的 SKILL.md**")
            st.code(skill_md, language="markdown")
            st.markdown("**生成的 manifest.json**")
            st.code(manifest_json, language="json")

            # Offer downloads
            col1, col2 = st.columns(2)
            with col1:
                st.download_button("📥 下载 SKILL.md", skill_md, file_name="SKILL.md", mime="text/markdown", use_container_width=True)
            with col2:
                st.download_button("📥 下载 manifest.json", manifest_json, file_name="manifest.json", mime="application/json", use_container_width=True)

# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.markdown(
    """
    <div class="footer">
        <strong>prajna 企业智能体平台</strong><br>
        记忆觉醒 · 智能新生｜让记忆型 AI 赋能企业数字化转型<br>
        出品方：惠州市锦辉人力资源管理有限公司 · 20+ 年一线 HRO 运营经验 · AAA 级和谐劳动关系企业<br>
        <small>广州「广智能」超级智能体大赛 · 开放赛道 · 自然语言智能体 · 原生 Agent 架构 · 全模态记忆核心</small>
    </div>
    """,
    unsafe_allow_html=True,
)
st.caption(
    "【人工智能生成-需人工核验】本 Demo 输出内容仅供参考，具体薪酬、销售、财务、招投标、招聘、采购、法务、生产、客服等决策请以当地法律法规及公司政策为准。"
)
