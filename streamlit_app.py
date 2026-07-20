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
SALARY_SCRIPT = APP_DIR / "skills" / "prajna-salary-template" / "scripts" / "generate_salary_template.py"
SALES_SCRIPT = APP_DIR / "skills" / "prajna-sales-weekly-report" / "scripts" / "generate_sales_weekly_report.py"

SALARY_SAMPLE = APP_DIR / "assets" / "sample_薪资模板_广州_电商运营助理.xlsx"
SALES_SAMPLE = APP_DIR / "assets" / "sample_销售周报_示例.xlsx"

SALARY_HISTORY_DIR = Path.home() / ".prajna" / "salary_template_history"
SALES_HISTORY_DIR = Path.home() / ".prajna" / "sales_weekly_report_history"

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
    </style>
    """,
    unsafe_allow_html=True,
)

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

SALARY_KEYWORDS = ["薪资", "工资", "薪酬", "salary", "收入", "待遇", "薪水"]
SALES_KEYWORDS = ["周报", "销售周报", "weekly report", "销售报告", "销售团队", "销售报表"]


def extract_city(text):
    for city in CITIES:
        if city in text:
            return city
    return None


def extract_level(text):
    m = re.search(r"P\d", text.upper())
    return m.group(0) if m else "P2"


def extract_industry(text):
    text = text.lower()
    for industry, kws in INDUSTRIES.items():
        for kw in kws:
            if kw in text:
                return industry
    return "互联网"


def extract_position(text):
    """Simple heuristic: remove known keywords/city/level and pick the noun phrase."""
    cleaned = text
    # remove level
    cleaned = re.sub(r"P\d", "", cleaned, flags=re.I)
    # remove cities
    for city in CITIES:
        cleaned = cleaned.replace(city, "")
    # remove industry keywords
    for kws in INDUSTRIES.values():
        for kw in kws:
            cleaned = cleaned.replace(kw, "")
    # remove salary keywords and generic words
    generic_words = ["模板", "生成", "做", "一份", "的", "和", "与", "给我", "帮我", "请", "需要", "表", "薪资", "工资", "薪酬", "薪水", "周报", "报告", "报表"]
    for kw in SALARY_KEYWORDS + SALES_KEYWORDS + generic_words:
        cleaned = cleaned.replace(kw, "")
    cleaned = re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9]", "", cleaned)
    cleaned = re.sub(r"(表|模板)+$", "", cleaned)
    return cleaned.strip()[:8] or "岗位"


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


def extract_week(text):
    m = re.search(r"第\s*(\d+)\s*周", text)
    if m:
        return f"{datetime.now().year}年第{int(m.group(1))}周"
    return f"{datetime.now().year}年第{datetime.now().isocalendar().week}周"


def infer_sales_preset(text):
    text = text.lower()
    if "saas" in text or "互联网" in text or "科技" in text:
        return "互联网/SaaS 销售团队"
    if "电商" in text:
        return "电商销售团队"
    if "零售" in text or "门店" in text or "线下" in text:
        return "线下零售销售团队"
    return "通用销售团队"


def parse_agent_input(text):
    text = text.strip()
    if not text:
        return None

    lowered = text.lower()
    is_salary = any(kw in text for kw in SALARY_KEYWORDS)
    is_sales = any(kw in text for kw in SALES_KEYWORDS)

    # Default to salary if ambiguous and contains P-level, else sales if sales keyword
    if is_salary and not is_sales:
        intent = "salary"
    elif is_sales and not is_salary:
        intent = "sales"
    elif re.search(r"P\d", text.upper()):
        intent = "salary"
    else:
        intent = "sales"

    city = extract_city(text) or "广州"
    level = extract_level(text) if intent == "salary" else None

    if intent == "salary":
        industry = extract_industry(text)
        position = extract_position(text)
        if position in ["", "岗位"] and industry:
            position = f"{industry}专员"
        return {
            "intent": "salary",
            "industry": industry,
            "position": position,
            "city": city,
            "level": level,
            "description": text,
        }
    else:
        team = text.split("周报")[0].strip() or "销售团队"
        team = re.sub(r"^(帮我|给我|请|做|一份|的|生成)", "", team).strip() or "销售团队"
        # remove week info and trailing '销售'
        team = re.sub(r"(本周|第\s*\d+\s*周|的|销售)+$", "", team).strip()
        team = re.sub(r"^销售", "", team).strip() or "销售团队"
        return {
            "intent": "sales",
            "team": team,
            "preset": infer_sales_preset(text),
            "week": extract_week(text),
            "target": extract_target(text),
            "author": "Prajna",
            "date": datetime.now().strftime("%Y-%m-%d"),
        }


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
        一键生成企业级 Excel 模板的能力。

        ### 当前支持
        - 💰 薪资模板（8 个工作表）
        - 📊 销售周报（5 个工作表）
        - 🤖 自然语言智能体模式
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
        <p>输入参数或自然语言，一键生成 HR 薪资模板与销售团队标准周报</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Quick stats
# ---------------------------------------------------------------------------
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown('<div class="metric-box"><h3>💰</h3><p>薪资模板</p><h5>8 个工作表</h5></div>', unsafe_allow_html=True)
with c2:
    st.markdown('<div class="metric-box"><h3>📊</h3><p>销售周报</p><h5>5 个工作表</h5></div>', unsafe_allow_html=True)
with c3:
    st.markdown('<div class="metric-box"><h3>🤖</h3><p>自然语言理解</p><h5>智能体模式</h5></div>', unsafe_allow_html=True)
with c4:
    st.markdown('<div class="metric-box"><h3>🔗</h3><p>Excel 公式</p><h5>全部可二次编辑</h5></div>', unsafe_allow_html=True)

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
                width="stretch",
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
                width="stretch",
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
    ["🤖 Prajna 智能体模式（自然语言）", "💰 薪资模板", "📊 销售团队标准周报"],
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
            placeholder="例如：帮我做一份深圳互联网产品经理 P5 的薪资模板 / 生成华南销售一部本周销售周报，目标 150 万",
            label_visibility="collapsed",
        )

        st.markdown(
            """
            <small>💡 你可以这样说：</small><br>
            <small>• "帮我做一份上海制造业生产主管 P4 的薪资模板"</small><br>
            <small>• "生成电商销售团队本周周报，目标 80 万"</small>
            """,
            unsafe_allow_html=True,
        )

    with right:
        st.markdown('<div class="thinking-card">', unsafe_allow_html=True)
        st.markdown("#### 🧠 Prajna 思考中...")
        parsed = parse_agent_input(agent_input)
        if parsed:
            if parsed["intent"] == "salary":
                st.markdown(
                    f"""
                    **识别意图：** 生成薪资模板  
                    **行业：** {parsed['industry']}  
                    **岗位：** {parsed['position']}  
                    **城市：** {parsed['city']}  
                    **职级：** {parsed['level']}
                    """
                )
            else:
                st.markdown(
                    f"""
                    **识别意图：** 生成销售周报  
                    **团队：** {parsed['team']}  
                    **类型：** {parsed['preset']}  
                    **周期：** {parsed['week']}  
                    **目标：** ¥{parsed['target']:,.0f}
                    """
                )
        else:
            st.markdown("请输入你想生成的内容...")
        st.markdown("</div>", unsafe_allow_html=True)

    execute_clicked = st.button("🚀 Prajna 执行", type="primary", width="stretch")

    if execute_clicked:
        if not parsed:
            st.error("无法识别输入内容，请尝试更清晰的描述。")
            st.stop()

        if parsed["intent"] == "salary":
            with st.spinner("Prajna 正在生成薪资模板..."):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_name = f"prajna_薪资模板_{parsed['industry']}_{parsed['position']}_{timestamp}.xlsx"
                out_path = SALARY_HISTORY_DIR / out_name
                out_path.parent.mkdir(parents=True, exist_ok=True)

                cmd = [
                    sys.executable,
                    str(SALARY_SCRIPT),
                    "--industry", parsed["industry"],
                    "--position", parsed["position"],
                    "--city", parsed["city"],
                    "--level", parsed["level"],
                    "--description", parsed["description"],
                    "--output", str(out_path),
                ]

                try:
                    result = subprocess.run(
                        cmd,
                        capture_output=True,
                        text=True,
                        check=True,
                        timeout=60,
                    )
                except subprocess.CalledProcessError as e:
                    st.error("生成失败，请检查输入或查看日志。")
                    st.code(e.stderr or e.stdout)
                    st.stop()
                except Exception as e:
                    st.error(f"运行异常：{e}")
                    st.stop()

            if out_path.exists():
                file_size = out_path.stat().st_size
                st.success(f"✅ 薪资模板已生成！（{file_size / 1024:.1f} KB）")
                c1, c2 = st.columns([1, 2])
                with c1:
                    with open(out_path, "rb") as f:
                        st.download_button(
                            label=f"📥 下载 {out_name}",
                            data=f,
                            file_name=out_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            width="stretch",
                        )
                with c2:
                    st.info("Prajna 已自动识别岗位、城市、职级，并生成了包含 8 个工作表的 Excel。")
            else:
                st.error("生成后未找到文件。")
                st.code(result.stdout)
        else:
            with st.spinner("Prajna 正在生成销售周报..."):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                safe_team = parsed["team"].replace(" ", "_")
                out_name = f"prajna_销售周报_{safe_team}_{timestamp}.xlsx"
                out_path = SALES_HISTORY_DIR / out_name
                out_path.parent.mkdir(parents=True, exist_ok=True)

                cmd = [
                    sys.executable,
                    str(SALES_SCRIPT),
                    "--preset", parsed["preset"],
                    "--team", parsed["team"],
                    "--week", parsed["week"],
                    "--sales-target", str(parsed["target"]),
                    "--author", parsed["author"],
                    "--date", parsed["date"],
                    "--output", str(out_path),
                ]

                try:
                    result = subprocess.run(
                        cmd,
                        capture_output=True,
                        text=True,
                        check=True,
                        timeout=60,
                    )
                except subprocess.CalledProcessError as e:
                    st.error("生成失败，请检查输入或查看日志。")
                    st.code(e.stderr or e.stdout)
                    st.stop()
                except Exception as e:
                    st.error(f"运行异常：{e}")
                    st.stop()

            if out_path.exists():
                file_size = out_path.stat().st_size
                st.success(f"✅ 销售周报已生成！（{file_size / 1024:.1f} KB）")
                c1, c2 = st.columns([1, 2])
                with c1:
                    with open(out_path, "rb") as f:
                        st.download_button(
                            label=f"📥 下载 {out_name}",
                            data=f,
                            file_name=out_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            width="stretch",
                        )
                with c2:
                    st.info("Prajna 已自动识别团队类型、周期与销售目标，并生成了包含 5 个工作表的 Excel。")
            else:
                st.error("生成后未找到文件。")
                st.code(result.stdout)

# ---------------------------------------------------------------------------
# Salary template
# ---------------------------------------------------------------------------
elif template_type == "💰 薪资模板":
    st.markdown("<br>", unsafe_allow_html=True)
    left, right = st.columns([2, 1])

    with left:
        st.markdown("#### 2️⃣ 输入岗位信息")

        st.markdown("**快速选择预设：**")
        preset_options = [
            "互联网-电商运营助理",
            "互联网-产品经理",
            "互联网-软件工程师",
            "制造业-生产主管",
            "零售业-门店店长",
            "通用-行政专员",
        ]
        selected_preset = st.selectbox("内置预设（可选，填入下方参数）", ["不使用预设"] + preset_options, label_visibility="collapsed")
        if selected_preset != "不使用预设":
            parts = selected_preset.split("-")
            default_industry, default_position = parts[0], parts[1]
        else:
            default_industry, default_position = "互联网", "电商运营助理"

        col1, col2 = st.columns(2)
        with col1:
            industry = st.text_input("行业", value=default_industry, placeholder="例如：互联网、制造业、零售")
            position = st.text_input("岗位", value=default_position, placeholder="例如：产品经理、生产主管")
            city = st.text_input("城市", value="广州", placeholder="例如：广州、深圳、上海")
        with col2:
            level = st.selectbox("职级", ["P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "P9"], index=1)
            description = st.text_area(
                "自然语言描述（可选）",
                value="",
                placeholder="例如：深圳新能源电池 pack 工艺工程师",
                height=100,
            )

        st.markdown(
            "<small>💡 小提示：如果不填行业/岗位，仅填写自然语言描述，Prajna 也会自动推断岗位类型。</small>",
            unsafe_allow_html=True,
        )

    with right:
        st.markdown('<div class="preview-card">', unsafe_allow_html=True)
        st.markdown("#### 🔍 生成预览")
        st.markdown(
            f"""
            **模板类型：** 薪资模板  
            **行业：** {industry or '自动推断'}  
            **岗位：** {position or '自动推断'}  
            **城市：** {city}  
            **职级：** {level}
            """
        )
        st.markdown("**将生成 8 个工作表：**")
        sheets = [
            "1. 薪资结构",
            "2. 绩效考核",
            "3. 调薪机制",
            "4. 奖金方案",
            "5. 福利明细",
            "6. 计算示例",
            "7. 合规校验",
            "8. 绩效系数对照表",
        ]
        for s in sheets:
            st.markdown(f"<small>{s}</small>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    generate_clicked = st.button("🚀 生成薪资模板", type="primary", width="stretch")

    if generate_clicked:
        if not industry.strip() and not position.strip() and not description.strip():
            st.error("请至少填写行业、岗位或自然语言描述中的一项。")
            st.stop()

        with st.spinner("Prajna 正在生成薪资模板，请稍候..."):
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_name = f"prajna_薪资模板_{industry}_{position}_{timestamp}.xlsx"
            out_path = SALARY_HISTORY_DIR / out_name
            out_path.parent.mkdir(parents=True, exist_ok=True)

            cmd = [
                sys.executable,
                str(SALARY_SCRIPT),
                "--industry", industry.strip(),
                "--position", position.strip(),
                "--city", city.strip(),
                "--level", level,
                "--output", str(out_path),
            ]
            if description.strip():
                cmd += ["--description", description.strip()]

            try:
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    check=True,
                    timeout=60,
                )
            except subprocess.CalledProcessError as e:
                st.error("生成失败，请检查输入或查看日志。")
                st.code(e.stderr or e.stdout)
                st.stop()
            except Exception as e:
                st.error(f"运行异常：{e}")
                st.stop()

        if out_path.exists():
            file_size = out_path.stat().st_size
            st.success(f"✅ 薪资模板已生成！（{file_size / 1024:.1f} KB）")
            c1, c2 = st.columns([1, 2])
            with c1:
                with open(out_path, "rb") as f:
                    st.download_button(
                        label=f"📥 下载 {out_name}",
                        data=f,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width="stretch",
                    )
            with c2:
                st.info(
                    "生成的 Excel 包含 8 个工作表，所有汇总、绩效加权、个税、全年总包均使用 Excel 公式，可二次编辑。"
                )
        else:
            st.error("生成后未找到文件，请检查脚本输出。")
            st.code(result.stdout)

# ---------------------------------------------------------------------------
# Sales weekly report
# ---------------------------------------------------------------------------
else:
    st.markdown("<br>", unsafe_allow_html=True)
    left, right = st.columns([2, 1])

    with left:
        st.markdown("#### 2️⃣ 输入销售周报信息")

        presets = ["互联网/SaaS 销售团队", "电商销售团队", "线下零售销售团队", "通用销售团队"]

        col1, col2 = st.columns(2)
        with col1:
            team = st.text_input("团队名称", value="华南销售一部", placeholder="例如：华南销售一部")
            preset = st.selectbox("团队类型", presets, index=0)
            week = st.text_input("报表周期", value=f"{datetime.now().year}年第{datetime.now().isocalendar().week}周", placeholder="例如：2026年第30周")
        with col2:
            sales_target = st.number_input("本周销售目标（元）", min_value=0, value=1200000, step=10000)
            author = st.text_input("填写人", value="销售主管", placeholder="例如：张经理")
            report_date = st.date_input("填写日期", value=datetime.now()).strftime("%Y-%m-%d")

    with right:
        st.markdown('<div class="preview-card">', unsafe_allow_html=True)
        st.markdown("#### 🔍 生成预览")
        st.markdown(
            f"""
            **模板类型：** 销售团队标准周报  
            **团队：** {team}  
            **类型：** {preset}  
            **周期：** {week}  
            **目标：** ¥{sales_target:,.0f}
            """
        )
        st.markdown("**将生成 5 个工作表：**")
        sheets = [
            "1. 封面汇总",
            "2. 核心业绩数据",
            "3. 重点商机进展",
            "4. 问题分析",
            "5. 下周工作计划",
        ]
        for s in sheets:
            st.markdown(f"<small>{s}</small>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

    generate_clicked = st.button("🚀 生成销售周报", type="primary", width="stretch")

    if generate_clicked:
        if not team.strip():
            st.error("请填写团队名称。")
            st.stop()

        with st.spinner("Prajna 正在生成销售周报，请稍候..."):
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_team = team.strip().replace(" ", "_")
            out_name = f"prajna_销售周报_{safe_team}_{timestamp}.xlsx"
            out_path = SALES_HISTORY_DIR / out_name
            out_path.parent.mkdir(parents=True, exist_ok=True)

            cmd = [
                sys.executable,
                str(SALES_SCRIPT),
                "--preset", preset,
                "--team", team.strip(),
                "--week", week.strip(),
                "--sales-target", str(sales_target),
                "--author", author.strip(),
                "--date", report_date,
                "--output", str(out_path),
            ]

            try:
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    check=True,
                    timeout=60,
                )
            except subprocess.CalledProcessError as e:
                st.error("生成失败，请检查输入或查看日志。")
                st.code(e.stderr or e.stdout)
                st.stop()
            except Exception as e:
                st.error(f"运行异常：{e}")
                st.stop()

        if out_path.exists():
            file_size = out_path.stat().st_size
            st.success(f"✅ 销售周报已生成！（{file_size / 1024:.1f} KB）")
            c1, c2 = st.columns([1, 2])
            with c1:
                with open(out_path, "rb") as f:
                    st.download_button(
                        label=f"📥 下载 {out_name}",
                        data=f,
                        file_name=out_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width="stretch",
                    )
            with c2:
                st.info(
                    "生成的 Excel 包含 5 个工作表，所有完成率、环比、同比、KPI 加权得分均使用 Excel 公式，可二次编辑。"
                )
        else:
            st.error("生成后未找到文件，请检查脚本输出。")
            st.code(result.stdout)

# ---------------------------------------------------------------------------
# Recent files
# ---------------------------------------------------------------------------
st.divider()
with st.expander("🕘 最近生成的文件（本次会话）"):
    salary_files = sorted(SALARY_HISTORY_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)[:5]
    sales_files = sorted(SALES_HISTORY_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)[:5]

    if not salary_files and not sales_files:
        st.info("暂无最近生成的文件。")
    else:
        recent = []
        for f in salary_files + sales_files:
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
    ("🧩", "多场景模板中台", "一个入口同时支持 HR 薪资模板与销售团队周报"),
    ("🎯", "多岗位智能适配", "内置 10 套薪资预设 + 4 种销售团队预设"),
    ("🌍", "城市薪酬基准", "根据城市等级自动调整薪资与社保公积金基数"),
    ("📈", "职级带宽体系", "P1-P9 职级对应薪酬与奖金带宽"),
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
    "【人工智能生成-需人工核验】本 Demo 输出内容仅供参考，具体薪酬、销售决策请以当地法律法规及公司政策为准。"
)
