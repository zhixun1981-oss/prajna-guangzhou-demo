import os
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

st.set_page_config(
    page_title="Prajna 企业智能体 · 多场景模板中台 Demo",
    page_icon="🧠",
    layout="centered",
)

# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("🧠 Prajna 企业智能体 — 多场景模板中台 Demo")
st.markdown(
    "本 Demo 展示 **Prajna** 作为企业多智能体模板中台的能力："
    "输入简单参数，即可一键生成结构完整、公式联动的 **Excel 薪资模板** 或 **销售团队标准周报**。"
)

with st.expander("📎 快速体验：下载示例模板"):
    c1, c2 = st.columns(2)
    if SALARY_SAMPLE.exists():
        with c1:
            with open(SALARY_SAMPLE, "rb") as f:
                st.download_button(
                    label="下载示例：广州 · 电商运营助理薪资模板",
                    data=f,
                    file_name=SALARY_SAMPLE.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
    else:
        c1.warning("薪资示例文件未找到。")

    if SALES_SAMPLE.exists():
        with c2:
            with open(SALES_SAMPLE, "rb") as f:
                st.download_button(
                    label="下载示例：销售团队标准周报",
                    data=f,
                    file_name=SALES_SAMPLE.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
    else:
        c2.warning("周报示例文件未找到。")

st.divider()

st.subheader("1️⃣ 选择模板类型")
template_type = st.radio(
    "要生成哪种模板？",
    ["💰 薪资模板", "📊 销售团队标准周报"],
    horizontal=True,
)

# ---------------------------------------------------------------------------
# Salary template inputs
# ---------------------------------------------------------------------------
if template_type == "💰 薪资模板":
    st.subheader("2️⃣ 输入岗位信息")

    col1, col2 = st.columns(2)
    with col1:
        industry = st.text_input("行业", value="互联网", placeholder="例如：互联网、制造业、零售")
        position = st.text_input("岗位", value="电商运营助理", placeholder="例如：产品经理、生产主管")
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
        "<small>如果不填行业/岗位，仅填写自然语言描述，Prajna 也会自动推断岗位类型。</small>",
        unsafe_allow_html=True,
    )

    generate_clicked = st.button("🚀 生成薪资模板", type="primary")

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
            st.success("✅ 薪资模板已生成！")
            with open(out_path, "rb") as f:
                st.download_button(
                    label=f"📥 下载 {out_name}",
                    data=f,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            st.info(
                "生成的 Excel 包含 8 个工作表：薪资结构、绩效考核、调薪机制、奖金方案、"
                "福利明细、计算示例、合规校验、绩效系数对照表。所有汇总均使用 Excel 公式，可二次编辑。"
            )
        else:
            st.error("生成后未找到文件，请检查脚本输出。")
            st.code(result.stdout)

# ---------------------------------------------------------------------------
# Sales weekly report inputs
# ---------------------------------------------------------------------------
else:
    st.subheader("2️⃣ 输入销售周报信息")

    # Load presets for display
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

    generate_clicked = st.button("🚀 生成销售周报", type="primary")

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
            st.success("✅ 销售周报已生成！")
            with open(out_path, "rb") as f:
                st.download_button(
                    label=f"📥 下载 {out_name}",
                    data=f,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            st.info(
                "生成的 Excel 包含 5 个工作表：封面汇总、核心业绩数据、重点商机进展、问题分析、下周工作计划。"
                "所有完成率、环比、同比、KPI 加权得分均使用 Excel 公式，可二次编辑。"
            )
        else:
            st.error("生成后未找到文件，请检查脚本输出。")
            st.code(result.stdout)

st.divider()

st.subheader("3️⃣ 这个 Demo 展示了什么？")
st.markdown(
    """
- **多场景模板中台**：一个入口同时支持 HR 薪资模板与销售团队周报两类高频企业文档。
- **多行业/多岗位适配**：薪资模板内置 10 套预设，支持自然语言岗位描述自动推断。
- **城市薪酬基准**：根据城市等级自动调整薪资、社保公积金基数。
- **职级带宽**：P1-P9 职级对应薪酬与奖金带宽。
- **销售漏斗与业绩追踪**：周报内置核心 KPI、重点商机、问题分析、下周计划。
- **公式联动**：所有汇总、绩效加权、个税、全年总包、完成率均使用 Excel 公式。
"""
)

st.caption(
    "【人工智能生成-需人工核验】本 Demo 输出内容仅供参考，具体薪酬、销售决策请以当地法律法规及公司政策为准。"
)
