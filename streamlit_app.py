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
SKILL_SCRIPT = APP_DIR / "skill" / "scripts" / "generate_salary_template.py"
SAMPLE_FILE = APP_DIR / "assets" / "sample_薪资模板_广州_电商运营助理.xlsx"
HISTORY_DIR = Path.home() / ".prajna" / "salary_template_history"

st.set_page_config(
    page_title="Prajna 企业智能体 · HR 薪酬场景 Demo",
    page_icon="🧠",
    layout="centered",
)

# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("🧠 Prajna 企业智能体 — HR 薪酬场景 Demo")
st.markdown(
    "本 Demo 展示 **Prajna** 为任意行业/岗位一键生成完整薪资结构模板的能力。"
    "输入岗位信息，即可在线生成包含薪资结构、绩效考核、调薪机制、奖金方案、"
    "福利明细、计算示例、合规校验的 Excel 模板。"
)

with st.expander("📎 快速体验：下载示例模板"):
    if SAMPLE_FILE.exists():
        with open(SAMPLE_FILE, "rb") as f:
            st.download_button(
                label="下载示例：广州 · 电商运营助理",
                data=f,
                file_name=SAMPLE_FILE.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.warning("示例文件未找到，请通过下方表单生成。")

st.divider()

st.subheader("1️⃣ 输入岗位信息")

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

st.divider()

# ---------------------------------------------------------------------------
# Generation
# ---------------------------------------------------------------------------
if generate_clicked:
    if not industry.strip() and not position.strip() and not description.strip():
        st.error("请至少填写行业、岗位或自然语言描述中的一项。")
        st.stop()

    with st.spinner("Prajna 正在生成薪资模板，请稍候..."):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"prajna_薪资模板_{industry}_{position}_{timestamp}.xlsx"
        out_path = HISTORY_DIR / out_name
        out_path.parent.mkdir(parents=True, exist_ok=True)

        cmd = [
            sys.executable,
            str(SKILL_SCRIPT),
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

st.divider()

st.subheader("2️⃣ 这个 Demo 展示了什么？")
st.markdown(
    """
- **多行业/多岗位适配**：内置 10 套预设，支持自然语言岗位描述自动推断。
- **城市薪酬基准**：根据城市等级自动调整薪资、社保公积金基数。
- **职级带宽**：P1-P9 职级对应薪酬与奖金带宽。
- **合规自动校验**：最低工资、社保基数、公积金比例、绩效占比等 7 项校验。
- **公式联动**：所有汇总、绩效加权、个税、全年总包均使用 Excel 公式。
"""
)

st.caption(
    "【人工智能生成-需人工核验】本 Demo 输出内容仅供参考，具体薪酬决策请以当地法律法规及公司政策为准。"
)
