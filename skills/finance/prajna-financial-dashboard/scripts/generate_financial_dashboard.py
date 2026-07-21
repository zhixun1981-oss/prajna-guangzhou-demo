#!/usr/bin/env python3
"""
prajna-financial-dashboard 生成器 v1.0.0
为管理层生成「财务核心指标分析看板」HTML 页面。
覆盖资产负债率、净利润率、现金流量比率、营业收入、毛利率、ROE、
流动比率、经营现金流净额等关键指标的趋势图表与同比环比分析。

依赖：
    - Python 3.8+
    - jinja2（推荐；未安装时使用内置占位符替换作为降级方案）
"""

import argparse
import json
import math
import os
import re
import sys
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Optional dependencies
# ---------------------------------------------------------------------------
try:
    from jinja2 import Template

    HAVE_JINJA = True
except ImportError:
    HAVE_JINJA = False

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
DATA_DIR = SKILL_DIR / "data"
DEFAULT_SAMPLES_DIR = Path.home() / ".prajna" / "prajna-financial-dashboard" / "samples"

# ---------------------------------------------------------------------------
# Metric definitions
# ---------------------------------------------------------------------------
METRICS = [
    {
        "key": "debt_ratio",
        "name": "资产负债率",
        "unit": "%",
        "fmt": ".2f",
        "invert": True,
        "thresholds": [(50.0, "good"), (65.0, "warning")],
        "desc": "负债总额 / 资产总额",
    },
    {
        "key": "net_profit_margin",
        "name": "净利润率",
        "unit": "%",
        "fmt": ".2f",
        "invert": False,
        "thresholds": [(15.0, "good"), (8.0, "warning")],
        "desc": "净利润 / 营业收入",
    },
    {
        "key": "cash_flow_ratio",
        "name": "现金流量比率",
        "unit": "%",
        "fmt": ".2f",
        "invert": False,
        "thresholds": [(80.0, "good"), (50.0, "warning")],
        "desc": "经营现金流净额 / 流动负债",
    },
    {
        "key": "revenue",
        "name": "营业收入",
        "unit": "万元",
        "fmt": ",.2f",
        "invert": False,
        "thresholds": [],
        "desc": "月度主营业务收入",
    },
    {
        "key": "gross_margin",
        "name": "毛利率",
        "unit": "%",
        "fmt": ".2f",
        "invert": False,
        "thresholds": [(30.0, "good"), (20.0, "warning")],
        "desc": "毛利润 / 营业收入",
    },
    {
        "key": "roe",
        "name": "净资产收益率(年化)",
        "unit": "%",
        "fmt": ".2f",
        "invert": False,
        "thresholds": [(12.0, "good"), (6.0, "warning")],
        "desc": "净利润年化 / 股东权益",
    },
    {
        "key": "current_ratio",
        "name": "流动比率",
        "unit": "倍",
        "fmt": ".2f",
        "invert": False,
        "thresholds": [(2.0, "good"), (1.2, "warning")],
        "desc": "流动资产 / 流动负债",
    },
    {
        "key": "operating_cash_flow",
        "name": "经营现金流净额",
        "unit": "万元",
        "fmt": ",.2f",
        "invert": False,
        "thresholds": [],
        "desc": "经营活动现金流量净额",
    },
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_presets():
    path = DATA_DIR / "financial_presets.json"
    if path.exists():
        try:
            return load_json(path)
        except Exception as exc:
            print(f"Warning: 无法读取预设文件 {path}: {exc}", file=sys.stderr)
    return {}


def list_presets():
    return list(load_presets().keys())


def get_preset(name):
    presets = load_presets()
    if name in presets:
        return dict(presets[name])
    # fuzzy match
    for key, value in presets.items():
        if key in name or name in key:
            return dict(value)
    return None


def _safe_filename(name):
    name = re.sub(r"[\\/:*?\"<>|\s]+", "_", str(name))
    return name.strip("_") or "company"


def _fmt(value, fmt):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "—"
    return format(value, fmt)


def _compute_monthly_metrics(raw_months):
    """由原始财务数据计算各项核心指标。"""
    months = []
    for r in raw_months:
        revenue = r.get("revenue", 0) or 1e-9
        equity = r.get("equity", 0) or 1e-9
        months.append(
            {
                "label": r["label"],
                "revenue": r.get("revenue", 0),
                "gross_profit": r.get("gross_profit", 0),
                "net_profit": r.get("net_profit", 0),
                "gross_margin": r.get("gross_profit", 0) / revenue * 100,
                "net_profit_margin": r.get("net_profit", 0) / revenue * 100,
                "total_assets": r.get("total_assets", 0),
                "total_liabilities": r.get("total_liabilities", 0),
                "debt_ratio": r.get("total_liabilities", 0)
                / max(r.get("total_assets", 0), 1e-9)
                * 100,
                "equity": equity,
                "current_assets": r.get("current_assets", 0),
                "current_liabilities": r.get("current_liabilities", 0),
                "current_ratio": r.get("current_assets", 0)
                / max(r.get("current_liabilities", 0), 1e-9),
                "operating_cash_flow": r.get("operating_cash_flow", 0),
                "cash_flow_ratio": r.get("operating_cash_flow", 0)
                / max(r.get("current_liabilities", 0), 1e-9)
                * 100,
                "roe": r.get("net_profit", 0) * 12 / equity * 100,
            }
        )
    return months


def _delta(cur, prev, metric):
    """返回 (显示文本, 变化方向)。

    百分比/倍数类指标采用绝对差，绝对金额类指标采用相对变化率。
    """
    if prev is None or prev == 0:
        return "—", "neutral"
    invert = metric.get("invert", False)
    if metric["unit"] in ("%", "倍"):
        d = cur - prev
    else:
        d = (cur - prev) / abs(prev) * 100
    if abs(d) < 0.005:
        text = f"0.00{metric['unit']}" if metric["unit"] in ("%", "倍") else "0.00%"
        return text, "neutral"
    text = f"{d:+.2f}{metric['unit']}" if metric["unit"] in ("%", "倍") else f"{d:+.2f}%"
    good = (d > 0 and not invert) or (d < 0 and invert)
    direction = "good" if good else "bad"
    return text, direction


def _status_for_metric(value, metric):
    """根据阈值返回 good / warning / danger。"""
    if not metric["thresholds"]:
        return "normal"
    invert = metric.get("invert", False)
    # 从优到劣排序：非反向指标按阈值降序，反向指标按阈值升序
    sorted_thresholds = sorted(metric["thresholds"], key=lambda x: x[0], reverse=not invert)
    for threshold, level in sorted_thresholds:
        if (not invert and value >= threshold) or (invert and value <= threshold):
            return level
    return "danger"


def _build_cards(latest, prev, yoy):
    cards = []
    for metric in METRICS:
        key = metric["key"]
        value = latest.get(key, 0)
        mom_text, mom_dir = _delta(value, prev.get(key) if prev else None, metric)
        yoy_text, yoy_dir = _delta(value, yoy.get(key) if yoy else None, metric)
        status = _status_for_metric(value, metric)
        cards.append(
            {
                "name": metric["name"],
                "desc": metric["desc"],
                "value": _fmt(value, metric["fmt"]),
                "unit": metric["unit"],
                "mom_text": mom_text,
                "mom_dir": mom_dir,
                "yoy_text": yoy_text,
                "yoy_dir": yoy_dir,
                "status": status,
            }
        )
    return cards


def _build_table_rows(months):
    """构建最近 12 个月的数据表格行（含同比/环比）。"""
    rows = []
    start_idx = max(0, len(months) - 12)
    for i in range(start_idx, len(months)):
        m = months[i]
        prev = months[i - 1] if i > 0 else None
        yoy_m = months[i - 12] if i >= 12 else None
        revenue_mom, _ = _delta(m["revenue"], prev["revenue"] if prev else None, {"unit": "万元", "invert": False})
        revenue_yoy, _ = _delta(m["revenue"], yoy_m["revenue"] if yoy_m else None, {"unit": "万元", "invert": False})
        rows.append(
            {
                "label": m["label"],
                "revenue": _fmt(m["revenue"], ",.2f"),
                "gross_margin": _fmt(m["gross_margin"], ".2f"),
                "net_profit_margin": _fmt(m["net_profit_margin"], ".2f"),
                "debt_ratio": _fmt(m["debt_ratio"], ".2f"),
                "current_ratio": _fmt(m["current_ratio"], ".2f"),
                "cash_flow_ratio": _fmt(m["cash_flow_ratio"], ".2f"),
                "operating_cash_flow": _fmt(m["operating_cash_flow"], ",.2f"),
                "roe": _fmt(m["roe"], ".2f"),
                "revenue_mom": revenue_mom,
                "revenue_yoy": revenue_yoy,
            }
        )
    return rows


def _health_rating_and_observations(latest):
    """基于最新指标给出健康评级与经营观察。"""
    score = 0
    observations = []

    debt = latest.get("debt_ratio", 0)
    if debt < 50:
        score += 2
        observations.append(f"资产负债率为 <strong>{debt:.2f}%</strong>，处于稳健区间，长期偿债风险可控。")
    elif debt < 65:
        score += 1
        observations.append(f"资产负债率为 <strong>{debt:.2f}%</strong>，建议关注有息负债结构与到期偿付安排。")
    else:
        observations.append(f"资产负债率为 <strong>{debt:.2f}%</strong>，高于警戒水平，需警惕财务杠杆风险。")

    npm = latest.get("net_profit_margin", 0)
    if npm >= 15:
        score += 2
        observations.append(f"净利润率达到 <strong>{npm:.2f}%</strong>，盈利能力表现优秀。")
    elif npm >= 8:
        score += 1
        observations.append(f"净利润率为 <strong>{npm:.2f}%</strong>，盈利能力处于行业中等水平。")
    else:
        observations.append(f"净利润率仅 <strong>{npm:.2f}%</strong>，建议排查成本费用结构与定价策略。")

    cfr = latest.get("cash_flow_ratio", 0)
    if cfr >= 80:
        score += 2
        observations.append(f"现金流量比率 <strong>{cfr:.2f}%</strong>，经营现金流对短期负债覆盖充足。")
    elif cfr >= 50:
        score += 1
        observations.append(f"现金流量比率 <strong>{cfr:.2f}%</strong>，短期流动性尚可，但需关注回款节奏。")
    else:
        observations.append(f"现金流量比率 <strong>{cfr:.2f}%</strong> 偏低，短期偿债依赖外部融资，需重点跟进。")

    current = latest.get("current_ratio", 0)
    if current >= 2:
        score += 1
        observations.append(f"流动比率 <strong>{current:.2f}倍</strong>，短期偿债能力良好。")
    elif current >= 1.2:
        observations.append(f"流动比率 <strong>{current:.2f}倍</strong>，处于常规水平。")
    else:
        observations.append(f"流动比率 <strong>{current:.2f}倍</strong> 偏低，存在短期流动性压力。")

    roe = latest.get("roe", 0)
    if roe >= 12:
        score += 1
        observations.append(f"年化 ROE 为 <strong>{roe:.2f}%</strong>，股东回报水平较高。")
    elif roe < 6:
        observations.append(f"年化 ROE 为 <strong>{roe:.2f}%</strong>，资本回报效率有待提升。")

    if score >= 6:
        rating = "A · 健康"
        summary = "整体财务结构稳健，盈利能力与现金流状况良好，可适度把握扩张机会。"
    elif score >= 4:
        rating = "B · 关注"
        summary = "财务基本面尚可，但部分指标出现边际弱化，建议加强营运资本与成本管控。"
    else:
        rating = "C · 预警"
        summary = "多项核心指标亮起黄灯，建议管理层立即组织专项复盘并制定改善计划。"
    return rating, summary, observations


# ---------------------------------------------------------------------------
# Render helpers
# ---------------------------------------------------------------------------
def _render_jinja_or_fallback(template_str, values):
    if HAVE_JINJA:
        return Template(template_str).render(values)

    # 降级方案：简单的占位符替换，适用于未安装 jinja2 的环境。
    print("Warning: jinja2 未安装，使用内置占位符渲染（建议安装 jinja2 以获得最佳体验）。", file=sys.stderr)

    def replacer(match):
        key = match.group(1).strip()
        return str(values.get(key, match.group(0)))

    return re.sub(r"\{\{\s*(\w+)\s*(?:\|[^}]*)?\s*\}\}", replacer, template_str)


# ---------------------------------------------------------------------------
# HTML template
# ---------------------------------------------------------------------------
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>{{ title }}</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.3/dist/chart.umd.min.js"></script>
  <style>
    :root {
      --primary: #1a2b4a;
      --primary-light: #2c4a75;
      --accent: #2563eb;
      --accent-light: #60a5fa;
      --success: #059669;
      --warning: #d97706;
      --danger: #dc2626;
      --bg: #f3f4f6;
      --card: #ffffff;
      --text: #1f2937;
      --muted: #6b7280;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.6;
    }
    header {
      background: linear-gradient(135deg, #0f2027 0%, #203a43 50%, #2c5364 100%);
      color: #fff;
      padding: 40px 30px 30px;
      text-align: center;
      box-shadow: 0 4px 20px rgba(0,0,0,0.15);
    }
    header h1 { margin: 0 0 8px; font-size: 32px; letter-spacing: 1px; }
    header p { margin: 0; opacity: 0.9; font-size: 15px; }
    .container { max-width: 1400px; margin: 0 auto; padding: 24px; }
    .summary-bar {
      background: var(--card);
      border-radius: 12px;
      padding: 20px 24px;
      margin-bottom: 24px;
      box-shadow: 0 2px 12px rgba(0,0,0,0.06);
      display: flex;
      flex-wrap: wrap;
      gap: 20px;
      align-items: center;
      justify-content: space-between;
    }
    .rating {
      display: inline-flex;
      align-items: center;
      gap: 12px;
    }
    .rating-badge {
      font-size: 22px;
      font-weight: 700;
      padding: 6px 18px;
      border-radius: 20px;
      color: #fff;
    }
    .rating-good { background: var(--success); }
    .rating-warning { background: var(--warning); }
    .rating-danger { background: var(--danger); }
    .observations { flex: 1 1 400px; }
    .observations li { margin-bottom: 6px; }
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
      gap: 20px;
      margin-bottom: 24px;
    }
    .kpi-card {
      background: var(--card);
      border-radius: 12px;
      padding: 20px;
      box-shadow: 0 2px 12px rgba(0,0,0,0.06);
      border-left: 5px solid var(--accent);
      transition: transform 0.15s ease;
    }
    .kpi-card:hover { transform: translateY(-3px); }
    .kpi-card.status-good { border-left-color: var(--success); }
    .kpi-card.status-warning { border-left-color: var(--warning); }
    .kpi-card.status-danger { border-left-color: var(--danger); }
    .kpi-name { font-size: 14px; color: var(--muted); margin-bottom: 6px; }
    .kpi-desc { font-size: 12px; color: #9ca3af; margin-bottom: 10px; }
    .kpi-value { font-size: 30px; font-weight: 700; color: var(--primary); }
    .kpi-unit { font-size: 14px; color: var(--muted); margin-left: 4px; }
    .kpi-deltas { margin-top: 12px; display: flex; gap: 16px; font-size: 13px; }
    .delta { display: flex; align-items: center; gap: 4px; }
    .delta-good { color: var(--success); font-weight: 600; }
    .delta-bad { color: var(--danger); font-weight: 600; }
    .delta-neutral { color: var(--muted); }
    .chart-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(520px, 1fr));
      gap: 24px;
      margin-bottom: 24px;
    }
    @media (max-width: 600px) {
      .chart-grid { grid-template-columns: 1fr; }
      header h1 { font-size: 24px; }
    }
    .chart-card {
      background: var(--card);
      border-radius: 12px;
      padding: 20px;
      box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    }
    .chart-card h3 { margin: 0 0 16px; font-size: 16px; color: var(--primary); }
    .chart-wrapper { position: relative; height: 320px; }
    .table-card {
      background: var(--card);
      border-radius: 12px;
      padding: 20px;
      box-shadow: 0 2px 12px rgba(0,0,0,0.06);
      overflow-x: auto;
    }
    .table-card h3 { margin: 0 0 16px; font-size: 16px; color: var(--primary); }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td { padding: 10px 12px; text-align: right; border-bottom: 1px solid #e5e7eb; white-space: nowrap; }
    th { background: #f9fafb; color: var(--primary); font-weight: 600; text-align: right; }
    td:first-child, th:first-child { text-align: left; }
    tr:hover td { background: #f9fafb; }
    .footer {
      text-align: center;
      color: var(--muted);
      font-size: 12px;
      margin: 30px 0;
    }
  </style>
</head>
<body>
  <header>
    <h1>财务核心指标分析看板</h1>
    <p>{{ company }} · {{ period }} · 生成于 {{ generated_at }}</p>
  </header>

  <div class="container">
    <section class="summary-bar">
      <div class="rating">
        <span>财务健康评级</span>
        <span class="rating-badge {{ rating_class }}">{{ health_rating }}</span>
      </div>
      <div class="observations">
        <p><strong>经营洞察：</strong>{{ health_summary }}</p>
        <ul>{{ observations_html }}</ul>
      </div>
    </section>

    <section class="kpi-grid">
      {{ kpi_cards_html }}
    </section>

    <section class="chart-grid">
      <div class="chart-card">
        <h3>资产负债率趋势</h3>
        <div class="chart-wrapper"><canvas id="debtChart"></canvas></div>
      </div>
      <div class="chart-card">
        <h3>营业收入与净利润率</h3>
        <div class="chart-wrapper"><canvas id="revenueChart"></canvas></div>
      </div>
      <div class="chart-card">
        <h3>现金流量比率与经营现金流净额</h3>
        <div class="chart-wrapper"><canvas id="cashFlowChart"></canvas></div>
      </div>
      <div class="chart-card">
        <h3>营业收入同比增速</h3>
        <div class="chart-wrapper"><canvas id="yoyChart"></canvas></div>
      </div>
    </section>

    <section class="table-card">
      <h3>近 12 个月核心指标明细（含同比/环比）</h3>
      <table>
        <thead>
          <tr>
            <th>月份</th>
            <th>营业收入(万元)</th>
            <th>毛利率(%)</th>
            <th>净利润率(%)</th>
            <th>资产负债率(%)</th>
            <th>流动比率(倍)</th>
            <th>现金流量比率(%)</th>
            <th>经营现金流净额(万元)</th>
            <th>ROE年化(%)</th>
            <th>营收环比</th>
            <th>营收同比</th>
          </tr>
        </thead>
        <tbody>
          {{ table_rows_html }}
        </tbody>
      </table>
    </section>

    <div class="footer">
      <p>本看板由 prajna 企智 prajna-financial-dashboard 技能自动生成，示例数据仅供经营分析参考，不构成投资决策依据。</p>
    </div>
  </div>

  <script>
    Chart.defaults.font.family = "'PingFang SC','Microsoft YaHei','Segoe UI',sans-serif";
    Chart.defaults.color = '#4b5563';
    const labels = {{ labels_json|safe }};
    const debtRatio = {{ debt_ratio_json|safe }};
    const debtGoodLine = {{ debt_good_line_json|safe }};
    const debtWarningLine = {{ debt_warning_line_json|safe }};
    const revenue = {{ revenue_json|safe }};
    const netMargin = {{ net_margin_json|safe }};
    const cashFlowRatio = {{ cash_flow_ratio_json|safe }};
    const operatingCashFlow = {{ operating_cash_flow_json|safe }};
    const yoyRevenue = {{ yoy_revenue_json|safe }};

    new Chart(document.getElementById('debtChart'), {
      type: 'line',
      data: {
        labels: labels,
        datasets: [
          { label: '资产负债率(%)', data: debtRatio, borderColor: '#2563eb', backgroundColor: 'rgba(37,99,235,0.1)', fill: true, tension: 0.3, pointRadius: 3 },
          { label: '稳健线 50%', data: debtGoodLine, borderColor: '#059669', borderDash: [5,5], pointRadius: 0, fill: false },
          { label: '警戒线 65%', data: debtWarningLine, borderColor: '#dc2626', borderDash: [5,5], pointRadius: 0, fill: false }
        ]
      },
      options: { responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false }, plugins: { legend: { position: 'bottom' } }, scales: { y: { beginAtZero: false, suggestedMax: 80 } } }
    });

    new Chart(document.getElementById('revenueChart'), {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [
          { label: '营业收入(万元)', data: revenue, backgroundColor: 'rgba(37,99,235,0.7)', yAxisID: 'y', order: 2 },
          { label: '净利润率(%)', data: netMargin, type: 'line', borderColor: '#d97706', backgroundColor: '#d97706', yAxisID: 'y1', tension: 0.3, pointRadius: 3, order: 1 }
        ]
      },
      options: { responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false }, plugins: { legend: { position: 'bottom' } }, scales: { y: { beginAtZero: true, title: { display: true, text: '营业收入(万元)' } }, y1: { position: 'right', beginAtZero: true, title: { display: true, text: '净利润率(%)' }, grid: { drawOnChartArea: false } } } }
    });

    new Chart(document.getElementById('cashFlowChart'), {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [
          { label: '经营现金流净额(万元)', data: operatingCashFlow, backgroundColor: 'rgba(5,150,105,0.7)', yAxisID: 'y', order: 2 },
          { label: '现金流量比率(%)', data: cashFlowRatio, type: 'line', borderColor: '#7c3aed', backgroundColor: '#7c3aed', yAxisID: 'y1', tension: 0.3, pointRadius: 3, order: 1 }
        ]
      },
      options: { responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false }, plugins: { legend: { position: 'bottom' } }, scales: { y: { beginAtZero: true, title: { display: true, text: '经营现金流净额(万元)' } }, y1: { position: 'right', beginAtZero: true, title: { display: true, text: '现金流量比率(%)' }, grid: { drawOnChartArea: false } } } }
    });

    new Chart(document.getElementById('yoyChart'), {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [{ label: '营收同比增速(%)', data: yoyRevenue, backgroundColor: yoyRevenue.map(v => v >= 0 ? 'rgba(5,150,105,0.7)' : 'rgba(220,38,38,0.7)'), borderRadius: 4 }]
      },
      options: { responsive: true, maintainAspectRatio: false, interaction: { mode: 'index', intersect: false }, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: false, title: { display: true, text: '同比增速(%)' } } } }
    });
  </script>
</body>
</html>
"""


def _build_html(values):
    return _render_jinja_or_fallback(HTML_TEMPLATE, values)


# ---------------------------------------------------------------------------
# Main builder
# ---------------------------------------------------------------------------
def build_dashboard(params, output_path):
    raw_months = params.get("months", [])
    if len(raw_months) < 2:
        raise ValueError("财务数据至少需要两个月的历史记录")

    months = _compute_monthly_metrics(raw_months)
    latest = months[-1]
    prev = months[-2]
    yoy = months[-13] if len(months) >= 13 else None

    # 用于图表的最近 12 个月数据
    chart_months = months[-12:]
    labels = [m["label"] for m in chart_months]

    health_rating, health_summary, observations = _health_rating_and_observations(latest)
    rating_class = "rating-good" if health_rating.startswith("A") else "rating-warning" if health_rating.startswith("B") else "rating-danger"

    cards = _build_cards(latest, prev, yoy)
    kpi_cards_html = "\n".join(
        f"""<div class=\"kpi-card status-{card['status']}\">
  <div class=\"kpi-name\">{card['name']}</div>
  <div class=\"kpi-desc\">{card['desc']}</div>
  <div class=\"kpi-value\">{card['value']}<span class=\"kpi-unit\">{card['unit']}</span></div>
  <div class=\"kpi-deltas\">
    <span class=\"delta delta-{card['mom_dir']}\">环比 {card['mom_text']}</span>
    <span class=\"delta delta-{card['yoy_dir']}\">同比 {card['yoy_text']}</span>
  </div>
</div>"""
        for card in cards
    )

    table_rows = _build_table_rows(months)
    table_rows_html = "\n".join(
        f"""<tr>
  <td>{r['label']}</td>
  <td>{r['revenue']}</td>
  <td>{r['gross_margin']}</td>
  <td>{r['net_profit_margin']}</td>
  <td>{r['debt_ratio']}</td>
  <td>{r['current_ratio']}</td>
  <td>{r['cash_flow_ratio']}</td>
  <td>{r['operating_cash_flow']}</td>
  <td>{r['roe']}</td>
  <td class=\"{'delta-good' if r['revenue_mom'].startswith('+') else 'delta-bad' if r['revenue_mom'].startswith('-') else 'delta-neutral'}\">{r['revenue_mom']}</td>
  <td class=\"{'delta-good' if r['revenue_yoy'].startswith('+') else 'delta-bad' if r['revenue_yoy'].startswith('-') else 'delta-neutral'}\">{r['revenue_yoy']}</td>
</tr>"""
        for r in table_rows
    )

    observations_html = "\n".join(f"<li>{obs}</li>" for obs in observations)

    # 计算营收同比增速序列（最近 12 个月）
    yoy_revenue = []
    for i, m in enumerate(chart_months):
        idx = len(months) - 12 + i
        if idx >= 12:
            base = months[idx - 12]["revenue"]
            yoy_revenue.append(round((m["revenue"] - base) / abs(base) * 100, 2) if base else 0)
        else:
            yoy_revenue.append(0)

    values = {
        "title": f"财务核心指标分析看板 - {params.get('company', '')}",
        "company": params.get("company", "示范企业"),
        "period": params.get("period", f"{latest['label']} 月度经营看板"),
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "health_rating": health_rating,
        "health_summary": health_summary,
        "rating_class": rating_class,
        "observations_html": observations_html,
        "kpi_cards_html": kpi_cards_html,
        "table_rows_html": table_rows_html,
        "labels_json": json.dumps(labels, ensure_ascii=False),
        "debt_ratio_json": json.dumps([round(m["debt_ratio"], 2) for m in chart_months], ensure_ascii=False),
        "debt_good_line_json": json.dumps([50.0] * len(chart_months), ensure_ascii=False),
        "debt_warning_line_json": json.dumps([65.0] * len(chart_months), ensure_ascii=False),
        "revenue_json": json.dumps([round(m["revenue"], 2) for m in chart_months], ensure_ascii=False),
        "net_margin_json": json.dumps([round(m["net_profit_margin"], 2) for m in chart_months], ensure_ascii=False),
        "cash_flow_ratio_json": json.dumps([round(m["cash_flow_ratio"], 2) for m in chart_months], ensure_ascii=False),
        "operating_cash_flow_json": json.dumps([round(m["operating_cash_flow"], 2) for m in chart_months], ensure_ascii=False),
        "yoy_revenue_json": json.dumps(yoy_revenue, ensure_ascii=False),
    }

    html = _build_html(values)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    return output_path


def build_params(args):
    preset_name = args.preset or "通用企业"
    preset = get_preset(preset_name)
    if not preset:
        print(f"Error: 未知预设 '{preset_name}'", file=sys.stderr)
        sys.exit(2)
    params = preset
    if args.company:
        params["company"] = args.company
    if args.period:
        params["period"] = args.period
    params.setdefault("company", "示范企业")
    return params


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="生成财务核心指标分析看板 HTML 页面")
    parser.add_argument(
        "--preset", "-p",
        help=f"财务预设类型，可选：{', '.join(list_presets())}，默认通用企业"
    )
    parser.add_argument(
        "--company", "-c",
        help="企业名称（覆盖预设默认值）"
    )
    parser.add_argument(
        "--period", "-r",
        help="报表周期描述，例如 '2026年6月'"
    )
    parser.add_argument(
        "--output", "-o",
        help="输出 HTML 文件完整路径（覆盖默认路径）"
    )
    parser.add_argument(
        "--list-presets", action="store_true",
        help="列出所有可用预设"
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)

    if args.list_presets:
        print("可用财务看板预设：")
        for name in list_presets():
            print(f"  - {name}")
        return 0

    params = build_params(args)

    if args.output:
        output_path = Path(args.output)
    else:
        DEFAULT_SAMPLES_DIR.mkdir(parents=True, exist_ok=True)
        months = params.get("months", [])
        latest_label = months[-1]["label"].replace("-", "") if months else datetime.now().strftime("%Y%m")
        safe_company = _safe_filename(params.get("company", "公司"))
        filename = f"prajna_财务核心指标看板_{safe_company}_{latest_label}.html"
        output_path = DEFAULT_SAMPLES_DIR / filename

    if output_path.suffix.lower() != ".html":
        output_path = output_path.with_suffix(".html")

    build_dashboard(params, output_path)
    print(f"已生成财务核心指标看板：{output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
