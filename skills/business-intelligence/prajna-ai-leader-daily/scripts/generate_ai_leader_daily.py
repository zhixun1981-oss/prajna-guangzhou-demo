#!/usr/bin/env python3
"""
prajna-ai-leader-daily 生成器 v1.0.0
一键生成「AI 领袖动态日报」HTML 页面。
支持：指定日期、自定义输出路径、自定义新闻数据、Jinja2/纯 Python 双渲染引擎。
"""

import argparse
import html
import json
import os
import sys
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
DATA_DIR = SKILL_DIR / "data"
DEFAULT_DATA = DATA_DIR / "ai_leader_curated.json"
SAMPLES_DIR = Path.home() / ".prajna" / "prajna-ai-leader-daily" / "samples"


# ---------------------------------------------------------------------------
# HTML template (uses bracket notation to avoid dict key / method name clash)
# ---------------------------------------------------------------------------
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>AI 领袖动态日报 | {{ date }}</title>
</head>
<body style="margin:0;padding:0;background-color:#f4f6f8;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','PingFang SC','Microsoft YaHei',sans-serif;color:#333;">
  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
    <tr>
      <td align="center" style="padding:24px 16px;">
        <table role="presentation" width="640" cellspacing="0" cellpadding="0" border="0" style="max-width:640px;width:100%;background-color:#ffffff;border-radius:8px;overflow:hidden;">
          <tr>
            <td style="padding:28px 24px 20px;border-bottom:2px solid #1a73e8;text-align:center;">
              <h1 style="margin:0 0 8px;font-size:24px;color:#1a237e;">&#127760; AI 领袖动态日报</h1>
              <p style="margin:0;font-size:14px;color:#666;">{{ date }} · Prajna 企智</p>
            </td>
          </tr>
          <tr>
            <td style="padding:20px 24px 8px;">
              <p style="margin:0;font-size:14px;color:#555;line-height:1.7;">
                精选 <strong>李开复 / 零一万物、Sam Altman / OpenAI、Anthropic、Google DeepMind、Meta</strong> 的最新技术与战略动态，帮您快速掌握全球 AI 风向。
              </p>
            </td>
          </tr>
          {% for leader in leaders %}
          <tr>
            <td style="padding:16px 24px 8px;">
              <h2 style="margin:0 0 12px;font-size:18px;color:{{ leader['color'] }};padding-left:10px;border-left:4px solid {{ leader['color'] }};">
                {{ leader['name'] }}
              </h2>
              {% for item in leader['items'] %}
              <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="margin-bottom:12px;background-color:#fafafa;border:1px solid #e0e0e0;border-left:4px solid {{ leader['color'] }};border-radius:8px;">
                <tr>
                  <td style="padding:14px 16px;">
                    <h3 style="margin:0 0 8px;font-size:15px;color:#222;">{{ item['title'] }}</h3>
                    <p style="margin:0 0 10px;font-size:14px;line-height:1.7;color:#444;">{{ item['summary'] }}</p>
                    <p style="margin:0;font-size:12px;color:#888;">
                      来源：<a href="{{ item['source_url'] }}" style="color:#1a73e8;text-decoration:none;" target="_blank" rel="noopener">{{ item['source_title'] }}</a>
                      {% if item.get('date') %} · {{ item['date'] }}{% endif %}
                    </p>
                  </td>
                </tr>
              </table>
              {% endfor %}
            </td>
          </tr>
          {% endfor %}
          <tr>
            <td style="padding:16px 24px 28px;border-top:1px solid #e0e0e0;">
              <p style="margin:0;font-size:12px;color:#888;line-height:1.6;">
                <strong>免责声明：</strong>本日报基于公开网络信息整理，仅供信息参考，不构成投资建议或商业决策依据。新闻摘要与来源链接可能随时间变化，请自行核实原始报道。
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def load_json(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _today() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def _escape(value: str) -> str:
    return html.escape(str(value), quote=True)


def _render_jinja2(template_str: str, context: dict) -> str:
    from jinja2 import Environment, BaseLoader
    env = Environment(loader=BaseLoader(), autoescape=True)
    template = env.from_string(template_str)
    return template.render(context)


def _render_fallback(context: dict) -> str:
    """Jinja2 缺失或渲染失败时的纯 Python 降级渲染。"""
    date_str = _escape(context["date"])
    leaders = context.get("leaders", [])

    leader_blocks = []
    for leader in leaders:
        name = _escape(leader.get("name", ""))
        color = _escape(leader.get("color", "#1a73e8"))
        items_html = []
        for item in leader.get("items", []):
            title = _escape(item.get("title", ""))
            summary = _escape(item.get("summary", ""))
            source_title = _escape(item.get("source_title", ""))
            source_url = _escape(item.get("source_url", "#"))
            item_date = item.get("date")
            date_html = f" · {_escape(item_date)}" if item_date else ""
            items_html.append(
                f"""<table role=\"presentation\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" style=\"margin-bottom:12px;background-color:#fafafa;border:1px solid #e0e0e0;border-left:4px solid {color};border-radius:8px;\">
  <tr><td style=\"padding:14px 16px;\">
    <h3 style=\"margin:0 0 8px;font-size:15px;color:#222;\">{title}</h3>
    <p style=\"margin:0 0 10px;font-size:14px;line-height:1.7;color:#444;\">{summary}</p>
    <p style=\"margin:0;font-size:12px;color:#888;\">来源：<a href=\"{source_url}\" style=\"color:#1a73e8;text-decoration:none;\" target=\"_blank\" rel=\"noopener\">{source_title}</a>{date_html}</p>
  </td></tr>
</table>"""
            )
        leader_blocks.append(
            f"""<tr><td style=\"padding:16px 24px 8px;\">
  <h2 style=\"margin:0 0 12px;font-size:18px;color:{color};padding-left:10px;border-left:4px solid {color};\">{name}</h2>
  {''.join(items_html)}
</td></tr>"""
        )

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>AI 领袖动态日报 | {date_str}</title>
</head>
<body style="margin:0;padding:0;background-color:#f4f6f8;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','PingFang SC','Microsoft YaHei',sans-serif;color:#333;">
  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
    <tr><td align="center" style="padding:24px 16px;">
      <table role="presentation" width="640" cellspacing="0" cellpadding="0" border="0" style="max-width:640px;width:100%;background-color:#ffffff;border-radius:8px;overflow:hidden;">
        <tr><td style="padding:28px 24px 20px;border-bottom:2px solid #1a73e8;text-align:center;">
          <h1 style="margin:0 0 8px;font-size:24px;color:#1a237e;">&#127760; AI 领袖动态日报</h1>
          <p style="margin:0;font-size:14px;color:#666;">{date_str} · Prajna 企智</p>
        </td></tr>
        <tr><td style="padding:20px 24px 8px;">
          <p style="margin:0;font-size:14px;color:#555;line-height:1.7;">精选 <strong>李开复 / 零一万物、Sam Altman / OpenAI、Anthropic、Google DeepMind、Meta</strong> 的最新技术与战略动态，帮您快速掌握全球 AI 风向。</p>
        </td></tr>
        {''.join(leader_blocks)}
        <tr><td style="padding:16px 24px 28px;border-top:1px solid #e0e0e0;">
          <p style="margin:0;font-size:12px;color:#888;line-height:1.6;"><strong>免责声明：</strong>本日报基于公开网络信息整理，仅供信息参考，不构成投资建议或商业决策依据。新闻摘要与来源链接可能随时间变化，请自行核实原始报道。</p>
        </td></tr>
      </table>
    </td></tr>
  </table>
</body>
</html>"""


def render_html(data: dict, date_str: str) -> str:
    context = {"date": date_str, "leaders": data.get("leaders", [])}
    try:
        return _render_jinja2(HTML_TEMPLATE, context)
    except Exception as exc:
        print(f"[WARN] Jinja2 渲染失败（{exc}），降级使用纯 Python 渲染。", file=sys.stderr)
        return _render_fallback(context)


def build_html(data: dict, output_path: Path, date_str: str) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    html = render_html(data, date_str)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    return output_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="生成 AI 领袖动态日报 HTML 页面"
    )
    parser.add_argument(
        "--date", "-d",
        help="日报日期，格式 YYYY-MM-DD，默认当天"
    )
    parser.add_argument(
        "--output", "-o",
        help="输出文件完整路径（覆盖默认输出目录）"
    )
    parser.add_argument(
        "--data",
        help=f"自定义新闻数据 JSON 路径（默认：{DEFAULT_DATA}）"
    )
    parser.add_argument(
        "--list-leaders", action="store_true",
        help="列出默认覆盖的 AI 领袖"
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)

    data_path = Path(args.data) if args.data else DEFAULT_DATA
    if not data_path.exists():
        print(f"错误：数据文件不存在 {data_path}", file=sys.stderr)
        return 1

    data = load_json(data_path)

    if args.list_leaders:
        print("默认覆盖的 AI 领袖：")
        for leader in data.get("leaders", []):
            print(f"  - {leader.get('name', leader.get('id'))}")
        return 0

    date_str = args.date or _today()

    if args.output:
        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
    else:
        SAMPLES_DIR.mkdir(parents=True, exist_ok=True)
        output_path = SAMPLES_DIR / f"prajna_ai_leader_daily_{date_str}.html"

    build_html(data, output_path, date_str)
    print(f"已生成 AI 领袖日报：{output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
