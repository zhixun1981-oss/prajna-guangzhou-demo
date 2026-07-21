#!/usr/bin/env python3
"""
prajna-ecommerce-finance-kb-catalog 生成器 v1.0.0
为电商企业生成财务知识库钉钉文档目录 Markdown 模板。
包含七大模块、每模块 3 个子文件夹、首页标题与更新日志。
"""

import argparse
import os
import re
import sys
from datetime import datetime
from pathlib import Path


# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
DEFAULT_SAMPLES_DIR = Path.home() / ".prajna" / "prajna-ecommerce-finance-kb-catalog" / "samples"


# ---------------------------------------------------------------------------
# Knowledge base structure
# ---------------------------------------------------------------------------
CATALOG_TITLE = "电商企业财务知识库"

MODULES = [
    {
        "name": "01_核算规范",
        "summary": "统一收入、成本、费用、凭证、结账与报表口径，确保财务数据真实、完整、可比。",
        "folders": [
            ("01_收入确认与开票规范", "平台订单收入确认时点、发票开具规则、退换货收入冲减、跨境收入特殊处理。"),
            ("02_费用报销与凭证管理", "费用报销 SOP、发票审核要点、费用科目归属、凭证模板与附件标准。"),
            ("03_期末结账与报表口径", "结账清单、关账节点、各平台对账差异处理、内部管理报表与法定报表口径。"),
        ],
    },
    {
        "name": "02_成本利润模型",
        "summary": "搭建电商全链路成本利润测算框架，支撑定价、促销、平台选择与SKU盈利决策。",
        "folders": [
            ("01_商品成本核算模型", "采购成本、头程/尾程物流、包装耗材、平台佣金、支付手续费归集与分摊。"),
            ("02_平台费用分摊模型", "广告推广费、平台佣金、仓储费、活动资源费按店铺/链接/SKU 多维分摊逻辑。"),
            ("03_利润测算与盈亏平衡", "单品毛利模型、店铺利润表、盈亏平衡销量测算、促销 ROI 与价格弹性分析。"),
        ],
    },
    {
        "name": "03_库存管理",
        "summary": "衔接进销存、仓储系统与财务账，保障库存账实相符、跌价准备充分、周转健康。",
        "folders": [
            ("01_进销存台账管理", "采购入库、销售出库、退货入库、调拨盘点、成本结转与台账模板。"),
            ("02_存货跌价与盘点", "库存呆滞预警、跌价准备计提、盘点差异处理、报损报废审批流程。"),
            ("03_WMS 与财务对账", "ERP/WMS 期末库存与总账核对、在途物资管理、平台仓与自营仓数据一致性校验。"),
        ],
    },
    {
        "name": "04_资金税务",
        "summary": "覆盖资金计划、现金流监控与电商常见税务合规事项，降低税务与资金风险。",
        "folders": [
            ("01_资金计划与现金流", "销售回款预测、采购付款计划、平台提现节奏、资金日报与头寸管理。"),
            ("02_增值税与发票管理", "一般纳税人/小规模纳税人申报、进项抵扣、电子发票管理、异常发票处理。"),
            ("03_企业所得税汇算", "季度预缴、年度汇算清缴、研发费用加计扣除、跨境电商税收优惠政策。"),
        ],
    },
    {
        "name": "05_经营分析",
        "summary": "从财务视角输出经营洞察，支持管理层做增长、效率与风险的动态决策。",
        "folders": [
            ("01_月度经营分析报表", "收入、毛利、费用、利润、现金流五维月度分析模板与解读示例。"),
            ("02_店铺维度利润分析", "按平台/店铺/品牌拆分收入、成本、费用与利润，定位亏损与增长引擎。"),
            ("03_预算执行与预测", "年度预算模板、滚动预测方法、预算偏差分析、预警与纠偏机制。"),
        ],
    },
    {
        "name": "06_内控审计",
        "summary": "建立授权审批、风险清单与审计底稿体系，保障资产安全与财务合规。",
        "folders": [
            ("01_财务授权与审批流程", "付款审批权限、合同审批、费用报销审批、钉钉审批流配置示例。"),
            ("02_内控风险清单与应对", "电商典型财务风险点（刷单、退款套利、资金挪用）识别与防控措施。"),
            ("03_内部审计工作底稿", "审计计划、抽样方法、问题清单、整改跟踪表与审计报告模板。"),
        ],
    },
    {
        "name": "07_工具模板",
        "summary": "沉淀可直接套用的 Excel、钉钉表单、思维导图与索引工具，提升知识库使用效率。",
        "folders": [
            ("01_Excel 核算工具包", "自动对账表、毛利测算表、库存台账、资金日报、预算执行表等。"),
            ("02_钉钉审批模板库", "报销单、借款单、付款申请、合同审批、库存报损审批等表单模板。"),
            ("03_知识库维护与索引", "文档命名规范、版本更新日志、标签体系、检索目录与常见问题 FAQ。"),
        ],
    },
]


def _safe_filename(name: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|\s]+", "_", str(name))
    return name.strip("_") or "catalog"


def _today_cn() -> str:
    return datetime.now().strftime("%Y年%m月%d日")


def _today_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d")


def build_markdown(company: str = "", author: str = "", date: str = "") -> str:
    company = company or "XX电商"
    author = author or "财务部"
    date = date or _today_cn()

    lines = []
    lines.append(f"# {company} · {CATALOG_TITLE}")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("> **版本**：v1.0.0  ")
    lines.append(f"> **适用对象**：电商企业财务、运营、管理层  ")
    lines.append(f"> **维护部门**：{author}  ")
    lines.append(f"> **最近更新**：{date}  ")
    lines.append("")
    lines.append("---")
    lines.append("")

    lines.append("## 📋 更新日志")
    lines.append("")
    lines.append("| 版本 | 更新日期 | 更新内容 | 维护人 |")
    lines.append("|------|----------|----------|--------|")
    lines.append(f"| v1.0.0 | {_today_iso()} | 初始发布：建立核算规范、成本利润模型、库存管理、资金税务、经营分析、内控审计、工具模板七大模块，每模块下设 3 个子文件夹。 | {author} |")
    lines.append("| v0.1.0 | - | 需求调研与目录框架设计。 | - |")
    lines.append("")
    lines.append("---")
    lines.append("")

    lines.append("## 🏗️ 知识库目录结构")
    lines.append("")
    lines.append("```")
    lines.append(f"{CATALOG_TITLE}/")
    for module in MODULES:
        lines.append(f"├── {module['name']}/")
        for folder, _desc in module["folders"]:
            lines.append(f"│   ├── {folder}/")
        lines.append("│   └── README.md")
    lines.append("└── 00_首页.md")
    lines.append("```")
    lines.append("")
    lines.append("---")
    lines.append("")

    for idx, module in enumerate(MODULES, 1):
        lines.append(f"## {idx}. {module['name']}")
        lines.append("")
        lines.append(module["summary"])
        lines.append("")
        for f_idx, (folder, desc) in enumerate(module["folders"], 1):
            lines.append(f"### {idx}.{f_idx} {folder}")
            lines.append("")
            lines.append(f"- **定位**：{desc}")
            lines.append("- **推荐文档**：")
            lines.append("  - 制度文件 / SOP")
            lines.append("  - 操作指引 / 填写示例")
            lines.append("  - 模板 / 表单 / 检查清单")
            lines.append("")
        lines.append("---")
        lines.append("")

    lines.append("## 📝 使用说明")
    lines.append("")
    lines.append("1. 本目录按 `数字_中文` 规范命名，方便在钉钉文档中按字典序自动排序。")
    lines.append("2. 每个模块下的子文件夹统一包含：制度文件、操作指引、模板表单三类资料。")
    lines.append("3. 新增文档时请同步更新本目录 `更新日志` 与对应模块 README.md。")
    lines.append("4. 涉及金额、税率、法规等政策类内容，需以最新政策及企业实际情况为准，建议人工复核。")
    lines.append("")
    lines.append("---")
    lines.append("")
    lines.append("## ⚠️ 免责声明")
    lines.append("")
    lines.append("> 【人工智能生成-需人工核验】本知识库目录模板由 prajna 企智自动生成，仅供参考。具体财务制度、税务处理、内控流程应根据企业所处行业、经营模式、所在地区法规及公司内部管理制度制定，并建议由财务负责人、法务或审计人员审核后使用。")
    lines.append("")

    return "\n".join(lines)


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="生成电商企业财务知识库钉钉文档目录 Markdown 模板"
    )
    parser.add_argument(
        "--company", "-c",
        help="企业名称，用于标题，默认 XX电商",
        default="XX电商",
    )
    parser.add_argument(
        "--author", "-a",
        help="维护部门/人，默认 财务部",
        default="财务部",
    )
    parser.add_argument(
        "--date", "-d",
        help="更新日期，格式 YYYY-MM-DD 或中文日期，默认今天",
        default=_today_cn(),
    )
    parser.add_argument(
        "--output", "-o",
        help="输出文件路径，覆盖默认路径",
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)

    if args.output:
        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
    else:
        DEFAULT_SAMPLES_DIR.mkdir(parents=True, exist_ok=True)
        safe_company = _safe_filename(args.company)
        filename = f"prajna_电商企业财务知识库目录_{safe_company}_{_today_iso()}.md"
        output_path = DEFAULT_SAMPLES_DIR / filename

    markdown = build_markdown(company=args.company, author=args.author, date=args.date)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(markdown)

    print(f"已生成财务知识库目录：{output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
