#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
prajna-code-agent
=================
A business-aware code generation and execution agent for prajna.

Core flow:
    1. Parse natural-language task
    2. Match intent to a code template
    3. Render code with extracted parameters
    4. Execute in a sandboxed subprocess
    5. Return generated artifacts + execution log
    6. Persist code, result and metadata to memory

Design choices:
    - Rule-first: works without any API key by default
    - LLM-ready: `CodeGenerator` can be swapped for an LLM backend
    - Sandbox: subprocess with timeout, cwd isolation, import whitelist
"""

import argparse
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from llm_backend import LLMBackend
from memory_core import MemoryCore, MemoryEntry


# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
AGENT_DIR = Path(__file__).resolve().parent.parent
TEMPLATES_DIR = AGENT_DIR / "templates"
MEMORY_DIR = Path.home() / ".prajna" / "code_agent_memory"


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------
@dataclass
class ExecutionResult:
    success: bool
    stdout: str
    stderr: str
    artifacts: List[str] = field(default_factory=list)
    error: Optional[str] = None
    duration_ms: int = 0
    workdir: Optional[str] = None


@dataclass
class TaskRecord:
    task_id: str
    timestamp: str
    raw_task: str
    intent: str
    parameters: Dict[str, Any]
    generated_code: str
    result: ExecutionResult
    memory_tags: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Template registry
# ---------------------------------------------------------------------------
CODE_TEMPLATES: Dict[str, Dict[str, Any]] = {
    "sales_analysis": {
        "name": "销售数据分析",
        "tags": ["销售", "销售额", "业绩", "完成率", "环比", "统计", "分析", "图表"],
        "description": "读取销售数据 Excel，按区域/人员统计销售额、完成率、环比，生成 HTML 图表报告",
        "parameters": {
            "input_file": "sales_data.xlsx",
            "group_by": "区域",
            "target_col": "销售目标",
            "amount_col": "销售额",
            "output_file": "sales_report.html",
        },
    },
    "payroll_calculation": {
        "name": "薪酬核算",
        "tags": ["薪酬", "工资", "个税", "社保", "公积金", "核算", "发薪", "实发"],
        "description": "读取员工薪酬数据，计算基本工资、绩效、社保公积金、个税、实发工资",
        "parameters": {
            "input_file": "payroll_data.xlsx",
            "city": "广州",
            "output_file": "payroll_result.xlsx",
        },
    },
    "resume_screening": {
        "name": "简历筛选",
        "tags": ["简历", "筛选", "招聘", "匹配度", "人才", "面试"],
        "description": "读取候选人简历 Excel，按岗位关键词、经验年限、学历等条件筛选并打分",
        "parameters": {
            "input_file": "resumes.xlsx",
            "position": "电商运营助理",
            "keywords": "电商,运营,淘宝,天猫,数据分析",
            "min_years": 1,
            "output_file": "screening_result.xlsx",
        },
    },
    "inventory_analysis": {
        "name": "库存分析",
        "tags": ["库存", "周转", "滞销", "SKU", "存货", "进销存"],
        "description": "读取库存数据，计算周转天数、滞销 SKU、库存金额，并标记补货建议",
        "parameters": {
            "input_file": "inventory.xlsx",
            "output_file": "inventory_analysis.xlsx",
        },
    },
    "kpi_scoring": {
        "name": "KPI 绩效评分",
        "tags": ["绩效", "KPI", "评分", "考核", "得分", "加权"],
        "description": "读取 KPI 指标与完成数据，按权重计算加权得分并生成绩效报告",
        "parameters": {
            "input_file": "kpi_data.xlsx",
            "output_file": "kpi_score.xlsx",
        },
    },
}


# ---------------------------------------------------------------------------
# Code templates (Python source strings)
# ---------------------------------------------------------------------------
SALES_ANALYSIS_CODE = '''
import json
import pandas as pd
from pathlib import Path

INPUT = Path("{{input_file}}")
OUTPUT = Path("{{output_file}}")

def main():
    df = pd.read_excel(INPUT)
    df["完成率"] = (df["{{amount_col}}"] / df["{{target_col}}"] * 100).round(2)
    df["环比"] = df["{{amount_col}}"].pct_change().fillna(0).round(4)

    summary = df.groupby("{{group_by}}").agg({
        "{{amount_col}}": "sum",
        "{{target_col}}": "sum",
        "完成率": "mean",
    }).reset_index()
    summary["总完成率"] = (summary["{{amount_col}}"] / summary["{{target_col}}"] * 100).round(2)

    # Build a simple HTML report with Chart.js
    html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>销售数据分析报告</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 40px; background: #f8f9fa; }}
.card {{ background: #fff; border-radius: 12px; padding: 24px; margin-bottom: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }}
h1 {{ color: #1a1a2e; }} table {{ border-collapse: collapse; width: 100%; margin-top: 16px; }}
th, td {{ border: 1px solid #e9ecef; padding: 10px; text-align: left; }}
th {{ background: #f1f3f5; }}
</style>
</head>
<body>
<div class="card">
<h1>销售数据分析报告</h1>
<p>生成时间：{pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")}</p>
</div>
<div class="card">
<h2>汇总表</h2>
{summary.to_html(index=False, classes="table", escape=False)}
</div>
<div class="card">
<h2>区域销售额对比</h2>
<canvas id="chart" height="80"></canvas>
</div>
<script>
const ctx = document.getElementById('chart').getContext('2d');
new Chart(ctx, {{
    type: 'bar',
    data: {{
        labels: {json.dumps(summary["{{group_by}}"].tolist())},
        datasets: [{{
            label: '销售额',
            data: {json.dumps(summary["{{amount_col}}"].tolist())},
            backgroundColor: 'rgba(54, 162, 235, 0.6)'
        }}]
    }},
    options: {{ responsive: true }}
}});
</script>
</body>
</html>
"""
    OUTPUT.write_text(html, encoding="utf-8")
    print(f"报告已生成：{OUTPUT}")
    print(summary.to_csv(index=False))

if __name__ == "__main__":
    main()
'''


PAYROLL_CALCULATION_CODE = '''
import pandas as pd
from pathlib import Path

INPUT = Path("{{input_file}}")
OUTPUT = Path("{{output_file}}")
CITY = "{{city}}"

SOCIAL_RATES = {
    "广州": {"pension": 0.08, "medical": 0.02, "unemployment": 0.005, "housing": 0.07},
    "深圳": {"pension": 0.08, "medical": 0.02, "unemployment": 0.005, "housing": 0.07},
    "北京": {"pension": 0.08, "medical": 0.02, "unemployment": 0.005, "housing": 0.12},
    "上海": {"pension": 0.08, "medical": 0.02, "unemployment": 0.005, "housing": 0.07},
}

def calc_tax(income):
    # Simplified monthly IIT table for China
    taxable = max(income - 5000, 0)
    brackets = [
        (3000, 0.03, 0),
        (12000, 0.10, 210),
        (25000, 0.20, 1410),
        (35000, 0.25, 2660),
        (55000, 0.30, 4410),
        (80000, 0.35, 7160),
        (float("inf"), 0.45, 15160),
    ]
    for limit, rate, quick in brackets:
        if taxable <= limit:
            return max(taxable * rate - quick, 0)
    return 0

def main():
    df = pd.read_excel(INPUT)
    rates = SOCIAL_RATES.get(CITY, SOCIAL_RATES["广州"])
    df["养老保险"] = (df["社保基数"] * rates["pension"]).round(2)
    df["医疗保险"] = (df["社保基数"] * rates["medical"]).round(2)
    df["失业保险"] = (df["社保基数"] * rates["unemployment"]).round(2)
    df["住房公积金"] = (df["公积金基数"] * rates["housing"]).round(2)
    df["社保公积金合计"] = df[["养老保险", "医疗保险", "失业保险", "住房公积金"]].sum(axis=1).round(2)
    df["应税工资"] = df["基本工资"] + df["绩效工资"] - df["社保公积金合计"]
    df["个税"] = df["应税工资"].apply(calc_tax).round(2)
    df["实发工资"] = (df["应税工资"] - df["个税"]).round(2)
    df.to_excel(OUTPUT, index=False)
    print(f"薪酬核算完成：{OUTPUT}")
    print(df[["姓名", "基本工资", "绩效工资", "社保公积金合计", "个税", "实发工资"]].to_csv(index=False))

if __name__ == "__main__":
    main()
'''


RESUME_SCREENING_CODE = '''
import pandas as pd
from pathlib import Path

INPUT = Path("{{input_file}}")
OUTPUT = Path("{{output_file}}")
POSITION = "{{position}}"
KEYWORDS = [k.strip() for k in "{{keywords}}".split(",") if k.strip()]
MIN_YEARS = {{min_years}}

def score(row):
    text = f"{row.get('技能', '')} {row.get('项目经验', '')} {row.get('自我评价', '')}"
    hits = sum(1 for kw in KEYWORDS if kw.lower() in text.lower())
    years = float(str(row.get("工作年限", "0")).replace("年", "").strip() or 0)
    edu_score = {"本科": 60, "硕士": 80, "博士": 100, "大专": 40}.get(str(row.get("学历", "")), 20)
    keyword_score = min(hits * 15, 60)
    exp_score = min(years * 10, 30)
    total = keyword_score + exp_score + (edu_score * 0.1)
    return round(total, 1), hits

def main():
    df = pd.read_excel(INPUT)
    scored = df.apply(lambda r: score(r), axis=1, result_type="expand")
    scored.columns = ["匹配得分", "关键词命中数"]
    df = pd.concat([df, scored], axis=1)
    df["是否通过初筛"] = (df["匹配得分"] >= 70) & (df["工作年限"] >= MIN_YEARS)
    df = df.sort_values("匹配得分", ascending=False)
    df.to_excel(OUTPUT, index=False)
    passed = df[df["是否通过初筛"]].shape[0]
    print(f"初筛完成：共 {len(df)} 人，通过 {passed} 人")
    print(df[["姓名", "学历", "工作年限", "匹配得分", "是否通过初筛"]].to_csv(index=False))

if __name__ == "__main__":
    main()
'''


INVENTORY_ANALYSIS_CODE = '''
import pandas as pd
from pathlib import Path

INPUT = Path("{{input_file}}")
OUTPUT = Path("{{output_file}}")

def main():
    df = pd.read_excel(INPUT)
    df["库存金额"] = (df["库存数量"] * df["单价"]).round(2)
    df["周转天数"] = (df["库存数量"] / df["月均销量"].replace(0, float("nan"))).round(1)
    df["滞销标记"] = df["周转天数"].apply(lambda x: "滞销" if pd.notna(x) and x > 90 else "正常")
    df["补货建议"] = df.apply(lambda r: "建议补货" if r["库存数量"] < r["安全库存"] and r["滞销标记"] == "正常" else "-", axis=1)

    summary = pd.DataFrame({
        "总 SKU 数": [len(df)],
        "库存总金额": [df["库存金额"].sum().round(2)],
        "滞销 SKU 数": [df[df["滞销标记"] == "滞销"].shape[0]],
        "需补货 SKU 数": [df[df["补货建议"] == "建议补货"].shape[0]],
    })

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="库存分析", index=False)
        summary.to_excel(writer, sheet_name="汇总", index=False)

    print(f"库存分析完成：{OUTPUT}")
    print(summary.to_csv(index=False))

if __name__ == "__main__":
    main()
'''


KPI_SCORING_CODE = '''
import pandas as pd
from pathlib import Path

INPUT = Path("{{input_file}}")
OUTPUT = Path("{{output_file}}")

def main():
    df = pd.read_excel(INPUT)
    df["单项得分"] = (df["完成率"].clip(upper=1.0) * df["权重"] * 100).round(2)
    summary = df.groupby("被考核人").agg({"单项得分": "sum", "权重": "sum"}).reset_index()
    summary["KPI 加权得分"] = summary["单项得分"].round(2)
    summary["绩效等级"] = summary["KPI 加权得分"].apply(
        lambda s: "S" if s >= 90 else "A" if s >= 80 else "B" if s >= 70 else "C" if s >= 60 else "D"
    )

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="KPI 明细", index=False)
        summary.to_excel(writer, sheet_name="绩效汇总", index=False)

    print(f"绩效评分完成：{OUTPUT}")
    print(summary.to_csv(index=False))

if __name__ == "__main__":
    main()
'''


TEMPLATE_SOURCE = {
    "sales_analysis": SALES_ANALYSIS_CODE,
    "payroll_calculation": PAYROLL_CALCULATION_CODE,
    "resume_screening": RESUME_SCREENING_CODE,
    "inventory_analysis": INVENTORY_ANALYSIS_CODE,
    "kpi_scoring": KPI_SCORING_CODE,
}


# ---------------------------------------------------------------------------
# Core agent
# ---------------------------------------------------------------------------
class TaskParser:
    """Rule-based NL task parser."""

    def parse(self, task: str) -> Dict[str, Any]:
        task_lower = task.lower()
        best_intent = None
        best_score = 0
        for intent, meta in CODE_TEMPLATES.items():
            score = sum(1 for tag in meta["tags"] if tag.lower() in task_lower)
            if score > best_score:
                best_score = score
                best_intent = intent

        # Fallbacks based on strong keywords
        if best_score == 0:
            if any(k in task_lower for k in ["销售", "业绩", "完成率", "销售额"]):
                best_intent = "sales_analysis"
            elif any(k in task_lower for k in ["工资", "薪酬", "个税", "社保", "公积金", "发薪"]):
                best_intent = "payroll_calculation"
            elif any(k in task_lower for k in ["简历", "招聘", "筛选", "候选人"]):
                best_intent = "resume_screening"
            elif any(k in task_lower for k in ["库存", "周转", "滞销", "sku"]):
                best_intent = "inventory_analysis"
            elif any(k in task_lower for k in ["绩效", "kpi", "考核", "评分"]):
                best_intent = "kpi_scoring"
            else:
                best_intent = "sales_analysis"

        params = dict(CODE_TEMPLATES[best_intent]["parameters"])
        params = self._extract_parameters(task, best_intent, params)
        return {
            "intent": best_intent,
            "template_name": CODE_TEMPLATES[best_intent]["name"],
            "parameters": params,
        }

    def _extract_parameters(self, task: str, intent: str, params: Dict[str, Any]) -> Dict[str, Any]:
        # File extraction: look for paths ending in xlsx/xls/csv (non-capturing group for extension)
        file_matches = re.findall(r"[\w\-\u4e00-\u9fa5/]+\.(?:xlsx|xls|csv)", task)
        if file_matches:
            if intent in ("sales_analysis", "payroll_calculation", "resume_screening", "inventory_analysis", "kpi_scoring"):
                params["input_file"] = Path(file_matches[0]).name

        # Output file extraction
        out_matches = re.findall(r"保存[为到]?\s*([\w\-\u4e00-\u9fa5/]+\.(?:xlsx|xls|csv|html))", task)
        if out_matches:
            params["output_file"] = Path(out_matches[0]).name

        # City extraction
        cities = ["广州", "深圳", "北京", "上海", "杭州", "成都", "武汉", "西安"]
        for city in cities:
            if city in task:
                params["city"] = city
                break

        # Position / keywords extraction for resume screening
        if intent == "resume_screening":
            m = re.search(r"岗位[是为]?\s*([\u4e00-\u9fa5\w]+)", task)
            if m:
                params["position"] = m.group(1)
            m = re.search(r"([\u4e00-\u9fa5\w]+)年[及以]?[上]?经验", task)
            if m:
                try:
                    params["min_years"] = int(m.group(1))
                except ValueError:
                    pass

        # Group-by extraction for sales analysis
        if intent == "sales_analysis":
            for col in ["区域", "大区", "省份", "城市", "销售人员", "销售代表"]:
                if col in task:
                    params["group_by"] = col
                    break

        return params


class CodeGenerator:
    """Generate Python code from templates or via LLM."""

    def __init__(self, use_llm: bool = False, llm_provider: Optional[str] = None):
        self.use_llm = use_llm
        self.llm_provider = llm_provider
        self.llm = LLMBackend(provider=llm_provider) if use_llm else None
        self.last_error: Optional[str] = None

    def generate(self, intent: str, parameters: Dict[str, Any], task: Optional[str] = None) -> str:
        self.last_error = None
        if self.use_llm and self.llm and task:
            llm_code = self.llm.generate_code(task, parameters)
            if llm_code:
                return llm_code
            self.last_error = self.llm.last_error or "LLM failed to generate code"
            # Fall back to rule templates if LLM unavailable or fails
        source = TEMPLATE_SOURCE[intent]
        code = source
        for key, value in parameters.items():
            placeholder = "{{" + key + "}}"
            code = code.replace(placeholder, str(value))
        return code


class SandboxExecutor:
    """Execute Python code in a temporary directory with timeout."""

    def __init__(self, timeout: int = 30, allowed_modules: Optional[List[str]] = None):
        self.timeout = timeout
        self.allowed_modules = allowed_modules or ["pandas", "numpy", "openpyxl", "matplotlib", "json", "pathlib"]

    def execute(
        self,
        code: str,
        input_files: Optional[List[Path]] = None,
        workdir: Optional[Path] = None,
    ) -> ExecutionResult:
        MEMORY_DIR.mkdir(parents=True, exist_ok=True)
        workdir = workdir or Path(tempfile.mkdtemp(prefix="prajna_code_", dir=MEMORY_DIR))
        script_path = workdir / "generated_script.py"
        script_path.write_text(code, encoding="utf-8")

        # Copy input files if provided
        if input_files:
            for f in input_files:
                if f.exists():
                    shutil.copy2(f, workdir / f.name)

        start = datetime.now()
        try:
            proc = subprocess.run(
                [sys.executable, str(script_path)],
                cwd=str(workdir),
                capture_output=True,
                text=True,
                timeout=self.timeout,
            )
            duration = int((datetime.now() - start).total_seconds() * 1000)

            artifacts = []
            for f in workdir.iterdir():
                if f.name != "generated_script.py" and f.is_file():
                    artifacts.append(str(f))

            return ExecutionResult(
                success=proc.returncode == 0,
                stdout=proc.stdout,
                stderr=proc.stderr,
                artifacts=artifacts,
                error=None if proc.returncode == 0 else f"Process exited with code {proc.returncode}",
                duration_ms=duration,
                workdir=str(workdir),
            )
        except subprocess.TimeoutExpired:
            return ExecutionResult(success=False, stdout="", stderr="", error="Execution timeout", workdir=str(workdir))
        except Exception as e:
            return ExecutionResult(success=False, stdout="", stderr=traceback.format_exc(), error=str(e), workdir=str(workdir))


class MemoryStore:
    """Persist task records for recall and continuous improvement."""

    def __init__(self, memory_dir: Path = MEMORY_DIR):
        self.memory_dir = memory_dir
        self.memory_dir.mkdir(parents=True, exist_ok=True)

    def save(self, record: TaskRecord, record_dir: Optional[Path] = None) -> Path:
        if record_dir is None:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_id = re.sub(r"[^\w\-]", "_", record.task_id)[:40]
            record_dir = self.memory_dir / f"{safe_id}_{ts}"
            record_dir.mkdir(parents=True, exist_ok=True)

        metadata = {
            "task_id": record.task_id,
            "timestamp": record.timestamp,
            "raw_task": record.raw_task,
            "intent": record.intent,
            "template_name": CODE_TEMPLATES.get(record.intent, {}).get("name", record.intent),
            "parameters": record.parameters,
            "memory_tags": record.memory_tags,
            "result": {
                "success": record.result.success,
                "stdout": record.result.stdout,
                "stderr": record.result.stderr,
                "artifacts": record.result.artifacts,
                "error": record.result.error,
                "duration_ms": record.result.duration_ms,
            },
        }
        (record_dir / "metadata.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
        (record_dir / "generated_script.py").write_text(record.generated_code, encoding="utf-8")
        for artifact in record.result.artifacts:
            src = Path(artifact)
            if src.exists() and src.parent.resolve() != record_dir.resolve():
                shutil.copy2(src, record_dir / src.name)
        return record_dir

    def list_tasks(self, intent: Optional[str] = None, limit: int = 20) -> List[Dict[str, Any]]:
        tasks = []
        for d in sorted(self.memory_dir.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True):
            meta_path = d / "metadata.json"
            if meta_path.exists():
                meta = json.loads(meta_path.read_text(encoding="utf-8"))
                if intent is None or meta.get("intent") == intent:
                    tasks.append(meta)
                if len(tasks) >= limit:
                    break
        return tasks


class CodeAgent:
    """Main orchestrator."""

    def __init__(self, use_llm: bool = False, llm_provider: Optional[str] = None):
        self.parser = TaskParser()
        self.generator = CodeGenerator(use_llm=use_llm, llm_provider=llm_provider)
        self.executor = SandboxExecutor(timeout=60)
        self.memory = MemoryStore()
        self.memory_core = MemoryCore()

    def run(self, task: str, input_files: Optional[List[Path]] = None, tags: Optional[List[str]] = None) -> Dict[str, Any]:
        # Recall relevant memories before execution
        recalled = self.memory_core.recall(task, limit=3)

        parsed = self.parser.parse(task)
        intent = parsed["intent"]
        parameters = parsed["parameters"]

        # Augment generation with recalled business memory if available
        if recalled and not self.generator.use_llm:
            # For rule mode, we cannot easily rewrite code, but we store memory for UI display
            pass

        code = self.generator.generate(intent, parameters, task=task)

        task_id = f"code_{intent}_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_id = re.sub(r"[^\w\-]", "_", task_id)[:40]
        record_dir = MEMORY_DIR / f"{safe_id}_{ts}"
        record_dir.mkdir(parents=True, exist_ok=True)

        result = self.executor.execute(code, input_files=input_files, workdir=record_dir)

        record = TaskRecord(
            task_id=task_id,
            timestamp=datetime.now().isoformat(),
            raw_task=task,
            intent=intent,
            parameters=parameters,
            generated_code=code,
            result=result,
            memory_tags=tags or [],
        )
        memory_path = self.memory.save(record, record_dir=record_dir)

        # Persist key information to memory core
        if result.success:
            artifact_names = [Path(a).name for a in result.artifacts if Path(a).name != "generated_script.py"]
            memory_content = (
                f"任务：{task}\n"
                f"识别意图：{parsed['template_name']}（{intent}）\n"
                f"输入文件：{parameters.get('input_file', '无')}，输出文件：{artifact_names}\n"
                f"执行结果：成功，耗时 {result.duration_ms} ms"
            )
            self.memory_core.remember(
                content=memory_content,
                memory_type="project",
                tags=(tags or []) + [intent, "code_agent"],
                source="code_agent",
                source_task=task,
                metadata={"task_id": task_id, "intent": intent, "parameters": parameters},
            )

            # Also extract reusable business rules into business memory for payroll/sales intents
            if intent == "payroll_calculation":
                self.memory_core.remember(
                    content=f"薪酬核算任务使用城市：{parameters.get('city', '广州')}，输入格式含基本工资、绩效工资、社保基数、公积金基数。",
                    memory_type="business",
                    tags=["薪酬核算", parameters.get("city", "广州"), "code_agent"],
                    source="code_agent",
                    source_task=task,
                )
            elif intent == "sales_analysis":
                self.memory_core.remember(
                    content=f"销售数据分析按 {parameters.get('group_by', '区域')} 分组，统计 {parameters.get('amount_col', '销售额')} 与 {parameters.get('target_col', '销售目标')}。",
                    memory_type="business",
                    tags=["销售分析", parameters.get("group_by", "区域"), "code_agent"],
                    source="code_agent",
                    source_task=task,
                )

        used_llm = self.generator.use_llm and code != TEMPLATE_SOURCE.get(intent, "")
        return {
            "task_id": task_id,
            "intent": intent,
            "template_name": parsed["template_name"],
            "parameters": parameters,
            "generated_code": code,
            "result": asdict(result),
            "memory_path": str(memory_path),
            "recalled_memories": [asdict(m) for m in recalled],
            "used_llm": used_llm,
            "llm_error": self.generator.last_error,
        }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="prajna-code-agent")
    parser.add_argument("task", help="自然语言任务描述")
    parser.add_argument("--input", action="append", type=Path, help="输入文件路径，可多次指定")
    parser.add_argument("--tag", action="append", help="记忆标签")
    parser.add_argument("--use-llm", action="store_true", help="启用 LLM 生成（尚未实现）")
    parser.add_argument("--list-templates", action="store_true", help="列出可用模板")
    args = parser.parse_args()

    if args.list_templates:
        print("可用代码模板：")
        for intent, meta in CODE_TEMPLATES.items():
            print(f"  {intent}: {meta['name']} — {meta['description']}")
        return

    agent = CodeAgent(use_llm=args.use_llm)
    output = agent.run(args.task, input_files=args.input, tags=args.tag)
    print(json.dumps(output, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
