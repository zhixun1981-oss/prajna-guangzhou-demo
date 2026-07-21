#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
prajna Agent Orchestrator
=========================
Coordinate multiple prajna Skills and the code agent into end-to-end workflows.

A workflow is a sequence of steps. Each step can be:
    - skill: run one of the document-generation Skills
    - code:  run the code agent on business data

The orchestrator handles parameter passing, file staging, and result collection.
"""

import argparse
import json
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from code_agent import CodeAgent
from memory_core import MemoryCore


# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
AGENT_DIR = Path(__file__).resolve().parent.parent
PROJECT_ROOT = AGENT_DIR.parent.parent
SKILLS_DIR = PROJECT_ROOT / "skills"
CODE_SAMPLES = AGENT_DIR / "samples"
HISTORY_DIR = Path.home() / ".prajna" / "demo_history"

# Map skill keys to their CLI scripts and default parameters
SKILL_CLI = {
    "recruitment": {
        "script": SKILLS_DIR / "hr" / "prajna-recruitment-assistant" / "scripts" / "generate_recruitment_kit.py",
        "params": {
            "position": "电商运营助理",
            "department": "电商运营部",
            "city": "广州",
            "level": "P2",
            "salary_min": 6000,
            "salary_max": 9000,
            "reports_to": "部门经理",
            "headcount": 1,
            "urgency": "中",
        },
    },
    "compensation": {
        "script": SKILLS_DIR / "hr" / "prajna-compensation-system" / "scripts" / "generate_compensation_system.py",
        "params": {
            "company": "智云科技",
            "industry": "互联网",
            "city": "广州",
            "scale": 200,
            "budget": 30000000,
            "growth": 5.0,
        },
    },
    "sales": {
        "script": SKILLS_DIR / "prajna-sales-weekly-report" / "scripts" / "generate_sales_weekly_report.py",
        "params": {
            "preset": "电商销售团队",
            "team": "华南销售一部",
            "week": f"{datetime.now().year}年第{datetime.now().isocalendar().week}周",
            "sales_target": 1200000,
            "author": "销售主管",
            "date": datetime.now().strftime("%Y-%m-%d"),
        },
    },
}


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------
@dataclass
class StepResult:
    step_type: str
    name: str
    ok: bool
    output_files: List[Path] = field(default_factory=list)
    stdout: str = ""
    stderr: str = ""
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class WorkflowResult:
    workflow_id: str
    results: List[StepResult]
    output_dir: Path
    recalled_memories: List[Any] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Workflow runner
# ---------------------------------------------------------------------------
class Orchestrator:
    def __init__(self, output_dir: Optional[Path] = None):
        self.output_dir = output_dir or HISTORY_DIR
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.memory_core = MemoryCore()

    def run_workflow(self, name: str, steps: List[Dict[str, Any]]) -> WorkflowResult:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        wf_id = f"{name}_{ts}"
        wf_dir = self.output_dir / wf_id
        wf_dir.mkdir(parents=True, exist_ok=True)

        # Recall relevant memories before executing the workflow
        query = f"{name} " + " ".join(s.get("name", "") or s.get("key", "") for s in steps)
        recalled = self.memory_core.recall(query, limit=3)

        results: List[StepResult] = []
        for idx, step in enumerate(steps, start=1):
            step_type = step.get("type")
            if step_type == "skill":
                r = self._run_skill_step(step, wf_dir, idx)
            elif step_type == "code":
                r = self._run_code_step(step, wf_dir, idx)
            else:
                r = StepResult(step_type=step_type or "unknown", name="unknown", ok=False, stderr=f"Unknown step type: {step_type}")
            results.append(r)

        # Persist workflow memory
        ok_count = sum(1 for r in results if r.ok)
        all_files = [f.name for r in results for f in r.output_files]
        self.memory_core.remember(
            content=f"执行工作流「{name}」，共 {len(steps)} 步，成功 {ok_count} 步，生成文件：{all_files}",
            memory_type="project",
            tags=[name, "workflow", "orchestrator"],
            source="orchestrator",
            source_task=name,
            metadata={"workflow_id": wf_id, "ok_count": ok_count, "total_steps": len(steps)},
        )

        return WorkflowResult(workflow_id=wf_id, results=results, output_dir=wf_dir, recalled_memories=recalled)

    def _run_skill_step(self, step: Dict[str, Any], wf_dir: Path, idx: int) -> StepResult:
        key = step["key"]
        cfg = SKILL_CLI.get(key)
        if not cfg:
            return StepResult(step_type="skill", name=key, ok=False, stderr=f"Skill {key} not configured in orchestrator")

        params = dict(cfg["params"])
        params.update(step.get("params", {}))
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        if key == "recruitment":
            safe_pos = re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9_-]", "_", params["position"])[:10]
            out_prefix = wf_dir / f"{idx}_招聘套件_{safe_pos}_{params['city']}_{timestamp}"
            cmd = [
                sys.executable, str(cfg["script"]),
                "--position", str(params["position"]),
                "--department", str(params["department"]),
                "--city", str(params["city"]),
                "--level", str(params["level"]),
                "--salary-min", str(params["salary_min"]),
                "--salary-max", str(params["salary_max"]),
                "--reports-to", str(params["reports_to"]),
                "--headcount", str(params["headcount"]),
                "--urgency", str(params["urgency"]),
                "--output", str(out_prefix),
                "--format", "all",
            ]
        elif key == "compensation":
            out_path = wf_dir / f"{idx}_薪酬体系_{params['company']}_{params['industry']}_{timestamp}.xlsx"
            cmd = [
                sys.executable, str(cfg["script"]),
                "--company", str(params["company"]),
                "--industry", str(params["industry"]),
                "--city", str(params["city"]),
                "--scale", str(params["scale"]),
                "--budget", str(params["budget"]),
                "--growth", str(params["growth"]),
                "--output", str(out_path),
            ]
        elif key == "sales":
            safe_team = params["team"].replace(" ", "_")
            out_path = wf_dir / f"{idx}_销售周报_{safe_team}_{timestamp}.xlsx"
            cmd = [
                sys.executable, str(cfg["script"]),
                "--preset", str(params["preset"]),
                "--team", str(params["team"]),
                "--week", str(params["week"]),
                "--sales-target", str(params["sales_target"]),
                "--author", str(params["author"]),
                "--date", str(params["date"]),
                "--output", str(out_path),
            ]
        else:
            return StepResult(step_type="skill", name=key, ok=False, stderr="Unsupported skill key")

        # Capture newly created files in the workflow directory
        before = set(wf_dir.iterdir())
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        ok = proc.returncode == 0

        output_files: List[Path] = []
        if ok:
            after = set(wf_dir.iterdir())
            output_files = sorted(after - before, key=lambda p: p.stat().st_mtime, reverse=True)

        return StepResult(
            step_type="skill",
            name=cfg["script"].parent.parent.name,
            ok=ok,
            output_files=output_files,
            stdout=proc.stdout,
            stderr=proc.stderr,
        )

    def _run_code_step(self, step: Dict[str, Any], wf_dir: Path, idx: int) -> StepResult:
        task = step["task"]
        intent = step.get("intent")
        sample_file = step.get("sample")
        input_files: List[Path] = []

        if sample_file:
            src = CODE_SAMPLES / sample_file
            if src.exists():
                dest = wf_dir / src.name
                shutil.copy2(src, dest)
                input_files.append(dest)

        for user_path in step.get("inputs", []):
            p = Path(user_path)
            if p.exists():
                input_files.append(p)

        agent = CodeAgent(use_llm=False)
        result = agent.run(task, input_files=input_files, tags=["workflow", step.get("name", "code")])
        ok = result["result"]["success"]

        output_files = [Path(a) for a in result["result"]["artifacts"] if Path(a).name != "generated_script.py"]
        # Move artifacts into workflow dir for unified packaging
        moved_files: List[Path] = []
        for f in output_files:
            if f.exists() and f.parent.resolve() != wf_dir.resolve():
                dest = wf_dir / f.name
                shutil.copy2(f, dest)
                moved_files.append(dest)
            else:
                moved_files.append(f)

        return StepResult(
            step_type="code",
            name=result["template_name"],
            ok=ok,
            output_files=moved_files,
            stdout=result["result"]["stdout"],
            stderr=result["result"]["stderr"],
            metadata={"intent": result["intent"], "task_id": result["task_id"]},
        )


# ---------------------------------------------------------------------------
# Pre-defined intelligent workflows
# ---------------------------------------------------------------------------
WORKFLOWS: Dict[str, Dict[str, Any]] = {
    "招聘筛选联动": {
        "desc": "生成招聘套件，并自动用代码智能体筛选候选人简历",
        "steps": [
            {"type": "skill", "key": "recruitment", "params": {"position": "电商运营助理", "city": "广州", "level": "P2"}},
            {"type": "code", "name": "简历智能筛选", "task": "筛选 resumes.xlsx 中电商运营助理岗位，要求1年以上经验", "sample": "resumes.xlsx", "intent": "resume_screening"},
        ],
    },
    "薪酬核算联动": {
        "desc": "生成薪酬体系后，自动用示例员工数据核算工资、社保、公积金与个税",
        "steps": [
            {"type": "skill", "key": "compensation", "params": {"company": "智云科技", "industry": "互联网", "city": "广州"}},
            {"type": "code", "name": "工资自动核算", "task": "核算 payroll_data.xlsx 的工资，按广州标准算社保公积金和个税", "sample": "payroll_data.xlsx", "intent": "payroll_calculation"},
        ],
    },
    "销售分析联动": {
        "desc": "生成销售周报后，自动对销售数据做区域分析与 HTML 可视化报告",
        "steps": [
            {"type": "skill", "key": "sales", "params": {"team": "华南销售一部", "sales_target": 1200000}},
            {"type": "code", "name": "销售数据分析", "task": "读取 sales_data.xlsx，按区域统计销售额和完成率，生成 HTML 报告", "sample": "sales_data.xlsx", "intent": "sales_analysis"},
        ],
    },
}


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="prajna Agent Orchestrator")
    parser.add_argument("workflow", choices=list(WORKFLOWS.keys()), help="要运行的工作流名称")
    parser.add_argument("--output", type=Path, help="输出目录")
    parser.add_argument("--list", action="store_true", help="列出可用工作流")
    args = parser.parse_args()

    if args.list:
        print("可用工作流：")
        for name, cfg in WORKFLOWS.items():
            print(f"  {name}: {cfg['desc']}")
        return

    orch = Orchestrator(output_dir=args.output)
    cfg = WORKFLOWS[args.workflow]
    result = orch.run_workflow(args.workflow, cfg["steps"])

    print(f"工作流 ID: {result.workflow_id}")
    print(f"输出目录: {result.output_dir}")
    for r in result.results:
        status = "✅" if r.ok else "❌"
        print(f"\n{status} {r.step_type.upper()}: {r.name}")
        for f in r.output_files:
            print(f"   → {f}")
        if r.stderr:
            print(f"   stderr: {r.stderr[:300]}")


if __name__ == "__main__":
    main()
