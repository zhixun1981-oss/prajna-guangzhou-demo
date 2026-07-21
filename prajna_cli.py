#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
prajna CLI
==========
Command-line interface for the prajna enterprise agent platform.

Usage examples:
    prajna run "帮我生成一份广州 P2 电商运营助理的薪资模板"
    prajna skill salary --city 广州 --level P2 --position 电商运营助理
    prajna workflow 招聘筛选联动
    prajna memory recall "广州 电商运营助理 薪酬"
    prajna memory remember "广州 P2 电商运营助理薪资带宽 6-9K" --type business --tags 薪酬,广州
"""

import argparse
import json
import subprocess
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT / "skills" / "prajna-code-agent" / "scripts"))

from code_agent import CodeAgent
from orchestrator import Orchestrator, WORKFLOWS
from memory_core import MemoryCore

SKILL_CLI = {
    "salary": {
        "script": PROJECT_ROOT / "skills" / "prajna-salary-template" / "scripts" / "generate_salary_template.py",
        "args": ["--industry", "--position", "--city", "--level"],
    },
    "sales": {
        "script": PROJECT_ROOT / "skills" / "prajna-sales-weekly-report" / "scripts" / "generate_sales_weekly_report.py",
        "args": ["--preset", "--team", "--week", "--sales-target", "--author", "--date"],
    },
    "recruitment": {
        "script": PROJECT_ROOT / "skills" / "hr" / "prajna-recruitment-assistant" / "scripts" / "generate_recruitment_kit.py",
        "args": ["--position", "--department", "--city", "--level", "--salary-min", "--salary-max", "--reports-to", "--headcount", "--urgency"],
    },
    "compensation": {
        "script": PROJECT_ROOT / "skills" / "hr" / "prajna-compensation-system" / "scripts" / "generate_compensation_system.py",
        "args": ["--company", "--industry", "--city", "--scale", "--budget", "--growth"],
    },
    "performance": {
        "script": PROJECT_ROOT / "skills" / "hr" / "prajna-performance-system" / "scripts" / "generate_performance_system.py",
        "args": ["--department", "--position", "--cycle", "--method", "--levels", "--company"],
    },
    "bidding": {
        "script": PROJECT_ROOT / "skills" / "sales" / "prajna-bidding-assistant" / "scripts" / "generate_bidding_kit.py",
        "args": ["--project", "--bidder", "--tenderer", "--amount", "--duration", "--industry"],
    },
}


def cmd_run(args):
    task = args.task
    agent = CodeAgent(use_llm=args.use_llm)
    result = agent.run(task, input_files=[Path(p) for p in (args.input or []) if Path(p).exists()], tags=args.tag)
    print(json.dumps(result, ensure_ascii=False, indent=2))


def cmd_skill(args):
    key = args.skill_key
    cfg = SKILL_CLI.get(key)
    if not cfg:
        print(f"Unknown skill: {key}. Supported: {', '.join(SKILL_CLI.keys())}")
        sys.exit(1)

    cmd = [sys.executable, str(cfg["script"])]
    for arg_name in cfg["args"]:
        param_name = arg_name.replace("--", "").replace("-", "_")
        value = getattr(args, param_name, None)
        if value is not None:
            cmd.extend([arg_name, str(value)])
    cmd.extend(["--output", args.output])

    proc = subprocess.run(cmd, capture_output=True, text=True)
    print(proc.stdout)
    if proc.returncode != 0:
        print(proc.stderr, file=sys.stderr)
        sys.exit(proc.returncode)


def cmd_workflow(args):
    name = args.workflow_name
    if name not in WORKFLOWS:
        print(f"Unknown workflow: {name}. Supported: {', '.join(WORKFLOWS.keys())}")
        sys.exit(1)
    orch = Orchestrator()
    result = orch.run_workflow(name, WORKFLOWS[name]["steps"])
    print(f"Workflow ID: {result.workflow_id}")
    print(f"Output dir: {result.output_dir}")
    for r in result.results:
        status = "✅" if r.ok else "❌"
        print(f"\n{status} {r.step_type.upper()}: {r.name}")
        for f in r.output_files:
            print(f"   → {f}")


def cmd_memory_recall(args):
    core = MemoryCore()
    results = core.recall(args.query, memory_type=args.type, tags=args.tag, limit=args.limit)
    for r in results:
        print(f"[{r.memory_type}] {r.timestamp[:19]} {r.source}")
        print(f"  {r.content}")
        print(f"  tags: {', '.join(r.tags)}\n")


def cmd_memory_remember(args):
    core = MemoryCore()
    entry = core.remember(
        content=args.content,
        memory_type=args.type,
        tags=args.tag,
        source="cli",
        source_task="prajna_cli",
    )
    print(f"Remembered: {entry.id}")


def cmd_memory_list(args):
    core = MemoryCore()
    results = core.list_all(memory_type=args.type, limit=args.limit)
    for r in results:
        print(f"[{r.memory_type}] {r.timestamp[:19]} {r.id}")
        print(f"  {r.content[:120]}")
        print(f"  tags: {', '.join(r.tags)}\n")


def main():
    parser = argparse.ArgumentParser(prog="prajna", description="prajna enterprise agent platform CLI")
    subparsers = parser.add_subparsers(dest="command", required=True)

    # run
    run_parser = subparsers.add_parser("run", help="Run a natural-language task via code agent")
    run_parser.add_argument("task", help="Natural language task description")
    run_parser.add_argument("--input", action="append", help="Input file path")
    run_parser.add_argument("--tag", action="append", help="Memory tag")
    run_parser.add_argument("--use-llm", action="store_true", help="Use LLM to generate code")
    run_parser.set_defaults(func=cmd_run)

    # skill
    skill_parser = subparsers.add_parser("skill", help="Run a single skill")
    skill_parser.add_argument("skill_key", choices=list(SKILL_CLI.keys()), help="Skill key")
    skill_parser.add_argument("--output", default="prajna_output", help="Output path prefix")
    skill_parser.add_argument("--industry", default="互联网")
    skill_parser.add_argument("--position", default="电商运营助理")
    skill_parser.add_argument("--city", default="广州")
    skill_parser.add_argument("--level", default="P2")
    skill_parser.add_argument("--team", default="华南销售一部")
    skill_parser.add_argument("--week", default="")
    skill_parser.add_argument("--sales-target", default="1200000")
    skill_parser.add_argument("--author", default="prajna")
    skill_parser.add_argument("--date", default="")
    skill_parser.add_argument("--department", default="电商运营部")
    skill_parser.add_argument("--salary-min", default="6000")
    skill_parser.add_argument("--salary-max", default="9000")
    skill_parser.add_argument("--reports-to", default="部门经理")
    skill_parser.add_argument("--headcount", default="1")
    skill_parser.add_argument("--urgency", default="中")
    skill_parser.add_argument("--company", default="智云科技")
    skill_parser.add_argument("--scale", default="200")
    skill_parser.add_argument("--budget", default="30000000")
    skill_parser.add_argument("--growth", default="5.0")
    skill_parser.add_argument("--cycle", default="月度")
    skill_parser.add_argument("--method", default="KPI")
    skill_parser.add_argument("--levels", default="A/B/C/D")
    skill_parser.add_argument("--project", default="智慧园区智能化建设项目")
    skill_parser.add_argument("--bidder", default="智讯科技有限公司")
    skill_parser.add_argument("--tenderer", default="广州高新区管委会")
    skill_parser.add_argument("--amount", default="5800000")
    skill_parser.add_argument("--duration", default="180天")
    skill_parser.set_defaults(func=cmd_skill)

    # workflow
    wf_parser = subparsers.add_parser("workflow", help="Run an intelligent multi-agent workflow")
    wf_parser.add_argument("workflow_name", choices=list(WORKFLOWS.keys()), help="Workflow name")
    wf_parser.set_defaults(func=cmd_workflow)

    # memory
    mem_parser = subparsers.add_parser("memory", help="Memory core operations")
    mem_sub = mem_parser.add_subparsers(dest="memory_command", required=True)

    recall_parser = mem_sub.add_parser("recall", help="Recall memories")
    recall_parser.add_argument("query", help="Recall query")
    recall_parser.add_argument("--type", choices=["project", "business", "feedback", "user"], help="Memory type filter")
    recall_parser.add_argument("--tag", action="append", help="Tag filter")
    recall_parser.add_argument("--limit", type=int, default=5)
    recall_parser.set_defaults(func=cmd_memory_recall)

    remember_parser = mem_sub.add_parser("remember", help="Remember a new entry")
    remember_parser.add_argument("content", help="Memory content")
    remember_parser.add_argument("--type", default="business", choices=["project", "business", "feedback", "user"])
    remember_parser.add_argument("--tag", action="append", help="Tags")
    remember_parser.set_defaults(func=cmd_memory_remember)

    list_parser = mem_sub.add_parser("list", help="List recent memories")
    list_parser.add_argument("--type", choices=["project", "business", "feedback", "user"], help="Memory type filter")
    list_parser.add_argument("--limit", type=int, default=20)
    list_parser.set_defaults(func=cmd_memory_list)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
