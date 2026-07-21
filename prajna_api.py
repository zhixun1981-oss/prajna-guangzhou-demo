#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
prajna FastAPI Service
======================
REST API for the prajna enterprise agent platform.

Run:
    uvicorn prajna_api:app --host 0.0.0.0 --port 8000 --reload

Endpoints:
    GET  /health
    GET  /skills
    GET  /workflows
    POST /run              { "task": "...", "use_llm": false, "tags": [...] }
    POST /skill/{skill_key} { "city": "广州", "level": "P2", ... }
    POST /workflow/{workflow_name}
    GET  /memory?query=...&type=...&limit=5
    POST /memory           { "content": "...", "type": "business", "tags": [...] }
"""

import json
import subprocess
import sys
import traceback
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT / "skills" / "prajna-code-agent" / "scripts"))

from code_agent import CodeAgent
from orchestrator import Orchestrator, WORKFLOWS
from memory_core import MemoryCore

app = FastAPI(title="prajna API", version="0.2.0")
memory_core = MemoryCore()


class RunRequest(BaseModel):
    task: str
    use_llm: bool = False
    tags: Optional[List[str]] = None


class SkillRequest(BaseModel):
    output: str = "prajna_api_output"
    industry: str = "互联网"
    position: str = "电商运营助理"
    city: str = "广州"
    level: str = "P2"
    team: str = "华南销售一部"
    week: str = ""
    sales_target: str = "1200000"
    author: str = "prajna"
    date: str = ""
    department: str = "电商运营部"
    salary_min: str = "6000"
    salary_max: str = "9000"
    reports_to: str = "部门经理"
    headcount: str = "1"
    urgency: str = "中"
    company: str = "智云科技"
    scale: str = "200"
    budget: str = "30000000"
    growth: str = "5.0"
    cycle: str = "月度"
    method: str = "KPI"
    levels: str = "A/B/C/D"
    project: str = "智慧园区智能化建设项目"
    bidder: str = "智讯科技有限公司"
    tenderer: str = "广州高新区管委会"
    amount: str = "5800000"
    duration: str = "180天"


class MemoryWriteRequest(BaseModel):
    content: str
    type: str = "business"
    tags: Optional[List[str]] = None


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


@app.get("/health")
def health():
    return {"status": "ok", "version": "0.2.0"}


@app.get("/skills")
def list_skills():
    return {
        "skills": [
            {"key": k, "name": k, "script": str(v["script"])} for k, v in SKILL_CLI.items()
        ]
    }


@app.get("/workflows")
def list_workflows():
    return {
        "workflows": [
            {"name": k, "description": v["desc"], "steps": len(v["steps"])} for k, v in WORKFLOWS.items()
        ]
    }


@app.post("/run")
def run_task(req: RunRequest):
    try:
        agent = CodeAgent(use_llm=req.use_llm)
        result = agent.run(req.task, tags=req.tags)
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=traceback.format_exc())


@app.post("/skill/{skill_key}")
def run_skill(skill_key: str, req: SkillRequest):
    cfg = SKILL_CLI.get(skill_key)
    if not cfg:
        raise HTTPException(status_code=404, detail=f"Unknown skill: {skill_key}")

    cmd = [sys.executable, str(cfg["script"])]
    for arg_name in cfg["args"]:
        param_name = arg_name.replace("--", "").replace("-", "_")
        value = getattr(req, param_name, None)
        if value is not None:
            cmd.extend([arg_name, str(value)])
    cmd.extend(["--output", req.output])

    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    if proc.returncode != 0:
        raise HTTPException(status_code=500, detail=proc.stderr or proc.stdout)

    # Locate generated file
    output_path = Path(f"{req.output}.xlsx")
    if not output_path.exists():
        output_path = Path(f"{req.output}.docx")
    if output_path.exists():
        return {"stdout": proc.stdout, "output_file": str(output_path), "size": output_path.stat().st_size}
    return {"stdout": proc.stdout}


@app.post("/workflow/{workflow_name}")
def run_workflow(workflow_name: str):
    if workflow_name not in WORKFLOWS:
        raise HTTPException(status_code=404, detail=f"Unknown workflow: {workflow_name}")
    try:
        orch = Orchestrator()
        result = orch.run_workflow(workflow_name, WORKFLOWS[workflow_name]["steps"])
        return {
            "workflow_id": result.workflow_id,
            "output_dir": str(result.output_dir),
            "results": [
                {
                    "step_type": r.step_type,
                    "name": r.name,
                    "ok": r.ok,
                    "output_files": [str(f) for f in r.output_files],
                }
                for r in result.results
            ],
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=traceback.format_exc())


@app.get("/memory")
def recall_memory(query: str = "", type: Optional[str] = None, limit: int = 10):
    results = memory_core.recall(query, memory_type=type, limit=limit)
    return {"query": query, "count": len(results), "memories": [json.loads(json.dumps(m.__dict__, ensure_ascii=False)) for m in results]}


@app.post("/memory")
def write_memory(req: MemoryWriteRequest):
    entry = memory_core.remember(
        content=req.content,
        memory_type=req.type,
        tags=req.tags or [],
        source="api",
        source_task="prajna_api",
    )
    return {"id": entry.id, "timestamp": entry.timestamp, "type": entry.memory_type}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
