#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
prajna Memory Core
==================
A lightweight, file-based memory system for prajna Agents.

Memory types:
    - project:  per-task/project context (e.g., a specific hiring pipeline, bid)
    - business: reusable business rules, formats, KPI definitions
    - feedback: user corrections and preferences
    - user:     user-level preferences and history

Core operations:
    - remember: store a memory entry
    - recall:   retrieve relevant memories by keyword/tag/time similarity
    - reflect:  summarize recent memories (placeholder for LLM-driven reflection)
    - forget:   delete old or filtered memories

Storage layout:
    ~/.prajna/memory_core/
    ├── project/
    ├── business/
    ├── feedback/
    └── user/
        └── YYYYMMDD/
            └── {uuid}.json
"""

import json
import re
import shutil
import uuid
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional


MEMORY_DIR = Path.home() / ".prajna" / "memory_core"


@dataclass
class MemoryEntry:
    id: str
    timestamp: str
    memory_type: str  # project | business | feedback | user
    content: str
    tags: List[str] = field(default_factory=list)
    source: str = ""          # e.g., "code_agent", "orchestrator", "user"
    source_task: str = ""     # original natural language task
    metadata: Dict[str, Any] = field(default_factory=dict)


class MemoryCore:
    def __init__(self, memory_dir: Path = MEMORY_DIR):
        self.memory_dir = memory_dir
        for mtype in ("project", "business", "feedback", "user"):
            (self.memory_dir / mtype).mkdir(parents=True, exist_ok=True)

    def remember(
        self,
        content: str,
        memory_type: str = "project",
        tags: Optional[List[str]] = None,
        source: str = "",
        source_task: str = "",
        metadata: Optional[Dict[str, Any]] = None,
    ) -> MemoryEntry:
        if memory_type not in ("project", "business", "feedback", "user"):
            raise ValueError(f"Unsupported memory type: {memory_type}")

        entry = MemoryEntry(
            id=str(uuid.uuid4())[:8],
            timestamp=datetime.now().isoformat(),
            memory_type=memory_type,
            content=content,
            tags=[t.strip().lower() for t in (tags or []) if t.strip()],
            source=source,
            source_task=source_task,
            metadata=metadata or {},
        )

        date_dir = self.memory_dir / memory_type / datetime.now().strftime("%Y%m%d")
        date_dir.mkdir(parents=True, exist_ok=True)
        path = date_dir / f"{entry.id}.json"
        path.write_text(json.dumps(asdict(entry), ensure_ascii=False, indent=2), encoding="utf-8")
        return entry

    def recall(
        self,
        query: str,
        memory_type: Optional[str] = None,
        tags: Optional[List[str]] = None,
        limit: int = 5,
        days: Optional[int] = None,
    ) -> List[MemoryEntry]:
        """Rule-based recall: score entries by keyword and tag overlap."""
        query_tokens = self._tokenize(query)
        tag_set = set(t.strip().lower() for t in (tags or []) if t.strip())

        cutoff = None
        if days is not None:
            cutoff = datetime.now() - timedelta(days=days)

        candidates = []
        dirs = [self.memory_dir / memory_type] if memory_type else [self.memory_dir / t for t in ("project", "business", "feedback", "user")]
        for d in dirs:
            if not d.exists():
                continue
            for path in d.rglob("*.json"):
                try:
                    data = json.loads(path.read_text(encoding="utf-8"))
                    entry = MemoryEntry(**data)
                    if cutoff:
                        entry_time = datetime.fromisoformat(entry.timestamp)
                        if entry_time < cutoff:
                            continue
                    candidates.append(entry)
                except Exception:
                    continue

        scored = []
        for entry in candidates:
            score = 0.0
            entry_tokens = self._tokenize(entry.content + " " + " ".join(entry.tags) + " " + entry.source_task)
            matched = sum(1 for t in query_tokens if t in entry_tokens)
            score += matched * 1.0

            if tag_set:
                entry_tag_set = set(entry.tags)
                overlap = len(tag_set & entry_tag_set)
                score += overlap * 2.0

            # Recency boost
            try:
                age_days = (datetime.now() - datetime.fromisoformat(entry.timestamp)).days
                score += max(0, 1.0 - age_days / 30.0)
            except Exception:
                pass

            if score > 0:
                scored.append((score, entry))

        scored.sort(key=lambda x: x[0], reverse=True)
        return [entry for _, entry in scored[:limit]]

    def reflect(self, memory_type: Optional[str] = None, days: int = 7, limit: int = 10) -> str:
        """Simple reflection: count top tags and summarize recent themes."""
        entries = self.recall("", memory_type=memory_type, days=days, limit=limit)
        if not entries:
            return "近期没有可反思的记忆。"

        tag_counts: Dict[str, int] = {}
        sources: Dict[str, int] = {}
        for e in entries:
            for t in e.tags:
                tag_counts[t] = tag_counts.get(t, 0) + 1
            sources[e.source] = sources.get(e.source, 0) + 1

        lines = [f"近 {days} 天共沉淀 {len(entries)} 条记忆。"]
        if tag_counts:
            top_tags = sorted(tag_counts.items(), key=lambda x: x[1], reverse=True)[:5]
            lines.append("高频标签：" + ", ".join(f"{t}({c})" for t, c in top_tags))
        if sources:
            lines.append("主要来源：" + ", ".join(f"{s}({c})" for s, c in sources.items()))
        return "\n".join(lines)

    def forget(self, memory_type: Optional[str] = None, days: Optional[int] = None, tags: Optional[List[str]] = None) -> int:
        """Delete matching memories. Use with caution."""
        dirs = [self.memory_dir / memory_type] if memory_type else [self.memory_dir / t for t in ("project", "business", "feedback", "user")]
        tag_set = set(t.strip().lower() for t in (tags or []) if t.strip())
        cutoff = datetime.now() - timedelta(days=days) if days else None
        removed = 0

        for d in dirs:
            if not d.exists():
                continue
            for path in list(d.rglob("*.json")):
                try:
                    data = json.loads(path.read_text(encoding="utf-8"))
                    entry = MemoryEntry(**data)
                    match = True
                    if cutoff:
                        entry_time = datetime.fromisoformat(entry.timestamp)
                        if entry_time > cutoff:
                            match = False
                    if tag_set and not tag_set.issubset(set(entry.tags)):
                        match = False
                    if match:
                        path.unlink()
                        removed += 1
                except Exception:
                    continue
        return removed

    def list_all(self, memory_type: Optional[str] = None, limit: int = 50) -> List[MemoryEntry]:
        entries = []
        dirs = [self.memory_dir / memory_type] if memory_type else [self.memory_dir / t for t in ("project", "business", "feedback", "user")]
        for d in dirs:
            if not d.exists():
                continue
            for path in d.rglob("*.json"):
                try:
                    data = json.loads(path.read_text(encoding="utf-8"))
                    entries.append(MemoryEntry(**data))
                except Exception:
                    continue
        entries.sort(key=lambda e: e.timestamp, reverse=True)
        return entries[:limit]

    @staticmethod
    def _tokenize(text: str) -> List[str]:
        text = text.lower()
        # Keep Chinese characters and alphanumeric tokens
        tokens = re.findall(r"[\u4e00-\u9fa5]{2,}|[a-z0-9]+", text)
        return tokens


if __name__ == "__main__":
    core = MemoryCore()
    core.remember(
        content="广州 P2 电商运营助理薪酬口径：基本工资 6000-9000，社保公积金按广州基数缴纳，绩效占比 20%。",
        memory_type="business",
        tags=["薪酬", "广州", "电商运营助理", "P2"],
        source="user",
        source_task="设定广州电商运营助理薪酬口径",
    )
    print(core.reflect())
    print("Recall:", core.recall("广州 电商运营助理 工资", limit=3))
