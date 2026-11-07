"""
workflows.py
============
Workflow Recorder and Macro Automation Engine.
Records sequences of validated JSON operations, persists them as reproducible workflows,
and replays them onto new datasets with graceful column-mapping and missing-data handling.
"""

from __future__ import annotations
import json
import os
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd


class WorkflowRecorder:
    """Manages active workflow recording and saved workflow profiles."""

    def __init__(self, storage_dir: str = "workflows"):
        self.storage_dir = storage_dir
        os.makedirs(self.storage_dir, exist_ok=True)

    def save_workflow(
        self,
        name: str,
        description: str,
        steps: List[Dict[str, Any]]
    ) -> str:
        """Saves an ordered list of action steps as a reusable JSON macro."""
        clean_name = "".join(c for c in name if c.isalnum() or c in (" ", "_", "-")).strip()
        filename = f"{clean_name.replace(' ', '_')}.json"
        path = os.path.join(self.storage_dir, filename)

        workflow_data = {
            "name": name,
            "description": description,
            "version": "1.0",
            "created_at": pd.Timestamp.now().isoformat(),
            "step_count": len(steps),
            "steps": steps
        }

        with open(path, "w", encoding="utf-8") as f:
            json.dump(workflow_data, f, indent=2, ensure_ascii=False)

        return path

    def list_workflows(self) -> List[Dict[str, Any]]:
        """Lists all saved workflow profiles."""
        workflows = []
        if not os.path.exists(self.storage_dir):
            return workflows

        for fname in os.listdir(self.storage_dir):
            if fname.endswith(".json"):
                fpath = os.path.join(self.storage_dir, fname)
                try:
                    with open(fpath, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        workflows.append({
                            "filename": fname,
                            "path": fpath,
                            "name": data.get("name", fname),
                            "description": data.get("description", ""),
                            "step_count": data.get("step_count", len(data.get("steps", []))),
                            "steps": data.get("steps", [])
                        })
                except Exception:
                    pass
        return workflows

    def replay_workflow(
        self,
        workflow_steps: List[Dict[str, Any]],
        tables_store: Dict[str, pd.DataFrame],
        dispatcher_fn: Any
    ) -> Tuple[Dict[str, pd.DataFrame], List[Dict[str, Any]]]:
        """
        Executes workflow steps sequentially on tables_store.
        Gracefully skips or warns if columns or tables are missing.
        """
        execution_log: List[Dict[str, Any]] = []
        current_store = {k: v.copy() for k, v in tables_store.items()}

        for idx, step in enumerate(workflow_steps, start=1):
            action_name = step.get("action", "Unknown")
            try:
                res_df, msg, meta = dispatcher_fn(step, current_store)
                out_name = meta.get("output_name", f"Step_{idx}_Result")
                current_store[out_name] = res_df
                execution_log.append({
                    "step_number": idx,
                    "action": action_name,
                    "status": "success",
                    "message": msg,
                    "output_table": out_name,
                    "rows": len(res_df)
                })
            except Exception as e:
                execution_log.append({
                    "step_number": idx,
                    "action": action_name,
                    "status": "warning",
                    "message": f"Step skipped due to mismatch: {str(e)}"
                })

        return current_store, execution_log
