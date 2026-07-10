import os
import json
from datetime import datetime
from typing import Dict, Any

class ReportGenerator:
    def __init__(self, metrics: Dict[str, Any], output_dir: str):
        self.metrics = metrics
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)

    def generate_json_report(self) -> str:
        """Saves the metrics dictionary as a JSON file."""
        project_name = self.metrics["project_name"]
        json_filename = f"{project_name}_weekly_report.json"
        json_filepath = os.path.join(self.output_dir, json_filename)
        
        # Datetime serializer for JSON
        def datetime_serializer(obj):
            if isinstance(obj, datetime) or hasattr(obj, "strftime"):
                return obj.strftime("%Y-%m-%d")
            raise TypeError("Type not serializable")
            
        with open(json_filepath, "w", encoding="utf-8") as jf:
            json.dump(self.metrics, jf, indent=2, default=datetime_serializer)
        return json_filepath

    def generate_markdown_report(self) -> str:
        """Saves the analysis as a Markdown weekly report."""
        m = self.metrics
        project_name = m["project_name"]
        md_filename = f"{project_name}_weekly_report.md"
        md_filepath = os.path.join(self.output_dir, md_filename)
        
        # Build Markdown content
        md_content = f"""# Weekly Project Health Report: {project_name}

**Report Date**: {m['today_date']}  
**Project Manager**: {m['project_manager']}  
**Overall RAG Status**: **{m['rag_status'].upper()}**  
**% Complete**: {m['pct_complete']*100:.1f}%  
**Schedule Health**: {m['calculated_score']:.1f}% (Calculated Score)  
**Project Stage**: {m['project_stage']}  
**At Risk**: {m['on_hold_count'] + m['blockers_count'] > 0 or m['rag_status'] != 'Green'}  

---

## 1. Executive Summary & Reasoning
{m['reasoning']}

---

## 2. Quantitative Health Dashboard

| Metric | Value |
| :--- | :--- |
| **Start Date** | {m['start_date']} |
| **Target End Date** | {m['end_date']} |
| **Overall Progress** | {m['pct_complete']*100:.1f}% |
| **Timeline Progress** | {m['time_elapsed_pct']*100:.1f}% |
| **Schedule Slippage** | {m['schedule_slippage']*100:.1f}% |
| **Total Tasks** | {m['total_tasks']} |
| **Completed Tasks** | {m['status_counts'].get('Completed', 0)} |
| **In Progress Tasks** | {m['status_counts'].get('In Progress', 0)} |
| **Not Started Tasks** | {m['status_counts'].get('Not Started', 0)} |
| **On Hold Tasks** | {m['status_counts'].get('On Hold', 0)} |
| **Overdue Milestones** | {m['overdue_milestones_count']} |
| **Active Blockers** | {m['blockers_count']} |

---

## 3. Detailed Risks & Blockers

### Overdue Milestones ({m['overdue_milestones_count']})
"""
        if m["overdue_milestones"]:
            md_content += "\n| Milestone Name | End Date | Status |\n| :--- | :--- | :--- |\n"
            for ms in m["overdue_milestones"]:
                md_content += f"| {ms['task_name']} | {ms['end_date']} | {ms['status']} |\n"
        else:
            md_content += "*No major overdue milestones identified.*\n"
            
        md_content += f"\n### Active Blockers & Issues ({m['blockers_count']})\n"
        if m["blockers"]:
            for idx, blk in enumerate(m["blockers"]):
                md_content += f"{idx+1}. **{blk['task_name']}** (Status: {blk['status']})\n"
                if blk["comments"]:
                    md_content += f"   * *Comments*: {blk['comments'][0]['comment']} (by {blk['comments'][0]['user']} on {blk['comments'][0]['date']})\n"
        else:
            md_content += "*No active critical blockers identified.*\n"
            
        md_content += "\n---\n\n## 4. Data Quality & Assumptions\n"
        if m["data_quality_notes"]:
            for note in m["data_quality_notes"]:
                md_content += f"- {note}\n"
        else:
            md_content += "*No major data quality issues reported. Project schedule structure is clean.*\n"
            
        with open(md_filepath, "w", encoding="utf-8") as mf:
            mf.write(md_content)
            
        return md_filepath

    def generate_powerpoint_report(self) -> str:
        """Save a readable four-slide weekly presentation for this project."""
        from src.weekly_ppt_generator import WeeklyPPTGenerator

        project_name = self.metrics["project_name"]
        ppt_filename = f"{project_name}_weekly_report.pptx"
        ppt_filepath = os.path.join(self.output_dir, ppt_filename)
        return WeeklyPPTGenerator(self.metrics, ppt_filepath).generate()
