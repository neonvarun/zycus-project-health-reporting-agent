import pandas as pd
from typing import Dict, Any, List, Tuple
from datetime import datetime
from src.excel_reader import ProjectData, TaskInfo, CommentInfo

class HealthRulesEvaluator:
    def __init__(self, project_data: ProjectData):
        self.project_data = project_data
        self.metrics = {}

    def evaluate(self) -> Dict[str, Any]:
        """Calculates project health metrics and determines RAG status."""
        p = self.project_data
        if hasattr(p, "precalculated_metrics") and p.precalculated_metrics is not None:
            return p.precalculated_metrics
        tasks = p.tasks
        
        # 1. Base Dates & Today Reference
        today_date = p.summary_dict.get("Today's Date")
        if isinstance(today_date, str):
            try:
                today_date = pd.to_datetime(today_date).to_pydatetime()
            except:
                today_date = datetime.now()
        elif not today_date or pd.isna(today_date):
            today_date = datetime.now()
        else:
            try:
                today_date = pd.to_datetime(today_date).to_pydatetime()
            except:
                today_date = datetime.now()

        start_date = p.project_start_date
        end_date = p.project_end_date

        # 2. Time elapsed and progress
        elapsed_days = (today_date - start_date).days if start_date else 0
        total_days = (end_date - start_date).days if start_date and end_date else 1
        time_elapsed_pct = elapsed_days / total_days if total_days > 0 else 0.0
        time_elapsed_pct = max(0.0, min(1.0, time_elapsed_pct))

        # Schedule Slippage
        slippage = time_elapsed_pct - p.percent_complete

        # SPI
        spi = p.percent_complete / time_elapsed_pct if time_elapsed_pct > 0 else 1.0

        # 3. Status and Health Counts
        status_counts = {"Completed": 0, "In Progress": 0, "Not Started": 0, "On Hold": 0, "Not Applicable": 0}
        health_counts = {"Green": 0, "Yellow": 0, "Red": 0}
        
        overdue_tasks = []
        blockers = []
        critical_incomplete = 0
        on_hold_count = 0
        
        # Blocker & Sentiment keywords
        risk_keywords = [
            "delay", "delayed", "impacted", "pending", "blocker", "issue", "risk", 
            "dependency", "waiting", "need", "not done", "missing", "failed", 
            "problem", "concern", "on hold", "escalate", "escalation", "stuck"
        ]
        
        positive_keywords = ["completed", "done", "covered", "resolved", "on track", "signed off", "approved"]

        neg_comments_count = 0
        pos_comments_count = 0
        all_comments_list = []

        for t in tasks:
            status = t.status
            if status in status_counts:
                status_counts[status] += 1
            else:
                status_counts["Not Started"] += 1

            if status == "On Hold":
                on_hold_count += 1

            health = t.schedule_health
            if health in health_counts:
                health_counts[health] += 1
            elif health == "Amber" or health == "Yellow":
                health_counts["Yellow"] += 1
            else:
                health_counts["Green"] += 1

            # Check if task is overdue
            if t.end_date and t.end_date < today_date and t.status != "Completed":
                overdue_tasks.append(t)

            # Check if critical path incomplete
            if t.level <= 2 and t.status != "Completed":
                critical_incomplete += 1

            # Blocker criteria: On Hold, At Risk, or negative comments
            is_blocked = False
            if t.status == "On Hold" or t.at_risk or t.schedule_health == "Red":
                is_blocked = True
            
            # Analyze comments
            for c in t.comments:
                c_text = c.comment.lower()
                all_comments_list.append({
                    "task_name": t.task_name,
                    "excel_row": t.excel_row,
                    "comment": c.comment,
                    "user": c.user,
                    "date": c.date
                })
                
                if any(w in c_text for w in risk_keywords):
                    neg_comments_count += 1
                    is_blocked = True
                if any(w in c_text for w in positive_keywords):
                    pos_comments_count += 1

            if is_blocked:
                blockers.append(t)

        # Overdue milestones
        overdue_milestones = [t for t in overdue_tasks if t.level in [1, 2] or t.is_milestone]

        # Calculate Scores (0 to 100 scale for each)
        
        # A. Schedule Health & Variance Score (30%)
        if slippage <= 0.05:
            schedule_score = 100
        elif slippage <= 0.15:
            schedule_score = 50
        else:
            schedule_score = 0
        
        # B. Progress & Milestone Score (25%)
        # Scale score with ratio of overdue milestones
        total_milestones = len([t for t in tasks if t.level in [1, 2] or t.is_milestone])
        milestone_ratio = len(overdue_milestones) / total_milestones if total_milestones > 0 else 0.0
        
        if len(overdue_milestones) == 0:
            milestone_score = 100
        elif milestone_ratio <= 0.05 or len(overdue_milestones) <= 2:
            milestone_score = 80
        elif milestone_ratio <= 0.15 or len(overdue_milestones) <= 5:
            milestone_score = 50
        else:
            milestone_score = 30
            
        # C. Risks & Blockers Score (20%)
        # Scale blockers score with total task count ratio
        total_tasks_count = len(tasks)
        blocker_ratio = len(blockers) / total_tasks_count if total_tasks_count > 0 else 0.0
        
        if len(blockers) == 0:
            blocker_score = 100
        elif blocker_ratio <= 0.03 or len(blockers) <= 3:
            blocker_score = 80
        elif blocker_ratio <= 0.06 or len(blockers) <= 10:
            blocker_score = 50
        else:
            blocker_score = 30

        # D. Budget Burn Score (10%)
        # Treated as neutral (100) if no budget columns or SPI cannot be computed
        # Or if SPI is reasonable
        if spi >= 0.95:
            budget_score = 100
        elif spi >= 0.80:
            budget_score = 50
        else:
            budget_score = 0
            
        # E. Stakeholder Sentiment Score (10%)
        # Scale based on count and ratio
        if len(all_comments_list) == 0:
            sentiment_score = 100  # Neutral / Green
        else:
            if neg_comments_count == 0:
                sentiment_score = 100
            elif neg_comments_count <= 2:
                sentiment_score = 80
            elif neg_comments_count <= 5:
                sentiment_score = 65
            elif neg_comments_count <= 10:
                sentiment_score = 50

            else:
                sentiment_score = 30

        # F. Data Quality Score (5%)
        dq_score = 100
        for note in p.data_quality_notes:
            if "missing" in note.lower() and "summary" in note.lower():
                dq_score -= 20
            elif "missing" in note.lower() and "comments" in note.lower():
                dq_score -= 10
            elif "could not be parsed" in note.lower():
                dq_score -= 5
            elif "missing plan sheet columns" in note.lower():
                dq_score -= 15
        dq_score = max(0, dq_score)

        # Weighted aggregate
        health_score = (
            (0.30 * schedule_score) +
            (0.25 * milestone_score) +
            (0.20 * blocker_score) +
            (0.10 * budget_score) +
            (0.10 * sentiment_score) +
            (0.05 * dq_score)
        )

        # Determine RAG Color
        if health_score >= 80:
            rag_color = "Green"
        elif health_score >= 60:
            rag_color = "Amber"
        else:
            rag_color = "Red"

        # Apply overrides
        # 1. Project-level schedule health override
        summary_health = p.schedule_health
        if summary_health == "Red" or summary_health == "Critical":
            rag_color = "Red"
        elif (summary_health == "Yellow" or summary_health == "Amber") and rag_color == "Green":
            rag_color = "Amber"

        # 2. Severe schedule slippage override
        if slippage > 0.25:
            rag_color = "Red"

        # 3. High blocker override
        # Only override to Red if the blocker count represents >10% of tasks AND slippage is positive
        if len(blockers) >= 15 and blocker_ratio >= 0.10 and slippage > 0.05:
            rag_color = "Red"


        # Safe serialized format for output
        self.metrics = {
            "project_name": p.project_name,
            "project_manager": p.project_manager,
            "project_stage": p.project_stage,
            "start_date": start_date.strftime("%Y-%m-%d") if start_date else "N/A",
            "end_date": end_date.strftime("%Y-%m-%d") if end_date else "N/A",
            "today_date": today_date.strftime("%Y-%m-%d"),
            "pct_complete": p.percent_complete,
            "time_elapsed_pct": time_elapsed_pct,
            "schedule_slippage": slippage,
            "spi": spi,
            "status_counts": status_counts,
            "health_counts": health_counts,
            "total_tasks": len(tasks),
            "overdue_tasks_count": len(overdue_tasks),
            "overdue_milestones_count": len(overdue_milestones),
            "blockers_count": len(blockers),
            "critical_incomplete": critical_incomplete,
            "on_hold_count": on_hold_count,
            "neg_comments_count": neg_comments_count,
            "pos_comments_count": pos_comments_count,
            "calculated_score": health_score,
            "rag_status": rag_color,
            "overdue_milestones": [
                {
                    "task_name": t.task_name,
                    "end_date": t.end_date.strftime("%Y-%m-%d") if t.end_date else "N/A",
                    "status": t.status
                } for t in overdue_milestones
            ],
            "blockers": [
                {
                    "task_name": t.task_name,
                    "status": t.status,
                    "excel_row": t.excel_row,
                    "comments": [c.dict() for c in t.comments]
                } for t in blockers
            ],
            "recent_comments": all_comments_list,
            "data_quality_notes": p.data_quality_notes
        }
        return self.metrics
