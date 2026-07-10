import unittest
from datetime import datetime, timedelta
from src.excel_reader import ProjectData, TaskInfo, CommentInfo
from src.health_rules import HealthRulesEvaluator

class TestHealthRulesEvaluator(unittest.TestCase):
    def setUp(self):
        # Create standard healthy tasks
        self.today = datetime(2026, 7, 10)
        self.start = self.today - timedelta(days=20)
        self.end = self.today + timedelta(days=20)
        
        self.task1 = TaskInfo(
            pandas_index=0,
            excel_row=2,
            task_name="Integration Kickoff",
            status="Completed",
            pct_complete=1.0,
            start_date=self.start,
            end_date=self.today - timedelta(days=5),
            schedule_health="Green",
            at_risk=False,
            level=1,
            is_milestone=True,
            comments=[]
        )
        
        self.task2 = TaskInfo(
            pandas_index=1,
            excel_row=3,
            task_name="Configuration Design",
            status="In Progress",
            pct_complete=0.5,
            start_date=self.today - timedelta(days=5),
            end_date=self.end,
            schedule_health="Green",
            at_risk=False,
            level=2,
            is_milestone=False,
            comments=[]
        )

    def test_green_project(self):
        # 50% elapsed, 75% complete overall -> ahead -> Green
        project_data = ProjectData(
            project_name="Test Green Project",
            project_manager="John Doe",
            project_start_date=self.start,
            project_end_date=self.end,
            percent_complete=0.75,
            schedule_health="Green",
            project_stage="Design",
            at_risk="No",
            tasks=[self.task1, self.task2],
            summary_dict={"Today's Date": "2026-07-10"},
            data_quality_notes=[]
        )
        
        evaluator = HealthRulesEvaluator(project_data)
        metrics = evaluator.evaluate()
        
        self.assertEqual(metrics["rag_status"], "Green")
        self.assertGreaterEqual(metrics["calculated_score"], 80)

    def test_red_due_to_slippage(self):
        # 50% elapsed, only 10% complete overall -> severe slippage -> Red
        project_data = ProjectData(
            project_name="Test Red Slippage Project",
            project_manager="John Doe",
            project_start_date=self.start,
            project_end_date=self.end,
            percent_complete=0.10,
            schedule_health="Green",
            project_stage="Design",
            at_risk="Yes",
            tasks=[self.task1, self.task2],
            summary_dict={"Today's Date": "2026-07-10"},
            data_quality_notes=[]
        )
        
        evaluator = HealthRulesEvaluator(project_data)
        metrics = evaluator.evaluate()
        
        self.assertEqual(metrics["rag_status"], "Red")

    def test_red_override_pmo_health(self):
        # Standard green numbers, but explicit Red health from Summary overrides to Red
        project_data = ProjectData(
            project_name="Test Override Project",
            project_manager="John Doe",
            project_start_date=self.start,
            project_end_date=self.end,
            percent_complete=0.80,
            schedule_health="Red",
            project_stage="Design",
            at_risk="No",
            tasks=[self.task1, self.task2],
            summary_dict={"Today's Date": "2026-07-10"},
            data_quality_notes=[]
        )
        
        evaluator = HealthRulesEvaluator(project_data)
        metrics = evaluator.evaluate()
        
        self.assertEqual(metrics["rag_status"], "Red")

    def test_missing_budget_neutral(self):
        # SPI cannot be calculated or is 0, check that budget score is treated as neutral (not failing the score)
        project_data = ProjectData(
            project_name="Test Budget Neutral",
            project_manager="John Doe",
            project_start_date=self.today, # Starts today -> 0 elapsed days
            project_end_date=self.end,
            percent_complete=0.00,
            schedule_health="Green",
            project_stage="Kickoff",
            at_risk="No",
            tasks=[self.task2],
            summary_dict={"Today's Date": "2026-07-10"},
            data_quality_notes=["Missing budget ledger data column."]
        )
        
        evaluator = HealthRulesEvaluator(project_data)
        metrics = evaluator.evaluate()
        
        self.assertEqual(metrics["rag_status"], "Green")
        self.assertGreaterEqual(metrics["calculated_score"], 80)

if __name__ == '__main__':
    unittest.main()
