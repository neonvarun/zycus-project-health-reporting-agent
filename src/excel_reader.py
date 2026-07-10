import os
import re
import pandas as pd
from typing import List, Dict, Any, Optional
from datetime import datetime
from pydantic import BaseModel, Field
from src.utils import parse_date, clean_text

class CommentInfo(BaseModel):
    comment: str
    user: str = "Unknown"
    date: str = "Unknown"

class TaskInfo(BaseModel):
    pandas_index: int
    excel_row: int
    task_name: str
    status: str = "Not Started"
    pct_complete: float = 0.0
    start_date: Optional[datetime] = None
    end_date: Optional[datetime] = None
    schedule_health: str = "Green"
    at_risk: bool = False
    level: int = 3
    is_milestone: bool = False
    comments: List[CommentInfo] = []
    rag_column_value: Optional[str] = None

class ProjectData(BaseModel):
    project_name: str
    project_manager: str = "Unknown"
    project_start_date: Optional[datetime] = None
    project_end_date: Optional[datetime] = None
    percent_complete: float = 0.0
    schedule_health: str = "Green"
    project_stage: str = "Unknown"
    at_risk: str = "No"
    tasks: List[TaskInfo] = []
    summary_dict: Dict[str, Any] = {}
    data_quality_notes: List[str] = []
    precalculated_metrics: Optional[Dict[str, Any]] = None

class ExcelReader:
    def __init__(self, file_source):
        """
        file_source: can be a file path (str/Path) or a file-like object/bytes from Streamlit file uploader.
        """
        self.file_source = file_source
        if isinstance(file_source, str):
            self.project_name = os.path.basename(file_source).replace(".xlsx", "")
        else:
            self.project_name = getattr(file_source, "name", "UploadedProject").replace(".xlsx", "")
            
        self.raw_excel = None
        self.summary_dict = {}
        self.data_quality_notes = []

    def read_project(self) -> ProjectData:
        """Reads and normalizes Excel data, returning a clean ProjectData object."""
        try:
            self.raw_excel = pd.ExcelFile(self.file_source)
        except Exception as e:
            # Fallback to pre-calculated JSON weekly report if available
            import json
            json_filename = f"{self.project_name}_weekly_report.json"
            json_path = os.path.join("outputs", "weekly", json_filename)
            if os.path.exists(json_path):
                try:
                    with open(json_path, "r", encoding="utf-8") as f:
                        metrics = json.load(f)
                    
                    start_date = None
                    if metrics.get("start_date") and metrics["start_date"] != "N/A":
                        start_date = datetime.strptime(metrics["start_date"], "%Y-%m-%d")
                    end_date = None
                    if metrics.get("end_date") and metrics["end_date"] != "N/A":
                        end_date = datetime.strptime(metrics["end_date"], "%Y-%m-%d")
                        
                    tasks = []
                    for b in metrics.get("blockers", []):
                        tasks.append(TaskInfo(
                            pandas_index=0,
                            excel_row=b.get("excel_row", 2),
                            task_name=b.get("task_name", "Blocked Task"),
                            status=b.get("status", "In Progress"),
                            pct_complete=0.0,
                            schedule_health="Red",
                            at_risk=True
                        ))
                    
                    return ProjectData(
                        project_name=metrics.get("project_name", self.project_name),
                        project_manager=metrics.get("project_manager", "Unknown"),
                        project_start_date=start_date,
                        project_end_date=end_date,
                        percent_complete=metrics.get("pct_complete", 0.0),
                        schedule_health=metrics.get("rag_status", "Green"),
                        project_stage=metrics.get("project_stage", "Unknown"),
                        at_risk="Yes" if metrics.get("blockers_count", 0) > 0 else "No",
                        tasks=tasks,
                        summary_dict={"Today's Date": metrics.get("today_date", "2026-07-02")},
                        data_quality_notes=metrics.get("data_quality_notes", []),
                        precalculated_metrics=metrics
                    )
                except Exception as json_err:
                    raise ValueError(f"Failed to load Excel file and fallback JSON: {e} | JSON err: {json_err}")
            raise ValueError(f"Failed to load Excel file: {e}")

        # 1. Parse Summary sheet
        self._parse_summary_sheet()

        # 2. Find and Parse main project plan sheet
        main_sheet_name = self._detect_main_sheet()
        if not main_sheet_name:
            raise ValueError(f"Could not identify main project plan sheet in {self.project_name}")

        self.data_quality_notes.append(f"Auto-detected main project plan sheet: '{main_sheet_name}'")
        main_df = self.raw_excel.parse(main_sheet_name)

        # 3. Parse Comments sheet if present
        comments_map = self._parse_comments_sheet()

        # 4. Normalize columns and parse tasks
        tasks = self._parse_tasks(main_df, comments_map)

        # 5. Extract basic project info from summary
        pm = self.summary_dict.get("Project Manager", self.summary_dict.get("PM", "Unknown"))
        stage = self.summary_dict.get("Project Stage", self.summary_dict.get("Stage", "Unknown"))
        at_risk_val = self.summary_dict.get("At Risk", self.summary_dict.get("At Risk?", "No"))
        
        # Read Project-level health
        proj_health = self.summary_dict.get("Schedule Health", self.summary_dict.get("Project Status", "Green"))

        # Safe date conversion
        start_date = parse_date(self.summary_dict.get("Project Start Date", self.summary_dict.get("Start Date")))
        end_date = parse_date(self.summary_dict.get("Project End Date", self.summary_dict.get("End Date")))
        if not start_date and tasks:
            valid_starts = [t.start_date for t in tasks if t.start_date]
            if valid_starts:
                start_date = min(valid_starts)
                self.data_quality_notes.append("Project Start Date missing from summary; derived from minimum task start date.")
        if not end_date and tasks:
            valid_ends = [t.end_date for t in tasks if t.end_date]
            if valid_ends:
                end_date = max(valid_ends)
                self.data_quality_notes.append("Project End Date missing from summary; derived from maximum task end date.")

        # Read overall progress
        overall_pct = 0.0
        pct_val = self.summary_dict.get("% Complete", self.summary_dict.get("Percent Complete", 0.0))
        if isinstance(pct_val, str):
            try:
                overall_pct = float(pct_val.replace("%", "").strip()) / 100.0
            except:
                overall_pct = 0.0
        elif isinstance(pct_val, (int, float)):
            overall_pct = float(pct_val)
            # If stored as e.g., 71 instead of 0.71
            if overall_pct > 1.0:
                overall_pct /= 100.0
        
        return ProjectData(
            project_name=self.project_name,
            project_manager=str(pm),
            project_start_date=start_date,
            project_end_date=end_date,
            percent_complete=overall_pct,
            schedule_health=str(proj_health),
            project_stage=str(stage),
            at_risk=str(at_risk_val),
            tasks=tasks,
            summary_dict=self.summary_dict,
            data_quality_notes=self.data_quality_notes
        )

    def _parse_summary_sheet(self):
        """Parses the 'Summary' sheet if available, extracting key-value metadata."""
        if "Summary" in self.raw_excel.sheet_names:
            try:
                summary_df = self.raw_excel.parse("Summary")
                for _, row in summary_df.iterrows():
                    if len(row) >= 2:
                        k = str(row.iloc[0]).strip()
                        v = row.iloc[1]
                        if k and k != "nan" and k != "Project Name":
                            # Handle unparseable/error cells
                            if isinstance(v, str) and (v.startswith("#") or "unparseable" in v.lower()):
                                self.data_quality_notes.append(f"Summary sheet cell '{k}' had unparseable Excel formula; treated as empty.")
                                v = None
                            self.summary_dict[k] = v
            except Exception as e:
                self.data_quality_notes.append(f"Failed to parse 'Summary' sheet: {e}")
        else:
            self.data_quality_notes.append("Summary sheet is missing. Using defaults for metadata.")

    def _detect_main_sheet(self) -> Optional[str]:
        """Detects the main project plan sheet."""
        # 1. Look for sheets containing 'project' or 'plan' (case insensitive)
        for name in self.raw_excel.sheet_names:
            if name not in ["Summary", "Comments"] and ("project" in name.lower() or "plan" in name.lower()):
                return name
        # 2. Fallback to first sheet that isn't Summary or Comments
        for name in self.raw_excel.sheet_names:
            if name not in ["Summary", "Comments"]:
                return name
        return None

    def _parse_comments_sheet(self) -> Dict[int, List[CommentInfo]]:
        """Parses the 'Comments' sheet and returns a mapping from pandas index to CommentInfo."""
        comments_map = {}
        if "Comments" in self.raw_excel.sheet_names:
            try:
                comments_df = self.raw_excel.parse("Comments", header=None)
                for _, row in comments_df.iterrows():
                    if len(row) >= 2 and pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]):
                        row_ref = str(row.iloc[0]).strip()
                        comment_text = str(row.iloc[1]).strip()
                        
                        # Safe parsing for author and date
                        author = str(row.iloc[2]).strip() if len(row) >= 3 and pd.notna(row.iloc[2]) else "Unknown"
                        timestamp = str(row.iloc[3]).strip() if len(row) >= 4 and pd.notna(row.iloc[3]) else "Unknown"
                        
                        # Match 'Row X'
                        match = re.search(r"Row\s+(\d+)", row_ref, re.IGNORECASE)
                        if match:
                            excel_row_num = int(match.group(1))
                            # Excel Row R corresponds to pandas index R-2
                            pandas_idx = excel_row_num - 2
                            
                            comment_obj = CommentInfo(
                                comment=comment_text,
                                user=author,
                                date=timestamp
                            )
                            if pandas_idx not in comments_map:
                                comments_map[pandas_idx] = []
                            comments_map[pandas_idx].append(comment_obj)
            except Exception as e:
                self.data_quality_notes.append(f"Failed to parse 'Comments' sheet: {e}")
        else:
            self.data_quality_notes.append("Comments sheet is missing. No external row comments linked.")
        return comments_map

    def _parse_tasks(self, main_df: pd.DataFrame, comments_map: Dict[int, List[CommentInfo]]) -> List[TaskInfo]:
        """Normalizes columns and parses rows from the main project plan sheet."""
        columns = main_df.columns.tolist()

        # Helper to find column matching patterns
        def find_col(patterns):
            for col in columns:
                for pat in patterns:
                    if re.search(pat, col, re.IGNORECASE):
                        return col
            return None

        # Detect columns with patterns
        task_name_col = find_col([r"^Task\s*Name", r"^Name"])
        status_col = find_col([r"^Status"])
        complete_col = find_col([r"^%\s*Complete", r"^Percent\s*Complete"])
        start_date_col = find_col([r"^Start\s*Date", r"^Baseline\s*Start", r"^Start"])
        end_date_col = find_col([r"^End\s*Date", r"^Baseline\s*Finish", r"^Finish"])
        health_col = find_col([r"^Schedule\s*Health", r"^Health"])
        at_risk_col = find_col([r"^At\s*Risk"])
        level_col = find_col([r"^Ancestors", r"^Level"])
        milestone_col = find_col([r"^Phase/Milestone", r"^Milestone"])
        rag_col = find_col([r"^RAG$"]) # Actual RAG column if present

        # Data quality checks for columns
        missing_cols = []
        if not task_name_col: missing_cols.append("Task Name")
        if not status_col: missing_cols.append("Status")
        if not complete_col: missing_cols.append("% Complete")
        if not start_date_col: missing_cols.append("Start Date")
        if not end_date_col: missing_cols.append("End Date")
        
        if missing_cols:
            self.data_quality_notes.append(f"Missing plan sheet columns: {', '.join(missing_cols)}. Fallbacks will be applied.")

        tasks = []
        for idx, row in main_df.iterrows():
            t_name = clean_text(row.get(task_name_col)) if task_name_col else ""
            # Skip empty rows
            if not t_name or t_name.lower() in ["nan", "null", "none"]:
                continue

            # Check for unparseable formulas in cells
            task_status = clean_text(row.get(status_col)) if status_col else "Not Started"
            if "unparseable" in task_status.lower() or task_status.startswith("#"):
                task_status = "Not Started"
                self.data_quality_notes.append(f"Row {idx+2}: status formula was unparseable. Defaulted to 'Not Started'.")

            # Parse completion percentage
            pct_complete = 0.0
            if complete_col:
                pct_val = row.get(complete_col)
                if pd.notna(pct_val):
                    if isinstance(pct_val, str):
                        if "unparseable" in pct_val.lower() or pct_val.startswith("#"):
                            pct_complete = 0.0
                            self.data_quality_notes.append(f"Row {idx+2}: % Complete was unparseable. Defaulted to 0%.")
                        else:
                            try:
                                pct_complete = float(pct_val.replace("%", "").strip()) / 100.0
                            except:
                                pct_complete = 0.0
                    elif isinstance(pct_val, (int, float)):
                        pct_complete = float(pct_val)
                        if pct_complete > 1.0:
                            pct_complete /= 100.0

            # Parse dates
            start_val = parse_date(row.get(start_date_col)) if start_date_col else None
            end_val = parse_date(row.get(end_date_col)) if end_date_col else None
            
            # Check for serialization warnings
            if start_date_col and pd.notna(row.get(start_date_col)) and not start_val:
                self.data_quality_notes.append(f"Row {idx+2}: Start Date '{row.get(start_date_col)}' could not be parsed.")
            if end_date_col and pd.notna(row.get(end_date_col)) and not end_val:
                self.data_quality_notes.append(f"Row {idx+2}: End Date '{row.get(end_date_col)}' could not be parsed.")

            # Schedule Health & At Risk
            t_health = clean_text(row.get(health_col)) if health_col else "Green"
            if not t_health or t_health.lower() in ["nan", "null"]:
                t_health = "Green"
                
            is_at_risk = False
            if at_risk_col:
                at_risk_raw = row.get(at_risk_col)
                if pd.notna(at_risk_raw):
                    if at_risk_raw in [1, 1.0, "Yes", "yes", "1", "true", True]:
                        is_at_risk = True

            # Hierarchical Levels
            level = 3
            if level_col:
                lvl_raw = row.get(level_col)
                try:
                    level = int(lvl_raw) if pd.notna(lvl_raw) else 3
                except:
                    level = 3

            # Milestone flag
            is_milestone = False
            if milestone_col:
                milestone_raw = row.get(milestone_col)
                if pd.notna(milestone_raw) and str(milestone_raw).strip() != "":
                    is_milestone = True

            # Comments parsing
            task_comments = comments_map.get(idx, [])
            inline_comment = clean_text(row.get("Comments")) if "Comments" in columns else ""
            if inline_comment and inline_comment.lower() not in ["nan", "null", "none"]:
                task_comments.append(CommentInfo(
                    comment=inline_comment,
                    user="System/Inline",
                    date="N/A"
                ))

            # RAG Column Value
            rag_val = clean_text(row.get(rag_col)) if rag_col else None

            tasks.append(TaskInfo(
                pandas_index=idx,
                excel_row=idx + 2,
                task_name=t_name,
                status=task_status,
                pct_complete=pct_complete,
                start_date=start_val,
                end_date=end_val,
                schedule_health=t_health,
                at_risk=is_at_risk,
                level=level,
                is_milestone=is_milestone,
                comments=task_comments,
                rag_column_value=rag_val
            ))

        return tasks
