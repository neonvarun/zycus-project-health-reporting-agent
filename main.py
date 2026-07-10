import os
import sys
import time
import argparse
from datetime import datetime
from src.utils import load_dotenv
from src.excel_reader import ExcelReader
from src.health_rules import HealthRulesEvaluator
from src.explainer import ProjectHealthExplainer
from src.report_generator import ReportGenerator
from src.ppt_generator import PPTPortfolioGenerator

# Load configuration
load_dotenv()

def run_reporting_pipeline(data_dir: str, output_dir: str):
    """Executes the analysis pipeline on all projects in data_dir."""
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Starting Project Health Reporting Agent...")
    print(f"Data Directory: {data_dir}")
    print(f"Output Directory: {output_dir}")
    
    # Subdirectories for outputs
    weekly_dir = os.path.join(output_dir, "weekly")
    monthly_dir = os.path.join(output_dir, "monthly")
    os.makedirs(weekly_dir, exist_ok=True)
    os.makedirs(monthly_dir, exist_ok=True)
    
    # Scan data_dir for xlsx
    if not os.path.exists(data_dir):
        print(f"Error: Data directory '{data_dir}' does not exist.")
        return
        
    excel_files = [f for f in os.listdir(data_dir) if f.endswith(".xlsx") and not f.startswith("~$")]
    if not excel_files:
        print(f"No project Excel plans (.xlsx) found in '{data_dir}'.")
        return
        
    print(f"Found {len(excel_files)} project file(s): {', '.join(excel_files)}")
    
    portfolio_summary = []
    
    for filename in excel_files:
        filepath = os.path.join(data_dir, filename)
        print(f"\nAnalyzing '{filename}'...")
        
        try:
            # 1. Read Excel
            reader = ExcelReader(filepath)
            project_data = reader.read_project()
            
            # 2. Evaluate Health Rules
            evaluator = HealthRulesEvaluator(project_data)
            metrics = evaluator.evaluate()
            
            # 3. Generate Explanation
            explainer = ProjectHealthExplainer(metrics)
            reasoning = explainer.generate_explanation()
            metrics["reasoning"] = reasoning
            
            # 4. Generate Weekly Outputs
            generator = ReportGenerator(metrics, weekly_dir)
            json_path = generator.generate_json_report()
            md_path = generator.generate_markdown_report()
            weekly_ppt_path = generator.generate_powerpoint_report()
            
            print(f"  Saved JSON: {json_path}")
            print(f"  Saved Markdown: {md_path}")
            print(f"  Saved weekly PowerPoint: {weekly_ppt_path}")
            
            # Add to portfolio metrics
            portfolio_summary.append(metrics)
            
        except Exception as e:
            print(f"  Error processing '{filename}': {e}")
            import traceback
            traceback.print_exc()

    if not portfolio_summary:
        print("\nNo projects successfully analyzed. Skipping PowerPoint generation.")
        return

    # 5. Generate Monthly PowerPoint Presentation
    ppt_path = os.path.join(monthly_dir, "project_health_monthly_synthesis.pptx")
    print(f"\nGenerating monthly synthesis PowerPoint...")
    try:
        ppt_builder = PPTPortfolioGenerator(portfolio_summary, ppt_path)
        ppt_builder.generate()
        print(f"  Saved presentation to {ppt_path}")
    except Exception as e:
        print(f"  Error generating PowerPoint: {e}")
        
    # 6. Print console summary
    print("\n" + "=" * 80)
    print(f" PORTFOLIO SUMMARY - {datetime.now().strftime('%Y-%m-%d')}")
    print("=" * 80)
    print(f"{'Project Name':<25} | {'Manager':<15} | {'RAG Status':<10} | {'Progress':<10} | {'Slippage':<10}")
    print("-" * 80)
    for p in portfolio_summary:
        print(f"{p['project_name'].upper():<25} | {p['project_manager']:<15} | {p['rag_status']:<10} | {p['pct_complete']*100:>7.1f}% | {p['schedule_slippage']*100:>7.1f}%")
    print("=" * 80)

def main():
    parser = argparse.ArgumentParser(description="CLI Project Health Reporting pipeline.")
    parser.add_argument("--data_dir", default="data", help="Directory with input Excel schedules")
    parser.add_argument("--output_dir", default="outputs", help="Base directory for output reports")
    parser.add_argument("--schedule", action="store_true", help="Run on a weekly/custom recurring schedule")
    parser.add_argument("--interval_seconds", type=int, default=604800, help="Interval in seconds for schedule mode (default: 1 week)")
    args = parser.parse_args()
    
    if args.schedule:
        print(f"Starting agent in schedule mode. Will run every {args.interval_seconds} seconds (approx. {args.interval_seconds/86400:.2f} days).")
        try:
            while True:
                run_reporting_pipeline(args.data_dir, args.output_dir)
                print(f"Waiting for next scheduled run in {args.interval_seconds} seconds...")
                time.sleep(args.interval_seconds)
        except KeyboardInterrupt:
            print("Schedule mode stopped by user.")
    else:
        run_reporting_pipeline(args.data_dir, args.output_dir)

if __name__ == "__main__":
    main()
