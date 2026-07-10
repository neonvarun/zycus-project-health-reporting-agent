import os
import io
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

from src.utils import load_dotenv
from src.excel_reader import ExcelReader
from src.health_rules import HealthRulesEvaluator
from src.explainer import ProjectHealthExplainer
from src.local_model import LocalModelConfig, is_model_cached
from src.report_generator import ReportGenerator
from src.ppt_generator import PPTPortfolioGenerator

# Page config
st.set_page_config(
    page_title="Project Health Reporting Agent",
    page_icon="🟢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom premium styling
st.markdown("""
<style>
    .main {
        background-color: #F8FAFC;
    }
    .reportview-container {
        background-color: #F8FAFC;
    }
    .stCard {
        background-color: #FFFFFF;
        padding: 1.5rem;
        border-radius: 12px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        margin-bottom: 1.5rem;
    }
    h1, h2, h3 {
        color: #0F172A;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .rag-badge {
        padding: 6px 12px;
        border-radius: 8px;
        font-weight: bold;
        display: inline-block;
        font-size: 14px;
        text-align: center;
    }
    .rag-green {
        background-color: #D1FAEC;
        color: #065F46;
        border: 1px solid #A7F3D0;
    }
    .rag-amber {
        background-color: #FEF3C7;
        color: #92400E;
        border: 1px solid #FDE68A;
    }
    .rag-red {
        background-color: #FEE2E2;
        color: #991B1B;
        border: 1px solid #FCA5A5;
    }
    .sidebar-title {
        font-size: 1.2rem;
        font-weight: bold;
        color: #4F46E5;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Load env configurations
load_dotenv()

st.title("💼 Project Health Reporting Agent")
st.markdown("Automated PMO dashboard mapping project slippage, milestones, blockers, and sentiment into client-ready deliverables.")

# Show the active explanation provider and local model state.
llm_provider = os.getenv("LLM_PROVIDER", "local").strip().lower()
if llm_provider == "local":
    local_config = LocalModelConfig.from_environment()
    if is_model_cached(local_config):
        st.sidebar.success("🧠 Local Qwen model: READY")
    else:
        st.sidebar.warning("🧠 Local Qwen model will download on first analysis (~600 MB).")
elif llm_provider == "gemini":
    if os.getenv("GEMINI_API_KEY"):
        st.sidebar.success("🔑 Gemini LLM Integration: ACTIVE")
    else:
        st.sidebar.warning("Gemini is selected but GEMINI_API_KEY is missing; rules fallback will be used.")
elif llm_provider == "rules":
    st.sidebar.info("ℹ️ Rules-based explanations: ACTIVE")
else:
    st.sidebar.warning(f"Unsupported LLM_PROVIDER '{llm_provider}'; rules fallback will be used.")

st.sidebar.markdown("<div class='sidebar-title'>📁 Data Sources</div>", unsafe_allow_html=True)

# Upload Files or Use Preloaded Samples
uploaded_files = st.sidebar.file_uploader(
    "Upload Project Schedule Excel Files (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True
)

st.sidebar.markdown("---")
use_preloaded = st.sidebar.checkbox("Use Sample Project Files", value=True)

# Local Sample files checking
sample_dir = "data"
sample_files = []
if use_preloaded and os.path.exists(sample_dir):
    sample_files = [
        os.path.join(sample_dir, f) for f in os.listdir(sample_dir) 
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]

# Core analysis collection
analyzed_projects = []

# Triggers for analysis
has_data = len(uploaded_files) > 0 or len(sample_files) > 0

if has_data:
    # Compile sources
    analysis_sources = []
    
    # 1. Add uploaded files
    for f in uploaded_files:
        analysis_sources.append(("upload", f))
        
    # 2. Add sample files if selected and not already uploaded
    uploaded_names = [f.name for f in uploaded_files]
    for s in sample_files:
        s_name = os.path.basename(s)
        if s_name not in uploaded_names:
            analysis_sources.append(("sample", s))

    # Process all sources
    for source_type, src in analysis_sources:
        try:
            reader = ExcelReader(src)
            project_data = reader.read_project()
            
            # Evaluate scoring rules
            evaluator = HealthRulesEvaluator(project_data)
            metrics = evaluator.evaluate()
            
            # Generate explanation
            explainer = ProjectHealthExplainer(metrics)
            reasoning = explainer.generate_explanation()
            metrics["reasoning"] = reasoning
            
            analyzed_projects.append(metrics)
        except Exception as e:
            st.error(f"Error processing {getattr(src, 'name', os.path.basename(str(src)))}: {e}")

# Page Navigation tabs
tab_dashboard, tab_projects, tab_risks, tab_downloads, tab_methodology = st.tabs([
    "📈 Portfolio Dashboard",
    "📂 Project Cards",
    "⚠️ Risks & Blockers",
    "📥 Deliverables & Downloads",
    "📜 RAG Methodology"
])

# ----------------------------------------------------
# TAB 1: Portfolio Dashboard
# ----------------------------------------------------
with tab_dashboard:
    if not analyzed_projects:
        st.warning("No projects analyzed. Please upload Excel files or enable sample files in the sidebar.")
    else:
        st.subheader("Portfolio Status Overview")
        
        # Portfolio KPI Summary Cards
        total_p = len(analyzed_projects)
        red_cnt = sum(1 for p in analyzed_projects if p["rag_status"] == "Red")
        amber_cnt = sum(1 for p in analyzed_projects if p["rag_status"] == "Amber")
        green_cnt = sum(1 for p in analyzed_projects if p["rag_status"] == "Green")
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Projects", total_p)
        col2.metric("Green Projects", green_cnt, delta=None)
        col3.metric("Amber Projects", amber_cnt, delta=None)
        col4.metric("Red Projects", red_cnt, delta=None)
        
        col_chart_l, col_chart_r = st.columns(2)
        
        with col_chart_l:
            # RAG Distribution Pie Chart
            fig_rag = px.pie(
                values=[green_cnt, amber_cnt, red_cnt],
                names=["Green", "Amber", "Red"],
                color=["Green", "Amber", "Red"],
                color_discrete_map={
                    "Green": "#10B981",
                    "Amber": "#F59E0B",
                    "Red": "#EF4444"
                },
                title="RAG Distribution",
                hole=0.4
            )
            st.plotly_chart(fig_rag, use_container_width=True)
            
        with col_chart_r:
            # Progress vs Schedule Elapsed Scatter
            scatter_data = []
            for p in analyzed_projects:
                scatter_data.append({
                    "Project": p["project_name"].upper(),
                    "Progress (%)": p["pct_complete"] * 100,
                    "Schedule Elapsed (%)": p["time_elapsed_pct"] * 100,
                    "RAG Status": p["rag_status"],
                    "Slippage (%)": p["schedule_slippage"] * 100
                })
            df_scatter = pd.DataFrame(scatter_data)
            
            fig_scatter = px.scatter(
                df_scatter,
                x="Schedule Elapsed (%)",
                y="Progress (%)",
                color="RAG Status",
                size=[10] * len(df_scatter),
                hover_name="Project",
                color_discrete_map={
                    "Green": "#10B981",
                    "Amber": "#F59E0B",
                    "Red": "#EF4444"
                },
                title="Progress vs. Schedule Elapsed"
            )
            # Add y=x reference line
            fig_scatter.add_trace(
                go.Scatter(
                    x=[0, 100], y=[0, 100],
                    mode="lines",
                    line=dict(color="#94A3B8", dash="dash"),
                    name="On Target"
                )
            )
            st.plotly_chart(fig_scatter, use_container_width=True)

        # Portfolio Data Table
        st.subheader("Analyzed Projects Summary")
        summary_table = []
        for p in analyzed_projects:
            summary_table.append({
                "Project": p["project_name"].upper(),
                "Project Manager": p["project_manager"],
                "RAG Status": "🟢 Green" if p["rag_status"] == "Green" else ("🟠 Amber" if p["rag_status"] == "Amber" else "🔴 Red"),
                "Progress": f"{p['pct_complete']*100:.1f}%",
                "Elapsed Schedule": f"{p['time_elapsed_pct']*100:.1f}%",
                "Slippage": f"{p['schedule_slippage']*100:.1f}%",
                "SPI": f"{p['spi']:.2f}",
                "Active Blockers": p["blockers_count"],
                "Overdue Milestones": p["overdue_milestones_count"]
            })
        st.dataframe(pd.DataFrame(summary_table), use_container_width=True, hide_index=True)

# ----------------------------------------------------
# TAB 2: Project Cards
# ----------------------------------------------------
with tab_projects:
    if not analyzed_projects:
        st.warning("No projects analyzed. Please upload Excel files or enable sample files in the sidebar.")
    else:
        st.subheader("Individual Project Health Details")
        
        for idx, proj in enumerate(analyzed_projects):
            # Select color variables
            rag = proj["rag_status"]
            badge_class = "rag-green" if rag == "Green" else ("rag-amber" if rag == "Amber" else "rag-red")
            badge_icon = "🟢" if rag == "Green" else ("🟠" if rag == "Amber" else "🔴")
            
            # Project card container
            st.markdown(f"""
            <div class="stCard">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                    <span style="font-size: 20px; font-weight: bold; color: #0F172A;">{proj['project_name'].upper()}</span>
                    <span class="rag-badge {badge_class}">{badge_icon} {rag.upper()} STATUS</span>
                </div>
                <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; margin-bottom: 1.5rem;">
                    <div>
                        <span style="font-size: 12px; color: #64748B;">PROJECT MANAGER</span><br/>
                        <strong style="color: #1E293B;">{proj['project_manager']}</strong>
                    </div>
                    <div>
                        <span style="font-size: 12px; color: #64748B;">STAGE</span><br/>
                        <strong style="color: #1E293B;">{proj['project_stage']}</strong>
                    </div>
                    <div>
                        <span style="font-size: 12px; color: #64748B;">COMPLETION</span><br/>
                        <strong style="color: #1E293B;">{proj['pct_complete']*100:.1f}%</strong>
                    </div>
                    <div>
                        <span style="font-size: 12px; color: #64748B;">TIMELINE SLIPPAGE</span><br/>
                        <strong style="color: #1E293B;">{proj['schedule_slippage']*100:.1f}%</strong>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # Explander for reasoning and tasks details
            with st.expander(f"🔍 Detailed Analysis for {proj['project_name'].upper()}", expanded=(idx == 0)):
                st.markdown(proj["reasoning"])
                
                # Show status sub-counts
                st.markdown("**Subtask Completion Distribution**")
                col_c1, col_c2, col_c3, col_c4 = st.columns(4)
                counts = proj["status_counts"]
                col_c1.metric("Completed", counts.get("Completed", 0))
                col_c2.metric("In Progress", counts.get("In Progress", 0))
                col_c3.metric("Not Started", counts.get("Not Started", 0))
                col_c4.metric("On Hold", counts.get("On Hold", 0))
            st.markdown("<br/>", unsafe_allow_html=True)

# ----------------------------------------------------
# TAB 3: Risks & Blockers Analysis
# ----------------------------------------------------
with tab_risks:
    if not analyzed_projects:
        st.warning("No projects analyzed. Please upload Excel files or enable sample files in the sidebar.")
    else:
        st.subheader("Emerging Portfolio Risks & Critical Blockers")
        
        # Display aggregated list of active blockers
        for proj in analyzed_projects:
            st.markdown(f"#### ⚠️ {proj['project_name'].upper()} Critical Issues")
            
            col_b1, col_b2 = st.columns(2)
            
            with col_b1:
                st.markdown(f"**Overdue Milestones ({proj['overdue_milestones_count']})**")
                if proj["overdue_milestones"]:
                    for ms in proj["overdue_milestones"]:
                        st.error(f"● **{ms['task_name']}** (Due: {ms['end_date']} | Status: {ms['status']})")
                else:
                    st.success("No overdue milestones identified.")
                    
            with col_b2:
                st.markdown(f"**Active Blockers & Flagged Risks ({proj['blockers_count']})**")
                if proj["blockers"]:
                    for idx, blk in enumerate(proj["blockers"][:5]):
                        st.warning(f"**{idx+1}. {blk['task_name']}** (Status: {blk['status']})")
                        if blk["comments"]:
                            st.markdown(f"   * *Notes*: {blk['comments'][0]['comment']} (by {blk['comments'][0]['user']})")
                else:
                    st.success("No active blockers flagged in comments.")
            
            st.markdown("---")

# ----------------------------------------------------
# TAB 4: Deliverables & Downloads
# ----------------------------------------------------
with tab_downloads:
    if not analyzed_projects:
        st.warning("No projects analyzed. Please upload Excel files or enable sample files in the sidebar.")
    else:
        st.subheader("Generate Weekly & Monthly Reports")
        st.markdown("Compile all analyzed project schedules into professional weekly summaries and a monthly synthesis executive deck.")
        
        col_dl1, col_dl2 = st.columns(2)
        
        with col_dl1:
            st.markdown("### 📅 Weekly Reports (PowerPoint / Markdown / JSON)")
            
            # Select project for download
            proj_names = [p["project_name"] for p in analyzed_projects]
            selected_proj_name = st.selectbox("Select Project to Preview & Download", proj_names)
            selected_proj = next(p for p in analyzed_projects if p["project_name"] == selected_proj_name)
            
            # Generate Report in-memory for download
            weekly_temp_dir = "outputs/weekly"
            os.makedirs(weekly_temp_dir, exist_ok=True)
            generator = ReportGenerator(selected_proj, weekly_temp_dir)
            json_filepath = generator.generate_json_report()
            md_filepath = generator.generate_markdown_report()
            weekly_ppt_filepath = generator.generate_powerpoint_report()
            
            # Preview report
            st.markdown("**Weekly Report Preview (Markdown)**")
            with open(md_filepath, "r", encoding="utf-8") as f:
                report_content = f.read()
            st.text_area("Markdown Preview", report_content, height=300)
            
            # Download buttons
            col_b1, col_b2, col_b3 = st.columns(3)
            with col_b1:
                st.download_button(
                    label="Download Markdown Report",
                    data=report_content,
                    file_name=f"{selected_proj_name}_weekly_report.md",
                    mime="text/markdown"
                )
            with col_b2:
                with open(json_filepath, "r", encoding="utf-8") as jf:
                    json_data = jf.read()
                st.download_button(
                    label="Download JSON Metrics",
                    data=json_data,
                    file_name=f"{selected_proj_name}_weekly_report.json",
                    mime="application/json"
                )
            with col_b3:
                with open(weekly_ppt_filepath, "rb") as weekly_ppt_file:
                    weekly_ppt_data = weekly_ppt_file.read()
                st.download_button(
                    label="Download Weekly PowerPoint",
                    data=weekly_ppt_data,
                    file_name=f"{selected_proj_name}_weekly_report.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                
        with col_dl2:
            st.markdown("### 📊 Monthly Synthesis presentation")
            st.markdown("""
            Generates a high-contrast widescreen (16:9) corporate executive slide deck covering:
            - Portfolio Health RAG Distribution
            - Individual Project Audits
            - Emerging Cross-Project Risks & Delivery Themes
            - Strategic Action Items & Recovery Recommendations
            """)
            
            monthly_temp_dir = "outputs/monthly"
            ppt_filepath = os.path.join(monthly_temp_dir, "project_health_monthly_synthesis.pptx")
            
            # Generate presentation button
            if st.button("🏗️ Compile Monthly Synthesis Slide Deck"):
                try:
                    ppt_builder = PPTPortfolioGenerator(analyzed_projects, ppt_filepath)
                    ppt_builder.generate()
                    st.success("Monthly PowerPoint synthesis compiled successfully!")
                    
                    with open(ppt_filepath, "rb") as pf:
                        ppt_bytes = pf.read()
                        
                    st.download_button(
                        label="📥 Download Executive Presentation (.pptx)",
                        data=ppt_bytes,
                        file_name="project_health_monthly_synthesis.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                except Exception as e:
                    st.error(f"Error compiling presentation: {e}")

# ----------------------------------------------------
# TAB 5: Methodology Viewer
# ----------------------------------------------------
with tab_methodology:
    st.subheader("RAG Calculation Methodology Framework")
    st.markdown("Below is the Professional Services PMO Framework explaining how Excel-based schedules translate to status metrics.")
    
    # Load docs/rag_methodology.md
    methodology_path = "docs/rag_methodology.md"
    if os.path.exists(methodology_path):
        with open(methodology_path, "r", encoding="utf-8") as f:
            meth_content = f.read()
        st.markdown(meth_content)
    else:
        st.info("Methodology document not found under docs/rag_methodology.md.")
