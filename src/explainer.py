import os
import json
import re
import requests
from typing import Dict, Any
from src.local_model import LocalModelConfig, get_local_model

class ProjectHealthExplainer:
    REQUIRED_SECTIONS = (
        "### Executive Summary",
        "### Why This Status Was Assigned",
        "### Top Risk Drivers",
        "### Recommended Actions",
        "### Data Quality & Assumptions",
    )

    def __init__(self, metrics: Dict[str, Any]):
        self.metrics = metrics

    def generate_explanation(self) -> str:
        """Generates a plain-English executive summary and explanation.
        Uses the configured local model by default and always has a deterministic
        rules-based fallback.
        """
        provider = os.getenv("LLM_PROVIDER", "local").strip().lower()

        if provider == "rules":
            return self._generate_rule_based_reasoning()

        try:
            if provider == "gemini":
                api_key = os.getenv("GEMINI_API_KEY")
                if not api_key:
                    raise ValueError("LLM_PROVIDER=gemini requires GEMINI_API_KEY")
                return self._generate_llm_reasoning(api_key)

            if provider == "local":
                return self._generate_local_reasoning()

            raise ValueError(f"Unsupported LLM_PROVIDER: {provider}")
        except Exception as e:
            print(f"Failed to generate {provider} reasoning, falling back to rule-based: {e}")
            return self._generate_rule_based_reasoning()

    def _build_context(self) -> Dict[str, Any]:
        """Build the model input from evaluator-owned metrics only."""
        m = self.metrics
        compact_blockers = [
            {
                **blocker,
                "comments": blocker.get("comments", [])[:1],
            }
            for blocker in m["blockers"][:6]
        ]
        return {
            "project_name": m["project_name"],
            "project_manager": m["project_manager"],
            "project_stage": m["project_stage"],
            "start_date": m["start_date"],
            "end_date": m["end_date"],
            "today_date": m["today_date"],
            "pct_complete": m["pct_complete"],
            "time_elapsed_pct": m["time_elapsed_pct"],
            "schedule_slippage": m["schedule_slippage"],
            "spi": m["spi"],
            "status_counts": m["status_counts"],
            "health_counts": m["health_counts"],
            "overdue_tasks_count": m["overdue_tasks_count"],
            "overdue_milestones_count": m["overdue_milestones_count"],
            "blockers_count": m["blockers_count"],
            "overdue_milestones": m["overdue_milestones"][:5],
            "blockers": compact_blockers,
            "data_quality_notes": m.get("data_quality_notes", []),
            "calculated_rag": m["rag_status"],
        }

    def _build_model_prompt(self) -> str:
        """Create a constrained prompt for either local or hosted reasoning."""
        m = self.metrics
        sections = "\n".join(f"- {section}" for section in self.REQUIRED_SECTIONS)
        return f"""
You are an expert Project Management Officer (PMO) Director reviewing a client delivery project.
Use only the structured project health data below. The calculated RAG status is authoritative:
{m['rag_status']}

Do not change the RAG status, invent facts, add unsupported project names, or recalculate the score.
Explain the supplied metrics in clear, professional, plain English. Return Markdown with exactly
these section headers:
{sections}

Keep the report concise and under 350 words. Use no more than 2 bullets per section.
Include 2-3 actionable recommendations. Mention data-quality assumptions explicitly.
Do not output <think> tags or any hidden reasoning.

Structured project health data:
{json.dumps(self._build_context(), indent=2)}
"""

    def _validate_generated_output(self, text: str) -> str:
        """Reject incomplete model output so callers receive a complete report."""
        text = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL | re.IGNORECASE).strip()
        first_section = text.find("### Executive Summary")
        if first_section < 0:
            raise ValueError("Model output did not start with the required report")
        text = text[first_section:].strip()
        prompt_artifacts = (
            "Structured project health data:",
            "Drafting - Section by Section",
            "Analyze the Data",
            "Constraint:",
            "Do not change the RAG status",
            "Do not output <think>",
        )
        if any(artifact in text for artifact in prompt_artifacts):
            raise ValueError("Model output contains prompt or reasoning artifacts")
        if not text.strip():
            raise ValueError("Model returned an empty explanation")
        missing = [section for section in self.REQUIRED_SECTIONS if section not in text]
        if missing:
            raise ValueError(f"Model output is missing sections: {', '.join(missing)}")
        return text.strip()

    def _generate_local_reasoning(self) -> str:
        """Generate reasoning with the cached local GGUF model."""
        config = LocalModelConfig.from_environment()
        model = get_local_model(config)
        
        if type(model).__name__ == "Llama":
            system_text = "You write concise, factual PMO health reports. Follow the requested headings exactly."
            user_text = self._build_model_prompt()
            prompt = (
                f"<|im_start|>system\n{system_text}<|im_end|>\n"
                f"<|im_start|>user\n{user_text}<|im_end|>\n"
                f"<|im_start|>assistant\n### Executive Summary\n"
            )
            response = model.create_completion(
                prompt=prompt,
                temperature=0.0,
                max_tokens=config.max_tokens,
                stop=["<|im_end|>"]
            )
            text = "### Executive Summary\n" + response["choices"][0]["text"]
        else:
            response = model.create_chat_completion(
                messages=[
                    {
                        "role": "system",
                        "content": "You write concise, factual PMO health reports. Follow the requested headings exactly.",
                    },
                    {"role": "user", "content": self._build_model_prompt()},
                ],
                temperature=0.0,
                top_p=1.0,
                seed=42,
                max_tokens=config.max_tokens,
            )
            text = response["choices"][0]["message"]["content"]
            
        return self._validate_generated_output(text)

    def _generate_rule_based_reasoning(self) -> str:
        """Deterministic, comprehensive PMO template-based explanation."""
        m = self.metrics
        rag = m["rag_status"]
        
        # Determine schedule description
        slippage_val = m["schedule_slippage"] * 100
        if slippage_val > 5:
            slippage_desc = f"schedule slippage of {slippage_val:.1f}%"
        elif slippage_val < -5:
            slippage_desc = f"tracking {-slippage_val:.1f}% ahead of schedule"
        else:
            slippage_desc = "tracking inline with schedule timeline"

        # Check for missing data
        data_assumptions = []
        # If budget burn / SPI is neutral
        if m["time_elapsed_pct"] == 0:
            data_assumptions.append("Budget burn (SPI) is neutral and not computed because the project start date is today or in the future.")
        else:
            data_assumptions.append("Project budget burn is assumed to scale linearly with elapsed schedule time.")
            
        dq_notes = m.get("data_quality_notes", [])
        for note in dq_notes:
            if "missing" in note.lower():
                data_assumptions.append(note)
            if "could not be parsed" in note.lower():
                data_assumptions.append(note)

        summary = ""
        
        if rag == "Green":
            summary += f"### Executive Summary\n"
            summary += f"The project **{m['project_name']}** is in a **HEALTHY (GREEN)** state. "
            summary += f"The project is currently in the **{m['project_stage']}** stage and is managed by **{m['project_manager']}**. "
            summary += f"Overall completion stands at **{m['pct_complete']*100:.1f}%** against **{m['time_elapsed_pct']*100:.1f}%** elapsed schedule, {slippage_desc}.\n\n"
            
            summary += f"### Why This Status Was Assigned\n"
            summary += f"The project has achieved its major delivery milestones on or ahead of time. "
            summary += f"There are **{m['overdue_milestones_count']}** overdue milestones and **{m['blockers_count']}** active blockers. "
            summary += f"The Schedule Performance Index (SPI) is **{m['spi']:.2f}**, showing solid delivery efficiency.\n\n"
            
            summary += f"### Top Risk Drivers\n"
            if m["overdue_tasks_count"] > 0:
                summary += f"While overall health is Green, there are **{m['overdue_tasks_count']}** overdue subtasks that must be monitored closely to prevent critical path impact. "
            else:
                summary += "No critical risk drivers or schedule blockages are currently active. "
            
            if m["blockers_count"] > 0:
                summary += f"The following minor items are flagged for monitoring: **{', '.join([b['task_name'] for b in m['blockers'][:2]])}**."
            summary += "\n\n"
            
            summary += f"### Recommended Actions\n"
            summary += f"1. **Maintain Cadence**: Continue normal weekly tracking and client reviews.\n"
            summary += f"2. **Address Minor Overdues**: Prompt task owners to update and close out the **{m['overdue_tasks_count']}** overdue subtasks.\n"
            
        elif rag == "Amber":
            summary += f"### Executive Summary\n"
            summary += f"The project **{m['project_name']}** is in an **AMBER (MODERATE RISK)** state. "
            summary += f"It is currently in the **{m['project_stage']}** stage, managed by **{m['project_manager']}**. "
            summary += f"Progress stands at **{m['pct_complete']*100:.1f}%** complete against **{m['time_elapsed_pct']*100:.1f}%** elapsed timeline.\n\n"
            
            summary += f"### Why This Status Was Assigned\n"
            summary += f"The project health is experiencing moderate friction, with **{m['overdue_milestones_count']}** overdue milestones "
            summary += f"and **{m['blockers_count']}** active blockers. "
            if slippage_val > 5:
                summary += f"Schedule slippage of {slippage_val:.1f}% is pulling the project behind schedule. "
            else:
                summary += f"Although the project is mathematically tracking ahead of schedule, delays on specific delivery milestones or dependencies prevent a Green classification. "
            summary += f"The current SPI is **{m['spi']:.2f}**.\n\n"
            
            summary += f"### Top Risk Drivers\n"
            if len(m["blockers"]) > 0:
                summary += f"The main concerns are active blockers on tasks: **{', '.join([b['task_name'] for b in m['blockers'][:3]])}**. "
            if len(m["overdue_milestones"]) > 0:
                summary += f"Overdue milestones include: **{', '.join([ms['task_name'] for ms in m['overdue_milestones'][:2]])}**."
            summary += "\n\n"
            
            summary += f"### Recommended Actions\n"
            summary += f"1. **Blocker Resolution**: Set up a working session to resolve technical blocks on **{', '.join([b['task_name'] for b in m['blockers'][:2]]) if m['blockers'] else 'flagged tasks'}**.\n"
            summary += f"2. **Client Alignment**: Align with the client project team regarding pending inputs and design sign-offs.\n"
            summary += f"3. **Resource Reallocation**: Adjust task ownership or inject support to complete overdue milestones.\n"
            
        else: # Red
            summary += f"### Executive Summary\n"
            summary += f"The project **{m['project_name']}** is in a **RED (CRITICAL STATUS)** state. "
            summary += f"It is currently in the **{m['project_stage']}** stage, managed by **{m['project_manager']}**. "
            summary += f"Overall completion stands at **{m['pct_complete']*100:.1f}%** against **{m['time_elapsed_pct']*100:.1f}%** elapsed schedule ({slippage_desc}).\n\n"
            
            summary += f"### Why This Status Was Assigned\n"
            summary += f"Immediate leadership intervention is required due to severe delivery bottlenecks: "
            summary += f"there are **{m['overdue_milestones_count']}** overdue milestones and **{m['blockers_count']}** active blockers. "
            summary += f"The project SPI is **{m['spi']:.2f}**, representing an inefficient progress-to-timeline burn rate.\n\n"
            
            summary += f"### Top Risk Drivers\n"
            if len(m["overdue_milestones"]) > 0:
                summary += f"Overdue milestones include: **{', '.join([ms['task_name'] for ms in m['overdue_milestones'][:3]])}**. "
            if len(m["blockers"]) > 0:
                summary += f"Key blockers have been flagged on tasks like: **{', '.join([b['task_name'] for b in m['blockers'][:3]])}**."
                blocker_comments = []
                for b in m["blockers"]:
                    if b["comments"]:
                        blocker_comments.append(f"\"{b['comments'][0]['comment']}\"")
                if blocker_comments:
                    summary += f" Notes indicate: {'; '.join(blocker_comments[:2])}."
            summary += "\n\n"
            
            summary += f"### Recommended Actions\n"
            summary += f"1. **Schedule Client Alignment Call**: Request an immediate executive escalation call with the client project sponsor to address critical blocks.\n"
            summary += f"2. **Resource Injection**: Reallocate senior developers and integration architects to configuration and mapping work.\n"
            summary += f"3. **Re-baselining**: If client-side dependencies cannot be unblocked, initiate a formal change request to adjust the project timeline."
            summary += "\n\n"

        # Attach Missing Data & Assumptions Section
        if data_assumptions:
            summary += f"### Data Quality & Assumptions\n"
            for assumption in data_assumptions:
                summary += f"- {assumption}\n"
                
        return summary

    def _generate_llm_reasoning(self, api_key: str) -> str:
        """Calls Gemini API to generate polished executive-grade reports."""
        model_name = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={api_key}"
        headers = {"Content-Type": "application/json"}

        prompt = self._build_model_prompt()
        payload = {
            "contents": [{
                "parts": [{"text": prompt}]
            }]
        }
        
        response = requests.post(url, headers=headers, json=payload, timeout=15)
        if response.status_code == 200:
            res_data = response.json()
            text = res_data['candidates'][0]['content']['parts'][0]['text']
            return self._validate_generated_output(text)
        else:
            raise ValueError(f"Gemini API status code {response.status_code}: {response.text}")
