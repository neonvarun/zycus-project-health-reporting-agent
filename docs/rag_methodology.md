# Project Health Reporting: RAG Methodology Framework

This document outlines the framework used by the automated **Project Health Reporting Agent** to calculate the Red/Amber/Green (RAG) status of delivery projects within the Professional Services portfolio.

---

## 1. Objectives & Architectural Principles

The Project Health Reporting Agent adheres to the following core tenets:
1. **Deterministic Core**: RAG calculations are strictly rule-based, mathematical, and reproducible. A project's color classification is never decided by an AI model, eliminating hallucination risks.
2. **AI-Enhanced Reasoning**: A local Qwen3.5-0.8B model is used exclusively to *summarize* and *polish* plain-English write-ups from quantitative indices. Gemini remains an optional explicit provider.
3. **Data Quality Resilience**: Large implementation plans often suffer from missing cells, unparseable formulas, or incomplete sheets. The agent calculates a metric-level "Completeness Indicator" and falls back to baseline/parent schedules rather than crashing.

The local explanation provider downloads `Qwen3.5-0.8B` in Q4_K_M GGUF format on first use, caches it under `.models/`, and falls back to the deterministic rules explanation if local inference is unavailable. The model is never allowed to change the calculated RAG status.

---

## 2. Core Health Signals & Weighted Framework

The agent aggregates six primary signals extracted from the project plans using a **weighted score of 0 to 100**:

$$\text{Project Health Score} = (0.30 \times S) + (0.25 \times M) + (0.20 \times B) + (0.10 \times C) + (0.10 \times E) + (0.05 \times D)$$

Where:
* **$S$ = Schedule Health & Variance (30%)**
* **$M$ = Progress & Milestone Completion (25%)**
* **$B$ = Risk & Blockers (20%)**
* **$C$ = Budget Burn (10%)**
* **$E$ = Stakeholder Sentiment (10%)**
* **$D$ = Data Quality & Completeness (5%)**

---

### Category A: Schedule Health & Variance (Weight: 30%)
Schedule slippage measures the variance between the elapsed duration of the project and its physical percentage complete.
* **Formulas**:
  $$\text{Time Elapsed \%} = \frac{\text{Today's Date} - \text{Project Start Date}}{\text{Project End Date} - \text{Project Start Date}}$$
  $$\text{Schedule Slippage} = \text{Time Elapsed \%} - \text{Overall \% Complete}$$
* **Slippage Scoring**:
  * **Score 100 (Green / Low Risk)**: Slippage $\le 5\%$ (or project is ahead of schedule)
  * **Score 50 (Amber / Medium Risk)**: Slippage $5\% - 15\%$
  * **Score 0 (Red / High Risk)**: Slippage $> 15\%$

### Category B: Progress & Milestone Completion (Weight: 25%)
Milestones represent critical gateways. Milestone health assesses whether key milestones are delayed.
* **Milestone Scoring**:
  * **Score 100 (Green)**: 0 overdue milestones.
  * **Score 80 (Low Amber)**: 1–2 overdue milestones, or an overdue milestone ratio $\le 5\%$.
  * **Score 50 (Amber)**: 3–5 overdue milestones, or an overdue milestone ratio $\le 15\%$.
  * **Score 30 (High Amber)**: More than 5 overdue milestones and a ratio above 15%.
  * *An overdue milestone is defined as having its End Date in the past relative to Today's Date and not marked 'Completed'.*

### Category C: Risk / Blockers / On-Hold Tasks (Weight: 20%)
Blockers directly halt downstream activities.
* **Scoring**:
  * **Score 100 (Green)**: 0 active blockers.
  * **Score 80 (Low Amber)**: Up to 3 blockers or blockers representing up to 3% of tasks.
  * **Score 50 (Amber)**: Up to 10 blockers or blockers representing up to 6% of tasks.
  * **Score 30 (High Amber)**: More than 10 blockers and more than 6% of tasks.

### Category D: Budget Burn & Efficiency (Weight: 10%)
Since explicit financial ledger columns are not present in standard project schedules, the agent infers budget efficiency using the **Schedule Performance Index (SPI)**:
* **Formula**:
  $$\text{SPI} = \frac{\text{Overall \% Complete}}{\text{Time Elapsed \%}}$$
* **Scoring**:
  * **Score 100 (Green)**: $\text{SPI} \ge 0.95$ (Budget is burning inline with or slower than progress).
  * **Score 50 (Amber)**: $\text{SPI} = 0.80 - 0.94$.
  * **Score 0 (Red)**: $\text{SPI} < 0.80$ (Severe budget overrun risk).
* **Missing Budget Data Rule**: Standard schedules do not contain a financial ledger, so SPI is used as a schedule-based budget proxy. If SPI cannot be computed (for example, elapsed time is 0), the score defaults to neutral **100** so the project is not penalized.

### Category E: Stakeholder Sentiment & Comments (Weight: 10%)
Sentiment is parsed from the qualitative notes and comments logged in the workspace.
* **Scoring**:
  * **Score 100 (Green)**: No comments contain negative keywords, or there are no comments.
  * **Score 80 (Light Amber)**: 1–2 comments contain negative keywords.
  * **Score 65 (Amber)**: 3–5 comments contain negative keywords.
  * **Score 50 (Amber)**: 6–10 comments contain negative keywords.
  * **Score 30 (High Amber)**: More than 10 comments contain negative keywords.
  * *Negative keywords include: delay, delayed, impacted, pending, blocker, issue, risk, dependency, waiting, need, not done, missing, failed, problem, concern, on hold, escalate, escalation, stuck.*

### Category F: Data Quality & Completeness (Weight: 5%)
Assesses schedule completeness:
* **Scoring (Starts at 100, deductions applied)**:
  * Missing Summary sheet: $-20$ points.
  * Missing Comments sheet: $-10$ points.
  * Unparseable Excel dates/serial dates: $-5$ points per instance.
  * Missing core plan sheet columns (like Task Name, Status): $-15$ points.

---

## 3. Overall RAG Mapping & Overrides

The weighted score maps directly to the final RAG Status:

| Overall Score | Final RAG Status | Portfolio Action Required |
| :--- | :---: | :--- |
| **80% to 100%** | **Green** | Project is healthy. Maintain normal tracking. |
| **60% to 79%** | **Amber** | Monitoring required. Review blockers in weekly status. PM to propose mitigation plan. |
| **Below 60%** | **Red** | Executive intervention required. Immediate client alignment and resource reallocation. |

### Governance Overrides
To ensure absolute alignment with PMO director logic, three overrides are enforced:
1. **Summary Status Sync**: If the plan schedule health is explicitly set to `Red` or `Critical` in the Summary sheet, the project RAG is overridden to `Red` regardless of calculated metrics. If `Yellow` or `Amber` and the score is `Green`, RAG is pushed to `Amber`.
2. **Severe Slippage Override**: If overall schedule slippage is $> 25\%$, the status is overridden to `Red` automatically.
3. **High Blocker Override**: If active blockers are $\ge 15$, represent $\ge 10\%$ of total tasks, and schedule slippage is positive ($>5\%$), the status is overridden to `Red` automatically. (This prevents large projects with negative slippage from being falsely flagged as Red).


---

## 4. Why Deterministic Rules + LLM is Better Than Pure LLM

1. **Repeatability**: Large Language Models (LLMs) are probabilistic and can classify the same schedule data differently across runs. Deterministic rules ensure consistent classifications.
2. **Auditability**: Leadership can trace a "Red" status back to exact math (e.g. 14.7% slippage, 4 blockers).
3. **Cost & Speed**: Evaluator logic runs locally in milliseconds. After the first model download, explanation generation also runs locally without an API key.
