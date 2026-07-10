import os
from pathlib import Path
import sys
from types import ModuleType
import tempfile
import unittest
from unittest.mock import patch

from src.explainer import ProjectHealthExplainer
from src.local_model import LocalModelConfig, _load_local_model, is_model_cached


VALID_MODEL_OUTPUT = """### Executive Summary
The project is progressing according to the supplied metrics.

### Why This Status Was Assigned
The calculated status is authoritative.

### Top Risk Drivers
The supplied blockers and milestones are the primary risks.

### Recommended Actions
1. Review the highest-priority blocker.
2. Confirm milestone owners and dates.

### Data Quality & Assumptions
The report uses the supplied schedule data and documented assumptions.
"""


def sample_metrics():
    return {
        "project_name": "Test Project",
        "project_manager": "Test Manager",
        "project_stage": "Design",
        "start_date": "2026-07-01",
        "end_date": "2026-08-01",
        "today_date": "2026-07-10",
        "pct_complete": 0.4,
        "time_elapsed_pct": 0.3,
        "schedule_slippage": -0.1,
        "spi": 1.33,
        "status_counts": {"Completed": 1},
        "health_counts": {"Green": 1},
        "overdue_tasks_count": 0,
        "overdue_milestones_count": 0,
        "blockers_count": 0,
        "overdue_milestones": [],
        "blockers": [],
        "data_quality_notes": [],
        "rag_status": "Green",
    }


class FakeLocalModel:
    def create_chat_completion(self, **kwargs):
        return {"choices": [{"message": {"content": VALID_MODEL_OUTPUT}}]}


class TestProjectHealthExplainer(unittest.TestCase):
    def setUp(self):
        self.metrics = sample_metrics()
        self.explainer = ProjectHealthExplainer(self.metrics)

    def test_local_provider_is_default_and_preserves_rag(self):
        with patch.dict(os.environ, {}, clear=True), patch(
            "src.explainer.get_local_model", return_value=FakeLocalModel()
        ):
            output = self.explainer.generate_explanation()

        self.assertEqual(self.metrics["rag_status"], "Green")
        self.assertIn("### Executive Summary", output)

    def test_rules_provider_skips_model(self):
        with patch.dict(os.environ, {"LLM_PROVIDER": "rules"}), patch(
            "src.explainer.get_local_model"
        ) as local_loader:
            output = self.explainer.generate_explanation()

        local_loader.assert_not_called()
        self.assertIn("### Recommended Actions", output)

    def test_gemini_provider_is_explicit(self):
        with patch.dict(
            os.environ,
            {"LLM_PROVIDER": "gemini", "GEMINI_API_KEY": "test-key"},
        ), patch.object(
            self.explainer, "_generate_llm_reasoning", return_value="gemini output"
        ) as gemini:
            output = self.explainer.generate_explanation()

        gemini.assert_called_once_with("test-key")
        self.assertEqual(output, "gemini output")

    def test_local_model_failure_falls_back_to_rules(self):
        with patch.dict(os.environ, {"LLM_PROVIDER": "local"}), patch(
            "src.explainer.get_local_model", side_effect=RuntimeError("offline")
        ):
            output = self.explainer.generate_explanation()

        self.assertIn("### Data Quality & Assumptions", output)

    def test_invalid_model_output_falls_back_to_rules(self):
        invalid_model = FakeLocalModel()
        invalid_model.create_chat_completion = lambda **kwargs: {
            "choices": [{"message": {"content": "incomplete"}}]
        }
        with patch.dict(os.environ, {"LLM_PROVIDER": "local"}), patch(
            "src.explainer.get_local_model", return_value=invalid_model
        ):
            output = self.explainer.generate_explanation()

        self.assertIn("### Top Risk Drivers", output)


class TestLocalModelLoader(unittest.TestCase):
    def test_first_load_downloads_once_and_reuses_cache(self):
        fake_huggingface = ModuleType("huggingface_hub")
        fake_llama_cpp = ModuleType("llama_cpp")
        download_calls = []

        class FakeLlama:
            def __init__(self, **kwargs):
                self.kwargs = kwargs

        def fake_download(**kwargs):
            download_calls.append(kwargs)
            path = Path(kwargs["local_dir"]) / kwargs["filename"]
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(b"test model")
            return str(path)

        fake_huggingface.hf_hub_download = fake_download
        fake_llama_cpp.Llama = FakeLlama

        with tempfile.TemporaryDirectory() as temp_dir, patch.dict(
            sys.modules,
            {"huggingface_hub": fake_huggingface, "llama_cpp": fake_llama_cpp},
        ):
            config = LocalModelConfig(model_dir=Path(temp_dir))
            _load_local_model.cache_clear()
            first = _load_local_model(config)
            second = _load_local_model(config)

            self.assertIs(first, second)
            self.assertEqual(len(download_calls), 1)
            self.assertTrue(is_model_cached(config))
            _load_local_model.cache_clear()
            _load_local_model(config)
            self.assertEqual(len(download_calls), 1)
            _load_local_model.cache_clear()


if __name__ == "__main__":
    unittest.main()
