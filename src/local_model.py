"""Lazy loading and inference configuration for the local Qwen model."""

from dataclasses import dataclass
from functools import lru_cache
import os
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class LocalModelConfig:
    """Runtime settings for the local GGUF model."""

    repo: str = "bartowski/Qwen_Qwen3.5-0.8B-GGUF"
    filename: str = "Qwen_Qwen3.5-0.8B-Q4_K_M.gguf"
    model_dir: Path = Path(".models")
    revision: str = "main"
    context_size: int = 8192
    max_tokens: int = 800

    @classmethod
    def from_environment(cls) -> "LocalModelConfig":
        """Load model settings from environment variables with safe defaults."""
        return cls(
            repo=os.getenv("LOCAL_MODEL_REPO", cls.repo),
            filename=os.getenv("LOCAL_MODEL_FILE", cls.filename),
            model_dir=Path(os.getenv("LOCAL_MODEL_DIR", str(cls.model_dir))),
            revision=os.getenv("LOCAL_MODEL_REVISION", cls.revision),
            context_size=_read_positive_int("LOCAL_MODEL_CONTEXT", cls.context_size),
            max_tokens=_read_positive_int("LOCAL_MODEL_MAX_TOKENS", cls.max_tokens),
        )

    @property
    def local_path(self) -> Path:
        """Return the expected path for the cached model file."""
        return self.model_dir / self.filename


def _read_positive_int(name: str, default: int) -> int:
    try:
        value = int(os.getenv(name, str(default)))
        return value if value > 0 else default
    except (TypeError, ValueError):
        return default


@lru_cache(maxsize=1)
def _load_local_model(config: LocalModelConfig) -> Any:
    """Download, cache, and load the model exactly once per process."""
    from huggingface_hub import hf_hub_download
    from llama_cpp import Llama

    config.model_dir.mkdir(parents=True, exist_ok=True)
    if config.local_path.exists():
        model_path = str(config.local_path)
    else:
        model_path = hf_hub_download(
            repo_id=config.repo,
            filename=config.filename,
            revision=config.revision,
            local_dir=str(config.model_dir),
        )
    return Llama(
        model_path=model_path,
        chat_format="chatml",
        n_ctx=config.context_size,
        n_threads=max(1, (os.cpu_count() or 2) - 1),
        n_gpu_layers=0,
        verbose=False,
    )


def get_local_model(config: LocalModelConfig | None = None) -> Any:
    """Return the cached local model, downloading it on first use."""
    return _load_local_model(config or LocalModelConfig.from_environment())


def is_model_cached(config: LocalModelConfig | None = None) -> bool:
    """Report whether the configured model file is already cached locally."""
    return (config or LocalModelConfig.from_environment()).local_path.exists()
