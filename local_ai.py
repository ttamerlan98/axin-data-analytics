"""Fully offline Qwen model lifecycle and Excel-audit inference."""

from __future__ import annotations

import threading
import time
from pathlib import Path
from typing import Any

PROJECT_DIR = Path(__file__).resolve().parent
MODELS_DIR = PROJECT_DIR / "models"
MODEL_FILENAME = "Qwen2.5-1.5B-Instruct-Q4_K_M.gguf"
MODEL_REPO = "bartowski/Qwen2.5-1.5B-Instruct-GGUF"
MODEL_URL = (
    f"https://huggingface.co/{MODEL_REPO}/resolve/main/{MODEL_FILENAME}"
)
MODEL_PATH = MODELS_DIR / MODEL_FILENAME
MODEL_SIZE_ESTIMATE = 1_100_000_000
SYSTEM_PROMPT = (
    "Sən yalnız Excel cədvəlləri və maliyyə məlumatlarının auditi üzrə "
    "köməkçisən. Yalnız verilən məlumatlara əsaslan, fərziyyələri açıq bildir, "
    "maliyyə məsləhəti vermə və Excel/maliyyə auditi ilə əlaqəsiz suallara "
    "nəzakətlə cavab verməkdən imtina et. Cavabı Azərbaycan dilində ver."
)

_lock = threading.RLock()
_model: Any | None = None
_worker: threading.Thread | None = None
_status: dict[str, Any] = {
    "state": "not_downloaded",
    "downloaded": False,
    "loaded": False,
    "progress": 0.0,
    "downloaded_bytes": 0,
    "total_bytes": MODEL_SIZE_ESTIMATE,
    "model": MODEL_FILENAME,
    "repo": MODEL_REPO,
    "url": MODEL_URL,
    "message": "Model hələ endirilməyib; endirmə yalnız açıq sorğu ilə başlayır.",
    "error": None,
    "updated_at": time.time(),
}


def _update(**changes: Any) -> None:
    with _lock:
        _status.update(changes, updated_at=time.time())


def model_status() -> dict[str, Any]:
    with _lock:
        status = dict(_status)
        path = MODEL_PATH
        if path.is_file():
            size = path.stat().st_size
            status.update(
                downloaded=True,
                downloaded_bytes=size,
                total_bytes=size,
                progress=100.0,
            )
            if status["state"] == "not_downloaded":
                status["state"] = "downloaded"
                status["message"] = "Model endirilib; yüklənməyə hazırdır."
        elif status["state"] == "downloading" and MODELS_DIR.exists():
            partial_size = sum(
                item.stat().st_size
                for item in MODELS_DIR.rglob("*.incomplete")
                if item.is_file()
            )
            status["downloaded_bytes"] = partial_size
            status["progress"] = round(
                min(99.0, partial_size / MODEL_SIZE_ESTIMATE * 100),
                2,
            )
        return status


def _download() -> None:
    try:
        MODELS_DIR.mkdir(parents=True, exist_ok=True)
        _update(state="downloading", error=None, message="Model Hugging Face-dən endirilir.")
        from huggingface_hub import hf_hub_download

        downloaded_path = Path(
            hf_hub_download(
                repo_id=MODEL_REPO,
                filename=MODEL_FILENAME,
                local_dir=str(MODELS_DIR),
            )
        )
        if downloaded_path.resolve() != MODEL_PATH.resolve():
            downloaded_path.replace(MODEL_PATH)
        size = MODEL_PATH.stat().st_size
        _update(
            state="downloaded",
            downloaded=True,
            downloaded_bytes=size,
            total_bytes=size,
            progress=100.0,
            message="Model endirildi; audit üçün yüklənməyə hazırdır.",
            error=None,
        )
    except Exception as error:
        _update(
            state="error",
            error=f"{type(error).__name__}: {error}",
            message="Model endirilə bilmədi. Şəbəkəni və Hugging Face bağlantısını yoxlayın.",
        )


def start_download() -> dict[str, Any]:
    global _worker
    with _lock:
        if MODEL_PATH.is_file():
            _status.update(
                state="ready" if _model is not None else "downloaded",
                downloaded=True,
                loaded=_model is not None,
                downloaded_bytes=MODEL_PATH.stat().st_size,
                message="Model artıq endirilib.",
                error=None,
                updated_at=time.time(),
            )
            return dict(_status)
        if _worker is not None and _worker.is_alive():
            return dict(_status)
        _status.update(
            state="queued",
            downloaded=False,
            progress=0.0,
            error=None,
            message="Model endirməsi növbəyə alındı.",
            updated_at=time.time(),
        )
        _worker = threading.Thread(
            target=_download,
            name="datalens-model-download",
            daemon=True,
        )
        _worker.start()
        return dict(_status)


def _load_in_background() -> None:
    try:
        _load_model()
    except Exception:
        # The failure details are already reflected in model_status().
        return


def start_loading() -> dict[str, Any]:
    global _worker
    with _lock:
        if _model is not None:
            return model_status()
        if not MODEL_PATH.is_file():
            return start_download()
        if _worker is not None and _worker.is_alive():
            return dict(_status)
        _status.update(
            state="queued",
            downloaded=True,
            error=None,
            message="Modelin yaddaşa yüklənməsi növbəyə alındı.",
            updated_at=time.time(),
        )
        _worker = threading.Thread(
            target=_load_in_background,
            name="datalens-model-load",
            daemon=True,
        )
        _worker.start()
        return dict(_status)


def _load_model() -> Any:
    global _model
    with _lock:
        if _model is not None:
            return _model
        if not MODEL_PATH.is_file():
            start_download()
            raise RuntimeError(
                "Yerli model mövcud deyil. Endirmə başladıldı; tamamlandıqdan sonra yenidən cəhd edin."
            )
        _status.update(
            state="loading",
            downloaded=True,
            message="Yerli model yaddaşa yüklənir.",
            error=None,
            updated_at=time.time(),
        )
        try:
            from llama_cpp import Llama

            _model = Llama(
                model_path=str(MODEL_PATH),
                n_ctx=4096,
                n_threads=4,
                verbose=False,
            )
        except Exception as error:
            _status.update(
                state="error",
                error=f"{type(error).__name__}: {error}",
                message="Yerli model yaddaşa yüklənə bilmədi.",
                updated_at=time.time(),
            )
            raise RuntimeError(
                "Yerli Qwen modeli yüklənə bilmədi; model faylını və llama-cpp-python quraşdırmasını yoxlayın."
            ) from error
        _status.update(
            state="ready",
            downloaded=True,
            loaded=True,
            progress=100.0,
            message="Yerli Qwen modeli audit üçün hazırdır.",
            error=None,
            updated_at=time.time(),
        )
        return _model


def audit_with_local_model(prompt: str) -> str:
    if not prompt.strip():
        raise ValueError("Audit üçün məlumat təqdim edilməyib.")
    if len(prompt) > 24_000:
        prompt = prompt[:24_000] + "\n[Məlumat kontekstinin qalan hissəsi ixtisar edildi.]"
    model = _load_model()
    try:
        response = model.create_chat_completion(
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": prompt},
            ],
            temperature=0.1,
            max_tokens=700,
        )
        answer = response["choices"][0]["message"]["content"]
        if not answer or not answer.strip():
            raise RuntimeError("Model boş audit cavabı qaytardı.")
        return answer.strip()
    except Exception as error:
        _update(
            state="error",
            error=f"{type(error).__name__}: {error}",
            message="Yerli model auditı icra edə bilmədi.",
        )
        raise RuntimeError(f"Yerli AI auditı uğursuz oldu: {error}") from error
