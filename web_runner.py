from __future__ import annotations

import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from runner import RunnerError, run_generation


ProgressCallback = Callable[[int, int, str], None]


@dataclass(frozen=True)
class UploadedGeneration:
    """In-memory result returned to the Streamlit interface."""

    workbook_bytes: bytes
    download_name: str
    family: str
    variant: int
    result: object

    @property
    def sheets(self) -> tuple[object, ...]:
        if self.family.startswith("HLG-320H"):
            return (self.result.sheet,)
        return tuple(self.result.sheets)

    @property
    def warnings(self) -> list[str]:
        return list(self.result.warnings)


def _safe_filename(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*]+', "_", value).strip(" .")
    return value or "產生完成的_BOM"


def _download_name(family: str, result: object) -> str:
    if family.startswith("HLG-320H"):
        return f"{_safe_filename(result.sheet.sheet_name)}_BOM.xlsx"

    first_sheet_name = result.sheets[0].sheet_name
    model_prefix = re.sub(r"-R19A$", "", first_sheet_name, flags=re.IGNORECASE)
    return f"{_safe_filename(model_prefix)}_四工作表_BOM.xlsx"


def generate_uploaded_bom(
    baseline_bytes: bytes,
    difference_bytes: bytes,
    *,
    variant: int | None = None,
    progress: ProgressCallback | None = None,
) -> UploadedGeneration:
    """Run the V4 file-based engine safely against uploaded workbook bytes."""

    if not baseline_bytes:
        raise RunnerError("基準 BOM 檔案是空白的。")
    if not difference_bytes:
        raise RunnerError("差異表檔案是空白的。")

    try:
        with tempfile.TemporaryDirectory(prefix="bom_streamlit_v4_") as directory:
            root = Path(directory)
            baseline_path = root / "baseline.xlsx"
            difference_path = root / "difference.xlsx"
            output_path = root / "generated.xlsx"
            baseline_path.write_bytes(baseline_bytes)
            difference_path.write_bytes(difference_bytes)

            saved = run_generation(
                baseline_path,
                difference_path,
                output_path,
                variant=variant,
                progress=progress,
            )
            workbook_bytes = output_path.read_bytes()
    except RunnerError:
        raise
    except Exception as exc:
        raise RunnerError(
            "Excel 無法讀取或產生失敗；請確認兩份檔案都是未損壞的 .xlsx 活頁簿。"
        ) from exc

    return UploadedGeneration(
        workbook_bytes=workbook_bytes,
        download_name=_download_name(saved.family, saved.result),
        family=saved.family,
        variant=saved.variant,
        result=saved.result,
    )


__all__ = ["UploadedGeneration", "generate_uploaded_bom"]
