from __future__ import annotations

import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable

from runner import RunnerError, run_generation


ProgressCallback = Callable[[int, int, str], None]


@dataclass(frozen=True)
class UploadedGeneration:
    workbook_bytes: bytes
    download_name: str
    family: str
    models: tuple[int, ...]
    target_keys: tuple[str, ...]
    result: object

    @property
    def variant(self) -> int | None:
        return self.models[0] if len(self.models) == 1 else None

    @property
    def sheets(self) -> tuple[object, ...]:
        return tuple(self.result.sheets)

    @property
    def warnings(self) -> list[str]:
        return list(self.result.warnings)

    @property
    def core_regression_checked(self) -> bool:
        return self.result.mode in {"legacy_rsp", "legacy_hlg"}


def _safe_filename(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*]+', "_", value).strip(" .")
    return value or "產生完成的_BOM"


def _download_name(saved: object) -> str:
    result = saved.result
    if len(result.sheets) == 1:
        return _safe_filename(f"{result.sheets[0].sheet_name}.xlsx")
    work_orders = len({sheet.work_order_group for sheet in result.sheets})
    target_count = len(saved.target_keys)
    return _safe_filename(
        f"自動辨識_{work_orders}工單_{target_count}機種_"
        f"共{len(result.sheets)}張_BOM表.xlsx"
    )


def generate_uploaded_bom(
    baseline_bytes: bytes,
    difference_bytes: bytes,
    *,
    selected_targets: Iterable[str] | None = None,
    models: Iterable[int] | None = None,
    variant: int | None = None,
    baseline_name: str = "BOM.xlsx",
    progress: ProgressCallback | None = None,
) -> UploadedGeneration:
    if not baseline_bytes:
        raise RunnerError("基準 BOM 檔案是空白的。")
    if not difference_bytes:
        raise RunnerError("差異表檔案是空白的。")

    try:
        with tempfile.TemporaryDirectory(prefix="bom_streamlit_v52_") as directory:
            root = Path(directory)
            baseline_path = root / "baseline.xlsx"
            difference_path = root / "difference.xlsx"
            output_path = root / "generated.xlsx"
            baseline_path.write_bytes(baseline_bytes)
            difference_path.write_bytes(difference_bytes)
            project_name = Path(baseline_name).stem or "BOM"

            saved = run_generation(
                baseline_path,
                difference_path,
                output_path,
                selected_targets=selected_targets,
                models=models,
                variant=variant,
                project_name=project_name,
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
        download_name=_download_name(saved),
        family=saved.family,
        models=saved.models,
        target_keys=saved.target_keys,
        result=saved.result,
    )


__all__ = ["UploadedGeneration", "generate_uploaded_bom"]
