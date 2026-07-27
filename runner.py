from __future__ import annotations

import os
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook

from bom_engine import BOMError as RSPBOMError
from bom_engine import GenerationResult as RSPGenerationResult
from bom_engine import generate_bom
from hlg_engine import BOMError as HLGBOMError
from hlg_engine import GenerationResult as HLGGenerationResult
from hlg_engine import default_variant_from_profile, generate_from_xlsx


ProgressCallback = Callable[[int, int, str], None]


class RunnerError(ValueError):
    """Raised when inputs, family detection, or output paths are invalid."""


@dataclass(frozen=True)
class SavedBOM:
    output_path: Path
    family: str
    variant: int
    result: RSPGenerationResult | HLGGenerationResult


def _report(
    callback: ProgressCallback | None,
    current: int,
    total: int,
    message: str,
) -> None:
    if callback:
        callback(current, total, message)


def _validate_xlsx(path: str | Path, label: str) -> Path:
    candidate = Path(path).expanduser()
    if not candidate.is_file():
        raise RunnerError(f"{label}不存在：{candidate}")
    if candidate.suffix.lower() != ".xlsx":
        raise RunnerError(f"{label}必須是 .xlsx 格式。")
    return candidate.resolve()


def _output_path(path: str | Path) -> Path:
    if not str(path).strip():
        raise RunnerError("輸出檔案不可空白。")
    candidate = Path(path).expanduser()
    if not candidate.suffix:
        candidate = candidate.with_suffix(".xlsx")
    if candidate.suffix.lower() != ".xlsx":
        raise RunnerError("輸出檔案必須是 .xlsx 格式。")
    return candidate.resolve()


def _save_atomically(data: bytes, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb",
            prefix=f".{output_path.stem}-",
            suffix=".xlsx",
            dir=output_path.parent,
            delete=False,
        ) as temporary_file:
            temporary_file.write(data)
            temporary_path = Path(temporary_file.name)
        os.replace(temporary_path, output_path)
    finally:
        if temporary_path and temporary_path.exists():
            temporary_path.unlink()


def detect_family(source: str | Path) -> str:
    """Detect the BOM profile from title/header content instead of the filename."""
    wb = load_workbook(source, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    values: list[str] = []
    for row in ws.iter_rows(
        min_row=1,
        max_row=min(ws.max_row, 25),
        max_col=min(ws.max_column, 15),
        values_only=True,
    ):
        values.extend(str(value).upper() for value in row if value is not None)
    text = " ".join(values)
    if "HLG-320H" in text or "9HLG-" in text:
        return "HLG-320H（單工作表）"
    if "RSP-3000" in text or "9RSP-" in text:
        return "RSP-3000（A/B/C/E 四工作表）"
    raise RunnerError(
        "無法從基準 BOM 標題或半成品辨識系列。目前支援 RSP-3000 與 HLG-320H。"
    )


def run_generation(
    baseline_path: str | Path,
    difference_path: str | Path,
    output_path: str | Path,
    *,
    variant: int | None = None,
    progress: ProgressCallback | None = None,
) -> SavedBOM:
    total_steps = 5
    _report(progress, 1, total_steps, "正在檢查基準 BOM、差異表與輸出位置……")
    baseline = _validate_xlsx(baseline_path, "基準 BOM")
    difference = _validate_xlsx(difference_path, "差異表")
    output = _output_path(output_path)
    if output in {baseline, difference}:
        raise RunnerError("輸出檔案不可覆蓋任何輸入檔案。")

    _report(progress, 2, total_steps, "正在依標題、半成品與欄位自動辨識系列……")
    family = detect_family(baseline)
    profile_path = Path(__file__).resolve().parent / "profiles" / "hlg_320h.json"

    selected_variant = variant
    if selected_variant is None:
        if family.startswith("HLG-320H"):
            selected_variant = default_variant_from_profile(profile_path)
        elif family.startswith("RSP-3000"):
            selected_variant = 3
        else:
            raise RunnerError("無法自動判斷機種編號，請手動選擇。")

    _report(
        progress,
        3,
        total_steps,
        f"已辨識：{family}；正在套用差異表機種編號 {selected_variant}……",
    )
    try:
        if family.startswith("HLG-320H"):
            result: RSPGenerationResult | HLGGenerationResult = generate_from_xlsx(
                baseline,
                difference,
                variant=int(selected_variant),
                profile_source=profile_path,
            )
        else:
            if int(selected_variant) != 3:
                raise RunnerError(
                    "RSP-3000 目前已驗證的機種編號為 3；"
                    f"編號 {selected_variant} 尚未建立輸出機種設定。"
                )
            result = generate_bom(
                baseline,
                difference,
                work_order="W2603D333A",
                target_model="AB-RSP-3000-012-R19A",
                variant=int(selected_variant),
                apply_calibration=True,
            )
    except (RSPBOMError, HLGBOMError) as exc:
        raise RunnerError(str(exc)) from exc

    _report(progress, 4, total_steps, "正在合併相同品號／規格的位置、檢查重複與單量……")
    _save_atomically(result.workbook_bytes, output)
    _report(progress, 5, total_steps, "BOM 已產生完成。")
    return SavedBOM(output, family, int(selected_variant), result)


__all__ = ["RunnerError", "SavedBOM", "detect_family", "run_generation"]
