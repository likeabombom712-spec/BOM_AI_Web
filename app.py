from __future__ import annotations

import hashlib

import streamlit as st

from adaptive_engine import AdaptiveAnalysis, AdaptiveBOMError, analyze_inputs
from runner import RunnerError
from web_runner import UploadedGeneration, generate_uploaded_bom


# 僅供快取失效使用；網頁不顯示版本文字。
APP_BUILD = "adaptive-layout-20260809-1"

st.set_page_config(
    page_title="BOM 自動產生工具",
    page_icon="📋",
    layout="wide",
)

st.title("BOM 自動產生工具")
st.caption(
    "上傳基準 BOM 與差異表；系統會自動尋找表頭、工單、機種、製程區、"
    "半成品、品號、規格、數量與位置。"
)

left, right = st.columns(2)
with left:
    baseline_file = st.file_uploader(
        "1. 上傳基準 BOM",
        type=["xlsx"],
        key="baseline",
    )
with right:
    difference_file = st.file_uploader(
        "2. 上傳差異表",
        type=["xlsx"],
        key="difference",
    )

analysis: AdaptiveAnalysis | None = None
analysis_error: str | None = None
if baseline_file and difference_file:
    try:
        analysis = analyze_inputs(
            baseline_file.getvalue(),
            difference_file.getvalue(),
        )
    except (AdaptiveBOMError, ValueError) as exc:
        analysis_error = str(exc)
    except Exception:
        analysis_error = (
            "Excel 內容無法安全解析。請確認兩份檔案未損壞，且差異表含有"
            "製程、位置、品號、規格及機種編號。"
        )

if analysis_error:
    st.error(analysis_error)

selected_targets: list[str] = []
if analysis:
    target_by_key = {target.key: target for target in analysis.targets}
    work_orders = {
        target.work_order for target in analysis.targets
    }
    metric_orders, metric_targets, metric_sheets = st.columns(3)
    metric_orders.metric("辨識到的工單群組", len(work_orders))
    metric_targets.metric("辨識到的機種", len(analysis.targets))
    metric_sheets.metric("全部工作表", analysis.expected_sheet_count)

    st.subheader("自動辨識結果")
    st.dataframe(
        analysis.matrix,
        use_container_width=True,
        hide_index=True,
    )

    def target_label(key: str) -> str:
        target = target_by_key[key]
        order = "" if target.work_order == "AUTO" else f"{target.work_order}｜"
        return (
            f"{order}第{target.model}機種｜{target.model_name}｜"
            f"{'/'.join(target.sections)}版"
        )

    options = list(target_by_key)
    selected_targets = st.multiselect(
        "要產生的工單／機種",
        options=options,
        default=options,
        format_func=target_label,
    )
    selected_count = sum(
        len(target_by_key[key].sections) for key in selected_targets
    )
    st.caption(
        f"已選 {len(selected_targets)} 個工單／機種，將產生 {selected_count} 張工作表。"
    )
    if analysis.warnings:
        with st.expander(f"辨識提醒（{len(analysis.warnings)}）"):
            for warning in analysis.warnings:
                st.write(f"- {warning}")
elif not (baseline_file and difference_file):
    st.caption("請先選擇兩份 .xlsx 檔案。")

current_cache_key: str | None = None
if baseline_file and difference_file and selected_targets:
    current_cache_key = hashlib.sha256(
        APP_BUILD.encode("ascii")
        + baseline_file.getvalue()
        + difference_file.getvalue()
        + "\n".join(selected_targets).encode("utf-8")
    ).hexdigest()

generate_clicked = st.button(
    "開始產生 BOM",
    type="primary",
    use_container_width=True,
    disabled=not (
        baseline_file
        and difference_file
        and selected_targets
        and not analysis_error
    ),
)

if generate_clicked:
    progress_bar = st.progress(0, text="正在準備檔案……")

    def update_progress(current: int, total: int, message: str) -> None:
        progress_bar.progress(min(current / total, 1.0), text=message)

    try:
        result = generate_uploaded_bom(
            baseline_file.getvalue(),
            difference_file.getvalue(),
            selected_targets=selected_targets,
            baseline_name=baseline_file.name,
            progress=update_progress,
        )
        progress_bar.progress(1.0, text="BOM 已產生完成。")
        st.session_state["bom_result"] = result
        st.session_state["bom_cache_key"] = current_cache_key
    except RunnerError as exc:
        st.session_state.pop("bom_result", None)
        st.session_state.pop("bom_cache_key", None)
        progress_bar.empty()
        st.error(str(exc))
    except Exception as exc:
        st.session_state.pop("bom_result", None)
        st.session_state.pop("bom_cache_key", None)
        progress_bar.empty()
        st.error(f"產生失敗：{exc}")

result: UploadedGeneration | None = (
    st.session_state.get("bom_result")
    if st.session_state.get("bom_cache_key") == current_cache_key
    else None
)

if result:
    st.subheader("產生結果")
    result_orders = {sheet.work_order_group for sheet in result.sheets}
    metric_orders, metric_targets, metric_sheets = st.columns(3)
    metric_orders.metric("工單群組", len(result_orders))
    metric_targets.metric("工單／機種", len(result.target_keys))
    metric_sheets.metric("工作表數", len(result.sheets))

    summary = [
        {
            "工單": (
                sheet.work_order_group
                if sheet.work_order_group != "AUTO"
                else "單一系列"
            ),
            "機種": f"第{sheet.model}機種",
            "機種名稱": sheet.model_name,
            "版別": sheet.section,
            "工作表": sheet.sheet_name,
            "半成品": sheet.semi_finished,
            "合併後料件數": len(sheet.items),
            "單量合計": sheet.total_quantity,
        }
        for sheet in result.sheets
    ]
    st.dataframe(summary, use_container_width=True, hide_index=True)

    st.download_button(
        "下載產生完成的 BOM Excel",
        data=result.workbook_bytes,
        file_name=result.download_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    if result.warnings:
        with st.expander(f"檢查訊息（{len(result.warnings)}）"):
            for warning in result.warnings:
                st.write(f"- {warning}")
    else:
        st.success("選料、單量、重複位置及合併檢查均已通過。")

    sheet_names = [sheet.sheet_name for sheet in result.sheets]
    preview_name = st.selectbox("預覽工作表", sheet_names)
    sheet = result.sheets[sheet_names.index(preview_name)]
    preview = [
        {
            "製程段": sheet.process if index == 0 else "",
            "半成品": sheet.semi_finished if index == 0 else "",
            "品號": item.part_no,
            "規格": item.specification,
            "單量": item.quantity,
            "位置": " ".join(item.positions),
        }
        for index, item in enumerate(sheet.items)
    ]
    st.dataframe(preview, use_container_width=True, hide_index=True, height=460)

st.caption(
    "系統只在表頭、工單、機種、製程區與料件資料能互相核對時輸出；"
    "若有衝突會停止，避免產生看似完成但內容錯誤的 BOM。"
)
