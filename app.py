from __future__ import annotations

import hashlib

import streamlit as st

from runner import RunnerError
from web_runner import UploadedGeneration, generate_uploaded_bom


APP_VERSION = "streamlit-v4-no-answer-1"

st.set_page_config(
    page_title="BOM 自動產生工具",
    page_icon="📋",
    layout="wide",
)

st.title("BOM 自動產生工具")
st.caption(
    "上傳基準 BOM 與差異表，系統會依內容自動辨識系列並產生新 BOM"
)

with st.sidebar:
    st.header("差異設定")
    variant_choice = st.selectbox(
        "差異表機種編號",
        ["自動判斷", *[str(number) for number in range(1, 16)]],
        index=0,
        help=(
            "建議保持自動判斷。目前 RSP-3000 會使用 3，"
            "HLG-320H 會使用 2。"
        ),
    )
    selected_variant = (
        None if variant_choice == "自動判斷" else int(variant_choice)
    )

    st.divider()

    st.caption("檔名可以不同；系統會讀取 Excel 內容與欄位判斷。")

left, right = st.columns(2)
with left:
    baseline_file = st.file_uploader(
        "1. 上傳基準 BOM（舊機種）",
        type=["xlsx"],
        key="baseline",
    )
with right:
    difference_file = st.file_uploader(
        "2. 上傳差異表",
        type=["xlsx"],
        key="difference",
    )

if baseline_file and difference_file:
    st.success(f"已選擇：{baseline_file.name}、{difference_file.name}")
else:
    st.info("請先上傳兩份 Excel 檔案，再按下產生按鈕。")

current_cache_key: str | None = None
if baseline_file and difference_file:
    current_cache_key = hashlib.sha256(
        APP_VERSION.encode("ascii")
        + baseline_file.getvalue()
        + difference_file.getvalue()
        + str(selected_variant).encode("ascii")
    ).hexdigest()

generate_clicked = st.button(
    "開始產生 BOM",
    type="primary",
    use_container_width=True,
    disabled=not (baseline_file and difference_file),
)

if generate_clicked:
    progress_bar = st.progress(0, text="正在準備檔案……")

    def update_progress(current: int, total: int, message: str) -> None:
        progress_bar.progress(min(current / total, 1.0), text=message)

    try:
        result = generate_uploaded_bom(
            baseline_file.getvalue(),
            difference_file.getvalue(),
            variant=selected_variant,
            progress=update_progress,
        )
        progress_bar.progress(1.0, text="BOM 已產生完成。")
        st.session_state["bom_v4_result"] = result
        st.session_state["bom_v4_cache_key"] = current_cache_key
    except RunnerError as exc:
        st.session_state.pop("bom_v4_result", None)
        st.session_state.pop("bom_v4_cache_key", None)
        progress_bar.empty()
        st.error(str(exc))
    except Exception as exc:
        st.session_state.pop("bom_v4_result", None)
        st.session_state.pop("bom_v4_cache_key", None)
        progress_bar.empty()
        st.error(f"產生失敗：{exc}")

result: UploadedGeneration | None = (
    st.session_state.get("bom_v4_result")
    if st.session_state.get("bom_v4_cache_key") == current_cache_key
    else None
)

if result:
    st.subheader("產生結果")
    metric_family, metric_variant, metric_sheets = st.columns(3)
    metric_family.metric("自動辨識系列", result.family)
    metric_variant.metric("差異表機種編號", result.variant)
    metric_sheets.metric("工作表數", len(result.sheets))

    summary = [
        {
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
        st.success("單量、重複位置與系列規則檢查均已通過。")

    tabs = st.tabs([sheet.sheet_name for sheet in result.sheets])
    for tab, sheet in zip(tabs, result.sheets):
        with tab:
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
            st.dataframe(
                preview,
                use_container_width=True,
                hide_index=True,
                height=460,
            )

st.caption(
    "若上傳尚未支援的新系列，系統會停止並顯示無法辨識，"
    "不會將現有系列規則硬套到新資料。"
)
