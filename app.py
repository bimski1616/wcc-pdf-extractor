
from __future__ import annotations

import streamlit as st
import pandas as pd

from extractor import extract_one_pdf, build_excel


st.set_page_config(
    page_title="WCC PDF Extractor",
    page_icon="📄",
    layout="wide",
)

st.title("📄 WCC PDF Extractor")
st.caption(
    "Batch extraction of WCC information and complete Approval Log history."
)

st.info(
    "Upload digital/text-based PDF files. The app extracts WCC information, "
    "reads all Approval Log entries, and calculates the latest approval "
    "and latest activity automatically."
)


# ==========================================
# 1. UPLOAD PDF
# ==========================================

uploaded_files = st.file_uploader(
    "Upload PDF files",
    type=["pdf"],
    accept_multiple_files=True,
)


# ==========================================
# 2. EXTRACTION PROCESS
# ==========================================

if uploaded_files:

    st.write(f"**Total PDF uploaded:** {len(uploaded_files)}")

    if st.button(
        "🚀 Extract Data",
        type="primary",
        use_container_width=True,
    ):

        summary_rows = []
        history_rows = []
        debug_rows = []

        progress = st.progress(0)
        status = st.empty()

        for index, uploaded_file in enumerate(uploaded_files, start=1):

            status.write(
                f"Processing {index}/{len(uploaded_files)}: "
                f"{uploaded_file.name}"
            )

            try:
                summary, history, raw_text = extract_one_pdf(
                    uploaded_file.name,
                    uploaded_file.getvalue(),
                )

                summary_rows.append(summary)
                history_rows.extend(history)

                debug_rows.append({
                    "File Name": uploaded_file.name,
                    "Extracted Text": raw_text,
                })

            except Exception as exc:

                summary_rows.append({
                    "File Name": uploaded_file.name,
                    "Extraction Status": (
                        f"Failed - {type(exc).__name__}: {exc}"
                    ),
                })

                debug_rows.append({
                    "File Name": uploaded_file.name,
                    "Extracted Text": "",
                })

            progress.progress(index / len(uploaded_files))

        st.session_state["wcc_summary"] = summary_rows
        st.session_state["wcc_history"] = history_rows
        st.session_state["wcc_debug"] = debug_rows

        status.success("Extraction process completed.")


# ==========================================
# 3. DISPLAY EXTRACTION RESULTS
# ==========================================

if "wcc_summary" in st.session_state:

    summary_df = pd.DataFrame(
        st.session_state["wcc_summary"]
    )

    history_df = pd.DataFrame(
        st.session_state["wcc_history"]
    )

    st.divider()

    st.subheader("📋 WCC Summary")

    st.dataframe(
        summary_df,
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("📝 Approval History")

    if not history_df.empty:

        st.dataframe(
            history_df,
            use_container_width=True,
            hide_index=True,
        )

    else:

        st.warning(
            "No Approval Log entries were detected."
        )


    # ==========================================
    # 4. EXTRACTION SUMMARY
    # ==========================================

    success_count = sum(
        1
        for row in st.session_state["wcc_summary"]
        if row.get("Extraction Status") == "Success"
    )

    failed_count = (
        len(st.session_state["wcc_summary"]) - success_count
    )

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "PDF Processed",
        len(st.session_state["wcc_summary"]),
    )

    col2.metric(
        "Successful",
        success_count,
    )

    col3.metric(
        "Approval Log Rows",
        len(st.session_state["wcc_history"]),
    )


    # ==========================================
    # 5. EXPORT EXCEL
    # ==========================================

    st.divider()

    excel_bytes = build_excel(
        st.session_state["wcc_summary"],
        st.session_state["wcc_history"],
    )

    st.download_button(
        "⬇️ Download Excel",
        data=excel_bytes,
        file_name="WCC_Extraction.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True,
    )


    # ==========================================
    # 6. DEBUG SECTION
    # ==========================================

    with st.expander("🔎 Debug: Extracted Raw PDF Text"):

        st.caption(
            "Gunakan bagian ini untuk mengecek hasil pembacaan "
            "teks mentah dari setiap PDF."
        )

        for idx, item in enumerate(
            st.session_state["wcc_debug"]
        ):

            st.markdown(
                f"**{item['File Name']}**"
            )

            st.text_area(
                "Extracted Text",
                item["Extracted Text"] or (
                    "(No selectable text found)"
                ),
                height=180,
                key=f"debug_{idx}_{item['File Name']}",
            )


# ==========================================
# 7. INITIAL INSTRUCTIONS
# ==========================================

else:

    st.markdown("### Cara penggunaan")

    st.markdown(
        """
        1. Upload satu atau banyak PDF.
        2. Klik **Extract Data**.
        3. Periksa hasil WCC Summary dan Approval History.
        4. Download satu file Excel berisi dua sheet.
        """
    )
