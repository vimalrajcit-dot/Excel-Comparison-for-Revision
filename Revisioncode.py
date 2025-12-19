import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import tempfile
import os

# =========================
# Page Config
# =========================
st.set_page_config(
    page_title="R0 vs R1 Tag Comparison",
    page_icon="📊",
    layout="wide"
)

# =========================
# UI Header
# =========================
st.markdown(
    """
    <style>
    .title {
        font-size: 42px;
        font-weight: 700;
        text-align: center;
        margin-bottom: 10px;
    }
    .subtitle {
        font-size: 18px;
        text-align: center;
        color: #6c757d;
        margin-bottom: 30px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown('<div class="title">📊 R0 vs R1 Tag Comparison</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Excel diffing, but make it smooth</div>', unsafe_allow_html=True)

# =========================
# File Upload
# =========================
col1, col2 = st.columns(2)

with col1:
    r0_file = st.file_uploader("Upload R0.xlsx", type=["xlsx"])

with col2:
    r1_file = st.file_uploader("Upload R1.xlsx", type=["xlsx"])

# =========================
# Processing
# =========================
if r0_file and r1_file:
    with st.spinner("🔍 Comparing files..."):
        r0_df = pd.read_excel(r0_file, dtype=str).fillna("")
        r1_df = pd.read_excel(r1_file, dtype=str).fillna("")

        if "Tag" not in r0_df.columns or "Tag" not in r1_df.columns:
            st.error("Both files must contain a 'Tag' column.")
            st.stop()

        r0_df = r0_df.drop_duplicates(subset="Tag").set_index("Tag")
        r1_df = r1_df.drop_duplicates(subset="Tag").set_index("Tag")

        r0_columns = list(r0_df.columns)
        all_columns = sorted(
            set(r0_df.columns).union(set(r1_df.columns)),
            key=lambda x: r0_columns.index(x) if x in r0_columns else float("inf")
        )

        all_tags = sorted(set(r0_df.index).union(set(r1_df.index)))

        comparison_rows = []

        for tag in all_tags:
            if tag not in r0_df.index:
                row = {"Tag": tag, "Change_Type": "✅ Added in R1"}
                row.update({col: r1_df.loc[tag].get(col, "") for col in all_columns})
                row["Change_Summary"] = ""
                comparison_rows.append(row)

            elif tag not in r1_df.index:
                row = {"Tag": tag, "Change_Type": "❌ Removed in R1"}
                row.update({col: r0_df.loc[tag].get(col, "") for col in all_columns})
                row["Change_Summary"] = ""
                comparison_rows.append(row)

            else:
                row_r0 = r0_df.loc[tag]
                row_r1 = r1_df.loc[tag]
                row_data = {"Tag": tag}
                summary = []
                changes_exist = False

                for col in all_columns:
                    v0 = row_r0.get(col, "")
                    v1 = row_r1.get(col, "")
                    if str(v0).strip() != str(v1).strip():
                        row_data[col] = f"{v0} → {v1}"
                        summary.append(f"{col}: {v0} → {v1}")
                        changes_exist = True
                    else:
                        row_data[col] = v1

                row_data["Change_Type"] = "✏️ Modified" if changes_exist else "No Change"
                row_data["Change_Summary"] = " | ".join(summary)
                comparison_rows.append(row_data)

        comparison_df = pd.DataFrame(comparison_rows)
        final_columns = ["Tag", "Change_Type"] + all_columns + ["Change_Summary"]
        comparison_df = comparison_df[final_columns]

    # =========================
    # Metrics
    # =========================
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Tags (R0)", len(r0_df))
    c2.metric("Total Tags (R1)", len(r1_df))
    c3.metric("Added", (set(r1_df.index) - set(r0_df.index)).__len__())
    c4.metric("Removed", (set(r0_df.index) - set(r1_df.index)).__len__())

    # =========================
    # Filter
    # =========================
    filter_option = st.selectbox(
        "Filter View",
        ["All", "Modified", "Added", "Removed", "No Change"]
    )

    filtered_df = comparison_df.copy()

    if filter_option == "Modified":
        filtered_df = filtered_df[filtered_df["Change_Type"] == "✏️ Modified"]
    elif filter_option == "Added":
        filtered_df = filtered_df[filtered_df["Change_Type"] == "✅ Added in R1"]
    elif filter_option == "Removed":
        filtered_df = filtered_df[filtered_df["Change_Type"] == "❌ Removed in R1"]
    elif filter_option == "No Change":
        filtered_df = filtered_df[filtered_df["Change_Type"] == "No Change"]

    st.dataframe(filtered_df, use_container_width=True, height=500)

    # =========================
    # Excel Export
    # =========================
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison Summary"

    highlight = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for r_idx, row in enumerate(dataframe_to_rows(comparison_df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if isinstance(value, str) and "→" in value:
                cell.fill = highlight

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        wb.save(tmp.name)
        output_path = tmp.name

    st.download_button(
        label="⬇️ Download Comparison Excel",
        data=open(output_path, "rb"),
        file_name=f"Vimal_Comparison_R0_vs_R1_{datetime.now().strftime('%d_%m_%Y_%H_%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👆 Upload both R0 and R1 Excel files to begin comparison.")
