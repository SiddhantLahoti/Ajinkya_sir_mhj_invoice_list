import io
import zipfile
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.utils import column_index_from_string

st.set_page_config(
    page_title="MHJ Invoice Packing list",
    page_icon="📊",
    layout="wide",
)

st.title("📊 MHJ Invoice Packing list")
st.markdown(
    "Upload **File A (Source/Lookup)** and **File B (Target Invoice(s))** to cross-reference order/PO numbers "
    "and map values while preserving all formatting, styles, and formulas."
)


def get_sheet_names(uploaded_file, fallback):
    if uploaded_file is not None:
        try:
            wb = openpyxl.load_workbook(uploaded_file, read_only=True)
            names = wb.sheetnames
            uploaded_file.seek(0)
            return names
        except Exception:
            uploaded_file.seek(0)
    return fallback


def process_single_invoice(
    file_b,
    po_lookup,
    col_mappings,
    selected_sheet_b,
    po_idx_b,
    start_row_b,
):
    file_b.seek(0)
    wb_b = openpyxl.load_workbook(file_b, data_only=False)
    ws_b = (
        wb_b[selected_sheet_b]
        if selected_sheet_b in wb_b.sheetnames
        else wb_b.active
    )

    updated_count = 0
    missing_records = []

    for r in range(start_row_b, ws_b.max_row + 1):
        po_val = ws_b.cell(row=r, column=po_idx_b).value
        if po_val is None or not str(po_val).strip():
            continue

        po_key = str(po_val).strip()
        if po_key in po_lookup:
            for src_idx, tgt_idx, _, _, _ in col_mappings:
                cell_val = po_lookup[po_key][src_idx]
                if isinstance(cell_val, str):
                    cell_val = cell_val.strip()
                ws_b.cell(row=r, column=tgt_idx).value = cell_val
            updated_count += 1
        else:
            missing_records.append(
                {"Row in Invoice": r, "Order / PO Number": po_key}
            )

    output_buffer = io.BytesIO()
    wb_b.save(output_buffer)
    output_buffer.seek(0)

    original_name = file_b.name
    if original_name.lower().endswith(".xlsx"):
        output_filename = original_name[:-5] + "_updated.xlsx"
    elif original_name.lower().endswith(".xlsm"):
        output_filename = original_name[:-5] + "_updated.xlsm"
    else:
        output_filename = f"updated_{original_name}"

    return {
        "original_name": original_name,
        "output_filename": output_filename,
        "buffer": output_buffer,
        "updated_count": updated_count,
        "missing_records": missing_records,
    }


def render_workflow_tab(
    tab_key: str,
    uploader_a_label: str,
    uploader_b_label: str,
    default_sheet_a: str,
    default_po_col_a: str,
    default_start_row_a: int,
    default_po_col_b: str,
    default_start_row_b: int,
    default_mappings: list[dict],
):
    # ----------------- 1. File Uploaders -----------------
    col_file_a, col_file_b = st.columns(2)

    with col_file_a:
        file_a = st.file_uploader(
            uploader_a_label,
            type=["xlsx", "xlsm"],
            key=f"{tab_key}_file_a",
        )

    with col_file_b:
        files_b = st.file_uploader(
            uploader_b_label,
            type=["xlsx", "xlsm"],
            accept_multiple_files=True,
            key=f"{tab_key}_files_b",
        )

    # ----------------- 2. Settings & Coordinate Configuration -----------------
    st.markdown("---")
    st.subheader("⚙️ Coordinate & Mapping Settings")

    sheet_names_a = get_sheet_names(file_a, [default_sheet_a])
    first_file_b = files_b[0] if files_b else None
    sheet_names_b = get_sheet_names(first_file_b, ["Invoice"])

    config_col1, config_col2 = st.columns(2)

    with config_col1:
        st.markdown("#### File A Settings (Source / Tag File)")
        default_a_idx = (
            sheet_names_a.index(default_sheet_a)
            if default_sheet_a in sheet_names_a
            else 0
        )
        selected_sheet_a = st.selectbox(
            "File A Sheet",
            sheet_names_a,
            index=default_a_idx,
            key=f"{tab_key}_sheet_a",
        )
        po_col_a = (
            st.text_input(
                "Lookup Column in File A",
                value=default_po_col_a,
                help="Column in File A containing Order / PO numbers to match against",
                key=f"{tab_key}_po_col_a",
            )
            .strip()
            .upper()
        )
        start_row_a = st.number_input(
            "Data Start Row in File A",
            min_value=1,
            value=default_start_row_a,
            step=1,
            key=f"{tab_key}_start_row_a",
        )

    with config_col2:
        st.markdown("#### File B Settings (Target Invoice)")
        selected_sheet_b = st.selectbox(
            "File B Sheet",
            sheet_names_b,
            index=0,
            key=f"{tab_key}_sheet_b",
        )
        po_col_b = (
            st.text_input(
                "Lookup Column in File B",
                value=default_po_col_b,
                help="Column in Invoice containing Order / PO numbers to search for",
                key=f"{tab_key}_po_col_b",
            )
            .strip()
            .upper()
        )
        start_row_b = st.number_input(
            "Data Start Row in File B",
            min_value=1,
            value=default_start_row_b,
            step=1,
            key=f"{tab_key}_start_row_b",
        )

    # ----------------- 3. Column Mapping Table -----------------
    st.markdown("#### 🔄 Column Mappings (File A ➔ File B)")
    st.caption(
        "Define which column in File A maps to which column in File B. You can modify, add, or remove rows below:"
    )

    mapping_df = pd.DataFrame(default_mappings)
    edited_mapping_df = st.data_editor(
        mapping_df,
        num_rows="dynamic",
        width="stretch",
        key=f"{tab_key}_mapping_editor",
        column_config={
            "File A Column": st.column_config.TextColumn(
                "File A Source Col", required=True
            ),
            "File B Column": st.column_config.TextColumn(
                "File B Target Col", required=True
            ),
            "Field Description": st.column_config.TextColumn("Description / Label"),
        },
    )

    # ----------------- 4. Processing Engine -----------------
    st.markdown("---")
    process_btn = st.button(
        "🚀 Process & Update Invoice File(s)",
        type="primary",
        key=f"{tab_key}_process_btn",
    )

    if process_btn:
        if not file_a or not files_b:
            st.error(
                "⚠️ Please upload both File A and at least one File B invoice to proceed."
            )
            return

        with st.spinner(f"Processing {len(files_b)} invoice file(s)..."):
            try:
                # 1. Parse and validate column mappings
                col_mappings = []
                for _, row in edited_mapping_df.iterrows():
                    src = str(row["File A Column"]).strip().upper()
                    tgt = str(row["File B Column"]).strip().upper()
                    if src and tgt and src != "NONE" and tgt != "NONE":
                        src_idx = column_index_from_string(src)
                        tgt_idx = column_index_from_string(tgt)
                        col_mappings.append(
                            (
                                src_idx,
                                tgt_idx,
                                src,
                                tgt,
                                row.get("Field Description", ""),
                            )
                        )

                po_idx_a = column_index_from_string(po_col_a)
                po_idx_b = column_index_from_string(po_col_b)

                # 2. Read File A (Source) once with evaluated formula values
                file_a.seek(0)
                wb_a = openpyxl.load_workbook(file_a, data_only=True)
                ws_a = (
                    wb_a[selected_sheet_a]
                    if selected_sheet_a in wb_a.sheetnames
                    else wb_a.active
                )

                po_lookup = {}
                for r in range(start_row_a, ws_a.max_row + 1):
                    po_val = ws_a.cell(row=r, column=po_idx_a).value
                    if po_val is not None:
                        po_key = str(po_val).strip()
                        if po_key and po_key not in po_lookup:
                            po_lookup[po_key] = {
                                src_idx: ws_a.cell(row=r, column=src_idx).value
                                for src_idx, _, _, _, _ in col_mappings
                            }

                # 3. Process every uploaded File B invoice
                results = []
                for fb in files_b:
                    res = process_single_invoice(
                        file_b=fb,
                        po_lookup=po_lookup,
                        col_mappings=col_mappings,
                        selected_sheet_b=selected_sheet_b,
                        po_idx_b=po_idx_b,
                        start_row_b=start_row_b,
                    )
                    results.append(res)

                total_updated = sum(r["updated_count"] for r in results)
                total_missing = sum(len(r["missing_records"]) for r in results)

                # 4. Overall Summary Metrics
                st.success(
                    f"✅ Processed **{len(results)}** invoice file(s) and updated **{total_updated}** total rows!"
                )

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Indexed Records in File A", len(po_lookup))
                m2.metric("Invoices Processed", len(results))
                m3.metric("Total Rows Updated", total_updated)
                m4.metric("Total Unmatched Rows", total_missing)

                # 5. Bulk ZIP download when multiple invoices are uploaded
                if len(results) > 1:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(
                        zip_buffer, "w", zipfile.ZIP_DEFLATED
                    ) as zf:
                        for res in results:
                            zf.writestr(
                                res["output_filename"],
                                res["buffer"].getvalue(),
                            )
                    zip_buffer.seek(0)

                    st.download_button(
                        label=f"📦 Download All {len(results)} Updated Files (ZIP)",
                        data=zip_buffer,
                        file_name="MHJ_Updated_Invoices.zip",
                        mime="application/zip",
                        key=f"{tab_key}_download_all_zip",
                    )

                # 6. Per-file breakdown & individual download buttons
                st.markdown("#### 📄 Processed Invoice Files")
                for idx, res in enumerate(results):
                    with st.container(border=True):
                        c_info, c_stat, c_btn = st.columns([3, 2, 2])
                        with c_info:
                            st.markdown(f"**{res['original_name']}**")
                            st.caption(f"Output: `{res['output_filename']}`")
                        with c_stat:
                            st.write(
                                f"✅ Updated: **{res['updated_count']}** | "
                                f"⚠️ Unmatched: **{len(res['missing_records'])}**"
                            )
                        with c_btn:
                            st.download_button(
                                label="📥 Download Excel",
                                data=res["buffer"].getvalue(),
                                file_name=res["output_filename"],
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"{tab_key}_dl_{idx}_{res['output_filename']}",
                            )

                        if res["missing_records"]:
                            with st.expander(
                                f"⚠️ View {len(res['missing_records'])} Unmatched Record(s) in {res['original_name']}"
                            ):
                                st.table(pd.DataFrame(res["missing_records"]))

            except Exception as e:
                st.error(f"❌ An error occurred during processing: {str(e)}")


# ----------------- Tabs Setup -----------------
tab_ajinkya, tab_nilesh = st.tabs(
    ["Tab A : Ajinkya Sir", "Tab B : Nilesh Sir"]
)

with tab_ajinkya:
    ajinkya_mappings = [
        {"File A Column": "E", "File B Column": "S", "Field Description": "Colour Code"},
        {"File A Column": "F", "File B Column": "T", "Field Description": "Size Code"},
        {"File A Column": "G", "File B Column": "L", "Field Description": "Barcode"},
        {"File A Column": "H", "File B Column": "M", "Field Description": "Item Description"},
        {"File A Column": "I", "File B Column": "K", "Field Description": "Style Code"},
        {"File A Column": "O", "File B Column": "R", "Field Description": "Product Code"},
    ]
    render_workflow_tab(
        tab_key="ajinkya",
        uploader_a_label="Upload File A (Lookup Source - e.g., MHJ SPO)",
        uploader_b_label="Upload File B (Target Invoice(s) - Multiple Allowed)",
        default_sheet_a="Tag",
        default_po_col_a="B",
        default_start_row_a=2,
        default_po_col_b="I",
        default_start_row_b=25,
        default_mappings=ajinkya_mappings,
    )

with tab_nilesh:
    nilesh_mappings = [
        {"File A Column": "G", "File B Column": "I", "Field Description": "Customer PO"},
        {"File A Column": "J", "File B Column": "B", "Field Description": "Sku#"},
        {"File A Column": "M", "File B Column": "J", "Field Description": "Sku SR"},
        {"File A Column": "N", "File B Column": "M", "Field Description": "Description on PO #"},
        {"File A Column": "AJ", "File B Column": "R", "Field Description": "Item number"},
        {"File A Column": "AL", "File B Column": "K", "Field Description": "Style code"},
        {"File A Column": "AN", "File B Column": "T", "Field Description": "Size"},
        {"File A Column": "AO", "File B Column": "S", "Field Description": "Colour"},
    ]
    render_workflow_tab(
        tab_key="nilesh",
        uploader_a_label="Upload File A (MHJ PO Tags / Shipment List)",
        uploader_b_label="Upload File B (MHJ Invoice(s) - Multiple Allowed)",
        default_sheet_a="MHJ OPS",
        default_po_col_a="E",
        default_start_row_a=4,
        default_po_col_b="G",
        default_start_row_b=25,
        default_mappings=nilesh_mappings,
    )