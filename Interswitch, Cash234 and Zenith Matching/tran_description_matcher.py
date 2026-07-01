from __future__ import annotations

import io
import re
from typing import Dict, Optional

import pandas as pd
import streamlit as st

from reconciliation import (
    ReconciliationError,
    calculate_excel_column_width,
    canonical_header,
    clean_headers,
    find_column,
    list_sheets,
    normalize_amount_series,
    read_transaction_file,
    safe_sheet_name,
)
from source_processing import (
    description_contains_tran_id_token,
    extract_interswitch_tran_id_token,
    normalize_tran_description_token,
)


def main() -> None:
    st.set_page_config(
        page_title="Tran ID vs Zenith Description Matcher",
        page_icon="",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.title("Tran ID vs Zenith Description Matcher")
    st.caption("Separate comparison for Interswitch Tran_ID tokens against useful Zenith description sections.")

    with st.sidebar:
        st.header("Upload Files")
        interswitch_file, interswitch_sheet = upload_file_with_sheet("Interswitch", "interswitch")
        zenith_file, zenith_sheet = upload_file_with_sheet("Zenith", "zenith")

    if interswitch_file is None or zenith_file is None:
        st.info("Upload the Interswitch and Zenith files to begin.")
        return

    try:
        interswitch_df = read_transaction_file(
            interswitch_file,
            interswitch_file.name,
            interswitch_sheet if interswitch_sheet != "CSV" else None,
        )
        zenith_df = read_transaction_file(
            zenith_file,
            zenith_file.name,
            zenith_sheet if zenith_sheet != "CSV" else None,
        )
        result = build_tran_description_comparison(interswitch_df, zenith_df)
    except Exception as exc:
        st.error(str(exc))
        return

    st.subheader("Summary")
    summary = result["Summary"]
    cols = st.columns(4)
    for idx, row in summary.iterrows():
        with cols[idx % 4]:
            st.metric(str(row["Metric"]), f"{row['Value']:,}" if isinstance(row["Value"], int) else row["Value"])

    st.divider()
    st.subheader("Downloads")
    st.download_button(
        "Download Tran ID vs Description Report",
        data=workbook_sheets_to_bytes(result),
        file_name="Tran_ID_vs_Zenith_Description.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    standalone_cols = st.columns(2)
    with standalone_cols[0]:
        st.download_button(
            "Download Interswitch Unmatched",
            data=workbook_sheets_to_bytes(
                {"Interswitch Unmatched": result["Interswitch Unmatched"]},
                clean_export=True,
            ),
            file_name="Interswitch_Unmatched.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with standalone_cols[1]:
        st.download_button(
            "Download Zenith Unmatched",
            data=workbook_sheets_to_bytes(
                {"Zenith Unmatched": result["Zenith Unmatched"]},
                clean_export=True,
            ),
            file_name="Zenith_Unmatched.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()
    st.subheader("Preview")
    sheet_name = st.selectbox("Choose sheet", list(result.keys()))
    st.dataframe(result[sheet_name], use_container_width=True, height=520, hide_index=True)


def upload_file_with_sheet(label: str, key_prefix: str) -> tuple[Optional[object], Optional[str]]:
    file = st.file_uploader(f"Upload {label} file", type=["xlsx", "xls", "csv"], key=f"{key_prefix}_file")
    if file is None:
        return None, None
    st.caption(f"{file.name} | {file.size / 1024:,.1f} KB")
    sheets = list_sheets(file, file.name)
    if len(sheets) == 1:
        st.caption(f"Sheet: {sheets[0]}")
        return file, sheets[0]
    sheet = st.selectbox(f"{label} sheet", sheets, key=f"{key_prefix}_sheet")
    return file, sheet


def build_tran_description_comparison(
    interswitch_df: pd.DataFrame,
    zenith_df: pd.DataFrame,
) -> Dict[str, pd.DataFrame]:
    interswitch = clean_headers(interswitch_df).copy()
    zenith = clean_headers(zenith_df).copy()

    tran_id_column = find_column(interswitch, "Tran_ID", "INTERSWITCH")
    description_column = find_column(zenith, "Description", "ZENITH")

    interswitch["INTERSWITCH_SOURCE_ROW_NUMBER"] = range(2, len(interswitch) + 2)
    zenith["ZENITH_SOURCE_ROW_NUMBER"] = range(2, len(zenith) + 2)
    interswitch["MATCH_TOKEN"] = interswitch[tran_id_column].map(extract_interswitch_tran_id_token)
    zenith["NORMALIZED_DESCRIPTION_SEARCH_TEXT"] = zenith[description_column].map(normalize_tran_description_token)

    interswitch_valid = interswitch.loc[interswitch["MATCH_TOKEN"].astype("string").str.strip().ne("")].copy()
    zenith_valid = zenith.loc[zenith["NORMALIZED_DESCRIPTION_SEARCH_TEXT"].astype("string").str.strip().ne("")].copy()

    interswitch_valid["TOKEN_SEQUENCE"] = interswitch_valid.groupby("MATCH_TOKEN").cumcount()
    interswitch_valid["INTERSWITCH_TOKEN_COUNT"] = interswitch_valid.groupby("MATCH_TOKEN")["MATCH_TOKEN"].transform("size")

    interswitch_prefixed = prefix_columns(interswitch_valid, "INTERSWITCH", {"MATCH_TOKEN", "TOKEN_SEQUENCE"})
    zenith_prefixed = prefix_columns(zenith_valid, "ZENITH")

    matched, matched_interswitch_rows, matched_zenith_rows = build_interswitch_driven_matches(
        interswitch_prefixed=interswitch_prefixed,
        zenith_prefixed=zenith_prefixed,
        description_column=f"ZENITH_{description_column}",
    )

    interswitch_unmatched = interswitch_prefixed.loc[
        ~interswitch_prefixed["INTERSWITCH_SOURCE_ROW_NUMBER"].isin(matched_interswitch_rows)
    ].copy()
    interswitch_unmatched["MATCH_STATUS"] = "Tran_ID token not found in Zenith descriptions"

    zenith_unmatched = zenith_prefixed.loc[
        ~zenith_prefixed["ZENITH_SOURCE_ROW_NUMBER"].isin(matched_zenith_rows)
    ].copy()
    zenith_unmatched["MATCH_STATUS"] = "Zenith description token not found in Interswitch Tran_ID"

    invalid_tokens = pd.concat(
        [
            add_invalid_token_status(
                prefix_columns(interswitch.loc[~interswitch.index.isin(interswitch_valid.index)], "INTERSWITCH"),
                "INTERSWITCH",
                "No useful Tran_ID token found",
            ),
            add_invalid_token_status(
                prefix_columns(zenith.loc[~zenith.index.isin(zenith_valid.index)], "ZENITH"),
                "ZENITH",
                "Blank Zenith description search text",
            ),
        ],
        ignore_index=True,
        sort=False,
    )

    summary = pd.DataFrame(
        [
            ("Total Interswitch rows", len(interswitch)),
            ("Total Zenith rows", len(zenith)),
            ("Interswitch rows with usable Tran_ID token", len(interswitch_valid)),
            ("Zenith rows with searchable description", len(zenith_valid)),
            ("Matched token pairs", len(matched)),
            ("Unmatched Interswitch token rows", len(interswitch_unmatched)),
            ("Zenith rows without matched Interswitch token", len(zenith_unmatched)),
            ("Rows without usable token", len(invalid_tokens)),
        ],
        columns=["Metric", "Value"],
    )

    return {
        "Summary": summary,
        "Matched": order_match_columns(matched),
        "Interswitch Unmatched": order_match_columns(interswitch_unmatched),
        "Zenith Unmatched": order_match_columns(zenith_unmatched),
        "Invalid Tokens": order_match_columns(invalid_tokens),
    }


def prefix_columns(df: pd.DataFrame, dataset: str, keep_columns: Optional[set[str]] = None) -> pd.DataFrame:
    keep_columns = keep_columns or set()
    renamed = {}
    for column in df.columns:
        if column in keep_columns:
            continue
        if not str(column).startswith(f"{dataset}_"):
            renamed[column] = f"{dataset}_{column}"
    return df.rename(columns=renamed)


def build_interswitch_driven_matches(
    interswitch_prefixed: pd.DataFrame,
    zenith_prefixed: pd.DataFrame,
    description_column: str,
) -> tuple[pd.DataFrame, set[object], set[object]]:
    matched_rows = []
    matched_interswitch_rows = set()
    matched_zenith_rows = set()

    for _index, interswitch_row in interswitch_prefixed.iterrows():
        token = interswitch_row.get("MATCH_TOKEN", "")
        if not str(token).strip():
            continue

        available_zenith = zenith_prefixed.loc[
            ~zenith_prefixed["ZENITH_SOURCE_ROW_NUMBER"].isin(matched_zenith_rows)
        ]
        match_candidates = available_zenith.loc[
            available_zenith[description_column].map(
                lambda description: description_contains_tran_id_token(description, token)
            )
        ]
        if match_candidates.empty:
            continue

        zenith_row = match_candidates.iloc[0]
        matched_interswitch_rows.add(interswitch_row["INTERSWITCH_SOURCE_ROW_NUMBER"])
        matched_zenith_rows.add(zenith_row["ZENITH_SOURCE_ROW_NUMBER"])
        matched_rows.append(
            {
                **interswitch_row.to_dict(),
                **zenith_row.to_dict(),
                "MATCH_STATUS": "Interswitch Tran_ID token found in Zenith description",
            }
        )

    if not matched_rows:
        columns = list(interswitch_prefixed.columns) + [
            column for column in zenith_prefixed.columns if column not in interswitch_prefixed.columns
        ]
        columns.append("MATCH_STATUS")
        return pd.DataFrame(columns=columns), matched_interswitch_rows, matched_zenith_rows

    return pd.DataFrame(matched_rows), matched_interswitch_rows, matched_zenith_rows


def add_invalid_token_status(df: pd.DataFrame, dataset: str, status: str) -> pd.DataFrame:
    output = df.copy()
    output["MATCH_STATUS"] = status
    output["SOURCE"] = dataset
    return output


def order_match_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    priority = ["MATCH_TOKEN", "TOKEN_SEQUENCE", "MATCH_STATUS", "SOURCE"]
    ordered = [column for column in priority if column in df.columns]
    ordered.extend([column for column in df.columns if column.startswith("INTERSWITCH_")])
    ordered.extend([column for column in df.columns if column.startswith("ZENITH_")])
    ordered.extend([column for column in df.columns if column not in ordered])
    return df.loc[:, ordered]


INTERNAL_EXPORT_COLUMNS = {
    "MATCH_TOKEN",
    "TOKEN_SEQUENCE",
    "MATCH_STATUS",
    "SOURCE",
    "NORMALIZED_DESCRIPTION_SEARCH_TEXT",
    "INTERSWITCH_TOKEN_COUNT",
    "INTERSWITCH_SOURCE_ROW_NUMBER",
    "ZENITH_SOURCE_ROW_NUMBER",
}
MONEY_HEADER_WORDS = (
    "AMOUNT",
    "VALUE",
    "SETTLEMENT",
    "IMPACT",
    "DEBIT",
    "CREDIT",
    "FEE",
    "CHARGE",
    "COMMISSION",
    "BALANCE",
    "TRANSACTIONAMOUNT",
)


def workbook_sheets_to_bytes(sheets: Dict[str, pd.DataFrame], clean_export: bool = True) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, frame in sheets.items():
            export_frame = prepare_export_frame(frame) if clean_export else frame.copy()
            write_sheet(writer, export_frame, safe_sheet_name(sheet_name))
    return output.getvalue()


def prepare_export_frame(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        export_df = drop_internal_export_columns(df.copy())
        export_df.columns = [remove_source_prefix(column) for column in export_df.columns]
        return drop_internal_export_columns(export_df)
    export_df = drop_internal_export_columns(df.copy())
    export_df.columns = [remove_source_prefix(column) for column in export_df.columns]
    export_df = drop_internal_export_columns(export_df)
    return export_df


def drop_internal_export_columns(df: pd.DataFrame) -> pd.DataFrame:
    return df.drop(columns=[column for column in INTERNAL_EXPORT_COLUMNS if column in df.columns], errors="ignore")


def remove_source_prefix(column: object) -> object:
    if not isinstance(column, str):
        return column
    for prefix in ("INTERSWITCH_", "ZENITH_"):
        if column.startswith(prefix):
            return column.removeprefix(prefix)
    return column


def write_sheet(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str) -> None:
    safe_df = df.copy()
    total_column_indices = find_monetary_column_indices(safe_df)
    if total_column_indices:
        safe_df = append_total_row(safe_df, total_column_indices)
    safe_df.to_excel(writer, sheet_name=sheet_name, index=False)
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    header_format = workbook.add_format(
        {
            "bold": True,
            "bg_color": "#1F4E78",
            "font_color": "white",
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        }
    )
    cell_format = workbook.add_format({"align": "center", "valign": "vcenter"})
    money_format = workbook.add_format({"num_format": "#,##0.00", "align": "center", "valign": "vcenter"})
    total_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "align": "center", "valign": "vcenter"})
    total_money_format = workbook.add_format(
        {"bold": True, "bg_color": "#D9EAF7", "num_format": "#,##0.00", "align": "center", "valign": "vcenter"}
    )
    for col_idx, column in enumerate(safe_df.columns):
        worksheet.write(0, col_idx, column, header_format)
        column_format = money_format if col_idx in total_column_indices else cell_format
        worksheet.set_column(col_idx, col_idx, calculate_excel_column_width(safe_df, col_idx, column), column_format)
        if col_idx in total_column_indices:
            coerce_excel_numeric_column(worksheet, safe_df.iloc[:, col_idx], col_idx, money_format, total_money_format)
    if total_column_indices:
        total_row_idx = len(safe_df)
        label_col_idx = find_total_label_column(safe_df, total_column_indices)
        worksheet.write(total_row_idx, label_col_idx, "Total", total_format)
        for col_idx in total_column_indices:
            worksheet.write_number(total_row_idx, col_idx, float(safe_df.iloc[-1, col_idx]), total_money_format)
    if len(safe_df) > 0 and len(safe_df.columns) > 0:
        worksheet.freeze_panes(1, 0)
    worksheet.set_row(0, 24, header_format)


def find_monetary_column_indices(df: pd.DataFrame) -> list[int]:
    money_columns = []
    for col_idx, column in enumerate(df.columns):
        if not is_monetary_column(column):
            continue
        numeric = numeric_money_series(df.iloc[:, col_idx])
        if numeric.notna().any():
            money_columns.append(col_idx)
    return money_columns


def is_monetary_column(column: object) -> bool:
    canonical = canonical_header(str(column))
    if canonical in {"RRN", "REFERENCENR", "RETRIEVALREFERENCENR", "TRANID", "DESCRIPTION"}:
        return False
    if re.search(r"(DATE|TIME|ROW|TOKEN|REFERENCE|RRN|ID|CODE|ACCOUNT|DESCRIPTION)", canonical):
        return False
    return any(word in canonical for word in MONEY_HEADER_WORDS)


def numeric_money_series(series: pd.Series) -> pd.Series:
    normalized = normalize_amount_series(series)
    return pd.to_numeric(normalized, errors="coerce")


def append_total_row(df: pd.DataFrame, total_column_indices: list[int]) -> pd.DataFrame:
    output = df.copy()
    total_row = [""] * len(output.columns)
    for col_idx in total_column_indices:
        total_row[col_idx] = numeric_money_series(output.iloc[:, col_idx]).sum()
    total_frame = pd.DataFrame([total_row], columns=output.columns)
    return pd.concat([output, total_frame], ignore_index=True, sort=False)


def find_total_label_column(df: pd.DataFrame, total_column_indices: list[int]) -> int:
    for index, column in enumerate(df.columns):
        if index not in total_column_indices:
            return index
    return 0


def coerce_excel_numeric_column(
    worksheet: object,
    series: pd.Series,
    col_idx: int,
    money_format: object,
    total_money_format: object,
) -> None:
    numeric = numeric_money_series(series.iloc[:-1])
    for row_offset, value in enumerate(numeric, start=1):
        if pd.notna(value):
            worksheet.write_number(row_offset, col_idx, float(value), money_format)
    if len(series) > 0 and pd.notna(series.iloc[-1]):
        worksheet.write_number(len(series), col_idx, float(series.iloc[-1]), total_money_format)


if __name__ == "__main__":
    main()
