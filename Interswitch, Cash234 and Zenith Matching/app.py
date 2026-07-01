from __future__ import annotations

import io
import hashlib
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, Optional
from uuid import uuid4

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

try:
    from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
except ImportError:  # Optional enhancement; the app still works offline without it.
    AgGrid = None
    GridOptionsBuilder = None
    GridUpdateMode = None

from reconciliation import (
    DATASETS,
    GAP_COLUMNS,
    INVALID_SHEET_NAME,
    REPORT_DOWNLOAD_ORDER,
    ReconciliationError,
    build_required_report_workbook,
    canonical_header,
    clean_headers,
    duplicates_workbook_to_bytes,
    find_column,
    filter_dataframe,
    invalid_rrn_workbook_to_bytes,
    interswitch_zenith_carryover_workbook_to_bytes,
    is_valid_rrn_series,
    list_sheets,
    normalize_rrn_series,
    parse_sort_date_series,
    read_transaction_file,
    reconcile,
    updated_export_filename,
    updated_source_workbook_to_bytes,
)


st.set_page_config(
    page_title="Three-Way Transaction Reconciliation",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded",
)


STATUS_COLORS = {
    "MATCHED_IN_ALL_THREE": "#1E7D32",
    "INVALID_OR_BLANK_RRN": "#6B7280",
    "ONLY": "#B42318",
    "PARTIAL": "#B7791F",
}
FOLDER_UPLOAD_HEADER_SCAN_ROWS = 12


def main() -> None:
    inject_css()
    st.title("Three-Way Transaction Reconciliation")
    st.caption("Offline comparison for INTERSWITCH, CASH234, and ZENITH using normalized RRN values.")

    with st.sidebar:
        st.header("Upload Files")
        st.info("Files are processed locally in memory. Nothing is uploaded to the internet.")
        uploads = collect_uploads()

    if not uploads_ready(uploads):
        render_empty_state()
        return

    input_fingerprint = build_upload_fingerprint(uploads)
    if st.session_state.get("input_fingerprint") != input_fingerprint:
        st.session_state["input_fingerprint"] = input_fingerprint
        st.session_state.pop("reconciliation_results", None)
        st.session_state.pop("raw_frames", None)
        st.session_state.pop("ingestion_stats", None)
        st.session_state["cache_run_id"] = uuid4().hex

    if "raw_frames" in st.session_state and "ingestion_stats" in st.session_state:
        raw_frames = st.session_state["raw_frames"]
        ingestion_stats = st.session_state["ingestion_stats"]
    else:
        try:
            raw_frames, ingestion_stats = load_uploaded_frames(uploads)
            st.session_state["raw_frames"] = raw_frames
            st.session_state["ingestion_stats"] = ingestion_stats
        except Exception as exc:
            st.error(f"Could not read one of the uploaded files: {exc}")
            return

    render_ingestion_preview(ingestion_stats)

    if st.button("Run reconciliation", type="primary", use_container_width=True):
        run_reconciliation(raw_frames, ingestion_stats)

    if "reconciliation_results" in st.session_state:
        render_results(st.session_state["reconciliation_results"])


def collect_uploads() -> Dict[str, Dict[str, object]]:
    upload_mode = st.radio("Upload Mode", ["Single Files", "Folder Upload"], horizontal=False)
    upload_config = {
        "INTERSWITCH": "Upload INTERSWITCH file",
        "CASH234": "Upload CASH234 file",
        "ZENITH": "Upload ZENITH file",
    }
    uploads: Dict[str, Dict[str, object]] = {"__mode__": {"value": upload_mode}}
    for dataset, label in upload_config.items():
        if upload_mode == "Folder Upload":
            st.caption("Close any open Excel/WPS workbooks first so temporary ~$ lock files are not selected.")
            files = st.file_uploader(
                label.replace("file", "folder"),
                type=["xlsx", "xls", "csv"],
                accept_multiple_files="directory",
                key=f"upload_folder_{dataset}",
            )
            uploads[dataset] = {"files": files or [], "sheet": None}
            if files:
                st.caption(f"{len(files):,} candidate files selected")
            continue

        file = st.file_uploader(label, type=["xlsx", "xls", "csv"], key=f"upload_{dataset}")
        sheet: Optional[str] = None
        if file is not None:
            st.caption(f"{file.name} | {file.size / 1024:,.1f} KB")
            try:
                sheets = list_sheets(file, file.name)
                if len(sheets) > 1:
                    sheet = st.selectbox(f"{dataset} sheet", sheets, key=f"sheet_{dataset}")
                else:
                    sheet = sheets[0]
                    st.caption(f"Sheet: {sheet}")
            except Exception as exc:
                st.error(f"Could not inspect sheets: {exc}")
        uploads[dataset] = {"file": file, "sheet": sheet}
    return uploads


def uploads_ready(uploads: Dict[str, Dict[str, object]]) -> bool:
    mode = str(uploads.get("__mode__", {}).get("value", "Single Files"))
    for dataset in DATASETS:
        item = uploads.get(dataset, {})
        if mode == "Folder Upload":
            if not item.get("files"):
                return False
        elif item.get("file") is None:
            return False
    return True


def build_upload_fingerprint(uploads: Dict[str, Dict[str, object]]) -> str:
    mode = str(uploads.get("__mode__", {}).get("value", "Single Files"))
    parts = [mode]
    for dataset in DATASETS:
        item = uploads.get(dataset, {})
        if mode == "Folder Upload":
            files = item.get("files", [])
            for file in sorted(files, key=lambda uploaded: str(getattr(uploaded, "name", ""))):
                parts.append(
                    "|".join(
                        [
                            dataset,
                            str(getattr(file, "name", "")),
                            str(getattr(file, "size", 0)),
                            uploaded_file_digest(file),
                        ]
                    )
                )
        else:
            file = item.get("file")
            if file is not None:
                parts.append(
                    "|".join(
                        [
                            dataset,
                            str(getattr(file, "name", "")),
                            str(getattr(file, "size", 0)),
                            str(item.get("sheet", "")),
                            uploaded_file_digest(file),
                        ]
                    )
                )
    return hashlib.sha256("\n".join(parts).encode("utf-8")).hexdigest()


def load_uploaded_frames(uploads: Dict[str, Dict[str, object]]) -> tuple[Dict[str, pd.DataFrame], Dict[str, Dict[str, object]]]:
    mode = str(uploads.get("__mode__", {}).get("value", "Single Files"))
    if mode == "Folder Upload":
        return load_folder_uploaded_frames(uploads)
    return load_single_uploaded_frames(uploads)


def load_single_uploaded_frames(
    uploads: Dict[str, Dict[str, object]],
) -> tuple[Dict[str, pd.DataFrame], Dict[str, Dict[str, object]]]:
    frames: Dict[str, pd.DataFrame] = {}
    stats: Dict[str, Dict[str, object]] = {}
    progress = st.progress(0, text="Reading files...")
    for index, dataset in enumerate(DATASETS, start=1):
        file = uploads[dataset]["file"]
        sheet = uploads[dataset]["sheet"]
        if file is None:
            raise ReconciliationError(f"{dataset} file is required.")
        frame = read_transaction_file(file, file.name, sheet if sheet != "CSV" else None)
        frame = force_reconciliation_key_to_string(frame, dataset)
        frames[dataset] = frame
        stats[dataset] = build_ingestion_stats(
            dataset=dataset,
            frame=frame,
            loaded_files=1,
            discovered_files=1,
            failed_files=[],
        )
        progress.progress(index / len(DATASETS), text=f"Loaded {dataset}")
    progress.empty()
    return frames, stats


def load_folder_uploaded_frames(
    uploads: Dict[str, Dict[str, object]],
) -> tuple[Dict[str, pd.DataFrame], Dict[str, Dict[str, object]]]:
    frames: Dict[str, pd.DataFrame] = {}
    stats: Dict[str, Dict[str, object]] = {}
    progress = st.progress(0, text="Discovering files...")

    for dataset_index, dataset in enumerate(DATASETS):
        selected_files = list(uploads[dataset].get("files", []))
        discovered_files, ignored_files = discover_supported_uploads(selected_files)
        progress.progress(
            dataset_index / len(DATASETS),
            text=f"{dataset}: {len(discovered_files):,} files discovered",
        )
        if not discovered_files:
            raise ReconciliationError(f"{dataset} folder does not contain any supported non-empty files.")

        frame, dataset_stats = load_dataset_folder_files(dataset, discovered_files, progress)
        dataset_stats["ignored_files"] = ignored_files
        frames[dataset] = frame
        stats[dataset] = dataset_stats

    progress.progress(0.95, text="Saving consolidated parquet cache...")
    cache_run_id = st.session_state.setdefault("cache_run_id", uuid4().hex)
    cache_master_frames(frames, str(cache_run_id))
    cached_frames = load_cached_master_frames(frames, str(cache_run_id))
    progress.progress(1.0, text="Folder upload consolidation complete")
    progress.empty()
    return cached_frames, stats


def discover_supported_uploads(files: list[object]) -> tuple[list[object], list[Dict[str, str]]]:
    discovered = []
    ignored = []
    seen_hashes: Dict[str, str] = {}
    for file in files:
        filename = str(getattr(file, "name", ""))
        size = int(getattr(file, "size", 0) or 0)
        reason = ignored_upload_reason(filename, size)
        if not reason:
            digest = uploaded_file_digest(file)
            if digest in seen_hashes:
                reason = f"Duplicate file content already loaded from {seen_hashes[digest]}"
            else:
                seen_hashes[digest] = filename
        if reason:
            ignored.append({"file": filename, "reason": reason})
        else:
            discovered.append(file)
    return discovered, ignored


def ignored_upload_reason(filename: str, size: int) -> Optional[str]:
    path = Path(filename)
    name = path.name
    path_parts = [part for part in filename.replace("\\", "/").split("/") if part]
    suffix = path.suffix.lower()
    if size <= 0:
        return "Empty file"
    if suffix not in {".xlsx", ".xls", ".csv"}:
        return "Unsupported file type"
    if name.startswith("~$") or any(part.startswith("~$") for part in path_parts):
        return "Temporary Excel lock file"
    if name.startswith(".") or any(part.startswith(".") for part in path_parts):
        return "Hidden file"
    if name.lower().endswith((".tmp", ".temp")):
        return "Temporary file"
    return None


def load_dataset_folder_files(
    dataset: str,
    files: list[object],
    progress: st.delta_generator.DeltaGenerator,
) -> tuple[pd.DataFrame, Dict[str, object]]:
    loaded_frames: list[pd.DataFrame] = []
    failed_files: list[Dict[str, str]] = []
    header_map: Dict[str, str] = {}
    rows_loaded = 0
    max_workers = min(8, max(1, len(files)))

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_filename = {
            executor.submit(read_uploaded_file_payload, file, dataset): str(getattr(file, "name", "Unknown file"))
            for file in files
        }
        for processed_count, future in enumerate(as_completed(future_to_filename), start=1):
            try:
                filename, frame = future.result()
                frame = align_headers_case_insensitive(frame, header_map)
                frame["SOURCE_FILE"] = Path(filename).name
                frame = force_reconciliation_key_to_string(frame, dataset)
                loaded_frames.append(frame)
                rows_loaded += len(frame)
            except Exception as exc:
                failed_files.append({"file": future_to_filename[future], "reason": str(exc)})

            progress.progress(
                min(processed_count / max(len(files), 1), 0.9),
                text=f"{display_dataset_label(dataset)}: {processed_count:,}/{len(files):,} files processed, {rows_loaded:,} rows loaded",
            )

    if not loaded_frames:
        failed_summary = summarize_failed_files(failed_files)
        raise ReconciliationError(
            f"No {dataset} files could be loaded successfully. {failed_summary}"
        )

    consolidated = pd.concat(loaded_frames, ignore_index=True, copy=False, sort=False)
    consolidated = optimize_master_dataframe(consolidated, dataset)
    consolidated = sort_raw_dataset_by_own_date(consolidated, dataset)

    if len(consolidated) > 500_000:
        st.warning("Large dataset detected. Processing may take several minutes.")

    return consolidated, build_ingestion_stats(
        dataset=dataset,
        frame=consolidated,
        loaded_files=len(loaded_frames),
        discovered_files=len(files),
        failed_files=failed_files,
    )


def uploaded_file_digest(file: object) -> str:
    payload = getattr(file, "getvalue")()
    return hashlib.sha256(payload).hexdigest()


def summarize_failed_files(failed_files: list[Dict[str, str]], limit: int = 5) -> str:
    if not failed_files:
        return ""
    examples = [
        f"{Path(item.get('file', 'Unknown file')).name}: {item.get('reason', 'Unknown error')}"
        for item in failed_files[:limit]
    ]
    suffix = "" if len(failed_files) <= limit else f" ...and {len(failed_files) - limit:,} more."
    return "Failed files: " + " | ".join(examples) + suffix


def read_uploaded_file_payload(file: object, dataset: Optional[str] = None) -> tuple[str, pd.DataFrame]:
    filename = str(getattr(file, "name", ""))
    payload = getattr(file, "getvalue")()
    frame = read_transaction_payload(io.BytesIO(payload), filename, dataset)
    return filename, frame


def read_transaction_payload(buffer: io.BytesIO, filename: str, dataset: Optional[str] = None) -> pd.DataFrame:
    suffix = filename.lower().rsplit(".", 1)[-1]
    buffer.seek(0)
    if suffix == "csv":
        frame = read_csv_payload_with_header_scan(buffer, filename, dataset)
    elif suffix in {"xlsx", "xls"}:
        frame = read_excel_payload_with_header_scan(buffer, filename, dataset)
    else:
        raise ReconciliationError(f"Unsupported file type for {filename}. Upload .xlsx, .xls, or .csv.")
    return clean_headers(frame)


def read_csv_payload_with_header_scan(
    buffer: io.BytesIO,
    filename: str,
    dataset: Optional[str],
) -> pd.DataFrame:
    first_frame: Optional[pd.DataFrame] = None
    first_error: Optional[Exception] = None
    for header_row in candidate_header_rows(dataset):
        try:
            buffer.seek(0)
            frame = pd.read_csv(
                buffer,
                header=header_row,
                dtype=str,
                keep_default_na=False,
                na_filter=False,
            )
            if first_frame is None:
                first_frame = frame
            if dataset is None or has_required_reconciliation_column(frame, dataset):
                return frame
        except Exception as exc:
            first_error = first_error or exc
    if first_frame is not None:
        raise_missing_folder_column_error(filename, dataset, first_frame, ["CSV"])
    raise first_error or ReconciliationError(f"Could not read {filename}.")


def read_excel_payload_with_header_scan(
    buffer: io.BytesIO,
    filename: str,
    dataset: Optional[str],
) -> pd.DataFrame:
    first_frame: Optional[pd.DataFrame] = None
    first_error: Optional[Exception] = None
    try:
        buffer.seek(0)
        excel = pd.ExcelFile(buffer)
        sheet_names = excel.sheet_names or [0]
    except Exception as exc:
        raise ReconciliationError(f"{Path(filename).name} could not be opened as an Excel workbook: {exc}")

    for sheet_name in sheet_names:
        for header_row in candidate_header_rows(dataset):
            try:
                frame = pd.read_excel(
                    excel,
                    sheet_name=sheet_name,
                    header=header_row,
                    dtype=str,
                    keep_default_na=False,
                    engine=None,
                )
                if first_frame is None:
                    first_frame = frame
                if dataset is None or has_required_reconciliation_column(frame, dataset):
                    return frame
            except Exception as exc:
                first_error = first_error or exc

    if first_frame is not None:
        raise_missing_folder_column_error(filename, dataset, first_frame, sheet_names)
    raise first_error or ReconciliationError(f"Could not read {filename}.")


def candidate_header_rows(dataset: Optional[str]) -> range:
    if dataset is None:
        return range(1)
    return range(FOLDER_UPLOAD_HEADER_SCAN_ROWS)


def has_required_reconciliation_column(frame: pd.DataFrame, dataset: str) -> bool:
    try:
        find_column(clean_headers(frame), str(DATASETS[dataset]["rrn_column"]), dataset)
        return True
    except ReconciliationError:
        return False


def raise_missing_folder_column_error(
    filename: str,
    dataset: Optional[str],
    frame: pd.DataFrame,
    sheet_names: object,
) -> None:
    if dataset is None:
        raise ReconciliationError(f"{Path(filename).name} could not be loaded.")
    cleaned = clean_headers(frame)
    available = ", ".join(map(str, cleaned.columns[:25]))
    sheets = ", ".join(map(str, sheet_names))
    raise ReconciliationError(
        f"{Path(filename).name} is missing required {dataset} column "
        f"'{DATASETS[dataset]['rrn_column']}'. Checked sheet(s): {sheets}; "
        f"header rows 1-{FOLDER_UPLOAD_HEADER_SCAN_ROWS}. Available columns from the first attempt include: {available}"
    )


def align_headers_case_insensitive(df: pd.DataFrame, header_map: Dict[str, str]) -> pd.DataFrame:
    renamed = {}
    existing_targets = set(df.columns)
    for column in df.columns:
        canonical = canonical_header(str(column))
        target = header_map.setdefault(canonical, str(column))
        if target != column and target not in existing_targets:
            renamed[column] = target
            existing_targets.add(target)
    return df.rename(columns=renamed, copy=False)


def force_reconciliation_key_to_string(df: pd.DataFrame, dataset: str) -> pd.DataFrame:
    rrn_column = find_column(df, str(DATASETS[dataset]["rrn_column"]), dataset)
    df[rrn_column] = df[rrn_column].astype("string").fillna("")
    return df


def sort_raw_dataset_by_own_date(df: pd.DataFrame, dataset: str) -> pd.DataFrame:
    if df.empty:
        return df
    config = DATASETS[dataset]
    try:
        date_column = find_column(df, str(config["date_column"]), dataset)
    except ReconciliationError:
        return df
    sorted_df = df.copy()
    sorted_df["__SORT_PRIORITY_DATE"] = parse_sort_date_series(
        sorted_df[date_column],
        dayfirst=bool(config["date_dayfirst"]),
    )
    sorted_df["__SORT_ORIGINAL_POSITION"] = range(len(sorted_df))
    sorted_df = sorted_df.sort_values(
        by=["__SORT_PRIORITY_DATE", "__SORT_ORIGINAL_POSITION"],
        ascending=[True, True],
        na_position="last",
        kind="mergesort",
    )
    return sorted_df.drop(columns=["__SORT_PRIORITY_DATE", "__SORT_ORIGINAL_POSITION"])


def optimize_master_dataframe(df: pd.DataFrame, dataset: str) -> pd.DataFrame:
    if df.empty:
        return df
    rrn_column = find_column(df, str(DATASETS[dataset]["rrn_column"]), dataset)
    protected_columns = {rrn_column}
    try:
        protected_columns.add(find_column(df, str(DATASETS[dataset]["date_column"]), dataset))
    except ReconciliationError:
        pass
    optimized = df.copy()
    for column in optimized.columns:
        if column in protected_columns:
            optimized[column] = optimized[column].astype("string").fillna("")
            continue
        if pd.api.types.is_string_dtype(optimized[column]) or optimized[column].dtype == object:
            unique_count = optimized[column].nunique(dropna=False)
            if len(optimized) and unique_count <= min(10_000, max(100, len(optimized) // 20)):
                optimized[column] = optimized[column].astype("category")
                if "" not in optimized[column].cat.categories:
                    optimized[column] = optimized[column].cat.add_categories([""])
    return optimized


def cache_master_frames(frames: Dict[str, pd.DataFrame], run_id: str) -> None:
    cache_dir = Path("cache") / run_id
    cache_dir.mkdir(parents=True, exist_ok=True)
    for dataset, frame in frames.items():
        frame.to_parquet(cache_dir / f"{dataset.lower()}.parquet", index=False)


def load_cached_master_frames(frames: Dict[str, pd.DataFrame], run_id: str) -> Dict[str, pd.DataFrame]:
    cached_frames: Dict[str, pd.DataFrame] = {}
    for dataset in frames:
        cache_path = Path("cache") / run_id / f"{dataset.lower()}.parquet"
        cached_frames[dataset] = pd.read_parquet(cache_path)
    return cached_frames


def build_ingestion_stats(
    dataset: str,
    frame: pd.DataFrame,
    loaded_files: int,
    discovered_files: int,
    failed_files: list[Dict[str, str]],
) -> Dict[str, object]:
    rrn_column = find_column(frame, str(DATASETS[dataset]["rrn_column"]), dataset)
    normalized_rrns = normalize_rrn_series(frame[rrn_column])
    valid_mask = is_valid_rrn_series(normalized_rrns)
    stats_frame = frame.copy()
    stats_frame["NORMALIZED_RRN"] = normalized_rrns
    duplicate_mask = valid_mask & stats_frame.duplicated("NORMALIZED_RRN", keep="first")
    return {
        "files_discovered": discovered_files,
        "files_loaded": loaded_files,
        "files_failed": len(failed_files),
        "failed_files": failed_files,
        "rows_loaded": len(frame),
        "rows_consolidated": len(frame),
        "rows_after_cleaning": len(frame),
        "rows_after_deduplication": int(len(frame) - duplicate_mask.sum()),
        "unique_rrns": int(normalized_rrns.loc[valid_mask].nunique(dropna=True)),
    }




def run_reconciliation(raw_frames: Dict[str, pd.DataFrame], ingestion_stats: Dict[str, Dict[str, object]]) -> None:
    progress = st.progress(0, text="Normalizing RRNs...")
    try:
        progress.progress(0.3, text="Comparing transaction sets...")
        results = reconcile(raw_frames)
        results["ingestion_stats"] = ingestion_stats
        progress.progress(0.9, text="Preparing preview...")
        st.session_state["reconciliation_results"] = results
        progress.progress(1.0, text="Done")
        st.success("Reconciliation completed successfully.")
    except ReconciliationError as exc:
        st.error(str(exc))
    except Exception as exc:
        st.exception(exc)
    finally:
        progress.empty()


def render_ingestion_preview(stats: Dict[str, Dict[str, object]]) -> None:
    st.subheader("Upload Preview")
    failed_total = sum(int(stats[dataset].get("files_failed", 0)) for dataset in DATASETS)
    ignored_total = sum(len(stats[dataset].get("ignored_files", [])) for dataset in DATASETS)
    if failed_total:
        st.error(f"{failed_total:,} file(s) failed to load. Review the Debug Panel before reconciling.")
    if ignored_total:
        st.warning(f"{ignored_total:,} file(s) were ignored, including empty, hidden, temporary, unsupported, or duplicate-content files.")

    preview_rows = []
    for dataset in DATASETS:
        item = stats[dataset]
        preview_rows.append(
            {
                "Source": display_dataset_label(dataset),
                "Total Files Loaded": item["files_loaded"],
                "Total Rows Loaded": item["rows_loaded"],
                "Rows After Deduplication": item["rows_after_deduplication"],
            }
        )
    st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

    with st.expander("Debug Panel", expanded=False):
        debug_rows = []
        for dataset in DATASETS:
            item = stats[dataset]
            debug_rows.append(
                {
                    "Source": display_dataset_label(dataset),
                    "Files Discovered": item["files_discovered"],
                    "Files Loaded": item["files_loaded"],
                    "Files Failed": item["files_failed"],
                    "Rows Loaded": item["rows_loaded"],
                    "Rows Consolidated": item["rows_consolidated"],
                    "Rows After Cleaning": item["rows_after_cleaning"],
                    "Rows After Deduplication": item["rows_after_deduplication"],
                    "Unique RRNs": item["unique_rrns"],
                }
            )
        st.dataframe(pd.DataFrame(debug_rows), use_container_width=True, hide_index=True)
        for dataset in DATASETS:
            failed_files = stats[dataset].get("failed_files", [])
            if failed_files:
                st.markdown(f"**{display_dataset_label(dataset)} failed files**")
                st.dataframe(pd.DataFrame(failed_files), use_container_width=True, hide_index=True)
            ignored_files = stats[dataset].get("ignored_files", [])
            if ignored_files:
                st.markdown(f"**{display_dataset_label(dataset)} ignored files**")
                st.dataframe(pd.DataFrame(ignored_files), use_container_width=True, hide_index=True)


def display_dataset_label(dataset: str) -> str:
    return {
        "INTERSWITCH": "Interswitch",
        "CASH234": "Cash234",
        "ZENITH": "Zenith",
    }[dataset]


def render_results(results: Dict[str, object]) -> None:
    result_sets: Dict[str, pd.DataFrame] = results["sets"]  # type: ignore[assignment]
    summary: pd.DataFrame = results["summary"]  # type: ignore[assignment]
    report_frames: Dict[str, pd.DataFrame] = results["report_frames"]  # type: ignore[assignment]
    reports: Dict[str, bytes] = results.setdefault("reports", {})  # type: ignore[assignment]
    prepared_downloads: Dict[str, bytes] = results.setdefault("prepared_downloads", {})  # type: ignore[assignment]
    duplicates_filename = str(results["duplicates_filename"])

    st.subheader("Summary")
    render_metric_cards(summary, result_sets)
    render_charts(result_sets)

    st.divider()
    st.subheader("Downloads")
    st.caption("Reconciliation is ready. Large Excel files are generated only when requested so the app stays responsive.")
    for row_start in range(0, len(REPORT_DOWNLOAD_ORDER), 3):
        cols = st.columns(3)
        for col, filename in zip(cols, REPORT_DOWNLOAD_ORDER[row_start : row_start + 3]):
            with col:
                label = filename.replace("_", " ").removesuffix(".xlsx")
                if filename not in reports:
                    if st.button(f"Prepare {label}", key=f"prepare_{filename}", use_container_width=True):
                        with st.spinner(f"Preparing {label}..."):
                            reports[filename] = build_required_report_workbook(report_frames, filename)
                            st.session_state["reconciliation_results"] = results
                        st.rerun()
                else:
                    st.download_button(
                        label,
                        data=reports[filename],
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{filename}",
                        on_click="ignore",
                        use_container_width=True,
                    )
    lazy_downloads = [
        ("duplicates", "Duplicate RRN Rows", duplicates_filename, lambda: duplicates_workbook_to_bytes(results)),
        (
            "invalid_rrns",
            "Invalid RRN Rows",
            str(results["invalid_filename"]),
            lambda: invalid_rrn_workbook_to_bytes(results),
        ),
        (
            "interswitch_zenith_carryover",
            "Interswitch-Zenith Carryover for Tran_ID Matcher",
            "Interswitch_Zenith_Tran_Matcher_Carryover.xlsx",
            lambda: interswitch_zenith_carryover_workbook_to_bytes(results),
        ),
    ]
    for download_key, label, filename, builder in lazy_downloads:
        if download_key not in prepared_downloads:
            if st.button(f"Prepare {label}", key=f"prepare_{download_key}", use_container_width=True):
                with st.spinner(f"Preparing {label}..."):
                    prepared_downloads[download_key] = builder()
                    st.session_state["reconciliation_results"] = results
                st.rerun()
        else:
            st.download_button(
                label,
                data=prepared_downloads[download_key],
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{download_key}",
                on_click="ignore",
                use_container_width=True,
            )
    updated_cols = st.columns(3)
    for col, dataset in zip(updated_cols, DATASETS):
        with col:
            label = f"Updated {display_dataset_label(dataset)}"
            download_key = f"updated_{dataset.lower()}"
            if download_key not in prepared_downloads:
                if st.button(f"Prepare {label}", key=f"prepare_{download_key}", use_container_width=True):
                    with st.spinner(f"Preparing {label}..."):
                        prepared_downloads[download_key] = updated_source_workbook_to_bytes(results, dataset)
                        st.session_state["reconciliation_results"] = results
                    st.rerun()
            else:
                st.download_button(
                    label,
                    data=prepared_downloads[download_key],
                    file_name=updated_export_filename(dataset),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{download_key}",
                    on_click="ignore",
                    use_container_width=True,
                )

    st.divider()
    st.subheader("Preview Result Sets")
    sheet_name = st.selectbox("Choose result set", list(result_sets.keys()))
    selected = result_sets[sheet_name]
    render_preview(sheet_name, selected)


def render_metric_cards(summary: pd.DataFrame, result_sets: Dict[str, pd.DataFrame]) -> None:
    metrics = {
        row["Metric"]: row["Value"]
        for _, row in summary.iterrows()
        if row["Metric"] and row["Metric"] != "Venn Set"
    }
    cols = st.columns(4)
    display = [
        "Total INTERSWITCH transactions",
        "Total CASH234 transactions",
        "Total ZENITH transactions",
        "Total unique RRNs",
        "Blank Zenith rows used for fallback matching",
        "Total matched across all 3",
        "Total mismatches",
        "Total invalid RRNs",
        "Total duplicate RRN rows removed",
    ]
    for idx, metric in enumerate(display):
        with cols[idx % 4]:
            st.metric(metric, f"{int(metrics.get(metric, 0)):,}")

    st.dataframe(
        pd.DataFrame(
            {
                "Result Set": list(result_sets.keys()),
                "Rows": [len(frame) for frame in result_sets.values()],
            }
        ),
        use_container_width=True,
        hide_index=True,
    )


def render_charts(result_sets: Dict[str, pd.DataFrame]) -> None:
    chart_data = pd.DataFrame(
        {
            "Set": list(result_sets.keys()),
            "Rows": [len(frame) for frame in result_sets.values()],
        }
    )
    col1, col2 = st.columns([1.2, 1])
    with col1:
        fig = go.Figure(
            data=[
                go.Bar(
                    x=chart_data["Rows"],
                    y=chart_data["Set"],
                    orientation="h",
                    marker_color=[
                        color_for_sheet(sheet)
                        for sheet in chart_data["Set"]
                    ],
                )
            ]
        )
        fig.update_layout(height=430, margin=dict(l=10, r=10, t=20, b=10), yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig, use_container_width=True)
    with col2:
        render_venn_visual(result_sets)


def render_venn_visual(result_sets: Dict[str, pd.DataFrame]) -> None:
    fig = go.Figure()
    circles = [
        ("INTERSWITCH", 0.38, 0.58, "rgba(36, 99, 235, 0.25)"),
        ("CASH234", 0.62, 0.58, "rgba(245, 158, 11, 0.25)"),
        ("ZENITH", 0.50, 0.36, "rgba(22, 163, 74, 0.25)"),
    ]
    for name, x, y, color in circles:
        fig.add_shape(type="circle", xref="paper", yref="paper", x0=x - 0.24, y0=y - 0.24, x1=x + 0.24, y1=y + 0.24, fillcolor=color, line_color=color)
        fig.add_annotation(x=x, y=y + 0.25, text=name, showarrow=False, font=dict(size=12))
    annotations = [
        ("INTERSWITCH_ONLY", 0.27, 0.63),
        ("CASH234_ONLY", 0.73, 0.63),
        ("ZENITH_ONLY", 0.50, 0.22),
        ("INTERSWITCH_AND_CASH234_ONLY", 0.50, 0.65),
        ("INTERSWITCH_AND_ZENITH_ONLY", 0.40, 0.42),
        ("CASH234_AND_ZENITH_ONLY", 0.60, 0.42),
        ("MATCHED_IN_ALL_THREE", 0.50, 0.51),
    ]
    for sheet, x, y in annotations:
        fig.add_annotation(x=x, y=y, text=f"{len(result_sets[sheet]):,}", showarrow=False, font=dict(size=13, color="#111827"))
    fig.update_xaxes(visible=False, range=[0, 1])
    fig.update_yaxes(visible=False, range=[0, 1])
    fig.update_layout(height=430, margin=dict(l=10, r=10, t=20, b=10), plot_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig, use_container_width=True)


def render_preview(sheet_name: str, df: pd.DataFrame) -> None:
    color = color_for_sheet(sheet_name)
    st.markdown(f"<div class='status-strip' style='border-left-color:{color}'><strong>{sheet_name}</strong>: {len(df):,} rows</div>", unsafe_allow_html=True)
    search = st.text_input("Search this result set", key=f"search_{sheet_name}")
    filtered = filter_dataframe(df, search)
    preview_filtered = filtered.drop(columns=[col for col in GAP_COLUMNS if col in filtered.columns])
    preview_filtered = make_unique_display_columns(preview_filtered)
    st.caption(f"Showing {len(filtered):,} of {len(df):,} rows")

    if AgGrid and GridOptionsBuilder and GridUpdateMode:
        grid_options = GridOptionsBuilder.from_dataframe(preview_filtered.head(100000))
        grid_options.configure_default_column(filter=True, sortable=True, resizable=True)
        grid_options.configure_pagination(enabled=True, paginationAutoPageSize=False, paginationPageSize=50)
        grid_options.configure_side_bar()
        AgGrid(
            preview_filtered,
            gridOptions=grid_options.build(),
            height=520,
            update_mode=GridUpdateMode.NO_UPDATE,
            fit_columns_on_grid_load=False,
            allow_unsafe_jscode=False,
            theme="streamlit",
        )
    else:
        page_size = st.selectbox("Rows per page", [25, 50, 100, 250, 500], index=1)
        page_count = max((len(filtered) - 1) // page_size + 1, 1)
        page = st.number_input("Page", min_value=1, max_value=page_count, value=1)
        start = (page - 1) * page_size
        st.dataframe(preview_filtered.iloc[start : start + page_size], use_container_width=True, height=520)


def make_unique_display_columns(df: pd.DataFrame) -> pd.DataFrame:
    display_df = df.copy()
    seen: Dict[str, int] = {}
    columns = []
    for column in display_df.columns:
        name = str(column)
        count = seen.get(name, 0)
        seen[name] = count + 1
        columns.append(name if count == 0 else f"{name} ({count + 1})")
    display_df.columns = columns
    return display_df


def color_for_sheet(sheet_name: str) -> str:
    if sheet_name == "MATCHED_IN_ALL_THREE":
        return STATUS_COLORS["MATCHED_IN_ALL_THREE"]
    if sheet_name == INVALID_SHEET_NAME:
        return STATUS_COLORS["INVALID_OR_BLANK_RRN"]
    if sheet_name.endswith("_ONLY") and "_AND_" not in sheet_name:
        return STATUS_COLORS["ONLY"]
    return STATUS_COLORS["PARTIAL"]


def render_empty_state() -> None:
    st.info("Upload the INTERSWITCH, CASH234, and ZENITH files from the sidebar to begin.")
    st.write(
        "The app compares normalized RRN values, removes duplicate valid RRNs per uploaded file, and keeps removed duplicates in a separate workbook."
    )


def inject_css() -> None:
    st.markdown(
        """
        <style>
        .status-strip {
            padding: 0.75rem 1rem;
            border-left: 0.45rem solid;
            background: rgba(127, 127, 127, 0.08);
            border-radius: 6px;
            margin-bottom: 0.75rem;
        }
        [data-testid="stMetricValue"] {
            font-size: 1.45rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
