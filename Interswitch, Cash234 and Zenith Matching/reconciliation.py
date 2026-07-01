from __future__ import annotations

import io
import re
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from typing import BinaryIO, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd

from matching import build_interswitch_zenith_description_links as match_interswitch_zenith_description_links
from source_processing import (
    expand_scientific_rrn,
    remove_hidden_spaces,
)


DATASETS = {
    "INTERSWITCH": {
        "rrn_column": "Retrieval_Reference_Nr",
        "date_column": "Local_Date_Time",
        "date_dayfirst": False,
        "amount_columns": ("Amount", "Transaction Amount", "Tran_Amount", "Tran Amount", "Transaction_Amount"),
        "present_column": "Present_In_INTERSWITCH",
    },
    "CASH234": {
        "rrn_column": "R R N",
        "date_column": "Transaction Date",
        "date_dayfirst": False,
        "amount_columns": ("Amount", "Transaction Amount", "Tran Amount", "Transaction_Amount"),
        "present_column": "Present_In_CASH234",
    },
    "ZENITH": {
        "rrn_column": "RRN",
        "date_column": "EffectiveDate",
        "date_dayfirst": True,
        "amount_columns": ("Amount", "Transaction Amount", "Tran Amount", "Transaction_Amount"),
        "present_column": "Present_In_ZENITH",
    },
}
DATE_PRIORITY = ("CASH234", "INTERSWITCH", "ZENITH")


SET_DEFINITIONS = [
    ("INTERSWITCH_ONLY", True, False, False, "Present only in INTERSWITCH"),
    ("CASH234_ONLY", False, True, False, "Present only in CASH234"),
    ("ZENITH_ONLY", False, False, True, "Present only in ZENITH"),
    ("INTERSWITCH_AND_CASH234_ONLY", True, True, False, "Present in INTERSWITCH and CASH234 only"),
    ("INTERSWITCH_AND_ZENITH_ONLY", True, False, True, "Present in INTERSWITCH and ZENITH only"),
    ("CASH234_AND_ZENITH_ONLY", False, True, True, "Present in CASH234 and ZENITH only"),
    ("MATCHED_IN_ALL_THREE", True, True, True, "Matched in all three"),
]


PAIRWISE_COMPLEMENT_DEFINITIONS = [
    (
        "NOT_INTERSWITCH_AND_CASH234",
        ("INTERSWITCH", "CASH234"),
        "Complement of INTERSWITCH and CASH234 intersection",
    ),
    (
        "NOT_INTERSWITCH_AND_ZENITH",
        ("INTERSWITCH", "ZENITH"),
        "Complement of INTERSWITCH and ZENITH intersection",
    ),
    (
        "NOT_CASH234_AND_ZENITH",
        ("CASH234", "ZENITH"),
        "Complement of CASH234 and ZENITH intersection",
    ),
]
PAIRWISE_COMPLEMENT_SHEETS = {definition[0] for definition in PAIRWISE_COMPLEMENT_DEFINITIONS}


INVALID_SHEET_NAME = "INVALID_OR_BLANK_RRN"
REPORT_DOWNLOAD_ORDER = [
    "Cash234_vs_Zenith_Reconciliation.xlsx",
    "Cash234_vs_Interswitch_Reconciliation.xlsx",
    "Interswitch_vs_Zenith_Reconciliation.xlsx",
    "Unmatched_Transactions.xlsx",
    "Interswitch_Unsettled.xlsx",
    "Zenith_Unsettled.xlsx",
]
GAP_AFTER_INTERSWITCH = "__GAP_AFTER_INTERSWITCH"
GAP_AFTER_CASH234 = "__GAP_AFTER_CASH234"
GAP_AFTER_ZENITH = "__GAP_AFTER_ZENITH"
GAP_COLUMNS = {GAP_AFTER_INTERSWITCH, GAP_AFTER_CASH234, GAP_AFTER_ZENITH}
EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384
WIDTH_SAMPLE_ROWS = 2_000


@dataclass(frozen=True)
class LoadedDataset:
    name: str
    data: pd.DataFrame
    report_data: pd.DataFrame
    duplicates: pd.DataFrame
    original_rows: int
    valid_rows: int
    invalid_rows: int
    duplicate_rows: int
    duplicate_rrn_rows: int
    blank_zenith_fallback_rows: int = 0


class ReconciliationError(ValueError):
    """Raised when an input file cannot be reconciled safely."""


def make_output_filename() -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"Transaction_Reconciliation_Output_{timestamp}.xlsx"


def make_duplicates_filename() -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"Duplicates_{timestamp}.xlsx"


def make_invalid_rrn_filename() -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"Invalid_RRN_Rows_{timestamp}.xlsx"


def updated_export_filename(dataset: str) -> str:
    return f"updated_{dataset.lower()}.xlsx"


def list_sheets(file: BinaryIO, filename: str) -> List[str]:
    suffix = filename.lower().rsplit(".", 1)[-1]
    file.seek(0)
    if suffix == "csv":
        return ["CSV"]
    try:
        excel = pd.ExcelFile(file)
        return excel.sheet_names or ["Sheet1"]
    finally:
        file.seek(0)


def read_transaction_file(file: BinaryIO, filename: str, sheet_name: Optional[str]) -> pd.DataFrame:
    suffix = filename.lower().rsplit(".", 1)[-1]
    file.seek(0)
    if suffix == "csv":
        df = pd.read_csv(file, dtype=str, keep_default_na=False, na_filter=False)
    elif suffix in {"xlsx", "xls"}:
        df = pd.read_excel(
            file,
            sheet_name=sheet_name,
            dtype=str,
            keep_default_na=False,
            engine=None,
        )
    else:
        raise ReconciliationError(f"Unsupported file type for {filename}. Upload .xlsx, .xls, or .csv.")
    file.seek(0)
    return clean_headers(df)


def clean_headers(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    columns = []
    seen: Dict[str, int] = {}
    for idx, col in enumerate(df.columns):
        cleaned = normalize_header(str(col)) if str(col).strip() else f"Unnamed_{idx + 1}"
        count = seen.get(cleaned, 0)
        seen[cleaned] = count + 1
        columns.append(cleaned if count == 0 else f"{cleaned}_{count + 1}")
    df.columns = columns
    text_view = df.astype("string").fillna("").apply(lambda col: col.str.strip())
    return df.loc[~text_view.eq("").all(axis=1)].copy()


def normalize_header(value: str) -> str:
    value = remove_hidden_spaces(value).strip()
    return re.sub(r"\s+", " ", value)


def canonical_header(value: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", remove_hidden_spaces(value).upper())


def find_column(df: pd.DataFrame, required_column: str, dataset_name: str) -> str:
    desired = canonical_header(required_column)
    exact_matches = [col for col in df.columns if canonical_header(col) == desired]
    if exact_matches:
        return exact_matches[0]
    available = ", ".join(map(str, df.columns[:25]))
    raise ReconciliationError(
        f"{dataset_name} is missing required column '{required_column}'. "
        f"Available columns include: {available}"
    )


def find_first_existing_column(
    df: pd.DataFrame,
    candidate_columns: Sequence[str],
    dataset_name: str,
    purpose: str,
) -> str:
    for candidate in candidate_columns:
        desired = canonical_header(candidate)
        exact_matches = [col for col in df.columns if canonical_header(col) == desired]
        if exact_matches:
            return exact_matches[0]

    amount_like = [
        col
        for col in df.columns
        if "AMOUNT" in canonical_header(str(col)) or canonical_header(str(col)) in {"DEBIT", "CREDIT"}
    ]
    if len(amount_like) == 1:
        return amount_like[0]

    available = ", ".join(map(str, df.columns[:25]))
    expected = ", ".join(candidate_columns)
    raise ReconciliationError(
        f"{dataset_name} is missing required {purpose} column. Expected one of: {expected}. "
        f"Available columns include: {available}"
    )


def normalize_rrn_series(series: pd.Series) -> pd.Series:
    normalized = series.fillna("").map(remove_hidden_spaces).astype("string")
    normalized = normalized.str.strip().str.upper()
    normalized = normalized.str.replace(r"\s+", "", regex=True)
    normalized = normalized.str.replace(r"\.0$", "", regex=True)
    normalized = normalized.map(expand_scientific_rrn)
    return normalized.fillna("")


def normalize_amount_series(series: pd.Series) -> pd.Series:
    cleaned = series.fillna("").map(remove_hidden_spaces).astype("string").str.strip()
    cleaned = cleaned.str.replace(",", "", regex=False)
    cleaned = cleaned.str.replace(r"^[A-Z]{3}\s*", "", regex=True)
    cleaned = cleaned.str.replace(r"[^0-9.\-()]", "", regex=True)
    return cleaned.map(normalize_amount_value).astype("string").fillna("")


def normalize_amount_value(value: object) -> str:
    text = "" if pd.isna(value) else str(value).strip()
    if not text:
        return ""
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    try:
        amount = Decimal(text)
    except InvalidOperation:
        return text.upper()
    if negative:
        amount = -amount
    normalized = amount.normalize()
    if normalized == normalized.to_integral_value():
        return format(normalized.quantize(Decimal("1")), "f")
    return format(normalized, "f")


def normalize_date_key_series(series: pd.Series, dayfirst: bool) -> pd.Series:
    parsed = parse_sort_date_series(series, dayfirst=dayfirst)
    formatted = parsed.dt.strftime("%Y-%m-%d").astype("string")
    fallback = series.fillna("").map(remove_hidden_spaces).astype("string").str.strip().str.upper()
    fallback = fallback.mask(fallback.isin({"", "NAN", "NONE", "NULL", "N/A", "NA", "-", "--"}), "")
    return formatted.fillna(fallback).fillna("")


def is_valid_rrn_series(series: pd.Series) -> pd.Series:
    stripped = series.fillna("").astype("string").str.strip()
    blank = stripped.eq("")
    invalid_tokens = stripped.str.upper().isin({"NAN", "NONE", "NULL", "N/A", "NA", "-", "--"})
    malformed = ~stripped.str.match(r"^[A-Z0-9]+$", na=False)
    return ~(blank | invalid_tokens | malformed)


def prepare_dataset(dataset_name: str, df: pd.DataFrame) -> LoadedDataset:
    config = DATASETS[dataset_name]
    rrn_column = find_column(df, config["rrn_column"], dataset_name)
    find_column(df, config["date_column"], dataset_name)
    working = df.copy()
    working["SOURCE_DATASET"] = dataset_name
    working["SOURCE_ROW_NUMBER"] = range(2, len(working) + 2)
    working["NORMALIZED_RRN"] = normalize_rrn_series(working[rrn_column])
    working["RRN_VALID"] = is_valid_rrn_series(working["NORMALIZED_RRN"])
    zenith_blank_rrn_mask = pd.Series(False, index=working.index)
    if dataset_name == "ZENITH":
        zenith_blank_rrn_mask = working["NORMALIZED_RRN"].fillna("").astype("string").str.strip().eq("")

    duplicate_counts = working.groupby("NORMALIZED_RRN", dropna=False)["NORMALIZED_RRN"].transform("size")
    valid_mask = working["RRN_VALID"]
    duplicate_counts = duplicate_counts.where(valid_mask, 0).astype(int)
    working[f"{dataset_name}_DUPLICATE_COUNT"] = duplicate_counts
    working[f"{dataset_name}_DUPLICATE_FLAG"] = duplicate_counts.gt(1)

    duplicate_mask = valid_mask & working.duplicated("NORMALIZED_RRN", keep="first")
    duplicates = working.loc[duplicate_mask].copy()
    report_mask = valid_mask
    if dataset_name == "ZENITH":
        report_mask = report_mask | zenith_blank_rrn_mask
    report_rows = working.loc[report_mask & ~duplicate_mask].copy()
    matching_rows = working.loc[~duplicate_mask].copy()

    prefixed = prefix_source_columns(matching_rows, dataset_name)
    prefixed_report = prefix_source_columns(report_rows, dataset_name)
    prefixed_duplicates = prefix_source_columns(duplicates, dataset_name)
    return LoadedDataset(
        name=dataset_name,
        data=prefixed,
        report_data=prefixed_report,
        duplicates=prefixed_duplicates,
        original_rows=len(df),
        valid_rows=len(report_rows),
        invalid_rows=int((~working["RRN_VALID"] & ~zenith_blank_rrn_mask).sum()),
        duplicate_rows=int(duplicate_mask.sum()),
        duplicate_rrn_rows=int(duplicate_mask.sum()),
        blank_zenith_fallback_rows=int(zenith_blank_rrn_mask.sum()),
    )


def build_exact_duplicate_mask(df: pd.DataFrame) -> pd.Series:
    ignored_columns = {
        "SOURCE_DATASET",
        "SOURCE_ROW_NUMBER",
        "RRN_VALID",
    }
    duplicate_columns = [
        column
        for column in df.columns
        if column not in ignored_columns
        and not str(column).endswith("_DUPLICATE_COUNT")
        and not str(column).endswith("_DUPLICATE_FLAG")
    ]
    normalized = df.loc[:, duplicate_columns].astype("string").fillna("")
    normalized = normalized.apply(lambda col: col.map(remove_hidden_spaces).str.strip())
    return normalized.duplicated(keep="first")


def prefix_source_columns(df: pd.DataFrame, dataset_name: str) -> pd.DataFrame:
    helper_columns = {
        "NORMALIZED_RRN",
        "RRN_VALID",
        f"{dataset_name}_DUPLICATE_COUNT",
        f"{dataset_name}_DUPLICATE_FLAG",
    }
    renamed = {}
    for col in df.columns:
        if col in helper_columns:
            continue
        if not str(col).startswith(f"{dataset_name}_"):
            renamed[col] = f"{dataset_name}_{col}"
    return df.rename(columns=renamed)


def reconcile(raw_frames: Dict[str, pd.DataFrame]) -> Dict[str, object]:
    loaded = {name: prepare_dataset(name, raw_frames[name]) for name in DATASETS}
    valid_frames = {
        name: ds.data.loc[ds.data["RRN_VALID"]].drop(columns=["RRN_VALID"])
        for name, ds in loaded.items()
    }
    if "ZENITH" in loaded:
        zenith_blank_rrn = loaded["ZENITH"].data["ZENITH_RRN"].fillna("").astype("string").str.strip().eq("")
        valid_frames["ZENITH"] = pd.concat(
            [
                valid_frames["ZENITH"],
                loaded["ZENITH"].data.loc[zenith_blank_rrn].drop(columns=["RRN_VALID"]),
            ],
            ignore_index=False,
            sort=False,
        )
    invalid_frames = [
        ds.data.loc[build_invalid_row_mask(name, ds.data)].drop(columns=["RRN_VALID"])
        for name, ds in loaded.items()
    ]

    unique_presence = build_presence_table(valid_frames)
    result_sets = build_result_sets(unique_presence, valid_frames)
    result_sets.update(build_pairwise_complement_result_sets(unique_presence, valid_frames))
    pairwise_exports = build_pairwise_export_sets(unique_presence, valid_frames, result_sets)
    result_sets[INVALID_SHEET_NAME] = build_invalid_sheet(invalid_frames)
    updated_exports = build_updated_source_exports(raw_frames, valid_frames, unique_presence)

    summary = build_summary(loaded, unique_presence, result_sets)
    statistics = build_statistics(loaded, result_sets)
    return {
        "loaded": loaded,
        "sets": result_sets,
        "pairwise_exports": pairwise_exports,
        "report_frames": valid_frames,
        "reports": {},
        "duplicates": {
            name: remove_dataset_prefixes_from_headers(sort_dataset_frame_by_own_date(ds.duplicates, name))
            for name, ds in loaded.items()
        },
        "updated_exports": updated_exports,
        "summary": summary,
        "statistics": statistics,
        "filename": make_output_filename(),
        "duplicates_filename": make_duplicates_filename(),
        "invalid_filename": make_invalid_rrn_filename(),
    }


def build_invalid_row_mask(dataset: str, frame: pd.DataFrame) -> pd.Series:
    invalid_mask = ~frame["RRN_VALID"]
    if dataset != "ZENITH":
        return invalid_mask
    return invalid_mask & ~frame["ZENITH_RRN"].fillna("").astype("string").str.strip().eq("")


def build_presence_table(valid_frames: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    all_rrns = pd.Index([], dtype="object")
    for frame in valid_frames.values():
        normalized_rrns = frame["NORMALIZED_RRN"].dropna().astype("string")
        normalized_rrns = normalized_rrns.loc[normalized_rrns.str.strip().ne("")]
        all_rrns = all_rrns.union(pd.Index(normalized_rrns.unique()))
    presence = pd.DataFrame({"NORMALIZED_RRN": all_rrns})
    for name, config in DATASETS.items():
        normalized_rrns = valid_frames[name]["NORMALIZED_RRN"].dropna().astype("string")
        normalized_rrns = normalized_rrns.loc[normalized_rrns.str.strip().ne("")]
        rrns = pd.Index(normalized_rrns.unique())
        presence[config["present_column"]] = presence["NORMALIZED_RRN"].isin(rrns)
    return presence


def build_interswitch_zenith_linked_frames(valid_frames: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    linked_frames = {name: frame.copy() for name, frame in valid_frames.items()}
    interswitch_frame = linked_frames.get("INTERSWITCH")
    zenith_frame = linked_frames.get("ZENITH")
    if interswitch_frame is None or zenith_frame is None or interswitch_frame.empty or zenith_frame.empty:
        return linked_frames

    interswitch_rrns = pd.Index(interswitch_frame["NORMALIZED_RRN"].dropna().unique())
    zenith_rrns = pd.Index(zenith_frame["NORMALIZED_RRN"].dropna().unique())
    already_settled_mask = interswitch_frame["NORMALIZED_RRN"].isin(zenith_rrns)
    already_matched_zenith_mask = zenith_frame["NORMALIZED_RRN"].isin(interswitch_rrns)
    links = build_interswitch_zenith_description_links(
        primary_frame=interswitch_frame,
        zenith_frame=zenith_frame,
        already_settled_mask=already_settled_mask,
        already_matched_zenith_mask=already_matched_zenith_mask,
    )
    if not links:
        return linked_frames

    linked_zenith_rows = []
    linked_zenith_indices = []
    for interswitch_index, zenith_index in links:
        linked_row = zenith_frame.loc[[zenith_index]].copy()
        linked_row["NORMALIZED_RRN"] = interswitch_frame.at[interswitch_index, "NORMALIZED_RRN"]
        linked_zenith_rows.append(linked_row)
        linked_zenith_indices.append(zenith_index)

    linked_frames["ZENITH"] = pd.concat(
        [zenith_frame.drop(index=pd.Index(linked_zenith_indices).unique(), errors="ignore"), *linked_zenith_rows],
        ignore_index=False,
        sort=False,
    )
    return linked_frames


def build_result_sets(
    unique_presence: pd.DataFrame,
    valid_frames: Dict[str, pd.DataFrame],
) -> Dict[str, pd.DataFrame]:
    output: Dict[str, pd.DataFrame] = {}
    for sheet_name, in_i, in_c, in_z, status in SET_DEFINITIONS:
        rrns = unique_presence.loc[
            (unique_presence["Present_In_INTERSWITCH"] == in_i)
            & (unique_presence["Present_In_CASH234"] == in_c)
            & (unique_presence["Present_In_ZENITH"] == in_z),
            "NORMALIZED_RRN",
        ]
        result = expand_rows_for_rrns(rrns, valid_frames)
        result["Present_In_INTERSWITCH"] = in_i
        result["Present_In_CASH234"] = in_c
        result["Present_In_ZENITH"] = in_z
        result["MATCH_STATUS"] = status
        output[sheet_name] = order_result_columns(result, include_gaps=True)
    return output


def build_pairwise_complement_result_sets(
    unique_presence: pd.DataFrame,
    valid_frames: Dict[str, pd.DataFrame],
) -> Dict[str, pd.DataFrame]:
    output: Dict[str, pd.DataFrame] = {}
    for sheet_name, dataset_pair, status in PAIRWISE_COMPLEMENT_DEFINITIONS:
        presence_columns = [DATASETS[dataset]["present_column"] for dataset in dataset_pair]
        pair_presence = unique_presence.loc[:, presence_columns]
        mask = pair_presence.any(axis=1) & ~pair_presence.all(axis=1)
        matching_presence = unique_presence.loc[mask, ["NORMALIZED_RRN", *presence_columns]].copy()
        result = expand_rows_for_rrns(matching_presence["NORMALIZED_RRN"], valid_frames, dataset_pair)
        result = result.merge(matching_presence, how="left", on="NORMALIZED_RRN")
        result["MATCH_STATUS"] = status
        output[sheet_name] = order_result_columns(result, include_gaps=True)
    return output


def build_pairwise_export_sets(
    unique_presence: pd.DataFrame,
    valid_frames: Dict[str, pd.DataFrame],
    result_sets: Dict[str, pd.DataFrame],
) -> Dict[str, Dict[str, pd.DataFrame]]:
    exports: Dict[str, Dict[str, pd.DataFrame]] = {}
    for sheet_name, dataset_pair, _status in PAIRWISE_COMPLEMENT_DEFINITIONS:
        first_dataset, second_dataset = dataset_pair
        first_presence = DATASETS[first_dataset]["present_column"]
        second_presence = DATASETS[second_dataset]["present_column"]

        first_only_rrns = unique_presence.loc[
            unique_presence[first_presence] & ~unique_presence[second_presence],
            "NORMALIZED_RRN",
        ]
        second_only_rrns = unique_presence.loc[
            unique_presence[second_presence] & ~unique_presence[first_presence],
            "NORMALIZED_RRN",
        ]

        exports[sheet_name] = {
            "COMBINED": result_sets[sheet_name],
            first_dataset: build_single_dataset_pairwise_export(
                first_only_rrns,
                valid_frames,
                first_dataset,
                second_dataset,
            ),
            second_dataset: build_single_dataset_pairwise_export(
                second_only_rrns,
                valid_frames,
                second_dataset,
                first_dataset,
            ),
        }
    return exports


def build_updated_source_exports(
    raw_frames: Dict[str, pd.DataFrame],
    valid_frames: Dict[str, pd.DataFrame],
    unique_presence: pd.DataFrame,
) -> Dict[str, pd.DataFrame]:
    matched_rrns = pd.Index(
        unique_presence.loc[
            unique_presence["Present_In_INTERSWITCH"]
            & unique_presence["Present_In_CASH234"]
            & unique_presence["Present_In_ZENITH"],
            "NORMALIZED_RRN",
        ].dropna()
    )
    exports: Dict[str, pd.DataFrame] = {}
    for dataset, raw_frame in raw_frames.items():
        matched_frame = valid_frames[dataset].loc[valid_frames[dataset]["NORMALIZED_RRN"].isin(matched_rrns)]
        source_row_column = f"{dataset}_SOURCE_ROW_NUMBER"
        if source_row_column not in matched_frame.columns:
            raise ReconciliationError(f"{dataset} is missing stable source row identifiers for updated export.")
        matched_source_rows = set(pd.to_numeric(matched_frame[source_row_column], errors="coerce").dropna().astype(int))
        raw_source_rows = pd.Series(range(2, len(raw_frame) + 2), index=raw_frame.index)
        exports[dataset] = raw_frame.loc[~raw_source_rows.isin(matched_source_rows)].copy()
    return exports


def build_composite_key_series(df: pd.DataFrame, dataset: str, prefixed: bool) -> pd.Series:
    config = DATASETS[dataset]
    if prefixed:
        rrn_column = "NORMALIZED_RRN"
        date_column = find_existing_prefixed_column(df, dataset, str(config["date_column"]))
        amount_column = find_existing_prefixed_column_any(
            df,
            dataset,
            tuple(str(column) for column in config["amount_columns"]),
        )
    else:
        rrn_column = find_column(df, str(config["rrn_column"]), dataset)
        date_column = find_column(df, str(config["date_column"]), dataset)
        amount_column = find_first_existing_column(
            df,
            tuple(str(column) for column in config["amount_columns"]),
            dataset,
            "amount",
        )

    if date_column is None or amount_column is None:
        raise ReconciliationError(
            f"{dataset} is missing date or amount data needed to build the RRN + amount + date exclusion key."
        )

    rrns = df[rrn_column].astype("string").fillna("") if prefixed else normalize_rrn_series(df[rrn_column])
    amounts = normalize_amount_series(df[amount_column])
    dates = normalize_date_key_series(df[date_column], dayfirst=bool(config["date_dayfirst"]))
    return rrns.astype("string") + "|" + amounts.astype("string") + "|" + dates.astype("string")


def build_single_dataset_pairwise_export(
    rrns: Iterable[str],
    valid_frames: Dict[str, pd.DataFrame],
    dataset: str,
    missing_dataset: str,
) -> pd.DataFrame:
    result = expand_rows_for_rrns(rrns, valid_frames, [dataset])
    for name, config in DATASETS.items():
        result[config["present_column"]] = name == dataset
    result["MATCH_STATUS"] = f"Present in {dataset}; not present in {missing_dataset}"
    return order_result_columns(result, include_gaps=False)


def expand_rows_for_rrns(
    rrns: Iterable[str],
    valid_frames: Dict[str, pd.DataFrame],
    datasets: Iterable[str] = DATASETS,
) -> pd.DataFrame:
    rrn_frame = pd.DataFrame({"NORMALIZED_RRN": pd.Series(list(rrns), dtype="string")})
    result = rrn_frame
    for name in datasets:
        result = result.merge(valid_frames[name], how="left", on="NORMALIZED_RRN")
    return result


def build_invalid_sheet(invalid_frames: List[pd.DataFrame]) -> pd.DataFrame:
    if not invalid_frames:
        return pd.DataFrame(columns=["NORMALIZED_RRN", "MATCH_STATUS"])
    aligned = pd.concat(invalid_frames, ignore_index=True, sort=False)
    aligned["Present_In_INTERSWITCH"] = aligned.get("INTERSWITCH_SOURCE_DATASET", "").eq("INTERSWITCH")
    aligned["Present_In_CASH234"] = aligned.get("CASH234_SOURCE_DATASET", "").eq("CASH234")
    aligned["Present_In_ZENITH"] = aligned.get("ZENITH_SOURCE_DATASET", "").eq("ZENITH")
    aligned["MATCH_STATUS"] = "Invalid, blank, or malformed RRN"
    return order_result_columns(aligned, include_gaps=True)


def order_result_columns(df: pd.DataFrame, include_gaps: bool = False) -> pd.DataFrame:
    df = sort_by_priority_date(df)
    helpers = [
        "NORMALIZED_RRN",
        "Present_In_INTERSWITCH",
        "Present_In_CASH234",
        "Present_In_ZENITH",
        "MATCH_STATUS",
    ]
    ordered = [col for col in helpers if col in df.columns]
    metadata = []

    dataset_gap_columns = {
        "INTERSWITCH": GAP_AFTER_INTERSWITCH,
        "CASH234": GAP_AFTER_CASH234,
        "ZENITH": GAP_AFTER_ZENITH,
    }
    for dataset in DATASETS:
        dataset_columns = [
            col
            for col in df.columns
            if col.startswith(f"{dataset}_")
            and not is_dataset_metadata_column(col, dataset)
        ]
        ordered.extend(dataset_columns)
        metadata.extend(
            [
                col
                for col in df.columns
                if col.startswith(f"{dataset}_") and is_dataset_metadata_column(col, dataset)
            ]
        )
        if include_gaps and dataset_columns:
            gap_column = dataset_gap_columns[dataset]
            df[gap_column] = ""
            ordered.append(gap_column)

    ordered.extend(metadata)
    ordered.extend([col for col in df.columns if col not in ordered and col not in GAP_COLUMNS])
    return remove_dataset_prefixes_from_headers(df.loc[:, ordered])


def sort_by_priority_date(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df

    sort_dates = []
    for dataset in DATE_PRIORITY:
        config = DATASETS[dataset]
        date_column = find_existing_prefixed_column(df, dataset, str(config["date_column"]))
        if date_column is None:
            sort_dates.append(pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]"))
            continue
        sort_dates.append(parse_sort_date_series(df[date_column], dayfirst=bool(config["date_dayfirst"])))

    priority_date = sort_dates[0]
    for candidate in sort_dates[1:]:
        priority_date = priority_date.fillna(candidate)

    sorted_df = df.copy()
    sorted_df["__SORT_PRIORITY_DATE"] = priority_date
    sorted_df["__SORT_ORIGINAL_POSITION"] = range(len(sorted_df))
    sorted_df = sorted_df.sort_values(
        by=["__SORT_PRIORITY_DATE", "NORMALIZED_RRN", "__SORT_ORIGINAL_POSITION"],
        ascending=[True, True, True],
        na_position="last",
        kind="mergesort",
    )
    return sorted_df.drop(columns=["__SORT_PRIORITY_DATE", "__SORT_ORIGINAL_POSITION"])


def sort_dataset_frame_by_own_date(df: pd.DataFrame, dataset: str) -> pd.DataFrame:
    if df.empty:
        return df
    config = DATASETS[dataset]
    date_column = find_existing_prefixed_column(df, dataset, str(config["date_column"]))
    if date_column is None:
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


def find_existing_prefixed_column(df: pd.DataFrame, dataset: str, source_column: str) -> Optional[str]:
    desired = canonical_header(source_column)
    prefix = f"{dataset}_"
    for column in df.columns:
        if isinstance(column, str) and column.startswith(prefix):
            unprefixed = column.removeprefix(prefix)
            if canonical_header(unprefixed) == desired:
                return column
    return None


def find_existing_prefixed_column_any(
    df: pd.DataFrame,
    dataset: str,
    source_columns: Sequence[str],
) -> Optional[str]:
    for source_column in source_columns:
        found = find_existing_prefixed_column(df, dataset, source_column)
        if found is not None:
            return found

    prefix = f"{dataset}_"
    amount_like = [
        column
        for column in df.columns
        if isinstance(column, str)
        and column.startswith(prefix)
        and (
            "AMOUNT" in canonical_header(column.removeprefix(prefix))
            or canonical_header(column.removeprefix(prefix)) in {"DEBIT", "CREDIT"}
        )
    ]
    if len(amount_like) == 1:
        return amount_like[0]
    return None


def parse_sort_date_series(series: pd.Series, dayfirst: bool) -> pd.Series:
    cleaned = series.fillna("").map(remove_hidden_spaces).astype("string").str.strip()
    cleaned = cleaned.mask(cleaned.str.upper().isin({"", "NAN", "NONE", "NULL", "N/A", "NA", "-", "--"}))
    parsed = pd.to_datetime(cleaned, errors="coerce", dayfirst=dayfirst, format="mixed")

    numeric_values = pd.to_numeric(cleaned, errors="coerce")
    max_excel_serial_supported = (datetime(2262, 4, 11) - datetime(1899, 12, 30)).days
    numeric_values = numeric_values.where(numeric_values.between(0, max_excel_serial_supported))
    numeric_dates = pd.to_datetime(
        numeric_values,
        errors="coerce",
        unit="D",
        origin="1899-12-30",
    )
    return parsed.fillna(numeric_dates)


def is_dataset_metadata_column(column: str, dataset: str) -> bool:
    return column in {
        f"{dataset}_SOURCE_DATASET",
        f"{dataset}_SOURCE_ROW_NUMBER",
        f"{dataset}_DUPLICATE_COUNT",
        f"{dataset}_DUPLICATE_FLAG",
    }


def remove_dataset_prefixes_from_headers(df: pd.DataFrame) -> pd.DataFrame:
    display_df = df.copy()
    display_df.columns = [remove_dataset_prefix(column) for column in display_df.columns]
    return display_df


def remove_dataset_prefix(column: object) -> object:
    if not isinstance(column, str) or column in GAP_COLUMNS:
        return column
    for dataset in DATASETS:
        prefix = f"{dataset}_"
        if column.startswith(prefix):
            return column.removeprefix(prefix)
    return column


def build_summary(
    loaded: Dict[str, LoadedDataset],
    unique_presence: pd.DataFrame,
    result_sets: Dict[str, pd.DataFrame],
) -> pd.DataFrame:
    total_valid_set_rows = sum(len(result_sets[name]) for name, *_ in SET_DEFINITIONS)
    total_mismatches = sum(
        len(result_sets[name]) for name, *_ in SET_DEFINITIONS if name != "MATCHED_IN_ALL_THREE"
    )
    rows = [
        ("Total INTERSWITCH transactions", loaded["INTERSWITCH"].original_rows, ""),
        ("Total CASH234 transactions", loaded["CASH234"].original_rows, ""),
        ("Total ZENITH transactions", loaded["ZENITH"].original_rows, ""),
        ("Total unique RRNs", len(unique_presence), ""),
        ("Blank Zenith rows used for fallback matching", loaded["ZENITH"].blank_zenith_fallback_rows, ""),
        ("Total matched across all 3", len(result_sets["MATCHED_IN_ALL_THREE"]), ""),
        ("Total mismatches", total_mismatches, ""),
        ("Total invalid RRNs", len(result_sets[INVALID_SHEET_NAME]), ""),
        ("Total duplicate RRN rows removed", sum(ds.duplicate_rows for ds in loaded.values()), ""),
    ]
    rows.append(("Output date sort priority", describe_date_priority(), ""))
    rows.append(("", "", ""))
    rows.append(("Venn Set", "Count", "Percentage Distribution"))
    denominator = max(total_valid_set_rows + len(result_sets[INVALID_SHEET_NAME]), 1)
    for sheet_name, *_ in SET_DEFINITIONS:
        count = len(result_sets[sheet_name])
        rows.append((sheet_name, count, count / denominator))
    invalid_count = len(result_sets[INVALID_SHEET_NAME])
    rows.append((INVALID_SHEET_NAME, invalid_count, invalid_count / denominator))
    return pd.DataFrame(rows, columns=["Metric", "Value", "Percentage"])


def describe_date_priority() -> str:
    labels = []
    for dataset in DATE_PRIORITY:
        labels.append(f"{dataset} {DATASETS[dataset]['date_column']}")
    return " > ".join(labels)


def build_statistics(
    loaded: Dict[str, LoadedDataset],
    result_sets: Dict[str, pd.DataFrame],
) -> pd.DataFrame:
    rows = []
    for name, ds in loaded.items():
        rows.append(
            {
                "Dataset": name,
                "Total Rows": ds.original_rows,
                "Valid RRN Rows After Deduplication": ds.valid_rows,
                "Invalid/Blank RRN Rows": ds.invalid_rows,
                "Duplicate RRN Rows Removed": ds.duplicate_rows,
            }
        )
    for sheet_name, frame in result_sets.items():
        rows.append(
            {
                "Dataset": sheet_name,
                "Total Rows": len(frame),
                "Valid RRN Rows After Deduplication": "",
                "Invalid/Blank RRN Rows": "",
                "Duplicate RRN Rows Removed": "",
            }
        )
    return pd.DataFrame(rows)


def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "RESULT") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        write_sheet(writer, df, safe_sheet_name(sheet_name))
    return output.getvalue()


def build_required_report_workbooks(valid_frames: Dict[str, pd.DataFrame]) -> Dict[str, bytes]:
    return {
        filename: build_required_report_workbook(valid_frames, filename)
        for filename in REPORT_DOWNLOAD_ORDER
    }


def build_required_report_workbook(valid_frames: Dict[str, pd.DataFrame], filename: str) -> bytes:
    return workbook_sheets_to_bytes(build_required_report_sheets(valid_frames, filename))


def build_required_report_sheets(valid_frames: Dict[str, pd.DataFrame], filename: str) -> Dict[str, pd.DataFrame]:
    report_sheets = {
        "Cash234_vs_Zenith_Reconciliation.xlsx": build_pairwise_reconciliation_report(
            valid_frames=valid_frames,
            primary_dataset="CASH234",
            secondary_dataset="ZENITH",
            reference_order=("ZENITH", "CASH234"),
            unmatched_sheet_name="Zenith Unmatched",
        ),
        "Cash234_vs_Interswitch_Reconciliation.xlsx": build_pairwise_reconciliation_report(
            valid_frames=valid_frames,
            primary_dataset="CASH234",
            secondary_dataset="INTERSWITCH",
            reference_order=("INTERSWITCH", "CASH234"),
            unmatched_sheet_name="Interswitch Unmatched",
        ),
        "Interswitch_vs_Zenith_Reconciliation.xlsx": build_pairwise_reconciliation_report(
            valid_frames=valid_frames,
            primary_dataset="INTERSWITCH",
            secondary_dataset="ZENITH",
            reference_order=("ZENITH", "INTERSWITCH"),
            unmatched_sheet_name="Zenith Unmatched",
        ),
        "Unmatched_Transactions.xlsx": build_true_unmatched_report(valid_frames),
        "Interswitch_Unsettled.xlsx": {
            "Interswitch Unsettled": build_dataset_minus_report(valid_frames, "CASH234", "INTERSWITCH"),
        },
        "Zenith_Unsettled.xlsx": {
            "Zenith Unsettled": build_dataset_minus_report(valid_frames, "INTERSWITCH", "ZENITH"),
        },
    }
    if filename not in report_sheets:
        raise ReconciliationError(f"Unknown report requested: {filename}")
    return report_sheets[filename]


def build_pairwise_reconciliation_report(
    valid_frames: Dict[str, pd.DataFrame],
    primary_dataset: str,
    secondary_dataset: str,
    reference_order: Tuple[str, str],
    unmatched_sheet_name: str,
) -> Dict[str, pd.DataFrame]:
    primary_frame = valid_frames[primary_dataset]
    secondary_frame = valid_frames[secondary_dataset]
    primary_rrns = primary_frame["NORMALIZED_RRN"]
    secondary_rrns = pd.Index(secondary_frame["NORMALIZED_RRN"].dropna().unique())
    primary_rrn_index = pd.Index(primary_rrns.dropna().unique())

    settled_mask = primary_rrns.isin(secondary_rrns)
    secondary_matched_mask = secondary_frame["NORMALIZED_RRN"].isin(primary_rrn_index)

    secondary_unmatched_mask = ~secondary_matched_mask

    sheets: Dict[str, pd.DataFrame] = {}
    for dataset in reference_order:
        sheets[f"{display_dataset_name(dataset)} Reference Sheet"] = dataset_report_frame(valid_frames[dataset], dataset)

    settled = dataset_report_frame(primary_frame.loc[settled_mask], primary_dataset)
    unsettled = dataset_report_frame(primary_frame.loc[~settled_mask], primary_dataset)
    unmatched = dataset_report_frame(secondary_frame.loc[secondary_unmatched_mask], secondary_dataset)

    sheets["Settled"] = settled
    sheets["Unsettled"] = unsettled
    sheets[unmatched_sheet_name] = unmatched
    sheets["Summary"] = build_pairwise_report_summary(
        primary_dataset=primary_dataset,
        secondary_dataset=secondary_dataset,
        primary_count=len(primary_frame),
        secondary_count=len(secondary_frame),
        settled_count=len(settled),
        unsettled_count=len(unsettled),
        unmatched_count=len(unmatched),
        unmatched_label=unmatched_sheet_name,
    )
    return sheets


def build_interswitch_zenith_tran_description_match_masks(
    primary_frame: pd.DataFrame,
    zenith_frame: pd.DataFrame,
    already_settled_mask: pd.Series,
    already_matched_zenith_mask: pd.Series,
) -> Tuple[pd.Series, pd.Series]:
    tran_matched_mask = pd.Series(False, index=primary_frame.index)
    description_matched_mask = pd.Series(False, index=zenith_frame.index)
    links = build_interswitch_zenith_description_links(
        primary_frame=primary_frame,
        zenith_frame=zenith_frame,
        already_settled_mask=already_settled_mask,
        already_matched_zenith_mask=already_matched_zenith_mask,
    )
    for interswitch_index, zenith_index in links:
        tran_matched_mask.loc[interswitch_index] = True
        description_matched_mask.loc[zenith_index] = True

    return tran_matched_mask, description_matched_mask


def build_interswitch_zenith_description_links(
    primary_frame: pd.DataFrame,
    zenith_frame: pd.DataFrame,
    already_settled_mask: pd.Series,
    already_matched_zenith_mask: pd.Series,
) -> List[Tuple[object, object]]:
    tran_id_column = find_existing_prefixed_column(primary_frame, "INTERSWITCH", "Tran_ID")
    description_column = find_existing_prefixed_column(zenith_frame, "ZENITH", "Description")
    if tran_id_column is None or description_column is None:
        return []
    return match_interswitch_zenith_description_links(
        primary_frame=primary_frame,
        zenith_frame=zenith_frame,
        already_settled_mask=already_settled_mask,
        already_matched_zenith_mask=already_matched_zenith_mask,
        tran_id_column=tran_id_column,
        description_column=description_column,
    )


def build_true_unmatched_report(valid_frames: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    sheets: Dict[str, pd.DataFrame] = {}
    for dataset in ("INTERSWITCH", "CASH234", "ZENITH"):
        other_rrns = pd.Index([], dtype="object")
        for other_dataset, other_frame in valid_frames.items():
            if other_dataset != dataset:
                other_rrns = other_rrns.union(pd.Index(other_frame["NORMALIZED_RRN"].dropna().unique()))
        frame = valid_frames[dataset]
        unmatched = frame.loc[~frame["NORMALIZED_RRN"].isin(other_rrns)]
        sheets[display_dataset_name(dataset)] = dataset_report_frame(unmatched, dataset)
    return sheets


def build_dataset_minus_report(
    valid_frames: Dict[str, pd.DataFrame],
    source_dataset: str,
    missing_from_dataset: str,
) -> pd.DataFrame:
    source_frame = valid_frames[source_dataset]
    target_frame = valid_frames[missing_from_dataset]
    target_rrns = pd.Index(target_frame["NORMALIZED_RRN"].dropna().unique())
    matched_mask = source_frame["NORMALIZED_RRN"].isin(target_rrns)

    return dataset_report_frame(source_frame.loc[~matched_mask], source_dataset)


def dataset_report_frame(df: pd.DataFrame, dataset: str) -> pd.DataFrame:
    sorted_frame = sort_dataset_frame_by_own_date(df, dataset)
    report_frame = remove_dataset_prefixes_from_headers(sorted_frame)
    metadata_columns = {
        "SOURCE_DATASET",
        "SOURCE_ROW_NUMBER",
        f"{dataset}_DUPLICATE_COUNT",
        f"{dataset}_DUPLICATE_FLAG",
        "DUPLICATE_COUNT",
        "DUPLICATE_FLAG",
    }
    return report_frame.drop(columns=[col for col in metadata_columns if col in report_frame.columns])


def build_pairwise_report_summary(
    primary_dataset: str,
    secondary_dataset: str,
    primary_count: int,
    secondary_count: int,
    settled_count: int,
    unsettled_count: int,
    unmatched_count: int,
    unmatched_label: str,
) -> pd.DataFrame:
    primary_label = display_dataset_name(primary_dataset)
    secondary_label = display_dataset_name(secondary_dataset)
    settlement_percentage = settled_count / primary_count if primary_count else 0
    return pd.DataFrame(
        [
            (f"Total {primary_label} Records", primary_count),
            (f"Total {secondary_label} Records", secondary_count),
            ("Settled Count", settled_count),
            ("Unsettled Count", unsettled_count),
            (f"{unmatched_label} Count", unmatched_count),
            ("Settlement Percentage", f"{settlement_percentage:.2%}"),
        ],
        columns=["Metric", "Value"],
    )


def workbook_sheets_to_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, frame in sheets.items():
            write_sheet(writer, frame, safe_sheet_name(sheet_name))
    return output.getvalue()


def display_dataset_name(dataset: str) -> str:
    return {
        "CASH234": "Cash234",
        "INTERSWITCH": "Interswitch",
        "ZENITH": "Zenith",
    }[dataset]


def workbook_to_bytes(results: Dict[str, object]) -> bytes:
    output = io.BytesIO()
    result_sets: Dict[str, pd.DataFrame] = results["sets"]  # type: ignore[assignment]
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        write_sheet(writer, results["summary"], "Summary")  # type: ignore[arg-type]
        write_sheet(writer, results["statistics"], "Statistics")  # type: ignore[arg-type]
        for sheet_name, frame in result_sets.items():
            if sheet_name in PAIRWISE_COMPLEMENT_SHEETS:
                continue
            write_sheet(writer, frame, safe_sheet_name(sheet_name))
    return output.getvalue()


def duplicates_workbook_to_bytes(results: Dict[str, object]) -> bytes:
    output = io.BytesIO()
    duplicates: Dict[str, pd.DataFrame] = results["duplicates"]  # type: ignore[assignment]
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for dataset_name in DATASETS:
            write_sheet(writer, duplicates[dataset_name], safe_sheet_name(dataset_name))
    return output.getvalue()


def invalid_rrn_workbook_to_bytes(results: Dict[str, object]) -> bytes:
    output = io.BytesIO()
    result_sets: Dict[str, pd.DataFrame] = results["sets"]  # type: ignore[assignment]
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        write_sheet(writer, result_sets[INVALID_SHEET_NAME], safe_sheet_name(INVALID_SHEET_NAME))
    return output.getvalue()


def updated_source_workbook_to_bytes(results: Dict[str, object], dataset: str) -> bytes:
    output = io.BytesIO()
    updated_exports: Dict[str, pd.DataFrame] = results["updated_exports"]  # type: ignore[assignment]
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        write_sheet(writer, updated_exports[dataset], safe_sheet_name(display_dataset_name(dataset)))
    return output.getvalue()


def build_interswitch_zenith_carryover_sheets(valid_frames: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    interswitch_frame = valid_frames["INTERSWITCH"]
    zenith_frame = valid_frames["ZENITH"]

    interswitch_rrns = pd.Index(
        interswitch_frame["NORMALIZED_RRN"].dropna().astype("string").loc[lambda series: series.str.strip().ne("")].unique()
    )
    zenith_rrns = pd.Index(
        zenith_frame["NORMALIZED_RRN"].dropna().astype("string").loc[lambda series: series.str.strip().ne("")].unique()
    )

    interswitch_carryover = interswitch_frame.loc[~interswitch_frame["NORMALIZED_RRN"].isin(zenith_rrns)]
    zenith_carryover = zenith_frame.loc[~zenith_frame["NORMALIZED_RRN"].isin(interswitch_rrns)]

    return {
        "Interswitch Carryover": dataset_report_frame(interswitch_carryover, "INTERSWITCH"),
        "Zenith Carryover": dataset_report_frame(zenith_carryover, "ZENITH"),
    }


def interswitch_zenith_carryover_workbook_to_bytes(results: Dict[str, object]) -> bytes:
    valid_frames: Dict[str, pd.DataFrame] = results["report_frames"]  # type: ignore[assignment]
    return workbook_sheets_to_bytes(build_interswitch_zenith_carryover_sheets(valid_frames))


def write_sheet(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str) -> None:
    validate_excel_sheet_size(df, sheet_name)
    safe_df = df.copy()
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
    gap_format = workbook.add_format({"bg_color": "#FFFFFF", "align": "center", "valign": "vcenter"})
    percent_format = workbook.add_format({"num_format": "0.00%", "align": "center", "valign": "vcenter"})
    for col_idx, column in enumerate(safe_df.columns):
        if column in GAP_COLUMNS:
            worksheet.write_blank(0, col_idx, None, gap_format)
            worksheet.set_column(col_idx, col_idx, 18, gap_format)
            continue
        worksheet.write(0, col_idx, column, header_format)
        width = calculate_excel_column_width(safe_df, col_idx, column)
        column_format = percent_format if str(column).lower().startswith("percentage") else cell_format
        worksheet.set_column(col_idx, col_idx, width, column_format)
        if str(column).lower().startswith("percentage"):
            worksheet.set_column(col_idx, col_idx, width, percent_format)
    if len(safe_df) > 0 and len(safe_df.columns) > 0:
        worksheet.freeze_panes(1, 0)
    worksheet.set_row(0, 24, header_format)


def calculate_excel_column_width(df: pd.DataFrame, col_idx: int, column: object) -> int:
    header_width = len(str(column))
    if df.empty:
        content_width = 0
    else:
        sample = df.iloc[:WIDTH_SAMPLE_ROWS, col_idx].astype("string").fillna("")
        content_width = int(sample.map(len).max())
    return min(max(header_width, content_width) + 3, 60)


def validate_excel_sheet_size(df: pd.DataFrame, sheet_name: str) -> None:
    if len(df) + 1 > EXCEL_MAX_ROWS:
        raise ReconciliationError(
            f"Sheet '{sheet_name}' has {len(df):,} data rows, which exceeds Excel's "
            f"{EXCEL_MAX_ROWS - 1:,} row export limit."
        )
    if len(df.columns) > EXCEL_MAX_COLUMNS:
        raise ReconciliationError(
            f"Sheet '{sheet_name}' has {len(df.columns):,} columns, which exceeds Excel's "
            f"{EXCEL_MAX_COLUMNS:,} column export limit."
        )


def safe_sheet_name(name: str) -> str:
    cleaned = re.sub(r"[\[\]:*?/\\]", "_", name)
    return cleaned[:31]


def filter_dataframe(df: pd.DataFrame, search: str) -> pd.DataFrame:
    if not search.strip() or df.empty:
        return df
    needle = search.strip().upper()
    searchable = df.astype("string").fillna("")
    mask = searchable.apply(lambda col: col.str.upper().str.contains(re.escape(needle), na=False))
    return df.loc[mask.any(axis=1)]
