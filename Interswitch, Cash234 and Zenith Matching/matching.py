from __future__ import annotations

from typing import Dict, List, Tuple

import pandas as pd

from source_processing import extract_interswitch_tran_id_token, extract_zenith_description_tran_id


def build_interswitch_zenith_description_links(
    primary_frame: pd.DataFrame,
    zenith_frame: pd.DataFrame,
    already_settled_mask: pd.Series,
    already_matched_zenith_mask: pd.Series,
    tran_id_column: str,
    description_column: str,
) -> List[Tuple[object, object]]:
    remaining_interswitch = primary_frame.loc[~already_settled_mask, tran_id_column]
    unmatched_zenith_descriptions = zenith_frame.loc[~already_matched_zenith_mask, description_column]

    zenith_token_index: Dict[str, List[object]] = {}
    for zenith_index, description in unmatched_zenith_descriptions.items():
        token = extract_zenith_description_tran_id(description)
        if token:
            zenith_token_index.setdefault(token, []).append(zenith_index)

    links: List[Tuple[object, object]] = []
    used_zenith_indices = set()
    for row_index, tran_id in remaining_interswitch.items():
        token = extract_interswitch_tran_id_token(tran_id)
        if not token:
            continue
        candidates = zenith_token_index.get(token, [])
        zenith_index = next((index for index in candidates if index not in used_zenith_indices), None)
        if zenith_index is None:
            continue
        links.append((row_index, zenith_index))
        used_zenith_indices.add(zenith_index)

    return links
