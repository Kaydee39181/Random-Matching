from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation

import pandas as pd


ZENITH_TRAN_TOKEN_PATTERN = re.compile(r"(?:^|\|)ZIB\|(\d{12})\|?(\d{6})(?=\D|$)")
GENERIC_TRAN_TOKEN_PATTERN = re.compile(r"(?:^|\|)(\d{12})\|?(\d{6})(?=\D|$)")


def remove_hidden_spaces(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value)
    return (
        text.replace("\u00a0", " ")
        .replace("\u200b", "")
        .replace("\u200c", "")
        .replace("\u200d", "")
        .replace("\ufeff", "")
    )


def expand_scientific_rrn(value: object) -> str:
    text = "" if pd.isna(value) else str(value)
    if not re.fullmatch(r"[+-]?\d+(?:\.\d+)?E[+-]?\d+", text, flags=re.IGNORECASE):
        return text
    try:
        decimal_value = Decimal(text)
    except InvalidOperation:
        return text
    if decimal_value == decimal_value.to_integral_value():
        return format(decimal_value.quantize(Decimal("1")), "f")
    return format(decimal_value.normalize(), "f")


def normalize_tran_description_token(value: object) -> str:
    text = remove_hidden_spaces(value).strip().upper()
    text = re.sub(r"\s*[/|]\s*", "|", text)
    text = re.sub(r"\s+", "", text)
    return text


def extract_interswitch_tran_id_token(value: object) -> str:
    normalized = normalize_tran_description_token(value)
    for pattern in (ZENITH_TRAN_TOKEN_PATTERN, GENERIC_TRAN_TOKEN_PATTERN):
        token_match = pattern.search(normalized)
        if token_match is not None:
            return f"{token_match.group(1)}|{token_match.group(2)}"
    return ""


def extract_zenith_description_tran_id(value: object) -> str:
    normalized = normalize_tran_description_token(value)
    token_match = ZENITH_TRAN_TOKEN_PATTERN.search(normalized)
    if token_match is None:
        return ""
    return f"{token_match.group(1)}|{token_match.group(2)}"


def description_contains_tran_id_token(description: object, token: object) -> bool:
    normalized_description = normalize_tran_description_token(description)
    normalized_token = normalize_tran_description_token(token)
    if not normalized_description or not normalized_token:
        return False
    token_pattern = re.escape(normalized_token)
    return re.search(rf"(?<![A-Z0-9]){token_pattern}(?![A-Z0-9])", normalized_description) is not None
