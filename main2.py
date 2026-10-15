# main_logic.py
from __future__ import annotations

import codecs
import csv
import gc
import os
import re
from typing import Iterator, Optional

import numpy as np
import pandas as pd

from location_codes_finder import Location_Codes_Finder
from matrix_mapping import MatrixMapper


# =========================
# Tunables
# =========================
CHUNK_SIZE = 500_000            # smaller chunks = far lower peak RAM
MAX_CACHE_ENTRIES = 2_000_000   # per-cache guard; cleared if exceeded
OUTPUT_ENCODING = "utf-8-sig"   # BOM so Power BI detects UTF-8 reliably


# =========================
# Encoding / delimiter detection (streaming, no full load)
# =========================
_ENCODING_CANDIDATES = ["utf-8", "cp1252", "latin1"]


def _detect_txt_encoding(file_path: str) -> str:
    """Fast BOM/null-byte sniff to decide if the file is UTF-16."""
    with open(file_path, "rb") as f:
        header = f.read(4)
    if header[:2] in (b"\xff\xfe", b"\xfe\xff") or b"\x00" in header:
        return "utf-16"
    return "utf-8"


def _validate_encoding(file_path: str, candidates: list[str]) -> str:
    """
    Stream the file in binary and find the first encoding that decodes the WHOLE
    file. Prevents the 'failed mid-file' restarts. Constant memory (1 MB buffer).
    """
    for enc in candidates:
        decoder = codecs.getincrementaldecoder(enc)(errors="strict")
        ok = True
        try:
            with open(file_path, "rb") as f:
                while True:
                    block = f.read(1024 * 1024)
                    if not block:
                        decoder.decode(b"", final=True)
                        break
                    decoder.decode(block)
        except (UnicodeDecodeError, LookupError):
            ok = False
        if ok:
            return enc
    return candidates[-1]


def _read_csv_compat(**kwargs) -> pd.DataFrame:
    """pd.read_csv that silently drops kwargs unsupported by older pandas."""
    try:
        return pd.read_csv(**kwargs)
    except TypeError:
        kwargs.pop("encoding_errors", None)
        return pd.read_csv(**kwargs)


def _detect_txt_delimiter(file_path: str, encoding: str) -> str:
    for sep in ["\t", "|", ",", ";"]:
        try:
            probe = _read_csv_compat(
                filepath_or_buffer=file_path, sep=sep, encoding=encoding,
                encoding_errors="replace", engine="python",
                on_bad_lines="skip", quoting=csv.QUOTE_NONE, nrows=100,
            )
            if probe.shape[1] > 1:
                return sep
        except Exception:
            continue
    return "\t"


# =========================
# File Reader (small/reference files only)
# =========================
def read_file_to_df(file_path: str) -> pd.DataFrame:
    name = file_path.lower()

    if name.endswith((".xlsx", ".xls", ".xlsm")):
        return pd.read_excel(file_path)

    if name.endswith(".csv"):
        enc = _validate_encoding(file_path, _ENCODING_CANDIDATES)
        return _read_csv_compat(
            filepath_or_buffer=file_path, encoding=enc,
            encoding_errors="replace", low_memory=False,
        )

    if name.endswith(".txt"):
        enc = _detect_txt_encoding(file_path)
        if enc != "utf-16":
            enc = _validate_encoding(file_path, _ENCODING_CANDIDATES)
        sep = _detect_txt_delimiter(file_path, enc)
        try:
            df = _read_csv_compat(
                filepath_or_buffer=file_path, sep=sep, encoding=enc,
                encoding_errors="replace", engine="python",
                on_bad_lines="skip", quoting=csv.QUOTE_NONE,
            )
            if df.shape[1] > 1:
                return df
        except Exception:
            pass
        try:
            return pd.read_fwf(file_path, encoding="latin1")
        except Exception:
            pass
        return _read_csv_compat(
            filepath_or_buffer=file_path, sep=None, engine="python",
            encoding="latin1", on_bad_lines="skip",
        )

    raise ValueError("Unsupported file type. Use .xlsx, .csv, or .txt")


# =========================
# Memoization helpers
# =========================
_KEYSEP = "\x01"


class _Miss:
    pass


_MISS = _Miss()


def _guard(cache: dict) -> None:
    if len(cache) > MAX_CACHE_ENTRIES:
        cache.clear()


def memo_map(values: pd.Series, fn, cache: dict) -> pd.Series:
    """
    Call `fn` once per DISTINCT value instead of once per row, reusing results
    across chunks via `cache`. Returns a Series aligned to `values`.
    """
    vals = values.fillna("").astype(str)
    local = {}
    for v in pd.unique(vals):
        got = cache.get(v, _MISS)
        if got is _MISS:
            got = fn(v)
            cache[v] = got
        local[v] = got
    _guard(cache)
    return vals.map(local)


def memo_map_multi(df: pd.DataFrame, cols: list[str], fn, cache: dict) -> pd.Series:
    """Same idea for multi-argument functions; key is the joined column values."""
    parts = [df[c].fillna("").astype(str) for c in cols]
    key = parts[0]
    for p in parts[1:]:
        key = key.str.cat(p, sep=_KEYSEP)

    local = {}
    for k in pd.unique(key):
        got = cache.get(k, _MISS)
        if got is _MISS:
            got = fn(*k.split(_KEYSEP))
            cache[k] = got
        local[k] = got
    _guard(cache)
    return key.map(local)


# =========================
# Column Standardization
# =========================
def _norm_col(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).lower())


def _pick_col(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    norm_map = {_norm_col(c): c for c in df.columns}
    for cand in candidates:
        key = _norm_col(cand)
        if key in norm_map:
            return norm_map[key]
    return None


def standardize_input_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    FIELD_ALIASES = {
        "Consignor": [
            "Consignor", "Origin Facility", "Shipper", "Ship From Name", "Origin Name", "Org Name",
            "Origin Company", "Org Company", "ShipFromName", "OrgName", "vcConsignor",
        ],
        "Consignee": [
            "Consignee", "Destination Facility", "Receiver", "Ship To Name", "Destination Name", "Dest Name",
            "Dest Company", "Destination Company", "ShipToName", "DestName", "vcConsignee",
        ],
        "Origin Address": [
            "Origin Address", "Origin Address1", "OriginStreet", "Origin Addresss", "Org Address", "Org Address1",
            "Ship From Address", "Ship From Address1", "Shipper Address", "Shipper Address1", "vcOrgAddress1",
        ],
        "Origin City": ["Origin City", "Org City", "Ship From City", "Shipper City", "vcOriginCity"],
        "Origin State": [
            "Origin State", "Origin State Code", "Org State", "Org State Code",
            "Ship From State", "Ship From State Code", "vcOriginStateCode",
        ],
        "Destination Address": [
            "Destination Address", "DestStreet", "Destination Address1", "Dest Address", "Dest Address1",
            "Ship To Address", "Ship To Address1", "Receiver Address", "Receiver Address1", "vcDestAddress1",
        ],
        "Destination City": ["Destination City", "Dest City", "Ship To City", "Receiver City", "vcDestCity"],
        "Destination State": [
            "Destination State", "Destination State Code", "Dest State", "Dest State Code",
            "Ship To State", "Ship To State Code", "vcDestStateCode",
        ],
    }

    for canonical, aliases in FIELD_ALIASES.items():
        found = _pick_col(df, aliases)
        df[canonical] = df[found] if found else ""

    for c in [
        "Consignor", "Consignee",
        "Origin Address", "Origin City", "Origin State",
        "Destination Address", "Destination City", "Destination State",
    ]:
        df[c] = df[c].fillna("").astype(str)

    return df


# =========================
# Utilities
# =========================
def clean_blank(s: pd.Series) -> pd.Series:
    return s.replace(r"^\s*$", pd.NA, regex=True)


def _normalize_loc_code(v) -> str:
    if v is None:
        return ""
    s = str(v).strip().upper()
    if not s:
        return ""
    if re.fullmatch(r"\d+", s):
        return s.zfill(4)
    return s


# Pure function -> safe to cache globally across every chunk.
_LOC_CODE_CACHE: dict = {}


def _normalize_loc_code_series(s: pd.Series) -> pd.Series:
    return memo_map(s, _normalize_loc_code, _LOC_CODE_CACHE)


def _fmt_keep_zeros(v: str, decimals: int = 5) -> str:
    if not v:
        return ""
    if not re.fullmatch(r"-?\d+(\.\d+)?", v):
        return v
    try:
        num = float(v)
    except Exception:
        return v
    return f"{num:.{decimals}f}"


_TEXT_ZEROS_CACHE: dict = {}


def _as_text_keep_zeros(series: pd.Series, decimals: int = 5) -> pd.Series:
    s = series.fillna("").astype(str).str.strip()
    return memo_map(s, lambda v: _fmt_keep_zeros(v, decimals), _TEXT_ZEROS_CACHE)


# =========================
# Pipeline Context (built ONCE, reused by every chunk)
# =========================
class PipelineContext:
    """
    Holds everything that doesn't change between chunks: normalized reference
    tables, the location finder, the mapper, and the memo caches. Building this
    once instead of per-chunk is what lets the caches survive across all 12M rows.
    """

    def __init__(
        self,
        cintas_master_data: pd.DataFrame,       # MY LOCATION TABLE
        cintas_master_data_2: pd.DataFrame,     # Master Location Table
        cintas_location_codes: pd.DataFrame,    # all_location_codes
    ):
        my_loc = cintas_master_data.copy()
        master_loc = cintas_master_data_2.copy()
        all_codes = cintas_location_codes.copy()

        if "Loc Code" in my_loc.columns:
            my_loc["Loc Code"] = _normalize_loc_code_series(my_loc["Loc Code"])
        if "Loc Code" in master_loc.columns:
            master_loc["Loc Code"] = _normalize_loc_code_series(master_loc["Loc Code"])

        codes_col = "Codes" if "Codes" in all_codes.columns else all_codes.columns[0]
        all_codes[codes_col] = all_codes[codes_col].fillna("").astype(str).str.strip().str.upper()

        self.my_loc = my_loc
        self.master_loc = master_loc

        self.finder = Location_Codes_Finder(
            master_table=my_loc,
            codes_list=all_codes[codes_col],
        )
        self.mapper = MatrixMapper()

        self.master_combined_set = set(
            my_loc.apply(
                lambda r: self.finder.combine_addr(
                    r.get("Loc_Address", ""), r.get("Loc_City", ""), r.get("Loc_ST", "")
                ),
                axis=1,
            ).dropna()
        )

        self.master_code_to_type = dict(
            zip(
                master_loc["Loc Code"].astype(str).str.upper(),
                master_loc["Type Code"].astype(str).str.upper().replace("NAN", pd.NA),
            )
        )

        ml_merge = master_loc.copy()
        missing_master = [c for c in ("ProfitCtr", "Cost Center") if c not in ml_merge.columns]
        if missing_master:
            raise KeyError(
                f"Master Location Table is missing required column(s): {missing_master}. "
                f"Found: {list(ml_merge.columns)}. Without these, every 'Profit Center EJ' "
                f"and 'Cost Center EJ' would come out blank."
            )
        ml_merge["ProfitCtr"] = ml_merge["ProfitCtr"].fillna("").astype(str).str.strip()
        ml_merge["Cost Center"] = _as_text_keep_zeros(ml_merge["Cost Center"], decimals=5)
        ml_merge = ml_merge.drop_duplicates(subset=["Loc Code"], keep="first")

        # Rename BEFORE merging so the accrual file can't collide with these
        # names. (With suffixes=("", "_master"), an input column literally named
        # "ProfitCtr" would win the rename and silently produce wrong values.)
        self.ml_merge = ml_merge[["Loc Code", "ProfitCtr", "Cost Center"]].rename(
            columns={
                "Loc Code": "__ej_loc_key",
                "ProfitCtr": "Profit Center EJ",
                "Cost Center": "Cost Center EJ",
            }
        )

        # One cache per call site, persisted across all chunks.
        self.cache_combine_addr: dict = {}
        self.cache_extract_text: dict = {}
        self.cache_extract_addr: dict = {}
        self.cache_org_dest_type: dict = {}

    def cache_report(self) -> str:
        return (
            f"addr={len(self.cache_combine_addr):,} "
            f"text={len(self.cache_extract_text):,} "
            f"lookup={len(self.cache_extract_addr):,} "
            f"orgdest={len(self.cache_org_dest_type):,}"
        )


# =========================
# Main Pipeline
# =========================
def run_pipeline(
    accrual_sheet: pd.DataFrame,
    cintas_master_data: pd.DataFrame = None,
    cintas_master_data_2: pd.DataFrame = None,
    cintas_location_codes: pd.DataFrame = None,
    ctx: Optional[PipelineContext] = None,
) -> pd.DataFrame:
    if ctx is None:
        ctx = PipelineContext(cintas_master_data, cintas_master_data_2, cintas_location_codes)

    finder = ctx.finder

    df = accrual_sheet.copy()
    df = standardize_input_columns(df)

    # Clean stray quotes
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = (
                df[col]
                .fillna("")
                .astype(str)
                .str.strip()
                .str.replace('"', '', regex=False)
            )

    # --- Combined addresses (memoized on the 3-column key)
    df["Consignor_Combined_Address"] = memo_map_multi(
        df, ["Origin Address", "Origin City", "Origin State"],
        finder.combine_addr, ctx.cache_combine_addr,
    )
    df["Consignee_Combined_Address"] = memo_map_multi(
        df, ["Destination Address", "Destination City", "Destination State"],
        finder.combine_addr, ctx.cache_combine_addr,
    )

    # --- Extract codes from text (memoized)
    df["Extracted Consignor Code"] = _normalize_loc_code_series(
        memo_map(df["Consignor"], finder.extract_from_text, ctx.cache_extract_text)
    )
    df["Extracted Consignee Code"] = _normalize_loc_code_series(
        memo_map(df["Consignee"], finder.extract_from_text, ctx.cache_extract_text)
    )

    # --- Address checker rule
    consignor_has_code = (
        df["Extracted Consignor Code"].notna()
        & (df["Extracted Consignor Code"].astype(str).str.strip() != "")
    )
    consignee_has_code = (
        df["Extracted Consignee Code"].notna()
        & (df["Extracted Consignee Code"].astype(str).str.strip() != "")
    )
    consignor_no_cintas_or_mat = ~df["Consignor"].fillna("").str.upper().str.contains(r"\b(?:CINTAS|MAT)\b", regex=True)
    consignee_no_cintas_or_mat = ~df["Consignee"].fillna("").str.upper().str.contains(r"\b(?:CINTAS|MAT)\b", regex=True)
    consignor_addr_ok = df["Consignor_Combined_Address"].fillna("").isin(ctx.master_combined_set)
    consignee_addr_ok = df["Consignee_Combined_Address"].fillna("").isin(ctx.master_combined_set)

    df.loc[consignor_has_code & consignor_no_cintas_or_mat & ~consignor_addr_ok, "Extracted Consignor Code"] = ""
    df.loc[consignee_has_code & consignee_no_cintas_or_mat & ~consignee_addr_ok, "Extracted Consignee Code"] = ""

    # --- Extract from Org/Dest Type columns (memoized on the 2-column key)
    org_type_col = _pick_col(df, ["Org Type Code", "Org Loc Code", "Origin Type Code", "OrgTypeCode", "Origin Location Code", "Org Location Code"])
    dest_type_col = _pick_col(df, ["Dest Type Code", "Dest Loc Code", "Destination Type Code", "DestTypeCode", "Destination Location Code", "Dest Location Code"])

    if org_type_col and dest_type_col:
        pairs = memo_map_multi(
            df, [org_type_col, dest_type_col],
            lambda a, b: tuple(finder.extract_from_org_dest_type(a, b)),
            ctx.cache_org_dest_type,
        )
        tmp = pd.DataFrame(
            pairs.tolist(),
            index=df.index,
            columns=["Org Type Consignor Code", "Dest Type Consignee Code"],
        )
        df[["Org Type Consignor Code", "Dest Type Consignee Code"]] = tmp
        del tmp, pairs
    else:
        df["Org Type Consignor Code"] = ""
        df["Dest Type Consignee Code"] = ""

    df["Org Type Consignor Code"] = _normalize_loc_code_series(df["Org Type Consignor Code"])
    df["Dest Type Consignee Code"] = _normalize_loc_code_series(df["Dest Type Consignee Code"])

    # --- Address looked-up codes (memoized)
    df["Addr_Lookup_Consignor_Code"] = _normalize_loc_code_series(
        memo_map(df["Consignor_Combined_Address"], finder.extract_from_address, ctx.cache_extract_addr)
    )
    df["Addr_Lookup_Consignee_Code"] = _normalize_loc_code_series(
        memo_map(df["Consignee_Combined_Address"], finder.extract_from_address, ctx.cache_extract_addr)
    )

    # --- Special override: Mississauga "Suite" consignee rule
    consignee_combo = df["Consignee_Combined_Address"].fillna("").astype(str).str.upper()
    consignor_txt = df["Consignor"].fillna("").astype(str).str.upper().str.strip()
    is_suite_mississauga = consignee_combo.str.startswith("SUITEMISSISSAUGAON")

    df.loc[is_suite_mississauga & consignor_txt.str.startswith(("LNK", "AMERICAN METAL CRAFTERS", "RADIANS", "EVER READY")), "Addr_Lookup_Consignee_Code"] = "097H"
    df.loc[is_suite_mississauga & consignor_txt.str.startswith(("VECTAIR", "ZEP")), "Addr_Lookup_Consignee_Code"] = "067N"
    df.loc[is_suite_mississauga & consignor_txt.str.startswith(("CHEMFREE", "BERRY GLOBAL")), "Addr_Lookup_Consignee_Code"] = "0897"

    # --- Final location codes
    # Extracted -> Address Lookup -> NON-CINTAS.
    # Org/Dest Type codes are NOT used as a fallback: if Extracted AND Address
    # Lookup are both blank, the final code is NON-CINTAS (exception rules below
    # can still assign a real code).
    df["Final Consignor Code"] = (
        clean_blank(df["Extracted Consignor Code"])
        .fillna(clean_blank(df["Addr_Lookup_Consignor_Code"]))
        .fillna("NON-CINTAS")
    )
    df["Final Consignee Code"] = (
        clean_blank(df["Extracted Consignee Code"])
        .fillna(clean_blank(df["Addr_Lookup_Consignee_Code"]))
        .fillna("NON-CINTAS")
    )

    df["Final Consignor Code"] = _normalize_loc_code_series(df["Final Consignor Code"])
    df["Final Consignee Code"] = _normalize_loc_code_series(df["Final Consignee Code"])

    # =========================
    # Exceptions
    # =========================
    EXCEPTION_RULES = [
        {"match_column": "Consignee", "contains": "MGA", "set_column": "Final Consignee Code", "value": "099H"},
        {"match_column": "Consignee", "contains": "HONDURAS", "set_column": "Final Consignee Code", "value": "0957"},
        {"match_column": "Consignee", "contains": "CARIBBEAN", "set_column": "Final Consignee Code", "value": "0W36"},
        {"match_column": "Destination Address", "contains": "6001 W", "set_column": "Final Consignee Code", "value": "0021"},
        {"match_column": "Consignor", "contains": "AVERITT TERMINAL", "set_column": "Final Consignor Code", "value": "0004"},
        {"match_column": "Consignee", "contains": "VALDEZ", "set_column": "Final Consignee Code", "value": "0K35"},
        {"match_column": "Destination Address", "contains": "ATTN: GARDNER", "set_column": "Final Consignee Code", "value": "0536"},
        {"match_column": "Destination Address", "contains": "ATTENTION - STEVEN FRILEY", "set_column": "Final Consignee Code", "value": "0311"},
        {"match_column": "Destination Address", "contains": "ATTN:  ROSIE BOCANEGRA", "set_column": "Final Consignee Code", "value": "0650"},
        {"match_column": "Consignee", "contains": "WILLIAM BROOM", "set_column": "Final Consignee Code", "value": "001M"},
        {"match_column": "Consignee", "contains": "COLTON NELSON", "set_column": "Final Consignee Code", "value": "012M"},
        {"match_column": "Consignee", "contains": "ANGELA CALDWELL", "set_column": "Final Consignee Code", "value": "0206"},
        {"match_column": "Consignee", "contains": "ANGELA JOHNSON", "set_column": "Final Consignee Code", "value": "0128"},
        {"match_column": "Consignee", "contains": "BELINDA LOCKHART HALL", "set_column": "Final Consignee Code", "value": "0004"},
        {"match_column": "Destination Address", "contains": "DAVE N. 11F", "set_column": "Final Consignee Code", "value": "011F"},
        {"match_column": "Destination Address", "contains": "ATTN TO: ADAM PAWLOWICZ", "set_column": "Final Consignee Code", "value": "0055"},
        {"match_column": "Destination Address", "contains": "GEORGE FREEMAN", "set_column": "Final Consignee Code", "value": "0280"},
        {"match_column": "Consignee", "contains": "BRENDA LOCKHARTHALL", "set_column": "Final Consignee Code", "value": "0004"},
        {"match_column": "Consignor", "contains": "AVERITT TERMINAL", "set_column": "Final Consignor Code", "value": "0004"},
        {"match_column": "Consignee", "contains": "BRET JOHNSON", "set_column": "Final Consignee Code", "value": "0452"},
        {"match_column": "Consignor", "contains": "COOPETRAJES", "set_column": "Final Consignor Code", "value": "0896"},
        {"match_column": "Consignee", "contains": "COOPETRAJES", "set_column": "Final Consignee Code", "value": "0896"},
        {"match_column": "Consignor", "contains": "MATHESON", "set_column": "Final Consignor Code", "value": "067N"},
        {"match_column": "Consignor", "contains": "EMPRESSA", "set_column": "Final Consignor Code", "value": "0972"},
        {"match_column": "Consignor", "contains": "EMPRESA", "set_column": "Final Consignor Code", "value": "0972"},
        {"match_column": "Consignee", "contains": "EMPRESSA", "set_column": "Final Consignee Code", "value": "0972"},
        {"match_column": "Consignee", "contains": "EMPRESA", "set_column": "Final Consignee Code", "value": "0972"},
        {"match_column": "Destination Address", "contains": "DEBBIE HUGHES", "set_column": "Final Consignee Code", "value": "0012"},
        {"match_column": "Destination Address", "contains": "LISA SIMPSON", "set_column": "Final Consignee Code", "value": "0210"},
        {"match_column": "Destination Address", "contains": "AMANDA SPEAKS", "set_column": "Final Consignee Code", "value": "0230"},
        {"match_column": "Destination Address", "contains": "ATTN: STOCKROOM", "set_column": "Final Consignee Code", "value": "019M"},
        {"match_column": "Destination Address", "contains": "JOHN VICARI", "set_column": "Final Consignee Code", "value": "0616"},
        {"match_column": "Destination Address", "contains": "AGUEDA", "set_column": "Final Consignee Code", "value": "0464"},
        {"match_column": "Destination Address", "contains": "40 BROADHEAD", "set_column": "Final Consignee Code", "value": "0531"},
        {"match_column": "Destination Address", "contains": "625 ELMWOOD", "set_column": "Final Consignee Code", "value": "0544"},
        {"match_column": "Destination Address", "contains": "TRADE ST", "set_column": "Final Consignee Code", "value": "0312"},
        {"match_column": "Destination Address", "contains": "TRADE STREET", "set_column": "Final Consignee Code", "value": "0312"},
        {"match_column": "Destination Address", "contains": "GWEN ARMSTRONG", "set_column": "Final Consignee Code", "value": "0016"},
        {"match_column": "Destination Address", "contains": "ANNESSA BRITTON", "set_column": "Final Consignee Code", "value": "0172"},
        {"match_column": "Destination Address", "contains": "BRIGHTSEAT", "set_column": "Final Consignee Code", "value": "0041"},
        {"match_column": "Destination Address", "contains": "320 WEST", "set_column": "Final Consignee Code", "value": "0006"},
        {"match_column": "Destination Address", "contains": "1111 N", "set_column": "Final Consignee Code", "value": "0017"},
        {"match_column": "Destination Address", "contains": "1111 NORTHWEST", "set_column": "Final Consignee Code", "value": "0017"},
        {"match_column": "Destination Address", "contains": "1100 REMINGTON", "set_column": "Final Consignee Code", "value": "0022"},
        {"match_column": "Destination Address", "contains": "CHERI WILLS", "set_column": "Final Consignee Code", "value": "0524"},
        {"match_column": "Destination Address", "contains": "TRICIA SMITH", "set_column": "Final Consignee Code", "value": "0701"},
        {"match_column": "Destination Address", "contains": "ROSELINE LICIANO", "set_column": "Final Consignee Code", "value": "0756"},
        {"match_column": "Destination Address", "contains": "3605 WHITE", "set_column": "Final Consignee Code", "value": "030R"},
        {"match_column": "Consignee", "contains": "TERRE", "set_column": "Final Consignee Code", "value": "0370"},
        {"match_column": "Consignee", "contains": "DAISY MEJIA", "set_column": "Final Consignee Code", "value": "082J"},
        {"match_column": "Consignee", "contains": "F78", "set_column": "Final Consignee Code", "value": "0F78"},
        {"match_column": "Consignee", "contains": "38R", "set_column": "Final Consignee Code", "value": "038R"},
        {"match_column": "Consignee", "contains": "73Q", "set_column": "Final Consignee Code", "value": "073Q"},
        {"match_column": "Consignee", "contains": "ADAM LEE", "set_column": "Final Consignee Code", "value": "0184"},
        {"match_column": "Consignee", "contains": "ALISIA PULELLA", "set_column": "Final Consignee Code", "value": "0K95"},
        {"match_column": "Consignee", "contains": "CAROL WALLINGER", "set_column": "Final Consignee Code", "value": "0044"},
        {"match_column": "Consignee", "contains": "CASEY LAYMAN", "set_column": "Final Consignee Code", "value": "0336"},
        {"match_column": "Consignee", "contains": "CINTAS CORPORATION LOCATION G4", "set_column": "Final Consignee Code", "value": "0G44"},
        {"match_column": "Consignee", "contains": "DAISY MEJIA", "set_column": "Final Consignee Code", "value": "0577"},
        {"match_column": "Consignee", "contains": "DARREN BURICH", "set_column": "Final Consignee Code", "value": "0310"},
        {"match_column": "Consignee", "contains": "JOSEPHINE MANCILLA", "set_column": "Final Consignee Code", "value": "0067"},
        {"match_column": "Consignee", "contains": "K90", "set_column": "Final Consignee Code", "value": "0K90"},
        {"match_column": "Consignee", "contains": "JOE WILLIAMS", "set_column": "Final Consignee Code", "value": "001M"},
    ]

    # Uppercase each match column ONCE instead of once per rule (65 passes -> 3).
    _upper: dict = {}
    for rule in EXCEPTION_RULES:
        col = rule["match_column"]
        if col not in df.columns:
            continue
        if col not in _upper:
            _upper[col] = df[col].fillna("").astype(str).str.upper()
        match_series = _upper[col].str.contains(rule["contains"], na=False)
        df.loc[match_series, rule["set_column"]] = rule["value"]
    del _upper

    df["Final Consignor Code"] = _normalize_loc_code_series(df["Final Consignor Code"])
    df["Final Consignee Code"] = _normalize_loc_code_series(df["Final Consignee Code"])

    # --- Type code mapping
    df["Final Consignor Type"] = (
        df["Final Consignor Code"].astype(str).str.upper().map(ctx.master_code_to_type).fillna("NON-CINTAS")
    )
    df["Final Consignee Type"] = (
        df["Final Consignee Code"].astype(str).str.upper().map(ctx.master_code_to_type).fillna("NON-CINTAS")
    )

    # --- Force NON-CINTAS type where the final code is NON-CINTAS, i.e. Extracted
    # AND Address Lookup were both blank and no exception rule assigned a code.
    # Placed BEFORE Responsible Party so the mapper sees the overridden types.
    consignor_no_code = df["Final Consignor Code"].eq("NON-CINTAS")
    consignee_no_code = df["Final Consignee Code"].eq("NON-CINTAS")
    df.loc[consignor_no_code, "Final Consignor Type"] = "NON-CINTAS"
    df.loc[consignee_no_code, "Final Consignee Type"] = "NON-CINTAS"

    # --- Responsible party
    # NOTE: still row-wise — this is now the slowest step in the pipeline.
    df["Responsible Party"] = df.apply(ctx.mapper.determine_profit_center, axis=1)

    # --- Responsible Party = NON-CINTAS for rows forced to NON-CINTAS above
    # (Extracted AND Address Lookup both blank on either side).
    # Profit Center EJ / Cost Center EJ are then blanked by the NON-CINTAS mask below.
    # To apply only when BOTH sides have no code, change "|" to "&".
    rp_non_cintas = consignor_no_code & consignee_no_code
    df.loc[rp_non_cintas, "Responsible Party"] = "NON-CINTAS"

    # --- Profit/Cost center lookup
    # Any pre-existing columns with our target names would be clobbered by the
    # merge, so drop them first — the master table is the source of truth here.
    df = df.drop(columns=[c for c in ("Profit Center EJ", "Cost Center EJ") if c in df.columns])

    df = df.merge(
        ctx.ml_merge,
        left_on="Responsible Party",
        right_on="__ej_loc_key",
        how="left",
    )
    df = df.drop(columns=["__ej_loc_key"])

    if "Profit Center" in df.columns:
        df["Profit Center"] = (
            df["Profit Center"].fillna("").astype(str).str.strip().str.replace('"', '', regex=False)
        )

    df["Profit Center EJ"] = (
        df.get("Profit Center EJ", pd.Series([""] * len(df))).fillna("").astype(str).str.strip()
    )
    df["Cost Center EJ"] = _as_text_keep_zeros(
        df.get("Cost Center EJ", pd.Series([""] * len(df))), decimals=5
    )

    # Blank if THIRD PARTY / NON-CINTAS
    rp_norm = df["Responsible Party"].fillna("").astype(str).str.upper().str.strip()
    mask_blank = rp_norm.isin(["THIRD PARTY", "NON-CINTAS"])
    df.loc[mask_blank, "Profit Center EJ"] = ""
    df.loc[mask_blank, "Cost Center EJ"] = ""

    # --- GL account logic (vectorized; same result as the old row-wise lambda)
    has_g59 = df["Profit Center EJ"].astype(str).str.contains("G59", regex=False, na=False)
    consignee_is_rp = (
        df["Final Consignee Code"].astype(str) == df["Responsible Party"].astype(str)
    )
    df["Account # EJ"] = np.where(has_g59 | consignee_is_rp, "621000", "621020")

    # --- Automation Accuracy
    if {"Profit Center", "Profit Center EJ"}.issubset(df.columns):
        match = (
            (df["Profit Center"] == df["Profit Center EJ"])
            & df["Profit Center"].notna()
            & df["Profit Center EJ"].notna()
        )
        df["Automation Accuracy"] = match.astype(int)
    else:
        df["Automation Accuracy"] = 0

    # --- Column order
    first_cols = [
        "Profit Center", "Cost Center", "Account #",
        "Automation Accuracy", "Profit Center EJ", "Cost Center EJ", "Account # EJ",
    ]
    first_cols = [c for c in first_cols if c in df.columns]
    other_cols = [c for c in df.columns if c not in first_cols]
    df = df[first_cols + other_cols]

    return df


# =========================
# Lazy chunk iterator (NEVER materializes the whole file)
# =========================
def _iter_accrual_chunks(accrual_path: str) -> Iterator[pd.DataFrame]:
    lower = accrual_path.lower()

    if lower.endswith((".xlsx", ".xls", ".xlsm")):
        print("\nReading Excel accrual file (Excel cannot be streamed)...")
        full_df = read_file_to_df(accrual_path)
        for i in range(0, len(full_df), CHUNK_SIZE):
            yield full_df.iloc[i:i + CHUNK_SIZE]
        return

    if lower.endswith(".csv"):
        enc = _validate_encoding(accrual_path, _ENCODING_CANDIDATES)
        print(f"\nDetected CSV — encoding: {enc}. Streaming in chunks...")
        reader = _read_csv_compat(
            filepath_or_buffer=accrual_path, encoding=enc, encoding_errors="replace",
            chunksize=CHUNK_SIZE, on_bad_lines="skip", dtype=str,
        )
        for chunk in reader:
            yield chunk
        return

    if lower.endswith(".txt"):
        enc = _detect_txt_encoding(accrual_path)
        if enc != "utf-16":
            print("Validating file encoding (streaming scan)...")
            enc = _validate_encoding(accrual_path, _ENCODING_CANDIDATES)
        sep = _detect_txt_delimiter(accrual_path, enc)
        print(f"\nDetected TXT — delimiter: {repr(sep)}, encoding: {enc}. Streaming in chunks...")
        reader = _read_csv_compat(
            filepath_or_buffer=accrual_path, sep=sep, encoding=enc,
            encoding_errors="replace", chunksize=CHUNK_SIZE,
            on_bad_lines="skip", quoting=csv.QUOTE_NONE, dtype=str,
        )
        for chunk in reader:
            yield chunk
        return

    raise ValueError("Unsupported accrual file type. Use .xlsx, .csv, or .txt")


# =========================
# Chunked Pipeline Runner
# =========================
def run_pipeline_chunked(
    accrual_path: str,
    master_data_path: str,
    master_data_2_path: str,
    location_codes_path: str,
    output_path: str,
) -> None:
    import time

    print("Loading reference tables...")
    cintas_master_data = read_file_to_df(master_data_path)
    cintas_master_data_2 = read_file_to_df(master_data_2_path)
    cintas_location_codes = read_file_to_df(location_codes_path)
    print(f"  MY LOCATION TABLE:     {len(cintas_master_data):,} rows")
    print(f"  Master Location Table: {len(cintas_master_data_2):,} rows")
    print(f"  All Location Codes:    {len(cintas_location_codes):,} rows")

    # Built ONCE — this is what lets the memo caches survive across chunks.
    ctx = PipelineContext(cintas_master_data, cintas_master_data_2, cintas_location_codes)

    if not output_path.lower().endswith(".csv"):
        print(f"NOTE: output is CSV-formatted; '{output_path}' does not end in .csv")

    header_cols: Optional[list[str]] = None
    rows_in = 0
    t_start = time.time()

    with open(output_path, "w", encoding=OUTPUT_ENCODING, newline="") as out_f:
        for i, chunk in enumerate(_iter_accrual_chunks(accrual_path), start=1):
            t0 = time.time()
            rows_in += len(chunk)

            result = run_pipeline(chunk, ctx=ctx)

            if header_cols is None:
                header_cols = list(result.columns)
                blanks = [
                    c for c in result.columns
                    if (result[c].astype(str).str.strip()
                        .replace({"nan": "", "None": ""}) == "").all()
                ]
                if blanks:
                    print(f"  NOTE: {len(blanks)} column(s) are blank for every row in "
                          f"chunk 1 — check the alias lists if unexpected:")
                    for c in blanks:
                        print(f"          {c}")
            else:
                result = result.reindex(columns=header_cols)

            result.to_csv(out_f, index=False, header=(i == 1), lineterminator="\n")

            elapsed = max(time.time() - t0, 1e-9)
            print(
                f"  Chunk {i}: {len(chunk):,} rows in {elapsed:,.1f}s "
                f"({len(chunk) / elapsed:,.0f} rows/s) | {rows_in:,} total "
                f"| caches: {ctx.cache_report()}"
            )

            del result, chunk
            gc.collect()

    total = time.time() - t_start
    size_gb = os.path.getsize(output_path) / 1024 ** 3
    print(f"\nProcessed {rows_in:,} rows in {total / 60:,.1f} min.")
    print(f"Output saved to: {output_path}  ({size_gb:,.2f} GB, {len(header_cols)} columns)")


# =========================
# Entry Point
# =========================
if __name__ == "__main__":
    # ---- Configure your file paths here ----
    ACCRUAL_FILE        = r"C:\Users\c1354623\OneDrive - Cintas Corporation\Documents\Miscellaneous_Excel_Analysis\Shipment Data\A3 Shipment Data\Cintas - Shipment Info FY24 to Present Additional Modes.txt"
    MASTER_DATA_FILE    = r"C:\Users\c1354623\OneDrive - Cintas Corporation\Documents\Miscellaneous_Excel_Analysis\Shipment Data\data\MY LOCATION TABLE (4).xlsx"
    MASTER_DATA_2_FILE  = r"C:\Users\c1354623\OneDrive - Cintas Corporation\Documents\Miscellaneous_Excel_Analysis\Shipment Data\data\Master Location Table.xlsx"
    LOCATION_CODES_FILE = r"C:\Users\c1354623\OneDrive - Cintas Corporation\Documents\Miscellaneous_Excel_Analysis\Shipment Data\data\all_location_codes.xlsx"
    OUTPUT_FILE         = r"C:\Users\c1354623\OneDrive - Cintas Corporation\Documents\Miscellaneous_Excel_Analysis\Shipment Data\Python-Processed Shipment Data\Cintas - Shipment Info FY24 to Present Additional Modes-Processed.csv"
    # -----------------------------------------

    run_pipeline_chunked(
        accrual_path=ACCRUAL_FILE,
        master_data_path=MASTER_DATA_FILE,
        master_data_2_path=MASTER_DATA_2_FILE,
        location_codes_path=LOCATION_CODES_FILE,
        output_path=OUTPUT_FILE,
    )