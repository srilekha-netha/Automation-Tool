"""
reconcile_engine.py
EDW vs Target Data Reconciliation Engine.

Output workbook sheets:
  Source | Target | Data_Rec | Summary | Column_Diff

Data_Rec uses ONLY columns that exist in BOTH source and target.
Columns only in source or only in target go to Column_Diff sheet.

Data_Rec layout (side-by-side):
  [SRC col1..colN] | [TGT col1..colN] | [REC col1..colN]

Missing record rules:
  MISSING IN SOURCE → SRC section blank, TGT = values, REC = "MISSING IN SOURCE" every col
  MISSING IN TARGET   → SRC = values, TGT section blank, REC = "MISSING IN TARGET" every col

Duplicate key handling:
  When duplicate keys exist, match rows positionally within each key group
  (first src row → first tgt row, second → second, etc.).
  Unmatched extras become MISSING IN TARGET or MISSING IN SOURCE.

Date normalisation:
  Dates in any format (YYYYMMDD, YYYY-MM-DD, YYYY-MM-DD HH:MM:SS, MM/DD/YYYY, etc.)
  are normalised to YYYYMMDD before comparison, so they compare as equal.

Financial REC  = Target − Source (numeric diff)
Non-financial  = TRUE / FALSE
"""

from __future__ import annotations

import gc
import io
import re
import time
from typing import List, Tuple

import numpy as np
import pandas as pd
import xlsxwriter
import xlsxwriter.utility as xlutil

MISSING_SRC = "MISSING IN SOURCE"
MISSING_TARGET = "MISSING IN TARGET"
TOLERANCE   = 0.01

C_GREEN  = "#C6EFCE"; C_GREEN_F  = "#276221"
C_YELLOW = "#FFEB9C"; C_YELLOW_F = "#9C6500"
C_RED    = "#FFC7CE"; C_RED_F    = "#9C0006"
C_PINK   = "#FFD7D7"; C_PINK_F   = "#7B0000"
C_SRC_H  = "#1F4E79"
C_TGT_H  = "#375623"
C_REC_H  = "#7B2C2C"
C_HDR    = "#2E4057"
C_WHITE  = "#FFFFFF"
C_ORANGE = "#F4B942"; C_ORANGE_F = "#7B3F00"

COL_KEY = 22; COL_FIN = 16; COL_STD = 20


# ═══════════════════════════════════════════════════════════════════════════════
# Public entry point
# ═══════════════════════════════════════════════════════════════════════════════

def reconcile(
    source_df: pd.DataFrame,
    target_df: pd.DataFrame,
    key_cols:  List[str],
    financial_cols: List[str],
    tolerance: float = TOLERANCE,
    col_mapping: dict = None,
) -> io.BytesIO:
    """
    col_mapping: {tgt_col_name: src_col_name}
    Matched target columns are renamed to their source counterpart BEFORE
    any processing, so they become common columns in Data_Rec.
    Unmapped source-only / target-only columns still go to Column_Diff.
    """
    t0 = time.perf_counter()

    source_df = source_df.copy()
    target_df = target_df.copy()

    # Deduplicate column names (in case of Excel weirdness)
    source_df.columns = _dedupe_cols(source_df.columns)
    target_df.columns = _dedupe_cols(target_df.columns)

    # Apply user-supplied column mapping: rename target cols to source names
    # so they become "common" columns and appear in Data_Rec together.
    if col_mapping:
        target_df = target_df.rename(columns=col_mapping)

    # ── Separate common vs source-only vs target-only columns ─────────────────
    src_cols_set = set(source_df.columns)
    tgt_cols_set = set(target_df.columns)
    common_cols  = [c for c in source_df.columns if c in tgt_cols_set]   # preserve source order
    src_only_cols = [c for c in source_df.columns if c not in tgt_cols_set]
    tgt_only_cols = [c for c in target_df.columns if c not in src_cols_set]

    # Data_Rec only uses common columns
    all_cols     = common_cols
    non_fin_cols = [c for c in all_cols if c not in key_cols and c not in financial_cols]

    # ── Build composite key (on common key cols only) ─────────────────────────
    # Key cols must exist in both sheets
    key_cols_valid = [k for k in key_cols if k in src_cols_set and k in tgt_cols_set]

    source_df["__key__"] = _make_key(source_df, key_cols_valid)
    target_df["__key__"] = _make_key(target_df, key_cols_valid)

    # ── Positional matching within key groups (prevents cross-join explosion) ──
    merged_src, merged_tgt = _positional_match(source_df, target_df, key_cols_valid)

    src_exists = merged_src["__key__"].notna().to_numpy()
    tgt_exists = merged_tgt["__key__"].notna().to_numpy()

    # Extract aligned section frames (common columns only)
    src_frame = _extract_frame(merged_src, all_cols)
    tgt_frame = _extract_frame(merged_tgt, all_cols)

    n_rows = len(src_frame)
    ms   = ~src_exists
    mt   = ~tgt_exists
    both =  src_exists & tgt_exists

    # ── Vectorized reconciliation ─────────────────────────────────────────────
    rec_arrays: dict[str, np.ndarray] = {}
    fin_flag    = np.zeros(n_rows, dtype=bool)
    nonfin_flag = np.zeros(n_rows, dtype=bool)

    for col in all_cols:
        sv = src_frame[col].to_numpy(dtype=object)
        tv = tgt_frame[col].to_numpy(dtype=object)
        if col in key_cols_valid:
            rec_arrays[col] = _rec_key(sv, tv, ms, mt)
        elif col in financial_cols:
            arr, flag = _rec_fin(sv, tv, ms, mt, both, tolerance)
            rec_arrays[col] = arr; fin_flag |= flag
        else:
            arr, flag = _rec_nonfin(sv, tv, ms, mt, both)
            rec_arrays[col] = arr; nonfin_flag |= flag

    # ── Summary stats ─────────────────────────────────────────────────────────
    total_src = int(src_exists.sum())
    total_tgt = int(tgt_exists.sum())
    both_cnt  = int(both.sum())
    miss_s    = int(ms.sum())
    miss_p    = int(mt.sum())
    fin_m     = int(fin_flag.sum())
    nfin_m    = int(nonfin_flag.sum())
    pct       = round(both_cnt / max(total_src, total_tgt) * 100, 2) if max(total_src, total_tgt) else 0.0
    elapsed   = round(time.perf_counter() - t0, 2)

    summary = {
        "Total Source Records":      total_src,
        "Total Target Records":      total_tgt,
        "Matched Records":           both_cnt,
        "Records Missing in Source": miss_s,
        "Records Missing in TARGET": miss_p,
        "Financial Differences":     fin_m,
        "Non-Financial Differences": nfin_m,
        "Match Percentage":          f"{pct}%",
        "Processing Time (Seconds)": elapsed,
    }

    # ── Write workbook ────────────────────────────────────────────────────────
    buf = io.BytesIO()
    _write_workbook(
        buf, summary,
        source_df.drop(columns="__key__"),
        target_df.drop(columns="__key__"),
        src_frame, tgt_frame, rec_arrays,
        all_cols, key_cols_valid, financial_cols, non_fin_cols,
        n_rows, tolerance,
        src_only_cols, tgt_only_cols,
    )

    del src_frame, tgt_frame, rec_arrays
    gc.collect()

    buf.seek(0)
    return buf


# ═══════════════════════════════════════════════════════════════════════════════
# Positional matching within key groups  (fixes duplicate explosion)
# ═══════════════════════════════════════════════════════════════════════════════

def _positional_match(
    source_df: pd.DataFrame,
    target_df: pd.DataFrame,
    key_cols: List[str],
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Match source and target rows by key, then positionally within each group.
    Uses vectorized pandas cumcount to assign position within group, then
    outer-merges on (key, position).  No Python loops over rows.

    Returns two aligned DataFrames (same number of rows) where:
      - Row i of src_out corresponds to row i of tgt_out
      - Missing rows have NaN in __key__
    """
    src = source_df.reset_index(drop=True).copy()
    tgt = target_df.reset_index(drop=True).copy()

    key = "__key__"

    # Assign within-group position (0, 1, 2, …) for each key
    src["__pos__"] = src.groupby(key, sort=False).cumcount()
    tgt["__pos__"] = tgt.groupby(key, sort=False).cumcount()

    # Outer merge on (key, position)
    merged = pd.merge(
        src.add_prefix("SRC_"),
        tgt.add_prefix("TGT_"),
        left_on=["SRC___key__",  "SRC___pos__"],
        right_on=["TGT___key__", "TGT___pos__"],
        how="outer",
    )

    # Reconstruct src_out and tgt_out as separate frames
    src_cols_raw = [c for c in src.columns if c not in ("__pos__",)]
    tgt_cols_raw = [c for c in tgt.columns if c not in ("__pos__",)]

    src_out_data = {}
    for c in src_cols_raw:
        pcol = f"SRC_{c}"
        src_out_data[c] = merged[pcol].to_numpy(dtype=object) if pcol in merged.columns else np.full(len(merged), np.nan, dtype=object)

    tgt_out_data = {}
    for c in tgt_cols_raw:
        pcol = f"TGT_{c}"
        tgt_out_data[c] = merged[pcol].to_numpy(dtype=object) if pcol in merged.columns else np.full(len(merged), np.nan, dtype=object)

    src_out = pd.DataFrame(src_out_data).reset_index(drop=True)
    tgt_out = pd.DataFrame(tgt_out_data).reset_index(drop=True)

    return src_out, tgt_out


# ═══════════════════════════════════════════════════════════════════════════════
# Helpers
# ═══════════════════════════════════════════════════════════════════════════════

def _dedupe_cols(columns):
    seen = {}; out = []
    for c in columns:
        s = "" if c is None else str(c)
        if s in seen:
            seen[s] += 1; out.append(f"{s}.{seen[s]}")
        else:
            seen[s] = 0; out.append(s)
    return out


def _make_key(df: pd.DataFrame, cols: List[str]) -> pd.Series:
    if not cols:
        return pd.Series(range(len(df)), dtype=str)
    if len(cols) == 1:
        return df[cols[0]].fillna("").astype(str).str.strip().str.upper()
    parts = [df[c].fillna("").astype(str).str.strip().str.upper() for c in cols]
    r = parts[0]
    for p in parts[1:]:
        r = r + "|||" + p
    return r


_NULLS = {"", " ", "nan", "none", "nat", "<na>", "null", "n/a", "#n/a"}

def _norm_val(v):
    """Return None for any null-like value, else return v unchanged."""
    if v is None:
        return None
    if isinstance(v, float) and (v != v):   # NaN
        return None
    if isinstance(v, str) and v.strip().lower() in _NULLS:
        return None
    return v


def _extract_frame(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    """Pull cols from df, normalising nulls, returning object-dtype DataFrame."""
    data = {}
    for col in cols:
        if col in df.columns:
            arr = df[col].to_numpy(dtype=object)
            data[col] = np.array([_norm_val(v) for v in arr], dtype=object)
        else:
            data[col] = np.full(len(df), None, dtype=object)
    return pd.DataFrame(data)


# ── Date normalisation ─────────────────────────────────────────────────────────
# Recognise common date formats and normalise to YYYYMMDD string for comparison.

_RE_YYYYMMDD_RAW  = re.compile(r"^\d{8}$")                        # 20161231
_RE_YYYY_MM_DD    = re.compile(r"^(\d{4})-(\d{2})-(\d{2})")       # 2016-12-31[...]
_RE_MM_DD_YYYY    = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})")   # 12/31/2016
_RE_DD_MM_YYYY    = re.compile(r"^(\d{1,2})-(\d{1,2})-(\d{4})")   # 31-12-2016


def _to_date_key(s: str) -> str | None:
    """
    Try to convert a string to YYYYMMDD.
    Returns the YYYYMMDD string if successful, else None.
    """
    s = s.strip()
    if _RE_YYYYMMDD_RAW.match(s):
        return s                                                 # already YYYYMMDD
    m = _RE_YYYY_MM_DD.match(s)
    if m:
        return f"{m.group(1)}{m.group(2)}{m.group(3)}"          # 2016-12-31 → 20161231
    m = _RE_MM_DD_YYYY.match(s)
    if m:
        return f"{m.group(3)}{m.group(1).zfill(2)}{m.group(2).zfill(2)}"
    m = _RE_DD_MM_YYYY.match(s)
    if m:
        return f"{m.group(3)}{m.group(2).zfill(2)}{m.group(1).zfill(2)}"
    return None


def _normalise_for_compare(v) -> str | None:
    """
    Normalise a value for string comparison:
    - None → None
    - Numeric → strip trailing .0 if integer-valued, uppercase string
    - Date-like string → YYYYMMDD
    - Everything else → stripped uppercase string
    """
    if v is None:
        return None
    s = str(v).strip().upper()
    if s.lower() in _NULLS:
        return None

    # Try date normalisation
    dk = _to_date_key(str(v).strip())
    if dk:
        return dk

    # Numeric-looking: remove trailing ".0" for integers
    try:
        f = float(s.replace(",", ""))
        if f == int(f):
            return str(int(f))
        return str(f)
    except (ValueError, TypeError):
        pass

    return s


# ── Reconciliation functions ──────────────────────────────────────────────────

def _rec_key(sv, tv, ms, mt):
    """Key column REC: TRUE/FALSE/MISSING – vectorized with date normalisation."""
    sv_norm = _norm_series_for_compare(sv)
    tv_norm = _norm_series_for_compare(tv)
    sv_none = sv_norm == "__NULL__"
    tv_none = tv_norm == "__NULL__"
    both_none   = sv_none & tv_none
    either_none = (sv_none | tv_none) & ~both_none
    eq = sv_norm == tv_norm
    result = np.where(ms, MISSING_SRC,
             np.where(mt, MISSING_TARGET,
             np.where(both_none, "TRUE",
             np.where(either_none, "FALSE",
             np.where(eq, "TRUE", "FALSE")))))
    return result.astype(object)


def _to_num(arr: np.ndarray) -> np.ndarray:
    out = np.full(len(arr), np.nan)
    for i, v in enumerate(arr):
        if v is None:
            continue
        if isinstance(v, (int, float)):
            if not (isinstance(v, float) and v != v):
                out[i] = float(v)
        else:
            try:
                out[i] = float(str(v).replace(",", "").strip())
            except (ValueError, TypeError):
                pass
    return out


def _rec_fin(sv, tv, ms, mt, both, tol):
    sv_n = _to_num(sv); tv_n = _to_num(tv)
    raw  = tv_n - sv_n
    diff = np.where(np.abs(raw) <= tol, 0.0, raw)
    result = np.where(ms, MISSING_SRC, np.where(mt, MISSING_TARGET, diff.astype(object)))
    return result.astype(object), both & (np.abs(diff) > tol)


def _rec_nonfin(sv, tv, ms, mt, both):
    """Non-financial comparison with date and format normalisation – vectorized."""
    n = len(sv)

    # Vectorized normalisation via pandas
    sv_norm = _norm_series_for_compare(sv)
    tv_norm = _norm_series_for_compare(tv)

    sv_none = sv_norm == "__NULL__"
    tv_none = tv_norm == "__NULL__"
    both_none   = sv_none & tv_none
    either_none = (sv_none | tv_none) & ~both_none
    eq = sv_norm == tv_norm

    result = np.where(ms, MISSING_SRC,
             np.where(mt, MISSING_TARGET,
             np.where(both_none, "TRUE",
             np.where(either_none, "FALSE",
             np.where(eq, "TRUE", "FALSE")))))
    flag = both & (result == "FALSE")
    return result.astype(object), flag


def _norm_series_for_compare(arr: np.ndarray) -> np.ndarray:
    """
    Vectorized normalisation of an object array for string comparison.
    Returns string array; null values become the sentinel '__NULL__'.
    Dates are normalised to YYYYMMDD. Numbers strip trailing .0.
    """
    s = pd.Series(arr, dtype=object)

    # First: identify nulls
    null_mask = s.isna() | s.astype(str).str.strip().str.lower().isin(_NULLS)

    s_str = s.fillna("").astype(str).str.strip().str.upper()

    # --- Date normalisation (vectorized regex) ---
    # Only apply regex to values that look date-like (contain - or / or are 8 digits)
    looks_date = (
        s_str.str.contains(r"[-/]", regex=True) |
        s_str.str.match(r"^\d{8}$")
    )

    # Pattern: YYYYMMDD → keep
    yyyymmdd_ok = s_str.str.match(r"^\d{8}$")

    # Pattern: YYYY-MM-DD (with optional time suffix)
    iso_match = s_str.str.extract(r"^(\d{4})-(\d{2})-(\d{2})", expand=True)
    iso_ok = iso_match[0].notna() & looks_date
    s_iso  = iso_match[0].fillna("") + iso_match[1].fillna("") + iso_match[2].fillna("")

    # Pattern: MM/DD/YYYY
    mdy_match = s_str.str.extract(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", expand=True)
    mdy_ok = mdy_match[0].notna() & looks_date
    s_mdy  = mdy_match[2].fillna("") + mdy_match[0].fillna("").str.zfill(2) + mdy_match[1].fillna("").str.zfill(2)

    # Apply date conversions (ISO first, then MDY, skip if already YYYYMMDD)
    out = s_str.copy()
    out = np.where(iso_ok & ~yyyymmdd_ok, s_iso,  out)
    out = np.where(mdy_ok,                s_mdy,  out)
    out = pd.Series(out, dtype=object)

    # --- Numeric normalisation: strip trailing .0 for integers ---
    # Only on non-date-looking values to avoid converting dates
    num_like = pd.to_numeric(out, errors="coerce")
    int_mask = num_like.notna() & (num_like % 1 == 0) & ~looks_date
    out = np.where(int_mask, num_like.fillna(0).astype(np.int64).astype(str), out)
    out = pd.Series(out, dtype=object)

    # Apply null sentinel
    final = np.where(null_mask | (out.astype(str).str.strip() == ""), "__NULL__", out)
    return np.array(final, dtype=object)


# ═══════════════════════════════════════════════════════════════════════════════
# Workbook writer
# ═══════════════════════════════════════════════════════════════════════════════

def _write_workbook(buf, summary,
                    source_df, target_df,
                    src_frame, tgt_frame, rec_arrays,
                    all_cols, key_cols, fin_cols, non_fin_cols,
                    n_rows, tolerance,
                    src_only_cols, tgt_only_cols):

    wb = xlsxwriter.Workbook(buf, {
        "in_memory":          True,
        "constant_memory":    False,
        "strings_to_numbers": False,
        "use_zip64":          True,
        "nan_inf_to_errors":  True,
    })
    wb.strings_to_urls = False

    fmt = _fmts(wb)

    _plain_sheet(wb, "Source", source_df, fmt, key_cols, fin_cols)
    _plain_sheet(wb, "Target", target_df, fmt, key_cols, fin_cols)
    _data_rec_sheet(wb, src_frame, tgt_frame, rec_arrays,
                    all_cols, key_cols, fin_cols, n_rows, tolerance, fmt)
    _summary_sheet(wb, summary, fmt)

    # Column_Diff sheet only if there are unmatched columns
    if src_only_cols or tgt_only_cols:
        _col_diff_sheet(wb, src_only_cols, tgt_only_cols, fmt)

    wb.close()


def _fmts(wb):
    def f(**kw): return wb.add_format(kw)
    return {
        "h_src":  f(bold=True, font_color=C_WHITE, bg_color=C_SRC_H, border=1, align="center", valign="vcenter", text_wrap=True),
        "h_tgt":  f(bold=True, font_color=C_WHITE, bg_color=C_TGT_H, border=1, align="center", valign="vcenter", text_wrap=True),
        "h_rec":  f(bold=True, font_color=C_WHITE, bg_color=C_REC_H, border=1, align="center", valign="vcenter", text_wrap=True),
        "h_pln":  f(bold=True, font_color=C_WHITE, bg_color=C_HDR,   border=1, align="center", valign="vcenter", text_wrap=True),
        "h_sum":  f(bold=True, font_color=C_WHITE, bg_color=C_HDR,   border=1, align="center", valign="vcenter"),
        "h_cdiff":f(bold=True, font_color=C_WHITE, bg_color=C_ORANGE, border=1, align="center", valign="vcenter"),
        "m_lbl":  f(bold=True, border=1),
        "m_val":  f(border=1, align="right"),
        "m_warn": f(border=1, align="right", bg_color=C_YELLOW, font_color=C_YELLOW_F),
        "m_ok":   f(border=1, align="right", bg_color=C_GREEN,  font_color=C_GREEN_F),
        "c_src":  f(border=1, bg_color="#DDEBF7"),   # light blue – source only
        "c_tgt":  f(border=1, bg_color="#E2EFDA"),   # light green – target only
    }


def _cw(col, key_cols, fin_cols):
    if col in key_cols: return COL_KEY
    if col in fin_cols: return COL_FIN
    return COL_STD


def _write_col_typed(ws, col_idx, start_row, values, is_numeric=False):
    """Write column using typed XlsxWriter methods for speed."""
    if is_numeric:
        for ri, v in enumerate(values):
            r = start_row + ri
            if v is None:
                ws.write_blank(r, col_idx, None)
            elif isinstance(v, str):
                ws.write_string(r, col_idx, v)
            else:
                fv = float(v)
                if fv != fv or fv == float('inf') or fv == float('-inf'):
                    ws.write_blank(r, col_idx, None)
                else:
                    ws.write_number(r, col_idx, fv)
    else:
        for ri, v in enumerate(values):
            r = start_row + ri
            if v is None:
                ws.write_blank(r, col_idx, None)
            else:
                ws.write_string(r, col_idx, str(v))


def _plain_sheet(wb, name, df, fmt, key_cols, fin_cols):
    ws = wb.add_worksheet(name)
    cols = list(df.columns); nrows = len(df)
    ws.freeze_panes(1, 0)
    if nrows:
        ws.autofilter(0, 0, nrows, len(cols) - 1)
    fin_set = set(fin_cols)
    for ci, col in enumerate(cols):
        ws.write_string(0, ci, col, fmt["h_pln"])
        ws.set_column(ci, ci, _cw(col, key_cols, fin_cols))
        raw = df[col].where(df[col].notna(), other=None).tolist()
        is_num = col in fin_set
        if is_num:
            col_data = []
            for v in raw:
                if v is None:
                    col_data.append(None)
                else:
                    try: col_data.append(float(str(v).replace(",", "")))
                    except: col_data.append(v)
            _write_col_typed(ws, ci, 1, col_data, is_numeric=True)
        else:
            _write_col_typed(ws, ci, 1, raw, is_numeric=False)


def _data_rec_sheet(wb, src_frame, tgt_frame, rec_arrays,
                    all_cols, key_cols, fin_cols, n_rows, tolerance, fmt):
    ws = wb.add_worksheet("Data_Rec")
    n  = len(all_cols)
    fin_set = set(fin_cols)

    ws.freeze_panes(1, 0)
    if n_rows > 0:
        ws.autofilter(0, 0, n_rows, 3 * n - 1)

    for ci, col in enumerate(all_cols):
        ws.write_string(0, ci,       col, fmt["h_src"])
        ws.write_string(0, n + ci,   col, fmt["h_tgt"])
        ws.write_string(0, 2*n + ci, col, fmt["h_rec"])
        w = _cw(col, key_cols, fin_cols)
        ws.set_column(ci,       ci,       w)
        ws.set_column(n + ci,   n + ci,   w)
        ws.set_column(2*n + ci, 2*n + ci, w)

    if n_rows == 0:
        return

    for ci, col in enumerate(all_cols):
        is_num = col in fin_set

        # SRC section
        src_vals = src_frame[col].tolist()
        if is_num:
            data = []
            for v in src_vals:
                if v is None or (isinstance(v, float) and v != v):
                    data.append(None)
                else:
                    try: data.append(float(str(v).replace(",", "")))
                    except: data.append(None)
        else:
            data = [None if (v is None or (isinstance(v, float) and v != v))
                    else str(v) for v in src_vals]
        _write_col_typed(ws, ci, 1, data, is_numeric=is_num)

        # TGT section
        tgt_vals = tgt_frame[col].tolist()
        if is_num:
            data = []
            for v in tgt_vals:
                if v is None or (isinstance(v, float) and v != v):
                    data.append(None)
                else:
                    try: data.append(float(str(v).replace(",", "")))
                    except: data.append(None)
        else:
            data = [None if (v is None or (isinstance(v, float) and v != v))
                    else str(v) for v in tgt_vals]
        _write_col_typed(ws, n + ci, 1, data, is_numeric=is_num)

        # REC section
        rec_arr = rec_arrays.get(col, np.full(n_rows, None, dtype=object))
        if is_num:
            rec_data = []
            for v in rec_arr:
                if v is None:
                    rec_data.append(None)
                elif isinstance(v, str):
                    rec_data.append(v)   # MISSING IN SOURCE / TARGET
                else:
                    try: rec_data.append(float(v))
                    except: rec_data.append(None)
            _write_col_typed(ws, 2*n + ci, 1, rec_data, is_numeric=True)
        else:
            rec_data = [None if v is None else str(v) for v in rec_arr]
            _write_col_typed(ws, 2*n + ci, 1, rec_data, is_numeric=False)

    # Conditional formatting on REC section
    r1c = xlutil.xl_rowcol_to_cell(1,      2*n,     row_abs=True, col_abs=True)
    r1e = xlutil.xl_rowcol_to_cell(n_rows, 3*n - 1, row_abs=True, col_abs=True)
    cr  = f"{r1c}:{r1e}"

    def _cf(rng, criteria, value, bg, fg, bold=False, minimum=None, maximum=None):
        fd = {"bg_color": bg, "font_color": fg}
        if bold: fd["bold"] = True
        opt = {"type": "cell", "criteria": criteria, "format": wb.add_format(fd)}
        if criteria == "not between":
            opt["minimum"] = minimum; opt["maximum"] = maximum
        else:
            opt["value"] = value
        ws.conditional_format(rng, opt)

    _cf(cr, "equal to",    '"TRUE"',           C_GREEN,  C_GREEN_F)
    _cf(cr, "equal to",    0,                  C_GREEN,  C_GREEN_F)
    _cf(cr, "not between", None,               C_YELLOW, C_YELLOW_F, minimum=-tolerance, maximum=tolerance)
    _cf(cr, "equal to",    f'"{MISSING_SRC}"', C_RED,    C_RED_F,   bold=True)
    _cf(cr, "equal to",    f'"{MISSING_TARGET}"', C_RED,    C_RED_F,   bold=True)
    _cf(cr, "equal to",    '"FALSE"',          C_PINK,   C_PINK_F)


def _summary_sheet(wb, summary, fmt):
    ws = wb.add_worksheet("Summary")
    ws.freeze_panes(1, 0)
    ws.set_column(0, 0, 38); ws.set_column(1, 1, 22)
    ws.write_string(0, 0, "Metric", fmt["h_sum"])
    ws.write_string(0, 1, "Value",  fmt["h_sum"])
    WARN = {"Records Missing in Source", "Records Missing in TARGET",
            "Financial Differences", "Non-Financial Differences"}
    OK   = {"Matched Records", "Match Percentage"}
    for ri, (k, v) in enumerate(summary.items(), start=1):
        ws.write_string(ri, 0, k, fmt["m_lbl"])
        vfmt = fmt["m_warn"] if k in WARN else fmt["m_ok"] if k in OK else fmt["m_val"]
        if isinstance(v, (int, float)):
            ws.write_number(ri, 1, v, vfmt)
        else:
            ws.write_string(ri, 1, str(v), vfmt)


def _col_diff_sheet(wb, src_only: List[str], tgt_only: List[str], fmt):
    """
    Write a Column_Diff sheet listing columns that exist only in Source
    or only in Target (i.e. unmatched columns excluded from Data_Rec).
    """
    ws = wb.add_worksheet("Column_Diff")
    ws.freeze_panes(1, 0)
    ws.set_column(0, 0, 40)
    ws.set_column(1, 1, 30)

    ws.write_string(0, 0, "Column Name",    fmt["h_cdiff"])
    ws.write_string(0, 1, "Present In",     fmt["h_cdiff"])

    ri = 1
    for col in src_only:
        ws.write_string(ri, 0, col, fmt["c_src"])
        ws.write_string(ri, 1, "Source Only (Missing in TARGET)", fmt["c_src"])
        ri += 1
    for col in tgt_only:
        ws.write_string(ri, 0, col, fmt["c_tgt"])
        ws.write_string(ri, 1, "Target Only (Missing in Source)", fmt["c_tgt"])
        ri += 1
