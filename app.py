# app.py
import pandas as pd
import numpy as np
import streamlit as st
import io
import re
import unicodedata
from datetime import datetime

# ============================================================
# Normalization utilities
# ============================================================

def strip_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def norm_text(s: str) -> str:
    """Uppercase, remove accents, collapse spaces, remove common punctuation."""
    s = "" if s is None else str(s)
    s = s.strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[\t\r\n]+", " ", s)
    s = re.sub(r"[^\w\s]", " ", s)  # remove punctuation
    s = re.sub(r"\s+", " ", s)
    return s.upper().strip()

def normalize_dni_value(x):
    """Normalize DNI to digits only; handles floats like 12345678.0."""
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    if re.match(r"^\d+\.0$", s):
        s = s[:-2]
    s = re.sub(r"\D", "", s)
    return s if s else np.nan

def extract_year(value):
    """Extract year 20xx from Periodo like 2024-I, 2025-II, etc."""
    if pd.isna(value):
        return pd.NA
    m = re.search(r"(20\d{2})", str(value))
    return int(m.group(1)) if m else pd.NA

def score_to_numeric(series: pd.Series) -> pd.Series:
    """
    Convert values like '-', '20/20', '16,5', '20 pts' to numeric.
    """
    s = series.copy()
    if s.dtype == object:
        s = s.replace({"-": np.nan, "—": np.nan, "–": np.nan, "": np.nan})
        s = s.astype(str).str.strip()
        s = s.replace({"nan": np.nan, "None": np.nan})
        s = s.str.replace(",", ".", regex=False)
        s = s.str.extract(r"(-?\d+(?:\.\d+)?)", expand=False)
    return pd.to_numeric(s, errors="coerce")

# ============================================================
# Sheet + column finding (robust)
# ============================================================

def find_sheet(xls: pd.ExcelFile, desired: str):
    want = norm_text(desired)
    for real in xls.sheet_names:
        if norm_text(real) == want:
            return real
    # contains fallback
    for real in xls.sheet_names:
        if want in norm_text(real):
            return real
    return None

def safe_read_sheet(xls: pd.ExcelFile, sheet_name: str, warnings: list[str]) -> pd.DataFrame:
    real = find_sheet(xls, sheet_name)
    if real is None:
        warnings.append(f"Missing sheet: '{sheet_name}'")
        return pd.DataFrame()
    return strip_cols(pd.read_excel(xls, sheet_name=real))

def find_col_by_keywords(df: pd.DataFrame, keyword_sets: list[list[str]]):
    """
    Find a column whose normalized header contains ALL keywords in any keyword_set.
    Example keyword_sets = [["TAREA", "PRODUCTO", "FINAL"], ["PROMEDIO"]]
    """
    if df is None or df.empty:
        return None
    norm_headers = {c: norm_text(c) for c in df.columns}
    for keyset in keyword_sets:
        keyset = [norm_text(k) for k in keyset]
        for col, h in norm_headers.items():
            if all(k in h for k in keyset):
                return col
    return None

def find_identity_columns(df: pd.DataFrame):
    """
    Try to locate DNI and name columns in many possible formats.
    Returns (dni_col, nombre_col, apellidos_col, ape_pat_col, ape_mat_col, nombre_completo_col)
    """
    if df is None or df.empty:
        return (None, None, None, None, None, None)

    dni_col = find_col_by_keywords(df, [
        ["DNI"],
        ["DOCUMENTO", "IDENTIDAD"],
        ["NRO", "DOCUMENTO"],
        ["N", "DOCUMENTO"],
        ["DOCUMENTO"]
    ])

    nombre_col = find_col_by_keywords(df, [
        ["NOMBRE"],
        ["NOMBRES"],
    ])

    apellidos_col = find_col_by_keywords(df, [
        ["APELLIDO"],
        ["APELLIDOS"],
        ["APELLIDO", "S"]
    ])

    ape_pat_col = find_col_by_keywords(df, [
        ["APELLIDO", "PATERNO"],
        ["APE", "PATERNO"]
    ])

    ape_mat_col = find_col_by_keywords(df, [
        ["APELLIDO", "MATERNO"],
        ["APE", "MATERNO"]
    ])

    nombre_completo_col = find_col_by_keywords(df, [
        ["NOMBRE", "COMPLETO"],
        ["APELLIDOS", "NOMBRES"],
        ["NOMBRES", "APELLIDOS"]
    ])

    return (dni_col, nombre_col, apellidos_col, ape_pat_col, ape_mat_col, nombre_completo_col)

def build_name_fields(df: pd.DataFrame):
    """
    Standardize to columns: Nombre, Apellido(s)
    Accepts various formats:
    - Nombre + Apellido(s)
    - Nombres + Apellidos
    - Apellido paterno + materno
    - Nombre completo
    """
    df = df.copy()
    dni_col, nombre_col, apellidos_col, ape_pat_col, ape_mat_col, nombre_completo_col = find_identity_columns(df)

    # Create Nombre/Apellido(s) if possible
    if "Nombre" not in df.columns:
        if nombre_col:
            df["Nombre"] = df[nombre_col].astype(str).str.strip()
        else:
            df["Nombre"] = np.nan

    if "Apellido(s)" not in df.columns:
        if apellidos_col:
            df["Apellido(s)"] = df[apellidos_col].astype(str).str.strip()
        elif ape_pat_col or ape_mat_col:
            ap = df[ape_pat_col].astype(str).str.strip() if ape_pat_col else ""
            am = df[ape_mat_col].astype(str).str.strip() if ape_mat_col else ""
            df["Apellido(s)"] = (ap + " " + am).str.replace(r"\s+", " ", regex=True).str.strip()
        elif nombre_completo_col:
            # best-effort split: last 2 tokens as surname(s) is risky, so keep full as Apellido(s) empty
            full = df[nombre_completo_col].astype(str).str.strip()
            df["Apellido(s)"] = np.nan
            df["Nombre"] = full
        else:
            df["Apellido(s)"] = np.nan

    # DNI
    if "DNI" not in df.columns:
        if dni_col:
            df["DNI"] = df[dni_col].apply(normalize_dni_value)
        else:
            df["DNI"] = np.nan
    else:
        df["DNI"] = df["DNI"].apply(normalize_dni_value)

    return df

def add_name_keys(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Nombre_key"] = df["Nombre"].apply(norm_text) if "Nombre" in df.columns else ""
    df["Apellido_key"] = df["Apellido(s)"].apply(norm_text) if "Apellido(s)" in df.columns else ""
    return df

# ============================================================
# Contract mapping (critical to fill DNI when missing)
# ============================================================

def load_teacher_contract(contract_file, warnings: list[str]) -> pd.DataFrame:
    df = strip_cols(pd.read_excel(contract_file, sheet_name=0))
    df = build_name_fields(df)

    # Require DNI + name
    if df["DNI"].isna().all():
        warnings.append("Teacher Contract: DNI column not found or empty.")
        return pd.DataFrame(columns=["DNI", "Nombre", "Apellido(s)", "Nombre_key", "Apellido_key"])

    df["Nombre"] = df["Nombre"].astype(str).str.strip()
    df["Apellido(s)"] = df["Apellido(s)"].astype(str).str.strip()
    df = df.dropna(subset=["DNI"]).drop_duplicates(subset=["DNI"])
    df = add_name_keys(df)
    return df[["DNI", "Nombre", "Apellido(s)", "Nombre_key", "Apellido_key"]]

def build_contract_lookup(contract: pd.DataFrame):
    """
    Build multiple matching keys -> DNI to map rows that have no DNI.
    Keys:
      1) full:  Nombre_key|Apellido_key
      2) swapped surnames (if 2 tokens): Nombre_key|swap(Apellido_key)
      3) first surname only: Nombre_key|first_token(Apellido_key)
    Only keep keys that map uniquely to ONE DNI.
    """
    def swap_surnames(ap_key: str):
        toks = ap_key.split()
        if len(toks) >= 2:
            toks[0], toks[1] = toks[1], toks[0]
        return " ".join(toks)

    def first_surname(ap_key: str):
        toks = ap_key.split()
        return toks[0] if toks else ""

    c = contract.copy()
    c["k_full"] = c["Nombre_key"] + "|" + c["Apellido_key"]
    c["k_swap"] = c["Nombre_key"] + "|" + c["Apellido_key"].apply(swap_surnames)
    c["k_first"] = c["Nombre_key"] + "|" + c["Apellido_key"].apply(first_surname)

    lookups = {}
    for kcol in ["k_full", "k_swap", "k_first"]:
        grp = c.groupby(kcol)["DNI"].nunique()
        unique_keys = grp[grp == 1].index
        tmp = c[c[kcol].isin(unique_keys)][[kcol, "DNI"]].drop_duplicates()
        lookups[kcol] = dict(zip(tmp[kcol], tmp["DNI"]))
    return lookups

def map_missing_dni_from_contract(df: pd.DataFrame, contract: pd.DataFrame, lookups: dict):
    """
    Fill df.DNI when missing using contract mappings and normalized names.
    """
    df = df.copy()
    df = add_name_keys(df)

    def swap_surnames(ap_key: str):
        toks = ap_key.split()
        if len(toks) >= 2:
            toks[0], toks[1] = toks[1], toks[0]
        return " ".join(toks)

    def first_surname(ap_key: str):
        toks = ap_key.split()
        return toks[0] if toks else ""

    missing = df["DNI"].isna() | (df["DNI"].astype(str).str.strip() == "")
    if missing.any():
        k_full = df["Nombre_key"] + "|" + df["Apellido_key"]
        k_swap = df["Nombre_key"] + "|" + df["Apellido_key"].apply(swap_surnames)
        k_first = df["Nombre_key"] + "|" + df["Apellido_key"].apply(first_surname)

        dni_full = k_full.map(lookups.get("k_full", {}))
        dni_swap = k_swap.map(lookups.get("k_swap", {}))
        dni_first = k_first.map(lookups.get("k_first", {}))

        # fill in priority order
        df.loc[missing, "DNI"] = df.loc[missing, "DNI"].combine_first(dni_full[missing])
        df.loc[missing, "DNI"] = df.loc[missing, "DNI"].combine_first(dni_swap[missing])
        df.loc[missing, "DNI"] = df.loc[missing, "DNI"].combine_first(dni_first[missing])

    return df

# ============================================================
# Component extraction + merge
# ============================================================

def aggregate_component(df: pd.DataFrame, value_col: str, out_col: str):
    """Return two columns: DNI, out_col with max score per DNI."""
    tmp = df[["DNI", value_col]].copy()
    tmp["DNI"] = tmp["DNI"].apply(normalize_dni_value)
    tmp[out_col] = score_to_numeric(tmp[value_col])
    tmp = tmp.dropna(subset=["DNI"])
    if tmp.empty:
        return pd.DataFrame(columns=["DNI", out_col])
    tmp = tmp.groupby("DNI", as_index=False)[out_col].max()
    return tmp

def merge_agg(all_data: pd.DataFrame, agg: pd.DataFrame, out_col: str):
    """Left merge on DNI and coalesce."""
    if out_col not in all_data.columns:
        all_data[out_col] = np.nan
    merged = all_data.merge(agg, on="DNI", how="left", suffixes=("", "__new"))
    if f"{out_col}__new" in merged.columns:
        merged[out_col] = merged[out_col].combine_first(merged[f"{out_col}__new"])
        merged = merged.drop(columns=[f"{out_col}__new"])
    return merged

# ============================================================
# Main extraction
# ============================================================

def extract_data_from_excel(master_file, contract_file=None):
    warnings = []
    xls = pd.ExcelFile(master_file)

    # Load contract first (so we can map DNIs everywhere)
    contract = None
    lookups = None
    if contract_file is not None:
        contract = load_teacher_contract(contract_file, warnings)
        if not contract.empty:
            lookups = build_contract_lookup(contract)
        else:
            contract = None
            lookups = None

    # Read master sheets
    induction_df       = safe_read_sheet(xls, "Inducción", warnings)
    nota_induccion_df  = safe_read_sheet(xls, "nota Inducción", warnings)
    bus_df             = safe_read_sheet(xls, "Bus. biblioteca", warnings)
    diseno_df          = safe_read_sheet(xls, "Diseño de sesión", warnings)
    comp_df            = safe_read_sheet(xls, "Comp. Tec", warnings)
    integracion_df     = safe_read_sheet(xls, "Integración", warnings)
    rsu_df             = safe_read_sheet(xls, "RSU", warnings)
    estress_df         = safe_read_sheet(xls, "estress", warnings)
    hab_df             = safe_read_sheet(xls, "Hab. comunicación", warnings)

    # -----------------------------
    # Build base from Inducción + nota Inducción
    # -----------------------------
    base_parts = []

    def build_base_from(df: pd.DataFrame, periodo_keys, score_keys, label):
        if df is None or df.empty:
            warnings.append(f"{label}: empty/missing")
            return None
        df = strip_cols(df)
        df = build_name_fields(df)

        periodo_col = find_col_by_keywords(df, periodo_keys)
        score_col = find_col_by_keywords(df, score_keys)

        if periodo_col is None or score_col is None:
            warnings.append(f"{label}: Periodo or score column not found")
            return None

        out = df.copy()
        out["Periodo"] = out[periodo_col]
        out["induccion"] = score_to_numeric(out[score_col])

        out = out[["Periodo", "DNI", "Nombre", "Apellido(s)", "induccion"]].copy()
        out["Year"] = out["Periodo"].apply(extract_year)

        # Map missing DNI from contract if available
        if contract is not None and lookups is not None:
            out = map_missing_dni_from_contract(out, contract, lookups)

        # Fill names from contract using DNI (canonical)
        if contract is not None:
            out = out.merge(contract[["DNI", "Nombre", "Apellido(s)"]], on="DNI", how="left", suffixes=("", "__c"))
            out["Nombre"] = out["Nombre"].combine_first(out["Nombre__c"])
            out["Apellido(s)"] = out["Apellido(s)"].combine_first(out["Apellido(s)__c"])
            out = out.drop(columns=["Nombre__c", "Apellido(s)__c"], errors="ignore")

        return out

    # nota Inducción
    b1 = build_base_from(
        nota_induccion_df,
        periodo_keys=[["PERIODO"], ["PERIODO", "ACADEMICO"], ["PERIODO", "CURSO"]],
        score_keys=[["TOTAL", "CURSO"], ["TOTAL"], ["PUNTAJE"], ["CALIFICACION"]],
        label="nota Inducción"
    )
    if b1 is not None:
        base_parts.append(b1)

    # Inducción
    b2 = build_base_from(
        induction_df,
        periodo_keys=[["PERIODO"], ["PERIODO", "ACADEMICO"], ["PERIODO", "CURSO"], ["PERIODO", "I"]],
        score_keys=[["CALIFICACION"], ["CALIFICACION", "REAL"], ["NOTA"], ["PUNTAJE"], ["SCORE"]],
        label="Inducción"
    )
    if b2 is not None:
        base_parts.append(b2)

    if not base_parts:
        return pd.DataFrame(), warnings + ["No base data could be extracted from Inducción sheets."]

    all_data = pd.concat(base_parts, ignore_index=True)

    # Keep only DNIs that exist (after mapping)
    all_data["DNI"] = all_data["DNI"].apply(normalize_dni_value)
    all_data = all_data.dropna(subset=["DNI"])

    # If contract exists, restrict to contract teachers
    if contract is not None:
        all_data = all_data[all_data["DNI"].isin(contract["DNI"])]

    # -----------------------------
    # Prepare numeric columns (start NaN to allow merge fill)
    # -----------------------------
    numeric_columns = [
        "induccion", "bus_biblioteca", "diseno_sesion",
        "Zoom_basico", "Zoom_Avanzado", "Grupos_Moodle", "Rubrica",
        "Padlet", "Nearpod", "Tareas_y_foros",
        "integracion", "rsu", "estress", "hab_comunicacion"
    ]
    for c in numeric_columns:
        if c not in all_data.columns:
            all_data[c] = np.nan

    # -----------------------------
    # Single-score component sheets
    # -----------------------------
    def process_component(sheet_df, out_col, value_keyword_sets, sheet_label):
        if sheet_df is None or sheet_df.empty:
            warnings.append(f"{sheet_label}: empty/missing -> {out_col} stays empty")
            return

        df = strip_cols(sheet_df)
        df = build_name_fields(df)

        # map missing DNI from contract
        if contract is not None and lookups is not None:
            df = map_missing_dni_from_contract(df, contract, lookups)

        val_col = find_col_by_keywords(df, value_keyword_sets)
        if val_col is None:
            warnings.append(f"{sheet_label}: score column not found -> {out_col} stays empty")
            return

        agg = aggregate_component(df, val_col, out_col)
        if agg.empty:
            warnings.append(f"{sheet_label}: no DNI-mapped rows found -> {out_col} stays empty")
            return

        nonlocal all_data
        all_data = merge_agg(all_data, agg, out_col)

    process_component(bus_df, "bus_biblioteca", [["PROMEDIO"], ["PROM", "EDIO"], ["NOTA"]], "Bus. biblioteca")
    process_component(diseno_df, "diseno_sesion", [["PROMEDIO"], ["NOTA"], ["CALIFICACION"]], "Diseño de sesión")
    process_component(integracion_df, "integracion", [["PRODUCTO", "FINAL"], ["TAREA", "PRODUCTO", "FINAL"]], "Integración")
    process_component(rsu_df, "rsu", [["PRODUCTO", "FINAL"], ["TAREA", "PRODUCTO", "FINAL"]], "RSU")
    process_component(estress_df, "estress", [["PRODUCTO", "FINAL"], ["TAREA", "PRODUCTO", "FINAL"]], "estress")
    process_component(hab_df, "hab_comunicacion", [["PRODUCTO", "FINAL"], ["TAREA", "PRODUCTO", "FINAL"]], "Hab. comunicación")

    # -----------------------------
    # Comp. Tec (multi-columns) - keyword-based column detection
    # -----------------------------
    if comp_df is None or comp_df.empty:
        warnings.append("Comp. Tec: empty/missing -> all Comp. Tec columns stay empty")
    else:
        df = strip_cols(comp_df)
        df = build_name_fields(df)

        # map missing DNI from contract (this is key!)
        if contract is not None and lookups is not None:
            df = map_missing_dni_from_contract(df, contract, lookups)

        # For each component, find the best matching column by keywords
        comp_specs = {
            "Zoom_basico": [["ZOOM", "BASICO"], ["ZOOM", "BASIC"]],
            "Zoom_Avanzado": [["ZOOM", "AVANZADO"], ["ZOOM", "ADVANCED"]],
            "Grupos_Moodle": [["GRUPOS", "MOODLE"], ["GROUPS", "MOODLE"]],
            "Rubrica": [["RUBRICA"], ["RUBRICA", "RETO"]],
            "Padlet": [["PADLET"]],
            "Nearpod": [["NEARPOD"]],
            "Tareas_y_foros": [["TAREAS", "FOROS"], ["TAREAS", "Y", "FOROS"], ["FOROS"], ["TAREAS"]],
        }

        # Build a per-DNI aggregated table with all comp columns
        df["DNI"] = df["DNI"].apply(normalize_dni_value)
        df = df.dropna(subset=["DNI"])

        if df.empty:
            warnings.append("Comp. Tec: no DNI-mapped rows found -> Comp. Tec columns stay empty")
        else:
            agg = pd.DataFrame({"DNI": df["DNI"]}).drop_duplicates()

            for out_col, kw_sets in comp_specs.items():
                col = find_col_by_keywords(df, kw_sets)
                if col is None:
                    warnings.append(f"Comp. Tec: column not found for {out_col}")
                    continue
                tmp = df[["DNI", col]].copy()
                tmp[out_col] = score_to_numeric(tmp[col])
                tmp = tmp.drop(columns=[col])
                tmp = tmp.groupby("DNI", as_index=False)[out_col].max()
                agg = agg.merge(tmp, on="DNI", how="left")

            # merge into all_data and coalesce
            for out_col in comp_specs.keys():
                if out_col not in agg.columns:
                    continue
                all_data = merge_agg(all_data, agg[["DNI", out_col]], out_col)

    # -----------------------------
    # Convert numeric columns to numbers (NaN -> 0) ONLY AT THE END
    # -----------------------------
    for col in numeric_columns:
        all_data[col] = score_to_numeric(all_data[col]).fillna(0)

    # -----------------------------
    # Metrics
    # -----------------------------
    all_data["Average"] = all_data[numeric_columns].mean(axis=1).round(2)
    all_data["Percentage"] = ((all_data[numeric_columns] > 0).sum(axis=1) / len(numeric_columns) * 100).round(2)
    all_data["Marks_Out_Of_20"] = (all_data["Percentage"] / 5).round(2)

    # -----------------------------
    # Filter 2024 vs 2025 and keep rows with any score
    # -----------------------------
    all_data["Year"] = all_data["Year"].astype("Int64")
    filtered = all_data[all_data["Year"].isin([2024, 2025])].copy()
    filtered = filtered[filtered[numeric_columns].sum(axis=1) > 0]

    if filtered.empty:
        return pd.DataFrame(), warnings + ["No records with scores found for 2024 or 2025 after merges."]

    # Dedup: keep best row per teacher (DNI)
    filtered["YearPref"] = filtered["Year"].apply(lambda y: 1 if y == 2025 else 0)
    sorted_df = filtered.sort_values(
        by=["Marks_Out_Of_20", "Average", "YearPref"],
        ascending=[False, False, False]
    )
    highest = sorted_df.drop_duplicates(subset=["DNI"], keep="first").copy()
    highest["Highest_Score_Year"] = highest["Year"]

    final_columns = [
        "Periodo", "Highest_Score_Year", "DNI", "Nombre", "Apellido(s)",
        *numeric_columns, "Average", "Marks_Out_Of_20", "Percentage"
    ]
    for c in final_columns:
        if c not in highest.columns:
            highest[c] = np.nan

    return highest[final_columns], warnings

# ============================================================
# Streamlit UI
# ============================================================

def main():
    st.set_page_config(page_title="📊 UMA Scores (Highest of 2024 vs 2025)", page_icon="📊", layout="wide")

    st.title("📊 UMA Scores — Highest Marks (2024 vs 2025)")
    st.markdown(
        "Upload the **Master** Excel and the **Teacher Contract** Excel.\n\n"
        "This version fixes missing/empty DNIs in Master sheets by mapping teachers using the Contract list, "
        "and detects columns using keywords (Nearpod/Padlet/etc.)."
    )

    uploaded_master = st.file_uploader("Choose the Master Excel file", type=["xlsx", "xls"], key="master")
    uploaded_contract = st.file_uploader("Choose the Teacher Contract Excel file", type=["xlsx", "xls"], key="contract")

    if uploaded_master is not None and uploaded_contract is not None:
        try:
            with st.spinner("Processing..."):
                final_df, warnings = extract_data_from_excel(uploaded_master, contract_file=uploaded_contract)

            with st.expander("⚠️ Diagnostics / Warnings"):
                if warnings:
                    for w in warnings:
                        st.warning(w)
                else:
                    st.success("No warnings.")

            if final_df.empty:
                st.warning("No records produced (after cleaning/merging).")
                return

            st.success("Done! Highest row per teacher selected.")

            st.subheader("Preview")
            st.dataframe(final_df.head(30), use_container_width=True)

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Teachers", len(final_df))
            with col2:
                st.metric("Avg Marks (Out of 20)", f"{final_df['Marks_Out_Of_20'].mean():.2f}")
            with col3:
                st.metric("Avg Percentage", f"{final_df['Percentage'].mean():.2f}%")

            st.subheader("Distribution by Highest Score Year")
            st.bar_chart(final_df["Highest_Score_Year"].value_counts().sort_index())

            # Download
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_name = f"Highest_Marks_2024_vs_2025_{timestamp}.xlsx"
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_df.to_excel(writer, index=False, sheet_name="Highest Marks (Unique)")
            output.seek(0)

            st.download_button(
                label="📥 Download Result",
                data=output,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {str(e)}")

    else:
        st.info("👆 Upload both files to start.")

if __name__ == "__main__":
    main()
