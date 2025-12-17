# app.py
import pandas as pd
import numpy as np
import streamlit as st
import io
import re
import unicodedata
from datetime import datetime

# -----------------------------
# Normalization helpers
# -----------------------------
def _strip_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _norm_text(s: str) -> str:
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s)
    return s.upper()

def _norm_header(s: str) -> str:
    return _norm_text(s)

def normalize_dni_value(x):
    """Normalize DNI to plain digits; handles floats like '12345678.0'."""
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    if re.match(r"^\d+\.0$", s):
        s = s[:-2]
    s = re.sub(r"\D", "", s)
    return s if s else np.nan

def extract_year(value):
    """Extract 20xx year from Periodo like '2025-II'."""
    if pd.isna(value):
        return pd.NA
    m = re.search(r"(20\d{2})", str(value))
    return int(m.group(1)) if m else pd.NA

def norm_person_key(s):
    """Normalize names for merges: no accents, no double spaces, uppercase."""
    if pd.isna(s):
        return ""
    return _norm_text(s)

def ensure_person_keys(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Nombre_key"] = df["Nombre"].apply(norm_person_key) if "Nombre" in df.columns else ""
    df["Apellido_key"] = df["Apellido(s)"].apply(norm_person_key) if "Apellido(s)" in df.columns else ""
    return df

def find_col(df: pd.DataFrame, candidates: list[str]):
    """Find a column in df matching any candidate (accent/case/space tolerant)."""
    if df is None or df.empty:
        return None
    norm_map = {_norm_header(c): c for c in df.columns}
    for cand in candidates:
        key = _norm_header(cand)
        if key in norm_map:
            return norm_map[key]
    # fallback: contains match
    for c in df.columns:
        hc = _norm_header(c)
        for cand in candidates:
            if _norm_header(cand) in hc:
                return c
    return None

def score_to_numeric(series: pd.Series) -> pd.Series:
    """
    Converts values like '-', '20/20', ' 16 ', '16,5', '20 pts' into numeric.
    Non-numeric -> NaN.
    """
    s = series.copy()
    if s.dtype == object:
        s = s.replace({"-": np.nan, "—": np.nan, "–": np.nan, "": np.nan})
        s = s.astype(str).str.strip().replace({"nan": np.nan, "None": np.nan})
        s = s.str.replace(",", ".", regex=False)
        s = s.str.extract(r"(-?\d+(?:\.\d+)?)", expand=False)  # first numeric chunk
    return pd.to_numeric(s, errors="coerce")

# -----------------------------
# Sheet reading (tolerant to accents/spaces)
# -----------------------------
def _find_sheet_name(xls: pd.ExcelFile, desired: str):
    want = _norm_text(desired)
    for real in xls.sheet_names:
        if _norm_text(real) == want:
            return real
    # fallback: contains
    for real in xls.sheet_names:
        if want in _norm_text(real):
            return real
    return None

def safe_read_sheet(xls: pd.ExcelFile, sheet_name: str, warnings: list[str]) -> pd.DataFrame:
    real = _find_sheet_name(xls, sheet_name)
    if real is None:
        warnings.append(f"Missing sheet: '{sheet_name}'")
        return pd.DataFrame()
    df = pd.read_excel(xls, sheet_name=real)
    return _strip_cols(df)

# -----------------------------
# Teacher contract loader
# -----------------------------
def load_teacher_contract(contract_file, warnings: list[str]) -> pd.DataFrame:
    df = _strip_cols(pd.read_excel(contract_file, sheet_name=0))

    dni_col = find_col(df, [
        "N° DE DOCUMENTO DE IDENTIDAD",
        "NRO DE DOCUMENTO DE IDENTIDAD",
        "NRO DE DOCUMENTO",
        "DOCUMENTO DE IDENTIDAD",
        "DNI"
    ])
    name_col = find_col(df, ["NOMBRES", "NOMBRE", "Nombre", "Nombres"])
    a_pat_col = find_col(df, ["APELLIDO PATERNO", "Apellido Paterno"])
    a_mat_col = find_col(df, ["APELLIDO MATERNO", "Apellido Materno"])

    if not all([dni_col, name_col, a_pat_col, a_mat_col]):
        warnings.append("Teacher Contract: missing required columns (DNI/NOMBRES/APELLIDOS).")
        return pd.DataFrame(columns=["DNI", "Nombre", "Apellido(s)"])

    out = pd.DataFrame({
        "DNI": df[dni_col].apply(normalize_dni_value),
        "Nombre": df[name_col].astype(str).str.strip(),
        "Apellido(s)": (
            df[a_pat_col].astype(str).str.strip() + " " +
            df[a_mat_col].astype(str).str.strip()
        ).str.replace(r"\s+", " ", regex=True).str.strip()
    })
    out = out.dropna(subset=["DNI"])
    out = out[out["DNI"] != ""]
    out = out.drop_duplicates(subset=["DNI"])
    return out[["DNI", "Nombre", "Apellido(s)"]]

# -----------------------------
# Merge component with DNI then Name fallback
# -----------------------------
def merge_component_dual(all_data: pd.DataFrame,
                         comp_df: pd.DataFrame,
                         out_col: str,
                         value_candidates: list[str],
                         warnings: list[str]):
    """
    Merge comp_df into all_data:
    1) merge by DNI if possible
    2) fill remaining missing by Nombre+Apellido(s)
    Keep NaN until final numeric step (do NOT prefill with 0).
    """
    if out_col not in all_data.columns:
        all_data[out_col] = np.nan

    if comp_df is None or comp_df.empty:
        warnings.append(f"{out_col}: sheet empty/missing")
        return all_data

    comp_df = _strip_cols(comp_df)
    value_col = find_col(comp_df, value_candidates)
    if value_col is None:
        warnings.append(f"{out_col}: score column not found")
        return all_data

    # --- (1) DNI merge
    dni_col = find_col(comp_df, ["DNI", "NRO DE DOCUMENTO", "DOCUMENTO", "DOCUMENTO DE IDENTIDAD"])
    if dni_col is not None:
        tmp = comp_df[[dni_col, value_col]].copy()
        tmp.columns = ["DNI_raw", out_col]
        tmp["DNI"] = tmp["DNI_raw"].apply(normalize_dni_value)
        tmp[out_col] = score_to_numeric(tmp[out_col])
        tmp = tmp.dropna(subset=["DNI"])
        if not tmp.empty:
            tmp = tmp.groupby("DNI", as_index=False)[out_col].max()
            merged = all_data.merge(tmp, on="DNI", how="left", suffixes=("", "__dni"))
            # coalesce
            all_data[out_col] = all_data[out_col].combine_first(merged[f"{out_col}__dni"])
            all_data = merged.drop(columns=[f"{out_col}__dni"])

    # --- (2) Name merge fallback
    name_col = find_col(comp_df, ["Nombre", "NOMBRES"])
    ap_col = find_col(comp_df, ["Apellido(s)", "Apellidos", "APELLIDOS", "APELLIDO(S)"])
    if name_col is not None and ap_col is not None:
        tmp = comp_df[[name_col, ap_col, value_col]].copy()
        tmp.columns = ["Nombre", "Apellido(s)", out_col]
        tmp[out_col] = score_to_numeric(tmp[out_col])
        tmp = ensure_person_keys(tmp)
        tmp = tmp.groupby(["Nombre_key", "Apellido_key"], as_index=False)[out_col].max()

        all_data = ensure_person_keys(all_data)
        merged = all_data.merge(
            tmp[["Nombre_key", "Apellido_key", out_col]],
            on=["Nombre_key", "Apellido_key"],
            how="left",
            suffixes=("", "__name")
        )
        all_data[out_col] = all_data[out_col].combine_first(merged[f"{out_col}__name"])
        all_data = merged.drop(columns=[f"{out_col}__name"])

    return all_data

# -----------------------------
# Core processing
# -----------------------------
def extract_data_from_excel(master_file, contract_file=None):
    warnings = []
    xls = pd.ExcelFile(master_file)

    # Read sheets safely
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
    # Base: Inducción + nota Inducción -> 'induccion'
    # -----------------------------
    base_parts = []

    # nota Inducción
    if not nota_induccion_df.empty:
        per_col   = find_col(nota_induccion_df, ["PERIODO", "Periodo"])
        dni_col   = find_col(nota_induccion_df, ["DNI", "DOCUMENTO", "NRO DE DOCUMENTO"])
        nom_col   = find_col(nota_induccion_df, ["Nombre", "NOMBRES"])
        ape_col   = find_col(nota_induccion_df, ["Apellido(s)", "Apellidos", "APELLIDOS"])
        mail_col  = find_col(nota_induccion_df, ["Dirección de correo", "Direccion de correo", "Correo", "Email"])
        score_col = find_col(nota_induccion_df, ["Total del curso (Real)", "Total del curso", "Total"])

        if all([per_col, dni_col, nom_col, ape_col, score_col]):
            cols = [per_col, dni_col, nom_col, ape_col] + ([mail_col] if mail_col else []) + [score_col]
            tmp = nota_induccion_df[cols].copy()
            tmp.columns = ["Periodo", "DNI", "Nombre", "Apellido(s)"] + (["Dirección de correo"] if mail_col else []) + ["induccion"]
            base_parts.append(tmp)
        else:
            warnings.append("nota Inducción: missing required columns (skipped).")

    # Inducción
    if not induction_df.empty:
        per_col   = find_col(induction_df, ["Periodo", "PERIODO"])
        dni_col   = find_col(induction_df, ["DNI", "DOCUMENTO", "NRO DE DOCUMENTO"])
        nom_col   = find_col(induction_df, ["Nombre", "NOMBRES"])
        ape_col   = find_col(induction_df, ["Apellido(s)", "Apellidos", "APELLIDOS"])
        mail_col  = find_col(induction_df, ["Dirección de correo", "Direccion de correo", "Correo", "Email"])
        score_col = find_col(induction_df, ["Calificación", "Calificacion", "Nota", "Score", "Puntaje"])

        if all([per_col, dni_col, nom_col, ape_col, score_col]):
            cols = [per_col, dni_col, nom_col, ape_col] + ([mail_col] if mail_col else []) + [score_col]
            tmp = induction_df[cols].copy()
            tmp.columns = ["Periodo", "DNI", "Nombre", "Apellido(s)"] + (["Dirección de correo"] if mail_col else []) + ["induccion"]
            base_parts.append(tmp)
        else:
            warnings.append("Inducción: missing required columns (skipped).")

    if not base_parts:
        return pd.DataFrame(), warnings + ["No base data found in Inducción/nota Inducción."]

    all_data = pd.concat(base_parts, ignore_index=True)

    # Normalize base
    all_data["DNI"] = all_data["DNI"].apply(normalize_dni_value)
    all_data["Year"] = all_data["Periodo"].apply(extract_year)

    # -----------------------------
    # Merge single-score sheets (DNI then Name fallback)
    # -----------------------------
    all_data = merge_component_dual(all_data, bus_df, "bus_biblioteca", ["Promedio", "promedio", "Nota"], warnings)
    all_data = merge_component_dual(all_data, diseno_df, "diseno_sesion", ["Promedio", "promedio", "Nota"], warnings)

    all_data = merge_component_dual(
        all_data,
        integracion_df,
        "integracion",
        [
            "Tarea:Producto final: Contenido académico, presentación y rúbrica con IA (Real)",
            "Tarea: Producto final: Contenido académico, presentación y rúbrica con IA (Real)",
            "Producto final",
            "Nota",
            "Score"
        ],
        warnings
    )

    all_data = merge_component_dual(all_data, rsu_df, "rsu",
                                   ["Tarea: Producto final", "Tarea:Producto final", "Producto final", "Nota"],
                                   warnings)

    all_data = merge_component_dual(all_data, estress_df, "estress",
                                   ["Tarea:Producto final", "Tarea: Producto final", "Producto final", "Nota"],
                                   warnings)

    all_data = merge_component_dual(all_data, hab_df, "hab_comunicacion",
                                   ["Tarea:Producto final", "Tarea: Producto final", "Producto final", "Nota"],
                                   warnings)

    # -----------------------------
    # Comp. Tec (multi-column) — IMPORTANT FIX: start as NaN, then coalesce
    # -----------------------------
    comp_map = {
        "Cuestionario:Reto: Zoom básico": "Zoom_basico",
        "Cuestionario:Reto: Zoom Basico": "Zoom_basico",
        "Cuestionario:Reto: Zoom Avanzado": "Zoom_Avanzado",
        "Cuestionario:Reto: Grupos Moodle": "Grupos_Moodle",
        "Cuestionario:Reto: Rúbrica": "Rubrica",
        "Cuestionario:Reto: Rubrica": "Rubrica",
        "Cuestionario:Reto: Padlet": "Padlet",
        "Cuestionario:Reto: Nearpod": "Nearpod",
        "Cuestionario:Reto: Tareas y foros": "Tareas_y_foros",
        "Cuestionario:Reto: Tareas y Foros": "Tareas_y_foros",
    }
    comp_out_cols = ["Zoom_basico", "Zoom_Avanzado", "Grupos_Moodle", "Rubrica", "Padlet", "Nearpod", "Tareas_y_foros"]

    # CRITICAL: initialize as NaN (NOT 0), so merge can fill real values like Nearpod=20
    for c in comp_out_cols:
        if c not in all_data.columns:
            all_data[c] = np.nan

    if comp_df is not None and not comp_df.empty:
        comp_df = _strip_cols(comp_df)
        ncol = find_col(comp_df, ["Nombre", "NOMBRES"])
        acol = find_col(comp_df, ["Apellido(s)", "Apellidos", "APELLIDOS", "APELLIDO(S)"])
        if ncol and acol:
            tmp = comp_df[[ncol, acol]].copy()
            tmp.columns = ["Nombre", "Apellido(s)"]
            tmp = ensure_person_keys(tmp)

            # Pull each component if exists
            for src, dst in comp_map.items():
                real = find_col(comp_df, [src])
                if real is not None:
                    tmp[dst] = score_to_numeric(comp_df[real])

            tmp = tmp.groupby(["Nombre_key", "Apellido_key"], as_index=False)[comp_out_cols].max()

            all_data = ensure_person_keys(all_data)
            merged = all_data.merge(tmp, on=["Nombre_key", "Apellido_key"], how="left", suffixes=("", "__comp"))

            # coalesce column-by-column (this fixes your Nearpod=20 -> 0 issue)
            for c in comp_out_cols:
                cc = f"{c}__comp"
                if cc in merged.columns:
                    merged[c] = merged[c].combine_first(merged[cc])
                    merged = merged.drop(columns=[cc])

            all_data = merged
        else:
            warnings.append("Comp. Tec: missing Nombre/Apellido(s) columns; skipped.")
    else:
        warnings.append("Comp. Tec: sheet empty/missing.")

    # -----------------------------
    # Teacher Contract filter + fill names
    # -----------------------------
    if contract_file is not None:
        contract = load_teacher_contract(contract_file, warnings)
        if not contract.empty:
            contract["DNI"] = contract["DNI"].apply(normalize_dni_value)
            contract = contract.dropna(subset=["DNI"])

            # Keep only teachers in contract
            all_data = all_data[all_data["DNI"].isin(contract["DNI"])]

            # Fill missing names from contract
            all_data = all_data.merge(contract, on="DNI", how="left", suffixes=("", "_contract"))
            all_data["Nombre"] = all_data["Nombre"].fillna(all_data["Nombre_contract"])
            all_data["Apellido(s)"] = all_data["Apellido(s)"].fillna(all_data["Apellido(s)_contract"])
            all_data = all_data.drop(columns=["Nombre_contract", "Apellido(s)_contract"], errors="ignore")

    # -----------------------------
    # Numeric columns (14 total) + Metrics
    # -----------------------------
    numeric_columns = [
        "induccion", "bus_biblioteca", "diseno_sesion",
        "Zoom_basico", "Zoom_Avanzado", "Grupos_Moodle", "Rubrica",
        "Padlet", "Nearpod", "Tareas_y_foros",
        "integracion", "rsu", "estress", "hab_comunicacion"
    ]

    for col in numeric_columns:
        if col not in all_data.columns:
            all_data[col] = np.nan
        all_data[col] = score_to_numeric(all_data[col]).fillna(0)

    all_data["Average"] = all_data[numeric_columns].mean(axis=1).round(2)
    all_data["Percentage"] = ((all_data[numeric_columns] > 0).sum(axis=1) / len(numeric_columns) * 100).round(2)
    all_data["Marks_Out_Of_20"] = (all_data["Percentage"] / 5).round(2)

    # Only compare 2024 vs 2025 and keep rows with any score
    filtered = all_data[all_data["Year"].isin([2024, 2025])].copy()
    filtered = filtered[filtered[numeric_columns].sum(axis=1) > 0]

    if filtered.empty:
        return pd.DataFrame(), warnings + ["No records with scores found for 2024 or 2025."]

    # Dedup: keep best row per DNI
    filtered["YearPref"] = filtered["Year"].apply(lambda y: 1 if y == 2025 else 0)
    sorted_df = filtered.sort_values(
        by=["Marks_Out_Of_20", "Average", "YearPref"],
        ascending=[False, False, False]
    )
    highest = sorted_df.drop_duplicates(subset=["DNI"], keep="first").copy()
    highest["Highest_Score_Year"] = highest["Year"]

    final_columns = [
        "Periodo", "Highest_Score_Year", "DNI", "Nombre", "Apellido(s)",
        *numeric_columns,
        "Average", "Marks_Out_Of_20", "Percentage"
    ]
    for c in final_columns:
        if c not in highest.columns:
            highest[c] = np.nan

    return highest[final_columns], warnings

# -----------------------------
# Streamlit App
# -----------------------------
def main():
    st.set_page_config(page_title="📊 UMA Scores (Highest of 2024 vs 2025)", page_icon="📊", layout="wide")

    st.title("📊 UMA Scores — Highest Marks (2024 vs 2025)")
    st.markdown(
        "Upload the **Master** Excel and the **Teacher Contract** Excel. "
        "Output has **one row per teacher** — the **highest Marks Out Of 20** between 2024 and 2025."
    )

    uploaded_file = st.file_uploader("Choose the Master Excel file", type=["xlsx", "xls"], key="master")
    uploaded_contract = st.file_uploader("Choose the Teacher Contract Excel file", type=["xlsx", "xls"], key="contract")

    if uploaded_file is not None and uploaded_contract is not None:
        try:
            with st.spinner("Processing..."):
                final_data, warnings = extract_data_from_excel(uploaded_file, contract_file=uploaded_contract)

            with st.expander("⚠️ Diagnostics / Warnings"):
                if warnings:
                    for w in warnings:
                        st.warning(w)
                else:
                    st.success("No warnings detected.")

            if final_data.empty:
                st.warning("No records with scores found for 2024 or 2025.")
                return

            st.success("Done! Showing only the highest marks per teacher (no duplicates).")

            st.subheader("Preview")
            st.dataframe(final_data.head(30), use_container_width=True)

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Teachers", len(final_data))
            with col2:
                st.metric("Avg Marks (Out of 20)", f"{final_data['Marks_Out_Of_20'].mean():.2f}")
            with col3:
                st.metric("Avg Percentage", f"{final_data['Percentage'].mean():.2f}%")

            st.subheader("Distribution by Highest Score Year")
            year_counts = final_data["Highest_Score_Year"].value_counts().sort_index()
            st.bar_chart(year_counts)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"Highest_Marks_2024_vs_2025_{timestamp}.xlsx"
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_data.to_excel(writer, index=False, sheet_name="Highest Marks (Unique)")
            output.seek(0)

            st.download_button(
                label="📥 Download (Unique, Highest of 2024/2025)",
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error while processing: {str(e)}")

    else:
        st.info("👆 Please upload both the Master Excel and the Teacher Contract Excel to get started.")

if __name__ == "__main__":
    main()
