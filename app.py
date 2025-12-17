# app.py
import pandas as pd
import numpy as np
import streamlit as st
import io
import re
import unicodedata
from datetime import datetime

# -----------------------------
# Helpers (text/columns)
# -----------------------------
def _strip_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def _norm_header(s: str) -> str:
    """Normalize header to compare robustly (uppercase, no accents, single spaces)."""
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s).upper()
    return s

def find_col(df: pd.DataFrame, candidates: list[str]):
    """Find a column in df that matches any of candidates (accent/case/space tolerant)."""
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

def normalize_dni_value(x):
    """Normalize DNI to a plain digit string (e.g., 12345678), handling floats like '12345678.0'."""
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    if re.match(r"^\d+\.0$", s):
        s = s[:-2]
    s = re.sub(r"\D", "", s)
    return s if s else np.nan

def extract_year(value):
    """Extract a 4-digit year (20xx) from Periodo fields like 2024, 2024-I, 2025-II."""
    if pd.isna(value):
        return pd.NA
    s = str(value)
    m = re.search(r"(20\d{2})", s)
    return int(m.group(1)) if m else pd.NA

def norm_person_key(s):
    """Normalize person name fields for safer merges."""
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = re.sub(r"\s+", " ", s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.upper()

def safe_read_sheet(xls: pd.ExcelFile, sheet_name: str, warnings: list[str]) -> pd.DataFrame:
    """Read a sheet if exists; else return empty df + warning."""
    if sheet_name not in xls.sheet_names:
        warnings.append(f"Missing sheet: '{sheet_name}' (filled with 0 where needed).")
        return pd.DataFrame()
    df = pd.read_excel(xls, sheet_name=sheet_name)
    return _strip_cols(df)

# -----------------------------
# Teacher contract loader
# -----------------------------
def load_teacher_contract(contract_file, warnings: list[str]) -> pd.DataFrame:
    """
    Load the teacher contract file and extract ['DNI','Nombre','Apellido(s)'].
    Tolerant with header variations.
    """
    df = pd.read_excel(contract_file, sheet_name=0)
    df = _strip_cols(df)

    dni_col = find_col(df, [
        "N° DE DOCUMENTO DE IDENTIDAD",
        "NRO DE DOCUMENTO DE IDENTIDAD",
        "NRO DE DOCUMENTO",
        "DOCUMENTO DE IDENTIDAD",
        "DNI"
    ])
    name_col = find_col(df, ["NOMBRES", "NOMBRE", "Nombres", "Nombre"])
    a_pat_col = find_col(df, ["APELLIDO PATERNO", "Apellido Paterno"])
    a_mat_col = find_col(df, ["APELLIDO MATERNO", "Apellido Materno"])

    if not all([dni_col, name_col, a_pat_col, a_mat_col]):
        warnings.append(
            "Teacher Contract file is missing required columns (DNI/Nombres/Apellidos). "
            "Cannot filter teachers reliably."
        )
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
# Merge helpers
# -----------------------------
def ensure_person_keys(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if "Nombre" in df.columns:
        df["Nombre_key"] = df["Nombre"].apply(norm_person_key)
    else:
        df["Nombre_key"] = ""
    if "Apellido(s)" in df.columns:
        df["Apellido_key"] = df["Apellido(s)"].apply(norm_person_key)
    else:
        df["Apellido_key"] = ""
    return df

def merge_component(all_data: pd.DataFrame,
                    comp_df: pd.DataFrame,
                    out_col: str,
                    value_candidates: list[str],
                    warnings: list[str],
                    prefer: str = "DNI"):
    """
    Merge a component sheet into all_data.
    prefer='DNI' tries DNI first; if missing, tries Nombre+Apellido(s); otherwise fills 0.
    """
    # default fill if sheet empty
    if comp_df is None or comp_df.empty:
        if out_col not in all_data.columns:
            all_data[out_col] = 0
        warnings.append(f"Sheet for '{out_col}' is empty or missing (filled with 0).")
        return all_data

    comp_df = _strip_cols(comp_df)

    value_col = find_col(comp_df, value_candidates)
    if value_col is None:
        # If value column missing, treat as empty
        if out_col not in all_data.columns:
            all_data[out_col] = 0
        warnings.append(f"'{out_col}': score column not found in sheet (filled with 0).")
        return all_data

    # Try merge by DNI
    dni_col = find_col(comp_df, ["DNI", "NRO DE DOCUMENTO", "DOCUMENTO", "DOCUMENTO DE IDENTIDAD"])
    can_by_dni = dni_col is not None

    # Try merge by names
    name_col = find_col(comp_df, ["Nombre", "Nombres"])
    ap_col = find_col(comp_df, ["Apellido(s)", "Apellidos", "Apellido", "APELLIDOS"])
    can_by_name = (name_col is not None and ap_col is not None)

    # Decide strategy
    strategies = []
    if prefer.upper() == "DNI":
        strategies = ["DNI", "NAME"]
    else:
        strategies = ["NAME", "DNI"]

    merged = all_data.copy()

    for strat in strategies:
        if strat == "DNI" and can_by_dni:
            tmp = comp_df[[dni_col, value_col]].copy()
            tmp.columns = ["DNI", out_col]
            tmp["DNI"] = tmp["DNI"].apply(normalize_dni_value)
            tmp = tmp.dropna(subset=["DNI"])
            merged = pd.merge(merged, tmp, on="DNI", how="left")
            return merged

        if strat == "NAME" and can_by_name:
            tmp = comp_df[[name_col, ap_col, value_col]].copy()
            tmp.columns = ["Nombre", "Apellido(s)", out_col]
            tmp = ensure_person_keys(tmp)

            merged = ensure_person_keys(merged)
            merged = pd.merge(
                merged,
                tmp[["Nombre_key", "Apellido_key", out_col]],
                on=["Nombre_key", "Apellido_key"],
                how="left"
            )
            return merged

    # If no merge key possible
    if out_col not in merged.columns:
        merged[out_col] = 0
    warnings.append(f"'{out_col}': sheet has no DNI and no Nombre/Apellido(s) keys (filled with 0).")
    return merged

# -----------------------------
# Core processing
# -----------------------------
def extract_data_from_excel(master_file, contract_file=None):
    warnings = []

    # Use ExcelFile so we can check available sheet names safely
    xls = pd.ExcelFile(master_file)

    # ---- Read sheets safely
    induction_df       = safe_read_sheet(xls, "Inducción", warnings)
    nota_induccion_df  = safe_read_sheet(xls, "nota Inducción", warnings)
    bus_biblioteca_df  = safe_read_sheet(xls, "Bus. biblioteca", warnings)
    diseno_sesion_df   = safe_read_sheet(xls, "Diseño de sesión", warnings)
    comp_tec_df        = safe_read_sheet(xls, "Comp. Tec", warnings)
    integracion_df     = safe_read_sheet(xls, "Integración", warnings)
    rsu_df             = safe_read_sheet(xls, "RSU", warnings)
    estress_df         = safe_read_sheet(xls, "estress", warnings)
    hab_com_df         = safe_read_sheet(xls, "Hab. comunicación", warnings)

    # ---- Build base from Inducción + nota Inducción (robust column lookup)
    base_parts = []

    # nota Inducción
    if not nota_induccion_df.empty:
        per_col = find_col(nota_induccion_df, ["PERIODO", "Periodo"])
        dni_col = find_col(nota_induccion_df, ["DNI", "NRO DE DOCUMENTO", "DOCUMENTO"])
        nom_col = find_col(nota_induccion_df, ["Nombre", "NOMBRES"])
        ape_col = find_col(nota_induccion_df, ["Apellido(s)", "Apellidos", "APELLIDOS"])
        mail_col = find_col(nota_induccion_df, ["Dirección de correo", "Direccion de correo", "Correo", "Email"])
        score_col = find_col(nota_induccion_df, ["Total del curso (Real)", "Total del curso", "Total"])

        if all([per_col, dni_col, nom_col, ape_col, score_col]):
            tmp = nota_induccion_df[[per_col, dni_col, nom_col, ape_col] + ([mail_col] if mail_col else []) + [score_col]].copy()
            cols = ["Periodo", "DNI", "Nombre", "Apellido(s)"] + (["Dirección de correo"] if mail_col else []) + ["induccion"]
            tmp.columns = cols
            base_parts.append(tmp)
        else:
            warnings.append("Sheet 'nota Inducción' is missing some required columns; skipped its rows.")

    # Inducción
    if not induction_df.empty:
        per_col = find_col(induction_df, ["Periodo", "PERIODO"])
        dni_col = find_col(induction_df, ["DNI", "NRO DE DOCUMENTO", "DOCUMENTO"])
        nom_col = find_col(induction_df, ["Nombre", "NOMBRES"])
        ape_col = find_col(induction_df, ["Apellido(s)", "Apellidos", "APELLIDOS"])
        mail_col = find_col(induction_df, ["Dirección de correo", "Direccion de correo", "Correo", "Email"])
        score_col = find_col(induction_df, ["Calificación", "Calificacion", "Nota", "Score"])

        if all([per_col, dni_col, nom_col, ape_col, score_col]):
            tmp = induction_df[[per_col, dni_col, nom_col, ape_col] + ([mail_col] if mail_col else []) + [score_col]].copy()
            cols = ["Periodo", "DNI", "Nombre", "Apellido(s)"] + (["Dirección de correo"] if mail_col else []) + ["induccion"]
            tmp.columns = cols
            base_parts.append(tmp)
        else:
            warnings.append("Sheet 'Inducción' is missing some required columns; skipped its rows.")

    if not base_parts:
        return pd.DataFrame(), ["No base data found in 'Inducción' / 'nota Inducción'. Check columns/sheets."]

    all_data = pd.concat(base_parts, ignore_index=True)

    # Normalize base fields
    all_data["DNI"] = all_data["DNI"].apply(normalize_dni_value)
    all_data["Year"] = all_data["Periodo"].apply(extract_year)

    # Ensure name keys for later name-based merges
    all_data = ensure_person_keys(all_data)

    # ---- Merge Bus. biblioteca -> bus_biblioteca (prefer DNI)
    all_data = merge_component(
        all_data,
        bus_biblioteca_df,
        out_col="bus_biblioteca",
        value_candidates=["Promedio", "promedio", "Nota", "Score"],
        warnings=warnings,
        prefer="DNI"
    )

    # ---- Diseño de sesión -> diseno_sesion (often by names)
    all_data = merge_component(
        all_data,
        diseno_sesion_df,
        out_col="diseno_sesion",
        value_candidates=["Promedio", "promedio", "Nota", "Score"],
        warnings=warnings,
        prefer="NAME"
    )

    # ---- Comp. Tec -> multiple columns (by names)
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

    # Make sure all expected comp columns exist even if sheet is missing/partial
    for outc in ["Zoom_basico","Zoom_Avanzado","Grupos_Moodle","Rubrica","Padlet","Nearpod","Tareas_y_foros"]:
        if outc not in all_data.columns:
            all_data[outc] = 0

    if not comp_tec_df.empty:
        comp_tec_df = _strip_cols(comp_tec_df)
        # build a temp df with name keys + each available component
        ncol = find_col(comp_tec_df, ["Nombre", "NOMBRES"])
        acol = find_col(comp_tec_df, ["Apellido(s)", "Apellidos", "APELLIDOS"])

        if ncol and acol:
            tmp = comp_tec_df[[ncol, acol]].copy()
            tmp.columns = ["Nombre", "Apellido(s)"]
            tmp = ensure_person_keys(tmp)

            for src, dst in comp_map.items():
                real = find_col(comp_tec_df, [src])
                if real is not None:
                    tmp[dst] = comp_tec_df[real]
                # else keep missing -> NaN; will become 0 later

            # merge
            all_data = pd.merge(
                all_data,
                tmp[["Nombre_key","Apellido_key"] + list(set(comp_map.values()))],
                on=["Nombre_key","Apellido_key"],
                how="left",
                suffixes=("", "_comp")
            )

            # if any duplicates of comp cols happened, coalesce
            for c in set(comp_map.values()):
                cc = c + "_comp"
                if cc in all_data.columns:
                    all_data[c] = all_data[c].fillna(all_data[cc])
                    all_data = all_data.drop(columns=[cc])
        else:
            warnings.append("Sheet 'Comp. Tec' has no Nombre/Apellido(s) columns; Comp. Tec filled with 0.")
    else:
        warnings.append("Sheet 'Comp. Tec' empty/missing; Comp. Tec filled with 0.")

    # ---- Integración -> integracion (by names)
    all_data = merge_component(
        all_data,
        integracion_df,
        out_col="integracion",
        value_candidates=[
            "Tarea:Producto final: Contenido académico, presentación y rúbrica con IA (Real)",
            "Tarea: Producto final: Contenido académico, presentación y rúbrica con IA (Real)",
            "Producto final",
            "Nota",
            "Score"
        ],
        warnings=warnings,
        prefer="NAME"
    )

    # ---- RSU -> rsu (prefer DNI, fallback to names, else 0)
    all_data = merge_component(
        all_data,
        rsu_df,
        out_col="rsu",
        value_candidates=["Tarea: Producto final", "Tarea:Producto final", "Producto final", "Nota", "Score"],
        warnings=warnings,
        prefer="DNI"
    )

    # ---- estress -> estress (prefer DNI, fallback to names, else 0)
    all_data = merge_component(
        all_data,
        estress_df,
        out_col="estress",
        value_candidates=["Tarea:Producto final", "Tarea: Producto final", "Producto final", "Nota", "Score"],
        warnings=warnings,
        prefer="DNI"
    )

    # ---- Hab. comunicación -> hab_comunicacion (prefer DNI)
    all_data = merge_component(
        all_data,
        hab_com_df,
        out_col="hab_comunicacion",
        value_candidates=["Tarea:Producto final", "Tarea: Producto final", "Producto final", "Nota", "Score"],
        warnings=warnings,
        prefer="DNI"
    )

    # ---- Teacher Contract (optional but recommended)
    if contract_file is not None:
        contract = load_teacher_contract(contract_file, warnings)
        if not contract.empty:
            contract["DNI"] = contract["DNI"].apply(normalize_dni_value)
            contract = contract.dropna(subset=["DNI"])

            # Keep only DNIs present in contract
            all_data = all_data[all_data["DNI"].isin(contract["DNI"])]

            # Fill names from contract when missing
            all_data = pd.merge(all_data, contract, on="DNI", how="left", suffixes=("", "_contract"))
            all_data["Nombre"] = all_data["Nombre"].fillna(all_data["Nombre_contract"])
            all_data["Apellido(s)"] = all_data["Apellido(s)"].fillna(all_data["Apellido(s)_contract"])
            all_data = all_data.drop(columns=["Nombre_contract", "Apellido(s)_contract"], errors="ignore")

    # -----------------------------
    # Numeric components (14 total)
    # -----------------------------
    numeric_columns = [
        "induccion", "bus_biblioteca", "diseno_sesion",
        "Zoom_basico", "Zoom_Avanzado", "Grupos_Moodle", "Rubrica",
        "Padlet", "Nearpod", "Tareas_y_foros",
        "integracion", "rsu", "estress", "hab_comunicacion"
    ]

    # Ensure all numeric cols exist
    for col in numeric_columns:
        if col not in all_data.columns:
            all_data[col] = 0

    # Convert
    for col in numeric_columns:
        all_data[col] = pd.to_numeric(all_data[col], errors="coerce").fillna(0)

    # ---- Compute metrics
    all_data["Average"] = all_data[numeric_columns].mean(axis=1).round(2)

    def calculate_percentage(row):
        scores = row[numeric_columns].values
        available = int(np.sum(np.array(scores) > 0))
        return round(available / len(numeric_columns) * 100, 2) if len(numeric_columns) else 0.0

    all_data["Percentage"] = all_data.apply(calculate_percentage, axis=1)
    all_data["Marks_Out_Of_20"] = (all_data["Percentage"] / 5).round(2)

    # ---- Only compare 2024 vs 2025 and keep rows with any score
    filtered = all_data[all_data["Year"].isin([2024, 2025])].copy()
    filtered = filtered[filtered[numeric_columns].sum(axis=1) > 0]

    if filtered.empty:
        return pd.DataFrame(), warnings + ["No records with scores found for 2024 or 2025."]

    # ---- DEDUP: one row per teacher (by DNI) with the HIGHEST Marks_Out_Of_20 across 2024 & 2025
    filtered["YearPref"] = filtered["Year"].apply(lambda y: 1 if y == 2025 else 0)
    sorted_df = filtered.sort_values(
        by=["Marks_Out_Of_20", "Average", "YearPref"],
        ascending=[False, False, False]
    )
    highest = sorted_df.drop_duplicates(subset=["DNI"], keep="first").copy()
    highest["Highest_Score_Year"] = highest["Year"]

    # ---- Final column order (create any missing columns safely)
    final_columns = [
        "Periodo", "Highest_Score_Year", "DNI", "Nombre", "Apellido(s)",
        "induccion", "bus_biblioteca", "diseno_sesion",
        "Zoom_basico", "Zoom_Avanzado", "Grupos_Moodle", "Rubrica",
        "Padlet", "Nearpod", "Tareas_y_foros",
        "integracion", "rsu", "estress", "hab_comunicacion",
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
        "The output contains **one row per teacher** — the **highest Marks Out Of 20** between **2024** and **2025**."
    )

    uploaded_file = st.file_uploader("Choose the Master Excel file", type=["xlsx", "xls"], key="master")
    uploaded_contract = st.file_uploader("Choose the Teacher Contract Excel file", type=["xlsx", "xls"], key="contract")

    if uploaded_file is not None and uploaded_contract is not None:
        try:
            with st.spinner("Processing your Excel files and comparing 2024 vs 2025..."):
                final_data, warnings = extract_data_from_excel(uploaded_file, contract_file=uploaded_contract)

            # Show warnings (important!)
            if warnings:
                with st.expander("⚠️ Warnings / Data issues detected (click to view)"):
                    for w in warnings:
                        st.warning(w)

            if final_data is None or len(final_data) == 0:
                st.warning("No usable records found for 2024 or 2025 (after cleaning/merging).")
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

            # Download
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
            st.info(
                "Tip: Open the Master file and confirm the sheet names/columns. "
                "This version should NOT crash on missing DNI; it will warn and fill with 0."
            )
    else:
        st.info("👆 Please upload both the Master Excel and the Teacher Contract Excel to get started.")

if __name__ == "__main__":
    main()
