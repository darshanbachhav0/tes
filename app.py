# app.py
import io
import re
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

# -----------------------------
# Helpers
# -----------------------------
def normalize_dni_value(x):
    if pd.isna(x):
        return pd.NA
    s = str(x).strip()
    if re.match(r'^\d+\.0$', s):
        s = s[:-2]
    s = re.sub(r'\D', '', s)
    return s if s else pd.NA

def normalize_email(x):
    if pd.isna(x):
        return pd.NA
    s = str(x).strip().lower()
    s = s.replace("mailto:", "").strip(" ,;.")
    s = re.sub(r"\s+", "", s)
    return s if ("@" in s and "." in s) else pd.NA

def extract_year(value):
    if pd.isna(value):
        return pd.NA
    s = str(value)
    m = re.search(r'(20\d{2})', s)
    return int(m.group(1)) if m else pd.NA

def strip_accents(s: str) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    return "".join([c for c in s if not unicodedata.combining(c)]).lower()

def find_col(df: pd.DataFrame, keywords, prefer_exact=None):
    cols = [c for c in df.columns if isinstance(c, str)]
    norm = {strip_accents(c): c for c in cols}
    if prefer_exact:
        for pe in prefer_exact:
            key = strip_accents(pe)
            if key in norm:
                return norm[key]
    for c in cols:
        cu = strip_accents(c)
        if any(k in cu for k in keywords):
            return c
    return None

def split_fullname(fullname):
    if pd.isna(fullname):
        return pd.NA, pd.NA
    s = str(fullname).strip()
    if not s:
        return pd.NA, pd.NA
    parts = s.split()
    if len(parts) == 1:
        return parts[0], pd.NA
    return parts[0], " ".join(parts[1:])

def load_teacher_email_roster(roster_file):
    wb = pd.read_excel(roster_file, sheet_name=0, dtype=str, engine="openpyxl")
    if isinstance(wb, dict):
        df = list(wb.values())[0]
    else:
        df = wb

    email_col = find_col(
        df, ["correo", "email", "e-mail", "mail"],
        prefer_exact=["Dirección de correo", "Correo", "Email", "E-mail"]
    )
    if email_col is None:
        for c in df.columns:
            frac = df[c].astype(str).str.contains("@", na=False).mean()
            if frac > 0.3:
                email_col = c
                break
    if email_col is None:
        return pd.DataFrame(columns=["Email", "Nombre", "Apellido(s)"])

    nombres_col = find_col(df, ["nombres", "nombre"], prefer_exact=["NOMBRES", "Nombre"])
    a_pat_col   = find_col(df, ["apellido paterno", "a. paterno"])
    a_mat_col   = find_col(df, ["apellido materno", "a. materno"])
    full_col    = find_col(df, ["docente", "profesor", "nombre completo",
                                "nombres y apellidos", "apellidos y nombres"])

    out = pd.DataFrame()
    out["Email"] = df[email_col].apply(normalize_email)

    if nombres_col and a_pat_col and a_mat_col:
        out["Nombre"] = df[nombres_col].astype(str).str.strip()
        out["Apellido(s)"] = (
            df[a_pat_col].astype(str).str.strip() + " " +
            df[a_mat_col].astype(str).str.strip()
        ).str.replace(r"\s+", " ", regex=True).str.strip()
    elif full_col:
        tmp = df[full_col].apply(split_fullname)
        out["Nombre"] = tmp.apply(lambda t: t[0])
        out["Apellido(s)"] = tmp.apply(lambda t: t[1])
    else:
        out["Nombre"] = pd.NA
        out["Apellido(s)"] = pd.NA

    out = out.dropna(subset=["Email"]).drop_duplicates(subset=["Email"])
    return out[["Email", "Nombre", "Apellido(s)"]]

def read_workbook(file):
    raw = pd.read_excel(file, sheet_name=None, dtype=str, engine="openpyxl")
    book = {}
    for name, df in raw.items():
        key = strip_accents(name).replace(" ", "")
        book[key] = df.copy()
    return book

def get_sheet(book, possible_names):
    keys = list(book.keys())
    for candidate in possible_names:
        cand = strip_accents(candidate).replace(" ", "")
        for k in keys:
            if cand == k or cand in k or k in cand:
                return book[k].copy()
    return None

def coalesce_keys(df: pd.DataFrame) -> pd.DataFrame:
    """After merges, collapse *_x/*_y to clean keys."""
    for key in ["Email", "DNI", "Nombre", "Apellido(s)"]:
        kx, ky = f"{key}_x", f"{key}_y"
        if kx in df.columns or ky in df.columns:
            if key not in df.columns:
                df[key] = pd.NA
            if kx in df.columns:
                df[key] = df[key].where(df[key].notna(), df[kx])
            if ky in df.columns:
                df[key] = df[key].where(df[key].notna(), df[ky])
            df = df.drop(columns=[c for c in [kx, ky] if c in df.columns])
    return df

# -----------------------------
# Core processing
# -----------------------------
def extract_data_from_excel(master_file, roster_file=None):
    # Read master workbook once
    book = read_workbook(master_file)
    df_induccion      = get_sheet(book, ["inducción", "induccion"])
    df_nota_ind       = get_sheet(book, ["nota inducción", "nota induccion"])
    df_biblio         = get_sheet(book, ["bus. biblioteca", "biblioteca", "bus biblioteca"])
    df_diseno         = get_sheet(book, ["diseño de sesión", "diseno de sesion", "diseño sesion"])
    df_comptec        = get_sheet(book, ["comp. tec", "competencias tecn", "comp tec"])
    df_integracion    = get_sheet(book, ["integración", "integracion"])
    df_rsu            = get_sheet(book, ["rsu"])
    df_estress        = get_sheet(book, ["estress", "estres", "estrés"])
    df_habcom         = get_sheet(book, ["hab. comunicación", "hab comunicacion", "habilidades comunicacion"])

    if df_induccion is None and df_nota_ind is None:
        raise ValueError("Neither 'Inducción' nor 'nota Inducción' sheet found in Master file.")

    frames = []
    if df_nota_ind is not None:
        cols_keep = ["PERIODO", "DNI", "Nombre", "Apellido(s)", "Dirección de correo", "Total del curso (Real)"]
        exist = [c for c in cols_keep if c in df_nota_ind.columns]
        tmp = df_nota_ind[exist].copy()
        if "PERIODO" in tmp.columns:
            tmp = tmp.rename(columns={"PERIODO": "Periodo"})
        if "Total del curso (Real)" in tmp.columns:
            tmp = tmp.rename(columns={"Total del curso (Real)": "induccion"})
        frames.append(tmp)

    if df_induccion is not None:
        cols_keep = ["Periodo", "DNI", "Nombre", "Apellido(s)", "Dirección de correo", "Calificación"]
        exist = [c for c in cols_keep if c in df_induccion.columns]
        tmp = df_induccion[exist].copy()
        if "Calificación" in tmp.columns:
            tmp = tmp.rename(columns={"Calificación": "induccion"})
        frames.append(tmp)

    all_data = pd.concat(frames, ignore_index=True)

    # Normalize base ids
    all_data["DNI"] = all_data["DNI"].apply(normalize_dni_value) if "DNI" in all_data.columns else pd.NA
    if "Dirección de correo" in all_data.columns:
        all_data["Email"] = all_data["Dirección de correo"].apply(normalize_email)
    else:
        all_data["Email"] = pd.NA
    all_data["Year"] = all_data["Periodo"].apply(extract_year) if "Periodo" in all_data.columns else pd.NA

    # Generic merge helper that avoids Email_x/Email_y
    def smart_merge(left: pd.DataFrame, right: pd.DataFrame, left_on, right_on, how="left"):
        # If the keys are the same string, use 'on=' to prevent _x/_y
        if isinstance(left_on, str) and isinstance(right_on, str) and left_on == right_on:
            out = pd.merge(left, right, on=left_on, how=how)
        else:
            out = pd.merge(left, right, left_on=left_on, right_on=right_on, how=how)
        return coalesce_keys(out)

    # Helper to add a sheet with multiple possible join keys
    def add_by_key(src_df, mapping, left_on_opts, how="left"):
        nonlocal all_data
        if src_df is None:
            return
        src = src_df.copy()
        for old, new in list(mapping.items()):
            if old in src.columns:
                src = src.rename(columns={old: new})

        for left_on, right_on in left_on_opts:
            # presence check
            if isinstance(left_on, list):
                left_ok = all(col in all_data.columns for col in left_on)
            else:
                left_ok = left_on in all_data.columns
            if isinstance(right_on, list):
                right_ok = all(col in src.columns for col in right_on)
            else:
                right_ok = right_on in src.columns
            if not (left_ok and right_ok):
                continue

            # normalize right keys
            if (isinstance(right_on, str) and ("correo" in strip_accents(right_on) or right_on == "Email")):
                src["Email"] = src[right_on].apply(normalize_email)
                right_on = "Email"
            if right_on == "DNI":
                src["DNI"] = src["DNI"].apply(normalize_dni_value)

            all_data = smart_merge(all_data, src, left_on, right_on, how=how)
            return  # done once

    # Bus. biblioteca
    if df_biblio is not None:
        prom = find_col(df_biblio, ["promedio"]) or "Promedio"
        df_biblio = df_biblio.rename(columns={prom: "bus_biblioteca"}) if prom in df_biblio.columns else df_biblio
        add_by_key(
            df_biblio, mapping={},
            left_on_opts=[
                ("Email", find_col(df_biblio, ["correo", "email"]) or "Email"),
                ("DNI", "DNI"),
                (["Nombre", "Apellido(s)"], ["Nombre", "Apellido(s)"])
            ]
        )

    # Diseño de sesión (usually by names)
    if df_diseno is not None:
        prom_col = find_col(df_diseno, ["promedio"]) or "Promedio"
        df_tmp = df_diseno.rename(columns={prom_col: "diseno_sesion"}) if prom_col in df_diseno.columns else df_diseno.copy()
        add_by_key(
            df_tmp, mapping={},
            left_on_opts=[
                (["Nombre", "Apellido(s)"], ["Nombre", "Apellido(s)"]),
                ("Email", find_col(df_tmp, ["correo", "email"]) or "Email"),
                ("DNI", "DNI")
            ]
        )

    # Comp. Tec
    if df_comptec is not None:
        ren = {
            "Cuestionario:Reto: Zoom básico": "Zoom_basico",
            "Cuestionario:Reto: Zoom Avanzado": "Zoom_Avanzado",
            "Cuestionario:Reto: Grupos Moodle": "Grupos_Moodle",
            "Cuestionario:Reto: Rúbrica": "Rubrica",
            "Cuestionario:Reto: Padlet": "Padlet",
            "Cuestionario:Reto: Nearpod": "Nearpod",
            "Cuestionario:Reto: Tareas y foros": "Tareas_y_foros",
        }
        df_tmp = df_comptec.rename(columns={k: v for k, v in ren.items() if k in df_comptec.columns})
        add_by_key(
            df_tmp, mapping={},
            left_on_opts=[
                (["Nombre", "Apellido(s)"], ["Nombre", "Apellido(s)"]),
                ("Email", find_col(df_tmp, ["correo", "email"]) or "Email"),
                ("DNI", "DNI")
            ]
        )

    # Integración
    if df_integracion is not None:
        integ_col = None
        for cand in [
            "Tarea:Producto final: Contenido académico, presentación y rúbrica con IA (Real)",
            "Tarea:Producto final", "Producto final", "Integración", "Integracion"
        ]:
            if cand in df_integracion.columns:
                integ_col = cand
                break
        df_tmp = df_integracion.copy()
        if integ_col:
            df_tmp = df_tmp.rename(columns={integ_col: "integracion"})
        add_by_key(
            df_tmp, mapping={},
            left_on_opts=[
                (["Nombre", "Apellido(s)"], ["Nombre", "Apellido(s)"]),
                ("Email", find_col(df_tmp, ["correo", "email"]) or "Email"),
                ("DNI", "DNI")
            ]
        )

    # RSU / Estress / Hab. comunicación
    for sub_df, out_col, default_old in [
        (df_rsu, "rsu", "Tarea: Producto final"),
        (df_estress, "estress", "Tarea:Producto final"),
        (df_habcom, "hab_comunicacion", "Tarea:Producto final"),
    ]:
        if sub_df is None:
            continue
        src = sub_df.copy()
        score_col = None
        for c in sub_df.columns:
            if strip_accents(c).startswith("tarea") or "promedio" in strip_accents(c):
                score_col = c
                break
        if score_col is None and default_old in sub_df.columns:
            score_col = default_old
        if score_col:
            src = src.rename(columns={score_col: out_col})
        else:
            src[out_col] = pd.NA

        add_by_key(
            src, mapping={},
            left_on_opts=[
                ("Email", find_col(src, ["correo", "email"]) or "Email"),
                ("DNI", "DNI"),
                (["Nombre", "Apellido(s)"], ["Nombre", "Apellido(s)"])
            ]
        )

    # Teacher Email Roster filter/fill
    if roster_file is not None:
        roster = load_teacher_email_roster(roster_file)
        roster["Email"] = roster["Email"].apply(normalize_email)
        roster = roster.dropna(subset=["Email"]).drop_duplicates(subset=["Email"])
        if len(roster) == 0:
            # Nothing to filter; keep going but warn in UI later
            pass
        else:
            valid_emails = set(roster["Email"].tolist())
            all_data = all_data[all_data["Email"].isin(valid_emails)]
            all_data = pd.merge(all_data, roster, on="Email", how="left", suffixes=("", "_roster"))
            # Fill names from roster when missing
            def blank(x):
                return (pd.isna(x)) or (str(x).strip() == "")
            if "Nombre" in all_data.columns and "Nombre_roster" in all_data.columns:
                all_data["Nombre"] = np.where(all_data["Nombre"].apply(blank), all_data["Nombre_roster"], all_data["Nombre"])
            if "Apellido(s)" in all_data.columns and "Apellido(s)_roster" in all_data.columns:
                all_data["Apellido(s)"] = np.where(all_data["Apellido(s)"].apply(blank), all_data["Apellido(s)_roster"], all_data["Apellido(s)"])
            all_data = all_data.drop(columns=[c for c in ["Nombre_roster", "Apellido(s)_roster"] if c in all_data.columns])

    # Numeric components
    numeric_columns = [
        "induccion", "bus_biblioteca", "diseno_sesion",
        "Zoom_basico", "Zoom_Avanzado", "Grupos_Moodle", "Rubrica",
        "Padlet", "Nearpod", "Tareas_y_foros",
        "integracion", "rsu", "estress", "hab_comunicacion"
    ]
    for col in numeric_columns:
        if col not in all_data.columns:
            all_data[col] = 0
        all_data[col] = pd.to_numeric(all_data[col], errors="coerce").fillna(0)

    # Metrics
    all_data["Average"] = all_data[numeric_columns].mean(axis=1).round(2)
    def calculate_percentage(row):
        scores = row[numeric_columns].values
        available = int(np.sum(np.array(scores) > 0))
        return round(available / len(numeric_columns) * 100, 2) if len(numeric_columns) else 0.0
    all_data["Percentage"] = all_data.apply(calculate_percentage, axis=1)
    all_data["Marks_Out_Of_20"] = (all_data["Percentage"] / 5).round(2)

    if "Year" not in all_data.columns:
        all_data["Year"] = pd.NA
    filtered = all_data[all_data["Year"].isin([2024, 2025])].copy()
    filtered = filtered[filtered[numeric_columns].sum(axis=1) > 0]
    if filtered.empty:
        return pd.DataFrame()

    # Dedup by Email with tie-breaks
    filtered["YearPref"] = filtered["Year"].apply(lambda y: 1 if y == 2025 else 0)
    sorted_df = filtered.sort_values(
        by=["Marks_Out_Of_20", "Average", "YearPref"],
        ascending=[False, False, False]
    )
    highest = sorted_df.drop_duplicates(subset=["Email"], keep="first").copy()
    highest["Highest_Score_Year"] = highest["Year"]

    final_columns = [
        "Periodo", "Highest_Score_Year", "Email", "DNI", "Nombre", "Apellido(s)",
        "induccion", "bus_biblioteca", "diseno_sesion",
        "Zoom_basico", "Zoom_Avanzado", "Grupos_Moodle", "Rubrica",
        "Padlet", "Nearpod", "Tareas_y_foros",
        "integracion", "rsu", "estress", "hab_comunicacion",
        "Average", "Marks_Out_Of_20", "Percentage"
    ]
    for c in final_columns:
        if c not in highest.columns:
            highest[c] = pd.NA
    return highest[final_columns]

# -----------------------------
# Streamlit App
# -----------------------------
def main():
    st.set_page_config(page_title="📊 UMA Scores (Highest of 2024 vs 2025)", page_icon="📊", layout="wide")

    st.title("📊 UMA Scores — Highest Marks (2024 vs 2025)")
    st.markdown(
        "Upload **Master** Excel and **Teacher Email Roster** (e.g., *C9_25-II_121025.xlsx*). "
        "Teachers are identified by **Email**; we keep the row with the **highest Marks Out Of 20** across 2024 & 2025."
    )

    up_master = st.file_uploader("Choose the Master Excel file", type=["xlsx", "xls"], key="master")
    up_roster = st.file_uploader("Choose the Teacher Email Roster Excel file", type=["xlsx", "xls"], key="roster")

    if up_master is not None and up_roster is not None:
        try:
            with st.spinner("Processing your Excel files and comparing 2024 vs 2025 (by Email)…"):
                final_data = extract_data_from_excel(up_master, roster_file=up_roster)

            if final_data.empty:
                st.warning("No records with scores found for 2024 or 2025 (after filtering by the email roster).")
                return

            st.success(f"Done! Showing highest marks per teacher (unique by Email). Total: {len(final_data)}")

            st.subheader("Preview")
            st.dataframe(final_data.head(30))

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
            output_filename = f"Highest_Marks_2024_vs_2025_by_Email_{timestamp}.xlsx"
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                final_data.to_excel(writer, index=False, sheet_name="Highest Marks (Unique by Email)")
            output.seek(0)
            st.download_button(
                label="📥 Download (Unique by Email, Highest of 2024/2025)",
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error while processing: {e}")
            st.info("If you still see this error, check Render logs for the stack trace (likely a header mismatch).")
    else:
        st.info("👆 Please upload both the Master Excel and the Teacher Email Roster Excel to get started.")
        st.subheader("Expected Columns (quick reference)")
        st.markdown("""
**Master Excel** (fuzzy sheet detection works):
- **Inducción / nota Inducción**: `Periodo/PERIODO`, `DNI`, `Nombre`, `Apellido(s)`, `Dirección de correo`, `Calificación` or `Total del curso (Real)`
- **Bus. biblioteca**: `DNI` and/or `Email`, `Promedio`
- **Diseño de sesión**: `Nombre`, `Apellido(s)`, `Promedio`
- **Comp. Tec**: Zoom básico/Avanzado, Grupos Moodle, Rúbrica, Padlet, Nearpod, Tareas y foros
- **Integración**: final score column (title varies)
- **RSU / estress / Hab. comunicación**: score column (title varies)

**Teacher Email Roster (e.g., C9_25-II_121025.xlsx)**:
- One email column (e.g., `Dirección de correo`, `Correo`, `Email`).
- Optional: names (`Nombres`, `Apellido Paterno`, `Apellido Materno`) or a single full-name column.
        """)

if __name__ == "__main__":
    main()
