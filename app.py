# app.py
import pandas as pd
import numpy as np
import streamlit as st
import io
import re
from datetime import datetime

# -----------------------------
# Helpers
# -----------------------------
EMAIL_REGEX = r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"

def normalize_dni_value(x):
    """Normalize DNI to a pure digit string and left-pad to 8 digits (e.g., '05403224')."""
    if pd.isna(x):
        return pd.NA
    s = str(x).strip()
    if re.match(r'^\d+\.0$', s):
        s = s[:-2]
    s = re.sub(r'\D', '', s)
    if not s:
        return pd.NA
    if s.isdigit() and len(s) <= 8:
        s = s.zfill(8)
    return s

def normalize_email(val, prefer_edu=True):
    """
    Extract and normalize a single email from free text.
    - Lowercases & trims.
    - If multiple, prefer one containing '.edu'.
    - Returns pd.NA if none found.
    """
    if pd.isna(val):
        return pd.NA
    text = str(val).strip().lower()
    emails = re.findall(EMAIL_REGEX, text)
    if not emails:
        tokens = re.split(r"[;,\s/]+", text)
        emails = [t for t in tokens if re.fullmatch(EMAIL_REGEX, t)]
    if not emails:
        return pd.NA
    if prefer_edu:
        edu_emails = [e for e in emails if ".edu" in e]
        if edu_emails:
            return edu_emails[0]
    return emails[0]

def extract_year(value):
    """Extract 4-digit year (20xx) from values like 2024, '2024-I', '2025-II', etc."""
    if pd.isna(value):
        return pd.NA
    s = str(value)
    m = re.search(r'(20\d{2})', s)
    return int(m.group(1)) if m else pd.NA

def load_email_roster(roster_file):
    """
    Read ANY Excel and return a roster based on .edu emails.
    - Scans ALL sheets & ALL columns for emails.
    - Tries to pick up DNI/Nombre/Apellidos when present.
    Returns columns: ['Email','DNI','Nombre','Apellido(s)'] deduped by Email.
    """
    xls = pd.ExcelFile(roster_file)
    collected = []

    # Column name hints
    name_first_candidates = {"NOMBRES", "Nombres", "Nombre"}
    apell_full = {"Apellido(s)", "Apellidos"}
    apell_pat = {"APELLIDO PATERNO", "Apellido Paterno"}
    apell_mat = {"APELLIDO MATERNO", "Apellido Materno"}

    dni_headers = [
        "N° DE DOCUMENTO DE IDENTIDAD", "N° DE DOCUMENTO DE IDENTIDAD ",
        "NRO DE DOCUMENTO DE IDENTIDAD", "NRO DE DOCUMENTO", "NRO DE DOCUMENTO",
        "DNI"
    ]

    for sheet in xls.sheet_names:
        df = pd.read_excel(roster_file, sheet_name=sheet)

        # Identify likely email columns (>=5% with '@')
        likely_email_cols = []
        for c in df.columns:
            if not isinstance(c, str):
                continue
            ratio = df[c].astype(str).str.contains("@", na=False).mean() if len(df) else 0
            if ratio >= 0.05:
                likely_email_cols.append(c)
        if not likely_email_cols:
            continue

        # Name/DNI columns (optional)
        name_col = next((c for c in df.columns if isinstance(c, str) and c in name_first_candidates), None)
        apell_s_col = next((c for c in df.columns if isinstance(c, str) and c in apell_full), None)
        apell_p_col = next((c for c in df.columns if isinstance(c, str) and c in apell_pat), None)
        apell_m_col = next((c for c in df.columns if isinstance(c, str) and c in apell_mat), None)
        dni_col = next((c for c in df.columns if isinstance(c, str) and c in dni_headers), None)

        # Build a tidy frame per email column
        for ec in likely_email_cols:
            tmp = pd.DataFrame()
            tmp["Email"] = df[ec].apply(lambda v: normalize_email(v, prefer_edu=True))
            if dni_col:
                tmp["DNI"] = df[dni_col].apply(normalize_dni_value)
            else:
                tmp["DNI"] = pd.NA

            # Names
            if name_col:
                tmp["Nombre"] = df[name_col].astype(str).str.strip()
            else:
                tmp["Nombre"] = ""

            if apell_s_col:
                tmp["Apellido(s)"] = df[apell_s_col].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
            else:
                # Combine paterno + materno if available
                if apell_p_col or apell_m_col:
                    paterno = df[apell_p_col].astype(str).str.strip() if apell_p_col else ""
                    materno = df[apell_m_col].astype(str).str.strip() if apell_m_col else ""
                    tmp["Apellido(s)"] = (paterno + " " + materno).str.replace(r"\s+", " ", regex=True).str.strip()
                else:
                    tmp["Apellido(s)"] = ""

            collected.append(tmp)

    if not collected:
        return pd.DataFrame(columns=["Email", "DNI", "Nombre", "Apellido(s)"])

    roster = pd.concat(collected, ignore_index=True)
    roster = roster.dropna(subset=["Email"])
    roster = roster[roster["Email"] != ""]
    roster = roster[roster["Email"].str.contains(r"\.edu", regex=True, na=False)]

    # Deduplicate by Email, prefer rows with any name/DNI filled
    roster["filled"] = (~roster["DNI"].isna()) | (roster["Nombre"].astype(str) != "") | (roster["Apellido(s)"].astype(str) != "")
    roster = roster.sort_values(by=["filled"], ascending=[False]).drop(columns=["filled"])
    roster = roster.drop_duplicates(subset=["Email"], keep="first")

    # Clean types
    roster["DNI"] = roster["DNI"].apply(normalize_dni_value)
    roster["Nombre"] = roster["Nombre"].astype(str).str.strip()
    roster["Apellido(s)"] = roster["Apellido(s)"].astype(str).str.strip()

    return roster[["Email", "DNI", "Nombre", "Apellido(s)"]]

# -----------------------------
# Core processing
# -----------------------------
def extract_data_from_excel(master_file, roster_file=None):
    # ---- Read master sheets (expected)
    induction_df       = pd.read_excel(master_file, sheet_name='Inducción')
    nota_induccion_df  = pd.read_excel(master_file, sheet_name='nota Inducción')
    bus_biblioteca_df  = pd.read_excel(master_file, sheet_name='Bus. biblioteca')
    diseno_sesion_df   = pd.read_excel(master_file, sheet_name='Diseño de sesión')
    comp_tec_df        = pd.read_excel(master_file, sheet_name='Comp. Tec')
    integracion_df     = pd.read_excel(master_file, sheet_name='Integración')
    rsu_df             = pd.read_excel(master_file, sheet_name='RSU')
    estress_df         = pd.read_excel(master_file, sheet_name='estress')
    hab_com_df         = pd.read_excel(master_file, sheet_name='Hab. comunicación')

    # ---- Base: Inducción + nota Inducción
    nota_induccion_clean = nota_induccion_df[
        ['PERIODO', 'DNI', 'Nombre', 'Apellido(s)', 'Dirección de correo', 'Total del curso (Real)']
    ].copy().rename(columns={'PERIODO': 'Periodo', 'Total del curso (Real)': 'induccion'})

    induction_clean = induction_df[
        ['Periodo', 'DNI', 'Nombre', 'Apellido(s)', 'Dirección de correo', 'Calificación']
    ].copy().rename(columns={'Calificación': 'induccion'})

    all_data = pd.concat([nota_induccion_clean, induction_clean], ignore_index=True)

    # Keys & year
    all_data['DNI']   = all_data['DNI'].apply(normalize_dni_value)
    all_data['Email'] = all_data['Dirección de correo'].apply(lambda v: normalize_email(v, prefer_edu=True))
    all_data['Year']  = all_data['Periodo'].apply(extract_year)

    # ---- Bus. biblioteca (by DNI)
    bus_biblioteca_df['DNI'] = bus_biblioteca_df['DNI'].apply(normalize_dni_value)
    bus_biblioteca_df = bus_biblioteca_df.rename(columns={'Promedio': 'bus_biblioteca'})
    all_data = pd.merge(all_data, bus_biblioteca_df[['DNI', 'bus_biblioteca']], on='DNI', how='left')

    # ---- Diseño de sesión (by names)
    diseno_sesion_df = diseno_sesion_df.rename(columns={'Promedio': 'diseno_sesion'})
    all_data = pd.merge(all_data, diseno_sesion_df[['Nombre', 'Apellido(s)', 'diseno_sesion']],
                        on=['Nombre', 'Apellido(s)'], how='left')

    # ---- Comp. Tec (by names)
    comp_tec_columns = {
        'Cuestionario:Reto: Zoom básico': 'Zoom_basico',
        'Cuestionario:Reto: Zoom Avanzado': 'Zoom_Avanzado',
        'Cuestionario:Reto: Grupos Moodle': 'Grupos_Moodle',
        'Cuestionario:Reto: Rúbrica': 'Rubrica',
        'Cuestionario:Reto: Padlet': 'Padlet',
        'Cuestionario:Reto: Nearpod': 'Nearpod',
        'Cuestionario:Reto: Tareas y foros': 'Tareas_y_foros'
    }
    comp_tec_df = comp_tec_df.rename(columns=comp_tec_columns)
    all_data = pd.merge(
        all_data,
        comp_tec_df[['Nombre', 'Apellido(s)', 'Zoom_basico', 'Zoom_Avanzado',
                     'Grupos_Moodle', 'Rubrica', 'Padlet', 'Nearpod', 'Tareas_y_foros']],
        on=['Nombre', 'Apellido(s)'], how='left'
    )

    # ---- Integración (by names)
    integracion_df = integracion_df.rename(columns={
        'Tarea:Producto final: Contenido académico, presentación y rúbrica con IA (Real)': 'integracion'
    })
    all_data = pd.merge(all_data, integracion_df[['Nombre', 'Apellido(s)', 'integracion']],
                        on=['Nombre', 'Apellido(s)'], how='left')

    # ---- RSU / estress / Hab. comunicación (by DNI)
    rsu_df['DNI'] = rsu_df['DNI'].apply(normalize_dni_value)
    estress_df['DNI'] = estress_df['DNI'].apply(normalize_dni_value)
    hab_com_df['DNI'] = hab_com_df['DNI'].apply(normalize_dni_value)

    rsu_df = rsu_df.rename(columns={'Tarea: Producto final': 'rsu'})
    estress_df = estress_df.rename(columns={'Tarea:Producto final': 'estress'})
    hab_com_df = hab_com_df.rename(columns={'Tarea:Producto final': 'hab_comunicacion'})

    all_data = pd.merge(all_data, rsu_df[['DNI', 'rsu']], on='DNI', how='left')
    all_data = pd.merge(all_data, estress_df[['DNI', 'estress']], on='DNI', how='left')
    all_data = pd.merge(all_data, hab_com_df[['DNI', 'hab_comunicacion']], on='DNI', how='left')

    # ---- Numeric components (14 total)
    numeric_columns = [
        'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
        'Padlet', 'Nearpod', 'Tareas_y_foros',
        'integracion', 'rsu', 'estress', 'hab_comunicacion'
    ]
    for col in numeric_columns:
        all_data[col] = pd.to_numeric(all_data.get(col, 0), errors='coerce').fillna(0)

    # ---- Metrics
    all_data['Average'] = all_data[numeric_columns].mean(axis=1).round(2)

    def calculate_percentage(row):
        scores = row[numeric_columns].values
        available = int(np.sum(np.array(scores) > 0))
        return round(available / len(numeric_columns) * 100, 2) if len(numeric_columns) else 0.0

    all_data['Percentage'] = all_data.apply(calculate_percentage, axis=1)
    all_data['Marks_Out_Of_20'] = (all_data['Percentage'] / 5).round(2)

    # ---- 2024/2025 only; keep zero-score rows
    base = all_data[all_data['Year'].isin([2024, 2025])].copy()

    # Best row per Email (if email present)
    base = base[~base['Email'].isna()].copy()
    if base.empty:
        best_rows = pd.DataFrame(columns=[
            'Periodo', 'Email', 'DNI', 'Nombre', 'Apellido(s)',
            *numeric_columns, 'Average', 'Marks_Out_Of_20', 'Percentage', 'Year'
        ])
    else:
        base['YearPref'] = base['Year'].apply(lambda y: 1 if y == 2025 else 0)
        sorted_df = base.sort_values(
            by=['Marks_Out_Of_20', 'Average', 'YearPref'],
            ascending=[False, False, False]
        )
        best_rows = sorted_df.drop_duplicates(subset=['Email'], keep='first').copy()
        best_rows['Highest_Score_Year'] = best_rows['Year']

    # ---- Roster (source of truth) from the provided file
    roster = load_email_roster(roster_file) if roster_file is not None else pd.DataFrame()
    if roster.empty:
        # No roster → fall back to whatever emails the master has (still deduped)
        # but user asked to take emails from "this file", so we keep the pipeline working.
        roster = best_rows[['Email', 'DNI']].copy()
        roster['Nombre'] = base.groupby('Email')['Nombre'].first().reindex(roster['Email']).fillna('')
        roster['Apellido(s)'] = base.groupby('Email')['Apellido(s)'].first().reindex(roster['Email']).fillna('')
        roster = roster.dropna(subset=['Email']).drop_duplicates(subset=['Email'])

    # Merge roster (left) with best_rows (right)
    keep_cols_best = [
        'Email', 'Periodo', 'Highest_Score_Year', 'DNI',
        *numeric_columns, 'Average', 'Marks_Out_Of_20', 'Percentage'
    ]
    best_rows = best_rows.reindex(columns=keep_cols_best)
    final = roster.merge(best_rows, on='Email', how='left', suffixes=('', '_best'))

    # Prefer names/DNI from roster; if missing, take from best_rows
    if 'DNI_best' in final.columns:
        final['DNI'] = final['DNI'].where(final['DNI'].notna() & (final['DNI'] != ''), final['DNI_best'])
        final.drop(columns=['DNI_best'], inplace=True)
    # Names
    if 'Nombre_best' in final.columns:
        final['Nombre'] = final['Nombre'].where(final['Nombre'].astype(str) != '', final['Nombre_best'])
        final.drop(columns=['Nombre_best'], inplace=True)
    if 'Apellido(s)_best' in final.columns:
        final['Apellido(s)'] = final['Apellido(s)'].where(final['Apellido(s)'].astype(str) != '', final['Apellido(s)_best'])
        final.drop(columns=['Apellido(s)_best'], inplace=True)

    # Placeholders for emails absent from master
    final['Periodo'] = final['Periodo'].fillna('2025')
    final['Highest_Score_Year'] = pd.to_numeric(final['Highest_Score_Year'], errors='coerce').fillna(2025).astype(int)
    for col in numeric_columns + ['Average', 'Marks_Out_Of_20', 'Percentage']:
        final[col] = pd.to_numeric(final[col], errors='coerce').fillna(0).round(2)

    # Final order
    final_columns = [
        'Periodo', 'Highest_Score_Year', 'Email', 'DNI', 'Nombre', 'Apellido(s)',
        'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
        'Padlet', 'Nearpod', 'Tareas_y_foros',
        'integracion', 'rsu', 'estress', 'hab_comunicacion',
        'Average', 'Marks_Out_Of_20', 'Percentage'
    ]
    for c in final_columns:
        if c not in final.columns:
            final[c] = '' if c in ['Periodo', 'Email', 'DNI', 'Nombre', 'Apellido(s)'] else 0
    return final[final_columns]

# -----------------------------
# Streamlit App
# -----------------------------
def main():
    st.set_page_config(page_title="📊 UMA Scores (Highest of 2024 vs 2025)", page_icon="📊", layout="wide")

    st.title("📊 UMA Scores — Highest Marks (2024 vs 2025)")
    st.markdown(
        "Upload the **Master** Excel and the **Professor Email Roster** Excel (the file you mentioned). "
        "We scan the roster for institutional **.edu** emails and produce **one row per email (no duplicates)**. "
        "For each email, we keep the **best** (highest _Marks Out Of 20_) row from **2024/2025** in the master "
        "(tie-breakers: higher Average, then prefer 2025). "
        "If an email has **no master data**, it still appears with `Periodo=2025` and all scores `0`."
    )

    uploaded_master = st.file_uploader("Choose the Master Excel file", type=["xlsx", "xls"], key="master")
    uploaded_roster = st.file_uploader("Choose the Professor Email Roster Excel (.xlsx/.xls)", type=["xlsx", "xls"], key="roster")

    if uploaded_master is not None and uploaded_roster is not None:
        try:
            with st.spinner("Processing files and comparing 2024 vs 2025..."):
                final_data = extract_data_from_excel(uploaded_master, roster_file=uploaded_roster)

            if len(final_data) == 0:
                st.warning("No professors found. Check that the roster contains .edu emails.")
                return

            st.success("Done! One row per roster .edu email.")
            st.subheader("Preview")
            st.dataframe(final_data.head(30))

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Professors", len(final_data))
            with col2:
                st.metric("Avg Marks (Out of 20)", f"{final_data['Marks_Out_Of_20'].mean():.2f}")
            with col3:
                st.metric("Avg Percentage", f"{final_data['Percentage'].mean():.2f}%")

            st.subheader("Distribution by Highest Score Year")
            counts = final_data['Highest_Score_Year'].value_counts().sort_index()
            st.bar_chart(counts)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"Highest_Marks_2024_vs_2025_byEmail_{timestamp}.xlsx"
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_data.to_excel(writer, index=False, sheet_name='Highest Marks (Unique by Email)')
            output.seek(0)

            st.download_button(
                label="📥 Download (Unique by .edu Email)",
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {str(e)}")
            st.info(
                "Master must contain the sheets: 'Inducción', 'nota Inducción', 'Bus. biblioteca', "
                "'Diseño de sesión', 'Comp. Tec', 'Integración', 'RSU', 'estress', 'Hab. comunicación'. "
                "Roster can be any Excel with .edu emails; names/DNI are optional."
            )
    else:
        st.info("👆 Please upload both the **Master** Excel and the **Professor Email Roster** Excel to get started.")

        st.subheader("Roster tips")
        st.markdown("""
        We auto-detect .edu emails from any sheet/column. If available, add:
        - `Nombre` (or `Nombres` / `Nombre`)
        - `Apellido(s)` (or `Apellido Paterno` + `Apellido Materno`)
        - `DNI` (any of the typical headers)
        """)

if __name__ == "__main__":
    main()
