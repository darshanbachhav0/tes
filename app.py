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
def normalize_dni_value(x):
    """Normalize DNI to a plain digit string (e.g., 12345678), handling floats like '12345678.0'."""
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    if re.match(r'^\d+\.0$', s):
        s = s[:-2]
    s = re.sub(r'\D', '', s)
    return s if s else np.nan

def extract_year(value):
    """
    Extract a 4-digit year (20xx) from 'Periodo' fields that might look like:
    2024, '2024-I', '2025-II', '2025', etc.
    """
    if pd.isna(value):
        return pd.NA
    s = str(value)
    m = re.search(r'(20\d{2})', s)
    return int(m.group(1)) if m else pd.NA

def infer_year_from_text(text):
    """
    Try to infer a 4-digit year from an arbitrary string (e.g., a filename like 'C9_25-II_121025.xlsx').
    Priority:
      1) direct '20xx' pattern,
      2) any standalone 2-digit '2x' (e.g., 25 -> 2025).
    Returns int or pd.NA
    """
    if not text:
        return pd.NA
    s = str(text)
    m4 = re.search(r'20\d{2}', s)
    if m4:
        return int(m4.group(0))
    # fall back: look for two-digit year like '25'
    m2 = re.search(r'(?<!\d)(2[0-9])(?!\d)', s)
    if m2:
        return 2000 + int(m2.group(1))
    return pd.NA

# -----------------------------
# C9 roster loader (replaces teacher contract)
# -----------------------------
def load_c9_roster(c9_file):
    """
    Load the C9 roster and extract ['DNI','Nombre','Apellido(s)','Email'].
    Tolerant with header variations commonly seen in C9 exports.
    """
    df = pd.read_excel(c9_file, sheet_name=0)

    # Canonicalize column name lookup
    def pick(colnames, *candidates):
        for cand in candidates:
            if cand in colnames:
                return colnames[cand]
        # fuzzy match fallbacks
        for c in df.columns:
            if not isinstance(c, str):
                continue
            up = c.upper()
            for cand in candidates:
                if isinstance(cand, str) and cand.upper() in up:
                    return c
        return None

    cols = { (c.strip() if isinstance(c, str) else c): c for c in df.columns }

    # DNI
    dni_col = pick(cols,
                   'N° DE DOCUMENTO DE IDENTIDAD', 'NRO DE DOCUMENTO DE IDENTIDAD',
                   'NRO DE DOCUMENTO', 'DNI', 'DOCUMENTO', 'DOC')
    # Names
    nombres_col   = pick(cols, 'NOMBRES', 'Nombres', 'NOMBRE', 'NOMBRE(S)')
    ap_pat_col    = pick(cols, 'APELLIDO PATERNO', 'Apellido Paterno', 'APE. PATERNO', 'APELLIDO  PATERNO')
    ap_mat_col    = pick(cols, 'APELLIDO MATERNO', 'Apellido Materno', 'APE. MATERNO', 'APELLIDO  MATERNO')
    # Email (be generous)
    email_col     = pick(cols, 'email', 'Email', 'EMAIL', 'Correo', 'CORREO', 'Dirección de correo',
                         'CORREO ELECTRÓNICO', 'Correo electrónico', 'Correo Institucional', 'CORREO INSTITUCIONAL',
                         'Correo Inst.', 'E-MAIL')

    if dni_col is None:
        # Return empty but compatible frame if DNI missing
        return pd.DataFrame(columns=['DNI', 'Nombre', 'Apellido(s)', 'Email'])

    out = pd.DataFrame()
    out['DNI'] = df[dni_col].apply(normalize_dni_value)

    # Names (optional; keep blanks if missing)
    if nombres_col is not None:
        out['Nombre'] = df[nombres_col].astype(str).str.strip()
    else:
        out['Nombre'] = ""

    if ap_pat_col is not None and ap_mat_col is not None:
        out['Apellido(s)'] = (
            df[ap_pat_col].astype(str).str.strip() + ' ' +
            df[ap_mat_col].astype(str).str.strip()
        ).str.replace(r'\s+', ' ', regex=True).str.strip()
    else:
        out['Apellido(s)'] = ""

    # Email
    if email_col is not None:
        out['Email'] = df[email_col].astype(str).str.strip()
    else:
        out['Email'] = ""

    out = out.dropna(subset=['DNI'])
    out = out[out['DNI'] != '']
    out = out.drop_duplicates(subset=['DNI'])

    # Final columns
    return out[['DNI', 'Nombre', 'Apellido(s)', 'Email']]

# -----------------------------
# Core processing
# -----------------------------
def extract_data_from_excel(file_path, c9_file=None):
    # ---- Read all source sheets (Master)
    induction_df       = pd.read_excel(file_path, sheet_name='Inducción')
    nota_induccion_df  = pd.read_excel(file_path, sheet_name='nota Inducción')
    bus_biblioteca_df  = pd.read_excel(file_path, sheet_name='Bus. biblioteca')
    diseno_sesion_df   = pd.read_excel(file_path, sheet_name='Diseño de sesión')
    comp_tec_df        = pd.read_excel(file_path, sheet_name='Comp. Tec')
    integracion_df     = pd.read_excel(file_path, sheet_name='Integración')
    rsu_df             = pd.read_excel(file_path, sheet_name='RSU')
    estress_df         = pd.read_excel(file_path, sheet_name='estress')
    hab_com_df         = pd.read_excel(file_path, sheet_name='Hab. comunicación')

    # ---- Prepare base: Inducción + nota Inducción (unified columns)
    nota_induccion_clean = nota_induccion_df[
        ['PERIODO', 'DNI', 'Nombre', 'Apellido(s)', 'Dirección de correo', 'Total del curso (Real)']
    ].copy()
    nota_induccion_clean = nota_induccion_clean.rename(columns={
        'PERIODO': 'Periodo',
        'Total del curso (Real)': 'induccion'
    })

    induction_clean = induction_df[
        ['Periodo', 'DNI', 'Nombre', 'Apellido(s)', 'Dirección de correo', 'Calificación']
    ].copy()
    induction_clean = induction_clean.rename(columns={'Calificación': 'induccion'})

    all_data = pd.concat([nota_induccion_clean, induction_clean], ignore_index=True)

    # Normalize DNI and derive Year
    all_data['DNI'] = all_data['DNI'].apply(normalize_dni_value)
    all_data['Year'] = all_data['Periodo'].apply(extract_year)

    # ---- Bus. biblioteca (by DNI)
    bus_biblioteca_df['DNI'] = bus_biblioteca_df['DNI'].apply(normalize_dni_value)
    bus_biblioteca_df = bus_biblioteca_df.rename(columns={'Promedio': 'bus_biblioteca'})
    all_data = pd.merge(
        all_data, bus_biblioteca_df[['DNI', 'bus_biblioteca']],
        on='DNI', how='left'
    )

    # ---- Diseño de sesión (join by names)
    diseno_sesion_df = diseno_sesion_df.rename(columns={'Promedio': 'diseno_sesion'})
    all_data = pd.merge(
        all_data, diseno_sesion_df[['Nombre', 'Apellido(s)', 'diseno_sesion']],
        on=['Nombre', 'Apellido(s)'], how='left'
    )

    # ---- Comp. Tec (join by names)
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

    # ---- Integración (join by names)
    integracion_df = integracion_df.rename(columns={
        'Tarea:Producto final: Contenido académico, presentación y rúbrica con IA (Real)': 'integracion'
    })
    all_data = pd.merge(
        all_data, integracion_df[['Nombre', 'Apellido(s)', 'integracion']],
        on=['Nombre', 'Apellido(s)'], how='left'
    )

    # ---- RSU (by DNI)
    rsu_df['DNI'] = rsu_df['DNI'].apply(normalize_dni_value)
    rsu_df = rsu_df.rename(columns={'Tarea: Producto final': 'rsu'})
    all_data = pd.merge(
        all_data, rsu_df[['DNI', 'rsu']],
        on='DNI', how='left'
    )

    # ---- estress (by DNI)
    estress_df['DNI'] = estress_df['DNI'].apply(normalize_dni_value)
    estress_df = estress_df.rename(columns={'Tarea:Producto final': 'estress'})
    all_data = pd.merge(
        all_data, estress_df[['DNI', 'estress']],
        on='DNI', how='left'
    )

    # ---- Hab. comunicación (by DNI)
    hab_com_df['DNI'] = hab_com_df['DNI'].apply(normalize_dni_value)
    hab_com_df = hab_com_df.rename(columns={'Tarea:Producto final': 'hab_comunicacion'})
    all_data = pd.merge(
        all_data, hab_com_df[['DNI', 'hab_comunicacion']],
        on='DNI', how='left'
    )

    # ---- Numeric components (14 total)
    numeric_columns = [
        'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
        'Padlet', 'Nearpod', 'Tareas_y_foros',
        'integracion', 'rsu', 'estress', 'hab_comunicacion'
    ]
    for col in numeric_columns:
        all_data[col] = pd.to_numeric(all_data.get(col, 0), errors='coerce').fillna(0)

    # ---- Compute metrics (before roster merge)
    def calculate_percentage_row(row):
        scores = row[numeric_columns].values
        available = int(np.sum(np.array(scores) > 0))
        return round(available / len(numeric_columns) * 100, 2) if len(numeric_columns) else 0.0

    all_data['Average'] = all_data[numeric_columns].mean(axis=1).round(2)
    all_data['Percentage'] = all_data.apply(calculate_percentage_row, axis=1)
    all_data['Marks_Out_Of_20'] = (all_data['Percentage'] / 5).round(2)

    # -----------------------------
    # Merge with C9 roster to include all teachers (even all-zero)
    # -----------------------------
    if c9_file is not None:
        roster = load_c9_roster(c9_file)
        roster['DNI'] = roster['DNI'].apply(normalize_dni_value)
        roster = roster.dropna(subset=['DNI'])
        roster = roster.drop_duplicates(subset=['DNI'])

        # Left join: keep ALL teachers from C9; bring scores if available
        merged = pd.merge(
            roster, all_data,
            on='DNI', how='left', suffixes=('_roster', '')
        )

        # Names: prefer Master if present, otherwise use C9
        merged['Nombre'] = merged['Nombre'].fillna(merged['Nombre_roster'])
        merged['Apellido(s)'] = merged['Apellido(s)'].fillna(merged['Apellido(s)_roster'])

        # Email: prefer C9 email; fallback to Master 'Dirección de correo' if empty
        if 'Email' not in merged.columns:
            merged['Email'] = merged.get('Email_roster', "")
        else:
            merged['Email'] = merged['Email']  # already from roster
        if 'Dirección de correo' in merged.columns:
            merged['Email'] = merged['Email'].replace(['', None, np.nan], np.nan)
            merged['Email'] = merged['Email'].fillna(merged['Dirección de correo'])
        merged['Email'] = merged['Email'].fillna("")

        # Ensure numeric columns exist and are numeric
        for col in numeric_columns:
            if col not in merged.columns:
                merged[col] = 0
            merged[col] = pd.to_numeric(merged[col], errors='coerce').fillna(0)

        # Fill missing Year/Periodo
        # Year hint: try master values (prefer latest among {2024, 2025}), else infer from C9 filename, else 2025
        year_candidates = pd.Series(merged['Year'].dropna(), dtype='float')
        year_hint = None
        if not year_candidates.empty:
            # Prefer 2025 if present, else 2024, else the max available
            years_present = sorted({int(y) for y in year_candidates if 2000 <= y <= 2100}, reverse=True)
            for y in (2025, 2024):
                if y in years_present:
                    year_hint = y
                    break
            if year_hint is None and years_present:
                year_hint = years_present[0]
        if year_hint is None:
            # Try from uploaded filename if available
            uploaded_name = getattr(c9_file, 'name', '') or ''
            y = infer_year_from_text(uploaded_name)
            year_hint = int(y) if (y is not pd.NA and pd.notna(y)) else 2025

        merged['Year'] = merged['Year'].fillna(year_hint).astype('Int64')
        merged['Periodo'] = merged['Periodo'].fillna(merged['Year'].astype(str))

        # Recompute metrics post-merge to cover newly created rows
        merged['Average'] = merged[numeric_columns].mean(axis=1).round(2)
        merged['Percentage'] = merged.apply(lambda r: calculate_percentage_row(r), axis=1)
        merged['Marks_Out_Of_20'] = (merged['Percentage'] / 5).round(2)

        # Keep only 2024 & 2025 comparisons
        merged = merged[merged['Year'].isin([2024, 2025])].copy()

        # If everything is still empty (edge case), return formatted empty
        if merged.empty:
            return pd.DataFrame(columns=[
                'Periodo', 'Highest_Score_Year', 'DNI', 'Nombre', 'Apellido(s)', 'Email',
                'induccion', 'bus_biblioteca', 'diseno_sesion',
                'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
                'Padlet', 'Nearpod', 'Tareas_y_foros',
                'integracion', 'rsu', 'estress', 'hab_comunicacion',
                'Average', 'Marks_Out_Of_20', 'Percentage'
            ])

        # ---- DEDUP: one row per teacher (by DNI) with the HIGHEST Marks_Out_Of_20 across 2024 & 2025
        # Tie-breaker: higher Average, then prefer 2025 over 2024
        merged['YearPref'] = merged['Year'].apply(lambda y: 1 if y == 2025 else 0)
        sorted_df = merged.sort_values(
            by=['Marks_Out_Of_20', 'Average', 'YearPref'],
            ascending=[False, False, False]
        )
        highest = sorted_df.drop_duplicates(subset=['DNI'], keep='first').copy()
        highest['Highest_Score_Year'] = highest['Year']

        # Final column order (with Email)
        final_columns = [
            'Periodo', 'Highest_Score_Year', 'DNI', 'Nombre', 'Apellido(s)', 'Email',
            'induccion', 'bus_biblioteca', 'diseno_sesion',
            'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
            'Padlet', 'Nearpod', 'Tareas_y_foros',
            'integracion', 'rsu', 'estress', 'hab_comunicacion',
            'Average', 'Marks_Out_Of_20', 'Percentage'
        ]
        # Ensure all columns exist
        for c in final_columns:
            if c not in highest.columns:
                highest[c] = "" if c in ['Periodo', 'Nombre', 'Apellido(s)', 'Email'] else 0
        final_df = highest[final_columns]
        return final_df

    # -----------------------------
    # Fallback path (no C9 uploaded) — still include zero-sum rows but only those present in Master
    # -----------------------------
    filtered = all_data[all_data['Year'].isin([2024, 2025])].copy()
    # NOTE: We NO LONGER drop zero-score rows; keep everyone in Master even if all zeros.

    if filtered.empty:
        return pd.DataFrame()

    filtered['YearPref'] = filtered['Year'].apply(lambda y: 1 if y == 2025 else 0)
    sorted_df = filtered.sort_values(
        by=['Marks_Out_Of_20', 'Average', 'YearPref'],
        ascending=[False, False, False]
    )
    highest = sorted_df.drop_duplicates(subset=['DNI'], keep='first').copy()
    highest['Highest_Score_Year'] = highest['Year']

    final_columns = [
        'Periodo', 'Highest_Score_Year', 'DNI', 'Nombre', 'Apellido(s)',
        'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
        'Padlet', 'Nearpod', 'Tareas_y_foros',
        'integracion', 'rsu', 'estress', 'hab_comunicacion',
        'Average', 'Marks_Out_Of_20', 'Percentage'
    ]
    final_df = highest[final_columns]
    return final_df

# -----------------------------
# Streamlit App
# -----------------------------
def main():
    st.set_page_config(page_title="📊 UMA Scores (Highest of 2024 vs 2025)", page_icon="📊", layout="wide")

    st.title("📊 UMA Scores — Highest Marks (2024 vs 2025)")
    st.markdown(
        "Upload the **Master** Excel and the **C9 Roster** Excel. "
        "The output will contain **only one row per teacher (no duplicates)** — "
        "the row corresponding to the **highest _Marks Out Of 20_ between 2024 and 2025**.\n\n"
        "Now includes **all teachers from C9** (even if they have zero marks) and adds **Email** to the export."
    )

    # File uploaders
    uploaded_file = st.file_uploader("Choose the Master Excel file", type=["xlsx", "xls"], key="master")
    uploaded_c9 = st.file_uploader("Choose the C9 Roster Excel file", type=["xlsx", "xls"], key="c9")

    if uploaded_file is not None and uploaded_c9 is not None:
        try:
            with st.spinner("Processing your Excel files and comparing 2024 vs 2025..."):
                final_data = extract_data_from_excel(uploaded_file, c9_file=uploaded_c9)

            if len(final_data) == 0:
                st.warning("No records found for 2024 or 2025 (after merging with the C9 roster).")
                return

            st.success("Done! Showing only the highest marks per teacher (no duplicates).")

            # Preview
            st.subheader("Preview")
            st.dataframe(final_data.head(30))

            # Metrics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Teachers", len(final_data))
            with col2:
                st.metric("Avg Marks (Out of 20)", f"{final_data['Marks_Out_Of_20'].mean():.2f}")
            with col3:
                st.metric("Avg Percentage", f"{final_data['Percentage'].mean():.2f}%")

            st.subheader("Distribution by Highest Score Year")
            year_counts = final_data['Highest_Score_Year'].value_counts().sort_index()
            st.bar_chart(year_counts)

            # Download
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"Highest_Marks_2024_vs_2025_{timestamp}.xlsx"
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_data.to_excel(writer, index=False, sheet_name='Highest Marks (Unique)')
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
                "Ensure your Master file contains these sheets: "
                "'Inducción', 'nota Inducción', 'Bus. biblioteca', 'Diseño de sesión', 'Comp. Tec', "
                "'Integración', 'RSU', 'estress', 'Hab. comunicación'.\n"
                "C9 Roster should include: DNI, nombres, apellidos y email."
            )
    else:
        st.info("👆 Please upload both the Master Excel and the C9 Roster Excel to get started.")

        st.subheader("Expected Columns (quick reference)")
        st.markdown("""
        **Master Excel**:
        - **Inducción**: `Periodo`, `DNI`, `Nombre`, `Apellido(s)`, `Dirección de correo`, `Calificación`
        - **nota Inducción**: `PERIODO`, `DNI`, `Nombre`, `Apellido(s)`, `Dirección de correo`, `Total del curso (Real)`
        - **Bus. biblioteca**: `DNI`, `Promedio`
        - **Diseño de sesión**: `Nombre`, `Apellido(s)`, `Promedio`
        - **Comp. Tec**: `Reto` columns (Zoom básico, Zoom Avanzado, Grupos Moodle, Rúbrica, Padlet, Nearpod, Tareas y foros)
        - **Integración**: `Tarea:Producto final: Contenido académico, presentación y rúbrica con IA (Real)`
        - **RSU**: `Tarea: Producto final`
        - **estress**: `Tarea:Producto final`
        - **Hab. comunicación**: `Tarea:Producto final`

        **C9 Roster**:
        - `N° DE DOCUMENTO DE IDENTIDAD` (DNI), `NOMBRES`, `APELLIDO PATERNO`, `APELLIDO MATERNO`, and **`email`** (or equivalent like *Correo*).
        """)

if __name__ == "__main__":
    main()
