# app.py
import pandas as pd
import numpy as np
import streamlit as st
import io
import re
from datetime import datetime

# =============================
# Helpers
# =============================
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
    Infer a year from a filename/text (e.g., 'C9_25-II_121025.xlsx' -> 2025).
    """
    if not text:
        return pd.NA
    s = str(text)
    m4 = re.search(r'20\d{2}', s)
    if m4:
        return int(m4.group(0))
    m2 = re.search(r'(?<!\d)(2[0-9])(?!\d)', s)
    if m2:
        return 2000 + int(m2.group(1))
    return pd.NA

def safe_read_excel(file_like, sheet_name):
    """Always read with openpyxl (more reliable on Render)."""
    return pd.read_excel(file_like, sheet_name=sheet_name, engine='openpyxl')

# =============================
# C9 roster loader (replaces teacher contract)
# =============================
def load_c9_roster(c9_file):
    """
    Load the C9 roster and extract ['DNI','Nombre','Apellido(s)','Email'].
    Tolerant with header variations commonly seen in C9 exports.
    """
    df = safe_read_excel(c9_file, sheet_name=0)

    def pick(colmap, *candidates):
        # exact first
        for cand in candidates:
            if cand in colmap:
                return colmap[cand]
        # fuzzy contains
        for col in df.columns:
            if not isinstance(col, str):
                continue
            up = col.upper()
            for cand in candidates:
                if isinstance(cand, str) and cand.upper() in up:
                    return col
        return None

    cols = {(c.strip() if isinstance(c, str) else c): c for c in df.columns}

    dni_col = pick(cols, 'N° DE DOCUMENTO DE IDENTIDAD', 'NRO DE DOCUMENTO DE IDENTIDAD',
                   'NRO DE DOCUMENTO', 'DNI', 'DOCUMENTO', 'DOC')
    nombres_col = pick(cols, 'NOMBRES', 'Nombres', 'NOMBRE', 'NOMBRE(S)')
    ap_pat_col  = pick(cols, 'APELLIDO PATERNO', 'Apellido Paterno', 'APE. PATERNO', 'APELLIDO  PATERNO')
    ap_mat_col  = pick(cols, 'APELLIDO MATERNO', 'Apellido Materno', 'APE. MATERNO', 'APELLIDO  MATERNO')
    email_col   = pick(cols, 'email', 'Email', 'EMAIL', 'Correo', 'CORREO', 'Dirección de correo',
                       'CORREO ELECTRÓNICO', 'Correo electrónico', 'Correo Institucional',
                       'CORREO INSTITUCIONAL', 'Correo Inst.', 'E-MAIL')

    if dni_col is None:
        return pd.DataFrame(columns=['DNI', 'Nombre', 'Apellido(s)', 'Email'])

    out = pd.DataFrame()
    out['DNI'] = df[dni_col].apply(normalize_dni_value)
    out['Nombre'] = df[nombres_col].astype(str).str.strip() if nombres_col is not None else ""
    if ap_pat_col is not None and ap_mat_col is not None:
        out['Apellido(s)'] = (
            df[ap_pat_col].astype(str).str.strip() + ' ' +
            df[ap_mat_col].astype(str).str.strip()
        ).str.replace(r'\s+', ' ', regex=True).str.strip()
    else:
        out['Apellido(s)'] = ""
    out['Email'] = df[email_col].astype(str).str.strip() if email_col is not None else ""

    out = out.dropna(subset=['DNI'])
    out = out[out['DNI'] != '']
    out = out.drop_duplicates(subset=['DNI'])
    return out[['DNI', 'Nombre', 'Apellido(s)', 'Email']]

# =============================
# Core processing
# =============================
def extract_data_from_excel(file_path, c9_file=None):
    # ---- Read all source sheets (Master)
    induction_df       = safe_read_excel(file_path, sheet_name='Inducción')
    nota_induccion_df  = safe_read_excel(file_path, sheet_name='nota Inducción')
    bus_biblioteca_df  = safe_read_excel(file_path, sheet_name='Bus. biblioteca')
    diseno_sesion_df   = safe_read_excel(file_path, sheet_name='Diseño de sesión')
    comp_tec_df        = safe_read_excel(file_path, sheet_name='Comp. Tec')
    integracion_df     = safe_read_excel(file_path, sheet_name='Integración')
    rsu_df             = safe_read_excel(file_path, sheet_name='RSU')
    estress_df         = safe_read_excel(file_path, sheet_name='estress')
    hab_com_df         = safe_read_excel(file_path, sheet_name='Hab. comunicación')

    # ---- Base: Inducción + nota Inducción
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
    rsu_df = rsu_df.rename(columns={'Tarea: Producto final': 'rsu'})
    all_data = pd.merge(all_data, rsu_df[['DNI', 'rsu']], on='DNI', how='left')

    estress_df['DNI'] = estress_df['DNI'].apply(normalize_dni_value)
    estress_df = estress_df.rename(columns={'Tarea:Producto final': 'estress'})
    all_data = pd.merge(all_data, estress_df[['DNI', 'estress']], on='DNI', how='left')

    hab_com_df['DNI'] = hab_com_df['DNI'].apply(normalize_dni_value)
    hab_com_df = hab_com_df.rename(columns={'Tarea:Producto final': 'hab_comunicacion'})
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

    # ---- Compute metrics (pre-merge)
    def availability_pct(row):
        scores = row[numeric_columns].values
        available = int(np.sum(np.array(scores) > 0))
        return round((available / len(numeric_columns)) * 100, 2) if numeric_columns else 0.0

    all_data['Average'] = all_data[numeric_columns].mean(axis=1).round(2)
    all_data['Percentage'] = all_data.apply(availability_pct, axis=1)
    all_data['Marks_Out_Of_20'] = (all_data['Percentage'] / 5).round(2)

    # ============================
    # Merge with C9 roster to include ALL teachers (even all-zero)
    # ============================
    if c9_file is not None:
        roster = load_c9_roster(c9_file)
        roster['DNI'] = roster['DNI'].apply(normalize_dni_value)
        roster = roster.dropna(subset=['DNI']).drop_duplicates(subset=['DNI'])

        merged = pd.merge(roster, all_data, on='DNI', how='left', suffixes=('_roster', ''))

        # Names: prefer Master if present
        merged['Nombre'] = merged['Nombre'].fillna(merged['Nombre_roster'])
        merged['Apellido(s)'] = merged['Apellido(s)'].fillna(merged['Apellido(s)_roster'])

        # Email: prefer C9; fallback to Master 'Dirección de correo'
        if 'Email' not in merged.columns:
            merged['Email'] = merged.get('Email_roster', "")
        else:
            merged['Email'] = merged['Email']
        if 'Dirección de correo' in merged.columns:
            merged['Email'] = merged['Email'].replace(['', None, np.nan], np.nan)
            merged['Email'] = merged['Email'].fillna(merged['Dirección de correo'])
        merged['Email'] = merged['Email'].fillna("")

        # Ensure numeric columns exist and numeric
        for col in numeric_columns:
            if col not in merged.columns:
                merged[col] = 0
            merged[col] = pd.to_numeric(merged[col], errors='coerce').fillna(0)

        # Year hint
        years_present = sorted(
            {int(y) for y in pd.to_numeric(merged.get('Year'), errors='coerce').dropna().tolist()
             if 2000 <= int(y) <= 2100},
            reverse=True
        )
        year_hint = 2025 if 2025 in years_present else (2024 if 2024 in years_present else (years_present[0] if years_present else None))
        if year_hint is None:
            inferred = infer_year_from_text(getattr(c9_file, 'name', '') or '')
            year_hint = int(inferred) if (inferred is not pd.NA and pd.notna(inferred)) else 2025

        # **Future-proof fix**: coerce to numeric before fillna+astype
        merged['Year'] = pd.to_numeric(merged.get('Year'), errors='coerce').fillna(year_hint).astype('Int64')
        merged['Periodo'] = merged['Periodo'].fillna(merged['Year'].astype(str))

        # Recompute metrics post-merge
        merged['Average'] = merged[numeric_columns].mean(axis=1).round(2)
        merged['Percentage'] = merged.apply(availability_pct, axis=1)
        merged['Marks_Out_Of_20'] = (merged['Percentage'] / 5).round(2)

        # Only 2024/2025
        merged = merged[merged['Year'].isin([2024, 2025])].copy()

        if merged.empty:
            return pd.DataFrame(columns=[
                'Periodo', 'Highest_Score_Year', 'DNI', 'Nombre', 'Apellido(s)', 'Email',
                'induccion', 'bus_biblioteca', 'diseno_sesion',
                'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
                'Padlet', 'Nearpod', 'Tareas_y_foros',
                'integracion', 'rsu', 'estress', 'hab_comunicacion',
                'Average', 'Marks_Out_Of_20', 'Percentage'
            ])

        # Deduplicate: keep row with highest Marks_Out_Of_20 (tie -> higher Average -> prefer 2025)
        merged['YearPref'] = merged['Year'].apply(lambda y: 1 if y == 2025 else 0)
        sorted_df = merged.sort_values(
            by=['Marks_Out_Of_20', 'Average', 'YearPref'],
            ascending=[False, False, False]
        )
        highest = sorted_df.drop_duplicates(subset=['DNI'], keep='first').copy()
        highest['Highest_Score_Year'] = highest['Year']

        final_columns = [
            'Periodo', 'Highest_Score_Year', 'DNI', 'Nombre', 'Apellido(s)', 'Email',
            'induccion', 'bus_biblioteca', 'diseno_sesion',
            'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
            'Padlet', 'Nearpod', 'Tareas_y_foros',
            'integracion', 'rsu', 'estress', 'hab_comunicacion',
            'Average', 'Marks_Out_Of_20', 'Percentage'
        ]
        for c in final_columns:
            if c not in highest.columns:
                highest[c] = "" if c in ['Periodo', 'Nombre', 'Apellido(s)', 'Email'] else 0
        return highest[final_columns]

    # -----------------------------
    # Fallback path (no C9 uploaded) — keep everyone present in Master (zeros included)
    # -----------------------------
    filtered = all_data[all_data['Year'].isin([2024, 2025])].copy()

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
    return highest[final_columns]

# =============================
# Streamlit App
# =============================
def main():
    st.set_page_config(page_title="📊 UMA Scores (Highest of 2024 vs 2025)", page_icon="📊", layout="wide")

    st.title("📊 UMA Scores — Highest Marks (2024 vs 2025)")
    st.markdown(
        "Upload the **Master** Excel and the **C9 Roster** Excel. "
        "The output shows **one row per teacher (no duplicates)** — the row with the **highest _Marks Out Of 20_** "
        "between **2024** and **2025**. Includes **all teachers from C9** (even all-zero) and adds **Email**."
    )

    uploaded_file = st.file_uploader("Choose the Master Excel file", type=["xlsx", "xls"], key="master")
    uploaded_c9   = st.file_uploader("Choose the C9 Roster Excel file", type=["xlsx", "xls"], key="c9")

    if uploaded_file is not None and uploaded_c9 is not None:
        try:
            with st.status("Processing your Excel files…", expanded=True) as s:
                st.write("Reading Master & C9…")
                final_data = extract_data_from_excel(uploaded_file, c9_file=uploaded_c9)
                st.write("Computing metrics & selecting highest marks…")
                s.update(label="Done", state="complete")

            if len(final_data) == 0:
                st.warning("No records found for 2024 or 2025 (after merging with the C9 roster).")
                return

            st.success("Ready! Showing one row per teacher (highest marks of 2024 vs 2025).")

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
            year_counts = final_data['Highest_Score_Year'].value_counts().sort_index()
            st.bar_chart(year_counts)

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
                "Make sure the Master has sheets: 'Inducción', 'nota Inducción', 'Bus. biblioteca', "
                "'Diseño de sesión', 'Comp. Tec', 'Integración', 'RSU', 'estress', 'Hab. comunicación'.\n"
                "C9 should contain DNI, Nombres, Apellidos, and Email."
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
        - **Comp. Tec**: Reto columns (Zoom básico, Zoom Avanzado, Grupos Moodle, Rúbrica, Padlet, Nearpod, Tareas y foros)
        - **Integración**: `Tarea:Producto final: Contenido académico, presentación y rúbrica con IA (Real)`
        - **RSU**: `Tarea: Producto final`
        - **estress**: `Tarea:Producto final`
        - **Hab. comunicación**: `Tarea:Producto final`

        **C9 Roster**:
        - `N° DE DOCUMENTO DE IDENTIDAD` (DNI), `NOMBRES`, `APELLIDO PATERNO`, `APELLIDO MATERNO`, **Email/Correo**.
        """)

if __name__ == "__main__":
    main()
