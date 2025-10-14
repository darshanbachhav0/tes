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

def clean_spaces(s):
    if pd.isna(s):
        return pd.NA
    return re.sub(r'\s+', ' ', str(s).strip())

def canonicalize_email(x):
    """Lower-case, trimmed email; returns NA if not an email."""
    if pd.isna(x):
        return pd.NA
    s = str(x).strip().lower()
    if '@' in s and '.' in s.split('@')[-1]:
        return s
    return pd.NA

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

def _pick_first_present(df, candidates=None, contains_all=None, contains_any=None):
    """
    Flexible header finder.
    """
    candidates = candidates or []
    cols = [c for c in df.columns if isinstance(c, str)]
    # exact
    for k in candidates:
        if k in cols:
            return k
    lc_map = {c.lower(): c for c in cols}
    for k in candidates:
        if k.lower() in lc_map:
            return lc_map[k.lower()]
    # contains_all
    if contains_all:
        for c in cols:
            up = c.upper()
            if all(sub.upper() in up for sub in contains_all):
                return c
    # contains_any
    if contains_any:
        for c in cols:
            up = c.upper()
            if any(sub.upper() in up for sub in contains_any):
                return c
    return None

# -----------------------------
# C9 loader (authoritative roster)
# -----------------------------
def load_c9_file(c9_file):
    """
    Load the C9 roster and extract:
      ['DNI','Nombre','Apellido(s)','Email','Periodo','Year']
    """
    df = pd.read_excel(c9_file, sheet_name=0)

    dni_col = _pick_first_present(
        df,
        candidates=['DNI', 'dni', 'N° DE DOCUMENTO DE IDENTIDAD', 'NRO DE DOCUMENTO DE IDENTIDAD',
                    'NRO DE DOCUMENTO', 'NUMERO DE DOCUMENTO'],
        contains_all=['DOCUMENTO', 'IDENTIDAD']
    )
    nombres_col = _pick_first_present(df, candidates=['NOMBRES', 'Nombres', 'Nombre'])
    a_pat_col   = _pick_first_present(df, candidates=['APELLIDO PATERNO', 'Apellido Paterno', 'Apellidos'])
    a_mat_col   = _pick_first_present(df, candidates=['APELLIDO MATERNO', 'Apellido Materno'])
    email_col   = _pick_first_present(
        df,
        candidates=['Dirección de correo', 'Correo', 'Correo institucional', 'Correo Institucional',
                    'Correo electrónico', 'Correo Electronico', 'EMAIL', 'Email', 'E-mail'],
        contains_any=['CORREO', 'EMAIL']
    )
    periodo_col = _pick_first_present(
        df,
        candidates=['Periodo', 'PERIODO', 'Periodo Académico', 'PERIODO ACADÉMICO', 'PERIODO ACADEMICO'],
        contains_any=['PERIODO', 'PERÍODO', 'ACADEM']
    )

    out = pd.DataFrame()
    out['DNI'] = df[dni_col].apply(normalize_dni_value) if dni_col else np.nan

    out['Nombre'] = df[nombres_col].astype(str).map(clean_spaces) if nombres_col else pd.NA

    if a_pat_col is not None and a_mat_col is not None:
        ap = df[a_pat_col].astype(str).map(clean_spaces)
        am = df[a_mat_col].astype(str).map(clean_spaces)
        out['Apellido(s)'] = (ap.fillna('') + ' ' + am.fillna('')).str.replace(r'\s+', ' ', regex=True).str.strip()
    elif a_pat_col is not None:
        out['Apellido(s)'] = df[a_pat_col].astype(str).map(clean_spaces)
    else:
        out['Apellido(s)'] = pd.NA

    out['Email'] = df[email_col].astype(str).map(clean_spaces) if email_col else pd.NA
    out['Email'] = out['Email'].map(canonicalize_email)

    if periodo_col:
        out['Periodo'] = df[periodo_col]
    else:
        # Try infer from file name (e.g., C9_25-II_*.xlsx)
        fname = getattr(c9_file, 'name', '') or ''
        yr = None
        m4 = re.search(r'(20\d{2})', fname)
        if m4:
            yr = int(m4.group(1))
        else:
            m2 = re.search(r'(^|[^0-9])(2[45])([^0-9]|$)', fname)  # 24/25
            if m2:
                yr = 2000 + int(m2.group(2))
        out['Periodo'] = f'{yr}' if yr else pd.NA

    out['Periodo'] = out['Periodo'].astype(str).str.strip()
    out['Year'] = out['Periodo'].apply(extract_year)

    out['DNI'] = out['DNI'].apply(normalize_dni_value)
    out = out.dropna(subset=['DNI'])
    out = out[out['DNI'] != '']
    out = out.drop_duplicates(subset=['DNI'])
    return out[['DNI', 'Nombre', 'Apellido(s)', 'Email', 'Periodo', 'Year']]

# -----------------------------
# Core processing
# -----------------------------
def extract_data_from_excel(file_path, c9_file=None):
    # ---- Read all source sheets
    induction_df       = pd.read_excel(file_path, sheet_name='Inducción')
    nota_induccion_df  = pd.read_excel(file_path, sheet_name='nota Inducción')
    bus_biblioteca_df  = pd.read_excel(file_path, sheet_name='Bus. biblioteca')
    diseno_sesion_df   = pd.read_excel(file_path, sheet_name='Diseño de sesión')
    comp_tec_df        = pd.read_excel(file_path, sheet_name='Comp. Tec')
    integracion_df     = pd.read_excel(file_path, sheet_name='Integración')
    rsu_df             = pd.read_excel(file_path, sheet_name='RSU')
    estress_df         = pd.read_excel(file_path, sheet_name='estress')
    hab_com_df         = pd.read_excel(file_path, sheet_name='Hab. comunicación')

    # ---- Base (Inducción + nota Inducción)
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
    all_data['DNI'] = all_data['DNI'].apply(normalize_dni_value)
    all_data['Nombre'] = all_data['Nombre'].astype(str).map(clean_spaces)
    all_data['Apellido(s)'] = all_data['Apellido(s)'].astype(str).map(clean_spaces)
    all_data['Year'] = all_data['Periodo'].apply(extract_year)

    # ---- Roster (C9) – authoritative for names & email
    if c9_file is not None:
        c9 = load_c9_file(c9_file)
        if 'Email' not in c9.columns:
            c9['Email'] = pd.NA

        # Append C9-only teachers so zero-mark teachers appear
        known = set(all_data['DNI'].dropna())
        c9_only = c9[~c9['DNI'].isin(known)].copy()
        if not c9_only.empty:
            extra = pd.DataFrame({
                'Periodo': c9_only['Periodo'],
                'DNI': c9_only['DNI'],
                'Nombre': c9_only['Nombre'],
                'Apellido(s)': c9_only['Apellido(s)'],
                'Dirección de correo': c9_only['Email'],
                'induccion': np.nan
            })
            all_data = pd.concat([all_data, extra], ignore_index=True)

        # Keep only DNIs present in C9
        all_data = all_data[all_data['DNI'].isin(c9['DNI'])].copy()

        # Merge roster & OVERWRITE names with C9 (canonical)
        all_data = pd.merge(
            all_data,
            c9[['DNI', 'Nombre', 'Apellido(s)', 'Email', 'Periodo', 'Year']],
            on='DNI', how='left', suffixes=('', '_c9')
        )
        # Overwrite names with C9 to avoid mismatches that look like duplicates
        all_data['Nombre'] = np.where(all_data['Nombre_c9'].notna(), all_data['Nombre_c9'], all_data['Nombre'])
        all_data['Apellido(s)'] = np.where(all_data['Apellido(s)_c9'].notna(), all_data['Apellido(s)_c9'], all_data['Apellido(s)'])
        # Prefer existing Periodo/Year; fill with C9 when missing
        all_data['Periodo'] = all_data['Periodo'].fillna(all_data['Periodo_c9'])
        all_data['Year'] = all_data['Year'].fillna(all_data['Year_c9'])
        # Email from C9; fallback to 'Dirección de correo'
        all_data['Email'] = all_data['Email'].map(canonicalize_email)
        if 'Dirección de correo' in all_data.columns:
            all_data['Email'] = all_data['Email'].fillna(all_data['Dirección de correo'].map(canonicalize_email))

        # Clean up merge helper columns
        drop_cols = [c for c in ['Nombre_c9', 'Apellido(s)_c9', 'Periodo_c9', 'Year_c9'] if c in all_data]
        all_data = all_data.drop(columns=drop_cols)

    # ---- Bus. biblioteca (by DNI)
    bus_biblioteca_df['DNI'] = bus_biblioteca_df['DNI'].apply(normalize_dni_value)
    bus_biblioteca_df = bus_biblioteca_df.rename(columns={'Promedio': 'bus_biblioteca'})
    all_data = pd.merge(all_data, bus_biblioteca_df[['DNI', 'bus_biblioteca']], on='DNI', how='left')

    # ---- RSU / estress / Hab. comunicación (by DNI)
    for src_df, col_src, col_dst in [
        (rsu_df, 'Tarea: Producto final', 'rsu'),
        (estress_df, 'Tarea:Producto final', 'estress'),
        (hab_com_df, 'Tarea:Producto final', 'hab_comunicacion')
    ]:
        src_df['DNI'] = src_df['DNI'].apply(normalize_dni_value)
        tmp = src_df.rename(columns={col_src: col_dst})
        all_data = pd.merge(all_data, tmp[['DNI', col_dst]], on='DNI', how='left')

    # ---- Clean names in name-joined modules and pre-aggregate to avoid row multiplication
    for df_name in [diseno_sesion_df, integracion_df, comp_tec_df]:
        if 'Nombre' in df_name.columns:
            df_name['Nombre'] = df_name['Nombre'].astype(str).map(clean_spaces)
        if 'Apellido(s)' in df_name.columns:
            df_name['Apellido(s)'] = df_name['Apellido(s)'].astype(str).map(clean_spaces)

    # Diseño de sesión (grouped by name)
    if 'Promedio' in diseno_sesion_df.columns:
        dis_tmp = diseno_sesion_df[['Nombre', 'Apellido(s)', 'Promedio']].copy()
        dis_tmp = dis_tmp.rename(columns={'Promedio': 'diseno_sesion'})
        dis_tmp = dis_tmp.groupby(['Nombre', 'Apellido(s)'], as_index=False)['diseno_sesion'].max()
        all_data = pd.merge(all_data, dis_tmp, on=['Nombre', 'Apellido(s)'], how='left')

    # Integración (grouped by name)
    integ_col = 'Tarea:Producto final: Contenido académico, presentación y rúbrica con IA (Real)'
    if integ_col in integracion_df.columns:
        int_tmp = integracion_df[['Nombre', 'Apellido(s)', integ_col]].copy()
        int_tmp = int_tmp.rename(columns={integ_col: 'integracion'})
        int_tmp = int_tmp.groupby(['Nombre', 'Apellido(s)'], as_index=False)['integracion'].max()
        all_data = pd.merge(all_data, int_tmp, on=['Nombre', 'Apellido(s)'], how='left')

    # Comp. Tec (grouped by name, max across columns)
    comp_tec_columns = {
        'Cuestionario:Reto: Zoom básico': 'Zoom_basico',
        'Cuestionario:Reto: Zoom Avanzado': 'Zoom_Avanzado',
        'Cuestionario:Reto: Grupos Moodle': 'Grupos_Moodle',
        'Cuestionario:Reto: Rúbrica': 'Rubrica',
        'Cuestionario:Reto: Padlet': 'Padlet',
        'Cuestionario:Reto: Nearpod': 'Nearpod',
        'Cuestionario:Reto: Tareas y foros': 'Tareas_y_foros'
    }
    present_map = {k: v for k, v in comp_tec_columns.items() if k in comp_tec_df.columns}
    if present_map:
        ct = comp_tec_df.rename(columns=present_map)
        keep_cols = ['Nombre', 'Apellido(s)'] + list(present_map.values())
        ct = ct[keep_cols].copy()
        # coerce to numeric
        for c in present_map.values():
            ct[c] = pd.to_numeric(ct[c], errors='coerce')
        # group by name (max per column)
        agg = {c: 'max' for c in present_map.values()}
        ct = ct.groupby(['Nombre', 'Apellido(s)'], as_index=False).agg(agg)
        all_data = pd.merge(all_data, ct, on=['Nombre', 'Apellido(s)'], how='left')

    # ---- Numeric components (14 total; coerce and fill zeros)
    numeric_columns = [
        'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
        'Padlet', 'Nearpod', 'Tareas_y_foros',
        'integracion', 'rsu', 'estress', 'hab_comunicacion'
    ]
    for col in numeric_columns:
        if col not in all_data.columns:
            all_data[col] = 0
        all_data[col] = pd.to_numeric(all_data[col], errors='coerce').fillna(0)

    # ---- Compute metrics
    all_data['Average'] = all_data[numeric_columns].mean(axis=1).round(2)

    def calculate_percentage(row):
        scores = row[numeric_columns].values
        available = int(np.sum(np.array(scores) > 0))
        return round(available / len(numeric_columns) * 100, 2) if len(numeric_columns) else 0.0

    all_data['Percentage'] = all_data.apply(calculate_percentage, axis=1)
    all_data['Marks_Out_Of_20'] = (all_data['Percentage'] / 5).round(2)

    # ---- Only compare 2024 vs 2025 (include zero-mark teachers too)
    filtered = all_data[all_data['Year'].isin([2024, 2025])].copy()
    if filtered.empty:
        return pd.DataFrame()

    # ---- Sort for "best row" selection
    filtered['YearPref'] = filtered['Year'].apply(lambda y: 1 if y == 2025 else 0)
    sorted_df = filtered.sort_values(
        by=['Marks_Out_Of_20', 'Average', 'YearPref'],
        ascending=[False, False, False]
    )

    # ---- Person-level de-duplication (email first, else DNI)
    sorted_df['Email'] = sorted_df.get('Email', pd.Series(index=sorted_df.index)).map(canonicalize_email)
    if 'Dirección de correo' in sorted_df.columns:
        sorted_df['Email'] = sorted_df['Email'].fillna(sorted_df['Dirección de correo'].map(canonicalize_email))

    sorted_df['person_key'] = sorted_df['Email']
    sorted_df.loc[sorted_df['person_key'].isna() | (sorted_df['person_key'] == ''), 'person_key'] = sorted_df['DNI']

    highest = sorted_df.drop_duplicates(subset=['person_key'], keep='first').copy()
    highest['Highest_Score_Year'] = highest['Year']

    # ---- Final column order (with Email)
    final_columns = [
        'Periodo', 'Highest_Score_Year', 'DNI', 'Nombre', 'Apellido(s)', 'Email',
        'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
        'Padlet', 'Nearpod', 'Tareas_y_foros',
        'integracion', 'rsu', 'estress', 'hab_comunicacion',
        'Average', 'Marks_Out_Of_20', 'Percentage'
    ]
    final_columns = [c for c in final_columns if c in highest.columns]
    final_df = highest[final_columns].reset_index(drop=True)
    return final_df

# -----------------------------
# Streamlit App
# -----------------------------
def main():
    st.set_page_config(page_title="📊 UMA Scores (Highest of 2024 vs 2025)", page_icon="📊", layout="wide")

    st.title("📊 UMA Scores — Highest Marks (2024 vs 2025)")
    st.markdown(
        "Upload the **Master** Excel and the **C9 roster** Excel. "
        "The output contains **one row per teacher** (no duplicates), using **email if present, else DNI** "
        "to define uniqueness. Names & Email are taken from **C9**. "
        "Teachers listed in C9 are included **even if all module scores are 0**."
    )

    uploaded_file = st.file_uploader("Choose the Master Excel file", type=["xlsx", "xls"], key="master")
    uploaded_c9   = st.file_uploader("Choose the C9 roster Excel file (for names & email)", type=["xlsx", "xls"], key="c9")

    if uploaded_file is not None and uploaded_c9 is not None:
        try:
            with st.spinner("Processing your Excel files and comparing 2024 vs 2025..."):
                final_data = extract_data_from_excel(uploaded_file, c9_file=uploaded_c9)

            if len(final_data) == 0:
                st.warning("No records found for 2024 or 2025 (even after including C9 roster).")
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
            if 'Highest_Score_Year' in final_data.columns:
                year_counts = final_data['Highest_Score_Year'].value_counts(sort=False).reindex([2024, 2025], fill_value=0)
                st.bar_chart(year_counts)
            else:
                st.info("No 'Highest_Score_Year' column to chart (unexpected).")

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
                "'Integración', 'RSU', 'estress', 'Hab. comunicación'. "
                "C9 roster file should include at least DNI and (preferably) Email; "
                "names/period help match name-based modules and set the Year."
            )
    else:
        st.info("👆 Please upload both the Master Excel and the C9 roster Excel to get started.")

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

        **C9 Roster Excel** (authoritative for names + Email; flexible headers supported):
        - Must have **DNI**.
        - Ideally includes **NOMBRES**, **APELLIDO PATERNO**, **APELLIDO MATERNO**, and an **Email/Correo** column.
        - If it includes **Periodo** (e.g., `2025-II`), that helps set the Year; otherwise we infer when possible.
        """)

if __name__ == "__main__":
    main()
