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

def _pick_first_present(df, candidates, contains_all=None, contains_any=None):
    """
    Find a column in df:
      - First try exact candidates (list of strings)
      - Else try 'contains_all' (list of substrings that must all be in col name)
      - Else try 'contains_any' (list of substrings where any match)
    Returns the column name or None.
    """
    cols = [c for c in df.columns if isinstance(c, str)]
    # exact matches (case-sensitive first, then case-insensitive)
    for k in candidates or []:
        if k in cols:
            return k
    if candidates:
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
# C9 loader (replaces contract)
# -----------------------------
def load_c9_file(c9_file):
    """
    Load the C9 roster and extract:
      ['DNI','Nombre','Apellido(s)','Email','Periodo','Year']
    Tolerant with header variations.
    """
    df = pd.read_excel(c9_file, sheet_name=0)

    # --- locate columns
    dni_col = _pick_first_present(
        df,
        candidates=[
            'DNI', 'dni', 'N° DE DOCUMENTO DE IDENTIDAD', 'N° DE DOCUMENTO DE IDENTIDAD ',
            'NRO DE DOCUMENTO DE IDENTIDAD', 'NRO DE DOCUMENTO', 'NUMERO DE DOCUMENTO'
        ],
        contains_all=['DOCUMENTO', 'IDENTIDAD']
    )

    # Names can be (NOMBRES + APELLIDO PATERNO + APELLIDO MATERNO) or already split
    nombres_col = _pick_first_present(df, candidates=['NOMBRES', 'Nombres', 'Nombre'])
    a_pat_col = _pick_first_present(df, candidates=['APELLIDO PATERNO', 'Apellido Paterno', 'Ap. Paterno', 'Ap Paterno', 'Apellidos'])
    a_mat_col = _pick_first_present(df, candidates=['APELLIDO MATERNO', 'Apellido Materno', 'Ap. Materno', 'Ap Materno'])

    # Email can vary a lot; pick the first reasonable email-like column
    email_col = _pick_first_present(
        df,
        candidates=['Dirección de correo', 'Correo', 'Correo institucional', 'Correo Institucional',
                    'Correo electrónico', 'Correo Electronico', 'EMAIL', 'Email', 'E-mail'],
        contains_any=['CORREO', 'EMAIL']
    )

    # Period may appear, else we infer from filename if possible
    periodo_col = _pick_first_present(
        df,
        candidates=['Periodo', 'PERIODO', 'Periodo Académico', 'PERIODO ACADÉMICO', 'PERIODO ACADEMICO'],
        contains_any=['PERIODO', 'PERÍODO', 'ACADEM']
    )

    # --- build normalized output
    out = pd.DataFrame()
    if dni_col is not None:
        out['DNI'] = df[dni_col].apply(normalize_dni_value)
    else:
        out['DNI'] = np.nan

    # names
    if nombres_col is not None:
        out['Nombre'] = df[nombres_col].astype(str).map(clean_spaces)
    else:
        out['Nombre'] = pd.NA

    if a_pat_col is not None and a_mat_col is not None:
        ap = df[a_pat_col].astype(str).map(clean_spaces)
        am = df[a_mat_col].astype(str).map(clean_spaces)
        out['Apellido(s)'] = (ap.fillna('') + ' ' + am.fillna('')).str.strip()
        out['Apellido(s)'] = out['Apellido(s)'].str.replace(r'\s+', ' ', regex=True)
    elif a_pat_col is not None and a_mat_col is None:
        out['Apellido(s)'] = df[a_pat_col].astype(str).map(clean_spaces)
    else:
        # If there's a single 'Apellidos' column we may have captured it in a_pat_col already.
        if 'Apellido(s)' not in out:
            out['Apellido(s)'] = pd.NA

    # email
    if email_col is not None:
        out['Email'] = df[email_col].astype(str).map(clean_spaces)
    else:
        out['Email'] = pd.NA

    # periodo/year
    if periodo_col is not None:
        out['Periodo'] = df[periodo_col]
    else:
        # try inferring period/year from file name (e.g., C9_25-II_*.xlsx -> 2025-II)
        fname = getattr(c9_file, 'name', '') or ''
        yr = None
        m4 = re.search(r'(20\d{2})', fname)
        if m4:
            yr = int(m4.group(1))
        else:
            m2 = re.search(r'(^|[^0-9])(2[45])([^0-9]|$)', fname)  # 24 or 25
            if m2:
                yr = 2000 + int(m2.group(2))
        out['Periodo'] = f'{yr}' if yr else pd.NA

    out['Periodo'] = out['Periodo'].astype(str).str.strip()
    out['Year'] = out['Periodo'].apply(extract_year)

    # finalize
    out['DNI'] = out['DNI'].apply(normalize_dni_value)
    out = out.dropna(subset=['DNI'])
    out = out[out['DNI'] != '']
    out = out.drop_duplicates(subset=['DNI'])

    # keep standard columns
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

    # ---- Prepare base from Inducción + nota Inducción (unified columns)
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

    # Normalize identifiers & year
    all_data['DNI'] = all_data['DNI'].apply(normalize_dni_value)
    all_data['Nombre'] = all_data['Nombre'].astype(str).map(clean_spaces)
    all_data['Apellido(s)'] = all_data['Apellido(s)'].astype(str).map(clean_spaces)
    all_data['Year'] = all_data['Periodo'].apply(extract_year)

    # ---- If C9 roster provided, use it as the teacher universe and email source
    c9 = None
    if c9_file is not None:
        c9 = load_c9_file(c9_file)
        # ensure Email column is present
        if 'Email' not in c9.columns:
            c9['Email'] = pd.NA

        # Append any C9-only teachers (so zero-mark teachers also appear)
        known = set(all_data['DNI'].dropna())
        c9_only = c9[~c9['DNI'].isin(known)].copy()

        # For C9-only rows, create minimal records so they flow through
        if not c9_only.empty:
            extra = pd.DataFrame({
                'Periodo': c9_only['Periodo'],
                'DNI': c9_only['DNI'],
                'Nombre': c9_only['Nombre'],
                'Apellido(s)': c9_only['Apellido(s)'],
                'Dirección de correo': c9_only['Email'],  # use roster email
                'induccion': np.nan
            })
            all_data = pd.concat([all_data, extra], ignore_index=True)

        # Keep only DNIs present in C9 (mirrors your previous "contract" filter)
        all_data = all_data[all_data['DNI'].isin(c9['DNI'])].copy()

        # Bring in roster names/emails when missing
        all_data = pd.merge(
            all_data,
            c9[['DNI', 'Nombre', 'Apellido(s)', 'Email', 'Periodo', 'Year']],
            on='DNI',
            how='left',
            suffixes=('', '_c9')
        )

        # Prefer existing names if present; otherwise fill from C9
        all_data['Nombre'] = all_data['Nombre'].fillna(all_data['Nombre_c9'])
        all_data['Apellido(s)'] = all_data['Apellido(s)'].fillna(all_data['Apellido(s)_c9'])

        # Prefer explicit Periodo/Year from Inducción/nota; if missing, use C9 period/year
        all_data['Periodo'] = all_data['Periodo'].fillna(all_data['Periodo_c9'])
        all_data['Year'] = all_data['Year'].fillna(all_data['Year_c9'])

        # Keep Email (from C9); don't drop the original correo column — we can use it as fallback later
        all_data = all_data.drop(columns=[c for c in ['Nombre_c9', 'Apellido(s)_c9', 'Periodo_c9', 'Year_c9'] if c in all_data])

    # ---- Bus. biblioteca (by DNI)
    bus_biblioteca_df['DNI'] = bus_biblioteca_df['DNI'].apply(normalize_dni_value)
    bus_biblioteca_df = bus_biblioteca_df.rename(columns={'Promedio': 'bus_biblioteca'})
    all_data = pd.merge(
        all_data, bus_biblioteca_df[['DNI', 'bus_biblioteca']],
        on='DNI', how='left'
    )

    # ---- RSU, estress, Hab. comunicación (by DNI)
    for src_df, col_src, col_dst in [
        (rsu_df, 'Tarea: Producto final', 'rsu'),
        (estress_df, 'Tarea:Producto final', 'estress'),
        (hab_com_df, 'Tarea:Producto final', 'hab_comunicacion')
    ]:
        src_df['DNI'] = src_df['DNI'].apply(normalize_dni_value)
        tmp = src_df.rename(columns={col_src: col_dst})
        all_data = pd.merge(all_data, tmp[['DNI', col_dst]], on='DNI', how='left')

    # ---- Diseño de sesión / Integración / Comp. Tec (join by names)
    # Clean names in source sheets to improve matching
    for df_name in [diseno_sesion_df, integracion_df, comp_tec_df]:
        if 'Nombre' in df_name.columns:
            df_name['Nombre'] = df_name['Nombre'].astype(str).map(clean_spaces)
        if 'Apellido(s)' in df_name.columns:
            df_name['Apellido(s)'] = df_name['Apellido(s)'].astype(str).map(clean_spaces)

    # Diseño de sesión
    diseno_sesion_df = diseno_sesion_df.rename(columns={'Promedio': 'diseno_sesion'})
    all_data = pd.merge(
        all_data, diseno_sesion_df[['Nombre', 'Apellido(s)', 'diseno_sesion']],
        on=['Nombre', 'Apellido(s)'], how='left'
    )

    # Integración
    integracion_df = integracion_df.rename(columns={
        'Tarea:Producto final: Contenido académico, presentación y rúbrica con IA (Real)': 'integracion'
    })
    all_data = pd.merge(
        all_data, integracion_df[['Nombre', 'Apellido(s)', 'integracion']],
        on=['Nombre', 'Apellido(s)'], how='left'
    )

    # Comp. Tec
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

    # ---- Numeric components (14 total)
    numeric_columns = [
        'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
        'Padlet', 'Nearpod', 'Tareas_y_foros',
        'integracion', 'rsu', 'estress', 'hab_comunicacion'
    ]
    for col in numeric_columns:
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

    # ---- DEDUP: one row per teacher (by DNI) with the HIGHEST Marks_Out_Of_20 across 2024 & 2025
    # Tie-breaker: higher Average, then prefer 2025 over 2024
    filtered['YearPref'] = filtered['Year'].apply(lambda y: 1 if y == 2025 else 0)
    sorted_df = filtered.sort_values(
        by=['Marks_Out_Of_20', 'Average', 'YearPref'],
        ascending=[False, False, False]
    )
    highest = sorted_df.drop_duplicates(subset=['DNI'], keep='first').copy()
    highest['Highest_Score_Year'] = highest['Year']

    # ---- Email column (from C9 if present, else fallback to 'Dirección de correo')
    if 'Email' not in highest.columns:
        highest['Email'] = pd.NA
    if 'Dirección de correo' in highest.columns:
        highest['Email'] = highest['Email'].fillna(highest['Dirección de correo'])

    # ---- Final column order (now with Email)
    final_columns = [
        'Periodo', 'Highest_Score_Year', 'DNI', 'Nombre', 'Apellido(s)', 'Email',
        'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
        'Padlet', 'Nearpod', 'Tareas_y_foros',
        'integracion', 'rsu', 'estress', 'hab_comunicacion',
        'Average', 'Marks_Out_Of_20', 'Percentage'
    ]
    # Keep columns that exist (in case some module sheets were entirely missing)
    final_columns = [c for c in final_columns if c in highest.columns]
    final_df = highest[final_columns]
    return final_df

# -----------------------------
# Streamlit App
# -----------------------------
def main():
    st.set_page_config(page_title="📊 UMA Scores (Highest of 2024 vs 2025)", page_icon="📊", layout="wide")

    st.title("📊 UMA Scores — Highest Marks (2024 vs 2025)")
    st.markdown(
        "Upload the **Master** Excel and the **C9 roster** Excel. "
        "The output will contain **only one row per teacher (no duplicates)** — "
        "the row corresponding to the **highest _Marks Out Of 20_ between 2024 and 2025**. "
        "Teachers listed in C9 are included **even if all module scores are 0**."
    )

    # File uploaders
    uploaded_file = st.file_uploader("Choose the Master Excel file", type=["xlsx", "xls"], key="master")
    uploaded_c9 = st.file_uploader("Choose the C9 roster Excel file (for Email & roster)", type=["xlsx", "xls"], key="c9")

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
                year_counts = final_data['Highest_Score_Year'].value_counts().sort_index()
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

        **C9 Roster Excel** (used for roster + Email; flexible headers supported):
        - Must have **DNI**.
        - Ideally includes **NOMBRES**, **APELLIDO PATERNO**, **APELLIDO MATERNO** (or equivalent), and an **Email/Correo** column.
        - If it includes **Periodo** (e.g., `2025-II`), that helps set the Year; otherwise we'll infer when possible.
        """)

if __name__ == "__main__":
    main()
