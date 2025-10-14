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
    """Extract a 4-digit year (20xx) from values like 2024, '2024-I', '2025-II', etc."""
    if pd.isna(value):
        return pd.NA
    s = str(value)
    m = re.search(r'(20\d{2})', s)
    return int(m.group(1)) if m else pd.NA

def _first_col_like(cols, *candidates, contains=None):
    """Find a column in 'cols' by exact name(s) or substring (case-insensitive)."""
    s2o = {c.strip(): c for c in cols if isinstance(c, str)}
    for c in candidates:
        if c in s2o:
            return s2o[c]
    # fallback: contains substring search
    if contains:
        up = {c.upper(): c for c in cols if isinstance(c, str)}
        for key, raw in up.items():
            ok = True
            for token in contains:
                if token.upper() not in key:
                    ok = False
                    break
            if ok:
                return raw
    return None

def _prep_names(df, name_col, ap_col=None, am_col=None):
    """Return standardized Nombre, Apellido(s) given available columns."""
    out = pd.DataFrame(index=df.index)
    if name_col is not None:
        out['Nombre'] = df[name_col].astype(str).str.strip()
    else:
        out['Nombre'] = pd.NA

    if ap_col is not None and am_col is not None:
        ap = df[ap_col].astype(str).str.strip()
        am = df[am_col].astype(str).str.strip()
        out['Apellido(s)'] = (ap + ' ' + am).str.replace(r'\s+', ' ', regex=True).str.strip()
    else:
        ap_single = _first_col_like(df.columns, 'Apellido(s)')
        out['Apellido(s)'] = df[ap_single].astype(str).str.strip() if ap_single else pd.NA
    return out

def to_numeric_inplace(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

def group_max(df, keys, num_cols):
    """Group by keys and take max over numeric columns only."""
    keep = [c for c in num_cols if c in df.columns]
    if not keep:
        return df[keys].drop_duplicates()
    tmp = df[keys + keep].copy()
    to_numeric_inplace(tmp, keep)
    g = tmp.groupby(keys, as_index=False)[keep].max()
    return g

# -----------------------------
# Load C9 roster (names & email; email not exported)
# -----------------------------
def load_c9_roster(c9_file):
    """
    Load the C9 roster and extract: DNI, Nombre, Apellido(s), Email, Periodo (if present).
    Tolerant to header variations.
    """
    df = pd.read_excel(c9_file, sheet_name=0)
    cols = df.columns

    dni_col = _first_col_like(
        cols,
        'DNI',
        'N° DE DOCUMENTO DE IDENTIDAD',
        'N° DE DOCUMENTO DE IDENTIDAD ',
        'NRO DE DOCUMENTO DE IDENTIDAD',
        'NRO DE DOCUMENTO',
        contains=('DOCUMENTO', 'IDENTIDAD')
    )
    email_col = _first_col_like(
        cols,
        'Dirección de correo',
        'Correo',
        'Correo institucional',
        'CORREO INSTITUCIONAL',
        'Email', 'E-mail', 'MAIL', 'Correo UMA', 'CORREO UMA',
        contains=('CORREO',)
    )
    periodo_col = _first_col_like(cols, 'Periodo', 'PERIODO', 'PERÍODO')

    # Names
    name_col = _first_col_like(cols, 'NOMBRES', 'Nombres')
    ap_col = _first_col_like(cols, 'APELLIDO PATERNO', 'Apellido Paterno')
    am_col = _first_col_like(cols, 'APELLIDO MATERNO', 'Apellido Materno')

    names = _prep_names(df, name_col, ap_col, am_col)
    out = pd.DataFrame({
        'DNI': df[dni_col].apply(normalize_dni_value) if dni_col else np.nan,
        'Email': (df[email_col].astype(str).str.strip() if email_col else pd.NA),
        'Periodo': df[periodo_col] if periodo_col else pd.NA
    })
    out['Nombre'] = names['Nombre']
    out['Apellido(s)'] = names['Apellido(s)']
    out['DNI'] = out['DNI'].astype(str)
    out = out.dropna(subset=['DNI'])
    out = out[out['DNI'] != '']
    out = out.sort_values(by=['Email'], na_position='last')
    out = out.drop_duplicates(subset=['DNI'], keep='first')
    return out[['DNI', 'Nombre', 'Apellido(s)', 'Email', 'Periodo']]

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

    # ---------------- Base from Inducción + nota Inducción ----------------
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
    all_data['DNI'] = all_data['DNI'].apply(normalize_dni_value).astype(str)
    all_data['Year'] = all_data['Periodo'].apply(extract_year)

    # If multiple rows per DNI/Periodo, keep highest induccion grade
    all_data['induccion'] = pd.to_numeric(all_data['induccion'], errors='coerce').fillna(0)
    all_data = all_data.groupby(
        ['DNI', 'Periodo', 'Year', 'Nombre', 'Apellido(s)'],
        as_index=False
    ).agg({'Dirección de correo': 'first', 'induccion': 'max'})

    # ---------------- Module: Bus. biblioteca (by DNI) ----------------
    bus_biblioteca_df['DNI'] = bus_biblioteca_df['DNI'].apply(normalize_dni_value).astype(str)
    bus_biblioteca_df = bus_biblioteca_df.rename(columns={'Promedio': 'bus_biblioteca'})
    bus_biblioteca_df['bus_biblioteca'] = pd.to_numeric(bus_biblioteca_df['bus_biblioteca'], errors='coerce').fillna(0)
    bus_biblioteca_g = group_max(bus_biblioteca_df, ['DNI'], ['bus_biblioteca'])
    all_data = all_data.merge(bus_biblioteca_g, on='DNI', how='left')

    # ---------------- Diseño de sesión (join by names; pre-aggregate) ----------------
    diseno_sesion_df = diseno_sesion_df.rename(columns={'Promedio': 'diseno_sesion'})
    diseno_sesion_df['diseno_sesion'] = pd.to_numeric(diseno_sesion_df['diseno_sesion'], errors='coerce').fillna(0)
    diseno_sesion_g = group_max(diseno_sesion_df, ['Nombre', 'Apellido(s)'], ['diseno_sesion'])
    all_data = all_data.merge(diseno_sesion_g, on=['Nombre', 'Apellido(s)'], how='left')

    # ---------------- Comp. Tec (join by names; pre-aggregate) ----------------
    comp_tec_columns_map = {
        'Cuestionario:Reto: Zoom básico': 'Zoom_basico',
        'Cuestionario:Reto: Zoom Avanzado': 'Zoom_Avanzado',
        'Cuestionario:Reto: Grupos Moodle': 'Grupos_Moodle',
        'Cuestionario:Reto: Rúbrica': 'Rubrica',
        'Cuestionario:Reto: Padlet': 'Padlet',
        'Cuestionario:Reto: Nearpod': 'Nearpod',
        'Cuestionario:Reto: Tareas y foros': 'Tareas_y_foros'
    }
    comp_tec_df = comp_tec_df.rename(columns=comp_tec_columns_map)
    ct_num_cols = [c for c in ['Zoom_basico','Zoom_Avanzado','Grupos_Moodle','Rubrica','Padlet','Nearpod','Tareas_y_foros'] if c in comp_tec_df.columns]
    to_numeric_inplace(comp_tec_df, ct_num_cols)
    comp_tec_g = group_max(comp_tec_df, ['Nombre','Apellido(s)'], ct_num_cols)
    all_data = all_data.merge(comp_tec_g, on=['Nombre','Apellido(s)'], how='left')

    # ---------------- Integración (join by names; pre-aggregate) ----------------
    integracion_df = integracion_df.rename(columns={
        'Tarea:Producto final: Contenido académico, presentación y rúbrica con IA (Real)': 'integracion'
    })
    integracion_df['integracion'] = pd.to_numeric(integracion_df['integracion'], errors='coerce').fillna(0)
    integracion_g = group_max(integracion_df, ['Nombre','Apellido(s)'], ['integracion'])
    all_data = all_data.merge(integracion_g, on=['Nombre','Apellido(s)'], how='left')

    # ---------------- RSU / estress / Hab. comunicación (by DNI) ----------------
    rsu_df['DNI'] = rsu_df['DNI'].apply(normalize_dni_value).astype(str)
    rsu_df = rsu_df.rename(columns={'Tarea: Producto final': 'rsu'})
    rsu_df['rsu'] = pd.to_numeric(rsu_df['rsu'], errors='coerce').fillna(0)
    rsu_g = group_max(rsu_df, ['DNI'], ['rsu'])
    all_data = all_data.merge(rsu_g, on='DNI', how='left')

    estress_df['DNI'] = estress_df['DNI'].apply(normalize_dni_value).astype(str)
    estress_df = estress_df.rename(columns={'Tarea:Producto final': 'estress'})
    estress_df['estress'] = pd.to_numeric(estress_df['estress'], errors='coerce').fillna(0)
    estress_g = group_max(estress_df, ['DNI'], ['estress'])
    all_data = all_data.merge(estress_g, on='DNI', how='left')

    hab_com_df['DNI'] = hab_com_df['DNI'].apply(normalize_dni_value).astype(str)
    hab_com_df = hab_com_df.rename(columns={'Tarea:Producto final': 'hab_comunicacion'})
    hab_com_df['hab_comunicacion'] = pd.to_numeric(hab_com_df['hab_comunicacion'], errors='coerce').fillna(0)
    hab_com_g = group_max(hab_com_df, ['DNI'], ['hab_comunicacion'])
    all_data = all_data.merge(hab_com_g, on='DNI', how='left')

    # ---------------- Bring in C9 roster (names & email; email not exported) ----------------
    if c9_file is not None:
        c9 = load_c9_roster(c9_file)
        c9['DNI'] = c9['DNI'].astype(str)
        c9['Year_roster'] = c9['Periodo'].apply(extract_year)

        # Merge to enrich/override names from C9
        all_data = all_data.merge(c9[['DNI','Nombre','Apellido(s)','Email','Periodo','Year_roster']],
                                  on='DNI', how='right', suffixes=('', '_c9'))

        # Prefer C9 names when available
        all_data['Nombre'] = all_data['Nombre_c9'].fillna(all_data['Nombre'])
        all_data['Apellido(s)'] = all_data['Apellido(s)_c9'].fillna(all_data['Apellido(s)'])

        # If Periodo/Year missing from induction sheets, use C9's
        all_data['Periodo'] = all_data['Periodo'].fillna(all_data['Periodo_c9'])
        all_data['Year'] = all_data['Year'].fillna(all_data['Year_roster'])

        # Clean tmp columns (also drop 'Dirección de correo' to avoid exporting email accidentally)
        drop_cols = [c for c in ['Nombre_c9','Apellido(s)_c9','Periodo_c9','Year_roster','Dirección de correo'] if c in all_data.columns]
        all_data = all_data.drop(columns=drop_cols)

    # ---------------- Numeric components (14 total) ----------------
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

    # ---------------- Compute metrics ----------------
    all_data['Average'] = all_data[numeric_columns].mean(axis=1).round(2)

    def calculate_percentage(row):
        scores = row[numeric_columns].values
        available = int(np.sum(np.array(scores) > 0))
        return round(available / len(numeric_columns) * 100, 2) if len(numeric_columns) else 0.0

    all_data['Percentage'] = all_data.apply(calculate_percentage, axis=1)
    all_data['Marks_Out_Of_20'] = (all_data['Percentage'] / 5).round(2)

    # ---------------- Keep only 2024 vs 2025 rows; infer Year if still missing ----------------
    all_data['Year'] = all_data['Year'].fillna(all_data['Periodo'].apply(extract_year))
    all_data['Year'] = all_data['Year'].fillna(2025)
    all_data = all_data[all_data['Year'].isin([2024, 2025])].copy()
    if all_data.empty:
        return pd.DataFrame()

    # ---------------- DEDUP: one row per DNI with the HIGHEST Marks (2024 vs 2025) ----------------
    all_data['YearPref'] = all_data['Year'].apply(lambda y: 1 if y == 2025 else 0)
    sorted_df = all_data.sort_values(
        by=['Marks_Out_Of_20', 'Average', 'YearPref'],
        ascending=[False, False, False]
    )
    highest = sorted_df.drop_duplicates(subset=['DNI'], keep='first').copy()
    highest['Highest_Score_Year'] = highest['Year']

    # Final column order (NO Email)
    final_columns = [
        'Periodo', 'Highest_Score_Year', 'DNI', 'Nombre', 'Apellido(s)',
        'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
        'Padlet', 'Nearpod', 'Tareas_y_foros',
        'integracion', 'rsu', 'estress', 'hab_comunicacion',
        'Average', 'Marks_Out_Of_20', 'Percentage'
    ]
    for c in final_columns:
        if c not in highest.columns:
            highest[c] = pd.NA
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
        "The output will contain **one row per teacher (unique by DNI)** — "
        "the row with the **highest _Marks Out Of 20_ across 2024 vs 2025**. "
        "Teachers with **zero in all modules** are included."
    )

    uploaded_file = st.file_uploader("Choose the Master Excel file", type=["xlsx", "xls"], key="master")
    uploaded_c9 = st.file_uploader("Choose the C9 roster Excel file (names source)", type=["xlsx", "xls"], key="c9")

    if uploaded_file is not None and uploaded_c9 is not None:
        try:
            with st.spinner("Processing your Excel files and comparing 2024 vs 2025..."):
                final_data = extract_data_from_excel(uploaded_file, c9_file=uploaded_c9)

            if len(final_data) == 0:
                st.warning("No records found for 2024 or 2025.")
                return

            st.success("Done! Showing the highest marks per teacher (unique by DNI).")

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

            # Download (NO Email column)
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
                "C9 roster should include at least **DNI** and names; Periodo helps set Year when missing."
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

        **C9 roster (names)**:
        - Must include `DNI`. Ideally also `NOMBRES`, `APELLIDO PATERNO`, `APELLIDO MATERNO`. `Periodo` helps set Year.
        """)

if __name__ == "__main__":
    main()
