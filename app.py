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

def load_teacher_contract(contract_file):
    """
    Load the teacher contract file and extract ['DNI','Nombre','Apellido(s)'].
    Tolerant with header variations.
    """
    df = pd.read_excel(contract_file, sheet_name=0)
    cols = {c.strip(): c for c in df.columns if isinstance(c, str)}

    dni_col = None
    for k in [
        'N° DE DOCUMENTO DE IDENTIDAD', 'N° DE DOCUMENTO DE IDENTIDAD ',
        'NRO DE DOCUMENTO DE IDENTIDAD', 'NRO DE DOCUMENTO', 'NRO DE DOCUMENTO'
    ]:
        if k in cols:
            dni_col = cols[k]
            break
    if dni_col is None:
        for c in df.columns:
            if isinstance(c, str) and 'DOCUMENTO' in c.upper() and 'IDENTIDAD' in c.upper():
                dni_col = c
                break

    name_col = cols.get('NOMBRES', cols.get('Nombres', None))
    a_pat_col = cols.get('APELLIDO PATERNO', cols.get('Apellido Paterno', None))
    a_mat_col = cols.get('APELLIDO MATERNO', cols.get('Apellido Materno', None))

    if dni_col is None or name_col is None or a_pat_col is None or a_mat_col is None:
        return pd.DataFrame(columns=['DNI', 'Nombre', 'Apellido(s)'])

    out = pd.DataFrame({
        'DNI': df[dni_col].apply(normalize_dni_value),
        'Nombre': df[name_col].astype(str).str.strip(),
        'Apellido(s)': (
            df[a_pat_col].astype(str).str.strip() + ' ' +
            df[a_mat_col].astype(str).str.strip()
        ).str.replace(r'\s+', ' ', regex=True).str.strip()
    })
    out = out.dropna(subset=['DNI'])
    out = out[out['DNI'] != '']
    out = out.drop_duplicates(subset=['DNI'])
    return out[['DNI', 'Nombre', 'Apellido(s)']]

# -----------------------------
# Core processing
# -----------------------------
def extract_data_from_excel(file_path, contract_file=None):
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

    # ---- Teacher Contract: FORCE include all contracted DNIs
    if contract_file is not None:
        contract = load_teacher_contract(contract_file)
        contract['DNI'] = contract['DNI'].apply(normalize_dni_value)
        contract = contract.dropna(subset=['DNI']).drop_duplicates(subset=['DNI'])

        # Right-join ensures every contracted teacher appears at least once
        all_data = pd.merge(
            all_data, contract, on='DNI', how='right', suffixes=('', '_contract')
        )
        # Prefer base names; if missing, take from contract
        all_data['Nombre'] = all_data['Nombre'].fillna(all_data['Nombre_contract'])
        all_data['Apellido(s)'] = all_data['Apellido(s)'].fillna(all_data['Apellido(s)_contract'])
        all_data = all_data.drop(columns=['Nombre_contract', 'Apellido(s)_contract'])

        # Ensure Periodo/Year exist for contract-only rows
        all_data['Periodo'] = all_data['Periodo'].fillna('2025')
        # Make 'Year' numeric and default to 2025 if missing
        all_data['Year'] = pd.to_numeric(all_data['Year'], errors='coerce').fillna(2025).astype(int)

    # ---- Numeric components (14 total)
    numeric_columns = [
        'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
        'Padlet', 'Nearpod', 'Tareas_y_foros',
        'integracion', 'rsu', 'estress', 'hab_comunicacion'
    ]
    for col in numeric_columns:
        all_data[col] = pd.to_numeric(all_data.get(col, 0), errors='coerce').fillna(0)

    # ---- Compute metrics
    all_data['Average'] = all_data[numeric_columns].mean(axis=1).round(2)

    def calculate_percentage(row):
        scores = row[numeric_columns].values
        available = int(np.sum(np.array(scores) > 0))
        return round(available / len(numeric_columns) * 100, 2) if len(numeric_columns) else 0.0

    all_data['Percentage'] = all_data.apply(calculate_percentage, axis=1)
    all_data['Marks_Out_Of_20'] = (all_data['Percentage'] / 5).round(2)

    # ---- Only compare 2024 vs 2025 (INCLUDE zero-value rows + placeholders default to 2025)
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

    # ---- Final column order
    final_columns = [
        'Periodo', 'Highest_Score_Year', 'DNI', 'Nombre', 'Apellido(s)',
        'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
        'Padlet', 'Nearpod', 'Tareas_y_foros',
        'integracion', 'rsu', 'estress', 'hab_comunicacion',
        'Average', 'Marks_Out_Of_20', 'Percentage'
    ]
    # Ensure presence even if some columns are missing
    for c in final_columns:
        if c not in highest.columns:
            highest[c] = 0 if c in numeric_columns + ['Average', 'Marks_Out_Of_20', 'Percentage'] else ''
    final_df = highest[final_columns]
    return final_df

# -----------------------------
# Streamlit App
# -----------------------------
def main():
    st.set_page_config(page_title="📊 UMA Scores (Highest of 2024 vs 2025)", page_icon="📊", layout="wide")

    st.title("📊 UMA Scores — Highest Marks (2024 vs 2025)")
    st.markdown(
        "Upload the **Master** Excel and the **Teacher Contract** Excel. "
        "The output contains **one row per teacher (no duplicates)** — "
        "the row corresponding to the **highest _Marks Out Of 20_ between 2024 and 2025**. "
        "**Teachers with all components = 0 are included.** "
        "_If a contracted teacher is missing from the master entirely, they’re still shown "
        "with `Periodo=2025`, `Year=2025`, and all scores = 0._"
    )

    # File uploaders
    uploaded_file = st.file_uploader("Choose the Master Excel file", type=["xlsx", "xls"], key="master")
    uploaded_contract = st.file_uploader("Choose the Teacher Contract Excel file", type=["xlsx", "xls"], key="contract")

    if uploaded_file is not None and uploaded_contract is not None:
        try:
            with st.spinner("Processing your Excel files and comparing 2024 vs 2025..."):
                final_data = extract_data_from_excel(uploaded_file, contract_file=uploaded_contract)

            if len(final_data) == 0:
                st.warning("No records found for 2024 or 2025 (even zero-score rows and placeholders).")
                return

            st.success("Done! Showing the highest marks per teacher (including zero-score and placeholder teachers).")

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
                "'Integración', 'RSU', 'estress', 'Hab. comunicación'. "
                "Teacher Contract file should include DNI and names."
            )
    else:
        st.info("👆 Please upload both the Master Excel and the Teacher Contract Excel to get started.")

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

        **Teacher Contract**:
        - `N° DE DOCUMENTO DE IDENTIDAD` (DNI), `NOMBRES`, `APELLIDO PATERNO`, `APELLIDO MATERNO`.
        """)

if __name__ == "__main__":
    main()
