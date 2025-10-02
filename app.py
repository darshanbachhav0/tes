# app.py
import pandas as pd
import numpy as np
import streamlit as st
import io
import re
from datetime import datetime

# -----------------------------
# Light, fast helpers
# -----------------------------
def normalize_dni_value(x):
    """Normalize DNI to digits only (e.g., '12345678.0' -> '12345678')."""
    if pd.isna(x):
        return np.nan
    s = str(x).strip()
    if re.match(r'^\d+\.0$', s):
        s = s[:-2]
    s = re.sub(r'\D', '', s)
    return s if s else np.nan

def clean_series(s):
    return s.astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()

def load_teacher_contract_from_bytes(contract_bytes):
    """Read minimal columns from the teacher contract, robust to header variants."""
    xf = pd.ExcelFile(io.BytesIO(contract_bytes))
    df = xf.parse(0)  # first sheet
    # Build a case-insensitive lookup over raw columns
    raw_cols = {str(c).strip(): c for c in df.columns}
    def find_col(options):
        for opt in options:
            if opt in raw_cols:
                return raw_cols[opt]
        # fallback: fuzzy search
        for c in df.columns:
            s = str(c).upper()
            if 'DOCUMENTO' in s and 'IDENTIDAD' in s:
                return c
        return None

    dni_col = find_col([
        'N° DE DOCUMENTO DE IDENTIDAD', 'N° DE DOCUMENTO DE IDENTIDAD ',
        'NRO DE DOCUMENTO DE IDENTIDAD', 'NRO DE DOCUMENTO', 'NRO DE DOCUMENTO'
    ])
    name_col   = raw_cols.get('NOMBRES', raw_cols.get('Nombres', None))
    a_pat_col  = raw_cols.get('APELLIDO PATERNO', raw_cols.get('Apellido Paterno', None))
    a_mat_col  = raw_cols.get('APELLIDO MATERNO', raw_cols.get('Apellido Materno', None))

    if not all([dni_col, name_col, a_pat_col, a_mat_col]):
        return pd.DataFrame(columns=['DNI', 'Nombre', 'Apellido(s)'])

    out = pd.DataFrame({
        'DNI': df[dni_col].apply(normalize_dni_value),
        'Nombre': clean_series(df[name_col]),
        'Apellido(s)': clean_series(df[a_pat_col]) + ' ' + clean_series(df[a_mat_col])
    })
    out = out.dropna(subset=['DNI'])
    out = out[out['DNI'] != ''].drop_duplicates(subset=['DNI'])
    return out[['DNI', 'Nombre', 'Apellido(s)']]

# -----------------------------
# Core processing (vectorized & trimmed)
# -----------------------------
def extract_data_from_excel_bytes(master_bytes, contract_bytes=None):
    xf = pd.ExcelFile(io.BytesIO(master_bytes))

    # 1) Base from Inducción + nota Inducción
    induccion = xf.parse('Inducción', usecols=['Periodo', 'DNI', 'Nombre', 'Apellido(s)', 'Dirección de correo', 'Calificación'])
    induccion = induccion.rename(columns={'Calificación': 'induccion'})
    nota_ind = xf.parse('nota Inducción', usecols=['PERIODO', 'DNI', 'Nombre', 'Apellido(s)', 'Dirección de correo', 'Total del curso (Real)'])
    nota_ind = nota_ind.rename(columns={'PERIODO': 'Periodo', 'Total del curso (Real)': 'induccion'})
    all_data = pd.concat([nota_ind, induccion], ignore_index=True)

    # Normalize IDs & names
    all_data['DNI'] = all_data['DNI'].apply(normalize_dni_value)
    all_data['Nombre'] = clean_series(all_data['Nombre'])
    all_data['Apellido(s)'] = clean_series(all_data['Apellido(s)'])

    # Vectorized Year from Periodo (handles '2024', '2024-I', etc.)
    all_data['Year'] = (
        all_data['Periodo'].astype(str)
        .str.extract(r'(20\d{2})', expand=False)
        .astype('Int64')
    )

    # 2) Optionally restrict to DNIs in Teacher Contract early (shrinks work)
    contract = None
    if contract_bytes is not None:
        contract = load_teacher_contract_from_bytes(contract_bytes)
        if not contract.empty:
            dnis = set(contract['DNI'])
            all_data = all_data[all_data['DNI'].isin(dnis)]

    # If nothing left, stop early
    if all_data.empty:
        return pd.DataFrame()

    # Convenience: only keep periods we care about asap
    all_data = all_data[all_data['Year'].isin([2024, 2025])]

    # 3) Bring in the rest, reading only necessary cols and trimming to relevant DNIs where possible
    needed_dnis = set(all_data['DNI'].dropna())

    # Bus. biblioteca (by DNI)
    bus_bib = xf.parse('Bus. biblioteca', usecols=['DNI', 'Promedio'])
    bus_bib['DNI'] = bus_bib['DNI'].apply(normalize_dni_value)
    if needed_dnis:
        bus_bib = bus_bib[bus_bib['DNI'].isin(needed_dnis)]
    bus_bib = bus_bib.rename(columns={'Promedio': 'bus_biblioteca'})
    all_data = all_data.merge(bus_bib, on='DNI', how='left')

    # Diseño de sesión (by names)
    dis_ses = xf.parse('Diseño de sesión', usecols=['Nombre', 'Apellido(s)', 'Promedio'])
    dis_ses = dis_ses.rename(columns={'Promedio': 'diseno_sesion'})
    dis_ses['Nombre'] = clean_series(dis_ses['Nombre'])
    dis_ses['Apellido(s)'] = clean_series(dis_ses['Apellido(s)'])
    all_data = all_data.merge(dis_ses, on=['Nombre', 'Apellido(s)'], how='left')

    # Comp. Tec (by names)
    comp = xf.parse('Comp. Tec')
    comp_map = {
        'Cuestionario:Reto: Zoom básico': 'Zoom_basico',
        'Cuestionario:Reto: Zoom Avanzado': 'Zoom_Avanzado',
        'Cuestionario:Reto: Grupos Moodle': 'Grupos_Moodle',
        'Cuestionario:Reto: Rúbrica': 'Rubrica',
        'Cuestionario:Reto: Padlet': 'Padlet',
        'Cuestionario:Reto: Nearpod': 'Nearpod',
        'Cuestionario:Reto: Tareas y foros': 'Tareas_y_foros'
    }
    keep_cols = ['Nombre', 'Apellido(s)'] + [k for k in comp_map.keys() if k in comp.columns]
    comp = comp[keep_cols].rename(columns=comp_map)
    comp['Nombre'] = clean_series(comp['Nombre'])
    comp['Apellido(s)'] = clean_series(comp['Apellido(s)'])
    all_data = all_data.merge(comp, on=['Nombre', 'Apellido(s)'], how='left')

    # Integración (by names)
    integ = xf.parse('Integración')
    integ_col = 'Tarea:Producto final: Contenido académico, presentación y rúbrica con IA (Real)'
    if integ_col in integ.columns:
        integ = integ[['Nombre', 'Apellido(s)', integ_col]].rename(columns={integ_col: 'integracion'})
        integ['Nombre'] = clean_series(integ['Nombre'])
        integ['Apellido(s)'] = clean_series(integ['Apellido(s)'])
        all_data = all_data.merge(integ, on=['Nombre', 'Apellido(s)'], how='left')
    else:
        all_data['integracion'] = np.nan

    # RSU (by DNI)
    rsu = xf.parse('RSU', usecols=['DNI', 'Tarea: Producto final'])
    rsu['DNI'] = rsu['DNI'].apply(normalize_dni_value)
    if needed_dnis:
        rsu = rsu[rsu['DNI'].isin(needed_dnis)]
    rsu = rsu.rename(columns={'Tarea: Producto final': 'rsu'})
    all_data = all_data.merge(rsu, on='DNI', how='left')

    # estress (by DNI)
    est = xf.parse('estress', usecols=['DNI', 'Tarea:Producto final'])
    est['DNI'] = est['DNI'].apply(normalize_dni_value)
    if needed_dnis:
        est = est[est['DNI'].isin(needed_dnis)]
    est = est.rename(columns={'Tarea:Producto final': 'estress'})
    all_data = all_data.merge(est, on='DNI', how='left')

    # Hab. comunicación (by DNI)
    hab = xf.parse('Hab. comunicación', usecols=['DNI', 'Tarea:Producto final'])
    hab['DNI'] = hab['DNI'].apply(normalize_dni_value)
    if needed_dnis:
        hab = hab[hab['DNI'].isin(needed_dnis)]
    hab = hab.rename(columns={'Tarea:Producto final': 'hab_comunicacion'})
    all_data = all_data.merge(hab, on='DNI', how='left')

    # 4) Fill names from contract (only where missing)
    if contract is not None and not contract.empty:
        all_data = all_data.merge(contract, on='DNI', how='left', suffixes=('', '_contract'))
        all_data['Nombre'] = all_data['Nombre'].fillna(all_data['Nombre_contract'])
        all_data['Apellido(s)'] = all_data['Apellido(s)'].fillna(all_data['Apellido(s)_contract'])
        all_data = all_data.drop(columns=['Nombre_contract', 'Apellido(s)_contract'], errors='ignore')

    # 5) Coerce numeric components (vectorized)
    numeric_cols = [
        'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
        'Padlet', 'Nearpod', 'Tareas_y_foros',
        'integracion', 'rsu', 'estress', 'hab_comunicacion'
    ]
    for c in numeric_cols:
        if c not in all_data.columns:
            all_data[c] = np.nan
    all_data[numeric_cols] = all_data[numeric_cols].apply(pd.to_numeric, errors='coerce').fillna(0)

    # Keep only rows with any non-zero score
    nonzero_mask = (all_data[numeric_cols].sum(axis=1) > 0)
    all_data = all_data[nonzero_mask]
    if all_data.empty:
        return pd.DataFrame()

    # Metrics (fully vectorized)
    all_data['Average'] = all_data[numeric_cols].mean(axis=1).round(2)
    non_zero_count = (all_data[numeric_cols] > 0).sum(axis=1)
    all_data['Percentage'] = ((non_zero_count / len(numeric_cols)) * 100).round(2)
    all_data['Marks_Out_Of_20'] = (all_data['Percentage'] / 5).round(2)

    # 6) Dedup to ONE row per teacher (best of 2024 vs 2025)
    # Prefer: higher Marks_Out_Of_20, then higher Average, then 2025
    yearpref = (all_data['Year'] == 2025).astype(int)
    all_data['_YearPref'] = yearpref
    all_data = all_data.sort_values(by=['Marks_Out_Of_20', 'Average', '_YearPref'], ascending=[False, False, False])
    highest = all_data.drop_duplicates(subset=['DNI'], keep='first').copy()
    highest['Highest_Score_Year'] = highest['Year']

    final_cols = [
        'Periodo', 'Highest_Score_Year', 'DNI', 'Nombre', 'Apellido(s)',
        'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
        'Padlet', 'Nearpod', 'Tareas_y_foros',
        'integracion', 'rsu', 'estress', 'hab_comunicacion',
        'Average', 'Marks_Out_Of_20', 'Percentage'
    ]
    # Some sheets may miss certain columns; guard selection
    final_cols = [c for c in final_cols if c in highest.columns]
    return highest[final_cols].reset_index(drop=True)

# Optional caching (speeds re-runs with the same files during a session)
@st.cache_data(show_spinner=False, ttl=3600)
def process_cached(master_bytes, contract_bytes):
    return extract_data_from_excel_bytes(master_bytes, contract_bytes)

# -----------------------------
# Streamlit App (lean UI)
# -----------------------------
def main():
    st.set_page_config(page_title="📊 UMA Scores (Highest of 2024 vs 2025)", page_icon="📊", layout="wide")
    st.title("📊 UMA Scores — Highest Marks (2024 vs 2025)")
    st.caption("Optimized for fast processing on Render (single-pass Excel parsing, trimmed columns, vectorized ops).")

    master_file = st.file_uploader("Master Excel", type=["xlsx", "xls"])
    contract_file = st.file_uploader("Teacher Contract Excel", type=["xlsx", "xls"])

    if master_file and contract_file:
        try:
            master_bytes = master_file.getvalue()
            contract_bytes = contract_file.getvalue()

            with st.spinner("Processing…"):
                final_df = process_cached(master_bytes, contract_bytes)

            if final_df.empty:
                st.warning("No records with non-zero scores for 2024/2025 after filtering.")
                return

            st.success(f"Done. Unique teachers: {len(final_df)}")
            st.dataframe(final_df.head(30), use_container_width=True)

            # Lightweight metrics (no charts)
            col1, col2, col3 = st.columns(3)
            col1.metric("Teachers", len(final_df))
            col2.metric("Avg Marks/20", f"{final_df['Marks_Out_Of_20'].mean():.2f}")
            col3.metric("Avg %", f"{final_df['Percentage'].mean():.2f}%")

            # Download
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            fname = f"Highest_Marks_2024_vs_2025_{ts}.xlsx"
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as w:
                final_df.to_excel(w, index=False, sheet_name='Highest Marks (Unique)')
            buf.seek(0)
            st.download_button("📥 Download", buf,
                               file_name=fname,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Processing error: {e}")
            st.info("Check sheet names/columns and try again.")
    else:
        st.info("Upload both files to begin.")

if __name__ == "__main__":
    main()
