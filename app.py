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

def normalize_email(x):
    """Normalize email: lowercase, strip spaces/punctuation, drop invalids."""
    if pd.isna(x):
        return pd.NA
    s = str(x).strip().lower()
    s = s.replace("mailto:", "")
    # Trim common trailing punctuation/spaces
    s = s.strip(" ,;.")
    # Remove spaces inside by collapsing
    s = re.sub(r"\s+", "", s)
    # Very light sanity check
    return s if ("@" in s and "." in s) else pd.NA

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

def _find_col(df, keywords, prefer_exact=None):
    """
    Find a column in df whose name contains any of the keywords (case-insensitive).
    If prefer_exact is provided, try exact case-insensitive match first.
    """
    cols = [c for c in df.columns if isinstance(c, str)]
    upper = {c.upper(): c for c in cols}
    if prefer_exact:
        for pe in prefer_exact:
            if pe.upper() in upper:
                return upper[pe.upper()]
    for c in cols:
        cu = c.upper()
        if any(k in cu for k in keywords):
            return c
    return None

def _split_name_into_nombre_apellidos(fullname):
    """
    Split a full name into 'Nombre' (first token) and 'Apellido(s)' (rest).
    This is a heuristic fallback if separate name columns are not present.
    """
    if pd.isna(fullname):
        return pd.NA, pd.NA
    s = str(fullname).strip()
    if not s:
        return pd.NA, pd.NA
    parts = [p for p in s.split() if p]
    if len(parts) == 1:
        return parts[0], pd.NA
    # Heuristic: first token as Nombre, the rest as Apellido(s)
    return parts[0], " ".join(parts[1:])

def load_teacher_email_roster(roster_file):
    """
    Load the teacher email roster and extract ['Email','Nombre','Apellido(s)'].
    Tolerant with header variations (Correo/Email, Docente/Profesor/Nombre...).
    """
    df = pd.read_excel(roster_file, sheet_name=0)

    # Find Email column
    email_col = _find_col(
        df,
        keywords=["CORREO", "EMAIL", "E-MAIL", "MAIL"],
        prefer_exact=["Dirección de correo", "Correo", "Email", "E-mail"]
    )

    if email_col is None:
        # Try any column that "looks like" email (by values)
        for c in df.columns:
            if not isinstance(c, str):
                continue
            # If many rows contain '@', assume it's email
            sample = df[c].astype(str).str.contains("@", na=False).mean()
            if sample > 0.3:
                email_col = c
                break

    # Try separate name columns first
    nombres_col = _find_col(df, keywords=["NOMBRES", "NOMBRE"], prefer_exact=["NOMBRES", "Nombre"])
    a_pat_col   = _find_col(df, keywords=["APELLIDO PATERNO", "APELLIDO_PATERNO", "A. PATERNO"])
    a_mat_col   = _find_col(df, keywords=["APELLIDO MATERNO", "APELLIDO_MATERNO", "A. MATERNO"])

    # Fallback to a single full-name style column if needed
    full_name_col = None
    if nombres_col is None and (a_pat_col is None or a_mat_col is None):
        full_name_col = _find_col(
            df,
            keywords=["DOCENTE", "PROFESOR", "NOMBRE COMPLETO", "NOMBRES Y APELLIDOS", "APELLIDOS Y NOMBRES", "NOMBRE DOCENTE", "DOCENTE NOMBRE"]
        )

    # Build output
    if email_col is None:
        # No email column found -> empty roster
        return pd.DataFrame(columns=["Email", "Nombre", "Apellido(s)"])

    out = pd.DataFrame()
    out["Email"] = df[email_col].apply(normalize_email)

    if nombres_col is not None and a_pat_col is not None and a_mat_col is not None:
        out["Nombre"] = df[nombres_col].astype(str).str.strip()
        out["Apellido(s)"] = (
            df[a_pat_col].astype(str).str.strip() + " " +
            df[a_mat_col].astype(str).str.strip()
        ).str.replace(r"\s+", " ", regex=True).str.strip()
    elif full_name_col is not None:
        split = df[full_name_col].apply(_split_name_into_nombre_apellidos)
        out["Nombre"] = split.apply(lambda t: t[0])
        out["Apellido(s)"] = split.apply(lambda t: t[1])
    else:
        # Names unknown; leave blank
        out["Nombre"] = pd.NA
        out["Apellido(s)"] = pd.NA

    out = out.dropna(subset=["Email"])
    out = out[out["Email"] != ""]
    out = out.drop_duplicates(subset=["Email"])
    return out[["Email", "Nombre", "Apellido(s)"]]

# -----------------------------
# Core processing
# -----------------------------
def extract_data_from_excel(file_path, roster_file=None):
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

    # Normalize DNI, Email and derive Year
    all_data['DNI'] = all_data['DNI'].apply(normalize_dni_value)
    all_data['Email'] = all_data['Dirección de correo'].apply(normalize_email)
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

    # ---- Teacher Email Roster (optional but recommended)
    if roster_file is not None:
        roster = load_teacher_email_roster(roster_file)
        roster['Email'] = roster['Email'].apply(normalize_email)
        roster = roster.dropna(subset=['Email'])

        # Keep only rows whose Email is present in roster
        valid_emails = set(roster['Email'].dropna().tolist())
        all_data = all_data[all_data['Email'].isin(valid_emails)]

        # Fill names from roster when missing
        all_data = pd.merge(
            all_data, roster, on='Email', how='left', suffixes=('', '_roster')
        )
        # If original names missing/blank, take from roster
        def _is_blank(x):
            return (pd.isna(x)) or (str(x).strip() == "")
        all_data['Nombre'] = np.where(
            all_data['Nombre'].apply(_is_blank),
            all_data['Nombre_roster'],
            all_data['Nombre']
        )
        all_data['Apellido(s)'] = np.where(
            all_data['Apellido(s)'].apply(_is_blank),
            all_data['Apellido(s)_roster'],
            all_data['Apellido(s)']
        )
        all_data = all_data.drop(columns=['Nombre_roster', 'Apellido(s)_roster'])

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

    # ---- Only compare 2024 vs 2025 and keep rows with any score
    filtered = all_data[all_data['Year'].isin([2024, 2025])].copy()
    filtered = filtered[filtered[numeric_columns].sum(axis=1) > 0]

    if filtered.empty:
        return pd.DataFrame()

    # ---- DEDUP: one row per teacher (by EMAIL) with the HIGHEST Marks_Out_Of_20 across 2024 & 2025
    # Tie-breaker: higher Average, then prefer 2025 over 2024
    filtered['YearPref'] = filtered['Year'].apply(lambda y: 1 if y == 2025 else 0)
    sorted_df = filtered.sort_values(
        by=['Marks_Out_Of_20', 'Average', 'YearPref'],
        ascending=[False, False, False]
    )
    highest = sorted_df.drop_duplicates(subset=['Email'], keep='first').copy()
    highest['Highest_Score_Year'] = highest['Year']

    # ---- Final column order
    final_columns = [
        'Periodo', 'Highest_Score_Year', 'Email', 'DNI', 'Nombre', 'Apellido(s)',
        'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica',
        'Padlet', 'Nearpod', 'Tareas_y_foros',
        'integracion', 'rsu', 'estress', 'hab_comunicacion',
        'Average', 'Marks_Out_Of_20', 'Percentage'
    ]
    # Ensure columns exist even if roster not provided
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
        "Upload the **Master** Excel and the **Teacher Email Roster** Excel (e.g., *C9_25-II_121025.xlsx*). "
        "The output will contain **only one row per teacher (no duplicates)** — "
        "identified by **Email** and chosen as the **highest _Marks Out Of 20_** between 2024 and 2025."
    )

    # File uploaders
    uploaded_file = st.file_uploader("Choose the Master Excel file", type=["xlsx", "xls"], key="master")
    uploaded_roster = st.file_uploader("Choose the Teacher Email Roster Excel file", type=["xlsx", "xls"], key="roster")

    if uploaded_file is not None and uploaded_roster is not None:
        try:
            with st.spinner("Processing your Excel files and comparing 2024 vs 2025 (by Email)..."):
                final_data = extract_data_from_excel(uploaded_file, roster_file=uploaded_roster)

            if len(final_data) == 0:
                st.warning("No records with scores found for 2024 or 2025 (after filtering by the email roster).")
                return

            st.success("Done! Showing only the highest marks per teacher (no duplicates, identified by Email).")

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
            output_filename = f"Highest_Marks_2024_vs_2025_by_Email_{timestamp}.xlsx"
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_data.to_excel(writer, index=False, sheet_name='Highest Marks (Unique by Email)')
            output.seek(0)

            st.download_button(
                label="📥 Download (Unique by Email, Highest of 2024/2025)",
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
                "Teacher Email Roster should include at least one column with emails "
                "(e.g., 'Dirección de correo', 'Correo', 'Email')."
            )
    else:
        st.info("👆 Please upload both the Master Excel and the Teacher Email Roster Excel to get started.")

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

        **Teacher Email Roster (e.g., C9_25-II_121025.xlsx)**:
        - One email column (like `Dirección de correo`, `Correo`, `Email`, etc.).
        - Optional name columns: either separate (`Nombres`, `Apellido Paterno`, `Apellido Materno`) **or** a full-name column (e.g., `Docente`, `Profesor`).
        """)

if __name__ == "__main__":
    main()
