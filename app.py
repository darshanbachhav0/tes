import pandas as pd
import numpy as np
import streamlit as st
import io
from datetime import datetime

# ---------------------------
# Helpers for contract logic
# ---------------------------
_CONTRACT_SUBSTRINGS = ("contrat", "docent", "teacher")  # matches: contrato, contratado, docente, teacher, etc.

def _detect_contract_columns(df: pd.DataFrame):
    cols = [c for c in df.columns if any(s in c.lower() for s in _CONTRACT_SUBSTRINGS)]
    return cols

def _contract_flag_series(df: pd.DataFrame) -> pd.Series:
    """
    Returns a boolean Series that is True if *any* teacher-contract column in the row
    is truthy/non-empty/non-zero. If no contract-like columns exist, returns all False.
    """
    cols = _detect_contract_columns(df)
    if not cols:
        return pd.Series(False, index=df.index)

    # Convert to strings where appropriate and check for non-empty / truthy values
    sub = df[cols].copy()

    # Normalize values: treat "", "0", "no", "false", NaN as False; anything else as True
    def _to_bool(v):
        if pd.isna(v):
            return False
        if isinstance(v, (int, float)):
            return v != 0
        s = str(v).strip().lower()
        return s not in ("", "0", "no", "false", "nan", "none")

    return sub.applymap(_to_bool).any(axis=1)


def extract_data_from_excel(file_path):
    # Read all sheets
    induction_df = pd.read_excel(file_path, sheet_name='Inducción')
    nota_induccion_df = pd.read_excel(file_path, sheet_name='nota Inducción')
    bus_biblioteca_df = pd.read_excel(file_path, sheet_name='Bus. biblioteca')
    diseno_sesion_df = pd.read_excel(file_path, sheet_name='Diseño de sesión')
    comp_tec_df = pd.read_excel(file_path, sheet_name='Comp. Tec')

    # ---------------------------
    # Prepare induction sources
    # ---------------------------
    # nota_induccion
    nota_induccion_clean = nota_induccion_df[['PERIODO', 'DNI', 'Nombre', 'Apellido(s)', 'Dirección de correo', 'Total del curso (Real)']].copy()
    nota_induccion_clean = nota_induccion_clean.rename(columns={
        'PERIODO': 'Periodo',
        'Total del curso (Real)': 'induccion'
    })
    # Add contract flag from original nota_induccion_df
    nota_induccion_clean['contract_flag'] = _contract_flag_series(nota_induccion_df)

    # inducción
    induction_clean = induction_df[['Periodo', 'DNI', 'Nombre', 'Apellido(s)', 'Dirección de correo', 'Calificación']].copy()
    induction_clean = induction_clean.rename(columns={'Calificación': 'induccion'})
    # Add contract flag from original induction_df
    induction_clean['contract_flag'] = _contract_flag_series(induction_df)

    # Combine both induction-based datasets
    all_data = pd.concat([nota_induccion_clean, induction_clean], ignore_index=True)

    # ---------------------------
    # Merge with other sheets
    # ---------------------------
    # Bus. biblioteca
    bus_biblioteca_df = bus_biblioteca_df.rename(columns={'Promedio': 'bus_biblioteca'})
    all_data = pd.merge(
        all_data,
        bus_biblioteca_df[['DNI', 'bus_biblioteca']],
        on='DNI',
        how='left'
    )

    # Diseño de sesión (by name)
    diseno_sesion_df = diseno_sesion_df.rename(columns={'Promedio': 'diseno_sesion'})
    all_data = pd.merge(
        all_data,
        diseno_sesion_df[['Nombre', 'Apellido(s)', 'diseno_sesion']],
        on=['Nombre', 'Apellido(s)'],
        how='left'
    )

    # Comp. Tec (by name)
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
    comp_keep = ['Nombre', 'Apellido(s)', 'Zoom_basico', 'Zoom_Avanzado',
                 'Grupos_Moodle', 'Rubrica', 'Padlet', 'Nearpod', 'Tareas_y_foros']
    comp_keep = [c for c in comp_keep if c in comp_tec_df.columns]
    all_data = pd.merge(
        all_data,
        comp_tec_df[comp_keep],
        on=['Nombre', 'Apellido(s)'],
        how='left'
    )

    # ---------------------------
    # Scoring
    # ---------------------------
    numeric_columns = [
        'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle',
        'Rubrica', 'Padlet', 'Nearpod', 'Tareas_y_foros'
    ]
    # Ensure all expected numeric columns exist
    for col in numeric_columns:
        if col not in all_data.columns:
            all_data[col] = 0

    # Replace blanks with 0 and coerce to numeric
    for col in numeric_columns:
        all_data[col] = all_data[col].replace('', 0)
        all_data[col] = pd.to_numeric(all_data[col], errors='coerce').fillna(0)

    # Average across 10 components (blanks already 0)
    all_data['Average'] = all_data[numeric_columns].mean(axis=1).round(2)

    # Percentage based on count of non-zero components
    def calculate_percentage(row):
        scores = row[numeric_columns].values
        available_components = int((scores > 0).sum())
        if available_components == 0:
            return 0.0
        return round((available_components / 10) * 100, 2)

    all_data['Percentage'] = all_data.apply(calculate_percentage, axis=1)

    # Marks out of 20
    all_data['Marks_Out_Of_20'] = (all_data['Percentage'] / 5).round(2)

    # ---------------------------
    # Identity + inclusion rule
    # ---------------------------
    all_data['Person_ID'] = (
        all_data['DNI'].astype(str) + '_' +
        all_data['Nombre'].astype(str) + '_' +
        all_data['Apellido(s)'].astype(str)
    )

    # NEW: Include rows if (Average > 0) OR (contract_flag == True)
    # (Previously we dropped Average==0 entirely.)
    all_data = all_data[(all_data['Average'] > 0) | (all_data['contract_flag'])].copy()

    # If still empty, return empty df
    if all_data.empty:
        return pd.DataFrame()

    # Normalize/assist sorting by period (extract numeric if possible)
    all_data['_Periodo_num'] = pd.to_numeric(all_data['Periodo'], errors='coerce').fillna(-1)

    # Choose one row per Person_ID:
    # 1) Highest Average
    # 2) If tie, prefer rows with contract_flag == True
    # 3) If still tie, prefer latest period (_Periodo_num largest)
    all_data.sort_values(
        by=['Person_ID', 'Average', 'contract_flag', '_Periodo_num'],
        ascending=[True, False, False, False],
        inplace=True
    )
    highest_scores = all_data.drop_duplicates(subset=['Person_ID'], keep='first').copy()

    # Highest score period (keep original value)
    highest_scores['Highest_Score_Period'] = highest_scores['Periodo']

    # Final column order (unchanged)
    final_columns = [
        'Periodo', 'DNI', 'Nombre', 'Apellido(s)', 'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica', 'Padlet', 'Nearpod', 'Tareas_y_foros',
        'Average', 'Marks_Out_Of_20', 'Percentage', 'Highest_Score_Period'
    ]
    final_columns = [c for c in final_columns if c in highest_scores.columns]
    final_df = highest_scores[final_columns].copy()

    return final_df


def main():
    st.set_page_config(page_title="Excel Data Processor", page_icon="📊", layout="wide")

    st.title("📊 Excel Data Processor")
    st.markdown("Upload your Excel file to process and combine data from multiple sheets.")
    st.info("This tool compares scores across periods and shows the row per professor with the highest average. "
            "Professors marked in teacher-contract columns are included even if all marks are 0.")

    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            with st.spinner("Processing your Excel file and comparing periods..."):
                final_data = extract_data_from_excel(uploaded_file)

            if len(final_data) == 0:
                st.warning("No records found after processing.")
                return

            st.success("File processed successfully!")

            # Preview
            st.subheader("Preview of Processed Data (One Row per Professor)")
            st.dataframe(final_data.head())

            # Metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Records", len(final_data))
            with col2:
                st.metric("Average Score", f"{final_data['Average'].mean():.2f}")
            with col3:
                st.metric("Avg Marks (Out of 20)", f"{final_data['Marks_Out_Of_20'].mean():.2f}")
            with col4:
                st.metric("Avg Percentage", f"{final_data['Percentage'].mean():.2f}%")

            # Safer count by year (handles string/numeric)
            period_year = pd.to_numeric(final_data['Highest_Score_Period'], errors='coerce')
            col5, col6, col7, col8 = st.columns(4)
            with col5:
                st.metric("2024 Records", int((period_year == 2024).sum()))
            with col6:
                st.metric("2025 Records", int((period_year == 2025).sum()))

            # Distribution by period label
            st.subheader("Highest Score Distribution by Period")
            period_counts = final_data['Highest_Score_Period'].value_counts()
            st.bar_chart(period_counts)

            # Download
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"Final_Report_Highest_Scores_{timestamp}.xlsx"

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_data.to_excel(writer, index=False, sheet_name='Highest Scores')

            output.seek(0)
            st.download_button(
                label="📥 Download Excel File with Highest Scores",
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Contains one row per professor (highest average per person). "
                     "Professors in teacher-contract columns are included even if marks are 0."
            )

            # Optional: sample listing
            st.subheader("Sample of Included Professors")
            st.info("Names shown here are exactly those in the exported file.")
            st.dataframe(
                final_data[['DNI', 'Nombre', 'Apellido(s)', 'Highest_Score_Period']].head(20)
            )

        except Exception as e:
            st.error(f"An error occurred while processing the file: {str(e)}")
            st.info("Please make sure your Excel file has the required sheets: "
                    "'Inducción', 'nota Inducción', 'Bus. biblioteca', 'Diseño de sesión', and 'Comp. Tec'.")

    else:
        st.info("👆 Please upload an Excel file to get started.")
        st.subheader("Expected Excel File Format")
        st.markdown("""
        Your Excel file should contain the following sheets:
        - **Inducción**: Basic professor information and grades
        - **nota Inducción**: Detailed course grades
        - **Bus. biblioteca**: Library search grades
        - **Diseño de sesión**: Session design grades
        - **Comp. Tec**: Technical competency grades
        
        The processor combines all these sheets, compares periods, and shows one row per professor.
        Professors marked in any teacher-contract column are included even if marks are 0.
        """)
if __name__ == "__main__":
    main()
