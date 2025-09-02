import pandas as pd
import numpy as np
import streamlit as st
import io
from datetime import datetime

def extract_data_from_excel(file_path):
    # Read all sheets
    induction_df = pd.read_excel(file_path, sheet_name='Inducción')
    nota_induccion_df = pd.read_excel(file_path, sheet_name='nota Inducción')
    bus_biblioteca_df = pd.read_excel(file_path, sheet_name='Bus. biblioteca')
    diseno_sesion_df = pd.read_excel(file_path, sheet_name='Diseño de sesión')
    comp_tec_df = pd.read_excel(file_path, sheet_name='Comp. Tec')
    
    # Create a master dataframe with all unique people from both induction sheets
    # First, prepare the nota_induccion data
    nota_induccion_clean = nota_induccion_df[['PERIODO', 'DNI', 'Nombre', 'Apellido(s)', 'Dirección de correo', 'Total del curso (Real)']].copy()
    nota_induccion_clean = nota_induccion_clean.rename(columns={
        'PERIODO': 'Periodo', 
        'Total del curso (Real)': 'induccion'
    })
    
    # Prepare the induction data
    induction_clean = induction_df[['Periodo', 'DNI', 'Nombre', 'Apellido(s)', 'Dirección de correo', 'Calificación']].copy()
    induction_clean = induction_clean.rename(columns={'Calificación': 'induccion'})
    
    # Combine both, prioritizing nota_induccion data
    master_df = pd.concat([nota_induccion_clean, induction_clean])
    
    # Remove duplicates, keeping the first occurrence (which will be from nota_induccion if both exist)
    master_df = master_df.drop_duplicates(subset=['DNI', 'Nombre', 'Apellido(s)'], keep='first')
    
    # Merge with bus biblioteca data
    bus_biblioteca_df = bus_biblioteca_df.rename(columns={'Promedio': 'bus_biblioteca'})
    master_df = pd.merge(master_df, bus_biblioteca_df[['DNI', 'bus_biblioteca']], 
                        on='DNI', how='left')
    
    # Merge with diseño de sesión data
    diseno_sesion_df = diseno_sesion_df.rename(columns={'Promedio': 'diseno_sesion'})
    # We need to match by name since DNI might not be available in this sheet
    master_df = pd.merge(master_df, diseno_sesion_df[['Nombre', 'Apellido(s)', 'diseno_sesion']], 
                        on=['Nombre', 'Apellido(s)'], how='left')
    
    # Merge with competencias técnicas data
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
    # Match by name for competencias técnicas
    master_df = pd.merge(master_df, comp_tec_df[['Nombre', 'Apellido(s)', 'Zoom_basico', 'Zoom_Avanzado', 
                                               'Grupos_Moodle', 'Rubrica', 'Padlet', 'Nearpod', 'Tareas_y_foros']], 
                        on=['Nombre', 'Apellido(s)'], how='left')
    
    # Select and order the required columns
    final_columns = [
        'Periodo', 'DNI', 'Nombre', 'Apellido(s)', 'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica', 'Padlet', 'Nearpod', 'Tareas_y_foros'
    ]
    
    final_df = master_df[final_columns]
    
    # Define numeric columns
    numeric_columns = ['induccion', 'bus_biblioteca', 'diseno_sesion', 'Zoom_basico', 
                      'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica', 'Padlet', 'Nearpod', 'Tareas_y_foros']
    
    # Replace empty strings and NaN with 0 for numeric columns
    for col in numeric_columns:
        final_df[col] = final_df[col].replace('', 0)
        final_df[col] = pd.to_numeric(final_df[col], errors='coerce').fillna(0)
    
    # Calculate average of all 10 marks (treating blanks as 0)
    final_df['Average'] = final_df[numeric_columns].mean(axis=1)
    
    # Format the average to 2 decimal places
    final_df['Average'] = final_df['Average'].round(2)
    
    return final_df

def main():
    st.set_page_config(page_title="Excel Data Processor", page_icon="📊", layout="wide")
    
    st.title("📊 Excel Data Processor")
    st.markdown("Upload your Excel file to process and combine data from multiple sheets.")
    
    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            # Process the file
            with st.spinner("Processing your Excel file..."):
                final_data = extract_data_from_excel(uploaded_file)
            
            st.success("File processed successfully!")
            
            # Display preview
            st.subheader("Preview of Processed Data")
            st.dataframe(final_data.head())
            
            # Show some statistics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Records", len(final_data))
            with col2:
                st.metric("Average Score", f"{final_data['Average'].mean():.2f}")
            with col3:
                st.metric("Columns", len(final_data.columns))
            
            # Create download button
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"Final_Report_{timestamp}.xlsx"
            
            # Convert DataFrame to Excel bytes
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_data.to_excel(writer, index=False, sheet_name='Final Report')
            
            output.seek(0)
            
            st.download_button(
                label="📥 Download Processed Excel File",
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"An error occurred while processing the file: {str(e)}")
            st.info("Please make sure your Excel file has the required sheets: 'Inducción', 'nota Inducción', 'Bus. biblioteca', 'Diseño de sesión', and 'Comp. Tec'")
    
    else:
        st.info("👆 Please upload an Excel file to get started.")

if __name__ == "__main__":
    main()