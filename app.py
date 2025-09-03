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
    
    # Combine both datasets
    all_data = pd.concat([nota_induccion_clean, induction_clean])
    
    # Merge with bus biblioteca data
    bus_biblioteca_df = bus_biblioteca_df.rename(columns={'Promedio': 'bus_biblioteca'})
    all_data = pd.merge(all_data, bus_biblioteca_df[['DNI', 'bus_biblioteca']], 
                        on='DNI', how='left')
    
    # Merge with diseño de sesión data
    diseno_sesion_df = diseno_sesion_df.rename(columns={'Promedio': 'diseno_sesion'})
    all_data = pd.merge(all_data, diseno_sesion_df[['Nombre', 'Apellido(s)', 'diseno_sesion']], 
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
    all_data = pd.merge(all_data, comp_tec_df[['Nombre', 'Apellido(s)', 'Zoom_basico', 'Zoom_Avanzado', 
                                               'Grupos_Moodle', 'Rubrica', 'Padlet', 'Nearpod', 'Tareas_y_foros']], 
                        on=['Nombre', 'Apellido(s)'], how='left')
    
    # Define numeric columns
    numeric_columns = ['induccion', 'bus_biblioteca', 'diseno_sesion', 'Zoom_basico', 
                      'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica', 'Padlet', 'Nearpod', 'Tareas_y_foros']
    
    # Replace empty strings and NaN with 0 for numeric columns
    for col in numeric_columns:
        all_data[col] = all_data[col].replace('', 0)
        all_data[col] = pd.to_numeric(all_data[col], errors='coerce').fillna(0)
    
    # Calculate average of all 10 marks (treating blanks as 0)
    all_data['Average'] = all_data[numeric_columns].mean(axis=1)
    
    # Format the average to 2 decimal places
    all_data['Average'] = all_data['Average'].round(2)
    
    # Calculate percentage based on available marks only
    def calculate_percentage(row):
        # Get the list of scores for the 10 components
        scores = row[numeric_columns].values
        # Count how many components have non-zero scores
        available_components = sum(score > 0 for score in scores)
        
        if available_components == 0:
            return 0
        
        # Calculate the sum of available scores
        total_score = sum(scores)
        # Calculate percentage (sum of scores / number of available components)
        percentage = total_score / available_components
        return round(percentage, 2)
    
    # Apply the percentage calculation to each row
    all_data['Percentage'] = all_data.apply(calculate_percentage, axis=1)
    
    # Create a unique identifier for each person
    all_data['Person_ID'] = all_data['DNI'].astype(str) + '_' + all_data['Nombre'] + '_' + all_data['Apellido(s)']
    
    # Filter out records where average is 0 (no scores available)
    all_data = all_data[all_data['Average'] > 0]
    
    # If no records with scores, return empty dataframe
    if len(all_data) == 0:
        return pd.DataFrame()
    
    # Group by person and find the highest average score across periods
    # Get the index of the row with the maximum average for each person
    idx = all_data.groupby('Person_ID')['Average'].idxmax()
    
    # Select the rows with the highest scores
    highest_scores = all_data.loc[idx].copy()
    
    # Add a column to indicate which period had the highest score
    highest_scores['Highest_Score_Period'] = highest_scores['Periodo']
    
    # Select and order the required columns for final output
    final_columns = [
        'Periodo', 'DNI', 'Nombre', 'Apellido(s)', 'induccion', 'bus_biblioteca', 'diseno_sesion',
        'Zoom_basico', 'Zoom_Avanzado', 'Grupos_Moodle', 'Rubrica', 'Padlet', 'Nearpod', 'Tareas_y_foros',
        'Average', 'Percentage', 'Highest_Score_Period'
    ]
    
    final_df = highest_scores[final_columns]
    
    return final_df

def main():
    st.set_page_config(page_title="Excel Data Processor", page_icon="📊", layout="wide")
    
    st.title("📊 Excel Data Processor")
    st.markdown("Upload your Excel file to process and combine data from multiple sheets.")
    st.info("This tool will compare scores from both 2024 and 2025 periods and show the highest score for each person.")
    
    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            # Process the file
            with st.spinner("Processing your Excel file and comparing periods..."):
                final_data = extract_data_from_excel(uploaded_file)
            
            if len(final_data) == 0:
                st.warning("No records with scores found in the uploaded file.")
                return
            
            st.success("File processed successfully!")
            
            # Display preview
            st.subheader("Preview of Processed Data (Showing Highest Scores)")
            st.dataframe(final_data.head())
            
            # Show some statistics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Records", len(final_data))
            with col2:
                st.metric("Average Score", f"{final_data['Average'].mean():.2f}")
            with col3:
                st.metric("Average Percentage", f"{final_data['Percentage'].mean():.2f}")
            with col4:
                st.metric("2024 Records", len(final_data[final_data['Highest_Score_Period'] == 2024]))
            
            col5, col6, col7, col8 = st.columns(4)
            with col5:
                st.metric("2025 Records", len(final_data[final_data['Highest_Score_Period'] == 2025]))
            
            # Show distribution of highest scores by period
            st.subheader("Highest Score Distribution by Period")
            period_counts = final_data['Highest_Score_Period'].value_counts()
            st.bar_chart(period_counts)
            
            # Create download button
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"Final_Report_Highest_Scores_{timestamp}.xlsx"
            
            # Convert DataFrame to Excel bytes
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_data.to_excel(writer, index=False, sheet_name='Highest Scores')
            
            output.seek(0)
            
            st.download_button(
                label="📥 Download Excel File with Highest Scores",
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="This file contains the highest scores for each person across both 2024 and 2025 periods"
            )
            
            # Show some examples of people with scores from both periods
            st.subheader("Sample of People with Scores from Both Periods")
            st.info("The downloaded file shows only the highest score for each person. Below are some examples where people have scores from both periods.")
            
        except Exception as e:
            st.error(f"An error occurred while processing the file: {str(e)}")
            st.info("Please make sure your Excel file has the required sheets: 'Inducción', 'nota Inducción', 'Bus. biblioteca', 'Diseño de sesión', and 'Comp. Tec'")
    
    else:
        st.info("👆 Please upload an Excel file to get started.")
        
        # Show expected format
        st.subheader("Expected Excel File Format")
        st.markdown("""
        Your Excel file should contain the following sheets:
        - **Inducción**: Contains basic student information and grades
        - **nota Inducción**: Contains detailed course grades
        - **Bus. biblioteca**: Contains library search grades
        - **Diseño de sesión**: Contains session design grades
        - **Comp. Tec**: Contains technical competency grades
        
        The processor will combine data from all these sheets, compare scores from 2024 and 2025 periods,
        and show the highest score for each person.
        """)

if __name__ == "__main__":
    main()