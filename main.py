import streamlit as st
import os
import pandas as pd
import datetime
import getpass  # To get the username dynamically
from pdf_editor import pdf_editors
import sys

# Streamlit UI
st.title("PDF Fyller")
st.write("Fyll ut PDF-er basert på Excel-data")

# File uploader
uploaded_file = st.file_uploader("Velg en Excel-fil", type=["xlsx"])

# Checkboxes for document selection
generate_ferdigattest = st.checkbox("Søknad om ferdigattest")
generate_tillatelse = st.checkbox("Søknad om tillatelse til tiltak")

# User-specified output folder
default_folder = os.path.join("C:\\Users", getpass.getuser(), "Downloads")
output_folder = st.text_input("Velg hvor filene skal lagres", default_folder)

if st.button("Skriv til PDF") and uploaded_file:
    try:
        # Define PDF template paths
        if getattr(sys, 'frozen', False):  # Check if running as an executable
            current_directory = os.path.dirname(sys.executable)
        else:
            current_directory = os.path.dirname(os.path.abspath(__file__))

        pdf_template_ferdigattest = os.path.join(current_directory, "Grunnlag skjemaer", "Søknad om ferdigattest-Blank.pdf")
        pdf_template_tillatelse = os.path.join(current_directory, "Grunnlag skjemaer", "Søknad om tillatelse til tiltak-Blank.pdf")

        # Check if PDF files exist
        if generate_ferdigattest and not os.path.exists(pdf_template_ferdigattest):
            st.error(f"Filen {pdf_template_ferdigattest} finnes ikke.")
            st.stop()
        if generate_tillatelse and not os.path.exists(pdf_template_tillatelse):
            st.error(f"Filen {pdf_template_tillatelse} finnes ikke.")
            st.stop()

        # Ensure the output folder exists
        os.makedirs(output_folder, exist_ok=True)

        # Read Excel file
        df_ferdigattest = pd.read_excel(uploaded_file, sheet_name='Skjema ferdigattest', engine='openpyxl').fillna('')
        df_tillatelse = pd.read_excel(uploaded_file, sheet_name='Skjema tillatelse til tiltak', engine='openpyxl').fillna('')

        # Process PDF generation
        def generate_pdfs(df, template, doc_type):
            columns_of_interest = df.iloc[:, 1:].values
            form_fields = list(pdf_editors.get_form_fields(template).keys())
            output_files = []

            counter = 0
            for col_index, column_data in enumerate(columns_of_interest.T):
                if counter >= 1:
                    continue
                data_dict = {}
                for row_index, cell_value in enumerate(column_data):
                    if isinstance(cell_value, datetime.datetime):
                        cell_value = cell_value.strftime("%Y-%m-%d")
                    if row_index == 0:
                        filename = f"{cell_value}"  # Ensure only "Utfylt" files are generated
                    else:
                        if row_index < len(form_fields):
                            data_dict[form_fields[row_index]] = cell_value

                output_file = os.path.join(output_folder, f"{filename}_{doc_type}.pdf")
                try:
                    pdf_editors.write_fillable_pdf(template, output_file, data_dict)
                    output_files.append(output_file)
                except PermissionError:
                    st.error(f"Kan ikke skrive til {output_file}. Filen er kanskje åpen i et annet program.")
                counter += 1

            return output_files[0] if len(output_files) > 1 else output_files

        all_generated_files = []

        if generate_ferdigattest:
            file = generate_pdfs(df_ferdigattest, pdf_template_ferdigattest, "ferdigattest")
            all_generated_files.append(file)

        if generate_tillatelse:
            file = generate_pdfs(df_tillatelse, pdf_template_tillatelse, "tillatelse")
            all_generated_files.append(file)

        if all_generated_files:
            st.success(f"PDF-generering fullført. Filene er lagret i: {output_folder}")

    except Exception as e:
        st.error(f"En feil oppstod: {e}")
