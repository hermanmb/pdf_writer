import streamlit as st
import os
import pandas as pd
import datetime
import zipfile
import tempfile  # Use temporary files
from pdf_editor import pdf_editors
import sys

# Streamlit UI
st.title("PDF Fyller")
st.write("Fyll ut PDF-er basert på Excel-data")

# File uploader
uploaded_file = st.file_uploader("Velg en Excel-fil", type=["xlsx"])

# Checkboxes to choose which PDF to generate
generate_ferdigattest = st.checkbox("Søknad om ferdigattest")
generate_tillatelse = st.checkbox("Søknad om tillatelse til tiltak")

# Initialize session state for storing generated files
if "generated_files" not in st.session_state:
    st.session_state.generated_files = []

if st.button("Skriv til PDF") and uploaded_file:
    try:
        # Define PDF template paths
        if getattr(sys, 'frozen', False):  # Running as an executable
            current_directory = os.path.dirname(sys.executable)
        else:
            current_directory = os.path.dirname(os.path.abspath(__file__))

        pdf_template_ferdigattest = os.path.join(current_directory, "Grunnlag skjemaer", "Søknad om ferdigattest-Blank.pdf")
        pdf_template_tillatelse = os.path.join(current_directory, "Grunnlag skjemaer", "Søknad om tillatelse til tiltak-Blank.pdf")

        # Check if the PDF templates exist
        if generate_ferdigattest and not os.path.exists(pdf_template_ferdigattest):
            st.error(f"Filen {pdf_template_ferdigattest} finnes ikke.")
            st.stop()
        if generate_tillatelse and not os.path.exists(pdf_template_tillatelse):
            st.error(f"Filen {pdf_template_tillatelse} finnes ikke.")
            st.stop()

        # Read Excel data
        df_ferdigattest = pd.read_excel(uploaded_file, sheet_name='Skjema ferdigattest', engine='openpyxl').fillna('')
        df_tillatelse = pd.read_excel(uploaded_file, sheet_name='Skjema tillatelse til tiltak', engine='openpyxl').fillna('')

        # Function to generate PDFs
        def generate_pdfs(df, template, doc_type):
            columns_of_interest = df.iloc[:, 1:].values
            form_fields = list(pdf_editors.get_form_fields(template).keys())
            output_files = []
            counter = 0  # Only process the first column of data

            for col_index, column_data in enumerate(columns_of_interest.T):
                if counter > 0:
                    continue

                data_dict = {}
                filename = "Utfylt"  # Default filename

                for row_index, cell_value in enumerate(column_data):
                    if isinstance(cell_value, datetime.datetime):
                        cell_value = cell_value.strftime("%Y-%m-%d")
                    if row_index == 0 and cell_value:
                        filename = str(cell_value)  # Use first row as filename
                    else:
                        if row_index < len(form_fields):
                            data_dict[form_fields[row_index]] = cell_value

                # Use a temporary file instead of saving to project folder
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
                output_file = f"{filename}_{doc_type}.pdf"

                try:
                    pdf_editors.write_fillable_pdf(template, output_file, data_dict)
                    output_files.append(output_file)
                except PermissionError:
                    st.error(f"Kan ikke skrive til {output_file}. Filen er kanskje åpen i et annet program.")

                counter += 1  # Only process one column
            return output_files

        # Generate PDFs and store in session state
        generated_files = []
        if generate_ferdigattest:
            generated_files += generate_pdfs(df_ferdigattest, pdf_template_ferdigattest, "ferdigattest")
        if generate_tillatelse:
            generated_files += generate_pdfs(df_tillatelse, pdf_template_tillatelse, "tillatelse")

        st.session_state.generated_files = generated_files  # Store generated files

        st.success("PDF-generering fullført!")

    except Exception as e:
        st.error(f"En feil oppstod: {e}")

# Download buttons (persist across reruns)
if st.session_state.generated_files:
    for i, file in enumerate(st.session_state.generated_files):
        with open(file, "rb") as f:
            st.download_button(label=f"Last ned {os.path.basename(file)}",
                               data=f,
                               file_name=os.path.basename(file),
                               mime="application/pdf",
                               key=f"download_{i}")  # Unique key

    # Offer ZIP download if multiple files exist
    if len(st.session_state.generated_files) > 1:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as temp_zip:
            zip_filename = temp_zip.name
            with zipfile.ZipFile(zip_filename, "w") as zipf:
                for file in st.session_state.generated_files:
                    zipf.write(file, os.path.basename(file))

        with open(zip_filename, "rb") as f:
            st.download_button(label="Last ned alle PDF-er som ZIP",
                               data=f,
                               file_name="PDFs.zip",
                               mime="application/zip",
                               key="download_zip")  # Unique key for ZIP
