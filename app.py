import os
import tempfile
import streamlit as st
from werkzeug.utils import secure_filename

# Import functions from the libre-docx2html5.py file.
# Make sure the file is named exactly "libre-docx2html5.py" and is in the same directory.
from libre_docx2html5 import convert_docx_to_html, allowed_file

# Streamlit UI layout
st.title("DOCX to Responsive HTML Converter")
st.write("Upload a DOCX file to convert it to responsive HTML along with its images packaged in a ZIP file.")

# File uploader for DOCX files
uploaded_file = st.file_uploader("Choose a DOCX file", type=["docx"])

if uploaded_file is not None:
    # Check if the file is allowed
    if allowed_file(uploaded_file.name):
        # Save the uploaded file to a temporary directory
        upload_dir = tempfile.mkdtemp()
        file_path = os.path.join(upload_dir, secure_filename(uploaded_file.name))
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.write(f"File saved to: {file_path}")
        
        # Convert the DOCX file using the function from libre-docx2html5.py
        package_path = convert_docx_to_html(file_path)
        
        # Check if conversion returned an error message
        if package_path.startswith("❌"):
            st.error(package_path)
        else:
            st.success("Conversion completed successfully!")
            # Read the resulting ZIP file and create a download button
            with open(package_path, "rb") as f:
                zip_data = f.read()
            st.download_button(
                label="Download Conversion Package",
                data=zip_data,
                file_name=os.path.basename(package_path),
                mime="application/zip"
            )
    else:
        st.error("Invalid file format. Please upload a DOCX file.")
