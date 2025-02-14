import os
import re
import zipfile
import xml.etree.ElementTree as ET
import subprocess
import tempfile
from bs4 import BeautifulSoup  # Make sure to install beautifulsoup4

import streamlit as st

# Hardcoded path for LibreOffice CLI on Linux
LIBREOFFICE_PATH = r"/usr/bin/libreoffice"

# Allowed file extensions
ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_namespaces(docx_path):
    """Extracts XML namespaces from document.xml inside a DOCX file."""
    namespaces = {}
    st.text("Extracting namespaces from DOCX...")
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            for event, elem in ET.iterparse(docx_zip.open('word/document.xml'), events=['start-ns']):
                namespaces[elem[0]] = elem[1]
    except Exception as e:
        st.error(f"⚠ Warning: Could not extract namespaces - {e}")
    st.text("Namespace extraction completed.")
    return namespaces

def extract_alt_text_from_docx(docx_path):
    """
    Extracts alternative text descriptions for images from a DOCX file,
    mapping the image's 'name' (as defined in <wp:docPr>) to its alt text.
    """
    alt_texts = {}
    st.text("Extracting alt texts from DOCX...")
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            xml_content = docx_zip.read('word/document.xml')
            tree = ET.ElementTree(ET.fromstring(xml_content))
            root = tree.getroot()
            namespaces = get_namespaces(docx_path)
            wp_ns = namespaces.get('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
            st.text("Processing <wp:docPr> elements...")
            for docPr in root.findall(f'.//{{{wp_ns}}}docPr'):
                alt_text = docPr.attrib.get('descr', '').strip()
                image_name = docPr.attrib.get('name', '').strip()
                if alt_text and image_name:
                    alt_texts[image_name] = alt_text
                    st.text(f"  Mapped '{image_name}' → '{alt_text}'")
                else:
                    st.warning(f"  ⚠ Skipping element, missing alt text or name: {docPr.attrib}")
    except Exception as e:
        st.error(f"⚠ Warning: Failed to extract alt text from DOCX - {e}")
    if not alt_texts:
        st.error("❌ No alt texts were extracted.")
    else:
        st.text("Alt text extraction completed.")
    return alt_texts

def optimize_html(html_file, alt_texts):
    """
    Cleans and optimizes the LibreOffice-generated HTML for responsiveness,
    using BeautifulSoup to remove extraneous markup.
    """
    st.text("Starting HTML optimization with BeautifulSoup...")
    if not html_file.lower().endswith(".html"):
        return f"❌ Error: The provided file is not an HTML file: {html_file}"
    try:
        with open(html_file, "r", encoding="utf-8", errors="ignore") as file:
            html_content = file.read()

        soup = BeautifulSoup(html_content, "html.parser")
        # Remove extraneous inline styles or tags (customize as needed)
        for span in soup.find_all("span", style=True):
            span.unwrap()

        new_head = BeautifulSoup("""
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
                <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
                <style>
                    :root {
                        --font-base: clamp(0.75rem, 1vw + 0.75rem, 1.25rem);
                        --font-headline: clamp(1.75rem, 4vw, 2.5rem);
                        --spacing-base: clamp(0.5rem, 1vw, 2rem);
                        --line-height-base: 1.5;
                        --vertical-spacing: clamp(1.3, 1vw + 1.3, 1.7);
                        --font-primary: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
                        --font-secondary: "Segoe UI Black", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
                    }
                    html { font-size: 100%; line-height: var(--line-height-base); font-family: var(--font-primary); }
                    header { background: rgba(255, 255, 255, 0.85); backdrop-filter: blur(10px); border-bottom: 1px solid rgba(0,0,0,0.1); padding: calc(var(--spacing-base) * 1.5); text-align: center; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
                    header h1 { margin: 0; font-family: var(--font-secondary); font-size: var(--font-headline); font-weight: 900; letter-spacing: -0.5pt; line-height: 1.3; }
                    body { padding: var(--spacing-base); }
                    img { max-width: 100% !important; height: auto !important; display: block; }
                    .img-line { width: 100% !important; height: auto !important; }
                    .table-responsive { overflow-x: auto; }
                    footer { margin-top: var(--spacing-base); padding: var(--spacing-base); background-color: #f8f9fa; text-align: center; font-size: clamp(0.75rem, 1vw, 1rem); }
                </style>
                <script async src="https://www.googletagmanager.com/gtag/js?id=G-P8LYBP9EDY"></script>
                <script defer>
                    window.dataLayer = window.dataLayer || [];
                    function gtag(){dataLayer.push(arguments);}
                    gtag('js', new Date());
                    gtag('config', 'G-P8LYBP9EDY');
                </script>
            </head>
        """, "html.parser")
        if soup.head:
            soup.head.replace_with(new_head.head)
        else:
            soup.insert(0, new_head.head)

        if soup.body:
            body_class = soup.body.get("class", [])
            if "container" not in body_class:
                body_class.append("container")
                soup.body["class"] = body_class

        updated_html = str(soup)

        with open(html_file, "w", encoding="utf-8") as file:
            file.write(updated_html)
        st.text("HTML optimization completed with BeautifulSoup.")
        return html_file
    except Exception as e:
        error_message = f"❌ Error processing HTML file: {e}"
        st.error(error_message)
        return error_message

def extract_images_from_docx(docx_path, destination_folder):
    """
    Extracts images from the DOCX file's word/media folder into destination_folder.
    """
    st.text("Extracting images from DOCX...")
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            for file in docx_zip.namelist():
                if file.startswith("word/media/"):
                    filename = os.path.basename(file)
                    if filename:
                        dest_path = os.path.join(destination_folder, filename)
                        with open(dest_path, "wb") as f:
                            f.write(docx_zip.read(file))
                        st.text(f"Extracted image: {filename}")
    except Exception as e:
        st.error(f"⚠ Error extracting images: {e}")
    st.text("Image extraction completed.")

def package_conversion(docx_path, responsive_html_file):
    """
    Extracts images from the DOCX and packages the responsive HTML and images folder into a ZIP file.
    """
    st.text("Starting packaging of conversion results...")
    output_dir = os.path.dirname(responsive_html_file)
    images_dir = os.path.join(output_dir, "images")
    os.makedirs(images_dir, exist_ok=True)
    extract_images_from_docx(docx_path, images_dir)
    with open(responsive_html_file, "r", encoding="utf-8") as f:
        html_content = f.read()
    updated_html = re.sub(
        r'src="(?!images/)(?:media/)?([^"]+)"',
        lambda m: f'src="images/{os.path.basename(m.group(1))}"',
        html_content
    )
    with open(responsive_html_file, "w", encoding="utf-8") as f:
        f.write(updated_html)
    zip_filename = responsive_html_file.replace("_responsive.html", "_package.zip")
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        zipf.write(responsive_html_file, arcname=os.path.basename(responsive_html_file))
        for root, _, files in os.walk(images_dir):
            for file in files:
                full_path = os.path.join(root, file)
                arcname = os.path.join("images", file)
                zipf.write(full_path, arcname=arcname)
    st.text(f"Packaging completed. Package file: {zip_filename}")
    return zip_filename

def convert_docx_to_html(docx_path):
    """
    Converts a DOCX file to HTML using LibreOffice CLI in headless mode,
    then optimizes the HTML and packages it with extracted images.
    """
    st.text("Starting DOCX to HTML conversion...")
    if not os.path.exists(docx_path):
        error_message = f"❌ Error: File '{docx_path}' not found."
        st.error(error_message)
        return error_message
    if not os.path.exists(LIBREOFFICE_PATH):
        error_message = f"❌ Error: LibreOffice not found at '{LIBREOFFICE_PATH}'."
        st.error(error_message)
        return error_message
    output_dir = os.path.dirname(docx_path)
    alt_texts = extract_alt_text_from_docx(docx_path)
    command = [
        LIBREOFFICE_PATH, "--headless", "--convert-to", "html", "--outdir", output_dir, docx_path
    ]
    try:
        st.text("Running LibreOffice conversion...")
        subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        st.text("LibreOffice conversion completed.")
        html_file = os.path.join(output_dir, os.path.basename(docx_path).replace(".docx", ".html"))
        if os.path.exists(html_file):
            optimized_html_file = optimize_html(html_file, alt_texts)
            package_file = package_conversion(docx_path, optimized_html_file)
            st.success("DOCX conversion and packaging completed successfully.")
            return package_file
        else:
            error_message = "❌ Error: Conversion failed. HTML file not created."
            st.error(error_message)
            return error_message
    except subprocess.CalledProcessError as e:
        error_message = f"❌ Error during conversion: {e}"
        st.error(error_message)
        return error_message

# --- Streamlit App Layout ---
st.title("DOCX to Responsive HTML Converter")
st.write("Upload a DOCX file to convert it to responsive HTML along with its images packaged in a ZIP file.")

uploaded_file = st.file_uploader("Choose a DOCX file", type=["docx"])

if uploaded_file is not None:
    if allowed_file(uploaded_file.name):
        # Save the uploaded file to a temporary directory
        upload_dir = tempfile.mkdtemp()
        file_path = os.path.join(upload_dir, secure_filename(uploaded_file.name))
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.write(f"File saved to: {file_path}")
        
        # Convert the DOCX file to a ZIP package
        package_path = convert_docx_to_html(file_path)
        
        if not package_path.startswith("❌"):
            # Read the ZIP file as binary for download
            with open(package_path, "rb") as f:
                zip_data = f.read()
            st.download_button(
                label="Download Conversion Package",
                data=zip_data,
                file_name=os.path.basename(package_path),
                mime="application/zip"
            )
        else:
            st.error("Conversion failed.")
    else:
        st.error("Invalid file format. Please upload a DOCX file.")
