import os
import re
import zipfile
import xml.etree.ElementTree as ET
import subprocess
import tempfile
import datetime
import threading
import time
import shutil
from lxml import html  # requires lxml package

# Hardcoded path for LibreOffice CLI (adjust for your platform)
#LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
# For Linux, you would use:
LIBREOFFICE_PATH = r"/usr/bin/libreoffice"

def get_namespaces(docx_path):
    """Extracts XML namespaces from document.xml inside a DOCX file."""
    namespaces = {}
    print("Extracting namespaces from DOCX...")
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            for event, elem in ET.iterparse(docx_zip.open('word/document.xml'), events=['start-ns']):
                namespaces[elem[0]] = elem[1]
    except Exception as e:
        print(f"⚠ Warning: Could not extract namespaces - {e}")
    print("Namespace extraction completed.")
    return namespaces

def extract_alt_text_from_docx(docx_path):
    """
    Extracts alternative text descriptions for images from a DOCX file,
    mapping the image's 'name' (as defined in <wp:docPr>) to its alt text.
    """
    alt_texts = {}
    print("Extracting alt texts from DOCX...")
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            xml_content = docx_zip.read('word/document.xml')
            tree = ET.ElementTree(ET.fromstring(xml_content))
            root = tree.getroot()
            namespaces = get_namespaces(docx_path)
            wp_ns = namespaces.get('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
            print("Processing <wp:docPr> elements...")
            for docPr in root.findall(f'.//{{{wp_ns}}}docPr'):
                alt_text = docPr.attrib.get('descr', '').strip()
                image_name = docPr.attrib.get('name', '').strip()
                if alt_text and image_name:
                    alt_texts[image_name] = alt_text
                    print(f"  Mapped '{image_name}' → '{alt_text}'")
                else:
                    print(f"  ⚠ Skipping element, missing alt text or name: {docPr.attrib}")
    except Exception as e:
        print(f"⚠ Warning: Failed to extract alt text from DOCX - {e}")
    if not alt_texts:
        print("❌ No alt texts were extracted.")
    else:
        print("Alt text extraction completed.")
    return alt_texts

def optimize_html(html_file, alt_texts):
    """
    Cleans and optimizes the LibreOffice-generated HTML for responsiveness.
    It injects a clean <head> section and updates image tags.
    """
    print("Starting HTML optimization...")
    if not html_file.lower().endswith(".html"):
        return f"❌ Error: The provided file is not an HTML file: {html_file}"
    try:
        with open(html_file, "r", encoding="utf-8", errors="ignore") as file:
            html_content = file.read()
        # Replace the <head> with a new responsive head
        responsive_head = """
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
        """
        html_content = re.sub(r'<head>.*?</head>', responsive_head, html_content, flags=re.DOTALL)
        # Ensure body is wrapped in a container
        if not re.search(r'<body[^>]*class="[^"]*container[^"]*"', html_content):
            html_content = re.sub(r'<body', '<body class="container"', html_content)
        # Remove fixed width/height attributes from <img> tags
        html_content = re.sub(r'\s*(width|height)="[^"]*"', '', html_content)
        # Update image tags with proper alt text and responsive classes
        def add_alt_attribute(match):
            img_tag = match.group(0)
            name_match = re.search(r'name="([^"]+)"', img_tag)
            src_match = re.search(r'src="([^"]+)"', img_tag)
            image_description = "Illustration from the document"
            if name_match:
                image_name = name_match.group(1)
                if image_name in alt_texts:
                    image_description = alt_texts[image_name]
                if image_name.lower().startswith("shape"):
                    if 'class=' in img_tag:
                        if 'img-line' not in img_tag:
                            img_tag = re.sub(r'class="([^"]+)"', lambda m: f'class="{m.group(1)} img-line"', img_tag)
                    else:
                        img_tag = re.sub(r'<img', '<img class="img-line"', img_tag)
            elif src_match:
                image_filename = os.path.basename(src_match.group(1))
                if image_filename in alt_texts:
                    image_description = alt_texts[image_filename]
            if not re.search(r'alt="[^"]*"', img_tag):
                img_tag = re.sub(r'<img', f'<img alt="{image_description}"', img_tag)
            else:
                img_tag = re.sub(r'alt="[^"]*"', f'alt="{image_description}"', img_tag)
            if 'class=' in img_tag:
                if 'img-fluid' not in img_tag:
                    img_tag = re.sub(r'class="([^"]+)"', lambda m: f'class="{m.group(1)} img-fluid"', img_tag)
            else:
                img_tag = re.sub(r'<img', '<img class="img-fluid"', img_tag)
            return img_tag
        html_content = re.sub(r'<img[^>]+>', add_alt_attribute, html_content)
        html_content = re.sub(r'(<table[^>]*>.*?</table>)', r'<div class="table-responsive">\1</div>', html_content, flags=re.DOTALL)
        footer_banner = """
        <footer>
            <hr>
            <p>© 2025 www.latest2all.com</p>
        </footer>
        """
        html_content = re.sub(r'</body>', footer_banner + '</body>', html_content, flags=re.IGNORECASE)
        # Save the optimized HTML file with a new name
        cleaned_html_file = html_file.replace(".html", "_responsive.html")
        with open(cleaned_html_file, "w", encoding="utf-8") as file:
            file.write(html_content)
        print("HTML optimization completed.")
        return cleaned_html_file
    except Exception as e:
        return f"❌ Error processing HTML file: {e}"

def extract_images_from_docx(docx_path, destination_folder):
    """
    Extracts images from the DOCX file's word/media folder into destination_folder.
    """
    print("Extracting images from DOCX...")
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            for file in docx_zip.namelist():
                if file.startswith("word/media/"):
                    filename = os.path.basename(file)
                    if filename:
                        dest_path = os.path.join(destination_folder, filename)
                        with open(dest_path, "wb") as f:
                            f.write(docx_zip.read(file))
                        print(f"Extracted image: {filename}")
    except Exception as e:
        print(f"❌ Error extracting images: {e}")
    print("Image extraction completed.")

def convert_docx_to_html(docx_path):
    """
    Converts a DOCX file to HTML using LibreOffice CLI in headless mode,
    then optimizes the HTML and packages it with its images in an output folder.
    
    The output folder is created based on the input file name and the current date/time.
    
    Returns:
        str: Path to the output ZIP package or an error message.
    """
    print("Starting DOCX to HTML conversion...")
    if not os.path.exists(docx_path):
        error_message = f"❌ Error: File '{docx_path}' not found."
        print(error_message)
        return error_message

    if not os.path.exists(LIBREOFFICE_PATH):
        error_message = f"❌ Error: LibreOffice not found at '{LIBREOFFICE_PATH}'."
        print(error_message)
        return error_message

    # Create an output folder using the base file name and current date/time
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    current_date = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_folder = os.path.join(os.path.dirname(docx_path), f"{base_name}_{current_date}")
    os.makedirs(output_folder, exist_ok=True)
    print(f"Output folder created: {output_folder}")

    # Extract alt texts from DOCX
    alt_texts = extract_alt_text_from_docx(docx_path)

    # Run LibreOffice conversion into the output folder
    command = [
        LIBREOFFICE_PATH, "--headless", "--convert-to", "html", "--outdir", output_folder, docx_path
    ]
    try:
        print("Running LibreOffice conversion...")
        subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        print("LibreOffice conversion completed.")
    except subprocess.CalledProcessError as e:
        error_message = f"❌ Error during conversion: {e}"
        print(error_message)
        return error_message

    # Get the generated HTML file from the output folder
    html_file = os.path.join(output_folder, os.path.basename(docx_path).replace(".docx", ".html"))
    if os.path.exists(html_file):
        optimized_html_file = optimize_html(html_file, alt_texts)
        print(f"Optimized HTML file: {optimized_html_file}")

        # Create an images folder in the output folder and extract images there
        images_folder = os.path.join(output_folder, "images")
        os.makedirs(images_folder, exist_ok=True)
        extract_images_from_docx(docx_path, images_folder)

        # Package the entire output folder into a ZIP file
        zip_filename = os.path.join(output_folder, f"{base_name}_{current_date}_package.zip")
        with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(output_folder):
                for file in files:
                    # Skip the ZIP file itself if it exists in the folder
                    if file == os.path.basename(zip_filename):
                        continue
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, output_folder)
                    zipf.write(full_path, arcname=arcname)
        print(f"Packaging completed. Package file: {zip_filename}")

        # Schedule deletion of the entire output folder (including the package and input file) after 10 minutes
        def schedule_deletion(folder_path, input_file, delay=600):
            print(f"Scheduling deletion of all files in {folder_path} and input file {input_file} in {delay} seconds...")
            time.sleep(delay)
            if os.path.exists(folder_path):
                shutil.rmtree(folder_path)
                print(f"Output folder {folder_path} deleted after {delay} seconds.")
            if os.path.exists(input_file):
                os.remove(input_file)
                print(f"Input file {input_file} deleted after {delay} seconds.")

        # Start the deletion thread without joining it
        deletion_thread = threading.Thread(target=schedule_deletion, args=(output_folder, docx_path), daemon=True)
        deletion_thread.start()

        return zip_filename
    else:
        error_message = "❌ Error: Conversion failed. HTML file not created."
        print(error_message)
        return error_message


# For command-line usage (if needed)
if __name__ == "__main__":
    docx_file = input("Enter the full path of the DOCX file: ").strip()
    result = convert_docx_to_html(docx_file)
    print(result)
