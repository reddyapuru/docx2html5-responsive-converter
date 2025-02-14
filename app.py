import os
import re
import zipfile
import xml.etree.ElementTree as ET
import subprocess
import tempfile

from flask import Flask, request, render_template_string, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'supersecretkey'  # change this in production

# Hardcoded path for LibreOffice CLI on Linux
LIBREOFFICE_PATH = r"/usr/bin/libreoffice"

# Allowed file extensions
ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_namespaces(docx_path):
    """Extracts XML namespaces from document.xml inside a DOCX file."""
    namespaces = {}
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            for event, elem in ET.iterparse(docx_zip.open('word/document.xml'), events=['start-ns']):
                namespaces[elem[0]] = elem[1]
    except Exception as e:
        print(f"⚠ Warning: Could not extract namespaces - {e}")
    return namespaces

def extract_alt_text_from_docx(docx_path):
    """
    Extracts alternative text descriptions for images from a DOCX file,
    mapping the image's 'name' (as defined in <wp:docPr>) to its alt text.
    """
    alt_texts = {}

    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            xml_content = docx_zip.read('word/document.xml')
            tree = ET.ElementTree(ET.fromstring(xml_content))
            root = tree.getroot()

            # Get namespaces dynamically
            namespaces = get_namespaces(docx_path)
            wp_ns = namespaces.get('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')

            print("\n🔍 Extracting Alt Texts from <wp:docPr> elements...")
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
    return alt_texts

def optimize_html(html_file, alt_texts):
    """
    Cleans and optimizes the LibreOffice-generated HTML for responsiveness.
    """
    if not html_file.lower().endswith(".html"):
        return f"❌ Error: The provided file is not an HTML file: {html_file}"

    try:
        with open(html_file, "r", encoding="utf-8", errors="ignore") as file:
            html_content = file.read()

        # Inject responsive meta tags, Bootstrap CSS, and custom CSS in the <head> section
        responsive_head = """
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <style>
      /* Custom CSS styles */
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
    <!-- Google Analytics Tag (if needed) -->
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

        # Wrap body content in a Bootstrap container if not already wrapped
        if not re.search(r'<body[^>]*class="[^"]*container[^"]*"', html_content):
            html_content = re.sub(r'<body', '<body class="container"', html_content)

        # Remove fixed width and height attributes from <img> tags
        html_content = re.sub(r'\s*(width|height)="[^"]*"', '', html_content)

        # Update image tags: add alt attribute and Bootstrap responsive classes
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
        
        with open(html_file, "w", encoding="utf-8") as file:
            file.write(html_content)
        return html_file

    except Exception as e:
        return f"❌ Error processing HTML file: {e}"

def extract_images_from_docx(docx_path, destination_folder):
    """
    Extracts images from the DOCX file's word/media folder into destination_folder.
    """
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
        print(f"⚠ Error extracting images: {e}")

def package_conversion(docx_path, responsive_html_file):
    """
    Extracts images from the DOCX and packages the responsive HTML and images folder into a ZIP file.
    """
    output_dir = os.path.dirname(responsive_html_file)
    images_dir = os.path.join(output_dir, "images")
    os.makedirs(images_dir, exist_ok=True)
    
    # Extract images from the DOCX file
    extract_images_from_docx(docx_path, images_dir)
    
    # Update the HTML file to reference the 'images' folder for any image sources not starting with http or images/
    with open(responsive_html_file, "r", encoding="utf-8") as f:
        html_content = f.read()
    updated_html = re.sub(
        r'src="(?!http)(?!images/)([^"]+)"',
        lambda m: f'src="images/{os.path.basename(m.group(1))}"',
        html_content
    )
    with open(responsive_html_file, "w", encoding="utf-8") as f:
        f.write(updated_html)
    
    # Create a ZIP file containing the HTML and the images folder
    zip_filename = responsive_html_file.replace("_responsive.html", "_package.zip")
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        # Add the responsive HTML file
        zipf.write(responsive_html_file, arcname=os.path.basename(responsive_html_file))
        # Add images
        for root, _, files in os.walk(images_dir):
            for file in files:
                full_path = os.path.join(root, file)
                arcname = os.path.join("images", file)
                zipf.write(full_path, arcname=arcname)
    return zip_filename

def convert_docx_to_html(docx_path):
    """
    Converts a DOCX file to HTML using LibreOffice CLI in headless mode,
    then optimizes the HTML and packages it with extracted images.
    """
    if not os.path.exists(docx_path):
        return f"❌ Error: File '{docx_path}' not found."

    if not os.path.exists(LIBREOFFICE_PATH):
        return f"❌ Error: LibreOffice not found at '{LIBREOFFICE_PATH}'."

    output_dir = os.path.dirname(docx_path)
    alt_texts = extract_alt_text_from_docx(docx_path)
    command = [
        LIBREOFFICE_PATH, "--headless", "--convert-to", "html", "--outdir", output_dir, docx_path
    ]
    
    try:
        subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        html_file = os.path.join(output_dir, os.path.basename(docx_path).replace(".docx", ".html"))
        if os.path.exists(html_file):
            optimized_html_file = optimize_html(html_file, alt_texts)
            # Package the HTML and images into a zip file
            package_file = package_conversion(docx_path, optimized_html_file)
            return package_file  # Return the path to the zip file
        else:
            return "❌ Error: Conversion failed. HTML file not created."
    except subprocess.CalledProcessError as e:
        return f"❌ Error during conversion: {e}"

# Flask Routes
@app.route("/", methods=["GET"])
def index():
    return render_template_string("""
    <!doctype html>
    <html lang="en">
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
        <title>DOCX to Responsive HTML Converter</title>
      </head>
      <body class="container mt-5">
        <h1>DOCX to Responsive HTML Converter</h1>
        <form method="post" action="{{ url_for('convert') }}" enctype="multipart/form-data">
          <div class="mb-3">
            <label for="docx_file" class="form-label">Upload DOCX File</label>
            <input class="form-control" type="file" id="docx_file" name="docx_file" required>
          </div>
          <button type="submit" class="btn btn-primary">Convert</button>
        </form>
        {% with messages = get_flashed_messages() %}
          {% if messages %}
            <div class="alert alert-info mt-3">
              {% for message in messages %}
                <p>{{ message }}</p>
              {% endfor %}
            </div>
          {% endif %}
        {% endwith %}
      </body>
    </html>
    """)

@app.route("/convert", methods=["POST"])
def convert():
    if "docx_file" not in request.files:
        flash("No file part")
        return redirect(url_for("index"))
    
    file = request.files["docx_file"]
    if file.filename == "":
        flash("No file selected")
        return redirect(url_for("index"))
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        upload_dir = tempfile.mkdtemp()
        file_path = os.path.join(upload_dir, filename)
        file.save(file_path)

        result = convert_docx_to_html(file_path)
        if result.startswith("❌"):
            flash(result)
            return redirect(url_for("index"))
        else:
            # Send the resulting zip file for download
            return send_file(result, as_attachment=True)
    else:
        flash("Invalid file format. Please upload a DOCX file.")
        return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)
