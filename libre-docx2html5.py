import subprocess
import os
import re
import zipfile
import xml.etree.ElementTree as ET
import streamlit as st

# Hardcoded path for LibreOffice CLI
LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"

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
    
    Args:
        docx_path (str): Path to the DOCX file.
        
    Returns:
        dict: Mapping of image names to alt text descriptions.
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
            # Use the 'name' attribute (present in both DOCX and HTML) as the key
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

def convert_docx_to_html(docx_path):
    """
    Converts a DOCX file to HTML using LibreOffice CLI in headless mode.
    
    Args:
        docx_path (str): Full path to the DOCX file.
        
    Returns:
        str: Path to the responsive HTML file or an error message.
    """
    if not os.path.exists(docx_path):
        return f"❌ Error: File '{docx_path}' not found."

    if not os.path.exists(LIBREOFFICE_PATH):
        return f"❌ Error: LibreOffice not found at '{LIBREOFFICE_PATH}'."

    output_dir = os.path.dirname(docx_path)
    # Extract alt text mapped to image names
    alt_texts = extract_alt_text_from_docx(docx_path)

    command = [
        LIBREOFFICE_PATH, "--headless", "--convert-to", "html", "--outdir", output_dir, docx_path
    ]
    
    try:
        subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        html_file = os.path.join(output_dir, os.path.basename(docx_path).replace(".docx", ".html"))
        if os.path.exists(html_file):
            responsive_html_file = optimize_html(html_file, alt_texts)
            return f"✅ Conversion successful! Responsive HTML5 saved at: {responsive_html_file}"
        else:
            return "❌ Error: Conversion failed. HTML file not created."
    except subprocess.CalledProcessError as e:
        return f"❌ Error during conversion: {e}"



def optimize_html(html_file, alt_texts):
    """
    Cleans and optimizes the LibreOffice-generated HTML for responsiveness.
    It ensures that each image's 'name' attribute is used to assign the correct alt text,
    that images with names starting with "Shape" get an extra 'img-line' class so they stretch to 100%,
    and that the HTML includes a responsive meta viewport, Bootstrap CSS, and a footer banner.
    Additionally, it removes fixed width and height attributes from <img> tags.

    Args:
        html_file (str): Path to the original HTML file.
        alt_texts (dict): Dictionary mapping image names to alt text.

    Returns:
        str: Path to the cleaned responsive HTML5 file.
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
    :root {
      /* Base font sizes and spacing */
      --font-base: clamp(0.75rem, 1vw + 0.75rem, 1.25rem);
      --font-headline: clamp(1.75rem, 4vw, 2.5rem);
      --spacing-base: clamp(0.5rem, 1vw, 2rem);
      --line-height-base: 1.5;
      /* Dynamic vertical spacing (line-height multiplier) */
      --vertical-spacing: clamp(1.3, 1vw + 1.3, 1.7);
      
      /* Fonts */
      --font-primary: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      --font-secondary: "Segoe UI Black", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    }
    html {
      font-size: 100%;
      line-height: var(--line-height-base);
      font-family: var(--font-primary);
    }
    /* iOS-inspired header styling for the page title */
    header {
      background: rgba(255, 255, 255, 0.85);
      -webkit-backdrop-filter: blur(10px);
      backdrop-filter: blur(10px);
      border-bottom: 1px solid rgba(0,0,0,0.1);
      padding: calc(var(--spacing-base) * 1.5);
      text-align: center;
      box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    }
    header h1 {
      margin: 0;
      font-family: var(--font-secondary);
      font-size: var(--font-headline);
      font-weight: 900;
      letter-spacing: -0.5pt;
      line-height: 1.3;
    }
    @media (min-width: 769px) and (max-width: 1440px) {
      header h1 { line-height: 1.5; }
    }
    @media (max-width: 768px) {
      header h1 {
        font-size: clamp(1.75rem, 2.8vw + 1rem, 2rem);
        line-height: clamp(1.4, 1vw + 1.4, 1.7);
      }
    }
    /* Fluid typography for headings and paragraphs */
    h2 { font-size: clamp(1.5rem, 3.5vw, 2rem); margin-bottom: var(--spacing-base); line-height: var(--vertical-spacing); }
    h3 { font-size: clamp(1.25rem, 3vw, 1.75rem); margin-bottom: var(--spacing-base); line-height: var(--vertical-spacing); }
    h4 { font-size: clamp(1.1rem, 2.5vw, 1.5rem); margin-bottom: var(--spacing-base); line-height: var(--vertical-spacing); }
    h5 { font-size: clamp(1rem, 2vw, 1.25rem); margin-bottom: var(--spacing-base); line-height: var(--vertical-spacing); }
    h6 { font-size: clamp(0.9rem, 1.5vw, 1rem); margin-bottom: var(--spacing-base); line-height: var(--vertical-spacing); }
    p  { font-size: var(--font-base); margin-bottom: var(--spacing-base); line-height: var(--vertical-spacing); }
    /* Responsive images */
    img { max-width: 100% !important; height: auto !important; display: block; }
    .img-line { width: 100% !important; height: auto !important; }
    /* Container and padding */
    body { padding: var(--spacing-base); }
    @media (max-width: 576px) { body { padding: calc(var(--spacing-base) / 2); } }
    /* Responsive tables */
    .table-responsive { overflow-x: auto; }
    .table-responsive table { width: 100%; }
    /* Footer styling */
    footer {
      margin-top: var(--spacing-base);
      padding: var(--spacing-base);
      background-color: #f8f9fa;
      text-align: center;
      font-size: clamp(0.75rem, 1vw, 1rem);
    }
    /* Adapt spacing for extreme screen sizes */
    @media (max-width: 400px) {
      :root { --spacing-base: 0.5rem; }
      header { padding: calc(var(--spacing-base) * 1.2); }
    }
    @media (min-width: 2000px) {
      :root { --spacing-base: 2rem; }
      header { padding: calc(var(--spacing-base) * 1.5); }
    }
  </style>






    <!-- Google tag (gtag.js) -->
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

        # Ensure images have correct alt text, are responsive, and add an extra class for line shapes
        def add_alt_attribute(match):
            img_tag = match.group(0)
            # Attempt to extract the 'name' attribute from the image tag
            name_match = re.search(r'name="([^"]+)"', img_tag)
            # Fallback: use the filename from the src attribute
            src_match = re.search(r'src="([^"]+)"', img_tag)
            image_description = "Illustration from the document"  # Default alt text

            if name_match:
                image_name = name_match.group(1)
                if image_name in alt_texts:
                    image_description = alt_texts[image_name]
                # If the image's name starts with "Shape", add the img-line class
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

            # Update or insert the alt attribute
            if not re.search(r'alt="[^"]*"', img_tag):
                img_tag = re.sub(r'<img', f'<img alt="{image_description}"', img_tag)
            else:
                img_tag = re.sub(r'alt="[^"]*"', f'alt="{image_description}"', img_tag)

            # Ensure Bootstrap's img-fluid class is present for responsiveness
            if 'class=' in img_tag:
                if 'img-fluid' not in img_tag:
                    img_tag = re.sub(r'class="([^"]+)"', lambda m: f'class="{m.group(1)} img-fluid"', img_tag)
            else:
                img_tag = re.sub(r'<img', '<img class="img-fluid"', img_tag)
            return img_tag

        html_content = re.sub(r'<img[^>]+>', add_alt_attribute, html_content)

        # Make tables responsive by wrapping them in a .table-responsive div
        html_content = re.sub(r'(<table[^>]*>.*?</table>)', r'<div class="table-responsive">\1</div>', html_content, flags=re.DOTALL)

        # Add a footer banner with the copyright notice before the closing </body> tag
        footer_banner = """
        <footer>
            <hr>
            <p>© 2025 www.latest2all.com</p>
        </footer>
        """
        html_content = re.sub(r'</body>', footer_banner + '</body>', html_content, flags=re.IGNORECASE)

        # Save the optimized HTML file
        cleaned_html_file = html_file.replace(".html", "_responsive.html")
        with open(cleaned_html_file, "w", encoding="utf-8") as file:
            file.write(html_content)
        return cleaned_html_file

    except Exception as e:
        return f"❌ Error processing HTML file: {e}" 

# 🚀 **User Input for File Path**
#docx_file = input("Enter the full path of the DOCX file: ").strip()
#result = convert_docx_to_html(docx_file)
#print(result)
def main():

    st.title("DOCX to HTML5 Converter")

    uploaded_file = st.file_uploader("Upload a DOCX file", type="docx")

    if uploaded_file is not None:
        st.success("File uploaded successfully!")
        html_content = convert_docx_to_html(uploaded_file)
        st.subheader("Converted HTML5:")
        st.code(html_content, language='html5')

if __name__ == "__main__":
    main()


