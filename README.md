# DOCX2HTML5 Responsive Converter

**DOCX2HTML5 Responsive Converter** is a Python tool that converts Microsoft Word DOCX files into responsive HTML5 documents. It leverages the LibreOffice CLI to perform the initial conversion and then post-processes the generated HTML to inject responsive meta tags, Bootstrap CSS, and custom styling. Additionally, the tool extracts alt text from DOCX image elements and applies them to the corresponding `<img>` tags in the HTML output.

---

## Features

- **DOCX to HTML Conversion**  
  Uses LibreOffice in headless mode to convert DOCX files into HTML.

- **Responsive HTML Output**  
  Injects responsive meta tags, Bootstrap CSS, and custom styles to ensure the resulting HTML adapts to various screen sizes.

- **Alt Text Extraction**  
  Parses the DOCX file’s XML to extract alternative text descriptions for images and updates the HTML accordingly.

- **Clean-up and Optimization**  
  Removes fixed width and height attributes from images and wraps tables in responsive containers.

- **Cross-Platform Compatibility**  
  Easily configurable for both Windows (using `C:\Program Files\LibreOffice\program\soffice.exe`) and Linux (e.g. `/usr/bin/libreoffice`) environments.

---

## Requirements

- **Python 3.x**  
  The code is written in Python and utilizes standard libraries such as:
  - `subprocess`
  - `os`
  - `re`
  - `zipfile`
  - `xml.etree.ElementTree`
  - `tempfile`

- **LibreOffice**  
  LibreOffice must be installed on your system and accessible via the command line in headless mode.
  - **Windows default:** `C:\Program Files\LibreOffice\program\soffice.exe`
  - **Linux default:** `/usr/bin/libreoffice`  
    *(Update the `LIBREOFFICE_PATH` variable in the code as needed.)*

---

## Installation

### 1. Clone the Repository:

```bash
git clone https://github.com/yourusername/docx2html5-responsive-converter.git
cd docx2html5-responsive-converter
```

### 2. Ensure LibreOffice is Installed

#### On Linux
If LibreOffice is not already installed, run:

```bash
sudo apt-get update
sudo apt-get install libreoffice
```

#### On Windows
Download and install LibreOffice from the official website.

### 3. Update the Configuration
In the `libre-docx2html5.py` file, adjust the `LIBREOFFICE_PATH` variable for your operating system:

```python
# For Windows:
LIBREOFFICE_PATH = r"C:\\Program Files\\LibreOffice\\program\\soffice.exe"

# For Linux (uncomment if using Linux):
# LIBREOFFICE_PATH = r"/usr/bin/libreoffice"
```

---

## Usage

### 1. Run the Converter
Open a terminal (or Command Prompt on Windows) in the repository directory and execute:

```bash
python libre-docx2html5.py
```

### 2. Provide Input
When prompted, enter the full path to the DOCX file you wish to convert:

```bash
Enter the full path of the DOCX file: /path/to/your/document.docx
```

### 3. Conversion Output
After the conversion is complete, the script will display a message similar to:

```bash
✅ Conversion successful! Responsive HTML5 saved at: /path/to/your/document_responsive.html
```

You can now open the resulting HTML file in your browser.

---

## Customization

### Responsive Styling
Modify the CSS within the `responsive_head` variable in the code to adjust fonts, spacing, and other styles as desired.

### Alt Text Extraction
The tool extracts alt text from `<wp:docPr>` elements in the DOCX file. You can adjust the behavior in the `extract_alt_text_from_docx` function if necessary.

---

## Troubleshooting

### LibreOffice Not Found
Verify that the `LIBREOFFICE_PATH` variable is correctly set for your system.

### Conversion Failures
Check the console output for error messages during conversion. Ensure that:

- The DOCX file is not corrupted.
- The LibreOffice CLI is functioning correctly in headless mode.

---

## Contributing

Contributions are welcome! Feel free to fork this repository, improve the code, and open pull requests.

For major changes, please open an issue first to discuss your ideas.

---

## License

This project is licensed under the **GNU General Public License v3.0**.

You are free to use, modify, and distribute this software under the terms of the GNU GPL v3.0.

For more details, please see the LICENSE file or visit [GNU GPL v3.0](https://www.gnu.org/licenses/gpl-3.0.html).

---

## Contact

For questions, feedback, or support, please open an issue in this repository or contact the maintainer:

🌍 [www.latest2all.com](https://www.latest2all.com)  
📧 reddyapuru@gmail.com

