from flask import Flask, request, send_file, render_template_string, flash, redirect, url_for
from werkzeug.utils import secure_filename
import os
import tempfile

# Import your conversion function from libre_docx2html5.py
from libre_docx2html5 import convert_docx_to_html

app = Flask(__name__)
app.secret_key = "supersecretkey"

ALLOWED_EXTENSIONS = {"docx"}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

UPLOAD_FORM = """
<!doctype html>
<html>
  <head>
    <title>DOCX to Responsive HTML Converter</title>
    <style>
      body { font-family: Arial, sans-serif; margin: 40px; }
      .upload-btn {
          font-size: 20px;
          padding: 12px 24px;
          background-color: #4CAF50;
          color: white;
          border: none;
          border-radius: 4px;
          cursor: pointer;
      }
      .upload-btn:hover {
          background-color: #45a049;
      }
    </style>
  </head>
  <body>
    <h1>DOCX to Responsive HTML Converter</h1>
    <p>Upload a DOCX file to convert it to responsive HTML along with its images packaged in a ZIP file. (The package will be deleted automatically after 10 minutes.)</p>
    <form method="post" enctype="multipart/form-data">
      <input type="file" name="docx_file" accept=".docx" required>
      <br><br>
      <input type="submit" value="Convert" class="upload-btn">
    </form>
  </body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "docx_file" not in request.files:
            flash("No file part")
            return redirect(request.url)
        file = request.files["docx_file"]
        if file.filename == "":
            flash("No selected file")
            return redirect(request.url)
        if file and allowed_file(file.filename):
            # Save the uploaded file to a temporary directory
            upload_dir = tempfile.mkdtemp()
            file_path = os.path.join(upload_dir, secure_filename(file.filename))
            file.save(file_path)
            # Convert the DOCX file using your conversion function.
            zip_path = convert_docx_to_html(file_path)
            if zip_path.startswith("❌"):
                flash(zip_path)
                return redirect(request.url)
            else:
                # send_file sets appropriate headers so the browser immediately downloads the file.
                return send_file(zip_path, as_attachment=True, download_name=os.path.basename(zip_path))
    return render_template_string(UPLOAD_FORM)

if __name__ == "__main__":
    app.run(debug=True)
