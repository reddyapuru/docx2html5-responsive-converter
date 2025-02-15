from flask import Flask, request, send_file, render_template_string, flash, redirect, url_for, session
from werkzeug.utils import secure_filename
import os
import tempfile
import datetime

# Import your conversion function from libre_docx2html5.py
from libre_docx2html5 import convert_docx_to_html

app = Flask(__name__)
app.secret_key = "supersecretkey"

ALLOWED_EXTENSIONS = {"docx"}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# HTML template for the upload form with a "please wait" overlay
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
      #loading {
          display: none;
          text-align: center;
          font-size: 20px;
          color: #555;
          margin-top: 20px;
      }
      header {
          background-color: #f5f5f5;
          padding: 20px 40px;
          text-align: center;
          border-bottom: 2px solid #ccc;
      }
      header h1 {
          font-size: 2.5rem;
          margin-bottom: 10px;
      }
      header p {
          font-size: 1.2rem;
          color: #333;
          margin-bottom: 15px;
      }
      header a {
          color: #007BFF;
          text-decoration: none;
      }
      header a:hover {
          text-decoration: underline;
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
    <script>
      function showLoading() {
          document.getElementById('loading').style.display = 'block';
      }
    </script>
  </head>
  <body>
<header>
  <h1>Welcome to Latest2All DOCX2HTML5 Converter</h1>
  <p>Effortlessly convert your DOCX files into a fully responsive HTML file with optimized images.</p>
  <p>
    To use this tool, simply upload your DOCX file below. Our converter extracts and optimizes the content—including images—and packages everything into a ZIP file. 
    The final output is a responsive HTML file designed to look great on any device, ready for immediate download. Your package will be automatically deleted after 10 minutes.
  </p>
  <p>docx2html5 online responsive converter sponsored by <a href="https://www.latest2all.com" target="_blank">www.latest2all.com</a> &copy; 2025</p>
</header>
    <h1>DOCX to Responsive HTML Converter</h1>
    <p>
      Convert your DOCX files into fully responsive HTML with images packaged in a ZIP file.
      <br>
      (Your files are processed securely and will be automatically deleted after 10 minutes.)
    </p>
    <form method="post" enctype="multipart/form-data" onsubmit="showLoading()">
      <input type="file" name="docx_file" accept=".docx" required>
      <br><br>
      <input type="submit" value="Convert" class="upload-btn">
    </form>
    <div id="loading">Conversion in progress... Please wait.</div>
  </body>
</html>
"""

# HTML template for the result page
RESULT_PAGE = """
<!doctype html>
<html>
  <head>
    <title>Conversion Result</title>
    <style>
      body { font-family: Arial, sans-serif; margin: 40px; text-align: center; }
      .btn {
          font-size: 20px;
          padding: 12px 24px;
          background-color: #4CAF50;
          color: white;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          text-decoration: none;
          margin: 10px;
      }
      .btn:hover {
          background-color: #45a049;
      }
      footer {
          margin-top: 40px;
          text-align: center;
          font-size: 14px;
          color: #666;
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
  <body>
<div style="max-width: 800px; margin: 40px auto;">
  <h1 style="margin-bottom: 30px;">Conversion Successful!</h1>
  <p style="margin-bottom: 20px;">Conversion Time: {{ conversion_time }} seconds.</p>
  <p style="margin-bottom: 30px;">Your package is ready for download. (It will be deleted automatically after 10 minutes.)</p>
  <div style="display: flex; flex-direction: column; gap: 1rem;">
    <a class="btn" href="{{ url_for('download_file') }}" style="display: block;">Download ZIP Package</a>
    <a class="btn" href="{{ url_for('clear') }}" style="display: block;">Clear and Start Over</a>
  </div>
</div>

    <footer>
      <p>docx2html5 online responsive converter sponsored by <a href="https://www.latest2all.com" target="_blank">www.latest2all.com</a></p>
      <p>&copy; 2025</p>
    </footer>
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
            
            # Measure conversion time
            start_time = datetime.datetime.now()
            zip_path = convert_docx_to_html(file_path)
            end_time = datetime.datetime.now()
            conversion_time = (end_time - start_time).total_seconds()
            
            if zip_path.startswith("❌"):
                flash(zip_path)
                return redirect(request.url)
            else:
                # Store result in session and redirect to the result page
                session["zip_path"] = zip_path
                session["conversion_time"] = conversion_time
                return redirect(url_for("result"))
    return render_template_string(UPLOAD_FORM)

@app.route("/result", methods=["GET"])
def result():
    if "zip_path" not in session or "conversion_time" not in session:
        flash("No conversion result found.")
        return redirect(url_for("index"))
    return render_template_string(RESULT_PAGE, conversion_time=session["conversion_time"])

@app.route("/download", methods=["GET"])
def download_file():
    if "zip_path" not in session:
        flash("No file to download.")
        return redirect(url_for("index"))
    return send_file(session["zip_path"], as_attachment=True, download_name=os.path.basename(session["zip_path"]))

@app.route("/clear", methods=["GET"])
def clear():
    # Clear session variables and redirect to the upload page
    session.pop("zip_path", None)
    session.pop("conversion_time", None)
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)
