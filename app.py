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
<header style="background-color: #f5f5f5; padding: 20px 40px; text-align: center; border-bottom: 2px solid #ccc;">
  <h1 style="font-size: 2.5rem; margin-bottom: 10px;">Welcome to Latest2All DOCX2HTML5 Converter</h1>
  <p style="font-size: 1.2rem; color: #333; margin-bottom: 15px;">
    Effortlessly convert your DOCX files into responsive HTML.
  </p>
  <p style="font-size: 1rem; color: #555; max-width: 800px; margin: 0 auto;">
    To use this tool, simply upload your DOCX file using the form below. Our converter will extract and optimize the content—including images—and package everything into a ZIP file ready for download. Your package will be available immediately and will be automatically deleted after a short period.
    <p>docx2html5 online responsive converter sponsored by <a href="https://www.latest2all.com" target="_blank">www.latest2all.com</a></p>
    <p>&copy; 2025</p>
  </p>
</header>
    <h1>DOCX to Responsive HTML Converter</h1>
    <p>Upload a DOCX file to convert it to responsive HTML along with its images packaged in a ZIP file.
    <br>(The package will be deleted automatically after 10 minutes.)</p>
    <form method="post" enctype="multipart/form-data">
      <input type="file" name="docx_file" accept=".docx" required>
      <br><br>
      <input type="submit" value="Convert" class="upload-btn">
    </form>
  </body>
</html>
"""

RESULT_PAGE = """
<!doctype html>
<html>
  <head>
    <title>Conversion Result</title>
    <style>
      body { font-family: Arial, sans-serif; margin: 40px; }
      .btn {
          font-size: 20px;
          padding: 12px 24px;
          background-color: #4CAF50;
          color: white;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          text-decoration: none;
      }
      .btn:hover {
          background-color: #45a049;
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
  <div style="max-width: 800px; margin: 40px auto; text-align: center;">
    <h1 style="margin-bottom: 30px;">Conversion Successful!</h1>
    <p style="margin-bottom: 20px;">Conversion Time: {{ conversion_time }} seconds.</p>
    <p style="margin-bottom: 30px;">Your package is ready for download. (It will be deleted automatically after 10 minutes.)</p>
    <div>
      <a class="btn" style="margin-right: 20px; padding: 12px 24px; font-size: 20px;" href="{{ url_for('download_file') }}">Download ZIP Package</a>
      <a class="btn" style="padding: 12px 24px; font-size: 20px;" href="{{ url_for('clear') }}">Clear and Start Over</a>
    </div>
  </div>
  <footer style="margin-top: 40px; text-align: center; font-size: 14px; color: #666;">
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
