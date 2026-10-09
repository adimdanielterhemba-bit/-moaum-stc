from flask import Flask, request, render_template_string, send_file
import pandas as pd
import os

app = Flask(__name__)

HTML_PAGE = """
<!DOCTYPE html>
<html>
<head>
<title>MOAUM STC - School Management</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body{font-family:Arial; background:#f4f6f9; padding:20px; margin:0}
.card{background:white; padding:30px; border-radius:10px; max-width:900px; margin:auto; box-shadow:0 2px 15px rgba(0,0,0,0.1)}
h1{color:#0b3d91; text-align:center; margin-bottom:5px}
.sub{text-align:center; color:#555; margin-bottom:20px}
.btn{background:#0b3d91; color:white; padding:12px 25px; border:none; border-radius:5px; cursor:pointer; font-size:16px}
.btn:hover{background:#09306f}
input[type=file]{padding:10px; width:100%; margin:15px 0; border:1px solid #ddd; border-radius:5px}
table{width:100%; border-collapse:collapse; margin-top:20px; font-size:14px}
th,td{border:1px solid #ddd; padding:8px; text-align:left}
th{background:#0b3d91; color:white}
tr:nth-child(even){background:#f9f9f9}
.footer{text-align:center; margin-top:30px; color:#888; font-size:12px}
</style>
</head>
<body>
<div class="card">
<h1>MOAUM STC - Secondary Technical College</h1>
<p class="sub">Student Results Processing System - Upload Excel File</p>
<p class="sub"><b>Excel Format:</b> Name | Class | Subject | Score (or any column with scores)</p>
<form method="post" enctype="multipart/form-data">
<input type="file" name="file" accept=".xlsx,.xls" required>
<div style="text-align:center"><button class="btn" type="submit">Upload & Process Results</button></div>
</form>
{% if tables %}
<hr>
<h2>Preview (Top 100 Rows)</h2>
<div style="overflow-x:auto">{{ tables|safe }}</div>
<br><br>
<div style="text-align:center">
<a href="/download"><button class="btn">Download Processed Excel File</button></a>
</div>
{% endif %}
<div class="footer">MOAUM STC Management System - Powered by Flask</div>
</div>
</body>
</html>
"""

OUTPUT_FILE = "/tmp/moaum_processed.xlsx"

@app.route("/", methods=["GET", "POST"])
def index():
    tables = None
    if request.method == "POST":
        file = request.files.get("file")
        if file and file.filename:
            try:
                df = pd.read_excel(file)
                if df.empty:
                    tables = "<p style='color:red'>Excel file is empty!</p>"
                else:
                    # Find score column
                    score_col = None
                    for c in df.columns:
                        cl = str(c).lower()
                        if "score" in cl or "mark" in cl or "total" in cl:
                            score_col = c
                            break
                    if score_col is None:
                        score_col = df.columns[-1]
                    
                    def grade(s):
                        try:
                            v = float(s)
                            if v >= 70: return "A"
                            if v >= 60: return "B"
                            if v >= 55: return "C"
                            if v >= 50: return "D"
                            if v >= 40: return "E"
                            return "F"
                        except: return ""
                    def remark(s):
                        try:
                            v = float(s)
                            return "PASS" if v >= 40 else "FAIL"
                        except: return ""
                    
                    df["GRADE"] = df[score_col].apply(grade)
                    df["REMARK"] = df[score_col].apply(remark)
                    df.to_excel(OUTPUT_FILE, index=False)
                    tables = df.head(100).to_html(index=False)
            except Exception as e:
                tables = f"<p style='color:red'>Error: {str(e)}</p>"
    return render_template_string(HTML_PAGE, tables=tables)

@app.route("/download")
def download():
    if os.path.exists(OUTPUT_FILE):
        return send_file(OUTPUT_FILE, as_attachment=True, download_name="MOAUM_Processed_Results.xlsx")
    return "No file processed yet. Please upload first."

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
