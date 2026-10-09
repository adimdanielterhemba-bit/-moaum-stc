from flask import Flask, request, render_template_string, send_file
import pandas as pd
import os

app = Flask(__name__)

HTML = """
<!DOCTYPE html>
<html>
<head>
<title>MOAUM STC - School Management</title>
<style>
body{font-family:Arial; background:#f4f6f9; padding:20px}
.card{background:white; padding:30px; border-radius:10px; max-width:800px; margin:auto; box-shadow:0 2px 10px #ccc}
h1{color:#0b3d91; text-align:center}
.btn{background:#0b3d91; color:white; padding:12px 20px; border:none; border-radius:5px; cursor:pointer}
input{padding:10px; width:100%; margin:10px 0}
table{width:100%; border-collapse:collapse; margin-top:20px}
th,td{border:1px solid #ccc; padding:8px; text-align:left}
th{background:#0b3d91; color:white}
</style>
</head>
<body>
<div class="card">
<h1>MOAUM STC School Management</h1>
<p style="text-align:center">Upload Student Excel File (Name, Class, Subject, Score)</p>
<form method="post" enctype="multipart/form-data">
<input type="file" name="file" accept=".xlsx,.xls" required>
<button class="btn" type="submit">Upload & Process</button>
</form>
{% if tables %}
<h2>Results Preview</h2>
{{ tables|safe }}
<br><br>
<a href="/download"><button class="btn">Download Processed Excel</button></a>
{% endif %}
</div>
</body>
</html>
"""

processed_path = "/tmp/processed_results.xlsx"
last_html = ""

@app.route("/", methods=["GET", "POST"])
def index():
    global last_html
    tables = None
    if request.method == "POST":
        f = request.files.get("file")
        if f:
            df = pd.read_excel(f)
            # Simple processing - add Grade and Remark
            def get_grade(score):
                try:
                    s = float(score)
                    if s >= 70: return "A"
                    elif s >= 60: return "B"
                    elif s >= 50: return "C"
                    elif s >= 45: return "D"
                    elif s >= 40: return "E"
                    else: return "F"
                except:
                    return ""
            def get_remark(score):
                try:
                    s = float(score)
                    return "Pass" if s >= 40 else "Fail"
                except:
                    return ""
            # try to find score column
            score_col = None
            for col in df.columns:
                if "score" in str(col).lower() or "mark" in str(col).lower():
                    score_col = col
                    break
            if score_col is None and len(df.columns) >= 2:
                score_col = df.columns[-1]
            if score_col:
                df["Grade"] = df[score_col].apply(get_grade)
                df["Remark"] = df[score_col].apply(get_remark)
            df.to_excel(processed_path, index=False)
            tables = df.head(100).to_html(classes="table", index=False)
            last_html = tables
    return render_template_string(HTML, tables=tables)

@app.route("/download")
def download():
    if os.path.exists(processed_path):
        return send_file(processed_path, as_attachment=True)
    return "No file yet"

if __name__ == "__main__":
    app.run(debug=True)
