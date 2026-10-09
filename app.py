from flask import Flask, request, render_template_string, send_file, redirect, url_for, session
import pandas as pd
import os
from functools import wraps

app = Flask(__name__)
app.secret_key = "moaum-stc-secret-2026"

# In-memory storage
students_db = []
teachers_db = []
OUTPUT_FILE = "/tmp/moaum_processed.xlsx"

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'logged_in' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

BASE_STYLE = """
<style>
body{font-family:Arial; background:#f0f2f5; margin:0; padding:0}
.header{background:#0b3d91; color:white; padding:15px 20px; display:flex; justify-content:space-between; align-items:center}
.header h1{margin:0; font-size:20px}
.header a{color:white; text-decoration:none; background:#ff4757; padding:8px 15px; border-radius:5px; margin-left:10px}
.nav{background:white; padding:10px 20px; box-shadow:0 2px 5px rgba(0,0,0,0.1); display:flex; gap:10px; flex-wrap:wrap}
.nav a{text-decoration:none; background:#eef2ff; color:#0b3d91; padding:8px 15px; border-radius:20px; font-size:14px; font-weight:bold}
.nav a:hover{background:#0b3d91; color:white}
.card{background:white; padding:25px; border-radius:10px; max-width:1100px; margin:20px auto; box-shadow:0 2px 10px rgba(0,0,0,0.08)}
.btn{background:#0b3d91; color:white; padding:10px 20px; border:none; border-radius:5px; cursor:pointer}
.btn:hover{background:#09306f}
.grid{display:grid; grid-template-columns:repeat(auto-fit, minmax(200px,1fr)); gap:15px; margin-top:20px}
.menu-card{background:white; border:2px solid #eef2ff; padding:20px; border-radius:10px; text-align:center; transition:0.3s; text-decoration:none; color:#333; display:block}
.menu-card:hover{border-color:#0b3d91; transform:translateY(-3px); box-shadow:0 5px 15px rgba(0,0,0,0.1)}
.menu-card .icon{font-size:40px; margin-bottom:10px}
table{width:100%; border-collapse:collapse; margin-top:15px}
th,td{border:1px solid #ddd; padding:10px; text-align:left; font-size:14px}
th{background:#0b3d91; color:white}
input,select{padding:10px; width:100%; margin:8px 0; border:1px solid #ddd; border-radius:5px; box-sizing:border-box}
.login-box{max-width:400px; margin:80px auto}
</style>
"""

LOGIN_PAGE = BASE_STYLE + """
<div class="login-box card">
<h2 style="text-align:center; color:#0b3d91">MOAUM STC Login</h2>
<p style="text-align:center">Secondary Technical College Portal</p>
<form method="post">
<input type="text" name="username" placeholder="Username" required>
<input type="password" name="password" placeholder="Password" required>
<button class="btn" style="width:100%; margin-top:10px" type="submit">Login</button>
</form>
<p style="font-size:12px; color:#888; text-align:center; margin-top:15px">Demo: admin / admin123</p>
{% if error %}<p style="color:red; text-align:center">{{ error }}</p>{% endif %}
</div>
"""

DASHBOARD_PAGE = BASE_STYLE + """
<div class="header"><h1>MOAUM STC - School Portal</h1><div><span>Admin</span> <a href="/logout">Logout</a></div></div>
<div class="nav">
<a href="/dashboard">🏠 Dashboard</a>
<a href="/students">👨‍🎓 Students</a>
<a href="/add-student">➕ Add Student</a>
<a href="/teachers">👨‍🏫 Teachers</a>
<a href="/results">📊 Results Processing</a>
<a href="/attendance">📅 Attendance</a>
<a href="/fees">💰 Fees</a>
<a href="/subjects">📚 Subjects</a>
</div>
<div class="card">
<h2>Welcome to MOAUM STC Dashboard</h2>
<p>Select a module to manage your school operations</p>
<div class="grid">
<a href="/students" class="menu-card"><div class="icon">👨‍🎓</div><h3>Student List</h3><p>{{ students_count }} Students</p></a>
<a href="/add-student" class="menu-card"><div class="icon">➕</div><h3>Add Student</h3><p>Register new student</p></a>
<a href="/teachers" class="menu-card"><div class="icon">👨‍🏫</div><h3>Teachers</h3><p>{{ teachers_count }} Teachers</p></a>
<a href="/results" class="menu-card"><div class="icon">📊</div><h3>Results Processing</h3><p>Upload & grade results</p></a>
<a href="/attendance" class="menu-card"><div class="icon">📅</div><h3>Attendance</h3><p>Mark attendance</p></a>
<a href="/fees" class="menu-card"><div class="icon">💰</div><h3>Fees Management</h3><p>Track payments</p></a>
<a href="/subjects" class="menu-card"><div class="icon">📚</div><h3>Subjects</h3><p>Manage subjects</p></a>
<a href="/reports" class="menu-card"><div class="icon">📈</div><h3>Reports</h3><p>View reports</p></a>
</div>
</div>
"""

RESULTS_PAGE = BASE_STYLE + """
<div class="header"><h1>MOAUM STC</h1><div><a href="/dashboard">Dashboard</a> <a href="/logout">Logout</a></div></div>
<div class="nav">
<a href="/dashboard">🏠 Dashboard</a>
<a href="/results">📊 Results Processing</a>
</div>
<div class="card">
<h2>📊 Student Results Processing - Upload Excel</h2>
<p><b>Excel Format:</b> Name | Class | Subject | Score</p>
<form method="post" enctype="multipart/form-data">
<input type="file" name="file" accept=".xlsx,.xls" required>
<button class="btn" type="submit">Upload & Process Results</button>
</form>
{% if tables %}
<hr>
<h3>Preview (Top 100)</h3>
<div style="overflow-x:auto">{{ tables|safe }}</div>
<br><br>
<a href="/download"><button class="btn">Download Processed Excel</button></a>
{% endif %}
</div>
"""

SIMPLE_PAGE = BASE_STYLE + """
<div class="header"><h1>MOAUM STC</h1><div><a href="/dashboard">Dashboard</a> <a href="/logout">Logout</a></div></div>
<div class="nav">
<a href="/dashboard">🏠 Dashboard</a>
<a href="/students">👨‍🎓 Students</a>
<a href="/add-student">➕ Add Student</a>
<a href="/teachers">👨‍🏫 Teachers</a>
<a href="/results">📊 Results</a>
<a href="/attendance">📅 Attendance</a>
<a href="/fees">💰 Fees</a>
<a href="/subjects">📚 Subjects</a>
</div>
<div class="card">
{{ content|safe }}
</div>
"""

@app.route("/")
def home():
    if 'logged_in' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route("/login", methods=["GET","POST"])
def login():
    error = None
    if request.method == "POST":
        u = request.form.get("username")
        p = request.form.get("password")
        if u == "admin" and p == "admin123":
            session['logged_in'] = True
            return redirect(url_for('dashboard'))
        else:
            error = "Invalid username or password! Use admin / admin123"
    return render_template_string(LOGIN_PAGE, error=error)

@app.route("/logout")
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('login'))

@app.route("/dashboard")
@login_required
def dashboard():
    return render_template_string(DASHBOARD_PAGE, students_count=len(students_db), teachers_count=len(teachers_db))

@app.route("/results", methods=["GET","POST"])
@login_required
def results():
    tables = None
    if request.method == "POST":
        file = request.files.get("file")
        if file and file.filename:
            try:
                df = pd.read_excel(file)
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
    return render_template_string(RESULTS_PAGE, tables=tables)

@app.route("/students")
@login_required
def students():
    if not students_db:
        content = "<h2>👨‍🎓 Student List</h2><p>No students registered yet. <a href='/add-student'>Add first student</a></p>"
    else:
        rows = ""
        for i,s in enumerate(students_db,1):
            rows += f"<tr><td>{i}</td><td>{s['name']}</td><td>{s['class']}</td><td>{s['reg']}</td><td>{s['gender']}</td></tr>"
        content = f"<h2>👨‍🎓 Student List - {len(students_db)} Students</h2><table><tr><th>#</th><th>Name</th><th>Class</th><th>Reg No</th><th>Gender</th></tr>{rows}</table>"
    return render_template_string(SIMPLE_PAGE, content=content)

@app.route("/add-student", methods=["GET","POST"])
@login_required
def add_student():
    if request.method == "POST":
        students_db.append({
            "name": request.form.get("name"),
            "class": request.form.get("class"),
            "reg": request.form.get("reg"),
            "gender": request.form.get("gender")
        })
        content = f"<h2>✅ Student Added!</h2><p>{request.form.get('name')} added successfully.</p><a href='/students'><button class='btn'>View All Students</button></a> <a href='/add-student'><button class='btn'>Add Another</button></a>"
        return render_template_string(SIMPLE_PAGE, content=content)
    content = """
    <h2>➕ Add New Student</h2>
    <form method="post">
    <input type="text" name="name" placeholder="Full Name" required>
    <input type="text" name="class" placeholder="Class e.g JSS2, SS1" required>
    <input type="text" name="reg" placeholder="Registration Number" required>
    <select name="gender" required><option value="">Select Gender</option><option>Male</option><option>Female</option></select>
    <button class="btn" type="submit">Register Student</button>
    </form>
    """
    return render_template_string(SIMPLE_PAGE, content=content)

@app.route("/teachers")
@login_required
def teachers():
    content = "<h2>👨‍🏫 Teachers</h2><p>Teachers module - Add teachers list here. Currently demo.</p><table><tr><th>Name</th><th>Subject</th><th>Phone</th></tr><tr><td>Mr. John</td><td>Mathematics</td><td>080...</td></tr><tr><td>Mrs. Mary</td><td>English</td><td>080...</td></tr></table>"
    return render_template_string(SIMPLE_PAGE, content=content)

@app.route("/attendance")
@login_required
def attendance():
    content = "<h2>📅 Attendance</h2><p>Attendance marking module.</p><p>Select Class: <select><option>JSS1</option><option>JSS2</option><option>SS1</option></select> <button class='btn'>Load Students</button></p>"
    return render_template_string(SIMPLE_PAGE, content=content)

@app.route("/fees")
@login_required
def fees():
    content = "<h2>💰 Fees Management</h2><p>Track school fees payments.</p><table><tr><th>Student</th><th>Class</th><th>Amount Paid</th><th>Balance</th><th>Status</th></tr><tr><td>John Doe</td><td>JSS2</td><td>₦15,000</td><td>₦5,000</td><td>Partial</td></tr></table>"
    return render_template_string(SIMPLE_PAGE, content=content)

@app.route("/subjects")
@login_required
def subjects():
    content = "<h2>📚 Subjects</h2><table><tr><th>Subject</th><th>Code</th><th>Class</th></tr><tr><td>Mathematics</td><td>MTH101</td><td>All</td></tr><tr><td>English</td><td>ENG101</td><td>All</td></tr><tr><td>Basic Tech</td><td>BTC</td><td>JSS</td></tr><tr><td>Physics</td><td>PHY</td><td>SSS</td></tr></table>"
    return render_template_string(SIMPLE_PAGE, content=content)

@app.route("/reports")
@login_required
def reports():
    content = "<h2>📈 Reports</h2><p>Generate reports.</p><button class='btn'>Generate Class Result Report</button> <button class='btn'>Generate Fees Report</button>"
    return render_template_string(SIMPLE_PAGE, content=content)

@app.route("/download")
@login_required
def download():
    if os.path.exists(OUTPUT_FILE):
        return send_file(OUTPUT_FILE, as_attachment=True, download_name="MOAUM_Processed_Results.xlsx")
    return "No file yet."

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
