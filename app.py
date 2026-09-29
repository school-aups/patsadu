from flask import Flask, render_template_string, request, redirect, url_for, send_file
import sqlite3
import pandas as pd
from datetime import datetime
import io
import os

app = Flask(__name__)

def init_db():
    conn = sqlite3.connect('assets.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS assets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE,
            name TEXT,
            category TEXT,
            purchase_date TEXT,
            price REAL,
            life_years INTEGER,
            is_low_value INTEGER
        )
    ''')
    conn.commit()
    conn.close()

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="UTF-8">
    <title>ระบบทะเบียนคุมครุภัณฑ์ - โรงเรียนอนุบาลอุทุมพรพิสัย</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Sarabun', sans-serif; background-color: #f8f9fa; }
        .school-logo { width: 80px; height: 80px; object-fit: contain; }
        .card { border: none; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
        h2, h3, h4, h5, .table th { font-family: 'Sarabun', sans-serif; font-weight: 600; }
        @media print { .no-print { display: none !important; } body { background-color: white !important; } }
    </style>
</head>
<body class="container mt-4 mb-5">
    <div class="d-flex align-items-center mb-4 p-3 bg-white rounded shadow-sm border no-print">
        <img src="https://school.aups.ac.th/images/logo.png" alt="ตราโรงเรียน" class="school-logo me-3" onerror="this.src='https://via.placeholder.com/80'">
        <div>
            <h3 class="mb-1 text-primary fw-bold">โรงเรียนอนุบาลอุทุมพรพิสัย</h3>
            <h5 class="text-secondary mb-0">ระบบทะเบียนคุมครุภัณฑ์และคำนวณค่าเสื่อมราคา (งานพัสดุ)</h5>
        </div>
    </div>
    
    {% if view_mode == 'schedule' %}
        <div class="card mb-4 p-4">
            <div class="text-center mb-4">
                <h4 class="fw-bold">ทะเบียนคุมทรัพย์สินและคำนวณค่าเสื่อมราคา</h4>
                <h5 class="text-secondary">โรงเรียนอนุบาลอุทุมพรพิสัย (งานพัสดุ)</h5>
            </div>
            <hr>
            <div class="row mb-3 fs-5">
                <div class="col-md-6">
                    <p class="mb-1"><strong>รหัสครุภัณฑ์:</strong> {{ asset.code }}</p>
                    <p class="mb-1"><strong>ชื่อรายการ:</strong> {{ asset.name }}</p>
                    <p class="mb-1"><strong>ประเภท:</strong> {{ asset.category }}</p>
                </div>
                <div class="col-md-6">
                    <p class="mb-1"><strong>วันที่จัดซื้อ:</strong> {{ asset.purchase_date }}</p>
                    <p class="mb-1"><strong>ราคาทุน:</strong> {{ "{:,.2f}".format(asset.price) }} บาท</p>
                    <p class="mb-1"><strong>อายุการใช้งาน:</strong> {{ asset.life_years }} ปี</p>
                </div>
            </div>
            <p class="fs-5"><strong>สถานะ:</strong> 
                <span class="badge {% if asset.is_low_value %}bg-warning text-dark{% else %}bg-success{% endif %}">
                    {{ "ต่ำกว่าเกณฑ์ (ไม่คิดค่าเสื่อม)" if asset.is_low_value else "ปกติ (วิธีเส้นตรง)" }}
                </span>
            </p>

            <table class="table table-bordered table-striped mt-3 align-middle fs-5">
                <thead class="table-dark text-center">
                    <tr>
                        <th>ปีพุทธศักราช (พ.ศ.)</th>
                        <th>มูลค่าต้นงวด (บาท)</th>
                        <th>ค่าเสื่อมราคาประจำปี (บาท)</th>
                        <th>ค่าเสื่อมราคาสะสม (บาท)</th>
                        <th>มูลค่าสุทธิปลายงวด (บาท)</th>
                    </tr>
                </thead>
                <tbody>
                    {% for row in schedule %}
                    <tr>
                        <td class="text-center">{{ row.year_be }}</td>
                        <td class="text-end">{{ "{:,.2f}".format(row.beginning_book_value) }}</td>
                        <td class="text-end">{{ "{:,.2f}".format(row.depreciation) }}</td>
                        <td class="text-end">{{ "{:,.2f}".format(row.accumulated) }}</td>
                        <td class="text-end">{{ "{:,.2f}".format(row.ending_book_value) }}</td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>

            <div class="mt-4 d-flex justify-content-between no-print">
                <a href="/" class="btn btn-secondary px-4">กลับหน้าหลัก</a>
                <div>
                    <a href="/schedule/export_excel/{{ asset.id }}" class="btn btn-success me-2 px-4">ดาวน์โหลด Excel</a>
                    <button type="button" class="btn btn-danger px-4" onclick="window.print()">พิมพ์ PDF / เอกสาร</button>
                </div>
            </div>
        </div>
    {% else %}
        <div class="card mb-4 p-4">
            <h4 class="mb-3 text-primary fw-bold">เพิ่มครุภัณฑ์ใหม่</h4>
            <form action="/add" method="POST" class="row g-3">
                <div class="col-md-3">
                    <label class="form-label fw-semibold">รหัสครุภัณฑ์</label>
                    <input type="text" class="form-control" name="code" required>
                </div>
                <div class="col-md-4">
                    <label class="form-label fw-semibold">ชื่อรายการ</label>
                    <input type="text" class="form-control" name="name" required>
                </div>
                <div class="col-md-3">
                    <label class="form-label fw-semibold">ประเภท</label>
                    <input type="text" class="form-control" name="category" required>
                </div>
                <div class="col-md-2">
                    <label class="form-label fw-semibold">วันที่จัดซื้อ</label>
                    <input type="date" class="form-control" name="purchase_date" required>
                </div>
                <div class="col-md-3">
                    <label class="form-label fw-semibold">ราคาทุน (บาท)</label>
                    <input type="number" step="0.01" class="form-control" name="price" required>
                </div>
                <div class="col-md-3">
                    <label class="form-label fw-semibold">อายุการใช้งาน (ปี)</label>
                    <input type="number" class="form-control" name="life_years" required>
                </div>
                <div class="col-md-3 d-flex align-items-end">
                    <button type="submit" class="btn btn-primary w-100 fw-semibold">บันทึกข้อมูล</button>
                </div>
            </form>
        </div>

        <div class="row mb-3 align-items-end">
            <form method="GET" action="/" class="row g-3 col-md-10">
                <div class="col-md-4">
                    <label class="form-label fw-semibold">ค้นหา (ชื่อ หรือ รหัส)</label>
                    <input type="text" class="form-control" name="search" value="{{ search }}">
                </div>
                <div class="col-md-4">
                    <label class="form-label fw-semibold">กรองตามประเภท</label>
                    <select name="category" class="form-select">
                        <option value="">-- ทุกประเภท --</option>
                        {% for cat in categories %}
                        <option value="{{ cat }}" {% if selected_cat == cat %}selected{% endif %}>{{ cat }}</option>
                        {% endfor %}
                    </select>
                </div>
                <div class="col-md-4 d-flex align-items-end">
                    <button type="submit" class="btn btn-secondary me-2 px-4">ค้นหา</button>
                    <a href="/" class="btn btn-outline-secondary px-3">รีเซ็ต</a>
                </div>
            </form>
            <div class="col-md-2 text-end">
                <a href="/export_excel" class="btn btn-success w-100 fw-semibold">ดาวน์โหลด Excel</a>
            </div>
        </div>

        <table class="table table-bordered table-striped align-middle bg-white shadow-sm rounded">
            <thead class="table-dark">
                <tr>
                    <th>รหัส</th>
                    <th>ชื่อรายการ</th>
                    <th>ประเภท</th>
                    <th>ราคาทุน</th>
                    <th>สถานะ</th>
                    <th>ค่าเสื่อมสะสม (ปีปัจจุบัน)</th>
                    <th>มูลค่าสุทธิ</th>
                    <th>จัดการ</th>
                </tr>
            </thead>
            <tbody>
                {% for item in assets %}
                <tr>
                    <td>{{ item.code }}</td>
                    <td>{{ item.name }}</td>
                    <td>{{ item.category }}</td>
                    <td>{{ "{:,.2f}".format(item.price) }}</td>
                    <td>
                        <span class="badge {% if item.status == 'ต่ำกว่าเกณฑ์' %}bg-warning text-dark{% else %}bg-success{% endif %}">
                            {{ item.status }}
                        </span>
                    </td>
                    <td>{{ "{:,.2f}".format(item.acc_dep) }}</td>
                    <td>{{ "{:,.2f}".format(item.net_val) }}</td>
                    <td>
                        <a href="/schedule/{{ item.id }}" class="btn btn-info btn-sm text-white fw-semibold">ดูค่าเสื่อมรายปี</a>
                    </td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    {% endif %}
</body>
</html>
'''

@app.route('/')
def index():
    search_query = request.args.get('search', '')
    category_filter = request.args.get('category', '')
    conn = sqlite3.connect('assets.db')
    cursor = conn.cursor()
    query = 'SELECT * FROM assets WHERE 1=1'
    params = []
    if search_query:
        query += ' AND (name LIKE ? OR code LIKE ?)'
        params.extend([f'%{search_query}%', f'%{search_query}%'])
    if category_filter:
        query += ' AND category = ?'
        params.append(category_filter)
    cursor.execute(query, params)
    rows = cursor.fetchall()
    cursor.execute('SELECT DISTINCT category FROM assets')
    categories = [row[0] for row in cursor.fetchall()]
    conn.close()
    
    assets = []
    current_year = datetime.now().year
    for row in rows:
        asset_id, code, name, category, purchase_date_str, price, life_years, is_low_value = row
        purchase_date = datetime.strptime(purchase_date_str, "%Y-%m-%d")
        salvage_value = 1.0
        years_passed = current_year - purchase_date.year
        if is_low_value == 1:
            acc_dep, net_val, status = 0.0, price, "ต่ำกว่าเกณฑ์"
        else:
            dep_per_year = (price - salvage_value) / life_years
            acc_dep = min(dep_per_year * max(0, years_passed), price - salvage_value)
            net_val = max(1.0, price - acc_dep)
            status = "ปกติ (เส้นตรง)"
        assets.append({"id": asset_id, "code": code, "name": name, "category": category, "price": price, "status": status, "acc_dep": round(acc_dep, 2), "net_val": round(net_val, 2)})
    return render_template_string(HTML_TEMPLATE, assets=assets, categories=categories, search=search_query, selected_cat=category_filter, view_mode='list')

@app.route('/schedule/<int:asset_id>')
def schedule(asset_id):
    conn = sqlite3.connect('assets.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM assets WHERE id = ?', (asset_id,))
    row = cursor.fetchone()
    conn.close()
    if not row: return redirect(url_for('index'))
    asset = {"id": row[0], "code": row[1], "name": row[2], "category": row[3], "purchase_date": row[4], "price": row[5], "life_years": row[6], "is_low_value": row[7]}
    purchase_year = datetime.strptime(asset["purchase_date"], "%Y-%m-%d").year
    life_years, price, salvage_value, is_low_value = asset["life_years"], asset["price"], 1.0, asset["is_low_value"]
    schedule_rows = []
    if is_low_value == 1:
        schedule_rows.append({"year_be": purchase_year + 543, "beginning_book_value": price, "depreciation": 0.0, "accumulated": 0.0, "ending_book_value": price})
    else:
        dep_per_year = (price - salvage_value) / life_years
        accumulated, current_book_value = 0.0, price
        for i in range(life_years):
            year_be = purchase_year + i + 543
            beg_val = current_book_value
            dep = (price - salvage_value - accumulated) if i == life_years - 1 else dep_per_year
            accumulated += dep
            end_val = max(salvage_value, price - accumulated)
            schedule_rows.append({"year_be": year_be, "beginning_book_value": round(beg_val, 2), "depreciation": round(dep, 2), "accumulated": round(accumulated, 2), "ending_book_value": round(end_val, 2)})
            current_book_value = end_val
    return render_template_string(HTML_TEMPLATE, view_mode='schedule', asset=asset, schedule=schedule_rows)

@app.route('/schedule/export_excel/<int:asset_id>')
def schedule_export_excel(asset_id):
    conn = sqlite3.connect('assets.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM assets WHERE id = ?', (asset_id,))
    row = cursor.fetchone()
    conn.close()
    if not row: return redirect(url_for('index'))
    code, purchase_year, price, life_years, is_low_value, salvage_value = row[1], datetime.strptime(row[4], "%Y-%m-%d").year, row[5], row[6], row[7], 1.0
    data = []
    if is_low_value == 1:
        data.append({"ปีพุทธศักราช (พ.ศ.)": purchase_year + 543, "มูลค่าต้นงวด (บาท)": price, "ค่าเสื่อมราคาประจำปี (บาท)": 0.0, "ค่าเสื่อมราคาสะสม (บาท)": 0.0, "มูลค่าสุทธิปลายงวด (บาท)": price})
    else:
        dep_per_year = (price - salvage_value) / life_years
        accumulated, current_book_value = 0.0, price
        for i in range(life_years):
            year_be = purchase_year + i + 543
            beg_val = current_book_value
            dep = (price - salvage_value - accumulated) if i == life_years - 1 else dep_per_year
            accumulated += dep
            end_val = max(salvage_value, price - accumulated)
            data.append({"ปีพุทธศักราช (พ.ศ.)": year_be, "มูลค่าต้นงวด (บาท)": round(beg_val, 2), "ค่าเสื่อมราคาประจำปี (บาท)": round(dep, 2), "ค่าเสื่อมราคาสะสม (บาท)": round(accumulated, 2), "มูลค่าสุทธิปลายงวด (บาท)": round(end_val, 2)})
            current_book_value = end_val
    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name='ค่าเสื่อมรายปี')
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f'depreciation_{code}.xlsx')

@app.route('/add', methods=['POST'])
def add_asset():
    code, name, category, purchase_date, price, life_years = request.form['code'], request.form['name'], request.form['category'], request.form['purchase_date'], float(request.form['price']), int(request.form['life_years'])
    is_low_value = 1 if price < 5000 else 0
    conn = sqlite3.connect('assets.db')
    cursor = conn.cursor()
    try:
        cursor.execute('INSERT INTO assets (code, name, category, purchase_date, price, life_years, is_low_value) VALUES (?, ?, ?, ?, ?, ?, ?)', (code, name, category, purchase_date, price, life_years, is_low_value))
        conn.commit()
    except sqlite3.IntegrityError: pass
    conn.close()
    return redirect(url_for('index'))

@app.route('/export_excel')
def export_excel():
    conn = sqlite3.connect('assets.db')
    df = pd.read_sql_query('SELECT code AS "รหัสครุภัณฑ์", name AS "ชื่อรายการ", category AS "ประเภท", purchase_date AS "วันที่จัดซื้อ", price AS "ราคาทุน", life_years AS "อายุการใช้งาน(ปี)" FROM assets', conn)
    conn.close()
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='asset_report.xlsx')

if __name__ == '__main__':
    init_db()
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)