from flask import Flask, render_template_string, request, redirect, url_for, send_file
import sqlite3
import pandas as pd
from datetime import datetime
import io
import os
from openpyxl.worksheet.datavalidation import DataValidation

app = Flask(__name__)

DB_PATH = '/tmp/assets.db'

def init_db():
    conn = sqlite3.connect(DB_PATH)
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
            is_low_value INTEGER,
            model_spec TEXT,
            location TEXT,
            money_type TEXT,
            acquire_method TEXT
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
                    <p class="mb-1"><strong>ยี่ห้อ/รุ่น:</strong> {{ asset.model_spec }}</p>
                </div>
                <div class="col-md-6">
                    <p class="mb-1"><strong>วันที่จัดซื้อ:</strong> {{ asset.purchase_date }}</p>
                    <p class="mb-1"><strong>ราคาทุน:</strong> {{ "{:,.2f}".format(asset.price) }} บาท</p>
                    <p class="mb-1"><strong>อายุการใช้งาน:</strong> {{ asset.life_years }} ปี</p>
                    <p class="mb-1"><strong>สถานที่ใช้งาน:</strong> {{ asset.location }}</p>
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
                    <a href="/schedule/export_excel/{{ asset.id }}" class="btn btn-success me-2 px-4">ดาวน์โหลด Excel ฟอร์มทางการ</a>
                    <button type="button" class="btn btn-danger px-4" onclick="window.print()">พิมพ์ PDF / เอกสาร</button>
                </div>
            </div>
        </div>
    {% else %}
        <div class="card mb-4 p-4">
            <h4 class="mb-3 text-primary fw-bold">เพิ่มครุภัณฑ์ใหม่ (ทะเบียนคุมทรัพย์สิน)</h4>
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
                <div class="col-md-3">
                    <label class="form-label fw-semibold">ยี่ห้อ / รุ่น / ลักษณะ</label>
                    <input type="text" class="form-control" name="model_spec" value="-">
                </div>
                <div class="col-md-3">
                    <label class="form-label fw-semibold">สถานที่ใช้งาน</label>
                    <input type="text" class="form-control" name="location" value="งานพัสดุ โรงเรียนอนุบาลอุทุมพรพิสัย">
                </div>
                <div class="col-md-4">
                    <label class="form-label fw-semibold">ประเภทเงิน</label>
                    <select name="money_type" class="form-select">
                        <option value="เงินงบประมาณ">เงินงบประมาณ</option>
                        <option value="เงินนอกงบประมาณ">เงินนอกงบประมาณ</option>
                        <option value="เงินบริจาค/เงินช่วยเหลือ">เงินบริจาค/เงินช่วยเหลือ</option>
                        <option value="อื่นๆ">อื่นๆ</option>
                    </select>
                </div>
                <div class="col-md-4">
                    <label class="form-label fw-semibold">วิธีการได้มา</label>
                    <select name="acquire_method" class="form-select">
                        <option value="วิธีเฉพาะเจาะจง">วิธีเฉพาะเจาะจง</option>
                        <option value="วิธีประกวดราคาอิเล็กทรอนิกส์ (e-bidding)">วิธีประกวดราคาอิเล็กทรอนิกส์ (e-bidding)</option>
                        <option value="วิธีคัดเลือก">วิธีคัดเลือก</option>
                        <option value="รับบริจาค">รับบริจาค</option>
                    </select>
                </div>
                <div class="col-md-4 d-flex align-items-end">
                    <button type="submit" class="btn btn-primary w-100 fw-semibold">บันทึกข้อมูลครุภัณฑ์</button>
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
                <a href="/export_excel" class="btn btn-success w-100 fw-semibold">ดาวน์โหลดภาพรวม</a>
            </div>
        </div>

        <table class="table table-bordered table-striped align-middle bg-white shadow-sm rounded">
            <thead class="table-dark">
                <tr>
                    <th>รหัส</th>
                    <th>ชื่อรายการ</th>
                    <th>ประเภท</th>
                    <th>วันที่จัดซื้อ</th>
                    <th>ราคาทุน</th>
                    <th>สถานะ</th>
                    <th>ค่าเสื่อมสะสม</th>
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
                    <td>{{ item.purchase_date }}</td>
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
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    query = 'SELECT * FROM assets WHERE 1=1'
    params = []
    if search_query:
        query += ' AND (name LIKE ? OR code LIKE ?)'
        params.extend([f'%{search_query}%', f'%{search_query}%'])
    if category_filter:
        query += ' AND category = ?'
        params.append(category_filter)
    
    query += ' ORDER BY purchase_date ASC'
    
    cursor.execute(query, params)
    rows = cursor.fetchall()
    cursor.execute('SELECT DISTINCT category FROM assets')
    categories = [row[0] for row in cursor.fetchall()]
    conn.close()
    
    assets = []
    current_year = datetime.now().year
    for row in rows:
        asset_id, code, name, category, purchase_date_str, price, life_years, is_low_value = row[:8]
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
        assets.append({"id": asset_id, "code": code, "name": name, "category": category, "purchase_date": purchase_date_str, "price": price, "status": status, "acc_dep": round(acc_dep, 2), "net_val": round(net_val, 2)})
    return render_template_string(HTML_TEMPLATE, assets=assets, categories=categories, search=search_query, selected_cat=category_filter, view_mode='list')

@app.route('/schedule/<int:asset_id>')
def schedule(asset_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM assets WHERE id = ?', (asset_id,))
    row = cursor.fetchone()
    conn.close()
    if not row: return redirect(url_for('index'))
    
    asset = {
        "id": row[0], "code": row[1], "name": row[2], "category": row[3], 
        "purchase_date": row[4], "price": row[5], "life_years": row[6], "is_low_value": row[7],
        "model_spec": row[8], "location": row[9], "money_type": row[10], "acquire_method": row[11]
    }
    
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
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM assets WHERE id = ?', (asset_id,))
    row = cursor.fetchone()
    conn.close()
    if not row: return redirect(url_for('index'))
    
    asset_id, code, name, category, purchase_date_str, price, life_years, is_low_value, model_spec, location, money_type, acquire_method = row
    purchase_year = datetime.strptime(purchase_date_str, "%Y-%m-%d").year
    salvage_value = 1.0
    
    schedule_data = []
    if is_low_value == 1:
        schedule_data.append({
            "ปีพุทธศักราช (พ.ศ.)": purchase_year + 543,
            "มูลค่าต้นงวด (บาท)": price,
            "ค่าเสื่อมราคาประจำปี (บาท)": 0.0,
            "ค่าเสื่อมราคาสะสม (บาท)": 0.0,
            "มูลค่าสุทธิปลายงวด (บาท)": price
        })
    else:
        dep_per_year = (price - salvage_value) / life_years
        accumulated, current_book_value = 0.0, price
        for i in range(life_years):
            year_be = purchase_year + i + 543
            beg_val = current_book_value
            dep = (price - salvage_value - accumulated) if i == life_years - 1 else dep_per_year
            accumulated += dep
            end_val = max(salvage_value, price - accumulated)
            schedule_data.append({
                "ปีพุทธศักราช (พ.ศ.)": year_be,
                "มูลค่าต้นงวด (บาท)": round(beg_val, 2),
                "ค่าเสื่อมราคาประจำปี (บาท)": round(dep, 2),
                "ค่าเสื่อมราคาสะสม (บาท)": round(accumulated, 2),
                "มูลค่าสุทธิปลายงวด (บาท)": round(end_val, 2)
            })
            current_book_value = end_val

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        header_data = [
            ["", "", "", "", "", "ทะเบียนคุมทรัพย์สิน", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "ส่วนราชการ", "สำนักงานคณะกรรมการการศึกษาขั้นพื้นฐาน", ""],
            ["", "", "", "", "", "", "", "", "หน่วยงาน", "โรงเรียนอนุบาลอุทุมพรพิสัย", ""],
            ["", "", "", "", "", "", "", "", "", "", ""],
            ["ประเภท", category, "", "", "หมายเลขครุภัณฑ์", code, "", "", "", "", ""],
            ["รายการ", name, "", "", "ยี่ห้อ/รุ่น/ลักษณะเฉพาะ", model_spec, "", "", "", "", ""],
            ["", "", "", "", "สถานที่ใช้งาน/หน่วยงานรับผิดชอบ", location, "", "", "", "", ""],
            ["ชื่อผู้ขาย/ผู้รับจ้าง/ผู้บริจาค", "-", "", "", "ที่อยู่", "-", "", "", "", "", ""],
            ["ประเภทเงิน", money_type, "", "", "", "", "", "", "", "", ""],
            ["วิธีการได้มา", acquire_method, "", "", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "", "", ""]
        ]
        
        df_header = pd.DataFrame(header_data)
        df_header.to_excel(writer, index=False, header=False, sheet_name='ทะเบียนคุมทรัพย์สิน')
        
        df_schedule = pd.DataFrame(schedule_data)
        df_schedule.to_excel(writer, index=False, startrow=12, sheet_name='ทะเบียนคุมทรัพย์สิน')
        
        workbook = writer.book
        worksheet = writer.sheets['ทะเบียนคุมทรัพย์สิน']
        
        dv_money = DataValidation(type="list", formula1='"เงินงบประมาณ, เงินนอกงบประมาณ, เงินบริจาค/เงินช่วยเหลือ, อื่นๆ"', allow_blank=True)
        worksheet.add_data_validation(dv_money)
        dv_money.add("B9")
        
        dv_method = DataValidation(type="list", formula1='"วิธีเฉพาะเจาะจง, วิธีประกวดราคาอิเล็กทรอนิกส์ (e-bidding), วิธีคัดเลือก, รับบริจาค"', allow_blank=True)
        worksheet.add_data_validation(dv_method)
        dv_method.add("B10")

    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f'form2_{code}.xlsx')

@app.route('/add', methods=['POST'])
def add_asset():
    code = request.form['code']
    name = request.form['name']
    category = request.form['category']
    purchase_date = request.form['purchase_date']
    price = float(request.form['price'])
    life_years = int(request.form['life_years'])
    model_spec = request.form.get('model_spec', '-')
    location = request.form.get('location', 'งานพัสดุ โรงเรียนอนุบาลอุทุมพรพิสัย')
    money_type = request.form.get('money_type', 'เงินงบประมาณ')
    acquire_method = request.form.get('acquire_method', 'วิธีเฉพาะเจาะจง')
    
    is_low_value = 1 if price < 5000 else 0
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    try:
        cursor.execute('''
            INSERT INTO assets (code, name, category, purchase_date, price, life_years, is_low_value, model_spec, location, money_type, acquire_method) 
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (code, name, category, purchase_date, price, life_years, is_low_value, model_spec, location, money_type, acquire_method))
        conn.commit()
    except sqlite3.IntegrityError: 
        pass
    conn.close()
    return redirect(url_for('index'))

@app.route('/export_excel')
def export_excel():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query('SELECT code AS "รหัสครุภัณฑ์", name AS "ชื่อรายการ", category AS "ประเภท", purchase_date AS "วันที่จัดซื้อ", price AS "ราคาทุน", life_years AS "อายุการใช้งาน(ปี)" FROM assets ORDER BY purchase_date ASC', conn)
    conn.close()
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name='รายการครุภัณฑ์ทั้งหมด')
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='all_assets_report.xlsx')

if __name__ == '__main__':
    init_db()
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
else:
    init_db()
