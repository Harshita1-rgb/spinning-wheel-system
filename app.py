from flask import Flask, request, jsonify, send_file, send_from_directory
import openpyxl
from openpyxl import Workbook
import os
from flask_cors import CORS
from datetime import datetime

app = Flask(__name__)
CORS(app)

EXCEL_FILE = "spin_responses.xlsx"
PASSWORD = "ecoil123"

# ✅ Serve static HTML, CSS, JS files from root
@app.route('/')
def root():
    return send_from_directory('.', 'index.html')  # or vendor.html

@app.route('/<path:filename>')
def serve_file(filename):
    return send_from_directory('.', filename)

# ✅ Save spin response to Excel
@app.route('/submit_spin', methods=['POST'])
def submit_spin():
    data = request.json
    name = data.get('name')
    phone = data.get('phone')
    role = data.get('role')
    prize = data.get('prize')
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.append(["Name", "Phone", "Role", "Prize", "Timestamp"])
        wb.save(EXCEL_FILE)

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    ws.append([name, phone, role, prize, timestamp])
    wb.save(EXCEL_FILE)

    return jsonify({"message": "Spin result saved successfully."}), 200

# ✅ Download Excel
@app.route('/download_excel', methods=['GET'])
def download_excel():
    pwd = request.args.get('password')
    if pwd != PASSWORD:
        return jsonify({"error": "Unauthorized"}), 403
    return send_file(EXCEL_FILE, as_attachment=True)

# ✅ Clear Excel
@app.route('/clear_excel', methods=['POST'])
def clear_excel():
    pwd = request.json.get('password')
    if pwd != PASSWORD:
        return jsonify({"error": "Unauthorized"}), 403

    wb = Workbook()
    ws = wb.active
    ws.append(["Name", "Phone", "Role", "Prize", "Timestamp"])
    wb.save(EXCEL_FILE)
    return jsonify({"message": "Excel data cleared successfully."}), 200

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
