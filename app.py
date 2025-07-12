from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl
import os

app = Flask(__name__, static_folder=".", static_url_path="")
CORS(app)

ADMIN_PASSWORD = "ecoil123"
EXCEL_FILE = "spin_responses.xlsx"

@app.route("/")
def index():
    return app.send_static_file("index.html")

@app.route("/admin")
def admin():
    return app.send_static_file("admin.html")

@app.route("/download_excel", methods=["POST"])
def download_excel():
    data = request.get_json()
    if data.get("password") != ADMIN_PASSWORD:
        return jsonify({"error": "Invalid password"}), 403

    if os.path.exists(EXCEL_FILE):
        return send_file(EXCEL_FILE, as_attachment=True)
    else:
        return jsonify({"error": "Excel file not found"}), 404

@app.route("/clear_excel", methods=["POST"])
def clear_excel():
    data = request.get_json()
    if data.get("password") != ADMIN_PASSWORD:
        return jsonify({"error": "Invalid password"}), 403

    try:
        if os.path.exists(EXCEL_FILE):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Name", "Phone", "Role", "Prize", "Timestamp"])
            wb.save(EXCEL_FILE)
            return jsonify({"message": "✅ Responses cleared."})
        else:
            return jsonify({"error": "Excel file not found"}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500
