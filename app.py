import io
import json
import os
import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file

app = Flask(__name__)
DATA_FILE = "data.json"


def load_data():
    if not os.path.exists(DATA_FILE):
        return []
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return []


def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/persons", methods=["GET"])
def get_persons():
    data = load_data()
    return jsonify(data)


@app.route("/api/persons", methods=["POST"])
def add_person():
    data = load_data()
    req_data = request.json
    new_id = max([item["id"] for item in data], default=0) + 1
    req_data["id"] = new_id
    data.append(req_data)
    save_data(data)
    return jsonify({"success": True, "message": "Thêm mới thành công!", "id": new_id})


@app.route("/api/persons/<int:person_id>", methods=["PUT"])
def update_person(person_id):
    data = load_data()
    req_data = request.json
    updated = False
    for item in data:
        if item["id"] == person_id:
            item.update(req_data)
            item["id"] = person_id
            updated = True
            break
    if updated:
        save_data(data)
        return jsonify({"success": True, "message": "Cập nhật thành công!"})
    return jsonify({"success": False, "message": "Không tìm thấy dữ liệu!"}), 404


@app.route("/api/persons/<int:person_id>", methods=["DELETE"])
def delete_person(person_id):
    data = load_data()
    new_data = [item for item in data if item["id"] != person_id]
    if len(new_data) < len(data):
        save_data(new_data)
        return jsonify({"success": True, "message": "Xóa thành công!"})
    return jsonify({"success": False, "message": "Không tìm thấy dữ liệu!"}), 404


# API XUẤT EXCEL (Đã thêm cột Quan hệ với nhân viên)
@app.route("/api/export", methods=["GET"])
def export_excel():
    data = load_data()
    if not data:
        return (
            jsonify({"success": False, "message": "Không có dữ liệu để xuất!"}),
            400,
        )

    export_data = []
    for index, item in enumerate(data, 1):
        export_data.append({
            "STT": index,
            "Họ và Tên CBNV": item.get("hoTenCBNV", ""),
            "Quan hệ với nhân viên": item.get("quanHeVoiNhanVien", ""),
            "Họ và Tên người liên quan": item.get("hoTenNguoiLienQuan", ""),
            "Địa chỉ người liên quan": item.get("diaChiNguoiLienQuan", ""),
            "Số CCCD người liên quan": item.get("soCCCD", ""),
            "Ngày cấp CCCD": item.get("ngayCapCCCD", ""),
            "Nơi cấp CCCD": item.get("noiCapCCCD", ""),
            "Số điện thoại": item.get("soDienThoai", ""),
        })

    df = pd.DataFrame(export_data)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="DanhSachNguoiLienQuan")
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="DanhSachNguoiLienQuan.xlsx",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)