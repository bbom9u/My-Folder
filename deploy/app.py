"""
데이터 분류 수집기 - 클라우드 서버
Render.com 배포용
"""

import os
import json
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS

app = Flask(__name__, static_folder="static")
CORS(app)

# 데이터 파일 경로 (Render에서는 /tmp 사용)
DATA_DIR = os.environ.get("DATA_DIR", os.path.dirname(os.path.abspath(__file__)))
DATA_FILE = os.path.join(DATA_DIR, "sync_data.json")


def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/api/boxes", methods=["GET"])
def get_boxes():
    return jsonify(load_data())


@app.route("/api/boxes", methods=["PUT"])
def update_all_boxes():
    data = request.get_json()
    save_data(data)
    return jsonify({"status": "ok"})


@app.route("/api/boxes/<box_name>", methods=["POST"])
def add_to_box(box_name):
    data = load_data()
    item = request.get_json()
    if box_name not in data:
        data[box_name] = []
    data[box_name].append(item)
    save_data(data)
    return jsonify({"status": "ok", "count": len(data[box_name])})


@app.route("/api/boxes/<box_name>", methods=["DELETE"])
def delete_box(box_name):
    data = load_data()
    if box_name in data:
        del data[box_name]
        save_data(data)
    return jsonify({"status": "ok"})


@app.route("/api/boxes/<box_name>/<int:idx>", methods=["DELETE"])
def delete_item(box_name, idx):
    data = load_data()
    if box_name in data and 0 <= idx < len(data[box_name]):
        data[box_name].pop(idx)
        save_data(data)
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
