from flask import Flask, render_template, request, jsonify
from flask_socketio import SocketIO
import pandas as pd
from datetime import datetime

app = Flask(__name__)
socketio = SocketIO(app, cors_allowed_origins="*")

DATA_RAW = pd.DataFrame(columns=["Tanggal", "Kategori", "Barang", "Masuk", "Keluar"])

def parse_excel(file_storage):
    df = pd.read_excel(file_storage)
    cols = {c.strip().lower(): c for c in df.columns}
    required = ["tanggal", "kategori", "barang", "masuk", "keluar"]
    for r in required:
        if r not in cols:
            raise ValueError(f"Kolom '{r}' tidak ditemukan di Excel. Kolom yang dibutuhkan: {required}")
    df = df[[cols["tanggal"], cols["kategori"], cols["barang"], cols["masuk"], cols["keluar"]]].copy()
    df[cols["tanggal"]] = pd.to_datetime(df[cols["tanggal"]], errors="coerce")
    df.rename(columns={cols["tanggal"]:"Tanggal", cols["kategori"]:"Kategori", cols["barang"]:"Barang",
                       cols["masuk"]:"Masuk", cols["keluar"]:"Keluar"}, inplace=True)
    for k in ["Masuk", "Keluar"]:
        df[k] = pd.to_numeric(df[k], errors="coerce").fillna(0).astype(int)
    df["Kategori"] = df["Kategori"].astype(str).str.title().replace({
        "Electrical":"Electrical", "Elektrikal":"Electrical",
        "Lighting":"Lighting", "Lampu":"Lighting"
    })
    df = df.dropna(subset=["Tanggal"])
    return df

def build_payload(df):
    if df.empty:
        return {
            "summary": {"totalMasuk":0, "totalKeluar":0, "totalStock":0, "lastUpdate": None},
            "dailySeries": [],
            "categoryMasuk": [],
            "categoryKeluar": [],
            "stockByItem": [],
            "keluarByItem": [],
            "table": []
        }

    total_masuk = int(df["Masuk"].sum())
    total_keluar = int(df["Keluar"].sum())
    stock_item = df.groupby("Barang").agg({"Masuk":"sum","Keluar":"sum"})
    stock_item["Stock"] = stock_item["Masuk"] - stock_item["Keluar"]
    total_stock = int(stock_item["Stock"].sum())

    daily = df.groupby(df["Tanggal"].dt.date).agg({"Masuk":"sum","Keluar":"sum"}).reset_index()
    daily = daily.sort_values("Tanggal")
    daily_series = [{"date": d.strftime("%Y-%m-%d"), "masuk": int(m), "keluar": int(k)} for d,m,k in zip(daily["Tanggal"], daily["Masuk"], daily["Keluar"])]
    cat_masuk = df.groupby("Kategori")["Masuk"].sum().reset_index()
    cat_keluar = df.groupby("Kategori")["Keluar"].sum().reset_index()
    category_masuk = [{"name": r["Kategori"], "value": int(r["Masuk"])} for _,r in cat_masuk.iterrows()]
    category_keluar = [{"name": r["Kategori"], "value": int(r["Keluar"])} for _,r in cat_keluar.iterrows()]
    stock_top = stock_item.reset_index().sort_values("Stock", ascending=False)
    stock_by_item = [{"name": r["Barang"], "value": int(r["Stock"])} for _,r in stock_top.iterrows()]
    keluar_top = df.groupby("Barang")["Keluar"].sum().sort_values(ascending=False).reset_index()
    keluar_by_item = [{"name": r["Barang"], "value": int(r["Keluar"])} for _,r in keluar_top.iterrows()]
    df_sorted = df.sort_values("Tanggal", ascending=False).head(100)
    table = [{
        "tanggal": r["Tanggal"].strftime("%Y-%m-%d"),
        "kategori": r["Kategori"],
        "barang": r["Barang"],
        "masuk": int(r["Masuk"]),
        "keluar": int(r["Keluar"])
    } for _, r in df_sorted.iterrows()]

    payload = {
        "summary": {
            "totalMasuk": int(total_masuk),
            "totalKeluar": int(total_keluar),
            "totalStock": int(total_stock),
            "lastUpdate": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        },
        "dailySeries": daily_series,
        "categoryMasuk": category_masuk,
        "categoryKeluar": category_keluar,
        "stockByItem": stock_by_item,
        "keluarByItem": keluar_by_item,
        "table": table
    }
    return payload

@app.route("/")
def index():
    initial = build_payload(DATA_RAW)
    return render_template("index.html", initial_json=initial)

@app.route("/upload", methods=["POST"])
def upload():
    global DATA_RAW
    file = request.files.get("file")
    if not file:
        return "File tidak ditemukan", 400
    try:
        df = parse_excel(file)
    except Exception as e:
        return f"Gagal memproses Excel: {e}", 400
    DATA_RAW = df.copy()
    payload = build_payload(DATA_RAW)
    socketio.emit("update_data", payload)
    return "Upload sukses"

@app.route("/api/data")
def api_data():
    payload = build_payload(DATA_RAW)
    return jsonify(payload)

if __name__ == "__main__":
    socketio.run(app, debug=True, allow_unsafe_werkzeug=True)
