import os
from flask import Flask, request, jsonify
from flask_cors import CORS
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ─── App setup ───────────────────────────────────────────────────────────
app = Flask(__name__)
CORS(app)  # allow all origins (Netlify ↔ Render)

# ─── Google Sheets auth ──────────────────────────────────────────────────
# path to your JSON key (upload it or point to env)
KEYFILE = os.environ.get("GOOGLE_CREDENTIALS_FILE", "credentials.json")
SCOPE   = ["https://spreadsheets.google.com/feeds",
           "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(KEYFILE, SCOPE)
gc    = gspread.authorize(creds)

# ─── Target spreadsheet & worksheet ──────────────────────────────────────
SPREADSHEET_NAME = os.environ.get("SPREADSHEET_NAME", "JR and Co.")
WORKSHEET_NAME   = os.environ.get("WORKSHEET_NAME", "Production Orders")
sheet = gc.open(SPREADSHEET_NAME).worksheet(WORKSHEET_NAME)

# ─── POST /submit ─────────────────────────────────────────────────────────
@app.route("/submit", methods=["POST"])
def submit():
    data = request.get_json(force=True)
    headers = sheet.row_values(1)            # first row = column names
    # build a row in same order as headers
    row = [ data.get(h, "") for h in headers ]
    sheet.append_row(row, value_input_option="RAW")
    return jsonify({"status":"ok"}), 200

# ─── Health check ─────────────────────────────────────────────────────────
@app.route("/healthz")
def healthz():
    return "OK", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
