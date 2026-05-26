from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd
import io
import os
import zipfile
import traceback

# Use calamine for fast xlsx reading if available, else fall back to openpyxl
try:
    import python_calamine  # noqa
    EXCEL_ENGINE = "calamine"
except ImportError:
    EXCEL_ENGINE = "openpyxl"

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB

COMPANY_MAPPING = {
    # Non Y&S
    "damon and crew": "Non Y&S",
    "the ticket guy": "Non Y&S",
    "yourtickets": "Non Y&S",
    # Y&S
    "gk llc": "Y&S",
    "jacks ys": "Y&S",
    "levovitz": "Y&S",
    "needle tickets llc": "Y&S",
    "pollak tickets": "Y&S",
    "yoni levine": "Y&S",
    "ys katz": "Y&S",
    "ys tickets": "Y&S",
    "ys tickets spec": "Y&S",
    "ys tl": "Y&S",
    "ysa": "Y&S",
    "ysa 2": "Y&S",
    "ysa 3": "Y&S",
    "ysm tickets": "Y&S",
    "yss tickets": "Y&S",
    "ys-seatgeek": "Y&S",
    "ys-seatgeek2": "Y&S",
    "ysw": "Y&S",
}

def get_main_company(company):
    c = company.strip().lower()
    if c in ("ys-seatgeek", "ys-seatgeek2", "ys tickets spec"):
        return "YS Tickets"
    if c in ("ysa 2", "ysa 3"):
        return "YSA"
    return company

def get_ys_mapping(company):
    return COMPANY_MAPPING.get(company.strip().lower(), "Non Y&S")

def build_excel_output(df, sheet_name="Available"):
    """Write Excel using xlsxwriter via pandas — fast bulk write with header styling."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        wb = writer.book
        ws = writer.sheets[sheet_name]

        hdr_fmt = wb.add_format({
            "bold": True,
            "font_name": "Arial",
            "font_size": 10,
            "font_color": "#FFFFFF",
            "bg_color": "#1F4E79",
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
            "border": 0,
        })
        money_fmt = wb.add_format({
            "font_name": "Arial",
            "font_size": 10,
            "num_format": "#,##0.00",
        })
        data_fmt = wb.add_format({
            "font_name": "Arial",
            "font_size": 10,
        })

        headers = list(df.columns)
        col_widths = [max(len(str(h)), 8) for h in headers]

        for ci, h in enumerate(headers):
            ws.write(0, ci, h, hdr_fmt)
            sample = df.iloc[:500, ci].astype(str)
            col_widths[ci] = min(max(col_widths[ci], sample.str.len().max() if len(sample) else 8) + 2, 40)

        for ci, h in enumerate(headers):
            fmt = money_fmt if h in ("Cost", "Total Cost") else data_fmt
            ws.set_column(ci, ci, col_widths[ci], fmt)

        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(df), len(headers) - 1)

    buf.seek(0)
    return buf

@app.errorhandler(Exception)
def handle_exception(e):
    return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process():
    try:
        files = request.files.getlist("files")
        month_end_date = request.form.get("month_end_date", "").strip()

        if not files or all(f.filename == "" for f in files):
            return jsonify({"error": "No files uploaded."}), 400
        if not month_end_date:
            return jsonify({"error": "Please enter a month end date."}), 400

        dfs = []
        for f in files:
            if f.filename.endswith(".xlsx"):
                try:
                    df = pd.read_excel(f, dtype=str, engine=EXCEL_ENGINE)
                    dfs.append(df)
                except Exception as e:
                    return jsonify({"error": f"Could not read {f.filename}: {str(e)}"}), 400

        if not dfs:
            return jsonify({"error": "No valid Excel files found."}), 400

        combined = pd.concat(dfs, ignore_index=True)

        combined["Cost"]       = pd.to_numeric(combined["Cost"],     errors="coerce")
        combined["Quantity"]   = pd.to_numeric(combined["Quantity"], errors="coerce")
        combined["Total Cost"] = combined["Quantity"] * combined["Cost"]

        cols = list(combined.columns)
        cost_idx = cols.index("Cost")
        cols.remove("Total Cost")
        cols.insert(cost_idx + 1, "Total Cost")
        combined = combined[cols]

        combined.insert(0, "Main Company", combined["Company"].apply(get_main_company))
        combined["_ys_flag"] = combined["Company"].apply(get_ys_mapping)

        df_ys     = combined[combined["_ys_flag"] == "Y&S"].drop(columns=["_ys_flag"])
        df_non_ys = combined[combined["_ys_flag"] == "Non Y&S"].drop(columns=["_ys_flag"])

        fname_ys     = f"Inventory {month_end_date} (YS).xlsx"
        fname_non_ys = f"Inventory {month_end_date} (Non YS).xlsx"

        buf_ys     = build_excel_output(df_ys,     sheet_name="Available")
        buf_non_ys = build_excel_output(df_non_ys, sheet_name="Available")

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(fname_ys,     buf_ys.read())
            zf.writestr(fname_non_ys, buf_non_ys.read())
        zip_buf.seek(0)

        response = send_file(
            zip_buf,
            mimetype="application/zip",
            as_attachment=True,
            download_name=f"Inventory {month_end_date}.zip"
        )
        response.headers["X-YS-Rows"]    = str(len(df_ys))
        response.headers["X-NonYS-Rows"] = str(len(df_non_ys))
        response.headers["X-Total-Rows"] = str(len(combined))
        response.headers["Access-Control-Expose-Headers"] = "X-YS-Rows, X-NonYS-Rows, X-Total-Rows"
        return response

    except Exception as e:
        return jsonify({"error": str(e), "detail": traceback.format_exc()}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
