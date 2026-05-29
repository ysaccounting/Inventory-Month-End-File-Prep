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
    buf = io.BytesIO()

    # Pivot: Main Company | Sum of Total Cost, sorted A-Z
    pivot = (
        df.groupby("Main Company", sort=True)["Total Cost"]
        .sum()
        .reset_index()
        .rename(columns={"Total Cost": "Sum of Total Cost"})
    )
    grand_total = pivot["Sum of Total Cost"].sum()
    n = len(pivot)  # number of companies

    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book

        # Create Summary FIRST so it is the first tab

        # ── Formats ────────────────────────────────────────────

        # Pivot / Summary header (dark blue)
        dark_hdr_left = wb.add_format({
            "bold": True, "font_name": "Calibri", "font_size": 11,
            "font_color": "#FFFFFF", "bg_color": "#1F4E79",
            "align": "left", "valign": "vcenter", "border": 0,
        })
        dark_hdr_right = wb.add_format({
            "bold": True, "font_name": "Calibri", "font_size": 11,
            "font_color": "#FFFFFF", "bg_color": "#1F4E79",
            "align": "right", "valign": "vcenter", "border": 0,
        })
        dark_hdr_center = wb.add_format({
            "bold": True, "font_name": "Calibri", "font_size": 11,
            "font_color": "#FFFFFF", "bg_color": "#1F4E79",
            "align": "center", "valign": "vcenter", "border": 0,
        })

        # Section title (merged, dark blue, centered) — "Current Month End" / "Prior Month End"
        section_title_fmt = wb.add_format({
            "bold": True, "font_name": "Calibri", "font_size": 11,
            "font_color": "#1F4E79", "bg_color": "#FFFFFF",
            "align": "center", "valign": "vcenter",
            "bottom": 2, "bottom_color": "#1F4E79",
        })

        # Pivot data rows (light blue alternating)
        row_label_fmt = wb.add_format({
            "font_name": "Calibri", "font_size": 11,
            "bg_color": "#DCE6F1", "align": "left", "valign": "vcenter",
        })
        row_money_fmt = wb.add_format({
            "font_name": "Calibri", "font_size": 11,
            "bg_color": "#DCE6F1", "num_format": '$#,##0.00',
            "align": "right", "valign": "vcenter",
        })

        # Grand total row (dark blue)
        total_label_fmt = wb.add_format({
            "bold": True, "font_name": "Calibri", "font_size": 11,
            "bg_color": "#1F4E79", "font_color": "#FFFFFF",
            "align": "left", "valign": "vcenter",
        })
        total_money_fmt = wb.add_format({
            "bold": True, "font_name": "Calibri", "font_size": 11,
            "bg_color": "#1F4E79", "font_color": "#FFFFFF",
            "num_format": '$#,##0.00', "align": "right", "valign": "vcenter",
        })
        total_pct_fmt = wb.add_format({
            "bold": True, "font_name": "Calibri", "font_size": 11,
            "bg_color": "#1F4E79", "font_color": "#FFFFFF",
            "num_format": '0.00%', "align": "right", "valign": "vcenter",
        })

        # Side table data rows (white background, user-input money/pct)
        side_money_fmt = wb.add_format({
            "font_name": "Calibri", "font_size": 11,
            "bg_color": "#DCE6F1", "num_format": '$#,##0.00',
            "align": "right", "valign": "vcenter",
        })
        side_pct_fmt = wb.add_format({
            "font_name": "Calibri", "font_size": 11,
            "bg_color": "#DCE6F1", "num_format": '0.00%',
            "align": "right", "valign": "vcenter",
        })
        # User-input cells (light yellow tint to indicate editable)
        input_money_fmt = wb.add_format({
            "font_name": "Calibri", "font_size": 11,
            "bg_color": "#FFFFC0", "num_format": '$#,##0.00',
            "align": "right", "valign": "vcenter",
        })
        input_pct_fmt = wb.add_format({
            "font_name": "Calibri", "font_size": 11,
            "bg_color": "#FFFFC0", "num_format": '0.00%',
            "align": "right", "valign": "vcenter",
        })

        # ── Summary sheet ──────────────────────────────────────
        ws2 = wb.add_worksheet("Summary")

        # Column widths
        # A: Main Company, B: Sum of Total Cost, C: spacer
        # D: Total per QBO (current), E: $ Diff (current), F: % Diff (current), G: spacer
        # H: Total per QBO (prior),  I: $ Diff (prior),  J: % Diff (prior)
        ws2.set_column(0, 0, 22)   # A - Main Company
        ws2.set_column(1, 1, 18)   # B - Sum of Total Cost
        ws2.set_column(2, 2, 3)    # C - spacer
        ws2.set_column(3, 3, 18)   # D - Current Total per QBO
        ws2.set_column(4, 4, 16)   # E - Current $ Diff
        ws2.set_column(5, 5, 10)   # F - Current % Diff
        ws2.set_column(6, 6, 3)    # G - spacer
        ws2.set_column(7, 7, 18)   # H - Prior Total per QBO
        ws2.set_column(8, 8, 16)   # I - Prior $ Diff
        ws2.set_column(9, 9, 10)   # J - Prior % Diff

        # Row 0: section titles (Current Month End / Prior Month End)
        ws2.merge_range(0, 3, 0, 5, "Current Month End", section_title_fmt)
        ws2.merge_range(0, 7, 0, 9, "Prior Month End",   section_title_fmt)
        ws2.set_row(0, 18)

        # Row 1: column headers
        ws2.write(1, 0, "Main Company",       dark_hdr_left)
        ws2.write(1, 1, "Sum of Total Cost",  dark_hdr_right)
        ws2.write(1, 3, "Total per QBO",      dark_hdr_right)
        ws2.write(1, 4, "$ Diff",             dark_hdr_right)
        ws2.write(1, 5, "% Diff",             dark_hdr_right)
        ws2.write(1, 7, "Total per QBO",      dark_hdr_right)
        ws2.write(1, 8, "$ Diff",             dark_hdr_right)
        ws2.write(1, 9, "% Diff",             dark_hdr_right)
        ws2.set_row(1, 18)

        # Pivot data rows + side table rows (data starts at Excel row 3, i.e. index 2)
        # Pivot values go in cols A-B.
        # Col D = user input (Current Total per QBO)
        # Col E = XLOOKUP pivot value for this company - D  ($ Diff current)
        # Col F = E / XLOOKUP pivot value                   (% Diff current)
        # Col H = user input (Prior Total per QBO)
        # Col I = user input (Prior $ Diff)
        # Col J = user input (Prior % Diff)

        # Pivot range for XLOOKUP: companies in A3:A{n+2}, values in B3:B{n+2}
        for ri, row in enumerate(pivot.itertuples(index=False), 2):  # 0-indexed row
            excel_row = ri  # 0-based for xlsxwriter
            company   = row[0]
            amount    = row[1]

            # Pivot cols
            ws2.write(excel_row, 0, company, row_label_fmt)
            ws2.write(excel_row, 1, amount,  row_money_fmt)

            # Current Month End — D=user input, E=B-D, F=E/B
            ws2.write_blank(excel_row, 3, None, input_money_fmt)

            b_cell = f"B{excel_row+1}"
            d_cell = f"D{excel_row+1}"
            e_cell = f"E{excel_row+1}"
            ws2.write_formula(excel_row, 4, f"={b_cell}-{d_cell}", side_money_fmt)
            ws2.write_formula(excel_row, 5, f"=IF({b_cell}=0,0,{e_cell}/{b_cell})", side_pct_fmt)

            # Prior Month End — H, I, J all user input (yellow)
            ws2.write_blank(excel_row, 7, None, input_money_fmt)
            ws2.write_blank(excel_row, 8, None, input_money_fmt)
            ws2.write_blank(excel_row, 9, None, input_pct_fmt)

            ws2.set_row(excel_row, 18)

        # Grand Total row
        total_row = n + 2  # 0-based
        ws2.write(total_row, 0, "Grand Total", total_label_fmt)
        ws2.write(total_row, 1, grand_total,   total_money_fmt)

        # Current grand total: sum of D col, sum of E col, E_total/B_total
        d_range = f"D3:D{n+2}"
        e_range = f"E3:E{n+2}"
        ws2.write_formula(total_row, 3, f"=SUM({d_range})", total_money_fmt)
        ws2.write_formula(total_row, 4, f"=SUM({e_range})", total_money_fmt)
        ws2.write_formula(total_row, 5, f"=IF(B{total_row+1}=0,0,E{total_row+1}/B{total_row+1})", total_pct_fmt)

        # Prior grand total: sum of H, I, J cols
        h_range = f"H3:H{n+2}"
        i_range = f"I3:I{n+2}"
        j_range = f"J3:J{n+2}"
        ws2.write_formula(total_row, 7, f"=SUM({h_range})", total_money_fmt)
        ws2.write_formula(total_row, 8, f"=SUM({i_range})", total_money_fmt)
        ws2.write_formula(total_row, 9, f"=SUM({j_range})", total_pct_fmt)

        ws2.set_row(total_row, 18)

        # ── Available tab ──────────────────────────────────────
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]

        hdr_fmt = wb.add_format({
            "bold": True, "font_name": "Arial", "font_size": 10,
            "font_color": "#FFFFFF", "bg_color": "#1F4E79",
            "align": "center", "valign": "vcenter", "text_wrap": True, "border": 0,
        })
        money_fmt = wb.add_format({"font_name": "Arial", "font_size": 10, "num_format": "#,##0.00"})
        data_fmt  = wb.add_format({"font_name": "Arial", "font_size": 10})

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

        # Put Summary first

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
