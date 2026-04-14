from flask import Flask, request, jsonify
import base64
import io
import pyxlsb
import openpyxl

app = Flask(__name__)

@app.route("/convert", methods=["POST"])
def convert():
    try:
        body = request.get_json()
        if not body or "file" not in body:
            return jsonify({"error": "No file provided"}), 400

        file_bytes = base64.b64decode(body["file"])
        xlsb_stream = io.BytesIO(file_bytes)

        rows_data = []
        with pyxlsb.open_workbook(xlsb_stream) as wb:
            sheet_name = wb.sheets[0]
            with wb.get_sheet(sheet_name) as sheet:
                for row in sheet.rows():
                    rows_data.append([
                        item.v if item.v is not None else ""
                        for item in row
                    ])

        if not rows_data:
            return jsonify({"error": "Empty sheet"}), 400

        rows_data = rows_data[:25000]
        rows_data = [row[:40] for row in rows_data]
        max_cols = max(len(r) for r in rows_data)
        rows_data = [r + [""] * (max_cols - len(r)) for r in rows_data]

        wb_out = openpyxl.Workbook()
        ws = wb_out.active
        for row in rows_data:
            ws.append(row)

        xlsx_stream = io.BytesIO()
        wb_out.save(xlsx_stream)
        xlsx_stream.seek(0)

        encoded = base64.b64encode(xlsx_stream.read()).decode("utf-8")
        return jsonify({
            "xlsx": encoded,
            "rows": len(rows_data),
            "cols": max_cols
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/ping", methods=["GET"])
def ping():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
