from flask import Flask, request, jsonify
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from io import BytesIO
import base64
import datetime

app = Flask(__name__)

@app.route('/')
def home():
    return "Excel Report API is running successfully!"

@app.route('/generate_excel', methods=['POST'])
def generate_excel():
    data = request.get_json().get("data", [])

    wb = Workbook()
    ws = wb.active
    ws.title = "Weekly Report"

    # Title row
    ws.merge_cells('A1:F1')
    title_cell = ws['A1']
    title_cell.value = "Weekly Site Report"
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    title_cell.fill = PatternFill(start_color="B7DEE8", end_color="B7DEE8", fill_type="solid")

    # Headers
    headers = ["Activity ID", "Site ID", "Date", "Before Photo URL", "After Photo URL"]
    ws.append(headers)
    for col in ws.iter_cols(min_row=2, max_row=2, min_col=1, max_col=len(headers)):
        for cell in col:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

    # Data rows
    for row in data:
        ws.append([
            row.get("activity_id", ""),
            row.get("site_id", ""),
            row.get("activity_date", ""),
            row.get("before_photo", ""),
            row.get("after_photo", "")
        ])

    # Auto-width
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column].width = max_length + 2

    # Save to memory
    output = BytesIO()
    wb.save(output)
    excel_data = output.getvalue()

    # Encode to base64
    encoded = base64.b64encode(excel_data).decode("utf-8")
    filename = f"Weekly_Report_{datetime.date.today().isoformat()}.xlsx"

    return jsonify({
        "fileData": encoded,
        "fileName": filename
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
