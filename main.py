from flask import Flask, request, jsonify
from werkzeug.exceptions import BadRequest, UnsupportedMediaType
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
    try:
        # Get JSON data with proper error handling
        try:
            json_data = request.get_json()
        except (BadRequest, UnsupportedMediaType) as e:
            return jsonify({"error": str(e.description)}), e.code
        
        if json_data is None:
            return jsonify({"error": "Invalid JSON or no data provided"}), 400
        
        data = json_data.get("data", [])
        
        # Validate data is a list
        if not isinstance(data, list):
            return jsonify({"error": "Data must be a list"}), 400

        wb = Workbook()
        ws = wb.active
        ws.title = "Weekly Report"

        # Title row - Fixed: Changed F1 to E1 to match 5 columns
        ws.merge_cells('A1:E1')
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

        # Data rows with validation
        for row in data:
            if not isinstance(row, dict):
                continue
            ws.append([
                row.get("activity_id", ""),
                row.get("site_id", ""),
                row.get("activity_date", ""),
                row.get("before_photo", ""),
                row.get("after_photo", "")
            ])

        # Auto-width - Fixed: Handle merged cells properly
        for col in ws.columns:
            max_length = 0
            column = None
            for cell in col:
                if hasattr(cell, 'column_letter'):
                    column = cell.column_letter
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            if column:
                ws.column_dimensions[column].width = max_length + 2

        # Save to memory
        output = BytesIO()
        try:
            wb.save(output)
            excel_data = output.getvalue()
        finally:
            # Properly close the BytesIO resource
            output.close()

        # Encode to base64
        encoded = base64.b64encode(excel_data).decode("utf-8")
        filename = f"Weekly_Report_{datetime.date.today().isoformat()}.xlsx"

        return jsonify({
            "fileData": encoded,
            "fileName": filename
        })
    
    except Exception as e:
        return jsonify({"error": f"Failed to generate Excel: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
