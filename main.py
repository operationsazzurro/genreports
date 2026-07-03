from flask import Flask, request, send_file, jsonify
from pest_report import pest_report_fn
from flask_cors import CORS
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.pagebreak import Break
import base64
import datetime
import requests
from PIL import Image as PILImage
import tempfile
from io import BytesIO
import os
from groupdocs_conversion_cloud import Configuration, ConvertApi, ConvertDocumentDirectRequest

app = Flask(__name__)
CORS(app)

# Setup conversion API
config = Configuration("c83e270b-cc7a-425d-a255-332e39c2df83",
                       "5d68a4ed789b847a256e4c5fb58625c6")

convert_api = ConvertApi(config)


@app.route('/')
def home():
    return "API is running successfully!"


@app.route('/generate_excel', methods=['POST'])
def generate_excel():

    payload = request.get_json(force=True, silent=True)
    if not payload:
        return jsonify({"error": "Invalid JSON payload"}), 400
    report_format = payload.get("format", "excel").lower()
    data_dict = payload.get("data", [])
    if isinstance(data_dict, dict):
        keys = list(data_dict.keys())
        num_rows = len(next(iter(data_dict.values()))) if data_dict else 0
        data = [{
            key: data_dict[key][i]
            for key in keys
        } for i in range(num_rows)]
    else:
        data = data_dict

    wb = Workbook()

    # ========== COVER PAGE ==========
    cover = wb.active
    cover.title = "Cover Page"

    # Company Logo

    try:

        azzurro_logo = XLImage("azzurro.png")
        azzurro_logo.width, azzurro_logo.height = 108, 24
        cover.add_image(azzurro_logo, "A2")
    except Exception as e:
        print(f"Logo load failed: {e}")

    # Subcontractor Logo

    try:

        client_logo = XLImage("imdaad.png")
        client_logo.width, client_logo.height = 64, 71
        cover.add_image(client_logo, "K1")
    except Exception as e:
        print(f"Logo load failed: {e}")

    # Main Client Logo

    try:

        mainclient_logo = XLImage("awqaf.png")
        mainclient_logo.width, mainclient_logo.height = 319, 124
        cover.add_image(mainclient_logo, "D8")
    except Exception as e:
        print(f"Logo load failed: {e}")

    # Title and Metadata
    cover.merge_cells('A17:K17')
    site_title = cover['A17']
    site_title.value = "GAIAE AD PACKAGE 3 - MOSQUES"
    site_title.font = Font(size=15, bold=True, color="808080")
    site_title.alignment = Alignment(horizontal="center", vertical="center")
    cover.merge_cells('A21:K21')
    title = cover['A21']
    title.value = "WEEKLY ACTIVITY  REPORT"
    title.font = Font(size=20, bold=True)
    title.alignment = Alignment(horizontal="center", vertical="center")

    cover.merge_cells('A23:K23')
    date_cell = cover['A23']
    date_cell.value = f"Generated on: {datetime.date.today().strftime('%d-%b-%Y')}"
    date_cell.alignment = Alignment(horizontal="center")

    #  footer.font = Font(italic=True, color="808080")

    cover.paperSize = cover.PAPERSIZE_A4
    cover.page_setup.orientation = cover.ORIENTATION_PORTRAIT
    cover.page_margins.left = 0.5
    cover.page_margins.right = 0.5
    cover.page_margins.top = 0.5
    cover.page_margins.bottom = 0.5
    # ✅ Force all columns to fit on one page width
    cover.page_setup.fitToWidth = 1
    cover.page_setup.fitToHeight = 0  # unlimited vertical pages
    cover.sheet_properties.pageSetUpPr.fitToPage = True

    cover.page_setup.centerHorizontally = True
    cover.page_setup.centerVertically = True

    # ========== DATA SHEET ==========
    ws = wb.create_sheet("Weekly Report")

    # Page setup
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5
    # ✅ Force all columns to fit on one page width
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0  # unlimited vertical pages
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Title row
    ws.merge_cells('A1:E1')
    title_cell = ws['A1']
    title_cell.value = "Weekly Activity Report"
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    title_cell.fill = PatternFill(start_color="B7DEE8",
                                  end_color="B7DEE8",
                                  fill_type="solid")
    ws.row_dimensions[1].height = 25

    # Headers
    headers = [
        "Finance Code", "Mosque Name", "Activity", "Before Photo",
        "After Photo"
    ]
    ws.append(headers)

    # Add borders to outer cells of Data sheet
    # === CONFIG ===

    rows_per_page = 8  # total rows INCLUDING header height
    header_rows = 2  # number of header rows
    max_col = 5
    header_height = 25
    data_row_height = 224

    thick = Side(style="thick", color="000000")

    # Count total data rows (excluding header)
    total_data_rows = len(data)
    data_rows_per_page = rows_per_page - header_rows

    # Total number of pages
    pages = (total_data_rows // data_rows_per_page) + (
        1 if total_data_rows % data_rows_per_page else 0)

    for page in range(pages):
        # ---- Calculate this page’s row range ----
        if page == 0:
            page_start_row = page * (data_rows_per_page) + 1
            page_end_row = page_start_row + rows_per_page - 1
        else:
            page_start_row = page * (rows_per_page) + 1
            page_end_row = page_start_row + data_rows_per_page - 1

# === 1️⃣ Apply Row Heights ===
        for r in range(page_start_row, page_end_row + 1):
            # Count header rows properly on each page
            local_row_index = (r - page_start_row) + 1

            if local_row_index <= header_rows:
                ws.row_dimensions[r].height = header_height
            else:
                ws.row_dimensions[r].height = data_row_height

        # === 2️⃣ Draw Page Border ===
        for row in range(page_start_row, page_end_row + 1):
            for col in range(1, max_col + 1):
                left = thick if col == 1 else None
                right = thick if col == max_col else None
                top = thick if row == page_start_row else None
                bottom = thick if row == page_end_row else None

                if (page == 0):
                    ws.cell(row=1, column=col).border = Border(left=left,
                                                               right=right,
                                                               top=top,
                                                               bottom=bottom)
                else:
                    ws.cell(row=row, column=col).border = Border(left=left,
                                                                 right=right,
                                                                 bottom=bottom)
    #Bonus to add empty rows to fill the page
    #total_rows_needed = pages * rows_per_page
    #for r in range(ws.max_row + 1, total_rows_needed + 1):
    #    ws.append([""] * max_col)

    # === 4️⃣ Repeat Header on Each Printed Page ===
    ws.print_title_rows = f"$1:${header_rows}"

    for col in ws.iter_cols(min_row=2,
                            max_row=2,
                            min_col=1,
                            max_col=len(headers)):
        for cell in col:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = PatternFill(start_color="D9EAD3",
                                    end_color="D9EAD3",
                                    fill_type="solid")
            cell.border = Border(left=Side(style="thin"),
                                 right=Side(style="thin"),
                                 top=Side(style="thin"),
                                 bottom=Side(style="thin"))

    # Column width preset
    column_widths = [18, 30, 40, 42, 42]
    for i, width in enumerate(column_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Data rows
    row_index = 3
    for row in data:
        ws.cell(row=row_index, column=1, value=row.get("fc", ""))
        ws.cell(row=row_index, column=2, value=row.get("name", ""))
        ws.cell(row=row_index, column=3, value=row.get("activity", ""))

        # Enable word wrap for first 3 columns
        for i in range(1, 4):
            ws.cell(row=row_index,
                    column=i).alignment = Alignment(wrap_text=True,
                                                    vertical="center",
                                                    horizontal="center")

#safe_image_from_url

        def safe_image_from_url(url, width=305, height=305):
            try:
                response = requests.get(url, timeout=5)
                if response.status_code != 200:
                    return None
                # Always convert image to PNG to avoid .mpo errors
                pil_img = PILImage.open(BytesIO(response.content))
                converted = BytesIO()
                pil_img.save(converted, format="PNG")
                converted.seek(0)

                img = XLImage(converted)
                img.width, img.height = width, height
                return img
            except Exception as e:
                print(f"Image load failed for {url}: {e}")
                return None

        # Before Photo
        before_img = safe_image_from_url(row.get("before_file_url", ""))
        if before_img:
            ws.add_image(before_img, f"D{row_index}")
        else:
            ws.cell(row=row_index, column=4, value="Image load failed")

        # After Photo
        after_img = safe_image_from_url(row.get("after_file_url", ""))
        if after_img:
            ws.add_image(after_img, f"E{row_index}")
        else:
            ws.cell(row=row_index, column=5, value="Image load failed")

        ws.row_dimensions[row_index].height = 224
        row_index += 1

    # Save to memory

#   output = BytesIO()
#   wb.save(output)
#   output.seek(0)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
        wb.save(tmp_xlsx.name)
        tmp_xlsx_path = tmp_xlsx.name
        tmp_xlsx.close

    if report_format == "excel":
        filename = "Weekly_Report.xlsx"

        return send_file(
            tmp_xlsx_path,
            download_name=filename,
            as_attachment=True,
            mimetype=
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # If user wants PDF via GroupDocs
    # Use Direct conversion (in-memory)
    request_conv = ConvertDocumentDirectRequest("pdf", tmp_xlsx_path)
    pdf_response = convert_api.convert_document_direct(request_conv)

    # If the API returns a path instead of bytes
    if isinstance(pdf_response, str) and os.path.exists(pdf_response):
        print(f"GroupDocs returned file path: {pdf_response}")
        with open(pdf_response, "rb") as f:
            pdf_bytes = f.read()
    elif isinstance(pdf_response, bytes):
        pdf_bytes = pdf_response
    elif isinstance(pdf_response, str):
        # Could be Base64
        try:
            import base64
            pdf_bytes = base64.b64decode(pdf_response)
        except Exception:
            raise TypeError(
                "Response string is not valid Base64 and not a path.")
    else:
        raise TypeError(f"Unexpected response type: {type(pdf_response)}")

    pdf_stream = BytesIO(pdf_bytes)
    pdf_stream.seek(0)
    print(f"PDF conversion successful:{tmp_xlsx_path}")
    print("PDF stream position:", pdf_stream.tell())
    print("PDF header preview:", pdf_bytes[:10])
    return send_file(pdf_stream,
                     download_name="report.pdf",
                     as_attachment=True,
                     mimetype="application/pdf")


#  except Exception as e:
#      return jsonify({"error": f"Conversion failed: {str(e)}"}), 500
#  finally:
# Clean up temporary file
#      if 'tmp_xlsx_path' in locals() and os.path.exists(tmp_xlsx_path):
#          os.remove(tmp_xlsx_path)


@app.route('/generate_excel_pest', methods=['POST'])
def generate_excel_pest():

    payload = request.get_json(force=True, silent=True)
    if not payload:
        return jsonify({"error": "Invalid JSON payload"}), 400

    data_dict = payload.get("data", [])
    if isinstance(data_dict, dict):
        keys = list(data_dict.keys())
        num_rows = len(next(iter(data_dict.values()))) if data_dict else 0
        data = [{
            key: data_dict[key][i]
            for key in keys
        } for i in range(num_rows)]
    else:
        data = data_dict

    path = pest_report_fn(data)

    return send_file(
        path,
        as_attachment=True,
        mimetype=
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
