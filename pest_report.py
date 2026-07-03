from flask import Flask, request, send_file, jsonify
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
import tempfile, os
from io import BytesIO


def pest_report_fn(data):
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
    cover.merge_cells("A17:K17")
    site_title = cover["A17"]
    site_title.value = "GAIAE AD PACKAGE 3 - MOSQUES"
    site_title.font = Font(size=15, bold=True, color="808080")
    site_title.alignment = Alignment(horizontal="center", vertical="center")
    cover.merge_cells("A21:K21")
    title = cover["A21"]
    title.value = "PEST CONTROL ACTIVITY REPORT"
    title.font = Font(size=20, bold=True)
    title.alignment = Alignment(horizontal="center", vertical="center")

    cover.merge_cells("A23:K23")
    date_cell = cover["A23"]
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
    ws.merge_cells("A1:E1")
    title_cell = ws["A1"]
    title_cell.value = "Weekly Pest Control Activity Report"
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill(
        start_color="B7DEE8", end_color="B7DEE8", fill_type="solid"
    )
    ws.row_dimensions[1].height = 25

    # Headers
    headers = ["Finance Code", "Mosque Name", "Activity", "Before Photo", "After Photo"]
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
    # total_data_rows = len(data)
    total_data_rows = 3
    data_rows_per_page = rows_per_page - header_rows

    # Total number of pages
    pages = (total_data_rows // data_rows_per_page) + (
        1 if total_data_rows % data_rows_per_page else 0
    )

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

                if page == 0:
                    ws.cell(row=1, column=col).border = Border(
                        left=left, right=right, top=top, bottom=bottom
                    )
                else:
                    ws.cell(row=row, column=col).border = Border(
                        left=left, right=right, bottom=bottom
                    )
    # Bonus to add empty rows to fill the page
    # total_rows_needed = pages * rows_per_page
    # for r in range(ws.max_row + 1, total_rows_needed + 1):
    #    ws.append([""] * max_col)

    # === 4️⃣ Repeat Header on Each Printed Page ===
    ws.print_title_rows = f"$1:${header_rows}"

    for col in ws.iter_cols(min_row=2, max_row=2, min_col=1, max_col=len(headers)):
        for cell in col:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = PatternFill(
                start_color="D9EAD3", end_color="D9EAD3", fill_type="solid"
            )
            cell.border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            )

    # Column width preset
    column_widths = [18, 30, 40, 42, 42]
    for i, width in enumerate(column_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # safe_image_from_url

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

    # Data rows
    row_index = 3
    for row in data:
        image1 = row.get("before_file_url").split(",")[0]
        image2 = row.get("before_file_url").split(",")[1]
        image3 = row.get("before_file_url").split(",")[2]
        image4 = row.get("before_file_url").split(",")[3]

        before_img1 = safe_image_from_url(image1)
        before_img2 = safe_image_from_url(image2)
        before_img3 = safe_image_from_url(image3)
        before_img4 = safe_image_from_url(image4)

        ws.cell(row=row_index, column=1, value=row.get("fc", ""))
        ws.cell(row=row_index, column=2, value=row.get("name", ""))
        ws.cell(row=row_index, column=3, value=row.get("activity", ""))
        if before_img1:
            ws.add_image(before_img1, f"D{row_index}")
        else:
            ws.cell(row=row_index, column=4, value="Image load failed")
        if before_img2:
            ws.add_image(before_img2, f"E{row_index}")
        else:
            ws.cell(row=row_index, column=5, value="Image load failed")

        ws.cell(row=row_index + 1, column=1, value=row.get("fc", ""))
        ws.cell(row=row_index + 1, column=2, value=row.get("name", ""))
        ws.cell(row=row_index + 1, column=3, value=row.get("activity", ""))
        if before_img3:
            ws.add_image(before_img3, f"D{row_index + 1}")
        else:
            ws.cell(row=row_index, column=4, value="Image load failed")
        if before_img4:
            ws.add_image(before_img4, f"E{row_index + 1}")
        else:
            ws.cell(row=row_index, column=5, value="Image load failed")

        #    print("Images are :", image1, ",", image2, ",", image3, ",", image4)
        # Enable word wrap for first 3 columns
        for i in range(1, 4):
            ws.cell(row=row_index, column=i).alignment = Alignment(
                wrap_text=True, vertical="center", horizontal="center"
            )
            ws.cell(row=row_index + 1, column=i).alignment = Alignment(
                wrap_text=True, vertical="center", horizontal="center"
            )

        ws.row_dimensions[row_index].height = 224
        ws.row_dimensions[row_index + 1].height = 224
        row_index += 2

    # Save to memory

    #   output = BytesIO()
    #   wb.save(output)
    #   output.seek(0)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_xlsx:
        wb.save(tmp_xlsx.name)
        tmp_xlsx_path = tmp_xlsx.name
        tmp_xlsx.close
        print("File exists:", os.path.exists(tmp_xlsx_path))
        print("File path:", tmp_xlsx_path)
        print(
            "File Size:",
            os.path.getsize(tmp_xlsx_path) if os.path.exists(tmp_xlsx_path) else "N/A",
        )
        return tmp_xlsx_path
