import subprocess
import os


def convert_excel_to_pdf(excel_path):
    output_dir = os.path.dirname(excel_path)

    subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to",
            "pdf",
            excel_path,
            "--outdir",
            output_dir,
        ],
        check=True,
    )

    pdf_path = os.path.splitext(excel_path)[0] + ".pdf"

    if not os.path.exists(pdf_path):
        raise Exception("PDF was not created")

    return pdf_path
