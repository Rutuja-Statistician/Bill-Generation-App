import io
import os
import zipfile
import tempfile
import subprocess
import shutil
from datetime import date

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from num2words import num2words


from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)   

# ---------------------------------------------------------
# Configuration
# ---------------------------------------------------------

APP_DIR = os.path.dirname(os.path.abspath(__file__))

TEMPLATE_FILE = os.path.join(
    APP_DIR,
    "bill_template.xlsx"
)

COUNTER_FILE = os.path.join(
    APP_DIR,
    "bill_counter.txt"
)

MEDICAL_TRANSCRIPTION_CHARGES = 14000
FIXED_TARGET_PM = 0


# ---------------------------------------------------------
# Find LibreOffice
# ---------------------------------------------------------

def find_libreoffice():

    possible_paths = [
        # Windows - normal installation
        r"C:\Program Files\LibreOffice\program\soffice.exe",

        # Windows - 32 bit installation
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",

        # If LibreOffice is available in PATH
        "soffice",
        "libreoffice",
    ]

    for path in possible_paths:

        if os.path.isfile(path):
            return path

        if shutil.which(path):
            return path

    return None


# ---------------------------------------------------------
# Financial Year
# ---------------------------------------------------------

def get_financial_year(dt):

    if dt.month >= 4:

        return (
            f"{str(dt.year)[-2:]}-"
            f"{str(dt.year + 1)[-2:]}"
        )

    return (
        f"{str(dt.year - 1)[-2:]}-"
        f"{str(dt.year)[-2:]}"
    )


# ---------------------------------------------------------
# Generate Unique Bill Number
# ---------------------------------------------------------

def get_next_bill_number():

    current_number = 0

    if os.path.exists(COUNTER_FILE):

        try:

            with open(
                COUNTER_FILE,
                "r",
                encoding="utf-8"
            ) as f:

                current_number = int(
                    f.read().strip() or "0"
                )

        except (ValueError, OSError):

            current_number = 0

    next_number = current_number + 1

    financial_year = get_financial_year(
        date.today()
    )

    with open(
        COUNTER_FILE,
        "w",
        encoding="utf-8"
    ) as f:

        f.write(str(next_number))

    return (
        f"{next_number:03d}/"
        f"{financial_year}"
    )


# ---------------------------------------------------------
# Amount in Words
# ---------------------------------------------------------

def amount_in_words(amount):

    words = num2words(
        int(amount),
        lang="en_IN"
    )

    return (
        f"Rupees {words.title()} Only."
    )


# ---------------------------------------------------------
# Safe Filename
# ---------------------------------------------------------

def clean_filename(name):

    name = str(name).strip()

    invalid_chars = '<>:"/\\|?*'

    for char in invalid_chars:

        name = name.replace(
            char,
            "_"
        )

    return name or "Employee"


# ---------------------------------------------------------
# Create Completed Excel Bill
# ---------------------------------------------------------

def create_bill_excel(
    employee,
    bill_no,
    bill_date,
    output_file
):

    if not os.path.exists(TEMPLATE_FILE):

        raise FileNotFoundError(
            "bill_template.xlsx was not found."
        )

    wb = load_workbook(
        TEMPLATE_FILE
    )

    ws = wb["Sheet1"]

    # ---------------------------------------------
    # Employee information
    # ---------------------------------------------

    name = str(
        employee.get("Name", "")
    ).strip()

    address = str(
        employee.get("Address", "")
    ).strip()

    sign = str(
        employee.get("Sign", "")
    ).strip()

    if not sign:

        sign = name


    # ---------------------------------------------
    # Employee Name
    # ---------------------------------------------

    ws["A4"] = (
        f"NAME: {name}"
    )


    # ---------------------------------------------
    # Employee Address
    # ---------------------------------------------

    ws["A5"] = (
        f"Address: {address}"
    )


    # ---------------------------------------------
    # Bill Date
    # ---------------------------------------------

    ws["G8"] = bill_date

    ws["G8"].number_format = (
        r'mmmm\, dd\ yyyy'
    )


    # ---------------------------------------------
    # Bill Number
    # ---------------------------------------------

    ws["G9"] = bill_no


    # ---------------------------------------------
    # Medical Transcription Charges
    # ---------------------------------------------

    ws["G19"] = (
        MEDICAL_TRANSCRIPTION_CHARGES
    )


    # ---------------------------------------------
    # Fixed Target P.M.
    # ---------------------------------------------

    ws["G20"] = (
        FIXED_TARGET_PM
    )


    # Other amount rows
    ws["G21"] = 0
    ws["G22"] = 0
    ws["G23"] = 0
    ws["G24"] = 0


    # ---------------------------------------------
    # Total
    # ---------------------------------------------

    ws["G25"] = (
        "=SUM(G19:G24)"
    )


    # ---------------------------------------------
    # Amount in words
    # ---------------------------------------------

    total_amount = (
        MEDICAL_TRANSCRIPTION_CHARGES
        + FIXED_TARGET_PM
    )

    ws["A26"] = amount_in_words(
        total_amount
    )


    # ---------------------------------------------
    # Signature
    # ---------------------------------------------

    ws["A31"] = sign

    ws["A32"] = "Contractor"


    # ---------------------------------------------
    # Excel calculation settings
    # ---------------------------------------------

    try:

        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"

    except AttributeError:

        pass


    # ---------------------------------------------
    # Save completed Excel
    # ---------------------------------------------

    wb.save(output_file)


# ---------------------------------------------------------
# Convert Excel → PDF
# ---------------------------------------------------------

def convert_excel_to_pdf(
    excel_file,
    output_directory
):

    libreoffice = find_libreoffice()

    if not libreoffice:

        raise RuntimeError(
            "LibreOffice was not found.\n\n"
            "Please install LibreOffice and make sure "
            "soffice.exe is available."
        )


    command = [
        libreoffice,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        output_directory,
        excel_file,
    ]


    result = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        timeout=120
    )


    if result.returncode != 0:

        raise RuntimeError(
            "Excel to PDF conversion failed.\n\n"
            f"{result.stderr}"
        )


    pdf_file = os.path.join(
        output_directory,
        os.path.splitext(
            os.path.basename(excel_file)
        )[0] + ".pdf"
    )


    if not os.path.exists(pdf_file):

        raise RuntimeError(
            "PDF was not created by LibreOffice."
        )


    return pdf_file


# ---------------------------------------------------------
# Generate PDF Bill
# ---------------------------------------------------------

from xml.sax.saxutils import escape
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


def create_pdf_bill(employee, bill_no, bill_date):
    """Generate PDF matching the supplied bill screenshot."""

    buffer = io.BytesIO()

    page_width, page_height = A4
    margin = 10 * mm
    content_width = page_width - 2 * margin

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=5 * mm,
        bottomMargin=5 * mm,
        title=f"Invoice {bill_no}",
    )

    name = escape(str(employee.get("Name", "")).strip())
    address = escape(str(employee.get("Address", "")).strip())
    sign = str(employee.get("Sign", "")).strip()
    if not sign:
        sign = str(employee.get("Name", "")).strip()
    sign = escape(sign)

    total = (
        MEDICAL_TRANSCRIPTION_CHARGES
        + FIXED_TARGET_PM
    )

    # Styles
    normal = ParagraphStyle(
        "BillNormal",
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
        alignment=TA_LEFT,
        spaceAfter=0,
    )

    bold = ParagraphStyle(
        "BillBold",
        parent=normal,
        fontName="Helvetica-Bold",
    )

    centered = ParagraphStyle(
        "BillCentered",
        parent=normal,
        alignment=TA_CENTER,
    )

    right = ParagraphStyle(
        "BillRight",
        parent=normal,
        alignment=TA_RIGHT,
    )

    story = []

    # -------------------------------------------------
    # Employee name and address at the top
    # -------------------------------------------------
    employee_header = Table(
        [
            [
                Paragraph(f"<b>NAME: {name}</b>", centered)
            ],
            [
                Paragraph(address, centered)
            ],
        ],
        colWidths=[content_width],
        rowHeights=[9 * mm, 9 * mm],
    )

    employee_header.setStyle(
        TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.9, colors.black),
            ("LINEBELOW", (0, 0), (-1, 0), 0, colors.white),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ])
    )

    story.append(employee_header)

    # -------------------------------------------------
    # Invoice title: centered in its own row
    # -------------------------------------------------
    invoice_title = Table(
        [[Paragraph("INVOICE", centered)]],
        colWidths=[content_width],
        rowHeights=[7 * mm],
    )

    invoice_title.setStyle(
        TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.8, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ])
    )

    story.append(invoice_title)

    # -------------------------------------------------
    # Recipient address on left; date and bill number
    # on the right, as shown in the screenshot
    # -------------------------------------------------
    recipient = (
        "<b>To,</b><br/>"
        "<b>DOVLIN ONE PRIVATE LIMITED</b><br/>"
        "Unit No. III E, Phase 2, 3rd Floor<br/>"
        "Carnival Infopark, Kakkanad,<br/>"
        "Kochi - 682042.<br/>"
        "KERALA"
    )

    date_text = bill_date.strftime("%B %d %Y")

    recipient_cell = Paragraph(recipient, normal)

    details_cell = Table(
        [
            [
                Paragraph("<b>Date:</b>", right),
                Paragraph(escape(date_text), normal),
            ],
            [
                Paragraph("<b>BILL No.</b>", right),
                Paragraph(escape(str(bill_no)), normal),
            ],
        ],
        colWidths=[22 * mm, 32 * mm],
    )

    details_cell.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ])
    )

    recipient_table = Table(
        [[recipient_cell, details_cell]],
        colWidths=[
            content_width - 54 * mm,
            54 * mm,
        ],
        rowHeights=[43 * mm],
    )

    recipient_table.setStyle(
        TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.8, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 1),
            ("RIGHTPADDING", (0, 0), (-1, -1), 1),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ])
    )

    story.append(recipient_table)

    # -------------------------------------------------
    # Particulars and amount table
    # -------------------------------------------------
    rows = [
        [
            Paragraph("Particulars", centered),
            Paragraph("Amount (Rs)", normal),
        ],
        [
            Paragraph("Medical Transcription Charges", normal),
            Paragraph(f"{MEDICAL_TRANSCRIPTION_CHARGES}", normal),
        ],
        [
            Paragraph("FIXED TARGET P.M.", normal),
            Paragraph(f"{FIXED_TARGET_PM}", normal),
        ],
        ["", Paragraph("0", normal)],
        ["", Paragraph("0", normal)],
        ["", Paragraph("0", normal)],
        ["", Paragraph("0", normal)],
        [
            Paragraph("<b>Total</b>", bold),
            Paragraph(
                f"<b>Rs.          {total:,.0f}</b>",
                right,
            ),
        ],
    ]

    amounts = Table(
        rows,
        colWidths=[
            content_width - 37 * mm,
            37 * mm,
        ],
        rowHeights=[
            7 * mm,
            10 * mm,
            10 * mm,
            7 * mm,
            7 * mm,
            7 * mm,
            7 * mm,
            7 * mm,
        ],
    )

    amounts.setStyle(
        TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.8, colors.black),
            ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.black),
            ("LINEBEFORE", (1, 0), (1, -1), 0.8, colors.black),
            ("LINEABOVE", (0, -1), (-1, -1), 0.8, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (0, 0), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ])
    )

    story.append(amounts)

    # -------------------------------------------------
    # Amount in words
    # -------------------------------------------------
    story.append(
        Table(
            [[Paragraph(escape(amount_in_words(total)), normal)]],
            colWidths=[content_width],
            rowHeights=[8 * mm],
            style=TableStyle([
                ("BOX", (0, 0), (-1, -1), 0.8, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ]),
        )
    )

    # -------------------------------------------------
    # Signature and contractor label near bottom
    # -------------------------------------------------
    story.append(Spacer(1, 24 * mm))

    signature = Table(
        [
            [
                Paragraph(f"{sign}<br/>Contractor", normal),
            ]
        ],
        colWidths=[content_width],
        rowHeights=[12 * mm],
    )

    signature.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ])
    )

    story.append(signature)

    doc.build(story)

    return buffer.getvalue()

# ---------------------------------------------------------
# Streamlit Configuration
# ---------------------------------------------------------

st.set_page_config(
    page_title="Employee Bill Generator",
    page_icon="🧾",
    layout="wide"
)


# ---------------------------------------------------------
# Header
# ---------------------------------------------------------

st.title(
    "🧾 Employee Bill Generator"
)

st.write(
    "Upload the employee Excel file and generate "
    "individual PDF bills using the predefined bill format."
)


# st.info(
#     "The application uses bill_template.xlsx as the "
#     "bill format. The completed Excel file is converted "
#     "to PDF automatically."
# )

st.info(
    "Upload an employee Excel file to generate individual PDF bills. "
    "PDFs are generated directly without Excel conversion or LibreOffice."
)

# ---------------------------------------------------------
# Upload Employee Excel
# ---------------------------------------------------------

uploaded_file = st.file_uploader(
    "Upload Employee Excel File",
    type=["xlsx", "xls"]
)


if uploaded_file is not None:

    try:

        # ---------------------------------------------
        # Read Excel
        # ---------------------------------------------

        if uploaded_file.name.lower().endswith(
            ".xlsx"
        ):

            df = pd.read_excel(
                uploaded_file,
                engine="openpyxl"
            )

        else:

            df = pd.read_excel(
                uploaded_file
            )


        # Clean column names

        df.columns = [
            str(column).strip()
            for column in df.columns
        ]


        # ---------------------------------------------
        # Required columns
        # ---------------------------------------------

        required_columns = [
            "Name",
            "Address",
            "Sign"
        ]


        missing_columns = [
            column
            for column in required_columns
            if column not in df.columns
        ]


        if missing_columns:

            st.error(
                "Missing required columns: "
                + ", ".join(missing_columns)
            )

            st.stop()


        # Remove completely blank rows

        df = df.dropna(
            how="all"
        ).copy()


        # ---------------------------------------------
        # Preview
        # ---------------------------------------------

        st.subheader(
            "Employee Data"
        )

        st.dataframe(
            df,
            use_container_width=True
        )


        # ---------------------------------------------
        # Bill information
        # ---------------------------------------------

        bill_date = date.today()


        col1, col2, col3 = st.columns(3)


        with col1:

            st.metric(
                "Employees",
                len(df)
            )


        with col2:

            st.metric(
                "Medical Transcription",
                f"₹{MEDICAL_TRANSCRIPTION_CHARGES:,}"
            )


        with col3:

            st.metric(
                "Fixed Target P.M.",
                f"₹{FIXED_TARGET_PM:,}"
            )


        st.write(
            f"**Bill Date:** "
            f"{bill_date.strftime('%d-%m-%Y')}"
        )


        # ---------------------------------------------
        # Generate Bills
        # ---------------------------------------------

        if st.button( "Generate PDF Bills", type="primary"):

            generated_files = []

            preview_rows = []


            progress = st.progress(0)

            status = st.empty()


            # -----------------------------------------
            # Generate bill for every employee
            # -----------------------------------------

            total_employees = len(df)


            for index, (_, employee) in enumerate(
                df.iterrows(),
                start=1
            ):


                name = str(
                    employee.get(
                        "Name",
                        ""
                    )
                ).strip()


                if not name:

                    continue


                # Generate unique bill number

                bill_no = get_next_bill_number()


                # Generate PDF

                pdf_bytes = create_pdf_bill(
                    employee,
                    bill_no,
                    bill_date
                )


                filename = (
                    f"BILL "
                    f"{bill_no.replace('/', '_')} "
                    f"{clean_filename(name)}.pdf"
                )


                generated_files.append(
                    (
                        filename,
                        pdf_bytes
                    )
                )


                preview_rows.append(
                    {
                        "Employee": name,
                        "Bill No.": bill_no,
                        "Date": bill_date.strftime(
                            "%d-%m-%Y"
                        ),
                        "Medical Transcription Charges":
                            MEDICAL_TRANSCRIPTION_CHARGES,
                        "Fixed Target P.M.":
                            FIXED_TARGET_PM,
                        "Total":
                            MEDICAL_TRANSCRIPTION_CHARGES
                            + FIXED_TARGET_PM
                    }
                )


                progress.progress(
                    index / total_employees
                )


                status.write(
                    f"Generating bill "
                    f"{index} of "
                    f"{total_employees}: "
                    f"{name}"
                )


            # -----------------------------------------
            # Generated successfully
            # -----------------------------------------

            if generated_files:

                st.success(
                    f"{len(generated_files)} "
                    f"PDF bill(s) generated successfully."
                )


                # -------------------------------------
                # Preview table
                # -------------------------------------

                st.subheader(
                    "Generated Bills"
                )


                st.dataframe(
                    pd.DataFrame(
                        preview_rows
                    ),
                    use_container_width=True
                )


                # -------------------------------------
                # Create ZIP
                # -------------------------------------

                zip_buffer = io.BytesIO()


                with zipfile.ZipFile(
                    zip_buffer,
                    mode="w",
                    compression=zipfile.ZIP_DEFLATED
                ) as zf:


                    for filename, content in generated_files:

                        zf.writestr(
                            filename,
                            content
                        )


                zip_buffer.seek(0)


                # -------------------------------------
                # Download all
                # -------------------------------------

                st.download_button(
                    label="⬇️ Download All PDF Bills (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=(
                        f"Employee_Bills_"
                        f"{bill_date.strftime('%Y%m%d')}.zip"
                    ),
                    mime="application/zip",
                    type="primary"
                )


                # -------------------------------------
                # Individual PDF downloads
                # -------------------------------------

                st.subheader(
                    "Individual PDF Bills"
                )


                for filename, content in generated_files:

                    st.download_button(
                        label=f"⬇️ {filename}",
                        data=content,
                        file_name=filename,
                        mime="application/pdf",
                        key=f"download_{filename}"
                    )


            status.empty()

            progress.empty()


    except Exception as e:

        st.error(
            f"Error processing the file: {e}"
        )








# Excel format

# import io
# import os
# import zipfile
# from datetime import date, datetime
# from copy import copy

# import openpyxl
# import pandas as pd
# import streamlit as st
# from openpyxl import load_workbook
# from num2words import num2words


# # ---------------------------------------------------------
# # Configuration
# # ---------------------------------------------------------
# APP_DIR = os.path.dirname(os.path.abspath(__file__))
# TEMPLATE_FILE = os.path.join(APP_DIR, "bill_template.xlsx")
# COUNTER_FILE = os.path.join(APP_DIR, "bill_counter.txt")

# MEDICAL_TRANSCRIPTION_CHARGES = 14000
# FIXED_TARGET_PM = 0


# # ---------------------------------------------------------
# # Utility functions
# # ---------------------------------------------------------
# def get_financial_year(dt):
#     """Return Indian financial year, e.g. 2026-27."""
#     if dt.month >= 4:
#         return f"{str(dt.year)[-2:]}-{str(dt.year + 1)[-2:]}"
#     return f"{str(dt.year - 1)[-2:]}-{str(dt.year)[-2:]}"


# def get_next_bill_number():
#     """
#     Generate a sequential unique bill number in the format:
#     001/26-27, 002/26-27, ...
#     """
#     current_number = 0

#     if os.path.exists(COUNTER_FILE):
#         try:
#             with open(COUNTER_FILE, "r", encoding="utf-8") as f:
#                 current_number = int(f.read().strip() or "0")
#         except (ValueError, OSError):
#             current_number = 0

#     next_number = current_number + 1
#     fy = get_financial_year(date.today())

#     with open(COUNTER_FILE, "w", encoding="utf-8") as f:
#         f.write(str(next_number))

#     return f"{next_number:03d}/{fy}"


# def amount_in_words(amount):
#     """Convert amount to words."""
#     words = num2words(int(amount), lang="en_IN")
#     return f"Rupees {words.title()} Only."


# def clean_filename(name):
#     """Make employee name safe for an Excel filename."""
#     name = str(name).strip()
#     invalid_chars = '<>:"/\\|?*'
#     for char in invalid_chars:
#         name = name.replace(char, "_")
#     return name or "Employee"


# def create_bill(employee, bill_no, bill_date):
#     """
#     Create one bill using bill_template.xlsx and return Excel bytes.
#     """
#     if not os.path.exists(TEMPLATE_FILE):
#         raise FileNotFoundError(
#             f"Template file not found: {TEMPLATE_FILE}. "
#             "Place your bill format file in the same folder as app.py "
#             "and rename it to bill_template.xlsx."
#         )

#     wb = load_workbook(TEMPLATE_FILE)
#     ws = wb["Sheet1"]

#     name = str(employee.get("Name", "")).strip()
#     address = str(employee.get("Address", "")).strip()
#     sign = str(employee.get("Sign", "")).strip()

#     if not sign:
#         sign = name

#     # Employee details
#     ws["A4"] = f"NAME: {name}"
#     ws["A5"] = f"Address: {address}"

#     # Automatically generated date and unique bill number
#     ws["G8"] = bill_date
#     ws["G8"].number_format = r'mmmm\, dd\ yyyy'
#     ws["G9"] = bill_no

#     # Charges
#     ws["G19"] = MEDICAL_TRANSCRIPTION_CHARGES

#     # Fixed Target P.M. = 0
#     # The supplied template has the amount area for this section in G20.
#     ws["G20"] = FIXED_TARGET_PM
#     ws["G21"] = 0
#     ws["G22"] = 0
#     ws["G23"] = 0
#     ws["G24"] = 0

#     # Total
#     ws["G25"] = "=SUM(G19:G24)"

#     # Amount in words
#     ws["A26"] = amount_in_words(MEDICAL_TRANSCRIPTION_CHARGES + FIXED_TARGET_PM)

#     # Signature name above "Contractor"
#     ws["A31"] = sign
#     ws["A32"] = "Contractor"

#     # Keep the workbook ready for Excel to recalculate formulas.
#     try:
#         wb.calculation.fullCalcOnLoad = True
#         wb.calculation.forceFullCalc = True
#         wb.calculation.calcMode = "auto"
#     except AttributeError:
#         pass

#     output = io.BytesIO()
#     wb.save(output)
#     output.seek(0)
#     return output.getvalue()


# # ---------------------------------------------------------
# # Streamlit UI
# # ---------------------------------------------------------
# st.set_page_config(
#     page_title="Employee Bill Generator",
#     page_icon="🧾",
#     layout="wide",
# )

# st.title("🧾 Employee Bill Generator")
# st.write(
#     "Upload the employee Excel file. The application will generate one bill "
#     "for each employee using the supplied bill format."
# )

# st.info(
#     "Before running the app, keep your bill format Excel file in the same "
#     "folder as app.py and rename it to: bill_template.xlsx"
# )

# uploaded_file = st.file_uploader(
#     "Upload Employee Excel File",
#     type=["xlsx", "xls"],
# )

# if uploaded_file is not None:
#     try:
#         if uploaded_file.name.lower().endswith(".xlsx"):
#             df = pd.read_excel(uploaded_file, engine="openpyxl")
#         else:
#             df = pd.read_excel(uploaded_file)

#         df.columns = [str(c).strip() for c in df.columns]

#         required_columns = ["Name", "Address", "Sign"]
#         missing = [c for c in required_columns if c not in df.columns]

#         if missing:
#             st.error(
#                 "Missing required column(s): "
#                 + ", ".join(missing)
#                 + ". Required columns are: Name, Address, Sign."
#             )
#             st.stop()

#         df = df.dropna(how="all").copy()

#         st.subheader("Employee Data")
#         st.dataframe(df, use_container_width=True)

#         st.write(f"**Employees found:** {len(df)}")
#         st.write(
#             f"**Bill Date:** {date.today().strftime('%d-%m-%Y')}  |  "
#             f"**Medical Transcription Charges:** ₹{MEDICAL_TRANSCRIPTION_CHARGES:,}  |  "
#             f"**Fixed Target P.M.:** ₹{FIXED_TARGET_PM:,}"
#         )

#         if st.button("Generate Bills", type="primary"):
#             if not os.path.exists(TEMPLATE_FILE):
#                 st.error(
#                     "bill_template.xlsx was not found. Put the bill format "
#                     "file in the same folder as app.py."
#                 )
#                 st.stop()

#             bill_date = date.today()
#             generated_files = []
#             preview_rows = []

#             progress = st.progress(0)
#             status = st.empty()

#             for index, (_, employee) in enumerate(df.iterrows(), start=1):
#                 name = str(employee.get("Name", "")).strip()

#                 if not name:
#                     status.warning(f"Skipping row {index}: employee name is blank.")
#                     continue

#                 bill_no = get_next_bill_number()
#                 excel_bytes = create_bill(employee, bill_no, bill_date)

#                 filename = f"BILL {bill_no.replace('/', '_')} {clean_filename(name)}.xlsx"
#                 generated_files.append((filename, excel_bytes))

#                 preview_rows.append(
#                     {
#                         "Employee": name,
#                         "Bill No.": bill_no,
#                         "Date": bill_date.strftime("%d-%m-%Y"),
#                         "Medical Transcription Charges": MEDICAL_TRANSCRIPTION_CHARGES,
#                         "Fixed Target P.M.": FIXED_TARGET_PM,
#                         "Total": MEDICAL_TRANSCRIPTION_CHARGES + FIXED_TARGET_PM,
#                     }
#                 )

#                 progress.progress(index / len(df))
#                 status.write(f"Generating bill {index} of {len(df)}: {name}")

#             if generated_files:
#                 # Create one ZIP containing all employee bills.
#                 zip_buffer = io.BytesIO()
#                 with zipfile.ZipFile(
#                     zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED
#                 ) as zf:
#                     for filename, content in generated_files:
#                         zf.writestr(filename, content)

#                 zip_buffer.seek(0)

#                 st.success(f"{len(generated_files)} bill(s) generated successfully.")

#                 st.subheader("Generated Bills")
#                 st.dataframe(pd.DataFrame(preview_rows), use_container_width=True)

#                 st.download_button(
#                     label="⬇️ Download All Bills (ZIP)",
#                     data=zip_buffer.getvalue(),
#                     file_name=f"Employee_Bills_{bill_date.strftime('%Y%m%d')}.zip",
#                     mime="application/zip",
#                     type="primary",
#                 )

#                 st.subheader("Individual Bills")
#                 for filename, content in generated_files:
#                     st.download_button(
#                         label=f"Download {filename}",
#                         data=content,
#                         file_name=filename,
#                         mime=(
#                             "application/vnd.openxmlformats-officedocument."
#                             "spreadsheetml.sheet"
#                         ),
#                         key=f"download_{filename}",
#                     )

#             status.empty()
#             progress.empty()

#     except Exception as e:
#         st.error(f"Error processing the file: {e}")
