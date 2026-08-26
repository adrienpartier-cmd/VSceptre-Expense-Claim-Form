"""
Receipt scanner -> Expense Claim Form filler

Workflow:
1. Reads receipt images/PDFs from ./receipts
2. Uses LlamaExtract OCR to identify:
   - Date
   - Merchant
   - Description
   - Expense category
   - Amount
3. Writes data into Expense_Claim_filled.xlsx

IMPORTANT:
- Column widths are NEVER changed.
- Long descriptions wrap onto new lines inside column D.
- E7:E16 always retain a real Excel dropdown list.
- Llama may only choose from the 36 dropdown options.
- H contains the amount.
- Excel itself handles all category totals and grand totals.
- The claim form can be reset after use without changing its layout.
"""

from llama_cloud_services import LlamaExtract
from pydantic import BaseModel, Field, ValidationError
from typing import Optional
from pathlib import Path
from dotenv import load_dotenv

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

import os
import textwrap


# ======================================================================
# PATHS
# ======================================================================

BASE_DIR = Path(__file__).resolve().parent

OUTPUT_FILE = BASE_DIR / "Expense_Claim_filled.xlsx"
RECEIPTS_FOLDER = BASE_DIR / "receipts"

SHEET_NAME = "Form"
DROPDOWN_SHEET = "Dropdown"

FIRST_DATA_ROW = 7
LAST_DATA_ROW = 16


# ======================================================================
# API KEY
# ======================================================================

load_dotenv(BASE_DIR / ".env")

API_KEY = os.getenv("LLAMA_CLOUD_API_KEY")


# ======================================================================
# EXACT DROPDOWN OPTIONS
# ======================================================================

# These are the 36 options in Dropdown!A1:A36.

ALLOWED_CATEGORIES = [
    "Office/Computer Equipment",
    "Furniture & Fixtures",
    "Purchase",
    "Software - Purchase",
    "Misc. Parts - Purchase",
    "Audit Fee",
    "Bank Charges",
    "Business Registration Fee",
    "Consulting and Secretarial Fee",
    "Legal and Professional Fee",
    "Finance Costs",
    "General Office Expense",
    "Insurance - Office",
    "Insurance -Staff (non-sales)",
    "Insurance - Contractor",
    "Print,Stationery, Due & Subs",
    "Postage & Courier",
    "Rent & Management Fee - Office",
    "Rates & Gov Rent - Office",
    "Repairs & Maintenance - Office",
    "Telephone & Broadband - Office",
    "Utilities - Office",
    "MPF (Non-Sales)",
    "Legal Fees",
    "Office Supplies",
    "Selling & Distribution Expense",
    "Advertising &Promotion Expense",
    "Entertainment",
    "Sundry Expense- Office",
    "Travelling - Overseas",
    "Travelling - Local",
    "Training and Seminar Costs",
    "Staff Amenities",
    "Salaries & Allowances (Sales)",
    "Wages & Salaries",
    "/",
]


CATEGORY_PROMPT = """
Choose EXACTLY ONE of the following Excel dropdown options.

Return the wording EXACTLY as written.

Office/Computer Equipment
Furniture & Fixtures
Purchase
Software - Purchase
Misc. Parts - Purchase
Audit Fee
Bank Charges
Business Registration Fee
Consulting and Secretarial Fee
Legal and Professional Fee
Finance Costs
General Office Expense
Insurance - Office
Insurance -Staff (non-sales)
Insurance - Contractor
Print,Stationery, Due & Subs
Postage & Courier
Rent & Management Fee - Office
Rates & Gov Rent - Office
Repairs & Maintenance - Office
Telephone & Broadband - Office
Utilities - Office
MPF (Non-Sales)
Legal Fees
Office Supplies
Selling & Distribution Expense
Advertising &Promotion Expense
Entertainment
Sundry Expense- Office
Travelling - Overseas
Travelling - Local
Training and Seminar Costs
Staff Amenities
Salaries & Allowances (Sales)
Wages & Salaries
/

Classification guidance:

Computers, laptops, monitors, computer hardware and similar equipment
-> Office/Computer Equipment

Furniture, desks, chairs and cabinets
-> Furniture & Fixtures

Software purchases or licences
-> Software - Purchase

Small electronic/computer replacement parts
-> Misc. Parts - Purchase

Office stationery, paper, pens and ordinary supplies
-> Office Supplies

Printing, stationery, subscriptions or membership dues
-> Print,Stationery, Due & Subs

Restaurants, cafes, meals, drinks, food and business meals
-> Entertainment

Hong Kong taxi, Uber, MTR, bus, ferry, parking and local transport
-> Travelling - Local

Flights, overseas taxis, overseas trains and overseas travel
-> Travelling - Overseas

Courses, seminars, conferences, examinations and certifications
-> Training and Seminar Costs

Advertising, promotion and marketing
-> Advertising &Promotion Expense

Courier, postage or delivery charges
-> Postage & Courier

Telephone, mobile service, internet or broadband
-> Telephone & Broadband - Office

Electricity, water or utilities
-> Utilities - Office

Repair or maintenance services
-> Repairs & Maintenance - Office

If no sensible category can be determined from the receipt
-> /

Do NOT create a new category.
"""


# ======================================================================
# EXTRACTION SCHEMA
# ======================================================================

class Receipt(BaseModel):

    merchant_name: str = Field(
        description=(
            "Business/shop/restaurant name printed on the receipt."
        )
    )

    receipt_date: str = Field(
        description=(
            "Transaction date in DD/MM/YYYY format."
        )
    )

    description: str = Field(
        description=(
            "Concise factual description of the items or services purchased. "
            "Include important product names or models visible on the receipt. "
            "Do not invent information."
        )
    )

    total_hkd: float = Field(
        description=(
            "Final amount paid as printed on the receipt. "
            "Do not perform currency conversion."
        )
    )

    currency: str = Field(
        description=(
            "Three-letter currency code, for example HKD, USD, GBP or TWD."
        )
    )

    category: str = Field(
        description=CATEGORY_PROMPT
    )

    confidence_note: Optional[str] = Field(
        default=None,
        description=(
            "Explain briefly if anything important is unclear or ambiguous. "
            "Otherwise return null."
        )
    )


# ======================================================================
# CHECK FILES
# ======================================================================

def check_output_file():

    if not OUTPUT_FILE.exists():

        raise SystemExit(
            f"\nCould not find:\n{OUTPUT_FILE}\n\n"
            "Make sure Expense_Claim_filled.xlsx is in the same folder "
            "as this Python script."
        )


def check_files():

    check_output_file()

    if not RECEIPTS_FOLDER.exists():

        RECEIPTS_FOLDER.mkdir()

        raise SystemExit(
            f"\nCreated receipts folder:\n{RECEIPTS_FOLDER}\n\n"
            "Put receipt images inside it and run the script again."
        )


# ======================================================================
# GET RECEIPTS
# ======================================================================

def get_receipt_files():

    files = sorted(
        file
        for file in RECEIPTS_FOLDER.iterdir()
        if file.suffix.lower() in (
            ".jpg",
            ".jpeg",
            ".png",
            ".pdf",
        )
    )

    if not files:

        raise SystemExit(
            "\nNo receipt files found in the receipts folder."
        )

    max_receipts = (
        LAST_DATA_ROW
        - FIRST_DATA_ROW
        + 1
    )

    if len(files) > max_receipts:

        raise SystemExit(
            f"\nFound {len(files)} receipts but the form "
            f"only supports {max_receipts} receipts."
        )

    return files


# ======================================================================
# LLAMA AGENT
# ======================================================================

def get_agent():

    # Only require the API key when actually processing receipts.
    # Resetting the workbook does NOT require LlamaCloud.

    if not API_KEY:

        raise SystemExit(
            "\nMissing LLAMA_CLOUD_API_KEY.\n\n"
            "Create a .env file beside this script containing:\n\n"
            "LLAMA_CLOUD_API_KEY=llx-your-real-key\n"
        )

    extractor = LlamaExtract(
        api_key=API_KEY
    )

    # Use a dedicated agent name so the correct category schema is used.
    agent_name = "expense-dropdown-receipts-v4"

    try:

        agent = extractor.get_agent(
            name=agent_name
        )

        print(
            f"Using existing extraction agent "
            f"'{agent_name}'."
        )

        return agent

    except Exception:

        print(
            f"Creating extraction agent "
            f"'{agent_name}'..."
        )

        return extractor.create_agent(
            name=agent_name,
            data_schema=Receipt
        )


# ======================================================================
# NORMALISE CATEGORY
# ======================================================================

def normalise_category(category):

    if not category:
        return "/"

    category = category.strip()

    # Exact match
    if category in ALLOWED_CATEGORIES:
        return category

    # Case-insensitive match
    for allowed in ALLOWED_CATEGORIES:

        if (
            category.casefold()
            == allowed.casefold()
        ):

            return allowed

    # Do NOT invent another category.
    return "/"


# ======================================================================
# EXTRACT RECEIPTS
# ======================================================================

def extract_receipts(files):

    agent = get_agent()

    extracted = []

    print(
        f"\nFound {len(files)} receipt(s).\n"
    )

    for number, file in enumerate(
        files,
        start=1
    ):

        print(
            f"[{number}/{len(files)}] "
            f"Extracting {file.name} ..."
        )

        try:

            result = agent.extract(
                str(file)
            )

            receipt = Receipt.model_validate(
                result.data
            )

            data = receipt.model_dump()

            data["category"] = (
                normalise_category(
                    data["category"]
                )
            )

            data["source_file"] = (
                file.name
            )

            extracted.append(
                data
            )

            print(
                f"    Merchant : "
                f"{data['merchant_name']}"
            )

            print(
                f"    Date     : "
                f"{data['receipt_date']}"
            )

            print(
                f"    Category : "
                f"{data['category']}"
            )

            print(
                f"    Amount   : "
                f"{data['currency']} "
                f"{data['total_hkd']}"
            )

            if data.get(
                "confidence_note"
            ):

                print(
                    f"    NOTE     : "
                    f"{data['confidence_note']}"
                )

            print()

        except ValidationError as error:

            print(
                f"\n[VALIDATION ERROR] "
                f"{file.name}\n"
                f"{error}\n"
            )

        except Exception as error:

            print(
                f"\n[EXTRACTION ERROR] "
                f"{file.name}\n"
                f"{error}\n"
            )

    if not extracted:

        raise SystemExit(
            "\nNo receipts were successfully extracted."
        )

    return extracted


# ======================================================================
# DESCRIPTION WRAPPING
# ======================================================================

def wrap_description(
    ws,
    row,
    description
):

    """
    Write receipt description while keeping the Description column
    width COMPLETELY FIXED.

    Behaviour:
    - NEVER changes column D width
    - NEVER changes borders
    - NEVER changes fills/fonts
    - manually inserts line breaks before text reaches the cell edge
    - enables Excel Wrap Text as an additional safeguard
    - increases ROW HEIGHT only
    """

    cell = ws[f"D{row}"]

    # --------------------------------------------------------------
    # READ EXISTING COLUMN WIDTH ONLY.
    # DO NOT CHANGE IT.
    # --------------------------------------------------------------

    current_width = (
        ws.column_dimensions["D"].width
        or 30
    )

    # --------------------------------------------------------------
    # CONSERVATIVE WRAPPING
    #
    # Wrap slightly early so the final word remains safely
    # within the visible Description cell.
    # --------------------------------------------------------------

    characters_per_line = max(
        15,
        int(current_width * 0.80)
    )

    lines = []

    # Preserve existing newlines if present.

    for paragraph in str(description).splitlines():

        paragraph = paragraph.strip()

        if not paragraph:
            lines.append("")
            continue

        wrapped_lines = textwrap.wrap(
            paragraph,
            width=characters_per_line,
            break_long_words=False,
            break_on_hyphens=False,
            replace_whitespace=True,
            drop_whitespace=True,
        )

        if wrapped_lines:
            lines.extend(wrapped_lines)

        else:
            lines.append(paragraph)

    if not lines:
        lines = [""]

    # Insert actual line breaks.

    final_text = "\n".join(lines)

    cell.value = final_text

    # --------------------------------------------------------------
    # ENABLE WRAP TEXT.
    #
    # Does not touch borders, font, fill or column width.
    # --------------------------------------------------------------

    cell.alignment = Alignment(
        horizontal=cell.alignment.horizontal,
        vertical="top",
        text_rotation=cell.alignment.text_rotation,
        wrap_text=True,
        shrink_to_fit=False,
        indent=cell.alignment.indent,
        relativeIndent=cell.alignment.relativeIndent,
        justifyLastLine=cell.alignment.justifyLastLine,
        readingOrder=cell.alignment.readingOrder,
    )

    # --------------------------------------------------------------
    # INCREASE ROW HEIGHT ONLY.
    # --------------------------------------------------------------

    line_count = max(
        1,
        len(lines)
    )

    HEIGHT_PER_LINE = 18
    EXTRA_PADDING = 4

    required_height = (
        line_count * HEIGHT_PER_LINE
        + EXTRA_PADDING
    )

    ws.row_dimensions[row].height = (
        required_height
    )


# ======================================================================
# ENSURE REAL EXCEL DROPDOWN
# ======================================================================

def ensure_category_dropdown(
    wb,
    ws
):

    """
    Reapply a genuine Excel data-validation dropdown to E7:E16.

    Python writes a value that belongs to this dropdown, but the
    dropdown arrow remains available when the employee opens Excel.
    """

    if DROPDOWN_SHEET not in wb.sheetnames:

        raise KeyError(
            f"Worksheet '{DROPDOWN_SHEET}' not found."
        )

    # --------------------------------------------------------------
    # Create workbook-level named range.
    # --------------------------------------------------------------

    range_name = "ExpenseCategories"

    try:

        if range_name in wb.defined_names:

            del wb.defined_names[
                range_name
            ]

    except Exception:

        pass

    category_range = DefinedName(
        range_name,
        attr_text="'Dropdown'!$A$1:$A$36"
    )

    try:

        wb.defined_names.add(
            category_range
        )

    except AttributeError:

        wb.defined_names.append(
            category_range
        )

    # --------------------------------------------------------------
    # Remove previous validation rules applying to E7:E16.
    # --------------------------------------------------------------

    existing_validations = list(
        ws.data_validations.dataValidation
    )

    for validation in existing_validations:

        validation_text = str(
            validation.sqref
        )

        if any(
            f"E{row}" in validation_text
            for row in range(
                FIRST_DATA_ROW,
                LAST_DATA_ROW + 1
            )
        ):

            try:

                ws.data_validations.dataValidation.remove(
                    validation
                )

            except ValueError:

                pass

    # --------------------------------------------------------------
    # Create real dropdown.
    # --------------------------------------------------------------

    dropdown = DataValidation(
        type="list",
        formula1="=ExpenseCategories",
        allow_blank=True
    )

    dropdown.error = (
        "Please select a category "
        "from the dropdown list."
    )

    dropdown.errorTitle = (
        "Invalid Category"
    )

    dropdown.prompt = (
        "Select one of the available "
        "expense categories."
    )

    dropdown.promptTitle = (
        "Expense Category"
    )

    dropdown.showErrorMessage = True
    dropdown.showInputMessage = True

    ws.add_data_validation(
        dropdown
    )

    dropdown.add(
        f"E{FIRST_DATA_ROW}:"
        f"E{LAST_DATA_ROW}"
    )


# ======================================================================
# CLEAR OLD RECEIPTS
# ======================================================================

def clear_existing_entries(
    ws
):

    """
    Clears ONLY receipt data.

    Does NOT alter:
    - borders
    - column widths
    - fonts
    - fills
    - summary formulas
    - total formulas
    """

    for row in range(
        FIRST_DATA_ROW,
        LAST_DATA_ROW + 1
    ):

        ws[f"B{row}"].value = None
        ws[f"C{row}"].value = None
        ws[f"D{row}"].value = None
        ws[f"E{row}"].value = None
        ws[f"H{row}"].value = None

        # Restore normal/default Excel row height after
        # descriptions previously made the row taller.
        ws.row_dimensions[row].height = None


# ======================================================================
# AUTOMATED RESET
# ======================================================================

def reset_expense_claim():

    """
    Return Expense_Claim_filled.xlsx to its clean reusable state.

    Clears:
    - Date
    - Merchant
    - Description
    - Category
    - Amount

    Restores:
    - normal row heights
    - category dropdowns

    Preserves:
    - borders
    - fills
    - fonts
    - column widths
    - merged cells
    - formulas
    - category summary section
    - grand total formulas
    """

    check_output_file()

    wb = load_workbook(
        OUTPUT_FILE
    )

    if SHEET_NAME not in wb.sheetnames:

        raise KeyError(
            f"Worksheet '{SHEET_NAME}' "
            "does not exist."
        )

    ws = wb[SHEET_NAME]

    print(
        "\nResetting expense claim..."
    )

    # --------------------------------------------------------------
    # CLEAR RECEIPT VALUES + RESTORE ROW HEIGHTS
    # --------------------------------------------------------------

    clear_existing_entries(
        ws
    )

    # --------------------------------------------------------------
    # MAKE SURE CATEGORY DROPDOWNS REMAIN AVAILABLE
    # --------------------------------------------------------------

    ensure_category_dropdown(
        wb,
        ws
    )

    # --------------------------------------------------------------
    # RECALCULATE TOTALS WHEN EXCEL OPENS
    # --------------------------------------------------------------

    try:

        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"

    except Exception:

        pass

    # --------------------------------------------------------------
    # SAVE RESET WORKBOOK
    # --------------------------------------------------------------

    try:

        wb.save(
            OUTPUT_FILE
        )

    except PermissionError:

        raise SystemExit(
            "\nERROR: Expense_Claim_filled.xlsx is currently "
            "open in Excel.\n\n"
            "Close the workbook completely and run the reset again.\n"
        )

    print(
        "\n----------------------------------------"
    )

    print(
        "Expense claim reset successfully."
    )

    print(
        "\nReceipt entries were cleared."
    )

    print(
        "Description row heights were restored."
    )

    print(
        "Category dropdowns were preserved."
    )

    print(
        "Borders were untouched."
    )

    print(
        "Column widths were untouched."
    )

    print(
        "Summary formulas were untouched."
    )

    print(
        "Grand total formula was untouched."
    )

    print(
        "\nThe workbook is ready for a new claim."
    )

    print(
        "----------------------------------------\n"
    )


# ======================================================================
# FILL FORM
# ======================================================================

def fill_form(
    extracted
):

    wb = load_workbook(
        OUTPUT_FILE
    )

    if SHEET_NAME not in wb.sheetnames:

        raise KeyError(
            f"Worksheet '{SHEET_NAME}' "
            "does not exist."
        )

    ws = wb[SHEET_NAME]

    # --------------------------------------------------------------
    # RECORD ORIGINAL COLUMN WIDTHS.
    #
    # They are restored before saving as an extra safeguard.
    # --------------------------------------------------------------

    original_column_widths = {}

    for column in (
        "A",
        "B",
        "C",
        "D",
        "E",
        "F",
        "G",
        "H",
        "I",
    ):

        original_column_widths[
            column
        ] = (
            ws.column_dimensions[
                column
            ].width
        )

    # --------------------------------------------------------------
    # CLEAR PREVIOUS RECEIPT ENTRIES.
    # --------------------------------------------------------------

    clear_existing_entries(
        ws
    )

    # --------------------------------------------------------------
    # GUARANTEE E7:E16 HAS REAL DROPDOWNS.
    # --------------------------------------------------------------

    ensure_category_dropdown(
        wb,
        ws
    )

    # --------------------------------------------------------------
    # WRITE RECEIPTS.
    # --------------------------------------------------------------

    for index, receipt in enumerate(
        extracted
    ):

        row = (
            FIRST_DATA_ROW
            + index
        )

        # ==========================================================
        # DATE
        # ==========================================================

        ws[f"B{row}"] = (
            receipt[
                "receipt_date"
            ]
        )

        # ==========================================================
        # MERCHANT
        # ==========================================================

        ws[f"C{row}"] = (
            receipt[
                "merchant_name"
            ]
        )

        # ==========================================================
        # DESCRIPTION
        #
        # Fixed column width.
        # Text wraps.
        # Row becomes taller only when necessary.
        # ==========================================================

        wrap_description(
            ws,
            row,
            receipt[
                "description"
            ]
        )

        # ==========================================================
        # CATEGORY
        # ==========================================================

        category = (
            receipt[
                "category"
            ]
        )

        if (
            category
            not in ALLOWED_CATEGORIES
        ):

            category = "/"

        ws[f"E{row}"] = (
            category
        )

        # ==========================================================
        # AMOUNT
        # ==========================================================

        if (
            receipt[
                "currency"
            ].upper()
            == "HKD"
        ):

            ws[f"H{row}"] = (
                receipt[
                    "total_hkd"
                ]
            )

        else:

            ws[f"H{row}"] = None

            print(
                f"\n[FOREIGN CURRENCY] "
                f"{receipt['source_file']}"
            )

            print(
                f"Receipt amount: "
                f"{receipt['currency']} "
                f"{receipt['total_hkd']}"
            )

            print(
                f"Enter the converted HKD "
                f"amount manually in H{row}."
            )

        print(
            f"\nReceipt {index + 1} written:"
        )

        print(
            f"    Row         : {row}"
        )

        print(
            f"    Merchant    : "
            f"{receipt['merchant_name']}"
        )

        print(
            f"    Category    : "
            f"{category}"
        )

        print(
            f"    Dropdown    : E{row}"
        )

        print(
            f"    Amount      : H{row}"
        )

    # --------------------------------------------------------------
    # RESTORE ALL ORIGINAL COLUMN WIDTHS.
    # --------------------------------------------------------------

    for (
        column,
        width
    ) in original_column_widths.items():

        ws.column_dimensions[
            column
        ].width = width

    # --------------------------------------------------------------
    # FORCE FORMULAS TO RECALCULATE WHEN EXCEL OPENS.
    # --------------------------------------------------------------

    try:

        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"

    except Exception:

        pass

    # --------------------------------------------------------------
    # SAVE
    # --------------------------------------------------------------

    try:

        wb.save(
            OUTPUT_FILE
        )

    except PermissionError:

        raise SystemExit(
            "\nERROR: Expense_Claim_filled.xlsx "
            "is currently open in Excel.\n\n"
            "Close the workbook completely and run:\n\n"
            "python llamacode.py\n"
        )

    print(
        "\n----------------------------------------"
    )

    print(
        f"Done - wrote "
        f"{len(extracted)} receipt(s)."
    )

    print(
        "\nColumn widths were preserved."
    )

    print(
        "Long descriptions were wrapped."
    )

    print(
        "Category dropdowns were restored "
        "to E7:E16."
    )

    print(
        "Excel totals/formulas were untouched."
    )

    print(
        "\nPlease open Expense_Claim_filled.xlsx "
        "and manually verify the extracted information "
        "before printing or submitting the claim."
    )

    print(
        "----------------------------------------\n"
    )


# ======================================================================
# MAIN MENU
# ======================================================================

def main_menu():

    print(
        "\n========================================"
    )

    print(
        "        EXPENSE CLAIM TOOL"
    )

    print(
        "========================================"
    )

    print(
        "\n1 - Process receipt images"
    )

    print(
        "2 - Reset claim form"
    )

    print(
        "3 - Exit"
    )

    choice = input(
        "\nChoose 1, 2 or 3: "
    ).strip()

    # ==================================================================
    # OPTION 1 - PROCESS RECEIPTS
    # ==================================================================

    if choice == "1":

        check_files()

        receipt_files = (
            get_receipt_files()
        )

        receipts = (
            extract_receipts(
                receipt_files
            )
        )

        fill_form(
            receipts
        )

    # ==================================================================
    # OPTION 2 - RESET FORM
    # ==================================================================

    elif choice == "2":

        reset_expense_claim()

        print(
            "IMPORTANT: Before processing the next claim, "
            "remove or replace the old receipt images "
            "inside the receipts folder."
        )

    # ==================================================================
    # OPTION 3 - EXIT
    # ==================================================================

    elif choice == "3":

        print(
            "\nExited Expense Claim Tool.\n"
        )

    # ==================================================================
    # INVALID INPUT
    # ==================================================================

    else:

        print(
            "\nInvalid selection."
        )

        print(
            "Run python llamacode.py again "
            "and choose 1, 2 or 3.\n"
        )


# ======================================================================
# RUN PROGRAM
# ======================================================================

if __name__ == "__main__":

    main_menu()