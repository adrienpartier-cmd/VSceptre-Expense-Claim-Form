"""
DYNAMIC EXPENSE CLAIM AUTOMATION

MASTER:
    Expense_Claim_MASTER.xlsx

WORKING FILE:
    Expense_Claim_filled.xlsx

RECEIPTS:
    ./receipts/


FINAL MASTER STRUCTURE
======================

Row 6
    A      Item
    B      Date
    C      Merchant
    D      Description
    E:F    Additional Context (if needed)
    G      Category
    H:I    Amount (HKD)

Row 7
    ONE blank receipt template row

Row 8
    A:D    Category / Categories
    E:I    Category Total (HKD)

Row 9
    ONE category-total template row

Row 10
    Grand Total

Row 11+
    Submitted By / Approved By / Notes / footer


DYNAMIC BEHAVIOUR
=================

If there are 5 successfully extracted receipts:
    Python creates exactly 5 receipt rows.

If those receipts use 3 unique categories:
    Python creates exactly 3 category-summary rows.

There is NO 10-receipt limit.


IMPORTANT
=========

Additional Context = E:F.

Python NEVER writes anything into Additional Context.

It remains completely blank for the user to complete manually.

Python also:
    - NEVER changes column widths
    - NEVER resizes/repositions the image
    - preserves borders and cell formatting
    - wraps Description in D
    - increases row HEIGHT only for long descriptions
    - dynamically extends Category dropdowns
    - rebuilds summary formulas correctly
    - moves the entire footer down correctly
    - reset restores the exact master workbook
"""


# ======================================================================
# IMPORTS
# ======================================================================

from llama_cloud_services import LlamaExtract

from pydantic import (
    BaseModel,
    Field,
    ValidationError,
)

from typing import Optional

from pathlib import Path

from dotenv import load_dotenv

from openpyxl import load_workbook

from openpyxl.styles import Alignment

from openpyxl.worksheet.datavalidation import DataValidation

from openpyxl.workbook.defined_name import DefinedName

from openpyxl.utils import range_boundaries

from copy import copy, deepcopy

import os
import shutil
import textwrap
import warnings


# ======================================================================
# SUPPRESS EXPECTED OPENPYXL WARNING
# ======================================================================

warnings.filterwarnings(
    "ignore",
    message=(
        "Data Validation extension is not supported "
        "and will be removed"
    ),
)


# ======================================================================
# PATHS
# ======================================================================

BASE_DIR = Path(__file__).resolve().parent


MASTER_FILE = (
    BASE_DIR
    / "Expense_Claim_MASTER.xlsx"
)


OUTPUT_FILE = (
    BASE_DIR
    / "Expense_Claim_filled.xlsx"
)


RECEIPTS_FOLDER = (
    BASE_DIR
    / "receipts"
)


SHEET_NAME = "Form"

DROPDOWN_SHEET = "Dropdown"


# ======================================================================
# FINAL MASTER ROW STRUCTURE
# ======================================================================

HEADER_ROW = 6

RECEIPT_TEMPLATE_ROW = 7

SUMMARY_HEADER_TEMPLATE_ROW = 8

CATEGORY_TEMPLATE_ROW = 9

GRAND_TOTAL_TEMPLATE_ROW = 10

FOOTER_FIRST_ROW = 11

# Meaningful formatted/footer area in the original master.
FOOTER_LAST_ROW = 20


# ======================================================================
# FINAL COLUMN STRUCTURE
# ======================================================================

ITEM_COLUMN = "A"

DATE_COLUMN = "B"

MERCHANT_COLUMN = "C"

DESCRIPTION_COLUMN = "D"

# E:F = MANUAL USER FIELD
ADDITIONAL_CONTEXT_START_COLUMN = "E"
ADDITIONAL_CONTEXT_END_COLUMN = "F"

CATEGORY_COLUMN = "G"

AMOUNT_COLUMN = "H"
AMOUNT_END_COLUMN = "I"


# ======================================================================
# API KEY
# ======================================================================

load_dotenv(
    BASE_DIR / "API_KEY.txt"
)


API_KEY = os.getenv(
    "LLAMA_CLOUD_API_KEY"
)


# ======================================================================
# EXACT CATEGORY OPTIONS
# ======================================================================

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


# ======================================================================
# CATEGORY PROMPT
# ======================================================================

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

General purchases which do not fit another specific category
-> Purchase

Software purchases or licences
-> Software - Purchase

Small electronic or computer replacement parts
-> Misc. Parts - Purchase

Audit charges
-> Audit Fee

Bank service or transaction charges
-> Bank Charges

Business registration payments
-> Business Registration Fee

Consulting or company-secretarial services
-> Consulting and Secretarial Fee

General professional services
-> Legal and Professional Fee

Finance or interest costs
-> Finance Costs

General routine office expenses
-> General Office Expense

Office insurance
-> Insurance - Office

Non-sales staff insurance
-> Insurance -Staff (non-sales)

Contractor insurance
-> Insurance - Contractor

Printing, stationery, subscriptions or membership dues
-> Print,Stationery, Due & Subs

Courier, delivery or postage
-> Postage & Courier

Office rental and management fees
-> Rent & Management Fee - Office

Government rates or office government rent
-> Rates & Gov Rent - Office

Repairs and maintenance
-> Repairs & Maintenance - Office

Telephone, mobile, internet or broadband
-> Telephone & Broadband - Office

Electricity, water or utilities
-> Utilities - Office

Mandatory Provident Fund
-> MPF (Non-Sales)

Legal work specifically
-> Legal Fees

Office stationery, paper, pens and ordinary office supplies
-> Office Supplies

Selling or distribution expenses
-> Selling & Distribution Expense

Advertising, promotion or marketing
-> Advertising &Promotion Expense

Restaurants, cafes, meals, drinks, food or business meals
-> Entertainment

Miscellaneous office expense that does not fit another category
-> Sundry Expense- Office

Flights and overseas transport/travel
-> Travelling - Overseas

Hong Kong taxi, Uber, MTR, bus, ferry, parking, tolls or local travel
-> Travelling - Local

Courses, seminars, conferences, examinations or certifications
-> Training and Seminar Costs

Staff refreshments, welfare or amenities
-> Staff Amenities

Sales staff salaries or allowances
-> Salaries & Allowances (Sales)

General wages and salaries
-> Wages & Salaries

If no sensible category can be determined
-> /

Do NOT create a new category.
"""


# ======================================================================
# LLAMA RECEIPT SCHEMA
# ======================================================================

class Receipt(BaseModel):

    merchant_name: str = Field(
        description=(
            "Business, shop, restaurant or organisation "
            "name printed on the receipt."
        )
    )


    receipt_date: str = Field(
        description=(
            "Transaction date in DD/MM/YYYY format."
        )
    )


    description: str = Field(
        description=(
            "Concise factual description of the items "
            "or services purchased. Include important "
            "product names or model numbers visible on "
            "the receipt where useful. Do not invent "
            "information."
        )
    )


    total_hkd: float = Field(
        description=(
            "Final amount paid as printed on the receipt. "
            "Do NOT perform currency conversion."
        )
    )


    currency: str = Field(
        description=(
            "Three-letter currency code such as "
            "HKD, USD, GBP, EUR, JPY or TWD."
        )
    )


    category: str = Field(
        description=CATEGORY_PROMPT
    )


    confidence_note: Optional[str] = Field(
        default=None,
        description=(
            "Briefly explain if important receipt information "
            "is unclear. Otherwise return null."
        )
    )


# ======================================================================
# FILE CHECKS
# ======================================================================

def check_master_file():

    if not MASTER_FILE.exists():

        raise SystemExit(
            "\nERROR: Expense_Claim_MASTER.xlsx "
            "was not found.\n\n"
            f"Expected:\n{MASTER_FILE}\n"
        )


def check_receipts_folder():

    if not RECEIPTS_FOLDER.exists():

        RECEIPTS_FOLDER.mkdir()

        raise SystemExit(
            "\nCreated receipts folder:\n\n"
            f"{RECEIPTS_FOLDER}\n\n"
            "Put receipt files inside it and run "
            "the program again."
        )


# ======================================================================
# GET RECEIPT FILES
# ======================================================================

def get_receipt_files():

    files = sorted(

        file

        for file in RECEIPTS_FOLDER.iterdir()

        if file.suffix.lower()
        in (
            ".jpg",
            ".jpeg",
            ".png",
            ".pdf",
        )
    )


    if not files:

        raise SystemExit(
            "\nNo receipt files found "
            "inside the receipts folder."
        )


    # NO MAXIMUM RECEIPT LIMIT

    return files


# ======================================================================
# LLAMA AGENT
# ======================================================================

def get_agent():

    if not API_KEY:

        raise SystemExit(
            "\nMissing LLAMA_CLOUD_API_KEY.\n\n"
            "Your API_KEY.txt file should contain:\n\n"
            "LLAMA_CLOUD_API_KEY=llx-your-real-key\n"
        )


    extractor = LlamaExtract(
        api_key=API_KEY
    )


    agent_name = (
        "expense-dynamic-final-v6"
    )


    try:

        agent = extractor.get_agent(
            name=agent_name
        )


        print(
            f"\nUsing extraction agent "
            f"'{agent_name}'."
        )


        return agent


    except Exception:

        print(
            f"\nCreating extraction agent "
            f"'{agent_name}'..."
        )


        return extractor.create_agent(
            name=agent_name,
            data_schema=Receipt
        )


# ======================================================================
# NORMALISE CATEGORY
# ======================================================================

def normalise_category(
    category
):

    if not category:

        return "/"


    category = str(
        category
    ).strip()


    # Exact match

    if category in ALLOWED_CATEGORIES:

        return category


    # Case-insensitive recovery

    for allowed in ALLOWED_CATEGORIES:

        if (
            category.casefold()
            == allowed.casefold()
        ):

            return allowed


    # Never invent another category.

    return "/"


# ======================================================================
# EXTRACT RECEIPTS
# ======================================================================

def extract_receipts(
    files
):

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


            receipt = (
                Receipt.model_validate(
                    result.data
                )
            )


            data = (
                receipt.model_dump()
            )


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
# GET UNIQUE CATEGORIES
# ======================================================================

def get_unique_categories(
    extracted
):

    """
    Number of category rows =
    number of unique categories selected by Llama.

    Categories remain in first-occurrence order.
    """

    categories = []


    for receipt in extracted:

        category = (
            normalise_category(
                receipt.get(
                    "category"
                )
            )
        )


        if category not in categories:

            categories.append(
                category
            )


    return categories


# ======================================================================
# COPY CELL FORMAT
# ======================================================================

def copy_cell_format(
    source,
    target
):

    """
    Copy formatting only.

    Does NOT copy:
        - value
        - formula
    """

    if source.has_style:

        target._style = copy(
            source._style
        )


    target.font = copy(
        source.font
    )


    target.fill = copy(
        source.fill
    )


    target.border = copy(
        source.border
    )


    target.alignment = copy(
        source.alignment
    )


    target.number_format = (
        source.number_format
    )


    target.protection = copy(
        source.protection
    )


# ======================================================================
# SNAPSHOT A ROW
# ======================================================================

def snapshot_row(
    ws,
    row,
    max_column=27
):

    """
    Store complete row formatting + contents.

    Columns A:AA are stored because the master workbook contains
    formatting through AA.
    """

    cells = []


    for column in range(
        1,
        max_column + 1
    ):

        cell = ws.cell(
            row=row,
            column=column
        )


        cells.append(
            {
                "value":
                    cell.value,

                "style":
                    copy(
                        cell._style
                    ),

                "font":
                    copy(
                        cell.font
                    ),

                "fill":
                    copy(
                        cell.fill
                    ),

                "border":
                    copy(
                        cell.border
                    ),

                "alignment":
                    copy(
                        cell.alignment
                    ),

                "number_format":
                    cell.number_format,

                "protection":
                    copy(
                        cell.protection
                    ),
            }
        )


    dimension = ws.row_dimensions[
        row
    ]


    return {

        "cells":
            cells,

        "height":
            dimension.height,

        "hidden":
            dimension.hidden,

        "outline_level":
            dimension.outlineLevel,

        "collapsed":
            dimension.collapsed,
    }


# ======================================================================
# RESTORE A SNAPSHOTTED ROW
# ======================================================================

def restore_row(
    ws,
    row,
    snapshot,
    copy_values=True
):

    for column, saved in enumerate(
        snapshot["cells"],
        start=1
    ):

        target = ws.cell(
            row=row,
            column=column
        )


        target._style = copy(
            saved["style"]
        )


        target.font = copy(
            saved["font"]
        )


        target.fill = copy(
            saved["fill"]
        )


        target.border = copy(
            saved["border"]
        )


        target.alignment = copy(
            saved["alignment"]
        )


        target.number_format = (
            saved[
                "number_format"
            ]
        )


        target.protection = copy(
            saved["protection"]
        )


        if copy_values:

            target.value = (
                saved["value"]
            )


        else:

            target.value = None


    dimension = ws.row_dimensions[
        row
    ]


    dimension.height = (
        snapshot["height"]
    )


    dimension.hidden = (
        snapshot["hidden"]
    )


    dimension.outlineLevel = (
        snapshot[
            "outline_level"
        ]
    )


    dimension.collapsed = (
        snapshot[
            "collapsed"
        ]
    )


# ======================================================================
# SNAPSHOT MERGED RANGES
# ======================================================================

def snapshot_merges(
    ws,
    first_row,
    last_row
):

    merges = []


    for merged in list(
        ws.merged_cells.ranges
    ):

        (
            min_col,
            min_row,
            max_col,
            max_row,
        ) = range_boundaries(
            str(
                merged
            )
        )


        if (
            min_row >= first_row
            and max_row <= last_row
        ):

            merges.append(
                (
                    min_col,
                    min_row,
                    max_col,
                    max_row,
                )
            )


    return merges


# ======================================================================
# COPY / OFFSET MERGES
# ======================================================================

def recreate_shifted_merges(
    ws,
    merges,
    source_first_row,
    destination_first_row
):

    offset = (
        destination_first_row
        - source_first_row
    )


    for (
        min_col,
        min_row,
        max_col,
        max_row,
    ) in merges:

        new_min_row = (
            min_row
            + offset
        )


        new_max_row = (
            max_row
            + offset
        )


        start = ws.cell(
            row=new_min_row,
            column=min_col
        ).coordinate


        end = ws.cell(
            row=new_max_row,
            column=max_col
        ).coordinate


        ws.merge_cells(
            f"{start}:{end}"
        )


# ======================================================================
# DESCRIPTION WRAPPING
# ======================================================================

def wrap_description(
    ws,
    row,
    description
):

    """
    Column D only.

    NEVER changes column width.

    Long descriptions:
        -> wrap to multiple lines
        -> increase row height only
    """

    cell = ws[
        f"{DESCRIPTION_COLUMN}{row}"
    ]


    current_width = (
        ws.column_dimensions[
            DESCRIPTION_COLUMN
        ].width
        or 30
    )


    # Conservative wrapping so text stays inside border.

    characters_per_line = max(
        15,
        int(
            current_width
            * 0.80
        )
    )


    lines = []


    for paragraph in str(
        description
    ).splitlines():


        paragraph = (
            paragraph.strip()
        )


        if not paragraph:

            lines.append("")

            continue


        wrapped_lines = (
            textwrap.wrap(
                paragraph,
                width=characters_per_line,
                break_long_words=False,
                break_on_hyphens=False,
                replace_whitespace=True,
                drop_whitespace=True,
            )
        )


        if wrapped_lines:

            lines.extend(
                wrapped_lines
            )


        else:

            lines.append(
                paragraph
            )


    if not lines:

        lines = [""]


    cell.value = (
        "\n".join(
            lines
        )
    )


    old_alignment = copy(
        cell.alignment
    )


    cell.alignment = Alignment(

        horizontal=
            old_alignment.horizontal,

        vertical="top",

        text_rotation=
            old_alignment.text_rotation,

        wrap_text=True,

        shrink_to_fit=False,

        indent=
            old_alignment.indent,

        relativeIndent=
            old_alignment.relativeIndent,

        justifyLastLine=
            old_alignment.justifyLastLine,

        readingOrder=
            old_alignment.readingOrder,
    )


    # --------------------------------------------------------------
    # ROW HEIGHT ONLY
    # --------------------------------------------------------------

    line_count = max(
        1,
        len(lines)
    )


    HEIGHT_PER_LINE = 18

    EXTRA_PADDING = 4


    required_height = (
        line_count
        * HEIGHT_PER_LINE
        + EXTRA_PADDING
    )


    ws.row_dimensions[
        row
    ].height = required_height


# ======================================================================
# CREATE CATEGORY DROPDOWN
# ======================================================================

def create_category_dropdown(
    wb,
    ws,
    first_receipt_row,
    last_receipt_row
):

    """
    Final Category column = G.

    Dropdown is attached to every dynamically generated receipt row.

    Yellow input message is disabled.
    """

    if DROPDOWN_SHEET not in wb.sheetnames:

        raise KeyError(
            f"Worksheet "
            f"'{DROPDOWN_SHEET}' "
            "was not found."
        )


    range_name = (
        "ExpenseCategories"
    )


    # --------------------------------------------------------------
    # REMOVE OLD NAMED RANGE
    # --------------------------------------------------------------

    try:

        if range_name in wb.defined_names:

            del wb.defined_names[
                range_name
            ]


    except Exception:

        pass


    # --------------------------------------------------------------
    # CREATE CLEAN NAMED RANGE
    # --------------------------------------------------------------

    category_range = (
        DefinedName(

            range_name,

            attr_text=(
                "'Dropdown'!"
                "$A$1:$A$36"
            )
        )
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
    # REMOVE OLD LIST VALIDATIONS
    # --------------------------------------------------------------

    for validation in list(
        ws.data_validations.dataValidation
    ):

        if (
            validation.type
            == "list"
        ):

            try:

                ws.data_validations.dataValidation.remove(
                    validation
                )


            except ValueError:

                pass


    # --------------------------------------------------------------
    # CREATE NEW DROPDOWN
    # --------------------------------------------------------------

    dropdown = DataValidation(

        type="list",

        formula1=(
            "=ExpenseCategories"
        ),

        allow_blank=True,
    )


    dropdown.showErrorMessage = True


    dropdown.errorTitle = (
        "Invalid Category"
    )


    dropdown.error = (
        "Please select a category "
        "from the dropdown list."
    )


    # IMPORTANT:
    # no yellow popup message

    dropdown.showInputMessage = False


    ws.add_data_validation(
        dropdown
    )


    dropdown.add(
        f"{CATEGORY_COLUMN}"
        f"{first_receipt_row}:"
        f"{CATEGORY_COLUMN}"
        f"{last_receipt_row}"
    )


# ======================================================================
# BUILD DYNAMIC FORM
# ======================================================================

def build_dynamic_form(
    wb,
    ws,
    extracted
):

    """
    Completely rebuild rows 7 onward from clean master snapshots.

    This avoids depending on openpyxl to correctly shift:
        - merged cells
        - formulas
        - row heights
        - footer formatting

    Those elements are explicitly recreated instead.
    """

    # ==============================================================
    # COUNTS
    # ==============================================================

    receipt_count = max(
        1,
        len(
            extracted
        )
    )


    unique_categories = (
        get_unique_categories(
            extracted
        )
    )


    category_count = max(
        1,
        len(
            unique_categories
        )
    )


    # ==============================================================
    # SNAPSHOT MASTER ROWS BEFORE CHANGING ANYTHING
    # ==============================================================

    receipt_template = snapshot_row(
        ws,
        RECEIPT_TEMPLATE_ROW
    )


    summary_header_template = snapshot_row(
        ws,
        SUMMARY_HEADER_TEMPLATE_ROW
    )


    category_template = snapshot_row(
        ws,
        CATEGORY_TEMPLATE_ROW
    )


    grand_total_template = snapshot_row(
        ws,
        GRAND_TOTAL_TEMPLATE_ROW
    )


    # --------------------------------------------------------------
    # FOOTER
    # --------------------------------------------------------------

    footer_snapshots = []


    for row in range(
        FOOTER_FIRST_ROW,
        FOOTER_LAST_ROW + 1
    ):

        footer_snapshots.append(
            snapshot_row(
                ws,
                row
            )
        )


    footer_merges = snapshot_merges(
        ws,
        FOOTER_FIRST_ROW,
        FOOTER_LAST_ROW
    )


    # ==============================================================
    # SAVE IMAGE EXACTLY
    # ==============================================================

    image_states = []


    for image in ws._images:

        image_states.append(
            {
                "width":
                    image.width,

                "height":
                    image.height,

                "anchor":
                    deepcopy(
                        image.anchor
                    ),
            }
        )


    # ==============================================================
    # SAVE ORIGINAL COLUMN WIDTHS
    # ==============================================================

    original_widths = {}


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

        original_widths[
            column
        ] = (
            ws.column_dimensions[
                column
            ].width
        )


    # ==============================================================
    # CALCULATE FINAL ROW POSITIONS
    # ==============================================================

    first_receipt_row = (
        RECEIPT_TEMPLATE_ROW
    )


    last_receipt_row = (
        first_receipt_row
        + receipt_count
        - 1
    )


    summary_header_row = (
        last_receipt_row
        + 1
    )


    first_category_row = (
        summary_header_row
        + 1
    )


    last_category_row = (
        first_category_row
        + category_count
        - 1
    )


    grand_total_row = (
        last_category_row
        + 1
    )


    footer_first_destination = (
        grand_total_row
        + 1
    )


    footer_shift = (
        footer_first_destination
        - FOOTER_FIRST_ROW
    )


    # ==============================================================
    # UNMERGE ALL MASTER MERGES FROM ROW 7 DOWNWARD
    #
    # Keep everything in rows 1-6 untouched.
    # ==============================================================

    for merged in list(
        ws.merged_cells.ranges
    ):

        (
            _,
            min_row,
            _,
            max_row,
        ) = range_boundaries(
            str(
                merged
            )
        )


        if min_row >= RECEIPT_TEMPLATE_ROW:

            ws.unmerge_cells(
                str(
                    merged
                )
            )


    # ==============================================================
    # INSERT THE REQUIRED EXTRA SPACE
    #
    # Master already has:
    #
    # 1 receipt row
    # 1 category row
    #
    # therefore:
    #
    # extra rows =
    #     (receipt_count - 1)
    #     +
    #     (category_count - 1)
    # ==============================================================

    extra_rows = (
        (receipt_count - 1)
        +
        (category_count - 1)
    )


    if extra_rows > 0:

        # Insert immediately before original footer.
        #
        # We will manually rebuild everything anyway.

        ws.insert_rows(
            FOOTER_FIRST_ROW,
            amount=extra_rows
        )


    # ==============================================================
    # CLEAR THE ENTIRE DYNAMIC AREA
    #
    # We rebuild every row explicitly.
    # ==============================================================

    final_footer_last_row = (
        footer_first_destination
        + len(
            footer_snapshots
        )
        - 1
    )


    for row in range(
        RECEIPT_TEMPLATE_ROW,
        final_footer_last_row + 1
    ):

        for column in range(
            1,
            28
        ):

            ws.cell(
                row=row,
                column=column
            ).value = None


    # ==============================================================
    # RECEIPT ROWS
    # ==============================================================

    for row in range(
        first_receipt_row,
        last_receipt_row + 1
    ):

        restore_row(
            ws,
            row,
            receipt_template,
            copy_values=False
        )


        # ----------------------------------------------------------
        # FINAL RECEIPT MERGES
        # ----------------------------------------------------------

        # Additional Context
        ws.merge_cells(
            f"E{row}:F{row}"
        )


        # Amount
        ws.merge_cells(
            f"H{row}:I{row}"
        )


        # ----------------------------------------------------------
        # ITEM NUMBER
        # ----------------------------------------------------------

        ws[
            f"A{row}"
        ] = (
            row
            - first_receipt_row
            + 1
        )


        # ----------------------------------------------------------
        # ADDITIONAL CONTEXT
        #
        # MANUAL USER FIELD.
        # ALWAYS EMPTY.
        # ----------------------------------------------------------

        ws[
            f"E{row}"
        ] = None


    # ==============================================================
    # SUMMARY HEADER
    # ==============================================================

    restore_row(
        ws,
        summary_header_row,
        summary_header_template,
        copy_values=False
    )


    ws.merge_cells(
        f"A{summary_header_row}:"
        f"D{summary_header_row}"
    )


    ws.merge_cells(
        f"E{summary_header_row}:"
        f"I{summary_header_row}"
    )


    ws[
        f"A{summary_header_row}"
    ] = (
        "Category / Categories"
    )


    ws[
        f"E{summary_header_row}"
    ] = (
        "Category Total (HKD)"
    )


    # ==============================================================
    # CATEGORY SUMMARY ROWS
    # ==============================================================

    for row in range(
        first_category_row,
        last_category_row + 1
    ):

        restore_row(
            ws,
            row,
            category_template,
            copy_values=False
        )


        ws.merge_cells(
            f"A{row}:D{row}"
        )


        ws.merge_cells(
            f"E{row}:I{row}"
        )


        # ----------------------------------------------------------
        # CATEGORY FORMULA
        #
        # Category appears only when:
        #
        # G is not blank
        # AND
        # H is not zero
        #
        # Exactly as previously requested.
        # ----------------------------------------------------------

        position = (
            row
            - first_category_row
            + 1
        )


        ws[
            f"A{row}"
        ] = (

            '=IFERROR('

            'INDEX('

            '_xlfn.UNIQUE('

            '_xlfn._xlws.FILTER('

            f'$G${first_receipt_row}:'
            f'$G${last_receipt_row},'

            f'('
            f'$G${first_receipt_row}:'
            f'$G${last_receipt_row}'
            f'<>"")'

            '*'

            f'('
            f'$H${first_receipt_row}:'
            f'$H${last_receipt_row}'
            f'<>0)'

            ')'

            '),'

            f'{position}'

            '),'

            '""'

            ')'
        )


        # ----------------------------------------------------------
        # CATEGORY TOTAL
        # ----------------------------------------------------------

        ws[
            f"E{row}"
        ] = (

            f'=IF('

            f'A{row}="",'

            f'"",'

            f'IF('

            f'SUMIF('

            f'$G${first_receipt_row}:'
            f'$G${last_receipt_row},'

            f'A{row},'

            f'$H${first_receipt_row}:'
            f'$H${last_receipt_row}'

            f')=0,'

            f'"",'

            f'SUMIF('

            f'$G${first_receipt_row}:'
            f'$G${last_receipt_row},'

            f'A{row},'

            f'$H${first_receipt_row}:'
            f'$H${last_receipt_row}'

            f')'

            f')'

            f')'
        )


    # ==============================================================
    # GRAND TOTAL
    # ==============================================================

    restore_row(
        ws,
        grand_total_row,
        grand_total_template,
        copy_values=False
    )


    ws.merge_cells(
        f"A{grand_total_row}:"
        f"D{grand_total_row}"
    )


    ws.merge_cells(
        f"E{grand_total_row}:"
        f"I{grand_total_row}"
    )


    ws[
        f"A{grand_total_row}"
    ] = (
        "Grand Total (HKD)"
    )


    ws[
        f"E{grand_total_row}"
    ] = (

        f'=SUM('

        f'$H${first_receipt_row}:'

        f'$H${last_receipt_row}'

        f')'
    )


    # ==============================================================
    # RESTORE FOOTER ROWS EXACTLY
    # ==============================================================

    for index, footer_snapshot in enumerate(
        footer_snapshots
    ):

        destination_row = (
            footer_first_destination
            + index
        )


        restore_row(
            ws,
            destination_row,
            footer_snapshot,
            copy_values=True
        )


    # ==============================================================
    # RESTORE FOOTER MERGES EXACTLY AT NEW POSITION
    # ==============================================================

    recreate_shifted_merges(
        ws,
        footer_merges,
        FOOTER_FIRST_ROW,
        footer_first_destination
    )


    # ==============================================================
    # CATEGORY DROPDOWN
    # ==============================================================

    create_category_dropdown(
        wb,
        ws,
        first_receipt_row,
        last_receipt_row
    )


    # ==============================================================
    # RESTORE ORIGINAL COLUMN WIDTHS
    # ==============================================================

    for (
        column,
        width
    ) in original_widths.items():

        ws.column_dimensions[
            column
        ].width = width


    # ==============================================================
    # RESTORE IMAGE EXACTLY
    # ==============================================================

    for (
        image,
        saved
    ) in zip(
        ws._images,
        image_states
    ):

        image.width = (
            saved["width"]
        )


        image.height = (
            saved["height"]
        )


        image.anchor = deepcopy(
            saved["anchor"]
        )


    return {

        "receipt_count":
            receipt_count,

        "unique_category_count":
            len(
                unique_categories
            ),

        "first_receipt_row":
            first_receipt_row,

        "last_receipt_row":
            last_receipt_row,

        "summary_header_row":
            summary_header_row,

        "first_category_row":
            first_category_row,

        "last_category_row":
            last_category_row,

        "grand_total_row":
            grand_total_row,

        "footer_first_row":
            footer_first_destination,
    }


# ======================================================================
# WRITE RECEIPTS
# ======================================================================

def write_receipts(
    ws,
    extracted,
    layout
):

    first_receipt_row = (
        layout[
            "first_receipt_row"
        ]
    )


    for index, receipt in enumerate(
        extracted
    ):

        row = (
            first_receipt_row
            + index
        )


        # ==========================================================
        # A — ITEM
        # ==========================================================

        ws[
            f"A{row}"
        ] = (
            index + 1
        )


        # ==========================================================
        # B — DATE
        # ==========================================================

        ws[
            f"B{row}"
        ] = (
            receipt[
                "receipt_date"
            ]
        )


        # ==========================================================
        # C — MERCHANT
        # ==========================================================

        ws[
            f"C{row}"
        ] = (
            receipt[
                "merchant_name"
            ]
        )


        # ==========================================================
        # D — DESCRIPTION
        # ==========================================================

        wrap_description(
            ws,
            row,
            receipt[
                "description"
            ]
        )


        # ==========================================================
        # E:F — ADDITIONAL CONTEXT
        #
        # ABSOLUTELY UNTOUCHED BY EXTRACTION.
        #
        # Blank manual-entry field.
        # ==========================================================

        ws[
            f"E{row}"
        ] = None


        # ==========================================================
        # G — CATEGORY
        # ==========================================================

        category = (
            normalise_category(
                receipt[
                    "category"
                ]
            )
        )


        ws[
            f"G{row}"
        ] = category


        # ==========================================================
        # H:I — AMOUNT
        # ==========================================================

        currency = str(
            receipt[
                "currency"
            ]
        ).upper()


        if currency == "HKD":

            ws[
                f"H{row}"
            ] = (
                receipt[
                    "total_hkd"
                ]
            )


        else:

            # ------------------------------------------------------
            # FOREIGN CURRENCY
            #
            # Leave HKD amount blank.
            #
            # User checks bank/card statement and manually enters
            # actual converted HKD amount.
            # ------------------------------------------------------

            ws[
                f"H{row}"
            ] = None


            print(
                "\n[FOREIGN CURRENCY]"
            )


            print(
                f"Receipt: "
                f"{receipt['source_file']}"
            )


            print(
                f"Original amount: "
                f"{currency} "
                f"{receipt['total_hkd']}"
            )


            print(
                f"Please manually enter "
                f"the converted HKD amount "
                f"in H{row}."
            )


        print(
            f"\nReceipt {index + 1}:"
        )


        print(
            f"    Row      : {row}"
        )


        print(
            f"    Merchant : "
            f"{receipt['merchant_name']}"
        )


        print(
            f"    Category : "
            f"{category}"
        )


# ======================================================================
# VERIFY GENERATED WORKBOOK
# ======================================================================

def verify_generated_workbook():

    """
    Immediately reopen the generated workbook after saving.

    This catches structural problems before the user opens it
    in Excel.
    """

    try:

        verification_wb = load_workbook(
            OUTPUT_FILE,
            data_only=False
        )


        if SHEET_NAME not in verification_wb.sheetnames:

            raise RuntimeError(
                "Form worksheet missing."
            )


        verification_ws = verification_wb[
            SHEET_NAME
        ]


        # Force workbook structures to be accessed.

        _ = list(
            verification_ws.merged_cells.ranges
        )


        _ = list(
            verification_ws.data_validations.dataValidation
        )


        _ = verification_ws.max_row


        verification_wb.close()


    except Exception as error:

        raise SystemExit(
            "\nERROR: Generated workbook failed "
            "the integrity check.\n\n"
            f"{error}\n\n"
            "The master workbook was not modified."
        )


# ======================================================================
# PROCESS RECEIPTS
# ======================================================================

def process_receipts():

    check_master_file()

    check_receipts_folder()


    receipt_files = (
        get_receipt_files()
    )


    # ==============================================================
    # OCR FIRST
    # ==============================================================

    extracted = (
        extract_receipts(
            receipt_files
        )
    )


    # ==============================================================
    # ALWAYS START FROM EXACT MASTER
    # ==============================================================

    try:

        shutil.copy2(
            MASTER_FILE,
            OUTPUT_FILE
        )


    except PermissionError:

        raise SystemExit(
            "\nERROR: Expense_Claim_filled.xlsx "
            "is currently open in Excel.\n\n"
            "Close it completely and run again."
        )


    # ==============================================================
    # LOAD FRESH WORKING COPY
    # ==============================================================

    wb = load_workbook(
        OUTPUT_FILE
    )


    if SHEET_NAME not in wb.sheetnames:

        raise SystemExit(
            f"\nERROR: Worksheet "
            f"'{SHEET_NAME}' was not found."
        )


    ws = wb[
        SHEET_NAME
    ]


    # ==============================================================
    # DYNAMICALLY BUILD FORM
    # ==============================================================

    layout = (
        build_dynamic_form(
            wb,
            ws,
            extracted
        )
    )


    # ==============================================================
    # WRITE RECEIPTS
    # ==============================================================

    write_receipts(
        ws,
        extracted,
        layout
    )


    # ==============================================================
    # RECALCULATE FORMULAS WHEN EXCEL OPENS
    # ==============================================================

    try:

        wb.calculation.fullCalcOnLoad = True

        wb.calculation.forceFullCalc = True

        wb.calculation.calcMode = "auto"


    except Exception:

        pass


    # ==============================================================
    # SAVE
    # ==============================================================

    try:

        wb.save(
            OUTPUT_FILE
        )


        wb.close()


    except PermissionError:

        raise SystemExit(
            "\nERROR: Expense_Claim_filled.xlsx "
            "is currently open in Excel.\n\n"
            "Close it completely and run again."
        )


    # ==============================================================
    # VERIFY
    # ==============================================================

    verify_generated_workbook()


    # ==============================================================
    # FINISHED
    # ==============================================================

    print(
        "\n========================================"
    )


    print(
        "EXPENSE CLAIM CREATED"
    )


    print(
        "========================================"
    )


    print(
        f"\nReceipts processed: "
        f"{layout['receipt_count']}"
    )


    print(
        f"Receipt rows: "
        f"{layout['first_receipt_row']}"
        f"-"
        f"{layout['last_receipt_row']}"
    )


    print(
        f"Unique categories: "
        f"{layout['unique_category_count']}"
    )


    print(
        f"Category rows: "
        f"{layout['first_category_row']}"
        f"-"
        f"{layout['last_category_row']}"
    )


    print(
        f"Grand Total row: "
        f"{layout['grand_total_row']}"
    )


    print(
        f"Footer begins at row: "
        f"{layout['footer_first_row']}"
    )


    print(
        "\nAdditional Context (E:F) "
        "was left completely blank."
    )


    print(
        "Column widths were preserved."
    )


    print(
        "Image size/location was preserved."
    )


    print(
        "Footer formatting was preserved."
    )


    print(
        "Summary formulas were rebuilt."
    )


    print(
        "Category dropdowns were extended."
    )


    print(
        "\nOpen Expense_Claim_filled.xlsx "
        "and manually check ALL information "
        "before printing/submitting."
    )


    print(
        "========================================\n"
    )


# ======================================================================
# RESET
# ======================================================================

def reset_expense_claim():

    """
    RESET IS DELIBERATELY SIMPLE.

    It does not try to delete individual dynamic rows.

    Instead:

        Expense_Claim_MASTER.xlsx

    is copied directly over:

        Expense_Claim_filled.xlsx

    Therefore reset ALWAYS restores the exact master:
        - exact columns
        - exact rows
        - exact borders
        - exact merges
        - exact formulas
        - exact image
        - exact row heights
        - exactly one receipt row
        - exactly one category row
    """

    check_master_file()


    try:

        shutil.copy2(
            MASTER_FILE,
            OUTPUT_FILE
        )


    except PermissionError:

        raise SystemExit(
            "\nERROR: Expense_Claim_filled.xlsx "
            "is currently open in Excel.\n\n"
            "Close Excel completely and run Reset again."
        )


    print(
        "\n========================================"
    )


    print(
        "EXPENSE CLAIM RESET"
    )


    print(
        "========================================"
    )


    print(
        "\nThe exact master workbook "
        "has been restored."
    )


    print(
        "\nThe working file is now back to:"
    )


    print(
        "• 1 blank receipt row"
    )


    print(
        "• 1 blank category row"
    )


    print(
        "• Original Grand Total"
    )


    print(
        "• Original signature section"
    )


    print(
        "• Original Notes"
    )


    print(
        "• Original borders and merges"
    )


    print(
        "• Original row heights"
    )


    print(
        "• Original column widths"
    )


    print(
        "• Original image"
    )


    print(
        "\nReady for the next claim."
    )


    print(
        "========================================\n"
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
        "\n1 - Process receipts"
    )


    print(
        "2 - Reset expense claim"
    )


    print(
        "3 - Exit"
    )


    choice = input(
        "\nChoose 1, 2 or 3: "
    ).strip()


    # ==============================================================
    # PROCESS
    # ==============================================================

    if choice == "1":

        process_receipts()


    # ==============================================================
    # RESET
    # ==============================================================

    elif choice == "2":

        reset_expense_claim()


        print(
            "IMPORTANT: Remove or replace "
            "the previous receipt files "
            "inside the receipts folder "
            "before processing a new claim.\n"
        )


    # ==============================================================
    # EXIT
    # ==============================================================

    elif choice == "3":

        input(
            "\nPress Enter to exit...\n"
        )


    # ==============================================================
    # INVALID
    # ==============================================================

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