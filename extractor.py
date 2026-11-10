
from __future__ import annotations

from io import BytesIO
from pathlib import Path
import re

import fitz
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# =========================================================
# CONFIGURATION
# =========================================================

WCC_FIELDS = [
    "Subcontractor Name",
    "WCC Type",
    "WCC Issuance Date",
    "Certificate Number",
    "Site ID",
    "Site Name",
]

SUMMARY_COLUMNS = [
    "File Name",
    *WCC_FIELDS,
    "Latest Approval By",
    "Latest Approval Date",
    "Extraction Status",
]

HISTORY_COLUMNS = [
    "File Name",
    "Activity",
    "Name",
    "Timestamp",
    "Timestamp Raw",
]


# =========================================================
# PDF TEXT EXTRACTION
# =========================================================

def extract_pdf_text(pdf_bytes: bytes) -> str:
    document = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    try:
        return "\n".join(
            page.get_text("text")
            for page in document
        )
    finally:
        document.close()


# =========================================================
# TEXT CLEANING
# =========================================================

def clean_text(value: object) -> str:
    if value is None:
        return ""

    text = str(value)

    text = text.replace("\xa0", " ")
    text = text.replace("\u200b", "")
    text = re.sub(r"[ \t]+", " ", text)

    return text.strip()


def normalize_lines(text: str) -> list[str]:
    return [
        cleaned
        for line in text.splitlines()
        if (cleaned := clean_text(line))
    ]


# =========================================================
# WCC INFORMATION EXTRACTION
# =========================================================

def extract_wcc_fields(pdf_bytes: bytes) -> dict:
    """
    Extract WCC fields using PDF word coordinates.

    The left column contains labels.
    The right column contains values.
    """

    result = {
        field: None
        for field in WCC_FIELDS
    }

    document = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    try:
        for page in document:

            words = page.get_text("words", sort=True)

            page_words = []

            for word in words:
                x0, y0, x1, y1, text = word[:5]

                page_words.append({
                    "x0": x0,
                    "x1": x1,
                    "y": (y0 + y1) / 2,
                    "text": clean_text(text),
                })

            rows = []

            for word in sorted(
                page_words,
                key=lambda item: (item["y"], item["x0"])
            ):

                matched_row = None

                for row in rows:
                    if abs(row["y"] - word["y"]) <= 4:
                        matched_row = row
                        break

                if matched_row is None:
                    rows.append({
                        "y": word["y"],
                        "words": [word]
                    })

                else:
                    matched_row["words"].append(word)

                    matched_row["y"] = sum(
                        item["y"]
                        for item in matched_row["words"]
                    ) / len(matched_row["words"])

            for row in rows:
                row["words"].sort(
                    key=lambda item: item["x0"]
                )

            for field in WCC_FIELDS:

                if result[field]:
                    continue

                target_tokens = [
                    re.sub(
                        r"[^a-z0-9]",
                        "",
                        token.casefold()
                    )
                    for token in field.split()
                ]

                for row in rows:

                    row_words = row["words"]

                    normalized_tokens = [
                        re.sub(
                            r"[^a-z0-9]",
                            "",
                            item["text"].casefold()
                        )
                        for item in row_words
                    ]

                    label_end_index = None

                    for start in range(
                        len(normalized_tokens) - len(target_tokens) + 1
                    ):

                        candidate = normalized_tokens[
                            start:start + len(target_tokens)
                        ]

                        if candidate == target_tokens:
                            label_end_index = (
                                start + len(target_tokens) - 1
                            )
                            break

                    if label_end_index is None:
                        continue

                    label_end_x = row_words[
                        label_end_index
                    ]["x1"]

                    value_words = [
                        item["text"]
                        for item in row_words
                        if item["x0"] > label_end_x + 1
                    ]

                    value = clean_text(
                        " ".join(value_words)
                    )

                    if value:
                        result[field] = value
                        break

            if all(result.values()):
                break

    finally:
        document.close()

    return result


# =========================================================
# APPROVAL LOG EXTRACTION
# =========================================================

def extract_approval_section(text: str) -> str:
    """
    Detect Approval Log heading in extracted PDF text.
    """

    matches = list(
        re.finditer(
            r"Approval\s+Log\s*:?",
            text,
            flags=re.IGNORECASE
        )
    )

    if not matches:
        return ""

    return text[matches[-1].end():]


def normalize_activity(activity: str) -> str:
    """
    Normalize activity capitalization and spacing.
    """

    activity = re.sub(
        r"\s+",
        " ",
        clean_text(activity)
    ).casefold()

    activity_mapping = {
        "submit by": "Submit By",
        "approved by": "Approved By",
    }

    return activity_mapping.get(
        activity,
        activity.title()
    )


def parse_approval_timestamp(timestamp_raw: str):
    """
    Parse timestamps including one-digit seconds.
    """

    try:
        date_part, time_part = timestamp_raw.split()

        hour, minute, second = time_part.split(":")

        normalized_time = (
            f"{hour}:{minute}:{second.zfill(2)}"
        )

        return pd.to_datetime(
            f"{date_part} {normalized_time}",
            format="%Y-%m-%d %H:%M:%S",
            errors="coerce"
        )

    except (ValueError, AttributeError):
        return pd.NaT


def build_approval_history(
    file_name: str,
    pdf_bytes: bytes
) -> list[dict]:
    """
    Extract Approval Log using visual row positions.

    Each activity label is matched with its name and timestamp
    on the same visual row.

    Duplicate labels are preserved when they have valid records.
    Empty activity rows are ignored.
    """

    document = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    history = []
    approval_section_started = False

    try:
        for page in document:

            words = page.get_text("words", sort=True)

            page_words = []

            for word in words:
                x0, y0, x1, y1, text = word[:5]

                page_words.append({
                    "x0": x0,
                    "y": (y0 + y1) / 2,
                    "text": clean_text(text),
                })

            rows = []

            # Group words into visual rows.
            for word in sorted(
                page_words,
                key=lambda item: (item["y"], item["x0"])
            ):

                matched_row = None

                for row in rows:
                    if abs(row["y"] - word["y"]) <= 4:
                        matched_row = row
                        break

                if matched_row is None:
                    rows.append({
                        "y": word["y"],
                        "words": [word]
                    })

                else:
                    matched_row["words"].append(word)

                    matched_row["y"] = sum(
                        item["y"]
                        for item in matched_row["words"]
                    ) / len(matched_row["words"])

            rows.sort(key=lambda row: row["y"])

            for row in rows:

                row["words"].sort(
                    key=lambda item: item["x0"]
                )

                row_text = clean_text(
                    " ".join(
                        item["text"]
                        for item in row["words"]
                    )
                )

                # Find Approval Log heading.
                if not approval_section_started:

                    if re.search(
                        r"\bApproval\s+Log\s*:?",
                        row_text,
                        flags=re.IGNORECASE
                    ):
                        approval_section_started = True

                    continue

                # Match numbered activity labels.
                activity_match = re.match(
                    r"^\s*(\d+)\s*[.)]\s*"
                    r"(Submit\s+by|Approved\s+by)\b"
                    r"\s*(.*?)\s*$",
                    row_text,
                    flags=re.IGNORECASE
                )

                if not activity_match:
                    continue

                activity = normalize_activity(
                    activity_match.group(2)
                )

                remaining_text = clean_text(
                    activity_match.group(3)
                )

                # Require both name and timestamp.
                record_match = re.match(
                    r"^(?P<name>.+?)\s+on\s+"
                    r"(?P<timestamp>\d{4}-\d{2}-\d{2}\s+"
                    r"\d{2}:\d{2}:\d{1,2})$",
                    remaining_text,
                    flags=re.IGNORECASE
                )

                if not record_match:
                    continue

                name = clean_text(
                    record_match.group("name")
                )

                timestamp_raw = clean_text(
                    record_match.group("timestamp")
                )

                if not name:
                    continue

                timestamp = parse_approval_timestamp(
                    timestamp_raw
                )

                if pd.isna(timestamp):
                    continue

                history.append({
                    "File Name": file_name,
                    "Activity": activity,
                    "Name": name,
                    "Timestamp": timestamp,
                    "Timestamp Raw": timestamp_raw,
                })

    finally:
        document.close()

    return history


# =========================================================
# LATEST APPROVAL
# =========================================================

def get_latest_information(history: list[dict]) -> dict:

    result = {
        "Latest Approval By": None,
        "Latest Approval Date": pd.NaT,
    }

    if not history:
        return result

    history_df = pd.DataFrame(history)

    valid_history = history_df.dropna(
        subset=["Timestamp"]
    )

    approval_df = valid_history[
        valid_history["Activity"].str.casefold() == "approved by"
    ]

    if approval_df.empty:
        return result

    latest_approval = approval_df.sort_values(
        "Timestamp"
    ).iloc[-1]

    result["Latest Approval By"] = latest_approval["Name"]
    result["Latest Approval Date"] = latest_approval["Timestamp"]

    return result


# =========================================================
# MAIN EXTRACTION FUNCTION
# =========================================================

def extract_one_pdf(file_name: str, pdf_bytes: bytes):
    """
    Main extraction function called by app.py.
    """

    file_name = Path(file_name).name

    raw_text = extract_pdf_text(pdf_bytes)

    wcc_data = extract_wcc_fields(pdf_bytes)

    approval_section = extract_approval_section(raw_text)

    history = build_approval_history(
        file_name=file_name,
        pdf_bytes=pdf_bytes
    )

    latest_data = get_latest_information(history)

    missing_fields = [
        field
        for field, value in wcc_data.items()
        if not value
    ]

    if missing_fields:

        extraction_status = (
            "Check - Missing: " + ", ".join(missing_fields)
        )

    elif not approval_section:

        extraction_status = "Check - Approval Log Not Found"

    elif not history:

        extraction_status = "Check - Approval Log Could Not Be Parsed"

    else:

        extraction_status = "Success"

    summary = {
        "File Name": file_name,
        **wcc_data,
        **latest_data,
        "Extraction Status": extraction_status,
    }

    return summary, history, raw_text


# =========================================================
# EXCEL EXPORT
# =========================================================

def build_excel(
    summary_data,
    history_data
) -> BytesIO:

    summary_df = pd.DataFrame(
        summary_data
    )

    history_df = pd.DataFrame(
        history_data
    )

    summary_df = summary_df.reindex(
        columns=SUMMARY_COLUMNS
    )

    history_df = history_df.reindex(
        columns=HISTORY_COLUMNS
    )

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        summary_df.to_excel(
            writer,
            sheet_name="WCC_Summary",
            index=False
        )

        history_df.to_excel(
            writer,
            sheet_name="Approval_History",
            index=False
        )

        for sheet_name in [
            "WCC_Summary",
            "Approval_History"
        ]:

            ws = writer.sheets[sheet_name]

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for cell in ws[1]:

                cell.font = Font(
                    bold=True,
                    color="FFFFFF"
                )

                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor="1F4E78"
                )

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

            for column_cells in ws.columns:

                max_length = 0

                column_letter = get_column_letter(
                    column_cells[0].column
                )

                for cell in column_cells:

                    if cell.value is not None:
                        max_length = max(
                            max_length,
                            len(str(cell.value))
                        )

                ws.column_dimensions[
                    column_letter
                ].width = min(max_length + 3, 45)

        # WCC Summary date formatting.
        summary_ws = writer.sheets["WCC_Summary"]

        column_index = (
            SUMMARY_COLUMNS.index("Latest Approval Date") + 1
        )

        for row in range(2, summary_ws.max_row + 1):

            summary_ws.cell(
                row=row,
                column=column_index
            ).number_format = "dd-mmm-yyyy hh:mm:ss"

        # Approval History date formatting.
        history_ws = writer.sheets["Approval_History"]

        timestamp_column = (
            HISTORY_COLUMNS.index("Timestamp") + 1
        )

        for row in range(2, history_ws.max_row + 1):

            history_ws.cell(
                row=row,
                column=timestamp_column
            ).number_format = "dd-mmm-yyyy hh:mm:ss"

    output.seek(0)

    return output
