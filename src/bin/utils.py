import sys
import os
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta

from definition import SPONSOR_DATABASE_PATH, SHEET_NAME

class DualLogger:
    """Class to log messages to both terminal and a file.
    This duplicates stdout messages to both the terminal and a log file.

    :param terminal: Original stdout `TextIO` for terminal output.
    :param log: File object `TextIO` where messages are logged.
    """
    def __init__(self, filepath: Path) -> None:
        """Initialize the DualLogger with a path to the log file.
        :param filepath: Path to the file where logs will be written.
        """
        self.terminal = sys.__stdout__  # keep original stdout
        self.log = open(filepath, "w")

    def write(self, message):
        """Write a message to both terminal and log file.
        :param message: The string message to write.
        """
        self.terminal.write(message)  # write to terminal
        self.log.write(message)       # write to file

    def flush(self):
        """Flush both terminal and file buffers."""
        self.terminal.flush()
        self.log.flush()


def get_today_formatted_date() -> str:
    """Get date of today in formatted format "dd.mm.yyyy".
    """
    today = datetime.now()
    today_formatted = today.strftime("%d.%m.%Y")
    return today_formatted


def get_deadline_formatted_date(date: datetime) -> str:
    """Get date plus 30 days in formatted format "dd.mm.yyyy".
    """
    in_30_days = date + timedelta(days=30)
    in_30_days_formatted = in_30_days.strftime("%d.%m.%Y")
    return in_30_days_formatted

def get_invoice_number() -> str:
        """Retrieve latest invoice number from Excel file and increment it by 1
        in order to get current invoice number.

        :return invoice_number: A string containing the computed current invoice number.
        """
        current_year = str(datetime.now().year)
        first_invoice_number = f"{current_year}0000"

        if not os.path.exists(SPONSOR_DATABASE_PATH):
            invoice_number = first_invoice_number
        else:
            existing_data_df = pd.read_excel(SPONSOR_DATABASE_PATH, sheet_name=SHEET_NAME)
            latest_invoice_number = str(existing_data_df["Invoice Number"].max())
            latest_invoice_year = latest_invoice_number[:4]
            if latest_invoice_year != current_year:
                # Set invoice number as first invoice number of the year
                invoice_number = first_invoice_number
            else:
                # Increment latest invoice number by 1
                invoice_number = str(int(latest_invoice_number) + 1)

        return invoice_number

def replace_text(paragraph, old_text, new_text):
        """Replaces old_text with new_text in the given paragraph while keeping the original font and bold style.
        """
        if old_text in paragraph.text:
            bold_text_list = ["[TOTAL]", "[COMPANY]"]
            bold = False
            # Capture the fact that text has to be bold for the "TOTAL" text
            if any(keyword in paragraph.text for keyword in bold_text_list):
                bold = True
            # Replace the text
            paragraph.text = paragraph.text.replace(old_text, new_text)
            # Apply font style
            paragraph.style.font.name = "Times New Roman"
            # Apply bold
            if bold:
                paragraph.style.font.bold = True
            else:
                paragraph.style.font.bold = False
                for run in paragraph.runs:
                    run.bold = False
            # if DEBUG_MODE:
            #     print(f"---\nparagraph.text: {paragraph.text}")
            #     print(f"paragraph.style.font.bold: {paragraph.style.font.bold}")