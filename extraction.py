import pdfplumber
import pandas as pd
from pathlib import Path
import re
import os

# Path to save Excel file in Downloads
downloads_path = Path.home() / "Downloads"
output_excel = downloads_path / "Tables_14_to_24.xlsx"
writer = pd.ExcelWriter(output_excel, engine='openpyxl')

# Regex to detect table titles (e.g., "Table 14: ...")
table_heading_pattern = re.compile(r"Table\s+(\d+):\s+(.*)")

# Input PDF file
pdf_file = "cclf_ip_508_v39.pdf"

# Control flags
collecting = False
current_table_number = None
current_table_title = None
current_rows = []

with pdfplumber.open(pdf_file) as pdf:
    for i, page in enumerate(pdf.pages):
        text = page.extract_text()
        tables = page.extract_tables()

        # Check each line for a new table header
        for line in text.splitlines():
            match = table_heading_pattern.match(line)
            if match:
                table_num = int(match.group(1))
                table_title = match.group(2)

                # If we were collecting a table and we just hit a new one
                if collecting and current_table_number and current_rows:
                    df = pd.DataFrame(current_rows)
                    sheet_name = f"Table_{current_table_number}_{current_table_title[:20]}"
                    sheet_name = sheet_name[:31]  # Excel sheet name limit
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    current_rows = []

                # Start collecting if table_num is in range 14–24
                if 14 <= table_num <= 24:
                    collecting = True
                    current_table_number = table_num
                    current_table_title = table_title
                else:
                    collecting = False

        # If we're in a valid table range, keep collecting rows
        if collecting:
            for table in tables:
                for row in table:
                    # Normalize row text
                    normalized_row = [cell.replace("\n", " ").strip().lower() if cell else "" for cell in row]

                    # Skip repeated headers
                    if (
                            "element #" in normalized_row[0]
                            and "claim field label" in normalized_row[1]
                            and "claim field name" in normalized_row[2]
                    ):
                        continue

                    # Skip non-table text blocks or "paragraphs"
                    non_empty_cells = [cell for cell in normalized_row if cell]
                    if len(non_empty_cells) < 4:
                        continue  # likely not a data row

                    current_rows.append(row)

# Save last table after final page
if collecting and current_table_number and current_rows:
    df = pd.DataFrame(current_rows)
    sheet_name = f"Table_{current_table_number}_{current_table_title[:20]}"
    sheet_name = sheet_name[:31]
    df.to_excel(writer, sheet_name=sheet_name, index=False)

writer.close()
print(f"✅ Extracted Tables 14–24 to {output_excel}")