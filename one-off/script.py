from openpyxl import load_workbook, Workbook
from datetime import datetime, date

INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "output.xlsx"

DEFAULT_START = date(2024, 10, 23)
DEFAULT_END = date(2026, 12, 31)

YEARS = [2024, 2025, 2026]


def parse_dates(value):
    if value is None or str(value).strip() == "":
        return DEFAULT_START, DEFAULT_END

    value = str(value).strip()

    if " to " in value:
        start, end = value.split(" to ")
        return (
            datetime.strptime(start, "%m/%d/%Y").date(),
            datetime.strptime(end, "%m/%d/%Y").date(),
        )

    if " and later" in value:
        start = value.replace(" and later", "")
        return datetime.strptime(start, "%m/%d/%Y").date(), DEFAULT_END

    if value.startswith("On or before "):
        end = value.replace("On or before ", "")
        return DEFAULT_START, datetime.strptime(end, "%m/%d/%Y").date()

    raise ValueError(f"Unknown format: {value}")


def overlaps_year(start, end, year):
    year_start = date(year, 1, 1)
    year_end = date(year, 12, 31)
    return not (end < year_start or start > year_end)


wb = load_workbook(INPUT_FILE)
ws = wb.active

headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
date_idx = headers.index("DATES")

output_wb = Workbook()
output_sheets = {}


for sheetname in wb.sheetnames:
    del wb[sheetname]

for year in YEARS:
    sheet = output_wb.create_sheet(str(year))
    output_sheets[year] = sheet

    new_headers = headers[:date_idx] + ["START DATE", "END DATE"] + headers[date_idx+1:]
    sheet.append(new_headers)

for row in ws.iter_rows(min_row=2, values_only=True):

    dates_value = row[date_idx]
    start, end = parse_dates(dates_value)

    new_row = (
        list(row[:date_idx])
        + [start, end]
        + list(row[date_idx+1:])
    )

    for year in YEARS:
        if overlaps_year(start, end, year):
            output_sheets[year].append(new_row)

output_wb.save(OUTPUT_FILE)