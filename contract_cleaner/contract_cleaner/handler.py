import logging
from pathlib import Path

from openpyxl import Workbook

from contract_cleaner.clean import clean_items
from contract_cleaner.io import load_workbook_from_path, write_items_to_csv, write_items_to_workbook
from contract_cleaner.parse import parse_sheet, parse_workbook
from contract_cleaner.models import ParsedSheet, SheetFormat


log = logging.getLogger(__name__)


def run(input_path: Path, output_path: Path, output_type: str = "csv") -> None:
    """Main entry point for cleaning a workbook"""
    input_wb = load_workbook_from_path(input_path)
    try:
        sheet_mapping = parse_workbook(input_wb)
    except ValueError as e:
        log.error(f"Error parsing workbook: {e}")
        return
    except NotImplementedError as e:
        log.error(f"Error parsing workbook: {e}")
        return
    
    cleaned_sheets = clean_sheets(input_wb, sheet_mapping)
    if output_type == "csv":
        write_items_to_csv(cleaned_sheets, output_path)
    elif output_type == "workbook":
        write_items_to_workbook(cleaned_sheets, output_path)


def clean_sheets(wb: Workbook, sheet_mapping: list[SheetFormat]) -> list[ParsedSheet]:
    output = []
    for mapping in sheet_mapping:
        sheet = wb[mapping.name]
        line_items = parse_sheet(sheet, mapping.header_mapping)
        cleaned_items = clean_items(line_items)
        output.append(ParsedSheet(sheet.title, cleaned_items))

    return output


if __name__ == "__main__":
    run(Path("TestData.xlsx"), Path("output.xlsx"), output_type="csv")
