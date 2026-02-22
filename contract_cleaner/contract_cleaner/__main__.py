import argparse
from pathlib import Path

from contract_cleaner.logging import setup_logging
from contract_cleaner.errors import install_exception_hook
from contract_cleaner.handler import run

def main():
    setup_logging()
    install_exception_hook()

    parser = argparse.ArgumentParser(description="Clean and normalize contract Excel files.")
    parser.add_argument("input_path", type=Path, help="Path to the input Excel file")
    parser.add_argument("output_path", type=Path, help="Path to save the output")
    parser.add_argument("--format", choices=["csv", "workbook"], default="csv", help="Output format (default: csv)")

    args = parser.parse_args()
    run(args.input_path, args.output_path, output_type=args.format)

if __name__ == "__main__":
    main()
