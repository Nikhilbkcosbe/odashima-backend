#!/usr/bin/env python3
import pandas as pd
import pdfplumber
import sys
sys.path.append('.')


def debug_page_22():
    """Debug the raw PDF data on page 22 to see what's actually in the cells"""

    pdf_path = '07_入札時（見積）積算参考資料.pdf'

    print('Debugging PDF page 22 raw data...')
    print(f'PDF: {pdf_path}')
    print()

    with pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) < 22:
            print(
                f"PDF only has {len(pdf.pages)} pages, cannot access page 22")
            return

        page = pdf.pages[21]  # 0-based indexing, so page 22 is index 21
        print(f"Page 22 content:")
        print("=" * 80)

        # Extract all tables from page 22
        tables = page.extract_tables()
        print(f"Found {len(tables)} tables on page 22")

        for table_idx, table in enumerate(tables):
            print(f"\n--- TABLE {table_idx + 1} ---")
            print(f"Table has {len(table)} rows")

            # Look for reference number in the table - now with one space
            reference_found = False
            reference_row = -1
            for row_idx, row in enumerate(table):
                if row and any(cell and '内 7号' in str(cell) for cell in row):
                    reference_found = True
                    reference_row = row_idx
                    print(f"Found reference '内 7号' at row {row_idx + 1}")
                    break

            if not reference_found:
                print("No '内 7号' reference found in this table")
                continue

            # Print more rows to see the pattern
            print(f"\nTable {table_idx + 1} data (rows 1-25):")
            for row_idx, row in enumerate(table[:25]):
                print(f"Row {row_idx + 1:2d}: {row}")

            # Look specifically for 橋りょう特殊工 and surrounding context
            print(f"\nSearching for '橋りょう特殊工' in table {table_idx + 1}:")
            bridge_special_found = False
            for row_idx, row in enumerate(table):
                if row:
                    for cell_idx, cell in enumerate(row):
                        if cell and '橋りょう特殊工' in str(cell):
                            bridge_special_found = True
                            print(
                                f"  Found '橋りょう特殊工' at Row {row_idx + 1}, Cell {cell_idx + 1}: '{cell}'")
                            print(f"  Full row {row_idx + 1}: {row}")

                            # Check the surrounding rows for context (more rows)
                            print(
                                f"  Extended context (rows around {row_idx + 1}):")
                            for context_idx in range(max(0, row_idx - 3), min(len(table), row_idx + 6)):
                                marker = ">>>" if context_idx == row_idx else "   "
                                print(
                                    f"  {marker} Row {context_idx + 1}: {table[context_idx]}")

            if not bridge_special_found:
                print(f"  No '橋りょう特殊工' found in table {table_idx + 1}")

            # Look for 橋梁点検車
            print(f"\nSearching for '橋梁点検車' in table {table_idx + 1}:")
            bridge_inspection_found = False
            for row_idx, row in enumerate(table):
                if row:
                    for cell_idx, cell in enumerate(row):
                        if cell and '橋梁点検車' in str(cell):
                            bridge_inspection_found = True
                            print(
                                f"  Found '橋梁点検車' at Row {row_idx + 1}, Cell {cell_idx + 1}: '{cell}'")
                            print(f"  Full row {row_idx + 1}: {row}")

                            # Check the surrounding rows for context
                            print(
                                f"  Extended context (rows around {row_idx + 1}):")
                            for context_idx in range(max(0, row_idx - 3), min(len(table), row_idx + 6)):
                                marker = ">>>" if context_idx == row_idx else "   "
                                print(
                                    f"  {marker} Row {context_idx + 1}: {table[context_idx]}")

            if not bridge_inspection_found:
                print(f"  No '橋梁点検車' found in table {table_idx + 1}")


def debug_page_24():
    """Debug the raw PDF data on page 24 to see what's actually in the cells, especially after 内9号."""

    pdf_path = '07_入札時（見積）積算参考資料.pdf'
    output_path = 'page24_debug.txt'
    with open(output_path, 'w', encoding='utf-8') as out:
        out.write('Debugging PDF page 24 raw data...\n')
        out.write(f'PDF: {pdf_path}\n\n')

        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) < 24:
                out.write(f"PDF only has {len(pdf.pages)} pages, cannot access page 24\n")
                return

            page = pdf.pages[23]  # 0-based indexing, so page 24 is index 23
            out.write(f"Page 24 content:\n")
            out.write("=" * 80 + "\n")

            # Extract all tables from page 24
            tables = page.extract_tables()
            out.write(f"Found {len(tables)} tables on page 24\n")

            for table_idx, table in enumerate(tables):
                out.write(f"\n--- TABLE {table_idx + 1} ---\n")
                out.write(f"Table has {len(table)} rows\n")

                # Look for reference number in the table - now with one space
                reference_found = False
                reference_row = -1
                for row_idx, row in enumerate(table):
                    if row and any(cell and ('内9号' in str(cell) or '内 9号' in str(cell)) for cell in row):
                        reference_found = True
                        reference_row = row_idx
                        out.write(f"Found reference '内9号' or '内 9号' at row {row_idx + 1}\n")
                        break

                if not reference_found:
                    out.write("No '内9号' or '内 9号' reference found in this table\n")
                    continue

                # Print more rows to see the pattern
                out.write(f"\nTable {table_idx + 1} data (rows {reference_row+1} to {reference_row+15}):\n")
                for row_idx, row in enumerate(table[reference_row:reference_row+15]):
                    out.write(f"Row {reference_row + row_idx + 1:2d}: {row}\n")
    print(f"Wrote debug output to {output_path}")


if __name__ == "__main__":
    debug_page_22()
    debug_page_24()
