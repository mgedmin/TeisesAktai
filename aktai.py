#!/usr/bin/env -S uv run --script
# /// script
# dependencies = [
#     "openpyxl",
# ]
# ///
"""
XLS failo visuose worksheetuose raskime unikalius teisės aktų identifikacinius
numerius.

XLS turi turėti stulpelį "Pavadinimas", o jame turi būti tekstas
'Identifikacinis kodas YYYY-NNNNN'.
"""
import argparse
import openpyxl


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('filename', nargs='?', default='DI.xlsx')
    args = parser.parse_args()

    seen = set()
    wb = openpyxl.load_workbook(args.filename)
    for name in wb.sheetnames:
        ws = wb[name]
        codes = {
            cd
            for row in ws.iter_rows(
                min_row=2, min_col=3, max_col=3, values_only=True,
            )
            for value in row
            if isinstance(value, str) and (
                cd := value.partition('Identifikacinis kodas ')[-1]
            )
        }
        print(f"{name}: aktų - {len(codes)}, naujų - {len(codes - seen)}")
        seen.update(codes)

    print(f"Viso: {len(seen)}")
    print(*sorted(seen), sep=', ')


if __name__ == '__main__':
    main()
