#!/usr/bin/env -S uv run --script
# /// script
# dependencies = [
#     "pandas",
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
import pandas


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('filename', nargs='?', default='DI.xlsx')
    args = parser.parse_args()

    dfs = pandas.read_excel(args.filename, sheet_name=None)
    seen = set()
    for name, df in dfs.items():
        codes = {
            cd
            for p in df['Pavadinimas']
            if isinstance(p, str) and (
                cd := p.partition('Identifikacinis kodas ')[-1]
            )
        }
        print(f"{name}: aktų - {len(codes)}, naujų - {len(codes - seen)}")
        seen.update(codes)

    print(f"Viso: {len(seen)}")
    print(*sorted(seen), sep=', ')
    

if __name__ == '__main__':
    main()
