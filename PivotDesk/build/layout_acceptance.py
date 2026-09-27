"""Package optional native Excel acceptance macros in a disposable QA copy.

This does not launch or automate Excel. Open the resulting workbook normally
and run PD_LayoutAcceptance. The release workbook contains none of these macros.
"""
from pathlib import Path
import argparse
import re
import zipfile

import build
import vbacompile
import vbalint


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--auto-run', action='store_true',
                        help='Run the synthetic acceptance macro after the QA copy opens.')
    args = parser.parse_args()
    if 'QA' not in args.output.stem.upper() or args.output.suffix.lower() != '.xlsm':
        parser.error('Use a disposable filename containing QA and ending in .xlsm.')
    if args.output.exists():
        parser.error('Output already exists; choose a new QA filename.')
    here = Path(__file__).resolve().parent
    sources = build.read_sources()
    if args.auto_run:
        marker = '    modPD_Desk.Opened\n'
        if sources['ThisWorkbook'].count(marker) != 1:
            raise SystemExit('The workbook startup hook changed; inspect it before adding QA autorun.')
        sources['ThisWorkbook'] = sources['ThisWorkbook'].replace(marker, marker +
            '    Application.OnTime Now + TimeSerial(0, 0, 2), "\'" & ThisWorkbook.Name & "\'!PD_LayoutAcceptance"\n')
    for name, filename in [('modPD_LayoutAcceptance', 'layout_acceptance.bas'),
                           ('modPD_ChartAcceptance', 'chart_acceptance.bas')]:
        source = (here / filename).read_text(encoding='cp1252')
        sources[name] = re.sub(r'^Attribute VB_Name.*\n', '', source)
    problems = vbalint.analyse(sources) + vbalint.arity_problems(sources) + vbacompile.problems(sources)
    if problems:
        raise SystemExit('\n'.join(str(p) for p in problems))
    with zipfile.ZipFile(build.DIST) as original:
        entries = [(item.filename, original.read(item.filename)) for item in original.infolist()]
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(args.output, 'w', zipfile.ZIP_DEFLATED) as out:
        for name, data in entries:
            if name == 'xl/vbaProject.bin':
                data = build.build_vba(data, sources)
            out.writestr(name, data)
    print(f'Created {args.output}. Open in Excel and run PD_LayoutAcceptance.')


if __name__ == '__main__':
    main()
