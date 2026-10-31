"""Check the refreshed synthetic fixture's TOC and page sequence in Word PDF."""
from __future__ import annotations

import argparse
import re
import subprocess
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pdf', type=Path)
    parser.add_argument('--pdftotext', default='pdftotext')
    args = parser.parse_args()
    result = subprocess.run([args.pdftotext, '-layout', '-enc', 'UTF-8', str(args.pdf), '-'],
                            capture_output=True, text=True, encoding='utf-8', errors='replace', check=True)
    pages = result.stdout.split('\f')
    toc_pages = [page for page in pages if 'Signal mapping' in page and 'Synthetic evaluation' in page and
                 'Evaluation method' in page and re.search(r'目\s+录', page)]
    if len(toc_pages) != 1:
        raise SystemExit('synthetic TOC page not found exactly once')
    toc = toc_pages[0]
    checks = [r'摘\s*要\s*\.{3,}\s*I\s*$', r'Abstract\s*\.{3,}\s*II\s*$',
              r'第\s*1\s*章\s*Signal mapping\s*\.{3,}\s*1\s*$',
              r'第\s*2\s*章\s*Synthetic evaluation\s*\.{3,}\s*3\s*$',
              r'2\.1\s*Evaluation method\s*\.{3,}\s*3\s*$']
    for pattern in checks:
        if not re.search(pattern, toc, re.M):
            raise SystemExit(f'TOC entry/page mismatch: {pattern}')
    for pattern in (r'^\s*1\.\s+First synthetic list item\.', r'^\s*2\.\s+Second synthetic list item\.',
                    r'^\s*1\.\s+New sequence first item\.', r'图\s*1-1\s+Nested synthetic pipeline',
                    r'图\s*1-2\s+First synthetic pipeline'):
        if not re.search(pattern, result.stdout, re.M):
            raise SystemExit(f'Word numbering mismatch: {pattern}')
    print('Word-rendered TOC and Roman/Arabic page numbers: PASS')
    print('Independent list and chapter/figure numbering: PASS')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
