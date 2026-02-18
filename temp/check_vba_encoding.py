#!/usr/bin/env python3
import sys
from pathlib import Path

EXTS = {'.bas', '.cls', '.frm', '.dcm'}

def bom_type(b: bytes) -> str:
    if b.startswith(b"\xef\xbb\xbf"):
        return 'UTF-8 BOM'
    if b.startswith(b"\xff\xfe"):
        return 'UTF-16 LE BOM'
    if b.startswith(b"\xfe\xff"):
        return 'UTF-16 BE BOM'
    return 'None'

def newline_stats(b: bytes):
    crlf = 0
    lf_alone = 0
    i = 0
    n = len(b)
    while i < n:
        if b[i] == 0x0D and i+1 < n and b[i+1] == 0x0A:
            crlf += 1
            i += 2
            continue
        if b[i] == 0x0A:
            lf_alone += 1
        i += 1
    return crlf, lf_alone

def roundtrip_ok(b: bytes, enc: str) -> bool:
    try:
        s = b.decode(enc)
        return s.encode(enc) == b
    except Exception:
        return False

def guess_encoding(b: bytes) -> str:
    bom = bom_type(b)
    if bom != 'None':
        return bom
    if roundtrip_ok(b, 'utf-8'):
        return 'UTF-8'
    # Python codec name for CP932
    if roundtrip_ok(b, 'cp932'):
        return 'CP932'
    return 'Unknown'

def main(path_str: str):
    root = Path(path_str)
    files = sorted([p for p in root.rglob('*') if p.is_file() and p.suffix.lower() in EXTS])
    rows = []
    for p in files:
        b = p.read_bytes()
        crlf, lf = newline_stats(b)
        enc = guess_encoding(b)
        crlf_ok = (lf == 0)
        needs = not (enc == 'CP932' and crlf_ok)
        rows.append((str(p), enc, crlf_ok, needs, crlf, lf))

    # Sort: needs conversion first, then by path
    rows.sort(key=lambda r: (not r[3], r[0]))

    # Print header
    print(f"{'Needs':5}  {'CRLF':4}  {'Encoding':12}  {'CRLF#':5}  {'LF#':3}  Path")
    for path, enc, crlf_ok, needs, crlf, lf in rows:
        print(f"{str(needs):5}  {str(crlf_ok):4}  {enc:12}  {crlf:5}  {lf:3}  {path}")

if __name__ == '__main__':
    p = sys.argv[1] if len(sys.argv) > 1 else '.'
    main(p)

