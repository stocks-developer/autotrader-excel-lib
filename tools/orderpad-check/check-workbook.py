"""Check the Order Pad workbook still matches the modules in this repository.

The workbook is self-contained: it carries its own copy of every module, so a
user has nothing to import. That copy is a SNAPSHOT. Change a module in
`direct/` and the workbook keeps shipping the old one until it is rebuilt, and
nothing about the file looks any different when that happens -- it opens, the
buttons work, and the code inside is simply older than the repository.

So run this after any change to `direct/` or to the Order Pad module:

    python tools/orderpad-check/check-workbook.py

It fails if the workbook has fallen behind, if a module is missing, or if the
embedded config carries anything other than the <API_KEY> placeholder.

Needs `oletools` (pip install oletools).

Two things worth knowing before reading a failure:

  Comparison is CASE-INSENSITIVE. The VBA editor normalises identifier case on
  import -- `value` becomes `Value`, `rows.count` becomes `Rows.Count` -- so a
  byte comparison reports hundreds of differences on a module that is in fact
  identical.

  VBA source must be plain ASCII. Modules are stored in a code page rather than
  UTF-8, so any character above ASCII does not survive an import and shows up
  here as a module that does not match its own source.
"""
import re
import sys
from pathlib import Path

try:
    from oletools.olevba import VBA_Parser
except ImportError:
    print("oletools is not installed. Run: pip install oletools")
    sys.exit(2)

REPO = Path(__file__).resolve().parents[2]
BOOK = REPO / "clients/excel/current/samples/AutoTraderWeb-OrderPad.xlsm"

# module name inside the workbook -> the file in this repository it must match
SOURCES = {
    "AutoTraderConfig": REPO / "direct/AutoTraderConfig.bas",
    "AutoTraderWebDirect": REPO / "direct/AutoTraderWebDirect.bas",
    "AutoTraderClient": REPO / "direct/AutoTraderClientDirect.bas",
    "AutoTraderOrderPad": REPO / "clients/excel/current/samples/modules/AutoTraderOrderPad.bas",
}


def normalise(text):
    """Drop what the VBA editor is entitled to change, keep everything else."""
    lines = []
    for line in text.replace("\r\n", "\n").split("\n"):
        if line.startswith("Attribute VB_"):
            continue
        lines.append(line.rstrip().lower())
    while lines and not lines[-1]:
        lines.pop()
    return lines


def main():
    if not BOOK.exists():
        print("workbook not found: %s" % BOOK)
        return 2

    problems = []

    for name, src in SOURCES.items():
        if not src.exists():
            problems.append("source missing: %s" % src)
            continue
        raw = src.read_text(encoding="utf-8", errors="replace")
        odd = sorted({c for c in raw if ord(c) > 127})
        if odd:
            problems.append("%s contains non-ASCII (%s). VBA source must be "
                            "plain ASCII: it will not survive import."
                            % (src.name, ", ".join("U+%04X" % ord(c) for c in odd)))

    if problems:
        for p in problems:
            print("  " + p)
        print("FAIL")
        return 1

    parser = VBA_Parser(str(BOOK))
    embedded = {}
    for _f, _s, name, code in parser.extract_macros():
        if isinstance(code, bytes):
            code = code.decode("latin-1")
        embedded[name.replace(".bas", "").replace(".cls", "")] = code
    parser.close()

    print("%s" % BOOK.relative_to(REPO))
    print()

    for name, src in SOURCES.items():
        if name not in embedded:
            problems.append("%s is not in the workbook" % name)
            print("  %-22s MISSING" % name)
            continue
        a = normalise(embedded[name])
        b = normalise(src.read_text(encoding="utf-8", errors="replace"))
        if a == b:
            print("  %-22s current (%d lines)" % (name, len(b)))
        else:
            differing = sum(1 for x, y in zip(a, b) if x != y) + abs(len(a) - len(b))
            problems.append("%s differs from %s in %d line(s). Rebuild the workbook."
                            % (name, src.name, differing))
            print("  %-22s BEHIND, %d line(s) differ" % (name, differing))

    unexpected = [k for k in embedded
                  if k not in SOURCES and not k.startswith(("Sheet", "ThisWorkbook"))]
    if unexpected:
        problems.append("unexpected module(s): %s" % ", ".join(unexpected))

    # The shipped workbook must never carry a real key.
    config = embedded.get("AutoTraderConfig", "")
    if "<API_KEY>" not in config:
        problems.append("AutoTraderConfig does not carry the <API_KEY> placeholder")
    for found in re.findall(r'AT_API_KEY[^"]*"([^"]*)"', config):
        if found != "<API_KEY>":
            masked = (found[:4] + "..." + found[-4:]) if len(found) > 12 else "(short value)"
            problems.append("AT_API_KEY is a real key (%s), not the placeholder. "
                            "The shipped workbook must never carry one." % masked)

    print()
    if problems:
        for p in problems:
            print("  " + p)
        print("FAIL")
        return 1

    print("PASS: the workbook matches this repository")
    return 0


if __name__ == "__main__":
    sys.exit(main())
