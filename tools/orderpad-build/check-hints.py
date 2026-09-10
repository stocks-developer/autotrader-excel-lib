"""Assert the finished workbook carries its guidance in the RIGHT place.

Two things went wrong here and both were invisible in the build output:

  a data-validation input message applies to a whole COLUMN, so guidance meant
  for a heading popped up on every data cell that was clicked;

  and a build that failed on a file lock left the canonical workbook stale
  while the fix sat under a different name, so the file being looked at was
  not the file being changed.

This reads the finished workbook and fails on either:

    python tools/orderpad-build/check-hints.py

With no argument it reads the workbook that SHIPS, so it needs neither Excel
nor a build. Pass a path to check a freshly built one before it replaces that.
"""
import re
import sys
import zipfile
from pathlib import Path

# tools/orderpad-build/check-hints.py -> the repository root is two levels up.
REPO = Path(__file__).resolve().parents[2]
BOOK = (Path(sys.argv[1]) if len(sys.argv) > 1
        else REPO / "clients/excel/current/samples/AutoTraderWeb-OrderPad.xlsm")

if not BOOK.exists():
    print("workbook not found: %s" % BOOK)
    sys.exit(2)

z = zipfile.ZipFile(BOOK)
names = z.namelist()
problems = []

# 1. No input messages anywhere. Dropdowns are fine; prompts are not.
prompts = 0
lists = 0
for n in names:
    if not (n.startswith("xl/worksheets/") and n.endswith(".xml")):
        continue
    x = z.read(n).decode("utf-8", "replace")
    for dv in re.findall(r"<dataValidation [^>]*?(?:/>|>)", x):
        if 'showInputMessage="1"' in dv or "prompt=" in dv:
            prompts += 1
        if 'type="list"' in dv:
            lists += 1

if prompts:
    problems.append("%d data-validation input message(s) left -- these show on "
                    "every data cell, not on the heading" % prompts)

# 2. The heading notes must be there instead.
comment_parts = [n for n in names if n.startswith("xl/comments")]
total_comments = 0
for n in comment_parts:
    x = z.read(n).decode("utf-8", "replace")
    total_comments += len(re.findall(r"<comment ", x))

if not comment_parts:
    problems.append("no comment parts at all -- the heading notes were never attached")

expected = 23
if total_comments != expected:
    problems.append("found %d heading note(s), expected %d" % (total_comments, expected))

# 3. Every note must be anchored on the heading row, never on a data cell.
for n in comment_parts:
    x = z.read(n).decode("utf-8", "replace")
    for ref in re.findall(r'<comment ref="([A-Z]+)(\d+)"', x):
        if int(ref[1]) != 2:
            problems.append("a note is anchored on %s%s, which is not the heading row"
                            % (ref[0], ref[1]))

print("%s" % BOOK)
print("  list dropdowns      : %d" % lists)
print("  input messages      : %d  (want 0)" % prompts)
print("  heading notes       : %d  (want %d)" % (total_comments, expected))

if problems:
    print()
    for p in problems:
        print("  " + p)
    print("FAIL")
    sys.exit(1)

print("PASS: guidance is on the headings and nowhere else")
