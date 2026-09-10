"""
Checks the VBA modules for the mistakes that stop a workbook compiling.

VBA compiles the WHOLE project at once, so one of these faults in one module
stops every macro in the workbook, not just the one containing it. That makes
them worth catching before a release rather than after.

Reaching into Excel to compile the project needs the "Trust access to the VBA
project object model" setting turned on, which is off by default and is a
security setting rather than a build step. These checks read the .bas files as
text instead, so they need nothing but Python and can run anywhere.

Checks:
  1. A procedure defined twice.
  2. Unclosed Function/Sub bodies, and stray End Function/End Sub.
  3. Unbalanced If/End If, For/Next, Do/Loop, With/End With, Select/End Select.
  4. Calls to helpers that are not defined anywhere in the module set.
  5. Variables assigned but never declared, in modules that set Option Explicit.
  6. Module-level declarations placed after a procedure. VBA keeps them in a
     section at the top; one further down stops the module compiling, and the
     error a caller then sees names an innocent function as "not defined".
  7. Characters outside ASCII. VBA stores modules in a code page rather than
     UTF-8, so anything above ASCII does not survive an import: it arrives as
     question marks, and the module in the workbook is then no longer the
     module in this repository.

Run from the repository root:

    python tools/vba-check/check-vba.py

Exit code is 0 when everything passes and 1 when anything does not, so it works
as a pre-push check.
"""

import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]

FILES = [
    REPO / "direct/AutoTraderWebDirect.bas",
    REPO / "direct/AutoTraderConfig.bas",
    REPO / "direct/AutoTraderClientDirect.bas",
]

PROC = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Function|Sub|Property\s+\w+)\s+(\w+)",
    re.I,
)
END_PROC = re.compile(r"^\s*End\s+(Function|Sub|Property)\b", re.I)
DECL = re.compile(r"^\s*(?:Public|Private)\s+(?:Const\s+)?(\w+)", re.I)

failures = []

def check_ascii(path, text):
    """Check 7 -- see the module docstring."""
    odd = {}
    for number, line in enumerate(text.splitlines(), 1):
        for ch in line:
            if ord(ch) > 127:
                odd.setdefault(ord(ch), number)
    for code, number in sorted(odd.items()):
        failures.append(
            "%s line %d: character U+%04X is outside ASCII. VBA modules are "
            "stored in a code page, so it will not survive an import."
            % (path.name, number, code))



def strip_comment(line):
    """Remove a trailing VBA comment, honouring quoted strings."""
    out = []
    in_str = False
    for ch in line:
        if ch == '"':
            in_str = not in_str
        if ch == "'" and not in_str:
            break
        out.append(ch)
    return "".join(out)


all_procs = {}
defined_names = set()

for path in FILES:
    if not path.exists():
        failures.append(f"MISSING: {path}")
        continue

    text = path.read_text(encoding="utf-8", errors="replace")
    check_ascii(path, text)
    lines = text.splitlines()

    # Join VBA line continuations so a multi-line signature reads as one line.
    joined = []
    buffer = ""
    start_no = 0
    for no, raw in enumerate(lines, 1):
        code = strip_comment(raw).rstrip()
        if not buffer:
            start_no = no
        if code.endswith("_"):
            buffer += code[:-1] + " "
            continue
        joined.append((start_no, buffer + code))
        buffer = ""
    if buffer:
        joined.append((start_no, buffer))

    depth = 0
    open_proc = None
    open_line = 0
    blocks = {"If": 0, "For": 0, "Do": 0, "With": 0, "Select": 0}

    for no, code in joined:
        bare = code.strip()
        if not bare:
            continue

        m = PROC.match(bare)
        if m:
            kind, name = m.group(1), m.group(2)
            if open_proc:
                failures.append(
                    f"{path.name}:{no}: '{name}' starts while '{open_proc}' "
                    f"(line {open_line}) is still open -- missing End"
                )
            open_proc, open_line = name, no
            key = name.lower()
            if key in all_procs:
                failures.append(
                    f"{path.name}:{no}: DUPLICATE procedure '{name}' -- "
                    f"already defined at {all_procs[key]}"
                )
            all_procs[key] = f"{path.name}:{no}"
            defined_names.add(key)
            continue

        if END_PROC.match(bare):
            if not open_proc:
                failures.append(f"{path.name}:{no}: End with no open procedure")
            open_proc = None
            continue

        d = DECL.match(bare)
        if d and not open_proc:
            defined_names.add(d.group(1).lower())

        # Block balance, only inside a procedure.
        if open_proc:
            low = " " + bare.lower() + " "
            if re.match(r"^if\b.*\bthen$", bare, re.I):
                blocks["If"] += 1
            elif re.match(r"^end\s+if\b", bare, re.I):
                blocks["If"] -= 1
            if re.match(r"^for\b", bare, re.I):
                blocks["For"] += 1
            elif re.match(r"^next\b", bare, re.I):
                blocks["For"] -= 1
            if re.match(r"^do\b", bare, re.I):
                blocks["Do"] += 1
            elif re.match(r"^loop\b", bare, re.I):
                blocks["Do"] -= 1
            if re.match(r"^with\b", bare, re.I):
                blocks["With"] += 1
            elif re.match(r"^end\s+with\b", bare, re.I):
                blocks["With"] -= 1
            if re.match(r"^select\s+case\b", bare, re.I):
                blocks["Select"] += 1
            elif re.match(r"^end\s+select\b", bare, re.I):
                blocks["Select"] -= 1

    if open_proc:
        failures.append(
            f"{path.name}: '{open_proc}' (line {open_line}) is never closed"
        )

    for name, count in blocks.items():
        if count != 0:
            failures.append(f"{path.name}: {name} blocks unbalanced by {count:+d}")

# Undefined helper calls -- only for the library's own At*/Read*/Get* namespace,
# since VBA built-ins are far too many to list.
CALL = re.compile(r"\b((?:At|Read|Get|Is|Place|Modify|Cancel|SquareOff)\w+)\s*\(")
# VBA built-ins that happen to start with one of the prefixes above.
BUILTINS = {
    "atn",
    "isnumeric", "isempty", "isnull", "isarray", "isdate", "iserror",
    "isobject", "ismissing",
    "getobject", "getsetting", "getallsettings", "getattr",
}

for path in FILES:
    if not path.exists():
        continue
    for no, raw in enumerate(path.read_text(encoding="utf-8", errors="replace").splitlines(), 1):
        code = strip_comment(raw)
        for name in CALL.findall(code):
            low = name.lower()
            if low in defined_names or low in BUILTINS:
                continue
            failures.append(f"{path.name}:{no}: calls '{name}', which is not defined")

# ---------------------------------------------------------------------------
# Option Explicit: every local must be declared.
#
# These modules all set Option Explicit, which turns an undeclared variable from
# a silently-created Variant into a compile error that stops the whole project.
# It is the easiest mistake to make when adding code and the least visible when
# reading it.
# ---------------------------------------------------------------------------

ASSIGN = re.compile(r"^\s*(?:Set\s+)?([A-Za-z_]\w*)\s*=(?!=)")
FOR_VAR = re.compile(r"^\s*For\s+(?:Each\s+)?([A-Za-z_]\w*)\s", re.I)
DIM = re.compile(r"^\s*(?:Dim|Static|Const|ReDim(?:\s+Preserve)?|Private|Public)\s+(.*)$", re.I)
PARAMS = re.compile(r"\((.*)\)", re.S)

VBA_KEYWORDS = {
    "if", "then", "else", "elseif", "end", "for", "next", "do", "loop", "while",
    "wend", "select", "case", "exit", "function", "sub", "set", "let", "with",
    "on", "error", "goto", "resume", "call", "dim", "const", "true", "false",
    "nothing", "empty", "null", "and", "or", "not", "to", "step", "each", "in",
    "as", "byval", "byref", "optional", "preserve", "redim", "static", "public",
    "private", "option", "explicit", "attribute", "type", "enum", "declare",
}

for path in FILES:
    if not path.exists():
        continue

    lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    if not any(re.match(r"^\s*Option\s+Explicit", l, re.I) for l in lines):
        continue

    module_level = set()
    proc_name = None
    declared = set()
    body = []
    start = 0

    def flush(proc, names, statements, first_line):
        for no, text in statements:
            for pattern in (ASSIGN, FOR_VAR):
                m = pattern.match(text)
                if not m:
                    continue
                target = m.group(1)
                low = target.lower()
                if low in VBA_KEYWORDS:
                    continue
                if low in names or low in module_level or low in defined_names:
                    continue
                failures.append(
                    f"{path.name}:{no}: '{target}' is assigned in "
                    f"'{proc}' but never declared (Option Explicit)"
                )

    # Collect module-level declarations first.
    for raw in lines:
        code = strip_comment(raw).strip()
        m = DIM.match(code)
        if m and not re.match(r"^\s*(?:Private|Public)\s+(?:Function|Sub|Const)", code, re.I):
            for chunk in m.group(1).split(","):
                nm = re.match(r"\s*([A-Za-z_]\w*)", chunk)
                if nm:
                    module_level.add(nm.group(1).lower())

    buffer = ""
    joined = []
    start_no = 0
    for no, raw in enumerate(lines, 1):
        code = strip_comment(raw).rstrip()
        if not buffer:
            start_no = no
        if code.endswith("_"):
            buffer += code[:-1] + " "
            continue
        joined.append((start_no, buffer + code))
        buffer = ""

    for no, code in joined:
        bare = code.strip()
        m = PROC.match(bare)
        if m:
            if proc_name:
                flush(proc_name, declared, body, start)
            proc_name = m.group(2)
            declared = {proc_name.lower()}
            body = []
            start = no
            p = PARAMS.search(bare)
            if p:
                for chunk in p.group(1).split(","):
                    nm = re.search(r"(?:ByVal\s+|ByRef\s+|Optional\s+)*([A-Za-z_]\w*)", chunk, re.I)
                    if nm:
                        declared.add(nm.group(1).lower())
            continue

        if END_PROC.match(bare):
            if proc_name:
                flush(proc_name, declared, body, start)
            proc_name = None
            continue

        if proc_name:
            d = DIM.match(bare)
            if d:
                for chunk in d.group(1).split(","):
                    nm = re.match(r"\s*([A-Za-z_]\w*)", chunk)
                    if nm:
                        declared.add(nm.group(1).lower())
                continue
            body.append((no, bare))

    if proc_name:
        flush(proc_name, declared, body, start)

# Check 6. Module-level declarations must come BEFORE every procedure.
#
# VBA keeps all module-level variables and constants in a declarations section
# at the top of the module. One placed between procedures does not merely warn
# -- the module fails to compile, so EVERY function in it disappears, and the
# error a caller sees is "Sub or Function not defined" pointing at something
# entirely innocent.
#
# This shipped: a `Private AtQuietMode As Boolean` added next to the comment
# that explained it, 30 lines below the first procedure. Checks 1 to 5 all
# passed, because nothing was duplicated, unclosed, unbalanced or undefined.
# Only real VBA objected, and only once someone tried to run it.
MODULE_DECL = re.compile(
    r"^\s*(?:Public|Private|Global|Dim)\s+(?!Function\b|Sub\b|Property\b|Declare\b|Type\b|Enum\b)\w+",
    re.I,
)

for path in FILES:
    if not path.exists():
        continue
    seen_proc = False
    in_proc = False
    for no, raw in enumerate(path.read_text(encoding="utf-8", errors="replace").splitlines(), 1):
        bare = strip_comment(raw)
        if PROC.match(bare):
            seen_proc = True
            in_proc = True
            continue
        if END_PROC.match(bare):
            in_proc = False
            continue
        if in_proc or not seen_proc:
            continue
        if MODULE_DECL.match(bare):
            failures.append(
                f"{path.name}:{no}: module-level declaration after a procedure "
                f"-- VBA needs it above the first one, or the module will not compile: "
                f"{bare.strip()}"
            )

print(f"Procedures found: {len(all_procs)}")
print(f"Files checked:    {len([p for p in FILES if p.exists()])}")
print()

if failures:
    print(f"FAIL -- {len(failures)} problem(s):")
    for f in failures:
        print(f"  {f}")
    sys.exit(1)

print("PASS -- no duplicates, no unclosed procedures, no unbalanced blocks,")
print("        no calls to undefined helpers, no declaration below a procedure.")
sys.exit(0)
