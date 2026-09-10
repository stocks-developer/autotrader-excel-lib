# VBA check

Checks the modules in [`direct/`](../../direct) for the mistakes that stop a
workbook compiling.

This is a maintainer's tool. It is not part of the library and is not included
in the download, so nothing here needs to be installed to use AutoTrader Web.

```
python tools/vba-check/check-vba.py
```

```
Procedures found: 149
Files checked:    3

PASS -- no duplicates, no unclosed procedures, no unbalanced blocks,
        no calls to undefined helpers.
```

Exit code is 0 when everything passes and 1 when anything does not, so it works
as a pre-push check. It needs only Python — no Excel, and no settings changed.

## Why this exists

VBA compiles the **whole project** at once. A procedure defined twice, or a
`Function` whose `End Function` is missing, does not break one macro — it stops
every macro in the workbook, including the ones that were fine. So the cost of
one of these reaching a release is the entire library, not one function.

They are also invisible to a reading. A duplicate procedure looks perfectly
correct at both definitions; only seeing them together shows the problem, and
they are usually hundreds of lines apart.

## What it checks

| Check | Why |
|---|---|
| A procedure defined twice | Stops the project compiling. Each definition looks right on its own. |
| An unclosed `Function` / `Sub` | The next procedure is swallowed into the previous one. |
| Unbalanced `If`, `For`, `Do`, `With`, `Select` | Same effect, and the error Excel reports points somewhere else. |
| A call to a helper defined nowhere | Catches a renamed or removed function whose callers were missed. |
| An undeclared variable | These modules set `Option Explicit`, so one missing `Dim` stops the project. |

It understands VBA line continuations, and it strips comments without being
fooled by an apostrophe inside a string literal.

## What it does not check

It is not a VBA compiler and does not claim to be. Type mismatches, wrong
argument counts and anything that depends on the Excel object model are outside
what reading the text can establish. Those still need Excel.
