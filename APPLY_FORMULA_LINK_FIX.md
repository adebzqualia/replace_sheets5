# Apply the Feuil1 formula repair

Extract `POPS_formula_links_fix.zip` into the project folder containing
`main.py`, replacing the supplied Python files. Keep the `pops_update` folder
structure. The new `pops_update/formula_links.py` file is required; replacing
only `backend.py` is not sufficient.

The archive contains code and this guide. It does not contain configuration
files or workbook data, so your current paths and UK sheet selection remain.

Run from your existing Python environment:

```powershell
python main.py --debug
```

This replaces the earlier fix that stopped when Excel's `ChangeLink` left
`Feuil1` behind. The updater now repairs the affected formulas, defined names
and chart series directly through Excel, verifies the changed formulas, then
performs a full recalculation before saving. For example:

```excel
'[Feuil1]Belgium_OBS KPI'!U16
```

becomes:

```excel
'Belgium_OBS KPI'!U16
```

The fallback checks that the local destination sheets exist. It preserves
text literals, ordinary external source links, and worksheet protection.
Missing targets and unresolved formulas produce an error naming the affected
sheet/cell or defined name, and an existing output is preserved on failure.

If Excel still lists `Feuil1` after the checked worksheet formulas, names and
chart series have been repaired, the output report records a warning instead
of failing solely on that link-list entry. References in other Excel objects
are outside this formula check. No links are broken into fixed values.

The report includes `repaired_formula_references`, counting formula cells,
array ranges, defined names and chart series written by the fallback.

If the output already exists, choose a new `output_workbook` or set
`overwrite_output: true` in your configuration. Close that output in Excel
before running. The input workbooks are not saved.

Portable regression tests pass, including a simulated ineffective
`ChangeLink` call followed by an actual formula repair. Native Excel tests
could not be run in this development environment.
