"""Repair known temporary workbook qualifiers through Excel's formula API.

Only reference tokens are changed. Text literals, other external workbooks,
cell formatting and cached business values are never used as replacements.
"""
from __future__ import annotations

import ntpath
import re

from .excel_session import ExcelSessionError


# Match strings first so reference-looking text inside them stays untouched.
# Quoted identifiers include Excel's doubled-apostrophe escape convention.
_TOKENS = re.compile(
    r'"(?:[^"]|"")*"'
    r"|'(?:[^']|'')*'!"
    r"|\[[^\[\]]+\][^\s!+*/^&=<>(),;{}\"']*!"
    r"|[\w.\\]+!"
)
_EXTERNAL = re.compile(r"^(.*?)\[([^\[\]]+)\](.*)$")
_NAME = re.compile(r"[\w.\\]+")


def localize_formula(formula, aliases, sheet_names, defined_names=()):
    """Remove known workbook qualifiers only when their local targets exist."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    folded = formula.casefold()
    if not any(alias.casefold() in folded for alias in aliases):
        return formula
    aliases = {name.casefold() for name in aliases}
    sheets = {name.casefold() for name in sheet_names}
    names = {name.casefold() for name in defined_names}

    def replace(match):
        token = match.group()
        if token.startswith('"'):
            return token
        qualifier = token[:-1]
        if qualifier.startswith("'"):
            qualifier = qualifier[1:-1].replace("''", "'")
        external = _EXTERNAL.fullmatch(qualifier)
        if external:
            _, book, sheet = external.groups()
            if book.casefold() not in aliases:
                return token
            if sheet:
                # Excel uses ':' for a 3-D reference between two sheet names.
                missing = [name for name in sheet.split(":") if name.casefold() not in sheets]
                if missing:
                    raise ExcelSessionError(
                        f"Cannot reconnect {token}: destination worksheet(s) missing: {', '.join(missing)}")
                return "'" + sheet.replace("'", "''") + "'!"
        else:
            if qualifier.casefold() in sheets or ntpath.basename(qualifier).casefold() not in aliases:
                return token
        # Workbook-scoped defined name, e.g. 'Feuil1'!Total or [Feuil1]!Total.
        name = _NAME.match(formula, match.end())
        if name is None or name.group().casefold() not in names:
            raise ExcelSessionError(f"Cannot reconnect {token}: matching local defined name is missing")
        return ""

    return _TOKENS.sub(replace, formula)


def _items(session, collection):
    for index in range(1, session.call(lambda: int(collection.Count)) + 1):
        yield session.call(lambda index=index: collection.Item(index))


def rewrite_workbook_links(session, workbook, aliases, *, verify_only=False, progress=None):
    """Repair/verify worksheet, name and chart formulas, including hidden sheets.

    The caller must unlock worksheets before repairing. Read-only verification
    works on protected sheets. Read rectangular formula areas in bulk, then
    write only changed cells, or entire legacy array ranges when necessary.
    """
    sheets = list(_items(session, workbook.Worksheets))
    sheet_names = [session.call(lambda ws=ws: str(ws.Name)) for ws in sheets]
    names = list(_items(session, workbook.Names))
    defined_names = [session.call(lambda name=name: str(name.Name)) for name in names]
    # Worksheet-scoped names cannot replace an external workbook-scoped name.
    defined_names = [name for name in defined_names if "!" not in name]
    repaired = 0

    def transform(value, label):
        try:
            updated = localize_formula(value, aliases, sheet_names, defined_names)
        except ExcelSessionError as exc:
            raise ExcelSessionError(f"{label}: {exc}") from exc
        if verify_only and updated != value:
            raise ExcelSessionError(f"Temporary workbook reference remains in {label}: {value}")
        return updated

    def write(target, attribute, value, label):
        nonlocal repaired
        session.run(f"Reconnecting {label}", lambda: setattr(target, attribute, value))
        actual = session.run(f"Checking {label}", lambda: getattr(target, attribute))
        if transform(actual, label) != actual:
            raise ExcelSessionError(f"Excel did not replace the temporary reference in {label}")
        if isinstance(actual, str) and "#REF!" in actual and "#REF!" not in value:
            raise ExcelSessionError(f"Excel introduced #REF! while reconnecting {label}")
        repaired += 1

    def property_formula(target, attribute, label):
        value = session.run(f"Reading {label}", lambda: getattr(target, attribute))
        updated = transform(value, label)
        if updated != value:
            write(target, attribute, updated, label)

    for worksheet, sheet_name in zip(sheets, sheet_names):
        used = session.call(lambda: worksheet.UsedRange)
        # HasFormula is None for mixed ranges, False only when no formulas exist.
        has_formulas = session.call(lambda: used.HasFormula)
        if has_formulas is None or bool(has_formulas):
            areas = session.run(f"Reading formula cells on {sheet_name!r}",
                                lambda: used.SpecialCells(-4123).Areas)
            processed_arrays = set()
            for area in _items(session, areas):
                attribute = "Formula2" if session.call(lambda: hasattr(area, "Formula2")) else "Formula"
                values = session.run(f"Reading formulas on {sheet_name!r}", lambda: getattr(area, attribute))
                if not isinstance(values, (tuple, list)):
                    values = ((values,),)
                for row, values_row in enumerate(values, start=1):
                    for column, value in enumerate(values_row, start=1):
                        # Resolve addresses only for affected cells to avoid a
                        # COM round trip for every formula in a large report.
                        try:
                            updated = localize_formula(value, aliases, sheet_names, defined_names)
                        except ExcelSessionError:
                            cell = session.call(lambda: area.Cells(row, column))
                            label = f"{sheet_name}!{session.call(lambda: cell.Address)}"
                            transform(value, label)  # raise with the exact cell
                            raise
                        if updated == value:
                            continue
                        cell = session.call(lambda: area.Cells(row, column))
                        label = f"{sheet_name}!{session.call(lambda: cell.Address)}"
                        transform(value, label)  # verify_only fails before writing
                        if session.call(lambda: cell.HasArray):
                            array = session.call(lambda: cell.CurrentArray)
                            address = session.call(lambda: str(array.Address))
                            if address in processed_arrays:
                                continue
                            processed_arrays.add(address)
                            value = session.call(lambda: array.FormulaArray)
                            updated = transform(value, f"{sheet_name}!{address}")
                            if len(updated) > 255:
                                raise ExcelSessionError(
                                    f"{sheet_name}!{address}: repaired legacy array formula exceeds Excel's "
                                    "255-character FormulaArray limit")
                            write(array, "FormulaArray", updated, f"{sheet_name}!{address}")
                        else:
                            write(cell, attribute, updated, label)
            if progress and not verify_only:
                progress(f"Checked temporary formula links: {sheet_name}")

    for name in names:
        label = "defined name " + session.call(lambda: str(name.Name))
        property_formula(name, "RefersTo", label)

    def chart_formulas(chart):
        chart_name = session.call(lambda: str(chart.Name))
        collection = session.call(lambda: chart.FullSeriesCollection()
                                  if hasattr(chart, "FullSeriesCollection") else chart.SeriesCollection())
        for index, series in enumerate(_items(session, collection), start=1):
            property_formula(series, "Formula", f"chart {chart_name!r}, series {index}")

    for worksheet in sheets:
        charts = session.call(lambda: worksheet.ChartObjects())
        for obj in _items(session, charts):
            chart_formulas(session.call(lambda: obj.Chart))
    for chart in _items(session, workbook.Charts):
        chart_formulas(chart)
    return repaired
