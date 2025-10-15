# workbook_builder/core/discovery.py
import importlib, pkgutil
import report_builder.sheets as sheets_pkg

def discover_sheets() -> None:
    for m in pkgutil.iter_modules(sheets_pkg.__path__, sheets_pkg.__name__ + "."):
        importlib.import_module(m.name)
