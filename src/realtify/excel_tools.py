from __future__ import annotations

from pathlib import Path
from typing import Any


def excel_com_available() -> bool:
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except Exception:
        return False
    return True


class ExcelApp:
    def __init__(self, *, visible: bool = False) -> None:
        self.visible = visible
        self.app: Any | None = None
        self._pythoncom: Any | None = None

    def __enter__(self):
        try:
            import pythoncom
            import win32com.client
        except Exception as exc:
            raise RuntimeError(
                "Microsoft Excel COM is unavailable on this machine. "
                "Run the workflow on Windows with Microsoft Excel installed, "
                "or use a cross-platform Excel implementation."
            ) from exc
        pythoncom.CoInitialize()
        self._pythoncom = pythoncom
        self.app = win32com.client.DispatchEx("Excel.Application")
        self.app.Visible = self.visible
        self.app.DisplayAlerts = False
        return self.app

    def __exit__(self, exc_type, exc, tb) -> None:
        if self.app is not None:
            self.app.Quit()
            self.app = None
        if self._pythoncom is not None:
            self._pythoncom.CoUninitialize()
            self._pythoncom = None


def excel_path(path: Path) -> str:
    resolved = path.resolve()
    try:
        import win32api

        return win32api.GetShortPathName(str(resolved))
    except Exception:
        return str(resolved)


def inspect_workbook(path: Path) -> dict[str, Any]:
    with ExcelApp() as excel:
        wb = excel.Workbooks.Open(excel_path(path), 0, True)
        try:
            sheets = []
            for ws in wb.Worksheets:
                used = ws.UsedRange
                sheets.append(
                    {
                        "name": ws.Name,
                        "used_range": used.Address,
                        "rows": used.Rows.Count,
                        "cols": used.Columns.Count,
                    }
                )
            return {"name": wb.Name, "sheets": sheets}
        finally:
            wb.Close(False)
