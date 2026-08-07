"""
Build final SalesIQ.xlsm with VBA automation and premium buttons.

Windows requirements:
- Desktop Microsoft Excel
- Python 3.10+
- pywin32

Install:
    python -m pip install pywin32

Excel:
File > Options > Trust Center > Trust Center Settings > Macro Settings >
Trust access to the VBA project object model
"""

from __future__ import annotations
import sys
from pathlib import Path

try:
    import win32com.client as win32
except ImportError as exc:
    raise SystemExit("Run: python -m pip install pywin32") from exc

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "excel" / "SalesIQ_Batch4_Source.xlsx"
OUTPUT = ROOT / "excel" / "SalesIQ.xlsm"

MODULES = [
    ROOT / "vba" / "modNavigation.bas",
    ROOT / "vba" / "modDataAutomation.bas",
    ROOT / "vba" / "modUI.bas",
    ROOT / "vba" / "modCharts.bas",
    ROOT / "vba" / "modButtons.bas",
    ROOT / "vba" / "modTheme.bas",
    ROOT / "vba" / "modValidation.bas",
]

def remove_if_exists(vbproject, name: str) -> None:
    for component in list(vbproject.VBComponents):
        if component.Name.lower() == name.lower():
            vbproject.VBComponents.Remove(component)
            return

def main() -> None:
    if sys.platform != "win32":
        raise SystemExit("Run this builder on Windows with desktop Excel installed.")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None

    try:
        wb = excel.Workbooks.Open(str(SOURCE.resolve()))
        wb.SaveAs(str(OUTPUT.resolve()), FileFormat=52)

        vbproject = wb.VBProject
        for module in MODULES:
            if not module.exists():
                raise FileNotFoundError(module)
            remove_if_exists(vbproject, module.stem)
            vbproject.VBComponents.Import(str(module.resolve()))

        # Call VBA to build buttons and restore the approved dark theme.
        excel.Run(f"'{wb.Name}'!CreateSalesIQButtons")
        excel.Run(f"'{wb.Name}'!ApplyDarkTheme")

        for sheet_name in [
            "Home","Dashboard","KPI_Center","Forecast","Product_Analysis",
            "Region_Analysis","AI_Insights","Reports","Data_Dictionary","Settings"
        ]:
            wb.Worksheets(sheet_name).Activate()
            excel.ActiveWindow.DisplayGridlines = False
            excel.ActiveWindow.DisplayHeadings = False
            excel.ActiveWindow.Zoom = 90

        wb.Worksheets("Home").Activate()
        wb.Save()
        print(f"CREATED: {OUTPUT}")

    except Exception as exc:
        raise SystemExit(
            "XLSM build failed. Confirm Excel is installed and Trust access "
            "to the VBA project object model is enabled.\n"
            f"Details: {exc}"
        ) from exc
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()

if __name__ == "__main__":
    main()
