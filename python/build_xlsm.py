"""
Create the final SalesIQ.xlsm on Windows using desktop Microsoft Excel.

Install:
    python -m pip install pywin32

Excel setting required:
File > Options > Trust Center > Trust Center Settings > Macro Settings >
Trust access to the VBA project object model
"""

from __future__ import annotations

import sys
from pathlib import Path

try:
    import win32com.client as win32
except ImportError as exc:
    raise SystemExit("Install pywin32 first: python -m pip install pywin32") from exc

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "excel" / "SalesIQ_Source.xlsx"
OUTPUT = ROOT / "excel" / "SalesIQ.xlsm"
VBA_MODULES = [
    ROOT / "vba" / "modNavigation.bas",
    ROOT / "vba" / "modDataAutomation.bas",
    ROOT / "vba" / "modUI.bas",
]

def delete_module(vbproject, module_name: str) -> None:
    for component in list(vbproject.VBComponents):
        if component.Name.lower() == module_name.lower():
            vbproject.VBComponents.Remove(component)
            return

def add_button(sheet, name, caption, macro, left, top, width, height, color):
    # 5 = msoShapeRoundedRectangle
    shape = sheet.Shapes.AddShape(5, left, top, width, height)
    shape.Name = name
    shape.TextFrame2.TextRange.Text = caption
    shape.TextFrame2.TextRange.Font.Name = "Segoe UI"
    shape.TextFrame2.TextRange.Font.Size = 12
    shape.TextFrame2.TextRange.Font.Bold = True
    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
    shape.Fill.ForeColor.RGB = color
    shape.Line.ForeColor.RGB = color
    shape.Shadow.Visible = True
    shape.OnAction = macro

def main():
    if sys.platform != "win32":
        raise SystemExit("Run this script on Windows with desktop Microsoft Excel installed.")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None

    try:
        wb = excel.Workbooks.Open(str(SOURCE.resolve()))
        wb.SaveAs(str(OUTPUT.resolve()), FileFormat=52)  # xlOpenXMLWorkbookMacroEnabled

        project = wb.VBProject
        for module in VBA_MODULES:
            delete_module(project, module.stem)
            project.VBComponents.Import(str(module.resolve()))

        home = wb.Worksheets("Home")

        for n in [
            "btnUpload","btnDashboard","btnKPI","btnForecast",
            "btnReports","btnRefresh","btnExport"
        ]:
            try:
                home.Shapes(n).Delete()
            except Exception:
                pass

        add_button(home,"btnUpload","UPLOAD CSV","UploadCSV",310,175,135,40,0xFF7929)
        add_button(home,"btnDashboard","OPEN DASHBOARD","ShowDashboard",475,175,135,40,0x55C522)
        add_button(home,"btnKPI","KPI CENTER","ShowKPI",640,175,135,40,0xF65C8B)
        add_button(home,"btnForecast","FORECAST","ShowForecast",805,175,135,40,0x1673F9)
        add_button(home,"btnReports","REPORTS","ShowReports",970,175,135,40,0x9948EC)
        add_button(home,"btnRefresh","REFRESH ALL","RefreshSalesIQ",800,410,135,36,0xC56321)
        add_button(home,"btnExport","EXPORT PDF","ExportDashboardPDF",950,410,135,36,0x391B07)

        # Workbook opens on Home.
        wb.Worksheets("Home").Activate()

        # Hide worksheet gridlines on key pages.
        for sheet_name in ["Home","Dashboard","KPI_Center","Forecast","Reports","Data_Dictionary","Settings"]:
            ws = wb.Worksheets(sheet_name)
            ws.Activate()
            excel.ActiveWindow.DisplayGridlines = False
            excel.ActiveWindow.DisplayHeadings = False
            excel.ActiveWindow.Zoom = 90

        wb.Worksheets("Home").Activate()
        wb.Save()

        print(f"CREATED: {OUTPUT}")
        print("Open SalesIQ.xlsm in desktop Excel and choose Enable Content.")

        import traceback

    except Exception:
        traceback.print_exc()
        raise
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()

if __name__ == "__main__":
    main()
