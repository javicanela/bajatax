import xlwings as xw
from pathlib import Path

ROOT = Path(r"C:\Users\LENOVO\Desktop\vscode\bajatax")
app = xw.App(visible=False)
wb = app.books.open(str(ROOT / "AUTOMATIZACION_v8.xlsm"))

print("Hojas encontradas:")
for sheet in wb.sheets:
    print(f"  '{sheet.name}'")

print("\nComponentes VBA:")
for comp in wb.api.VBProject.VBComponents:
    print(f"  '{comp.Name}' tipo={comp.Type}")

wb.close()
app.quit()
