import xlwings as xw
from pathlib import Path

ROOT = Path(r"C:\Users\LENOVO\Desktop\vscode\bajatax")
app = xw.App(visible=False)
wb = app.books.open(str(ROOT / "AUTOMATIZACION_v8.xlsm"))

# Primero mapear nombre visible -> CodeName
print("Mapeando hojas...")
mapa = {}
for sheet in wb.sheets:
    code_name = sheet.api.CodeName
    mapa[sheet.name] = code_name
    print(f"  '{sheet.name}' -> CodeName: '{code_name}'")

hojas = [
    ("05_Hoja_OPERACIONES.bas",     "OPERACIONES"),
    ("06_Hoja_DIRECTORIO.bas",      "DIRECTORIO"),
    ("10_Hoja_BuscadorCliente.bas", "BUSCADOR CLIENTE"),
    ("11_Hoja_REGISTROS.bas",       "REGISTROS"),
]

for bas_file, sheet_name in hojas:
    codigo = (ROOT / "src" / "vba-modules" / bas_file).read_text(encoding="utf-8-sig")
    lineas = [l for l in codigo.split("\n") if not l.startswith("Attribute VB_")]
    codigo_limpio = "\n".join(lineas)
    try:
        code_name = mapa.get(sheet_name)
        if not code_name:
            print(f"❌ Sin CodeName: '{sheet_name}'")
            continue
        comp = wb.api.VBProject.VBComponents(code_name)
        cm = comp.CodeModule
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(codigo_limpio)
        print(f"✅ {sheet_name} -> {code_name}")
    except Exception as e:
        print(f"❌ {sheet_name}: {e}")

wb.save()
wb.close()
app.quit()
print("Listo")
