#!/usr/bin/env python3
"""
instalar_vba.py â€” BajaTax v4 FINAL
Instala todos los mÃ³dulos VBA en una copia del archivo v2.xlsm
Requiere: pip install xlwings openpyxl
Ejecutar con Excel cerrado (o al menos el archivo destino cerrado).
"""

from pathlib import Path
import json

ROOT     = Path(__file__).parent.parent
config   = json.loads((ROOT / "bajatax.config.json").read_text(encoding="utf-8"))
SRC_FILE = str(ROOT / config["xlsm_source"])
DST_FILE = str(ROOT / config["xlsm_output"])
VBA_DIR  = str(ROOT / config["vba_modules_dir"])
import shutil
import os
import sys
import time
import subprocess

BASE_DIR = str(ROOT)
VBA_DIR  = str(ROOT / config["vba_modules_dir"])
SRC_FILE = str(ROOT / config["xlsm_source"])
DST_FILE = str(ROOT / config["xlsm_output"])

# MÃ³dulos estÃ¡ndar (.bas) con su nombre dentro de VBA
MODULOS_STD = [
    ("01_Mod_Sistema.bas",          "Mod_Sistema"),
    ("02_Mod_ImportarArchivos.bas", "Mod_ImportarArchivos"),
    ("03_Mod_WhatsApp.bas",         "WhatsApp"),
    ("04_Mod_PDF.bas",              "PDF"),
    ("07_Mod_MasivoPDF.bas",        "Mod_MasivoPDF"),
    ("08_Mod_BuscadorCliente.bas",  "Mod_BuscadorCliente"),
    ("09_Mod_FormatoGlobal.bas",      "Mod_FormatoGlobal"),
]

# MÃ³dulos de hoja (el nombre del atributo coincide con la hoja)
MODULOS_HOJA = [
    ("05_Hoja_OPERACIONES.bas",     "OPERACIONES"),
    ("06_Hoja_DIRECTORIO.bas",      "DIRECTORIO"),
    ("10_Hoja_BuscadorCliente.bas", "BUSCADOR CLIENTE"),
    ("11_Hoja_REGISTROS.bas",       "REGISTROS"),
]


def log(msg):
    print(f"  {msg}")


def hacer_copia():
    if not os.path.exists(SRC_FILE):
        print(f"ERROR: No se encuentra el archivo fuente:\n  {SRC_FILE}")
        sys.exit(1)
    shutil.copy2(SRC_FILE, DST_FILE)
    log(f"Copia creada: {os.path.basename(DST_FILE)}")


def leer_bas(nombre_archivo):
    ruta = os.path.join(VBA_DIR, nombre_archivo)
    if not os.path.exists(ruta):
        print(f"  ADVERTENCIA: No se encuentra {nombre_archivo}")
        return None
    with open(ruta, "r", encoding="utf-8") as f:
        return f.read()


def instalar_con_xlwings():
    try:
        import xlwings as xw
    except ImportError:
        log("xlwings no instalado. Instalando...")
        subprocess.run([sys.executable, "-m", "pip", "install", "xlwings"], check=True)
        import xlwings as xw

    log("Abriendo archivo en Excel...")
    app = xw.App(visible=True, add_book=False)
    wb  = app.books.open(DST_FILE)
    vba = wb.api.VBProject.VBComponents

    # â”€â”€ Eliminar mÃ³dulos estÃ¡ndar existentes â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    nombres_a_borrar = [m[1] for m in MODULOS_STD]
    componentes = list(vba)
    for comp in componentes:
        try:
            if comp.Name in nombres_a_borrar and comp.Type == 1:  # 1=vbext_ct_StdModule
                vba.Remove(comp)
                log(f"  Eliminado mÃ³dulo existente: {comp.Name}")
        except Exception:
            pass

    # â”€â”€ Instalar mÃ³dulos estÃ¡ndar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    for bas_file, mod_name in MODULOS_STD:
        codigo = leer_bas(bas_file)
        if codigo is None:
            continue
        try:
            nuevo = vba.Add(1)  # vbext_ct_StdModule
            nuevo.Name = mod_name
            nuevo.CodeModule.AddFromString(codigo)
            log(f"  âœ“ MÃ³dulo instalado: {mod_name}")
        except Exception as e:
            log(f"  âœ— Error en {mod_name}: {e}")

    # â”€â”€ Instalar cÃ³digo de hojas â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    for bas_file, sheet_name in MODULOS_HOJA:
        codigo = leer_bas(bas_file)
        if codigo is None:
            continue
        try:
            # Buscar el componente de la hoja por nombre
            hoja_comp = None
            for comp in list(vba):
                if comp.Type == 100:  # vbext_ct_Document (hoja)
                    try:
                        if comp.Properties("Name").Value == sheet_name:
                            hoja_comp = comp
                            break
                    except Exception:
                        if comp.Name.replace("Hoja", "").replace("Sheet", "") in sheet_name:
                            hoja_comp = comp
                            break

            if hoja_comp is None:
                # Buscar por caption de la hoja
                for sh in wb.sheets:
                    if sh.name == sheet_name:
                        for comp in list(vba):
                            if comp.Type == 100:
                                try:
                                    nm = comp.Properties("_CodeName").Value
                                    if nm == sh.api.CodeName:
                                        hoja_comp = comp
                                        break
                                except Exception:
                                    pass
                        break

            if hoja_comp is not None:
                # Limpiar cÃ³digo existente
                cm = hoja_comp.CodeModule
                if cm.CountOfLines > 0:
                    cm.DeleteLines(1, cm.CountOfLines)
                # Eliminar la primera lÃ­nea Attribute VB_Name del cÃ³digo de hoja
                lineas = codigo.split("\n")
                codigo_limpio = "\n".join([l for l in lineas if not l.startswith("Attribute VB_Name")])
                cm.AddFromString(codigo_limpio)
                log(f"  âœ“ CÃ³digo de hoja instalado: {sheet_name}")
            else:
                log(f"  ! No se encontrÃ³ la hoja '{sheet_name}' en el proyecto VBA")
        except Exception as e:
            log(f"  âœ— Error en hoja {sheet_name}: {e}")

    # â”€â”€ Guardar y cerrar â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
    log("Guardando...")
    wb.save()
    wb.close()
    app.quit()
    log(f"âœ“ Archivo guardado: {os.path.basename(DST_FILE)}")


def instalar_con_applescript():
    """Fallback: usar AppleScript para instalar VBA en macOS"""
    log("Usando AppleScript como mÃ©todo alternativo...")

    # Generar script de instalaciÃ³n VBA interno
    macro_installer = generar_macro_instalador()

    # Guardar el macro en un archivo temporal
    tmp_bas = os.path.join(BASE_DIR, "_temp_installer.bas")
    with open(tmp_bas, "w", encoding="utf-8") as f:
        f.write(macro_installer)

    script = f'''
    tell application "Microsoft Excel"
        activate
        set wb to open workbook workbook file name "{DST_FILE}"
        set vbp to VBA project of wb

        -- Ejecutar macro de instalaciÃ³n
        run macro macro name "InstalarTodo" of wb

        save wb
        close wb
    end tell
    '''

    result = subprocess.run(["osascript", "-e", script],
                           capture_output=True, text=True, timeout=120)
    if result.returncode == 0:
        log("âœ“ InstalaciÃ³n vÃ­a AppleScript completada")
    else:
        log(f"âœ— Error AppleScript: {result.stderr}")

    if os.path.exists(tmp_bas):
        os.remove(tmp_bas)


def generar_macro_instalador():
    """Genera un macro VBA que se auto-instala leyendo los .bas del disco"""
    partes = []
    partes.append("Sub InstalarTodo()")
    partes.append("    Dim vbp As Object")
    partes.append("    Set vbp = ThisWorkbook.VBProject")
    partes.append(f"    Dim baseDir As String: baseDir = \"{VBA_DIR.replace(chr(92), '/')}/\"")
    partes.append("")

    for bas_file, mod_name in MODULOS_STD:
        ruta = os.path.join(VBA_DIR, bas_file).replace("\\", "/")
        partes.append(f"    ' Instalar {mod_name}")
        partes.append(f"    On Error Resume Next")
        partes.append(f"    vbp.VBComponents.Remove vbp.VBComponents(\"{mod_name}\")")
        partes.append(f"    On Error GoTo 0")
        partes.append(f"    vbp.VBComponents.Import \"{ruta}\"")
        partes.append("")

    partes.append("    MsgBox \"MÃ³dulos instalados correctamente.\", vbInformation, \"BajaTax\"")
    partes.append("End Sub")
    return "\n".join(partes)


def crear_carpetas_necesarias():
    """Crea las carpetas de trabajo que el sistema necesita"""
    carpetas = [
        os.path.join(BASE_DIR, "IMPORTAR"),
        os.path.join(BASE_DIR, "SALIDA_PDF"),
        os.path.join(BASE_DIR, "LOGOS"),
    ]
    for c in carpetas:
        os.makedirs(c, exist_ok=True)
        log(f"  Carpeta: {os.path.basename(c)}")


def main():
    print("=" * 60)
    print("  BajaTax v4 â€” Instalador VBA")
    print("=" * 60)
    print()

    print("â–º Paso 1: Crear copia del archivo base...")
    hacer_copia()

    print()
    print("â–º Paso 2: Crear carpetas necesarias...")
    crear_carpetas_necesarias()

    print()
    print("â–º Paso 3: Instalar mÃ³dulos VBA...")

    try:
        instalar_con_xlwings()
        ok = True
    except Exception as e:
        log(f"xlwings fallÃ³: {e}")
        log("Intentando mÃ©todo alternativo (AppleScript)...")
        try:
            instalar_con_applescript()
            ok = True
        except Exception as e2:
            log(f"AppleScript tambiÃ©n fallÃ³: {e2}")
            ok = False

    print()
    if ok:
        print("=" * 60)
        print("  âœ“ INSTALACIÃ“N COMPLETADA")
        print(f"  Archivo: {os.path.basename(DST_FILE)}")
        print("=" * 60)
        print()
        print("PRÃ“XIMOS PASOS:")
        print("1. Abre AUTOMATIZACION_v4_FINAL.xlsm")
        print("2. Habilita macros cuando Excel lo solicite")
        print("3. En DIRECTORIO, ejecuta InicializarEncabezadosDirectorio()")
        print("4. En REPORTES CXC, asigna botÃ³n â†’ ActualizarReportesCXC")
        print("5. En BUSCADOR CLIENTE, asigna botÃ³n â†’ EjecutarBusqueda / LimpiarBuscador")
    else:
        print("=" * 60)
        print("  âœ— La instalaciÃ³n automÃ¡tica fallÃ³.")
        print("  Usa la instalaciÃ³n manual con los archivos .bas")
        print("  Ver: VBA_CODIGO/00_GUIA_INSTALACION.md")
        print("=" * 60)


if __name__ == "__main__":
    main()


