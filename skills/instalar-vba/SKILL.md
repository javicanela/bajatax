---
name: instalar-vba
description: >
  Usar esta skill cuando el usuario necesite inyectar módulos VBA al archivo Excel
  de BajaTax sin abrir manualmente el editor VBA. Activar cuando el usuario diga
  "instalar los módulos", "inyectar el VBA", "cómo ejecuto instalar_vba.py",
  "el Excel no tiene los cambios que generé", "cómo paso el .bas al Excel",
  o cuando termine de generar o modificar módulos .bas y quiera verlos en el Excel.
  También activar si instalar_vba.py da error al ejecutarse. Esta skill explica
  el flujo completo de instalación y los requisitos previos — úsala siempre
  que haya que sincronizar los módulos .bas con AUTOMATIZACION_v7.xlsm.
---

# Instalar VBA — BajaTax

## Qué hace instalar_vba.py

Inyecta automáticamente los módulos `.bas` de `src/vba-modules/` al archivo
`AUTOMATIZACION_v7.xlsm` sin necesidad de abrir el editor VBA ni copiar/pegar.
Detecta automáticamente el OS (Mac/Windows) y usa la ruta correcta.

---

## Requisitos previos OBLIGATORIOS

```
1. Excel CERRADO — si está abierto, el script falla con error de acceso
2. Python 3 instalado y accesible desde terminal
3. xlwings instalado: pip install xlwings
4. En Mac: xlwings necesita permisos de automatización (System Preferences → Privacy → Automation)
```

---

## Cómo ejecutar

```powershell
# Desde la raíz del proyecto BajataxV2 (ajustar a tu ruta real):
cd ruta\a\BajataxV2

# Verificar que Excel esté cerrado primero
# Luego ejecutar:
python src/python/instalar_vba.py
```

En Mac:
```bash
cd ~/Desktop/vscode/BajataxV2
python3 src/python/instalar_vba.py
```

---

## Qué hace el script paso a paso

```python
"""
instalar_vba.py — Inyecta módulos .bas al Excel de BajaTax

Flujo:
1. Detectar OS y encontrar AUTOMATIZACION_v7.xlsm en la raíz del proyecto
2. Listar todos los .bas en src/vba-modules/ ordenados por nombre
3. Abrir el Excel programáticamente con xlwings
4. Para cada .bas:
   a. Leer el contenido del archivo
   b. Si ya existe un módulo con ese nombre en el Excel → reemplazarlo
   c. Si no existe → importarlo como módulo nuevo
5. Guardar el Excel
6. Cerrar Excel (opcional — preguntar al usuario)
7. Reportar: X módulos instalados, Y actualizados, Z errores
"""
```

---

## Orden de instalación de módulos

Los módulos deben instalarse en este orden (el prefijo numérico lo garantiza):

```
00_Bootstrap_Installer.bas   → funciones de arranque
01_Mod_Sistema.bas           → LogEvento, funciones globales
02_Mod_ImportarArchivos.bas  → importación inteligente
03_Mod_WhatsApp.bas          → envío WA y URLEncodeUTF8
04_Mod_PDF.bas               → generación PDF via Python
05_Hoja_OPERACIONES.bas      → eventos de la hoja principal
06_Hoja_DIRECTORIO.bas       → eventos del directorio
07_Mod_MasivoPDF.bas         → generación PDF masiva
08_Mod_BuscadorCliente.bas   → búsqueda avanzada
09_Mod_FormatoGlobal.bas     → formatos y estilos
10_Hoja_BuscadorCliente.bas  → eventos del buscador
11_Hoja_REGISTROS.bas        → BeforeDoubleClick principal
```

> Importante: 01_Mod_Sistema.bas debe instalarse antes que cualquier otro
> porque contiene LogEvento() que todos los demás usan.

---

## Errores comunes y solución

| Error | Causa | Solución |
|---|---|---|
| `PermissionError: [Errno 13]` | Excel está abierto | Cerrar Excel completamente y reintentar |
| `ModuleNotFoundError: xlwings` | xlwings no instalado | `pip install xlwings` |
| `xlwings.XlwingsError: Couldn't connect` | Excel no encontrado por xlwings | En Mac: dar permisos en Privacy → Automation |
| `FileNotFoundError: AUTOMATIZACION_v7.xlsm` | Script no está en la raíz del proyecto | Ejecutar desde `BajataxV2/`, no desde subcarpeta |
| Módulo instalado pero sin cambios | El nombre del módulo en el .bas no coincide | Verificar que la primera línea del .bas tenga el nombre correcto |

---

## Verificar que la instalación funcionó

Después de ejecutar el script:

```
1. Abrir AUTOMATIZACION_v7.xlsm
2. Presionar Alt+F11 (o Tools → Macros → Visual Basic Editor en Mac)
3. En el panel izquierdo (Project Explorer) verificar que aparecen los módulos:
   - Mod_Sistema
   - Mod_ImportarArchivos
   - Mod_WhatsApp
   - Mod_PDF
   - etc.
4. Abrir uno y verificar que el código es el correcto
```

---

## Instalación manual (si el script falla)

Si `instalar_vba.py` no funciona, importar módulos manualmente:

```
En el editor VBA (Alt+F11):
1. Click derecho en el proyecto → "Import File..."
2. Navegar a src/vba-modules/
3. Seleccionar el .bas a importar
4. Repetir para cada módulo en orden numérico
```

---

## Flujo completo recomendado

```
Roo Code genera/modifica módulos .bas en src/vba-modules/
          ↓
git add . && git commit -m "feat: módulo X actualizado"
          ↓
Cerrar Excel si está abierto
          ↓
python src/python/instalar_vba.py
          ↓
Abrir Excel → verificar en editor VBA
          ↓
Probar la funcionalidad en el Excel
          ↓
Si funciona → git push
Si falla → git diff para ver qué cambió, reportar error en Roo Code
```
