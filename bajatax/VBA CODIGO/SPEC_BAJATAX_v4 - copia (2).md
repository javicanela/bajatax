# ESPECIFICACIÓN TÉCNICA COMPLETA — BajaTax v4_FINAL
> Documento de referencia para construir el proyecto desde cero.
> Fecha: 2026-03-03

---

## 1. CONTEXTO DEL PROYECTO

Sistema de automatización para el despacho fiscal **BajaTax**.
Construido en **Excel VBA (.xlsm)** para Mac. Sin servidores web.

**Función principal**: gestionar cobranza, enviar recordatorios por WhatsApp, generar estados de cuenta en PDF, registrar pagos y mantener un directorio de clientes.

**Responsables del despacho**: JOSSELYN, DENISSE, FERNANDA, OMAR, JAVIER

---

## 2. ARQUITECTURA

| Campo | Valor |
|---|---|
| Archivo base de instalación | `TEST_zip_rebuild.xlsm` (usa vbaProject.bin de v2, funcional) |
| Método de instalación | Importar `00_Bootstrap_Installer.bas` en VBE → ejecutar `InstalarSistema_BajaTax()` → se autodestruye |
| Hojas activas (9) | CONFIGURACION, OPERACIONES, REGISTROS, DIRECTORIO, BUSCADOR CLIENTE, LOG ENVÍOS, LOG ENVIOS, Soportes |
| Hoja eliminada | ~~REPORTES CXC~~ |
| Módulo eliminado | ~~09_Mod_ReportesCXC.bas~~ |
| KPIs | Eliminados completamente — no existen en ninguna forma |

---

## 3. MÓDULOS VBA — LISTA COMPLETA (11 archivos)

| # | Archivo | Nombre módulo VBA | Tipo |
|---|---|---|---|
| 0 | `00_Bootstrap_Installer.bas` | Bootstrap_Installer | Module (se autodestruye al terminar) |
| 1 | `01_Mod_Sistema.bas` | Mod_Sistema | Module |
| 2 | `02_Mod_ImportarArchivos.bas` | Mod_ImportarArchivos | Module |
| 3 | `03_WhatsApp.bas` | Mod_WhatsApp | Module |
| 4 | `04_PDF.bas` | Mod_PDF | Module |
| 5 | `05_Hoja_OPERACIONES.bas` | Hoja_OPERACIONES | Sheet module |
| 6 | `06_Hoja_DIRECTORIO.bas` | Hoja_DIRECTORIO | Sheet module |
| 7 | `07_Hoja_LOGENVIOS.bas` | Hoja_LOGENVIOS | Sheet module |
| 8 | `08_Mod_BuscadorCliente.bas` | Mod_BuscadorCliente | Module |
| 10 | `10_Hoja_BuscadorCliente.bas` | Hoja_BUSCADOR | Sheet module |
| 11 | `11_Hoja_REGISTROS.bas` | Hoja_REGISTROS | Sheet module (**NUEVO**) |

> **Nota**: `09_Mod_ReportesCXC.bas` fue eliminado. No debe incluirse.

---

## 4. CONFIGURACION (hoja sin cambios de lógica)

| Celda | Contenido |
|---|---|
| B2 | Modo: `PRUEBA` / `PRODUCCION` |
| B5 | Nombre despacho |
| B6 | Beneficiario |
| B7 | Banco |
| B8 | CLABE |
| B9 | Teléfono despacho |
| B10 | Correo |
| B11 | Web |
| B12 | Departamento |
| B14 | Número prueba (WA va aquí en modo PRUEBA) |
| B16 | Plantilla mensaje tipo 1 |
| B19 | Plantilla mensaje tipo 2 |
| B22 | Plantilla mensaje tipo 3 |

**Lógica PRUEBA/PRODUCCION**: En modo `PRUEBA` todos los envíos WA van al número de B14, no al cliente real.

---

## 5. SISTEMA DE BOTONES — COLUMNA Z

### Reglas generales

- Todos los botones de acción son **celdas en columna Z** de su respectiva hoja
- Activación exclusiva por **doble-clic** (`Worksheet_BeforeDoubleClick`)
- Si Z1 está ocupado, el siguiente botón va a Z2, Z3, etc.
- Se **auto-regeneran** cuando se borran: `Worksheet_Change` detecta celda Z vaciada y restaura el texto
- **No usar Form Controls ni ActiveX** en ninguna hoja

### Caracteres seguros (dentro del BMP, sin problemas en Mac Excel)

```vba
▶  = ChrW(&H25B6)   ' botones WA, IMPORTAR, PROCESAR
■  = ChrW(&H25A0)   ' botones PDF
✕  = ChrW(&H2715)   ' LIMPIAR
↺  = ChrW(&H21BA)   ' REGENERAR, INICIALIZAR
★  = ChrW(&H2605)   ' COLORIZAR (reemplaza emoji fuera de BMP)
```

### Botones Z por hoja

#### OPERACIONES
| Celda | Texto | Macro |
|---|---|---|
| Z1 | `▶ IMPORTAR` | `ImportarArchivosExternos` |
| Z2 | `▶ PROCESAR TODO` | `ProcesarTodoBajaTax` |
| Z3 | `▶ ENVÍO MASIVO WA` | `EnvioMasivoAutomatico` |
| Z4 | `■ PDF MASIVO` | `GenerarPDFMasivo` |
| Z5 | `↺ REGENERAR` | `RegenerarFaltantes` |

#### DIRECTORIO
| Celda | Texto | Macro |
|---|---|---|
| Z1 | `↺ INICIALIZAR` | `InicializarEncabezadosDirectorio` |
| Z2 | `★ COLORIZAR` | `ColorizarEstados` |

#### REGISTROS
| Celda | Texto | Macro |
|---|---|---|
| Z1 | `▶ IMPORTAR` | `ImportarArchivosExternos` |
| Z2 | `▶ PROCESAR TODO` | `ProcesarTodoBajaTax` |
| Z3 | `▶ ENVÍO MASIVO WA` | `EnvioMasivoAutomatico` |
| Z4 | `■ PDF MASIVO` | `GenerarPDFMasivo` |
| Z5 | `↺ REGENERAR` | `RegenerarFaltantes` |

#### BUSCADOR CLIENTE
| Celda | Texto | Macro |
|---|---|---|
| Z1 | `✕ LIMPIAR` | `LimpiarBuscador` |

---

## 6. HOJA: OPERACIONES

### Columnas (A–T + Z)

```
A  = Responsable        B  = ID_Factura      C  = Régimen
D  = Cliente            E  = RFC             F  = Fecha_Cob
G  = Concepto           H  = Monto           I  = Estatus      ← fórmula auto-generada
J  = Vencimiento        K  = Días_Venc       ← fórmula auto-generada
L  = Reg_Pago           M  = Teléfono        N  = Correo
O  = ▶WA (cell-btn)    ← auto-generado      P  = ■PDF (cell-btn) ← auto-generado
Q  = Excluir            R  = Prox_Envio      S  = Intentos      T  = Ult_Envio
Z  = Botones acción (Z1–Z5)
```

### Constantes VBA (en Mod_Sistema)

```vba
COL_OP_RESP    = 1   ' A - Responsable
COL_OP_ID      = 2   ' B - ID_Factura
COL_OP_REGIMEN = 3   ' C - Régimen
COL_OP_CLIENTE = 4   ' D - Cliente
COL_OP_RFC     = 5   ' E - RFC
COL_OP_FECHA   = 6   ' F - Fecha_Cob
COL_OP_CONCEP  = 7   ' G - Concepto
COL_OP_MONTO   = 8   ' H - Monto
COL_OP_ESTATUS = 9   ' I - Estatus
COL_OP_VENC    = 10  ' J - Vencimiento
COL_OP_DIAS    = 11  ' K - Días_Venc
COL_OP_PAGO    = 12  ' L - Reg_Pago
COL_OP_TEL     = 13  ' M - Teléfono
COL_OP_CORREO  = 14  ' N - Correo
COL_OP_WA      = 15  ' O - Botón WA
COL_OP_PDF     = 16  ' P - Botón PDF
COL_OP_EXCL    = 17  ' Q - Excluir
COL_OP_PROX    = 18  ' R - Prox_Envio
COL_OP_INT     = 19  ' S - Intentos
COL_OP_ULT     = 20  ' T - Ult_Envio
```

### Fórmulas auto-generadas (fila f)

```vba
' Col I — Estatus
"=IF(D" & f & "="""","""",IF(L" & f & "<>"""",""PAGADO""," & _
"IF(J" & f & "="""",""PENDIENTE"",IF(TODAY()>J" & f & _
",""VENCIDO"",IF(TODAY()=J" & f & ",""HOY VENCE"",""PENDIENTE"")))))"

' Col K — Días_Venc
"=IF(J" & f & "="""","""",J" & f & "-TODAY())"
```

### Celdas auto-generadas en OPERACIONES

| Celda | Tipo | Restaurar cuando... |
|---|---|---|
| `I{f}` fila≥2 | Fórmula estatus | Col I se vacía en fila con datos |
| `K{f}` fila≥2 | Fórmula días | Col K se vacía en fila con datos |
| `O{f}` fila≥2 | Texto `▶ WA` | Col O se vacía y D + H no están vacíos |
| `P{f}` fila≥2 | Texto `■ PDF` | Col P se vacía y D + H no están vacíos |
| Z1–Z5 | Texto botones | Celda Z1–Z5 se vacía |

### Triggers — Hoja OPERACIONES (05_Hoja_OPERACIONES.bas)

| Evento | Condición | Acción |
|---|---|---|
| `BeforeDoubleClick` | Col 15 (O), fila≥2, D y H no vacíos | `EnviarWhatsAppIndividual(fila)` |
| `BeforeDoubleClick` | Col 16 (P), fila≥2, D y H no vacíos | `GenerarPDFIndividual(fila)` |
| `BeforeDoubleClick` | Col 12 (L), fila≥2, D no vacío | Flujo registro de pago |
| `BeforeDoubleClick` | Col 26 (Z), filas 1–5 | Ejecutar macro botón Z |
| `Worksheet_Change` | Col 9 (I) vacía, fila≥2, D no vacío | Restaurar fórmula estatus |
| `Worksheet_Change` | Col 11 (K) vacía, fila≥2, J no vacío | Restaurar fórmula días |
| `Worksheet_Change` | Col 15 (O) vacía, fila≥2, D y H no vacíos | Restaurar `▶ WA` |
| `Worksheet_Change` | Col 16 (P) vacía, fila≥2, D y H no vacíos | Restaurar `■ PDF` |
| `Worksheet_Change` | Col 26 (Z) vacía, fila 1–5 | Restaurar texto botón Z |

### Flujo registro de pago (BeforeDoubleClick col L)

```
¿Col L ya tiene valor?
├── SÍ → MsgBox "Ya hay pago registrado ($X). ¿Cancelar pago?" [Sí / No]
│        Sí → Limpiar col L
│             col I recalcula automáticamente (fórmula detecta L="")
│             Quitar color verde de celda I + quitar color fila
│        No → No hacer nada
│
└── NO → InputBox "¿Monto recibido?" [texto / Cancel]
         Cancel o vacío → No hacer nada
         Monto válido →
           1. Escribir monto en col L
           2. col I muestra "PAGADO" automáticamente (fórmula: L<>"" → "PAGADO")
           3. Colorear fondo celda I: RGB(198, 239, 206) verde
           4. Colorear fila entera con color PAGADO
           5. Registrar en LOG ENVÍOS
```

> ⚠️ **REGLA CRÍTICA**: NO se escribe "PAGADO" directamente en col I.
> La fórmula `IF(L<>"","PAGADO",...)` maneja el valor automáticamente.
> Al cancelar pago: limpiar col L es suficiente — col I recalcula sola.
> Solo se gestiona el **color** de la celda I, nunca el valor.

---

## 7. HOJA: DIRECTORIO

### Columnas

```
A = RFC
B = Nombre / Razón Social
C = Email
D = Teléfono
E = Régimen
F = Responsable
G = Estatus
Z = Botones acción (Z1–Z2)
```

### Constantes VBA (en Mod_Sistema)

```vba
COL_DIR_RFC    = 1   ' A
COL_DIR_NOMBRE = 2   ' B
COL_DIR_EMAIL  = 3   ' C
COL_DIR_TEL    = 4   ' D
COL_DIR_REG    = 5   ' E
COL_DIR_RESP   = 6   ' F
COL_DIR_EST    = 7   ' G
```

### Triggers — Hoja DIRECTORIO (06_Hoja_DIRECTORIO.bas)

| Evento | Condición | Acción |
|---|---|---|
| `BeforeDoubleClick` | Col 26 (Z), fila 1 | `InicializarEncabezadosDirectorio` |
| `BeforeDoubleClick` | Col 26 (Z), fila 2 | `ColorizarEstados` |
| `Worksheet_Change` | Col 1 (A) borrada, fila≥2 | Limpiar toda la fila |
| `Worksheet_Change` | Col 26 (Z) vacía, fila 1–2 | Restaurar texto botón Z |

---

## 8. HOJA: REGISTROS (módulo NUEVO — 11_Hoja_REGISTROS.bas)

Vista consolidada de todos los datos. **Solo lectura por defecto**.
Edición requiere confirmación explícita por doble-clic.

### Columnas (A–N + Z)

```
A = Responsable     B = ID_Factura    C = Régimen     D = Cliente
E = RFC             F = Fecha_Cob     G = Concepto    H = Monto
I = Estatus         J = Vencimiento   K = Días_Venc   L = Reg_Pago
M = Teléfono        N = Correo
Z = Botones acción (Z1–Z5)
```

### Variables de módulo (privadas)

```vba
Private mEditando   As Boolean   ' True cuando se aprobó edición
Private mEditRow    As Long       ' Fila que se está editando
Private mEditCol    As Integer    ' Columna que se está editando
Private mEditOldVal As String     ' Valor original antes de la edición
```

### Triggers — Hoja REGISTROS (11_Hoja_REGISTROS.bas)

| Evento | Condición | Acción |
|---|---|---|
| `BeforeDoubleClick` | Col 26 (Z), filas 1–5 | Ejecutar macro Z correspondiente |
| `BeforeDoubleClick` | Col 1–14, fila≥2, celda **no vacía** | Diálogo ¿actualizar? → permitir/bloquear |
| `Worksheet_Change` | `mEditando = True`, nuevo ≠ "" | `SincronizarEdicionRegistros(...)` → `mEditando = False` |
| `Worksheet_Change` | `mEditando = True`, nuevo = "" | Dialog ¿restaurar desde OPERACIONES? → `mEditando = False` |
| `Worksheet_Change` | Col 26 (Z) vacía, fila 1–5 | Restaurar texto botón Z |

### Flujo de edición controlada

```
BeforeDoubleClick (col A–N, fila≥2, celda no vacía):
  → Cancel = True  (bloquea por defecto)
  → MsgBox "¿Deseas actualizar este dato del cliente?" [Sí / No]
  → No  → salir (celda no editable)
  → Sí  → mEditRow    = Target.Row
           mEditCol    = Target.Column
           mEditOldVal = CStr(Target.Value)
           mEditando   = True
           Cancel      = False   ← permite la edición

Worksheet_Change (si mEditando = True):
  → nuevoVal = Target.Value
  → mEditando = False
  → nuevoVal ≠ "" → Call SincronizarEdicionRegistros(mEditRow, mEditCol, nuevoVal, mEditOldVal)
  → nuevoVal = "" → MsgBox "Celda vaciada. ¿Restaurar valor desde OPERACIONES?" [Sí / No]
                     Sí → buscar por RFC (col E) + ID_Factura (col B) en OPERACIONES
                           restaurar valor original
                     No → dejar vacío (inconsistencia elegida por el usuario)
```

### Sub InicializarBotonesRegistros()

Sub pública que escribe/restaura los 5 botones Z en REGISTROS.
Llamada por `RegenerarFaltantes` y por Bootstrap.

---

## 9. FUNCIÓN: SincronizarEdicionRegistros (en Mod_ImportarArchivos)

```
Parámetros: editRow As Long, editCol As Integer, nuevoVal As String, oldVal As String

Clave compuesta de búsqueda:
  RFC       = ws_REGISTROS.Cells(editRow, 5)   ' col E
  ID_Fact   = ws_REGISTROS.Cells(editRow, 2)   ' col B

1. Buscar en OPERACIONES la fila donde (col E = RFC AND col B = ID_Fact)
2. Si encuentra: actualizar celda OPERACIONES(filaEnc, editCol) = nuevoVal

3. Si editCol corresponde a campo de DIRECTORIO, también actualizar DIRECTORIO
   (buscar por RFC = col A de DIRECTORIO):
```

### Mapeo REGISTROS → DIRECTORIO

| Col REGISTROS | Columna DIRECTORIO |
|---|---|
| D (col 4) — Nombre/Cliente | B (col 2) — Nombre |
| E (col 5) — RFC | A (col 1) — RFC |
| M (col 13) — Teléfono | D (col 4) — Teléfono |
| N (col 14) — Correo | C (col 3) — Email |
| C (col 3) — Régimen | E (col 5) — Régimen |

---

## 10. HOJA: BUSCADOR CLIENTE (rediseño completo)

### Layout fila 3 — Filtros

| Celda | Tipo | Valores |
|---|---|---|
| B3 | Dropdown (Data Validation) | TODOS / JOSSELYN / DENISSE / FERNANDA / OMAR / JAVIER |
| C3 | Dropdown (Data Validation) | TODOS + valores únicos de OPERACIONES col C (Régimen) |
| D3 | Dropdown (Data Validation) | TODOS / PENDIENTE / VENCIDO / HOY VENCE / PAGADO |
| E3 | Texto libre | Filtro Cliente (busca al salir de celda) |
| F3 | Texto libre | Filtro RFC (busca al salir de celda) |
| G3 | Texto libre | Filtro Concepto (busca al salir de celda) |
| H3 | Cell-btn activador | `BUSCAR ▶` — doble-clic ejecuta búsqueda |
| I3 | Dropdown (Data Validation) | Vencimiento / Monto / Cliente / Días |
| J3 | Dropdown (Data Validation) | Mayor a menor / Menor a mayor |

### Layout fila 6 — Encabezados resultados (fijos, no borrar)

```
A=No | B=Cliente | C=Responsable | D=RFC | E=Régimen | F=Concepto |
G=Monto | H=Vencimiento | I=Estatus | J=▶WA | K=■PDF
```

### Fila 7+ — Resultados dinámicos

- **Col A**: número secuencial (1, 2, 3…)
- **Cols B–I**: datos correspondientes de OPERACIONES
- **Col J**: texto `▶ WA` → un solo clic activa `EnviarWADesdeBuscador(fila)`
- **Col K**: texto `■ PDF` → un solo clic activa `GenerarPDFDesdeBuscador(fila)`

### Constantes (en Mod_Sistema)

```vba
BUSC_FILA_HEADERS = 6    ' Fila de encabezados de resultados
BUSC_FILA_DATOS   = 7    ' Primera fila de resultados
BUSC_COL_WA       = 10   ' J
BUSC_COL_PDF      = 11   ' K
```

### Celdas auto-generadas en BUSCADOR

| Celda | Valor restaurado |
|---|---|
| H3 | `"BUSCAR ▶"` |
| J7+ (con datos en B de esa fila) | `"▶ WA"` |
| K7+ (con datos en B de esa fila) | `"■ PDF"` |
| Z1 | `"✕ LIMPIAR"` |

### Triggers — Hoja BUSCADOR (10_Hoja_BuscadorCliente.bas)

| Evento | Condición | Acción |
|---|---|---|
| `BeforeDoubleClick` | H3 | `EjecutarBusqueda()` |
| `BeforeDoubleClick` | Z1 | `LimpiarBuscador()` |
| `SelectionChange` | Col 10 (J), fila≥7, valor = `▶ WA` | `EnviarWADesdeBuscador(Target.Row)` |
| `SelectionChange` | Col 11 (K), fila≥7, valor = `■ PDF` | `GenerarPDFDesdeBuscador(Target.Row)` |
| `SelectionChange` | Salir de E3 / F3 / G3 hacia otra celda | `EjecutarBusqueda()` |
| `Worksheet_Change` | H3 vacía | Restaurar `"BUSCAR ▶"` |
| `Worksheet_Change` | Col J vacía, fila≥7, B de esa fila ≠ "" | Restaurar `"▶ WA"` |
| `Worksheet_Change` | Col K vacía, fila≥7, B de esa fila ≠ "" | Restaurar `"■ PDF"` |
| `Worksheet_Change` | Z1 vacía | Restaurar `"✕ LIMPIAR"` |

### Detección de Enter en E3/F3/G3 (SelectionChange)

```vba
' Variable de módulo para rastrear celda anterior
Private mCeldaFiltro As String

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim eraCelda As String
    eraCelda = mCeldaFiltro
    mCeldaFiltro = ""

    ' Rastrear si entramos a celda de filtro texto
    If Target.Address = "$E$3" Or Target.Address = "$F$3" Or _
       Target.Address = "$G$3" Then
        mCeldaFiltro = Target.Address
    End If

    ' Si se salió de una celda de filtro → ejecutar búsqueda
    If eraCelda <> "" And Target.Address <> eraCelda Then
        Call EjecutarBusqueda
    End If

    ' Botón WA
    If Target.Column = BUSC_COL_WA And Target.Row >= BUSC_FILA_DATOS Then
        If Target.Value = ChrW(&H25B6) & " WA" Then
            Call EnviarWADesdeBuscador(Target.Row)
        End If
    End If

    ' Botón PDF
    If Target.Column = BUSC_COL_PDF And Target.Row >= BUSC_FILA_DATOS Then
        If Target.Value = ChrW(&H25A0) & " PDF" Then
            Call GenerarPDFDesdeBuscador(Target.Row)
        End If
    End If
End Sub
```

---

## 11. FUNCIONES PRINCIPALES POR MÓDULO

### 01_Mod_Sistema.bas

| Función / Sub | Descripción |
|---|---|
| `ProcesarTodoBajaTax()` | Recorre OPERACIONES fila a fila: actualiza días, envía WA automático si aplica según Prox_Envio |
| `RegenerarFaltantes()` | Recorre OPERACIONES y restaura todas las celdas auto-generadas faltantes (I, K, O, P, Z) |
| `InicializarBotonFila(ws, fila)` | Escribe `▶ WA` en col O y `■ PDF` en col P de la fila dada |
| `InicializarBotones(ws)` | Escribe todos los botones Z de la hoja pasada como parámetro |
| `ConstruirMensaje(tipo, fila)` | Construye texto WA para: VENCIDO / HOY_VENCE / RECORDATORIO usando plantillas de CONFIGURACION |
| `ObtenerConfig(celda)` | Retorna valor de celda en hoja CONFIGURACION |
| `EsModoProduccion()` | Retorna True si B2 = "PRODUCCION" |

### 02_Mod_ImportarArchivos.bas

| Función / Sub | Descripción |
|---|---|
| `ImportarArchivosExternos()` | Abre selector de archivo Excel → importa datos a hoja REGISTROS |
| `SincronizarEdicionRegistros(row, col, nuevoVal, oldVal)` | Propaga edición de REGISTROS → OPERACIONES → DIRECTORIO usando RFC + ID_Factura como clave |

### 03_WhatsApp.bas

| Función / Sub | Descripción |
|---|---|
| `EnviarWhatsAppIndividual(fila)` | Construye URL `wa.me/52{tel}?text=...` con mensaje UTF-8 encoded, abre en navegador |
| `EnvioMasivoAutomatico()` | Recorre OPERACIONES, filtra elegibles (Excluir≠"X", estatus activo, Prox_Envio≤hoy), envía con delay random 8–15s entre cada uno |

**Formato URL WhatsApp**:
```vba
"https://wa.me/52" & tel & "?text=" & EncodeURL(mensaje)
' Mensaje usa *negrita* de WhatsApp
' En modo PRUEBA: tel = ObtenerConfig("B14")
' En modo PRODUCCION: tel = col M de la fila
```

### 04_PDF.bas

| Función / Sub | Descripción |
|---|---|
| `GenerarPDFIndividual(fila)` | Crea hoja temporal `TEMP_BAJATAX`, llena estado de cuenta con datos de la fila, exporta como PDF, elimina hoja temporal |
| `GenerarPDFMasivo()` | **Silencioso** — genera PDF por cada cliente elegible sin diálogo SaveAs; path automático (carpeta del workbook + nombre cliente + fecha) |

### 08_Mod_BuscadorCliente.bas

| Función / Sub | Descripción |
|---|---|
| `EjecutarBusqueda()` | Lee filtros B3–J3, filtra filas de OPERACIONES, escribe resultados en filas 7+ con columnas A–K |
| `LimpiarBuscador()` | Limpia filtros B3–G3 (dropdowns a "TODOS", textos a ""), borra resultados fila 7+ |
| `PopularDropdownsRegimen()` | Lee valores únicos de OPERACIONES col C y los asigna como DV a celda C3 |
| `EnviarWADesdeBuscador(fila)` | Lee RFC de col D en fila resultado → busca fila origen en OPERACIONES → llama `EnviarWhatsAppIndividual` |
| `GenerarPDFDesdeBuscador(fila)` | Lee RFC de col D en fila resultado → busca fila origen en OPERACIONES → llama `GenerarPDFIndividual` |

---

## 12. BOOTSTRAP INSTALLER (00_Bootstrap_Installer.bas)

```vba
' Módulo de instalación automática.
' Uso:
'   1. Abrir TEST_zip_rebuild.xlsm
'   2. Herramientas → Macro → Editor de Visual Basic
'   3. Archivo → Importar → seleccionar 00_Bootstrap_Installer.bas
'   4. Ejecutar Sub InstalarSistema_BajaTax()
'   5. El módulo se autodestruye al terminar

Sub InstalarSistema_BajaTax()
    ' Verificar acceso a VBE (requiere "Confiar en VBA" habilitado)
    ' Ruta base: "/Users/javieravila/Documents/AUTOMATIZACION/VBA_CODIGO/"
    ' Importa módulos regulares con vbp.VBComponents.Import
    ' Actualiza sheet modules con CodeModule.DeleteLines + AddFromString
    ' Omite líneas "Attribute VB_" al actualizar sheet modules
    ' Al final: vbp.VBComponents.Remove vbp.VBComponents("Bootstrap_Installer")
End Sub
```

**Requisito previo**: En Excel → Preferencias → Seguridad → habilitar
"Confiar en el acceso al modelo de objetos de proyectos de VBA"

---

## 13. AUTO-REGENERACIÓN UNIVERSAL

Cada sheet module implementa `Worksheet_Change` que detecta cuando una celda auto-generada se vacía y la restaura automáticamente.

### Resumen de qué se auto-regenera por hoja

| Hoja | Celdas auto-regeneradas |
|---|---|
| OPERACIONES | I{f} (fórmula estatus), K{f} (fórmula días), O{f} (▶ WA), P{f} (■ PDF), Z1–Z5 (botones) |
| DIRECTORIO | Z1–Z2 (botones) |
| REGISTROS | Z1–Z5 (botones) |
| BUSCADOR | H3 (BUSCAR ▶), J7+ (▶ WA), K7+ (■ PDF), Z1 (✕ LIMPIAR) |

---

## 14. REGLAS DE NEGOCIO

### Estatus posibles (col I OPERACIONES)
| Valor | Condición |
|---|---|
| *(vacío)* | col D vacía |
| `PAGADO` | col L tiene valor (fórmula: `L<>""`) |
| `VENCIDO` | col J < hoy |
| `HOY VENCE` | col J = hoy |
| `PENDIENTE` | col J vacía o col J > hoy |

### Criterios para envío WA automático
- `Excluir` ≠ `"X"`
- Estatus en: VENCIDO, HOY VENCE, PENDIENTE
- `Prox_Envio` ≤ hoy o vacío
- Teléfono (col M) no vacío

### Anti-ban WA masivo
- Delay random entre 8 y 15 segundos entre cada envío
- `Application.Wait Now + TimeSerial(0, 0, Int((15-8+1)*Rnd+8))`

### PDF Masivo (silencioso)
- Path: carpeta del workbook & "\PDF_BajaTax\" & NombreCliente & "_" & Format(Date,"YYYY-MM-DD") & ".pdf"
- Sin `Application.FileDialog` — exporta directo con `ExportAsFixedFormat`

---

## 15. LOG ENVÍOS

Hoja LOG ENVÍOS (o LOG ENVIOS — existe duplicado legacy).
Columnas sugeridas: Fecha, Hora, Cliente, RFC, Teléfono, Tipo (WA/PDF), Estatus, Responsable.
Se escribe una fila por cada envío WA individual o registro de pago.

---

## 16. RESUMEN DE CAMBIOS VS V2/V3

| Elemento | V2/V3 | V4_FINAL |
|---|---|---|
| Botones acción | Form Controls | Celdas col Z, doble-clic |
| Botones WA/PDF por fila | Form Controls | Celdas col O/P, doble-clic (OPERACIONES) |
| Auto-regeneración | Solo botones O/P | Universal: I, K, O, P, Z en OPERACIONES; H3, J+, K+, Z en BUSCADOR; Z en DIRECTORIO y REGISTROS |
| Col I al pagar | Sin cambio | Manejado exclusivamente por fórmula; NO se escribe directo |
| Col I al cancelar | Sin cambio | Limpiar col L → fórmula recalcula sola |
| REPORTES CXC | Existía | **Eliminada** (hoja + módulo + KPIs) |
| KPIs | Existían | **Eliminados completamente** |
| `LimpiarRegistrosProcesados` | Existía | **Eliminada** |
| BUSCADOR CLIENTE | Layout antiguo | Rediseño completo (filtros B3–J3, resultados A–K, WA en J, PDF en K) |
| PDF Masivo | Con diálogo | **Silencioso**, path automático |
| REGISTROS edición | Sin control | Doble-clic + dialog + sync a OPERACIONES + DIRECTORIO |
| Módulo REGISTROS | No existía | `11_Hoja_REGISTROS.bas` (NUEVO) |
| Caracteres botones | `Chr(128241)` → `?` en Mac | `ChrW(&H25B6)` etc. — seguros en Mac Excel |

---

## 17. ESTRUCTURA DE ARCHIVOS VBA_CODIGO/

```
VBA_CODIGO/
├── SPEC_BAJATAX_v4.md              ← este documento
├── 00_Bootstrap_Installer.bas      ← instalador (se autodestruye)
├── 01_Mod_Sistema.bas
├── 02_Mod_ImportarArchivos.bas
├── 03_WhatsApp.bas
├── 04_PDF.bas
├── 05_Hoja_OPERACIONES.bas         ← reescrito
├── 06_Hoja_DIRECTORIO.bas          ← actualizado
├── 07_Hoja_LOGENVIOS.bas
├── 08_Mod_BuscadorCliente.bas      ← reescrito
├── 10_Hoja_BuscadorCliente.bas     ← reescrito
└── 11_Hoja_REGISTROS.bas           ← NUEVO
```

> `09_Mod_ReportesCXC.bas` fue eliminado del proyecto.

---

*Fin del documento — BajaTax v4_FINAL spec © 2026*
