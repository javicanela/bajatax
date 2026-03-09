# GUÍA DE INSTALACIÓN — AUTOMATIZACIÓN v3 BajaTax

## Resumen de Cambios v2 → v3

| Problema | Solución |
|----------|----------|
| Signos `?` en botones WhatsApp y PDF | Reemplazados por ▶ y ■ (Unicode BMP, compatible Mac/Win) |
| No existía importación de archivos externos | Nuevo módulo `Mod_ImportarArchivos` con auto-detección de columnas |
| PROCESAR REGISTROS no confirmaba duplicados | Ahora pregunta si concepto+monto ya existen en OPERACIONES |
| REGISTROS no alimentaba directamente OPERACIONES | Ahora escribe datos directamente + fórmulas de estatus |
| Faltaban encabezados en DIRECTORIO | Agregados: RFC, NOMBRE, EMAIL, TELÉFONO, RÉGIMEN |
| No había columna PROCESADO en REGISTROS | Nueva columna N marca estado de cada registro importado |
| No había envío masivo automático | Nueva función `EnvioMasivoAutomatico` con intervalos 8-20 seg |

---

## PASO 1: Abrir el archivo v3

1. Abre `AUTOMATIZACION_v3.xlsm` (las celdas ya están corregidas)
2. Si Excel pregunta sobre macros, haz clic en **Habilitar contenido**

---

## PASO 2: Abrir el Editor VBA

- **Windows:** `Alt + F11`
- **Mac:** `Option + F11` (o menú Herramientas > Macro > Editor de Visual Basic)

---

## PASO 3: Instalar los módulos (en este ORDEN)

### 3.1 — Módulo `Mod_Sistema` (PRIMERO)
1. En el panel izquierdo, busca el módulo `Mod_Sistema`
2. Si no existe: clic derecho en `VBAProject` → Insertar → Módulo → Renombrar a `Mod_Sistema`
3. Selecciona TODO el contenido existente (`Ctrl+A`) y bórralo
4. Abre el archivo `01_Mod_Sistema.bas` y copia TODO su contenido
5. Pega en el editor VBA

### 3.2 — Módulo `Mod_ImportarArchivos` (NUEVO)
1. Clic derecho en `VBAProject` → **Insertar → Módulo**
2. Renombra el nuevo módulo a `Mod_ImportarArchivos`
3. Copia y pega todo el contenido de `02_Mod_ImportarArchivos.bas`

### 3.3 — Módulo `WhatsApp`
1. Busca el módulo `WhatsApp` existente
2. Selecciona todo (`Ctrl+A`) y borra
3. Copia y pega todo el contenido de `03_Mod_WhatsApp.bas`

### 3.4 — Módulo `PDF`
1. Busca el módulo `PDF` existente
2. Selecciona todo (`Ctrl+A`) y borra
3. Copia y pega todo el contenido de `04_Mod_PDF.bas`

### 3.5 — Código de hoja OPERACIONES
1. En el panel izquierdo, busca en "Microsoft Excel Objetos" la hoja `OPERACIONES`
2. Haz **DOBLE CLIC** en ella (el nombre entre paréntesis)
3. Selecciona todo (`Ctrl+A`) y borra
4. Copia y pega todo el contenido de `05_Hoja_OPERACIONES.bas`
5. **IMPORTANTE:** Este código va EN LA HOJA, NO en un módulo estándar

### 3.6 — Código de hoja DIRECTORIO
1. Haz **DOBLE CLIC** en la hoja `DIRECTORIO` en el panel izquierdo
2. Selecciona todo y borra
3. Copia y pega todo el contenido de `06_Hoja_DIRECTORIO.bas`

---

## PASO 4: Asignar botones a macros

### Botón "IMPORTAR ARCHIVOS" (en hoja REGISTROS)
1. Ve a la hoja REGISTROS
2. Inserta un botón: menú **Desarrollador → Insertar → Botón (control de formulario)**
3. Dibuja el botón donde desees
4. Cuando pregunte la macro, selecciona: `ImportarArchivosExternos`
5. Cambia el texto del botón a: `IMPORTAR ARCHIVOS`

### Botón "PROCESAR REGISTROS" (en hoja REGISTROS)
1. Si ya existe, clic derecho → **Asignar macro**
2. Selecciona: `ProcesarTodoBajaTax`
3. Si no existe, créalo igual que el anterior

### Botón "LIMPIAR PROCESADOS" (opcional, en hoja REGISTROS)
1. Crea un botón y asígnale: `LimpiarRegistrosProcesados`
2. Texto: `LIMPIAR PROCESADOS`

### Botón "ENVÍO MASIVO" (opcional, en hoja OPERACIONES)
1. Crea un botón y asígnale: `EnvioMasivoAutomatico`
2. Texto: `ENVÍO MASIVO WA`

### Botón "REGENERAR FALTANTES" (opcional)
1. Asígnale: `RegenerarFaltantes`

---

## PASO 5: Crear carpeta IMPORTAR

Junto al archivo Excel, crea esta carpeta:
```
TU_CARPETA/
  ├── AUTOMATIZACION_v3.xlsm
  ├── IMPORTAR/          ← Aquí dejas los Excel externos
  ├── SALIDA_PDF/        ← Opcional: para PDFs generados
  └── LOGOS/             ← Opcional: logo del despacho
```

---

## FLUJO DE USO

### Importar datos nuevos:
1. Coloca los archivos Excel (.xlsx) en la carpeta `IMPORTAR`
2. Presiona **IMPORTAR ARCHIVOS** → selecciona los archivos
3. El sistema auto-detecta las columnas y las agrega a REGISTROS
4. Presiona **PROCESAR REGISTROS**:
   - Si concepto+monto YA existe en OPERACIONES → te pregunta
   - Si NO existe → agrega automáticamente
   - Los datos del cliente se envían a DIRECTORIO
   - Se marca como PROCESADO en columna N

### Enviar WhatsApp:
- Clic en el botón ▶ de la columna O en OPERACIONES
- O usa ENVÍO MASIVO para enviar a todos los pendientes

### Generar PDF:
- Clic en el botón ■ de la columna P en OPERACIONES

### Registrar pago:
- Doble clic en la columna L (Registro de Pago)

---

## NOTAS TÉCNICAS

- **Costo:** $0 — todo funciona con VBA nativo de Excel
- **Compatibilidad:** Windows y macOS
- **Símbolos Unicode:** ▶ ■ ✓ ✔ (rango BMP, sin emojis problemáticos)
- **Intervalos de envío masivo:** 8-20 segundos aleatorios entre mensajes
- **Modo PRUEBA:** Todos los mensajes van al número de CONFIGURACION B14
