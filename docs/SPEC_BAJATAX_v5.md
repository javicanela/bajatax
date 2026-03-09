# SPEC_BAJATAX_v5.md — Especificación Técnica BajaTax Automatización v7

> Versión 5, actualizada marzo 2026. Basada en INSTRUCCIONES.pdf (reglas de negocio) + TEST_zip_rebuild.xlsm (estructura real verificada). Esta versión corrige las divergencias entre la spec v4 y el código v7.

---

## 1. OBJETIVO DEL SISTEMA

Automatización integral de cobranza y gestión contable para un despacho fiscal. Operación basada en "pagos instantáneos" al generar la obligación. Diseñado para un solo usuario sin conocimientos técnicos, con capacidad de gestionar 500+ clientes activos.

**Stack tecnológico:**
- Excel con macros VBA (Mac como plataforma principal)
- Python 3 / ReportLab (generación de PDFs vía pdf_server.py)
- Evolution API / Docker (envío de WhatsApp vía REST API local)
- Ollama (modelos de IA locales para asistencia en desarrollo)

---

## 2. ARQUITECTURA DE HOJAS

### 2.1 Jerarquía Funcional

1. **CONFIGURACION** → Cerebro: parámetros globales, interruptor PRUEBA/PRODUCCIÓN
2. **REGISTROS** → Entrada: punto único de ingreso de datos externos
3. **DIRECTORIO** → Identidad: base maestra de clientes (RFC único)
4. **OPERACIONES** → Motor: gestión diaria, envíos WA, generación PDF
5. **BUSCADOR CLIENTE** → Consulta: búsqueda avanzada con filtros
6. **REPORTES CXC** → Análisis: resumen de cuentas por cobrar por asesor
7. **LOG ENVIOS** → Auditoría: bitácora de mensajes WhatsApp
8. **Soportes** → Auxiliar: lista de asesores para dropdowns

### 2.2 Flujo Principal

```
Archivos externos → REGISTROS → DIRECTORIO + OPERACIONES → PDF + WhatsApp → LOG ENVIOS
```

> Detalle completo de columnas por hoja en: DOCS/estructura-datos.md y .roo/rules/05-hojas-excel.md

---

## 3. MÓDULO DE CONFIGURACIÓN

### 3.1 Interruptor Maestro (B2)

| Valor | Comportamiento |
|-------|---------------|
| PRUEBA | Todos los WA → número B14. PDFs → OUTPUT/PRUEBA/. Ventana de advertencia. |
| PRODUCCIÓN | WA → números reales de clientes. PDFs → OUTPUT/DDMMYYYY/. Ventana de confirmación. |
| Vacío/otro | BLOQUEAR toda operación hasta que se configure. |

### 3.2 Protocolo de Validación

Antes de CUALQUIER envío (individual o masivo):
1. Leer B2
2. Mostrar ventana: "Modo activo: [valor]. ¿Desea continuar?"
3. Si usuario cancela → abortar sin cambios
4. Si PRUEBA → redirigir destino a B14
5. Si PRODUCCIÓN → usar números reales

---

## 4. MÓDULO DE IMPORTACIÓN INTELIGENTE

### 4.1 Principio

El sistema detecta datos por su CONTENIDO (patrones regex), no por posición ni nombre de columna. Funciona con archivos sin headers, headers en cualquier fila, columnas en cualquier orden, múltiples hojas, y datos mezclados con información irrelevante.

### 4.2 Detección por Contenido

| Dato | Patrón | Umbral |
|------|--------|--------|
| RFC | 3-4 letras + 6 dígitos + 2-3 alfanuméricos (12-13 chars) | >50% celdas |
| Email | Contiene @ y punto posterior | >50% |
| Teléfono | 10 dígitos con/sin prefijo (+52, 044, paréntesis) | >50% |
| Monto | Numérico con decimales, rango razonable | >50% |
| Fecha | Datetime o formato reconocible | >50% |
| Nombre | String >10 chars sin match a otros patrones | >50% |
| Régimen | Catálogo SAT o abreviaciones (PF, PM, AC) | >30% |

### 4.3 Diccionario de Aliases (complementario)

Cuando hay headers reconocibles, se comparan contra un diccionario de variaciones conocidas (ver agent.md sección 3.3). La normalización antes de comparar incluye: trim, lowercase, quitar acentos, quitar puntuación.

### 4.4 Estrategia

**Fase 1**: Escaneo automático → propuesta de mapeo con % de confianza.
**Fase 2**: Confirmación del usuario → corrección manual si necesario.

### 4.5 Destino: REGISTROS

Todos los datos importados se insertan en REGISTROS (cols A-K). Las columnas L-N son trackers internos que el sistema gestiona.

### 4.6 Deduplicación

Clave compuesta: **RFC + Concepto + Fecha Cobranza + Monto**. Si los 4 coinciden → preguntar al usuario. ID Factura es verificación adicional cuando existe, pero NO es requisito (no todos los clientes tienen factura).

### 4.7 Distribución

Al procesar un registro (doble clic en REGISTROS):
1. **DIRECTORIO**: Buscar RFC. Si no existe → crear. Si existe con datos diferentes → preguntar si actualizar.
2. **OPERACIONES**: Verificar duplicado con clave compuesta. Si no existe → insertar.
3. **Trackers**: Marcar L = "✓ OPERACIONES", M = "✓ DIRECTORIO", N = "✓ PROCESADO".

---

## 5. MÓDULO DE ESTATUS Y LÓGICA DE NEGOCIO

### 5.1 Algoritmo de Estatus (Col.I de OPERACIONES)

Evaluar en este orden estricto de prioridad:

| Prioridad | Condición | Resultado |
|-----------|-----------|-----------|
| 1 | Col.D (CLIENTE) vacío | Celda vacía |
| 2 | Col.L (Registro de Pago) tiene valor | PAGADO |
| 3 | Col.J (Fecha Vencimiento) vacía | PENDIENTE |
| 4 | @TODAY() = Col.J | HOY VENCE |
| 5 | @TODAY() > Col.J | VENCIDO |
| 6 | Cualquier otro caso | PENDIENTE |

### 5.2 Días Vencidos (Col.K)

```
Col.K = @TODAY() - Col.J
```
- Positivo = días de retraso
- Cero = vence hoy
- Negativo = días restantes hasta vencimiento

### 5.3 Activadores de Doble Clic

| Columna | Acción | Detalle |
|---------|--------|---------|
| L | Marcar PAGADO | Estampa Now(), cambia estatus automáticamente |
| O | Enviar WhatsApp | Dispara EnviarMensajeInteligente según estatus |
| P | Generar PDF | Consolida por RFC, llama a pdf_server.py |
| I (DIRECTORIO) | Alternar estado | ACTIVO ↔ SUSPENDIDO |

### 5.4 Columnas de Control (Q-T)

- **Q (EXCLUIR)**: Cualquier valor → procesos masivos saltan la fila. Uso: acuerdos especiales o datos pendientes de corrección.
- **R (FECHA_PROX_ENVIO)**: Siguiente contacto programado. Si estatus VENCIDO, calcular +3 o +7 días.
- **S (INTENTOS_ENVIO)**: Contador +1 por envío. Si ≥ 5 y sigue VENCIDO → celda roja, sugerir llamada humana.
- **T (ULTIMO_ENVIO_FECHA)**: Timestamp del último envío. Evita envíos duplicados el mismo día.

---

## 6. MÓDULO DE MENSAJERÍA WHATSAPP

### 6.1 Motor

**Principal**: Evolution API (Docker local, REST API en localhost:8080).
**Fallback**: URLs wa.me con mensaje encodeado.

### 6.2 Selección de Variante

| Días Vencidos (Col.K) | Variante | Tono | Color |
|------------------------|----------|------|-------|
| = 0 | MSG_HOY (prioridad máxima) | Último momento | Naranja |
| > 0 | MSG_VENCIDO | Urgente/firme | Rojo |
| < 0 | MSG_RECORDATORIO | Amigable/preventivo | Azul |
| Estatus PAGADO | BLOQUEAR | No enviar | — |

### 6.3 Variables Dinámicas

| Variable | Origen |
|----------|--------|
| {CLIENTE} | OPERACIONES Col.D |
| {MONTO} | OPERACIONES Col.H |
| {CONCEPTO} | OPERACIONES Col.G |
| {FECHA} | OPERACIONES Col.J |
| {DIAS} | OPERACIONES Col.K (absoluto) |
| {BENEFICIARIO} | CONFIGURACION B6 |
| {BANCO} | CONFIGURACION B7 |
| {CLABE} | CONFIGURACION B8 |
| {DEPTO} | CONFIGURACION B12 |
| {TEL_DESPACHO} | CONFIGURACION B9 |
| {EMAIL_DESPACHO} | CONFIGURACION B10 |

> Plantillas completas de las 4 variantes en agent.md sección 1.2.

### 6.4 Envío Masivo — Consolidación por Teléfono

1. Filtrar OPERACIONES: Estatus ≠ PAGADO y Col.Q vacía
2. Agrupar por Col.M (Teléfono)
3. Por cada número: sumar montos, concatenar conceptos
4. Construir UN mensaje consolidado por número
5. Confirmación: "Se detectaron X registros → Y envíos. ¿Proceder?"
6. Enviar con pausa 8-15s entre mensajes
7. Actualizar Col.S y Col.T de cada fila incluida

### 6.5 Codificación UTF-8

Acentos, ñ, caracteres especiales → códigos hex para URL. Negritas con asteriscos `*texto*` antes de encodear. Saltos de línea: `Chr(10)` → `%0A`.

> Tabla completa en agent.md sección 1.3.

### 6.6 Anti-Baneo

1. Pausa aleatoria 8-15 segundos entre envíos
2. Variaciones invisibles (espacios aleatorios al final)
3. Consolidar por teléfono (un mensaje por número)
4. Límite de intentos (≥5 → alerta visual)
5. Respetar Col.Q (EXCLUIR)
6. Respetar DIRECTORIO.I (SUSPENDIDO)

### 6.7 Post-Envío

- Col.S += 1
- Col.T = Now()
- Col.O → negritas
- LOG ENVIOS → nueva fila: fecha/hora, responsable, cliente, teléfono, variante, modo, resultado

---

## 7. MÓDULO PDF — ESTADOS DE CUENTA

### 7.1 Motor

Python 3 + ReportLab vía `pdf_server.py`. VBA construye JSON → ejecuta Python → Python genera PDF.

> NO usar ExportAsFixedFormat — falla en Mac sin impresora, no maneja acentos, formato depende del zoom.

### 7.2 Consolidación por RFC

Al generar PDF, buscar TODOS los registros del RFC en OPERACIONES y separar en:
- **Sección 1: Pendientes** (VENCIDO, HOY VENCE, PENDIENTE)
- **Sección 2: Liquidados** (PAGADO, con fecha/hora de pago desde Col.L)

### 7.3 Diseño Visual

**Formato**: Vertical (Portrait), tamaño Carta, optimizado para 1 página.
**Fuente**: Arial o Calibri (Sans Serif).

| Elemento | Especificación |
|----------|---------------|
| Encabezado pendientes | #1F4E78 (azul oscuro), texto blanco, negrita |
| Encabezado liquidados | #385623 (verde bosque), texto blanco, negrita |
| Sombreado filas pendientes | Zebra: blanco / gris muy claro |
| Sombreado filas liquidados | Zebra: blanco / verde muy claro |
| Estatus VENCIDO | Fondo rojo/rosa, letras rojas |
| Total pendiente | Rojo |
| Total liquidado | Verde |
| Bloque transferencia | Fondo azul acero, texto blanco |
| Moneda | $#,##0.00 MXN |
| Fechas | dd-mmm-yyyy |

### 7.4 Nomenclatura

- **Archivo**: `EdoCuenta_[NombreCliente]_[DDMMYYYY].pdf`
- **Carpeta**: `OUTPUT/DDMMYYYY/` (producción) o `OUTPUT/PRUEBA/` (modo prueba)
- **Sanitizar**: ñ→n, acentos→sin acento, espacios→_ en nombre de archivo
- **Peso**: < 500 KB

### 7.5 Generación Masiva

1. Filtrar por estatus (VENCIDO, HOY VENCE, PENDIENTE)
2. Excluir Col.Q y ESTADO_CLIENTE = SUSPENDIDO
3. Agrupar por RFC
4. Generar PDF por cada RFC único
5. Barra de progreso
6. Si un cliente falla → log error, continuar
7. Reporte final

**Sin pausas de seguridad** — proceso 100% local, no hay riesgo de baneo.

---

## 8. DIRECTORIO Y GESTIÓN DE CLIENTES

### 8.1 Unicidad

RFC es la clave única. Un RFC = un cliente en DIRECTORIO. Un cliente puede tener N registros en OPERACIONES.

### 8.2 Estado del Cliente (Col.I)

| Valor | Efecto |
|-------|--------|
| ACTIVO | Permite envíos WA y generación PDF normalmente |
| SUSPENDIDO | Bloquea WA y PDF para ese RFC. Fondo rojo tenue visual. |

### 8.3 Actualización de Datos

Si al procesar registros nuevos el RFC ya existe pero teléfono o correo difieren → ventana preguntando si actualizar en DIRECTORIO y propagar a OPERACIONES.

### 8.4 Clasificación (Col.G)

Campo de texto libre para etiquetar comportamiento del cliente: PAGADOR PUNTUAL, DEUDOR CRÓNICO, etc. Actualmente vacía — se llenará conforme se use el sistema.

---

## 9. SEGURIDAD Y AUDITORÍA

### 9.1 Modo PRUEBA vs PRODUCCIÓN

| Aspecto | PRUEBA | PRODUCCIÓN |
|---------|--------|------------|
| WA destino | B14 (número prueba) | Col.M (número cliente) |
| PDF carpeta | OUTPUT/PRUEBA/ | OUTPUT/DDMMYYYY/ |
| LOG registro | MODO = "PRUEBA" | MODO = "PRODUCCIÓN" |
| Ventana | "⚠ Modo PRUEBA activo" | "Modo PRODUCCIÓN — envío real" |

### 9.2 LOG ENVIOS

Cada envío de WhatsApp genera un registro con: fecha/hora, responsable, cliente, teléfono destino, variante de mensaje, modo del sistema, y resultado.

### 9.3 Protección de Datos

- CLABE (18 dígitos) con validación estricta de longitud
- Contraseñas, Fiel, Sellos → NUNCA importar al sistema
- Datos bancarios solo en CONFIGURACION, nunca hardcodeados en código

---

## 10. REGLAS TÉCNICAS VBA (Mac)

| Regla | Detalle |
|-------|---------|
| Fechas | @TODAY() en fórmulas, Date en VBA |
| Rutas | ThisWorkbook.Path & Application.PathSeparator |
| Eventos | EnableEvents = False/True al modificar celdas |
| Errores | On Error GoTo ErrorHandler en cada Sub |
| Workbooks | Cerrar en ErrorHandler si quedaron abiertos |
| Columnas | Buscar por header, nunca por índice |
| Módulos | Option Explicit obligatorio |
| Variables | camelCase inglés, comentarios español |

---

## 11. ROADMAP DE EVOLUCIÓN

| Fase | Alcance | Requisito previo |
|------|---------|-----------------|
| 1 | Importación inteligente | — |
| 2 | PDF masivo + Evolution API | Fase 1 completa |
| 3 | WhatsApp inteligente con variables | Fase 2 completa |
| 4 | Web App (React + FastAPI) | Javi valida Excel al 100% |
| 5 | SaaS multi-tenant | Web App validada |

**Principio**: Excel es el MVP. No se migra a web hasta validación completa en formato, funcionalidad y compatibilidad.
