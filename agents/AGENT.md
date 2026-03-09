# agent.md — BajaTax Especificación Extendida

> Este archivo complementa CLAUDE.md con las especificaciones detalladas de cada módulo.
> CLAUDE.md define QUÉ es el sistema. agent.md define CÓMO funciona cada parte.

---

## 1. MÓDULO DE MENSAJERÍA WHATSAPP

### 1.1 Mapeo de Variables

Cada `{VARIABLE}` se reemplaza dinámicamente antes de construir la URL:

| Variable | Origen | Celda |
|----------|--------|-------|
| {CLIENTE} | OPERACIONES | Col.D |
| {MONTO} | OPERACIONES | Col.H (formato $#,##0.00) |
| {CONCEPTO} | OPERACIONES | Col.G |
| {FECHA} | OPERACIONES | Col.J (formato dd-mmm-yyyy) |
| {DIAS} | OPERACIONES | Col.K (valor absoluto) |
| {BENEFICIARIO} | CONFIGURACION | B6 |
| {BANCO} | CONFIGURACION | B7 |
| {CLABE} | CONFIGURACION | B8 |
| {DEPTO} | CONFIGURACION | B12 |
| {TEL_DESPACHO} | CONFIGURACION | B9 |
| {EMAIL_DESPACHO} | CONFIGURACION | B10 |

### 1.2 Variantes de Mensaje

**VARIANTE 1: MSG_VENCIDO (Rojo) — Col.K > 0**

```
Baja Tax - Recordatorio de Pago Vencido

Estimado {CLIENTE},

Su cuenta presenta un saldo vencido de {MONTO} correspondiente a: {CONCEPTO}

Fecha de vencimiento: {FECHA} ({DIAS} días de retraso)

Le pedimos regularizar su situación a la brevedad para evitar la suspensión de servicios.

Apreciamos su pronto pago:

Datos para Transferencia:
Beneficiario: {BENEFICIARIO}
Banco: {BANCO}
CLABE: {CLABE}

Cualquier duda estamos a sus ordenes.
{DEPTO} | {TEL_DESPACHO}
{EMAIL_DESPACHO}
```

**VARIANTE 2: MSG_HOY (Naranja) — Col.K = 0 — PRIORIDAD MÁXIMA**

```
Baja Tax - Vencimiento Hoy

Estimado {CLIENTE},

Le recordamos que hoy {FECHA} es la fecha límite para realizar su pago.

Saldo pendiente: {MONTO}
Concepto: {CONCEPTO}

Evite recargos realizando su pago el día de hoy. Apreciamos su puntualidad:

Datos para Transferencia:
Beneficiario: {BENEFICIARIO}
Banco: {BANCO}
CLABE: {CLABE}

Cualquier duda estamos a sus ordenes.
{DEPTO} | {TEL_DESPACHO}
```

**VARIANTE 3: MSG_RECORDATORIO (Azul) — Col.K < 0**

```
Baja Tax - Próximo Vencimiento

Estimado {CLIENTE},

Le recordamos que el próximo {FECHA} es la fecha límite para realizar su pago.

Saldo pendiente: {MONTO}
Concepto: {CONCEPTO} ({DIAS} días restantes)

Agradecemos de antemano su gestión.
Datos para su depósito:
Beneficiario: {BENEFICIARIO}
Banco: {BANCO}
CLABE: {CLABE}

{DEPTO} | {TEL_DESPACHO}
```

**VARIANTE CONSOLIDADA (Envío Masivo — mismo teléfono, múltiples facturas)**

```
Baja Tax - Recordatorio de Saldo Pendiente

Estimado {CLIENTE},

Su cuenta presenta un saldo pendiente por la suma de {SUMA_TOTAL_MONTO} correspondiente a los siguientes conceptos:
• {CONCEPTO_1}
• {CONCEPTO_2}
• {CONCEPTO_N}

Le pedimos regularizar su situación a la brevedad.

Datos para Transferencia:
Banco: {BANCO} | CLABE: {CLABE}

{DEPTO} | {TEL_DESPACHO}
```

### 1.3 Codificación UTF-8 para WhatsApp (URLEncodeUTF8)

| Carácter | Código Hex |
|----------|-----------|
| á | %C3%A1 |
| é | %C3%A9 |
| í | %C3%AD |
| ó | %C3%B3 |
| ú | %C3%BA |
| ñ | %C3%B1 |
| Á | %C3%81 |
| É | %C3%89 |
| Í | %C3%8D |
| Ó | %C3%93 |
| Ú | %C3%9A |
| Ñ | %C3%91 |
| ü | %C3%BC |
| Ü | %C3%9C |
| espacio | %20 |
| salto línea | %0A |
| $ | %24 |
| , | %2C |
| : | %3A |
| ( | %28 |
| ) | %29 |

**Negritas en WhatsApp**: Rodear con asteriscos ANTES de encodear → `*{MONTO}*` → `*%242%2C100.00*`

**Saltos de línea en VBA**: Usar `Chr(10)` internamente, convertir a `%0A` en la URL final.

### 1.4 Protocolo Anti-Baneo

1. Pausa aleatoria entre envíos: `Application.Wait Now + TimeSerial(0, 0, Int(Rnd * 7) + 8)` → 8-15 segundos
2. Variaciones mínimas invisibles: espacios aleatorios al final del texto
3. Consolidar por teléfono: un mensaje por número, no por fila
4. Límite de intentos: si Col.S ≥ 5 y estatus sigue VENCIDO → pintar celda rojo, sugerir llamada humana
5. Respetar EXCLUIR (Col.Q): si tiene cualquier valor, saltar la fila en envío masivo

### 1.5 Post-Envío (actualización de celdas)

Después de cada envío exitoso:
- Col.S (INTENTOS_ENVIO) = Col.S + 1
- Col.T (ULTIMO_ENVIO_FECHA) = Now()
- Col.O: marcar en negritas para indicar visualmente que se despachó
- LOG ENVIOS: nueva fila con fecha/hora, responsable, cliente, teléfono, variante, modo, resultado

---

## 2. MÓDULO PDF — ESTADOS DE CUENTA

### 2.1 Motor: pdf_server.py

VBA recopila datos → genera JSON temporal → ejecuta Python → Python genera PDF → VBA verifica resultado.

```
Flujo técnico:
1. VBA identifica RFC de la fila activa
2. VBA busca TODOS los registros de ese RFC en OPERACIONES
3. VBA construye objeto JSON con: datos cliente, registros pendientes, registros pagados, datos bancarios
4. VBA escribe JSON a archivo temporal: TEMP/pdf_data_[RFC].json
5. VBA ejecuta: Shell "python3 " & ThisWorkbook.Path & "/python/pdf_server.py " & jsonPath
6. Python lee JSON, genera PDF con ReportLab
7. Python guarda en OUTPUT/DDMMYYYY/EdoCuenta_[Cliente]_[DDMMYYYY].pdf
8. VBA verifica que el archivo existe
9. VBA registra en LOG ENVIOS
10. VBA limpia archivo JSON temporal
```

### 2.2 Diseño Visual del PDF

**Orientación**: Vertical (Portrait), tamaño Carta, optimizado para 1 página.

**Encabezado**:
- Superior derecha: Logo (desde B16) + "Baja Tax" (azul oscuro) + "Dpto. de Impuestos" + contacto
- Superior izquierda: "ESTADO DE CUENTA" (negrita, azul oscuro, subrayado) + "Generado el [dd-mmm-yyyy]"
- Bloque ID: franja gris clara con CLIENTE: [NOMBRE] | RFC: [RFC] (negritas, mayúsculas)

**Sección 1: Conceptos Pendientes**

| Propiedad | Valor |
|-----------|-------|
| Color encabezado | Azul Oscuro #1F4E78 |
| Texto encabezado | Blanco, negrita |
| Columnas | No, Concepto, F.Cobro, Vencimiento, Monto, Estatus, Días Venc. |
| Sombreado filas | Zebra: blanco / gris muy claro |
| Estatus VENCIDO | Fondo rojo/rosa, letras rojas |
| Totalizador | "TOTAL PENDIENTE: $X,XXX.XX" en rojo |

**Sección 2: Historial Liquidados**

| Propiedad | Valor |
|-----------|-------|
| Color encabezado | Verde Bosque #385623 |
| Texto encabezado | Blanco, negrita |
| Columnas | No, Concepto, F.Cobro, Fecha Pago, Monto, Método |
| Sombreado filas | Zebra: blanco / verde muy claro |
| Registrado | Texto en verde |
| Totalizador | "TOTAL LIQUIDADO: $X,XXX.XX" en verde |

**Bloque de Pago (inferior)**:
- Encabezado "DATOS PARA TRANSFERENCIA" → fondo azul acero, texto blanco
- Beneficiario, Banco, CLABE alineados a la izquierda
- Despedida: "Cualquier duda estamos a sus órdenes." en cursiva

**Pie de página**:
- Izquierda: Baja Tax | Teléfono | Correo
- Centro: CLABE | Banco | Beneficiario
- Derecha: "Página 1 de 1"

### 2.3 Nomenclatura y Almacenamiento

- **Nombre archivo**: `EdoCuenta_[NombreCliente]_[DDMMYYYY].pdf`
- **Carpeta**: `OUTPUT/DDMMYYYY/` (se crea automáticamente si no existe)
- **Peso máximo**: optimizar para < 500 KB (envío rápido por WhatsApp)
- **Caracteres especiales**: sanitizar ñ, acentos en nombre de archivo (reemplazar por equivalente sin acento)

### 2.4 Generación Masiva

1. Filtrar: solo registros con Estatus = VENCIDO, HOY VENCE o PENDIENTE
2. Excluir: saltar filas con valor en Col.Q (EXCLUIR)
3. Excluir: saltar clientes con ESTADO_CLIENTE = "SUSPENDIDO" en DIRECTORIO
4. Agrupar por RFC: un PDF por cliente, no por fila
5. Barra de progreso: "15 de 48 completados"
6. Continuar si un cliente falla (no detener todo el proceso)
7. Reporte final: cuántos PDFs generados, cuántos fallaron y por qué

---

## 3. MÓDULO DE IMPORTACIÓN INTELIGENTE

### 3.1 Problema

Los archivos de origen llegan en CUALQUIER formato imaginable: headers en filas diferentes (o sin headers), nombres de columnas inconsistentes, datos irrelevantes mezclados, múltiples hojas con estructuras distintas, datos horizontales (meses como columnas), filas vacías intermedias, títulos decorativos, celdas merged, y formatos mixtos dentro de la misma columna.

### 3.2 Detección por Contenido (MÉTODO PRIMARIO)

El sistema identifica columnas por el PATRÓN de sus datos, no por el nombre del header. Esto funciona incluso sin headers:

```
RFC → Regex: 3-4 letras + 6 dígitos + 2-3 alfanuméricos = 12-13 chars
       Si >50% de celdas en una columna matchean → es la columna RFC
       Ejemplo: "AALR930211MX2", "CAT931006E16"

EMAIL → Contiene @ y al menos un punto después
        Regex: .*@.*\..*
        Ejemplo: "cliente@gmail.com", "dr.pgeffroy@gmail.com"

TELÉFONO → 10 dígitos consecutivos (con o sin prefijo)
            Detectar: (664) 123-4567, 664.123.4567, 6641234567, +526641234567
            También: prefijos 044, 045, +52, +1

MONTOS → Números con formato $X,XXX.XX o decimales
          Columna donde >50% son numéricos y rango razonable (100-500,000)
          Ejemplo: 2100, $2,100.00, 1600.00

FECHAS → Detectar múltiples formatos:
          dd/mm/yyyy, mm/dd/yyyy, yyyy-mm-dd
          "01-ene-2026", datetime objects de Excel
          Si una columna tiene >50% fechas válidas → es columna de fecha

NOMBRES → Strings largos (>10 chars) que no matchean ningún otro patrón
           Texto mixto mayúsculas/minúsculas, promedio >15 chars
           Ejemplo: "Colegio de Anestesiólogos de Tijuana Dr Benigno M"

RÉGIMEN → Coincide con catálogo SAT (601, 605, 606, 612, 625, 626, 616)
           O abreviaciones: PF, PM, AC, AC FNL, RESICO
```

### 3.3 Diccionario de Aliases (MÉTODO COMPLEMENTARIO — solo cuando hay headers)

Cuando el archivo SÍ tiene headers reconocibles, estos aliases ayudan al mapeo. Son un complemento a la detección por contenido, NO el método principal:

```
REGISTROS.B (NOMBRE DEL CONTRIBUYENTE):
  → "nombre del contribuyente"
  → "cliente"
  → "razon social"
  → "razón social"
  → "contribuyente"
  → "nombre"
  → "nombre completo"

REGISTROS.C (RFC):
  → "rfc"
  → "r.f.c."
  → "rfc del contribuyente"
  → "registro federal"

REGISTROS.D (EMAIL):
  → "email"
  → "e-mail"
  → "correo"
  → "correo electrónico"
  → "e-mail [1]"
  → "email [1]"
  → "mail"

REGISTROS.E (TELÉFONO):
  → "teléfono"
  → "telefono"
  → "tel"
  → "tel [1]"
  → "celular"
  → "número"
  → "numero"
  → "whatsapp"

REGISTROS.G (CONCEPTO):
  → "concepto"
  → "descripción"
  → "descripcion"
  → "servicio"
  → "comentarios"

REGISTROS.H (MONTO):
  → "monto"
  → "monto base"
  → "total"
  → "importe"
  → "honorarios"
  → "cantidad"
  → "envio de honorarios"

REGISTROS.I (FACTURA):
  → "factura"
  → "id factura"
  → "folio"
  → "folio de factura"
  → "numero de factura"
  → "factura o excel"

REGISTROS.J (REGIMEN):
  → "regimen"
  → "régimen"
  → "regimen fiscal"
  → "régimen fiscal"
  → "tipo régimen"

REGISTROS.K (VENCIMIENTO):
  → "vencimiento"
  → "fecha vencimiento"
  → "fecha de vencimiento"
  → "fecha límite"
  → "fecha limite"
  → "vigencia impuestos"
```

### 3.4 Normalización antes de comparar headers

```
1. Trim (quitar espacios al inicio/final)
2. LCase (convertir a minúsculas)
3. Quitar acentos: á→a, é→e, í→i, ó→o, ú→u, ñ→n
4. Quitar caracteres especiales: puntos, guiones, corchetes
5. Quitar palabras vacías: "de", "del", "la", "el"
```

### 3.5 Estrategia de Detección en Dos Fases

**Fase 1 — Escaneo Automático**: El sistema analiza el archivo completo y PROPONE un mapeo:
1. Escanear CADA hoja del workbook (no solo la primera)
2. Para cada hoja, escanear las primeras 20 filas
3. Detectar headers: buscar la fila donde hay más texto tipo-etiqueta y menos datos tipo-valor
4. Si no hay headers → clasificar columnas por patrones de contenido
5. Generar propuesta: "Columna C parece RFC (85% coincidencia), Columna E parece email (92%)"

**Fase 2 — Confirmación del Usuario**: Mostrar ventana con el mapeo propuesto. El usuario confirma o corrige antes de importar.

**Orden de prioridad para identificar columnas:**
1. Detección por contenido (patrones regex) — PRIMERO
2. Diccionario de aliases (si hay headers) — SEGUNDO
3. Posición heurística (último recurso) — TERCERO
4. Preguntar al usuario (si confianza < 60%) — SIEMPRE disponible

### 3.6 Situaciones que debe manejar

1. **Sin headers**: Escanear contenido, clasificar por patrón
2. **Headers en cualquier fila** (1, 2, 3, 4, 5...): Detectar automáticamente
3. **Múltiples hojas**: Analizar cada una, reportar qué encontró
4. **Columnas revueltas**: Orden no importa — detectar por contenido
5. **Datos horizontales**: Meses como columnas → detectar y transponer si necesario
6. **Basura mezclada**: Filas vacías, títulos decorativos, notas al pie, celdas merged → filtrar
7. **Formatos mixtos**: Teléfonos con/sin formato, RFC mayúsculas/minúsculas → normalizar
8. **Archivos .csv y .xlsx**: Ambos soportados
9. **Encoding**: UTF-8, Latin-1, Windows-1252 → detectar automáticamente

### 3.7 Columnas Obligatorias vs Opcionales

**Obligatorias** (la importación falla sin estas):
- RFC o NOMBRE DEL CONTRIBUYENTE (al menos uno)

**Opcionales** (se importan si existen, vacías si no):
- EMAIL, TELÉFONO, CONCEPTO, MONTO, FACTURA, REGIMEN, VENCIMIENTO

**Ignorar POR AHORA** (basado en archivos analizados hasta la fecha):
- Contraseña, Fiel, Sellos, Vigencia, DIOT, Cedula, Rep. Legal
- Esta lista puede crecer conforme se analicen más tipos de archivos
- Si se detecta una columna desconocida → preguntar al usuario qué hacer con ella

### 3.8 Protocolo de Distribución (REGISTROS → OPERACIONES + DIRECTORIO)

Al procesar un registro en REGISTROS (doble clic):

**Paso 1: DIRECTORIO**
- Buscar RFC en DIRECTORIO.Col.A
- Si NO existe: crear fila nueva con RFC, Cliente, Correo, Teléfono, Régimen, Responsable, fecha alta = Now(), estado = "ACTIVO"
- Si EXISTE pero teléfono o correo difiere: ventana → "¿Actualizar datos de contacto?"
- Marcar REGISTROS.Col.M = "✓ DIRECTORIO"

**Paso 2: OPERACIONES**
- Verificar duplicados con **clave compuesta**: RFC (Col.E) + Concepto (Col.G) + Fecha Cobranza (Col.F) + Monto (Col.H)
- Si los 4 campos coinciden con un registro existente → ventana: "¿Re-procesar o solo nuevos?"
- Si el registro tiene ID Factura, usarlo como verificación adicional pero NO como requisito (no todos los clientes tienen factura)
- Si no hay coincidencia → insertar nueva fila mapeando:
  - OPERACIONES.A = REGISTROS.A (Responsable)
  - OPERACIONES.B = REGISTROS.I (Factura → ID Factura)
  - OPERACIONES.C = REGISTROS.J (Régimen)
  - OPERACIONES.D = REGISTROS.B (Nombre → Cliente)
  - OPERACIONES.E = REGISTROS.C (RFC)
  - OPERACIONES.F = REGISTROS.F (Fecha → Fecha Cobranza)
  - OPERACIONES.G = REGISTROS.G (Concepto)
  - OPERACIONES.H = REGISTROS.H (Monto → Monto Base)
  - OPERACIONES.J = REGISTROS.K (Vencimiento → Fecha Vencimiento)
  - OPERACIONES.M = REGISTROS.E (Teléfono)
  - OPERACIONES.N = REGISTROS.D (Email → Correo)
- Marcar REGISTROS.Col.L = "✓ OPERACIONES"

**Paso 3: Confirmación**
- Marcar REGISTROS.Col.N = "✓ PROCESADO"
- Celda final de la fila → fondo verde
- Ventana resumen: "Registros procesados: X nuevos, Y actualizados, Z duplicados saltados"

---

## 4. REGÍMENES FISCALES SAT

| Código | Nombre Técnico | Explicación |
|--------|---------------|-------------|
| 601 | General de Ley Personas Morales | Empresas y sociedades con fines de lucro |
| 605 | Sueldos y Salarios e Ingresos Asimilados | Empleados por nómina |
| 606 | Arrendamiento | Renta de inmuebles |
| 612 | Personas Físicas con Actividades Empresariales | Freelancers, médicos, abogados |
| 625 | Actividades por Plataformas Tecnológicas | Amazon, Uber, Airbnb |
| 626 | Régimen Simplificado de Confianza (RESICO) | Pequeños negocios |
| 616 | Sin obligaciones fiscales | Estudiantes, sin ingresos |

**Abreviaciones usadas en el sistema**:
- PF = Persona Física (RFC 13 chars) → Regímenes 605, 606, 612, 625, 626, 616
- PM = Persona Moral (RFC 12 chars) → Régimen 601
- AC = Asociación Civil → variante de PM
- AC FNL = Asociación Civil sin Fines de Lucro

---

## 5. SEGURIDAD: MODO PRUEBA vs PRODUCCIÓN

### Protocolo de Validación

Antes de CUALQUIER envío (individual o masivo):

```
1. Leer B2 de CONFIGURACION
2. Mostrar ventana: "Modo activo: [PRUEBA/PRODUCCIÓN]. ¿Desea continuar?"
3. Si PRUEBA:
   - Todos los mensajes WA → número B14
   - PDFs se generan con datos reales pero en carpeta OUTPUT/PRUEBA/
   - LOG ENVIOS registra MODO = "PRUEBA"
4. Si PRODUCCIÓN:
   - Mensajes WA → número real del cliente (Col.M de OPERACIONES)
   - PDFs en carpeta OUTPUT/DDMMYYYY/
   - LOG ENVIOS registra MODO = "PRODUCCIÓN"
5. Si B2 está vacío o tiene otro valor → BLOQUEAR y pedir configurar
```

### Protección adicional: DIRECTORIO.ESTADO_CLIENTE

Si ESTADO_CLIENTE (Col.I) = "SUSPENDIDO":
- Bloquear envío de WhatsApp para ese RFC
- Bloquear generación de PDF para ese RFC
- El registro aparece en OPERACIONES pero con indicador visual (fondo rojo tenue)

---

## 6. HOJAS AUXILIARES

### BUSCADOR CLIENTE
- Fila 2-3: Filtros (Responsable, Estatus, Concepto, Rango fechas vencimiento, Días vencidos)
- Fila 3: Botones BUSCAR ▶ y ✕ LIMPIAR
- Fila 5: Contadores (Registros encontrados, Total monto, Total vencido)
- Fila 6+: Resultados con columnas: No, Cliente, Responsable, Concepto, Monto, Estatus, Vencimiento, Días, Teléfono, WhatsApp, PDF

### REPORTES CXC (Cuentas por Cobrar)
- Resumen por responsable: Cobrado ($), Pendiente ($), Vencido, Total, Flujo (%)
- Responsables: Joselyn, Denisse, Fernanda, Omar, Javier

### Soportes
- Lista simple de asesores en Col.A
- Alimenta las listas desplegables de "Responsable" en otras hojas

---

## 7. DATOS BANCARIOS DEL DESPACHO (referencia)

Los siguientes son valores reales configurados en CONFIGURACION:
- **Despacho**: Baja Tax
- **Beneficiario**: Oscar Abel Zambrano Rentería
- **Banco**: Banco Inbursa S.A.
- **CLABE**: 036 028 5005 7775 7743
- **Departamento**: Dpto. de Impuestos

> Estos valores se leen dinámicamente de CONFIGURACION, NUNCA hardcodear en el código.
