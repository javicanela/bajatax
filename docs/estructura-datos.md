# estructura-datos.md — Mapa de Estructura de Datos BajaTax v7

> Documento de referencia técnica. Describe la estructura real del archivo AUTOMATIZACION_v7.xlsm verificada contra el Excel de producción.

---

## 1. Arquitectura General

El sistema se organiza en 8 hojas de trabajo + 1 hoja auxiliar:

```
┌─────────────────────────────────────────────────────┐
│                    CONFIGURACION                     │
│         Panel de control — parámetros globales       │
│              B2: Interruptor PRUEBA/PROD             │
└──────────────────────┬──────────────────────────────┘
                       │ alimenta mensajes y PDFs
                       ▼
┌──────────────────────────────────────────────────────┐
│                     REGISTROS                         │
│      Punto de entrada — datos externos caen aquí      │
│           A-K: datos | L-N: trackers                  │
└────────┬─────────────────────────┬───────────────────┘
         │ distribuye              │ distribuye
         ▼                        ▼
┌────────────────────┐   ┌────────────────────┐
│    OPERACIONES     │   │    DIRECTORIO      │
│  Motor operativo   │   │  Base maestra      │
│  A-T (20 cols)     │   │  A-I (9 cols)      │
│  Envíos WA + PDF   │   │  RFC único         │
└────────┬───────────┘   └────────────────────┘
         │ genera
         ▼
┌────────────────────┐   ┌────────────────────┐
│    LOG ENVIOS      │   │   REPORTES CXC     │
│  Bitácora WA       │   │  Resumen x asesor  │
│  A-G               │   │  A-F               │
└────────────────────┘   └────────────────────┘

┌────────────────────┐   ┌────────────────────┐
│  BUSCADOR CLIENTE  │   │     Soportes       │
│  Búsqueda avanzada │   │  Lista asesores    │
│  Filtros + results │   │  (auxiliar)        │
└────────────────────┘   └────────────────────┘
```

---

## 2. Detalle por Hoja

### 2.1 CONFIGURACION

Tipo: formulario vertical en columna B. No tiene headers tradicionales.

```
Fila 1: A="MODO DEL SISTEMA"     B=(etiqueta)
Fila 2: A=                        B="PRUEBA" o "PRODUCCIÓN"  ← INTERRUPTOR MAESTRO
Fila 4: A="DATOS DEL DESPACHO"   B=(etiqueta)
Fila 5: A=                        B="Baja Tax"               ← Nombre despacho
Fila 6: A=                        B="Oscar Abel Zambrano..."  ← Beneficiario
Fila 7: A=                        B="Banco Inbursa S.A."     ← Banco
Fila 8: A=                        B="036 028 5005 7775 7743" ← CLABE (18 dígitos)
Fila 9: A=                        B="+52 (XXX) XXX XXXX"    ← Teléfono despacho
Fila 10: A=                       B="correo@bajatax.com"     ← Email
Fila 12: A=                       B="Dpto. de Impuestos"     ← Departamento
Fila 14: A=                       B="52XXXXXXXXXX"           ← Número prueba WA
Fila 16: A=                       B="/ruta/al/logo.png"      ← Logo para PDFs
```

Columna C: descripciones de ayuda para el usuario.
Columnas D-E: opciones de la lista desplegable de B2 ("PRODUCCIÓN", "PRUEBA").

### 2.2 REGISTROS (A-N)

Punto de entrada de datos externos. Headers en fila 1.

```
DATOS IMPORTADOS (A-K):
  A: RESPONSABLE         → Asesor asignado
  B: NOMBRE DEL CONTRIBUYENTE → Razón social
  C: RFC                 → 12-13 caracteres
  D: EMAIL               → Correo del cliente
  E: TELÉFONO            → 10 dígitos
  F: FECHA               → Fecha cobranza/emisión
  G: CONCEPTO            → Descripción servicio
  H: MONTO               → Cantidad a cobrar
  I: FACTURA             → ID factura (OPCIONAL)
  J: REGIMEN             → Código SAT
  K: VENCIMIENTO         → Fecha límite pago

TRACKERS DE DISTRIBUCIÓN (L-N):
  L: "✓ OPERACIONES"     → Se marcó cuando el registro se copió a OPERACIONES
  M: "✓ DIRECTORIO"      → Se marcó cuando el cliente se registró/actualizó en DIRECTORIO
  N: "✓ PROCESADO"       → Se marcó cuando ambas distribuciones completaron
```

### 2.3 OPERACIONES (A-T)

Motor operativo diario. Headers en fila 1. 20 columnas.

```
IDENTIFICACIÓN (A-E):
  A: RESPONSABLE         → Asesor
  B: ID FACTURA          → Folio (puede estar vacío)
  C: REGIMEN             → Código SAT
  D: CLIENTE             → Razón social
  E: RFC                 → Clave de vinculación con DIRECTORIO

TRANSACCIÓN (F-H):
  F: FECHA COBRANZA      → Emisión obligación
  G: Concepto            → Descripción servicio
  H: Monto Base          → Cantidad MXN

ESTADO (I-K) — FÓRMULAS AUTOMÁTICAS:
  I: Estatus de Pago     → PAGADO|HOY VENCE|VENCIDO|PENDIENTE
  J: Fecha Vencimiento   → Límite de pago
  K: Días Vencidos       → @TODAY() - Col.J

ACTIVADORES DOBLE CLIC (L, O, P):
  L: Registro de Pago    → ⬡ Doble clic estampa Now(), marca PAGADO
  O: Acción WhatsApp     → ⬡ Doble clic dispara envío WA
  P: Estado De Cuenta    → ⬡ Doble clic genera PDF consolidado por RFC

CONTACTO (M-N):
  M: Teléfono            → 10 dígitos
  N: Correo              → Email

CONTROL AVANZADO (Q-T):
  Q: EXCLUIR             → Si tiene valor, masivos saltan esta fila
  R: FECHA_PROX_ENVIO    → Próximo contacto programado
  S: INTENTOS_ENVIO      → Contador +1 por envío. Rojo si ≥5
  T: ULTIMO_ENVIO_FECHA  → Stamp último envío exitoso
```

### 2.4 DIRECTORIO (A-I)

Base maestra de clientes. Headers en fila 1. RFC es ÚNICO.

```
  A: RFC                 → CLAVE ÚNICA (12 morales, 13 físicas)
  B: CLIENTE             → Razón social
  C: CORREO              → Email principal
  D: TELEFONO            → 10 dígitos
  E: REGIMEN             → Código SAT
  F: RESPONSIBLE         → ⚠ Typo en header real (debería ser RESPONSABLE)
  G: CLASIFICACIÓN       → PAGADOR PUNTUAL / DEUDOR CRÓNICO / etc. (vacía aún)
  H: (sin header)        → FECHA ALTA — registro de ingreso
  I: (sin header)        → ESTADO_CLIENTE: ACTIVO / SUSPENDIDO
```

⚠ Cols H e I no tienen header en el Excel real pero contienen datos funcionales.

### 2.5 LOG ENVIOS (A-G)

Bitácora de mensajes WhatsApp. Headers en fila 1.

```
  A: FECHA/HORA          → Timestamp del envío
  B: RESPONSABLE         → Quién ejecutó
  C: CLIENTE             → Destinatario
  D: TELEFONO            → Número destino (prueba o real)
  E: VARIANTE            → VENCIDO / HOY VENCE / RECORDATORIO
  F: MODO                → PRUEBA / PRODUCCIÓN
  G: RESULTADO           → ENVIADO / CANCELADO / ERROR
```

⚠ Existen DOS hojas de log: "LOG ENVÍOS" (con acento) y "LOG ENVIOS" (sin acento). Usar "LOG ENVIOS" (sin acento) como principal.

### 2.6 Hojas Auxiliares

**BUSCADOR CLIENTE**: Interfaz de búsqueda con filtros en filas 2-3 (Responsable, Estatus, Concepto, Fechas, Días vencidos), botones en fila 3, contadores en fila 5, resultados desde fila 6.

**REPORTES CXC**: Resumen por responsable con columnas Cobrado, Pendiente, Vencido, Total, Flujo (%). Responsables: Joselyn, Denisse, Fernanda, Omar, Javier.

**Soportes**: Col A = lista de asesores (JOSSELYN, DENISSE, FERNANDA, OMAR, JAVIER). Alimenta dropdowns de Responsable.

---

## 3. Relaciones y Vínculos

### 3.1 Claves de Vinculación

| Relación | Hoja A | Campo A | Hoja B | Campo B | Tipo |
|----------|--------|---------|--------|---------|------|
| Identidad cliente | REGISTROS | C (RFC) | DIRECTORIO | A (RFC) | 1:1 |
| Transacciones | REGISTROS | C (RFC) | OPERACIONES | E (RFC) | 1:N |
| Cliente↔Movimientos | DIRECTORIO | A (RFC) | OPERACIONES | E (RFC) | 1:N |
| Control envío | DIRECTORIO | I (ESTADO_CLIENTE) | OPERACIONES | — | Bloqueo |
| Modo sistema | CONFIGURACION | B2 | OPERACIONES | — | Global |
| Bitácora | OPERACIONES | (envío) | LOG ENVIOS | (nueva fila) | 1:N |
| Asesores | Soportes | A | Todas las hojas | Responsable | Dropdown |

### 3.2 Flujo de Datos al Importar

```
Archivo externo
    ↓ (detección por contenido + aliases)
REGISTROS fila nueva (cols A-K)
    ↓ (doble clic → procesar)
    ├→ DIRECTORIO: buscar RFC
    │   Si no existe → crear entidad nueva
    │   Si existe + datos diferentes → preguntar actualizar
    │   Marcar REGISTROS.M = "✓ DIRECTORIO"
    │
    └→ OPERACIONES: verificar duplicado (RFC+Concepto+Fecha+Monto)
        Si existe → preguntar re-procesar
        Si no existe → insertar fila nueva
        Marcar REGISTROS.L = "✓ OPERACIONES"

    → Marcar REGISTROS.N = "✓ PROCESADO"
```

### 3.3 Flujo de Envío WhatsApp

```
Usuario doble clic Col.O de OPERACIONES
    ↓
Validar B2 CONFIGURACION (PRUEBA/PRODUCCIÓN)
    ↓
Verificar DIRECTORIO.I ≠ "SUSPENDIDO" para ese RFC
    ↓
Verificar OPERACIONES.I ≠ "PAGADO"
    ↓
Seleccionar variante por Col.K (0=HOY, >0=VENCIDO, <0=RECORDATORIO)
    ↓
Construir mensaje con variables
    ↓
Enviar (Evolution API o wa.me)
    ↓
Actualizar Col.S (+1), Col.T (Now()), Col.O (negritas)
    ↓
Registrar en LOG ENVIOS
```

### 3.4 Flujo de Generación PDF

```
Usuario doble clic Col.P de OPERACIONES
    ↓
Tomar RFC de la fila activa
    ↓
Buscar TODOS los registros con ese RFC en OPERACIONES
    ↓
Separar: pendientes (VENCIDO/HOY VENCE/PENDIENTE) vs liquidados (PAGADO)
    ↓
Construir JSON con datos cliente + pendientes + liquidados + datos bancarios
    ↓
Ejecutar pdf_server.py
    ↓
PDF → OUTPUT/DDMMYYYY/EdoCuenta_[Cliente]_[DDMMYYYY].pdf
```

---

## 4. Formatos de Datos

| Dato | Formato Display | Formato Interno |
|------|----------------|-----------------|
| Moneda | $2,100.00 | Número con 2 decimales |
| Fecha | dd/mm/yyyy | Serial de Excel |
| Timestamp | dd/mm/yyyy hh:mm:ss | Now() |
| RFC | Mayúsculas | String 12-13 chars |
| Teléfono | 10 dígitos | Número o texto |
| Estatus | PAGADO/HOY VENCE/VENCIDO/PENDIENTE | Fórmula en Col.I |
| Modo | PRUEBA/PRODUCCIÓN | Texto en B2 |
| Estado cliente | ACTIVO/SUSPENDIDO | Texto en DIRECTORIO.I |
