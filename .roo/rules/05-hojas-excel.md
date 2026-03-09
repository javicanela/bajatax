# 05-hojas-excel.md — Mapa Maestro de Hojas y Columnas

> Roo Code: CONSULTA ESTE ARCHIVO antes de escribir cualquier referencia a celdas, columnas o hojas. Este es el mapa real verificado contra el Excel TEST_zip_rebuild.xlsm.

## CONFIGURACION

Columna B, celdas específicas. NO tiene headers en fila 1 — es un formulario vertical.

| Celda | Campo | Tipo | Uso |
|-------|-------|------|-----|
| B2 | MODO DEL SISTEMA | Lista: PRUEBA / PRODUCCIÓN | Interruptor maestro de todo el sistema |
| B5 | Nombre Despacho | Texto | "Baja Tax" — firma en mensajes y PDFs |
| B6 | Beneficiario | Texto | Titular cuenta bancaria |
| B7 | Banco | Texto | Institución bancaria |
| B8 | CLABE | Numérico 18 dígitos | Validación estricta |
| B9 | Teléfono Despacho | Texto | Formato +52 (XXX) XXX XXXX |
| B10 | Email Despacho | Texto | Correo oficial |
| B12 | Departamento | Texto | "Dpto. de Impuestos" |
| B14 | Número de Prueba | Texto | Destino WA cuando B2=PRUEBA. Formato: 52+10 dígitos |
| B16 | Ruta del Logo | Texto | Ruta local de imagen para PDFs |

**Notas**: Col A tiene etiquetas descriptivas (fila 1: "MODO DEL SISTEMA", fila 4: "DATOS DEL DESPACHO"). Col C tiene descripciones de ayuda. D y E tienen las opciones de la lista desplegable de B2.

---

## REGISTROS (Punto de Entrada — A-N)

Headers en fila 1. Esta hoja es donde CAE la información importada ANTES de distribuirse.

| Col | Header Real | Tipo | Notas |
|-----|------------|------|-------|
| A | RESPONSABLE | Texto | Asesor (de hoja Soportes) |
| B | NOMBRE DEL CONTRIBUYENTE | Texto | Razón social o nombre |
| C | RFC | Texto | 12-13 caracteres |
| D | EMAIL | Texto | Correo del cliente |
| E | TELÉFONO | Número | 10 dígitos |
| F | FECHA | Fecha | Fecha cobranza/emisión |
| G | CONCEPTO | Texto/Fecha | Descripción del servicio |
| H | MONTO | Número | Cantidad a cobrar |
| I | FACTURA | Texto | ID factura (OPCIONAL — no todos tienen) |
| J | REGIMEN | Texto | Código SAT o abreviación |
| K | VENCIMIENTO | Fecha | Fecha límite de pago |
| L | OPERACIONES | Texto | Tracker: "✓ OPERACIONES" cuando se distribuyó |
| M | DIRECTORIO | Texto | Tracker: "✓ DIRECTORIO" cuando se distribuyó |
| N | (Columna1) | Texto | Tracker: "✓ PROCESADO" cuando ambos completaron |

**Activador doble clic**: Al hacer doble clic en una fila, procesa el registro → distribuye a OPERACIONES y DIRECTORIO → marca trackers L, M, N.

**Lógica de procesamiento**: Ver agent.md sección 3.8 para el protocolo completo.

---

## OPERACIONES (Motor Operativo — A-T)

Headers en fila 1. 20 columnas. Esta es la hoja de trabajo diario.

| Col | Header Real | Tipo | Notas |
|-----|------------|------|-------|
| A | RESPONSABLE | Texto | Asesor asignado |
| B | ID FACTURA | Texto | Folio único (puede estar vacío) |
| C | REGIMEN | Texto | Código SAT |
| D | CLIENTE | Texto | Razón social |
| E | RFC | Texto | 12-13 chars — clave de vinculación |
| F | FECHA COBRANZA | Fecha | Emisión de la obligación |
| G | Concepto | Texto/Fecha | Descripción del servicio |
| H | Monto Base | Moneda | Cantidad total MXN |
| I | Estatus de Pago | **FÓRMULA** | PAGADO / HOY VENCE / VENCIDO / PENDIENTE |
| J | Fecha Vencimiento | Fecha | Límite de pago |
| K | Días Vencidos | **FÓRMULA** | @TODAY() - Col.J |
| L | Registro de Pago | Fecha/Hora | ⬡ Doble clic → estampa Now(), marca PAGADO |
| M | Teléfono | Número | 10 dígitos |
| N | Correo | Texto | Email del cliente |
| O | Acción WhatsApp | Texto | ⬡ Doble clic → dispara envío WA |
| P | Estado De Cuenta | Texto | ⬡ Doble clic → genera PDF consolidado por RFC |
| Q | EXCLUIR | Texto | Si tiene CUALQUIER valor → masivos la saltan |
| R | FECHA_PROX_ENVIO | Fecha | Próximo contacto programado |
| S | INTENTOS_ENVIO | Número | +1 cada envío WA. Rojo si ≥ 5 |
| T | ULTIMO_ENVIO_FECHA | Fecha/Hora | Stamp del último envío exitoso |

**⬡ = Activadores doble clic (Worksheet_BeforeDoubleClick):**
- Col.L → Marca PAGADO: estampa Now() y cambia estatus
- Col.O → Envía WhatsApp: dispara EnviarMensajeInteligente
- Col.P → Genera PDF: consolida por RFC y llama a pdf_server.py

**Fórmula Col.I (Estatus) — orden de prioridad:**
1. Si Col.D vacío → vacío
2. Si Col.L tiene valor → PAGADO
3. Si Col.J vacía → PENDIENTE
4. Si @TODAY() = Col.J → HOY VENCE
5. Si @TODAY() > Col.J → VENCIDO
6. Otro → PENDIENTE

**Fórmula Col.K (Días Vencidos):**
```
= @TODAY() - Col.J
```
Positivo = días de retraso. Negativo = días restantes.

**Formato visual Col.O**: Muestra texto dinámico según estatus → "▶ VENCIDO\nENVIAR WA" o "▶ HOY VENCE\nENVIAR WA"

**Formato visual Col.P**: Muestra "■ GENERAR PDF" con fondo gris claro, negritas.

---

## DIRECTORIO (Base Maestra Clientes — A-I)

Headers en fila 1. 9 columnas. RFC es ÚNICO por cliente.

| Col | Header Real | Tipo | Notas |
|-----|------------|------|-------|
| A | RFC | Texto | CLAVE ÚNICA. 12 chars (moral) o 13 (física) |
| B | CLIENTE | Texto | Razón social completa |
| C | CORREO | Texto | Email principal |
| D | TELEFONO | Número | 10 dígitos |
| E | REGIMEN | Texto | Código SAT |
| F | RESPONSIBLE | Texto | ⚠ TYPO en header real — debería ser RESPONSABLE |
| G | CLASIFICACIÓN | Texto | Etiqueta: PAGADOR PUNTUAL, DEUDOR CRÓNICO, etc. (vacía aún) |
| H | (sin header) | Fecha | FECHA ALTA — registro automático de ingreso |
| I | (sin header) | Texto | ESTADO_CLIENTE: ACTIVO / SUSPENDIDO |

**Activador Col.I**: Doble clic alterna entre ACTIVO y SUSPENDIDO. SUSPENDIDO → fondo rojo tenue, bloquea WA y PDF para ese RFC.

**⚠ Headers faltantes**: Col H y Col I no tienen header en la fila 1 del Excel real, pero contienen datos. Al referenciar, usar la posición o agregar headers.

---

## LOG ENVIOS (Bitácora WhatsApp — A-G)

Headers en fila 1.

| Col | Header | Tipo |
|-----|--------|------|
| A | FECHA/HORA | Timestamp |
| B | RESPONSABLE | Texto |
| C | CLIENTE | Texto |
| D | TELEFONO | Número |
| E | VARIANTE | VENCIDO / HOY VENCE / RECORDATORIO |
| F | MODO | PRUEBA / PRODUCCIÓN |
| G | RESULTADO | ENVIADO / CANCELADO / ERROR |

**⚠ Existen dos hojas**: "LOG ENVÍOS" (con acento) y "LOG ENVIOS" (sin acento). Usar "LOG ENVIOS" (sin acento) como la principal.

---

## BUSCADOR CLIENTE

Interfaz de búsqueda avanzada:
- Fila 2-3: Filtros (Responsable, Estatus, Concepto, Rango fechas, Días vencidos)
- Fila 3: Botones: `BUSCAR ▶` y `✕ LIMPIAR`
- Fila 5: Contadores: Registros encontrados, Total monto, Total vencido
- Fila 6: Headers de resultados: No, Cliente, Responsable, Concepto, Monto, Estatus, Vencimiento, Días, Teléfono, WhatsApp, PDF

---

## REPORTES CXC

Resumen de cuentas por cobrar por responsable:

| Col | Campo |
|-----|-------|
| A | Responsable |
| B | Cobrado ($) |
| C | Pendiente ($) |
| D | Vencido |
| E | Total |
| F | Flujo (%) |

Responsables: Joselyn, Denisse, Fernanda, Omar, Javier

---

## Soportes (hoja oculta/auxiliar)

Lista de asesores en Col.A: JOSSELYN, DENISSE, FERNANDA, OMAR, JAVIER.
Alimenta listas desplegables de "Responsable" en otras hojas.

---

## Relaciones entre Hojas

```
REGISTROS.C (RFC) ↔ DIRECTORIO.A (RFC) → vínculo de identidad
REGISTROS.C (RFC) ↔ OPERACIONES.E (RFC) → vínculo transaccional
DIRECTORIO.A (RFC) ↔ OPERACIONES.E (RFC) → cliente ↔ movimientos
DIRECTORIO.I (ESTADO_CLIENTE) → controla si se permite WA/PDF en OPERACIONES
CONFIGURACION.B2 (MODO) → controla destino de todos los envíos
OPERACIONES → LOG ENVIOS (cada envío WA genera una fila de log)
Soportes → listas desplegables de Responsable en todas las hojas
```

## Mapeo REGISTROS → OPERACIONES

| OPERACIONES | ← | REGISTROS |
|-------------|---|-----------|
| A (RESPONSABLE) | ← | A (RESPONSABLE) |
| B (ID FACTURA) | ← | I (FACTURA) — puede estar vacío |
| C (REGIMEN) | ← | J (REGIMEN) |
| D (CLIENTE) | ← | B (NOMBRE DEL CONTRIBUYENTE) |
| E (RFC) | ← | C (RFC) |
| F (FECHA COBRANZA) | ← | F (FECHA) |
| G (Concepto) | ← | G (CONCEPTO) |
| H (Monto Base) | ← | H (MONTO) |
| J (Fecha Vencimiento) | ← | K (VENCIMIENTO) |
| M (Teléfono) | ← | E (TELÉFONO) |
| N (Correo) | ← | D (EMAIL) |

Columnas I, K, L, O, P, Q, R, S, T de OPERACIONES son calculadas o gestionadas internamente — NO se importan.
