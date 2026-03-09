# 04-pdf.md — Reglas del Módulo PDF

> Roo Code: lee estas reglas cuando trabajes en 04_Mod_PDF.bas, 07_Mod_MasivoPDF.bas o python/pdf_server.py.

## Motor: pdf_server.py (Python)

NO usar ExportAsFixedFormat de VBA — falla en Mac sin impresora configurada, no maneja acentos en nombres de archivo, y el formato cambia si el zoom ≠ 100%.

**Flujo técnico:**
```
1. VBA identifica RFC de la fila activa
2. VBA busca TODOS los registros de ese RFC en OPERACIONES
3. VBA construye JSON con: datos cliente, pendientes, pagados, datos bancarios
4. VBA escribe JSON → TEMP/pdf_data_[RFC].json
5. VBA ejecuta: Shell "python3 " & ThisWorkbook.Path & "/python/pdf_server.py " & jsonPath
6. Python lee JSON → genera PDF con ReportLab
7. Python guarda → OUTPUT/DDMMYYYY/EdoCuenta_[Cliente]_[DDMMYYYY].pdf
8. VBA verifica que el archivo existe (Dir() ≠ "")
9. VBA registra en LOG ENVIOS
10. VBA limpia JSON temporal
```

## Estructura del JSON (VBA → Python)

```json
{
  "cliente": "Nombre del Cliente",
  "rfc": "AALR930211MX2",
  "despacho": {
    "nombre": "Baja Tax",
    "departamento": "Dpto. de Impuestos",
    "telefono": "+52 (664) ...",
    "email": "correo@bajatax.com",
    "beneficiario": "Oscar Abel Zambrano Rentería",
    "banco": "Banco Inbursa S.A.",
    "clabe": "036 028 5005 7775 7743",
    "logo_path": "/ruta/al/logo.png"
  },
  "pendientes": [
    {
      "no": 1,
      "concepto": "01-2026",
      "fecha_cobro": "01/02/2026",
      "vencimiento": "25/02/2026",
      "monto": 2100.00,
      "estatus": "VENCIDO",
      "dias_vencidos": 6
    }
  ],
  "liquidados": [
    {
      "no": 1,
      "concepto": "12-2025",
      "fecha_cobro": "01/01/2026",
      "fecha_pago": "15/01/2026 14:30",
      "monto": 2100.00,
      "metodo": "Transferencia"
    }
  ],
  "fecha_generacion": "03-mar-2026"
}
```

## Diseño Visual del PDF

**Orientación**: Vertical (Portrait), tamaño Carta, optimizado para 1 página.
**Fuente**: Arial o Calibri (Sans Serif).

### Encabezado
- Superior derecha: Logo (desde B16) + "Baja Tax" (azul oscuro) + datos contacto
- Superior izquierda: "ESTADO DE CUENTA" (negrita, azul oscuro, subrayado) + fecha generación
- Bloque ID: franja gris clara → CLIENTE: [NOMBRE] | RFC: [RFC] (negritas, mayúsculas)

### Sección 1: Conceptos Pendientes
| Propiedad | Valor |
|-----------|-------|
| Encabezado tabla | **#1F4E78** (azul oscuro), texto blanco, negrita |
| Columnas | No, Concepto, F.Cobro, Vencimiento, Monto, Estatus, Días Venc. |
| Sombreado | Zebra: blanco / gris muy claro |
| Estatus VENCIDO | Fondo rojo/rosa, letras rojas |
| Totalizador | "TOTAL PENDIENTE: $X,XXX.XX" → **rojo** |

### Sección 2: Historial Liquidados
| Propiedad | Valor |
|-----------|-------|
| Encabezado tabla | **#385623** (verde bosque), texto blanco, negrita |
| Columnas | No, Concepto, F.Cobro, Fecha Pago, Monto, Método |
| Sombreado | Zebra: blanco / verde muy claro |
| Totalizador | "TOTAL LIQUIDADO: $X,XXX.XX" → **verde** |

### Bloque de Pago (inferior)
- Encabezado "DATOS PARA TRANSFERENCIA" → fondo azul acero, texto blanco
- Beneficiario, Banco, CLABE → alineados a la izquierda
- Despedida: "Cualquier duda estamos a sus órdenes." en cursiva

### Pie de página
- Izquierda: Baja Tax | Teléfono | Correo
- Centro: CLABE | Banco | Beneficiario
- Derecha: "Página 1 de 1"

### Tamaños de letra
- Títulos sección: 12 pts negrita
- Cuerpo tablas: 10 pts
- Datos cliente/RFC: 11 pts negrita
- Contacto encabezado: 9 pts
- Pie de página: 8 pts

## Nomenclatura y Almacenamiento

- **Archivo**: `EdoCuenta_[NombreCliente]_[DDMMYYYY].pdf`
- **Carpeta**: `OUTPUT/DDMMYYYY/` → crear si no existe
- **Sanitizar nombre**: reemplazar ñ→n, acentos→sin acento, espacios→_ en nombre de archivo
- **Peso máximo**: < 500 KB para envío rápido por WhatsApp

## Consolidación por RFC

Al generar PDF (doble clic Col.P), NO imprimir solo la fila activa:
1. Tomar RFC de la fila activa
2. Buscar TODOS los registros con ese RFC en OPERACIONES
3. Separar en pendientes (VENCIDO, HOY VENCE, PENDIENTE) y liquidados (PAGADO)
4. Construir UN PDF con ambas secciones

## Generación Masiva (07_Mod_MasivoPDF.bas)

```
1. Filtrar: Estatus = VENCIDO, HOY VENCE o PENDIENTE
2. Excluir: Col.Q con valor (EXCLUIR)
3. Excluir: ESTADO_CLIENTE = "SUSPENDIDO" en DIRECTORIO
4. Agrupar por RFC → lista única
5. Por cada RFC: generar PDF consolidado
6. Barra de progreso: "15 de 48 completados"
7. Si un cliente falla → log del error, continuar con el siguiente
8. Reporte final: X generados, Y fallaron (con razón)
```

**Velocidad**: Proceso 100% local, sin pausas de seguridad (no hay riesgo de baneo). Puede generar decenas por minuto.

## Modo PRUEBA

Si B2 = "PRUEBA":
- PDFs se generan con datos reales del cliente (para verificar exactitud)
- Carpeta destino: `OUTPUT/PRUEBA/` en lugar de `OUTPUT/DDMMYYYY/`
- LOG ENVIOS registra MODO = "PRUEBA"

## Dependencias Python

```bash
pip3 install reportlab
```
- `reportlab` es el motor de generación PDF
- Python 3 viene pre-instalado en Mac
- pdf_server.py debe manejar su propio try/except y escribir errores a stderr
