---
name: python-pdf-server
description: >
  Usar esta skill para cualquier tarea relacionada con la generación de PDFs en
  BajaTax. Activar cuando el usuario mencione pdf_server.py, estados de cuenta,
  ReportLab, generación de PDF, diseño del estado de cuenta, colores del PDF,
  tablas de pendientes o liquidados, datos bancarios en PDF, pie de página,
  o cuando VBA necesite generar un PDF de un cliente. También activar si el PDF
  no se genera, tiene formato incorrecto, o el usuario quiere cambiar el diseño.
  Esta skill contiene el diseño visual exacto aprobado y la estructura del JSON
  que VBA envía a Python — úsala siempre que haya PDF involucrado en BajaTax.
---

# Python PDF Server — BajaTax

## Arquitectura del flujo

```
VBA (recopila datos del cliente desde OPERACIONES)
  → Construye dict Python / JSON
  → Escribe a TEMP/pdf_data_{RFC}.json
  → Shell: python3 src/python/pdf_server.py {jsonPath}
  → Python lee JSON → genera PDF con ReportLab
  → Guarda en OUTPUT/DDMMYYYY/EdoCuenta_{Cliente}_{DDMMYYYY}.pdf
  → VBA verifica existencia del archivo (timeout 30s)
  → VBA registra en LOG ENVIOS
  → VBA limpia archivo JSON temporal
```

---

## Schema del JSON que VBA envía

```json
{
  "cliente": "Nombre Completo del Cliente",
  "rfc": "XXXX000000XXX",
  "fecha_generacion": "09/03/2025",
  "despacho": {
    "nombre": "Baja Tax",
    "departamento": "Dpto. de Impuestos",
    "telefono": "+52 (664) XXX-XXXX",
    "email": "contacto@bajatax.mx",
    "beneficiario": "Oscar Abel Zambrano Rentería",
    "banco": "Banco Inbursa S.A.",
    "clabe": "036028500577757743",
    "logo_path": "/ruta/local/logo.png"
  },
  "pendientes": [
    {
      "no": 1,
      "concepto": "Declaración Anual 2024",
      "fecha_cobro": "01/01/2025",
      "vencimiento": "31/01/2025",
      "monto": 2500.00,
      "estatus": "VENCIDO",
      "dias_vencidos": 37
    }
  ],
  "liquidados": [
    {
      "no": 1,
      "concepto": "Declaración Mensual Nov 2024",
      "fecha_cobro": "01/11/2024",
      "fecha_pago": "15/11/2024",
      "monto": 1800.00,
      "metodo": "Transferencia"
    }
  ],
  "output_path": "OUTPUT/09032025/EdoCuenta_NombreCliente_09032025.pdf"
}
```

---

## pdf_server.py — estructura completa

```python
#!/usr/bin/env python3
"""
BajaTax PDF Server — Genera estados de cuenta en PDF
Uso: python3 pdf_server.py <ruta_json>
"""
import sys
import json
import os
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, HRFlowable)
from reportlab.lib.units import inch, cm

# Paleta de colores BajaTax
AZUL_OSCURO  = colors.HexColor('#1F4E78')  # Headers sección pendientes
VERDE_BOSQUE = colors.HexColor('#385623')  # Headers sección liquidados
AZUL_ACERO   = colors.HexColor('#2E75B6')  # Datos bancarios
GRIS_CLARO   = colors.HexColor('#F2F2F2')  # Zebra filas impares
VERDE_CLARO  = colors.HexColor('#EBF1DE')  # Zebra filas liquidados
ROJO_FONDO   = colors.HexColor('#FFE4E1')  # Fondo estatus VENCIDO
ROJO_TEXTO   = colors.HexColor('#C00000')  # Texto estatus VENCIDO / total pendiente
VERDE_TEXTO  = colors.HexColor('#375623')  # Texto total liquidado

def generate_pdf(data: dict) -> str:
    """Genera PDF y devuelve ruta del archivo creado."""
    output_path = data['output_path']
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        rightMargin=1.5*cm, leftMargin=1.5*cm,
        topMargin=1.5*cm, bottomMargin=2*cm
    )

    styles = getSampleStyleSheet()
    story = []

    # Encabezado
    story += _build_header(data, styles)
    story.append(Spacer(1, 0.3*cm))

    # Bloque de identificación del cliente
    story += _build_client_id(data, styles)
    story.append(Spacer(1, 0.4*cm))

    # Sección pendientes
    if data.get('pendientes'):
        story += _build_pendientes_table(data['pendientes'], styles)
        story.append(Spacer(1, 0.4*cm))

    # Sección liquidados
    if data.get('liquidados'):
        story += _build_liquidados_table(data['liquidados'], styles)
        story.append(Spacer(1, 0.4*cm))

    # Datos bancarios
    story += _build_payment_block(data['despacho'], styles)

    doc.build(story, onFirstPage=_footer_factory(data['despacho']),
              onLaterPages=_footer_factory(data['despacho']))
    return output_path


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIONES PENDIENTES DE IMPLEMENTAR
# Las siguientes funciones son necesarias para que generate_pdf() funcione.
# Se listan aquí como spec; deben implementarse en el mismo archivo.
# ─────────────────────────────────────────────────────────────────────────────

# TODO: def _build_header(data: dict, styles) -> list
#   Encabezado: logo + "Baja Tax" (der) | "ESTADO DE CUENTA" + fecha generación (izq)

# TODO: def _build_client_id(data: dict, styles) -> list
#   Franja gris con: CLIENTE: [nombre] | RFC: [rfc] en negrita

# TODO: def _build_liquidados_table(liquidados: list, styles) -> list
#   Igual que _build_pendientes_table pero header verde bosque #385623
#   Columnas: No, Concepto, F.Cobro, Fecha Pago, Monto, Método
#   Total en verde VERDE_TEXTO

# TODO: def _build_payment_block(despacho: dict, styles) -> list
#   Header "DATOS PARA TRANSFERENCIA" (fondo AZUL_ACERO, texto blanco)
#   Beneficiario / Banco / CLABE alineados izquierda
#   Despedida en cursiva

# TODO: def _footer_factory(despacho: dict)
#   Retorna función canvas para onFirstPage/onLaterPages
#   Izq: Baja Tax | Tel | Email  |  Centro: CLABE | Banco | Beneficiario  |  Der: "Página X de Y"


def _build_pendientes_table(pendientes: list, styles) -> list:
    headers = ['No', 'Concepto', 'F.Cobro', 'Vencimiento', 'Monto', 'Estatus', 'Días Venc.']
    table_data = [headers]

    for item in pendientes:
        row = [
            str(item['no']),
            item['concepto'],
            item['fecha_cobro'],
            item['vencimiento'],
            f"${item['monto']:,.2f}",
            item['estatus'],
            str(item.get('dias_vencidos', ''))
        ]
        table_data.append(row)

    # Fila de total
    total = sum(i['monto'] for i in pendientes)
    table_data.append(['', '', '', 'TOTAL PENDIENTE:', f"${total:,.2f}", '', ''])

    col_widths = [0.8*cm, 6*cm, 2*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2*cm]
    t = Table(table_data, colWidths=col_widths, repeatRows=1)

    style = TableStyle([
        # Header
        ('BACKGROUND', (0,0), (-1,0), AZUL_OSCURO),
        ('TEXTCOLOR',  (0,0), (-1,0), colors.white),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,0), 9),
        ('ALIGN',      (0,0), (-1,0), 'CENTER'),
        # Cuerpo
        ('FONTSIZE',   (0,1), (-1,-2), 8),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, GRIS_CLARO]),
        # Total
        ('FONTNAME',   (3,-1), (4,-1), 'Helvetica-Bold'),
        ('TEXTCOLOR',  (3,-1), (4,-1), ROJO_TEXTO),
        ('ALIGN',      (3,-1), (4,-1), 'RIGHT'),
        # Bordes
        ('GRID',       (0,0), (-1,-1), 0.5, colors.grey),
    ])

    # Colorear filas VENCIDO
    for i, item in enumerate(pendientes, start=1):
        if item['estatus'] == 'VENCIDO':
            style.add('BACKGROUND', (0,i), (-1,i), ROJO_FONDO)
            style.add('TEXTCOLOR',  (5,i), (5,i), ROJO_TEXTO)

    t.setStyle(style)
    return [t]


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Uso: python3 pdf_server.py <ruta_json>")
        sys.exit(1)

    json_path = sys.argv[1]
    if not os.path.exists(json_path):
        print(f"ERROR: No se encontró {json_path}")
        sys.exit(1)

    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    output = generate_pdf(data)
    print(f"PDF generado: {output}")
    sys.exit(0)
```

---

## Diseño visual aprobado

### Encabezado
- Superior derecha: Logo + "Baja Tax" (azul oscuro, bold) + "Dpto. de Impuestos" + contacto
- Superior izquierda: "ESTADO DE CUENTA" (negrita, azul oscuro, subrayado) + "Generado el [fecha]"
- Franja gris: CLIENTE: [NOMBRE] | RFC: [RFC] en negrita

### Tabla Pendientes
- Header: fondo `#1F4E78` (azul oscuro), texto blanco bold
- Columnas: No, Concepto, F.Cobro, Vencimiento, Monto, Estatus, Días Venc.
- Zebra: blanco / `#F2F2F2`
- Fila VENCIDO: fondo `#FFE4E1`, estatus en rojo `#C00000`
- Total: "TOTAL PENDIENTE: $X,XXX.XX" en rojo, bold

### Tabla Liquidados
- Header: fondo `#385623` (verde bosque), texto blanco bold
- Columnas: No, Concepto, F.Cobro, Fecha Pago, Monto, Método
- Zebra: blanco / `#EBF1DE`
- Total: "TOTAL LIQUIDADO: $X,XXX.XX" en verde, bold

### Bloque de pago
- Header "DATOS PARA TRANSFERENCIA": fondo azul acero `#2E75B6`, texto blanco
- Beneficiario, Banco, CLABE alineados izquierda
- Despedida en cursiva

### Pie de página
- Izquierda: Baja Tax | Teléfono | Correo
- Centro: CLABE | Banco | Beneficiario
- Derecha: "Página 1 de 1"

---

## Reglas críticas

- Peso máximo: < 500 KB (envío rápido por WhatsApp)
- Nombre archivo: `EdoCuenta_{NombreCliente}_{DDMMYYYY}.pdf` — sin acentos, espacios→guión bajo
- Carpeta: `OUTPUT/DDMMYYYY/` — crear automáticamente si no existe
- Un PDF por RFC (no por fila) — incluir TODOS los registros del cliente
- Si el cliente no tiene pendientes, omitir esa sección (no tabla vacía)
- Subscripts/superscripts: usar tags `<sub>` y `<super>` en Paragraph, nunca Unicode
