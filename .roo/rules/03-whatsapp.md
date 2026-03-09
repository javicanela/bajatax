# 03-whatsapp.md — Reglas del Módulo de WhatsApp

> Roo Code: lee estas reglas cuando trabajes en 03_Mod_WhatsApp.bas o cualquier lógica de mensajería.

## Motor de Envío: Evolution API

El sistema usa Evolution API (Docker local) en lugar de WhatsApp Web con Selenium:
- Endpoint base: `http://localhost:8080`
- Enviar texto: `POST /message/sendText/{instance}`
- Enviar media (PDF): `POST /message/sendMedia/{instance}`
- El VBA construye el mensaje y hace la petición HTTP al API local

**Fallback temporal**: Si Evolution API no está configurada aún, usar URLs `wa.me` como método alternativo:
```
https://wa.me/{numero}?text={mensaje_encodeado}
```

## Validación OBLIGATORIA Antes de Cualquier Envío

```vba
' SIEMPRE ejecutar esto ANTES de enviar
Dim modo As String
modo = Sheets("CONFIGURACION").Range("B2").Value

If modo = "" Then
    MsgBox "Configure el modo del sistema en CONFIGURACION celda B2", vbCritical
    Exit Sub
End If

' Mostrar ventana de confirmación
If MsgBox("Modo activo: " & modo & vbCrLf & "¿Desea continuar?", vbYesNo + vbQuestion) = vbNo Then
    Exit Sub
End If

' Determinar número destino
Dim destino As String
If modo = "PRUEBA" Then
    destino = Sheets("CONFIGURACION").Range("B14").Value
Else
    destino = celda_telefono_cliente
End If
```

## Codificación UTF-8 (URLEncodeUTF8)

**Tabla de reemplazos obligatorios:**

| Carácter | Hex | Carácter | Hex |
|----------|-----|----------|-----|
| á | %C3%A1 | Á | %C3%81 |
| é | %C3%A9 | É | %C3%89 |
| í | %C3%AD | Í | %C3%8D |
| ó | %C3%B3 | Ó | %C3%93 |
| ú | %C3%BA | Ú | %C3%9A |
| ñ | %C3%B1 | Ñ | %C3%91 |
| ü | %C3%BC | espacio | %20 |
| salto línea | %0A | $ | %24 |
| , | %2C | : | %3A |

**Negritas**: Rodear con `*` ANTES de encodear → `*$2,100.00*`
**Saltos de línea en VBA**: `Chr(10)` internamente → `%0A` en URL final

## Selección de Variante

Basado en Días Vencidos (Col.K de OPERACIONES):

| Condición | Variante | Color visual |
|-----------|----------|-------------|
| K = 0 | MSG_HOY — PRIORIDAD MÁXIMA | Naranja |
| K > 0 | MSG_VENCIDO | Rojo |
| K < 0 | MSG_RECORDATORIO | Azul |
| Estatus = PAGADO | BLOQUEAR envío + mensaje informativo | — |

**Seguro**: Si ESTADO_CLIENTE = "SUSPENDIDO" en DIRECTORIO → bloquear envío para ese RFC.

## Consolidación por Teléfono (Envío Masivo)

NO enviar un mensaje por fila. Agrupar por número:

```
1. Recorrer OPERACIONES filtrando: Estatus ≠ PAGADO AND Col.Q vacía
2. Agrupar filas por Col.M (Teléfono)
3. Por cada teléfono único:
   a. Sumar montos de todas sus filas → SUMA_TOTAL_MONTO
   b. Concatenar conceptos → lista con bullets
   c. Construir UN mensaje consolidado
   d. Enviar
4. Pausa aleatoria entre envíos
5. Actualizar Col.S y Col.T de CADA fila incluida en el consolidado
```

**Confirmación previa**: "Se detectaron {X} registros → se consolidarán en {Y} envíos. ¿Proceder?"

## Protocolo Anti-Baneo

1. **Pausa**: `Application.Wait Now + TimeSerial(0, 0, Int(Rnd * 7) + 8)` → 8-15 segundos
2. **Variación**: Espacios aleatorios invisibles al final del texto
3. **Límite**: Si Col.S ≥ 5 y estatus sigue VENCIDO → pintar celda rojo, sugerir llamada humana
4. **Excluir**: Si Col.Q tiene valor → saltar fila en masivo
5. **Consolidar**: Un mensaje por número, no por fila

## Post-Envío (actualización de celdas)

Después de cada envío exitoso:
```vba
' Incrementar intentos
ws.Cells(fila, colS).Value = ws.Cells(fila, colS).Value + 1
' Estampar fecha/hora
ws.Cells(fila, colT).Value = Now()
' Marcar visualmente Col.O en negritas
ws.Cells(fila, colO).Font.Bold = True
```

Registrar en LOG ENVIOS: fecha/hora, responsable, cliente, teléfono destino, variante, modo, resultado.

## Variables del Mensaje

| Variable | Origen |
|----------|--------|
| {CLIENTE} | OPERACIONES Col.D |
| {MONTO} | OPERACIONES Col.H (formato $#,##0.00) |
| {CONCEPTO} | OPERACIONES Col.G |
| {FECHA} | OPERACIONES Col.J (formato dd-mmm-yyyy) |
| {DIAS} | OPERACIONES Col.K (valor absoluto) |
| {BENEFICIARIO} | CONFIGURACION B6 |
| {BANCO} | CONFIGURACION B7 |
| {CLABE} | CONFIGURACION B8 |
| {DEPTO} | CONFIGURACION B12 |
| {TEL_DESPACHO} | CONFIGURACION B9 |
| {EMAIL_DESPACHO} | CONFIGURACION B10 |

> Ver plantillas completas de las 4 variantes en agent.md sección 1.2
