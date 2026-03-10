---
name: whatsapp-cobranza
description: >
  Usar esta skill para cualquier tarea relacionada con el envío de mensajes WhatsApp
  en BajaTax. Activar cuando el usuario mencione Evolution API, envío masivo,
  anti-baneo, URLEncodeUTF8, mensajes de cobranza, variantes de mensaje (VENCIDO,
  HOY VENCE, RECORDATORIO), mensaje consolidado, pausa entre envíos, actualizar
  Col.S o Col.T de OPERACIONES, o cuando el código necesite construir y enviar
  mensajes WhatsApp con datos de clientes. También activar si el usuario pregunta
  qué número usar en modo PRUEBA vs PRODUCCIÓN, o cómo evitar que WhatsApp banee
  el número. Esta skill contiene las plantillas exactas de mensaje de BajaTax y
  la lógica completa de envío — úsala siempre que haya WhatsApp involucrado.
---

# WhatsApp Cobranza — BajaTax

## Arquitectura de envío

```
VBA lee datos de OPERACIONES
  → Selecciona variante de mensaje (por Col.K Días Vencidos)
  → Construye texto del mensaje
  → URLEncodeUTF8 (caracteres especiales)
  → Valida modo PRUEBA/PRODUCCIÓN (CONFIGURACION.B2)
  → POST a Evolution API (localhost:8080)
  → Actualiza Col.S (INTENTOS_ENVIO +1) y Col.T (ULTIMO_ENVIO_FECHA = Now())
  → Registra en LOG ENVIOS
```

---

## Evolution API — Endpoints

```vba
' URL base — Docker local, siempre disponible sin costo ni límites
Const C_EVO_BASE As String = "http://localhost:8080"

' Enviar texto:
' POST http://localhost:8080/message/sendText/{instanceName}
' Body: {"number":"52XXXXXXXXXX","text":"mensaje"}

' Enviar media + texto (para PDF adjunto):
' POST http://localhost:8080/message/sendMedia/{instanceName}
' Body: {"number":"52XXXXXXXXXX","mediatype":"document","media":"<base64>",
'        "fileName":"EdoCuenta.pdf","caption":"texto"}
```

### Función de envío cross-platform (Mac: curl / Windows: WinHTTP)
```vba
' Requiere Util_IsMac() de Mod_Sistema (ver skill cross-platform)
Function EnviarWA(numero As String, mensaje As String, instanceName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim url As String
    url = "http://localhost:8080/message/sendText/" & instanceName
    
    Dim body As String
    body = "{""number"":""" & numero & """,""text"":""" & _
           Replace(mensaje, """", "\""") & """}"
    
    If Util_IsMac() Then
        ' Mac: escribir body a JSON temporal y enviar con curl
        Dim tmpPath As String
        tmpPath = ThisWorkbook.Path & Application.PathSeparator & _
                  "TEMP" & Application.PathSeparator & "wa_payload.json"
        Dim f As Integer: f = FreeFile
        Open tmpPath For Output As #f
        Print #f, body
        Close #f
        ' curl síncrono via bash -c; Evolution API corre en localhost
        Shell "bash -c 'curl -s -X POST """ & url & """ " & _
              "-H ""Content-Type: application/json"" " & _
              "-d @""" & tmpPath & """'"
        ' En Mac Shell no devuelve status HTTP — asumir éxito si no hubo excepción
        EnviarWA = True
    Else
        ' Windows: WinHTTP síncrono permite leer status de respuesta
        Dim http As Object
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.Open "POST", url, False
        http.SetRequestHeader "Content-Type", "application/json"
        http.Send body
        If http.Status = 200 Or http.Status = 201 Then
            EnviarWA = True
        Else
            LogEvento "WA_ERROR", "EnviarWA", "Status: " & http.Status & " | " & numero
            EnviarWA = False
        End If
    End If
    Exit Function
ErrorHandler:
    LogEvento "WA_ERROR", "EnviarWA", "Excepción: " & Err.Description
    EnviarWA = False
End Function
```

---

## Selección de variante por Col.K (Días Vencidos)

```vba
Function GetVarianteMensaje(diasVencidos As Long) As String
    Select Case True
        Case diasVencidos = 0:  GetVarianteMensaje = "HOY_VENCE"    ' Naranja — prioridad máxima
        Case diasVencidos > 0:  GetVarianteMensaje = "VENCIDO"      ' Rojo — tono urgente
        Case diasVencidos < 0:  GetVarianteMensaje = "RECORDATORIO" ' Azul — tono amigable
    End Select
End Function
```

---

## Plantillas de mensaje (texto exacto aprobado por el despacho)

### VENCIDO (Col.K > 0)
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

### HOY VENCE (Col.K = 0)
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

### RECORDATORIO (Col.K < 0)
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

### CONSOLIDADO (mismo teléfono, múltiples facturas)
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

---

## URLEncodeUTF8 — tabla completa

```vba
Function URLEncodeUTF8(texto As String) As String
    Dim result As String: result = texto
    ' Salto de línea PRIMERO (antes de reemplazar otros chars)
    result = Replace(result, Chr(13) & Chr(10), "%0A")
    result = Replace(result, Chr(10), "%0A")
    result = Replace(result, Chr(13), "%0A")
    ' Espacios
    result = Replace(result, " ", "%20")
    ' Caracteres especiales
    result = Replace(result, "$", "%24")
    result = Replace(result, ",", "%2C")
    result = Replace(result, ":", "%3A")
    result = Replace(result, "(", "%28")
    result = Replace(result, ")", "%29")
    result = Replace(result, "•", "%E2%80%A2")
    ' Vocales con acento — minúsculas
    result = Replace(result, "á", "%C3%A1")
    result = Replace(result, "é", "%C3%A9")
    result = Replace(result, "í", "%C3%AD")
    result = Replace(result, "ó", "%C3%B3")
    result = Replace(result, "ú", "%C3%BA")
    result = Replace(result, "ñ", "%C3%B1")
    result = Replace(result, "ü", "%C3%BC")
    ' Vocales con acento — mayúsculas
    result = Replace(result, "Á", "%C3%81")
    result = Replace(result, "É", "%C3%89")
    result = Replace(result, "Í", "%C3%8D")
    result = Replace(result, "Ó", "%C3%93")
    result = Replace(result, "Ú", "%C3%9A")
    result = Replace(result, "Ñ", "%C3%91")
    result = Replace(result, "Ü", "%C3%9C")
    URLEncodeUTF8 = result
End Function

' Negritas en WhatsApp: rodear con * ANTES de encodear
' Ejemplo: "*" & Format(monto,"$#,##0.00") & "*"
' Resultado: *$2,100.00* → *%242%2C100.00* (WhatsApp lo muestra en negrita)
```

---

## Protocolo anti-baneo

```vba
' 1. Pausa aleatoria entre envíos: 8-15 segundos
Application.Wait Now() + TimeSerial(0, 0, Int(Rnd() * 7) + 8)

' 2. Variación mínima invisible (espacio aleatorio al final)
Dim padding As String
padding = String(Int(Rnd() * 3) + 1, " ")
mensajeFinal = mensaje & padding

' 3. Consolidar por teléfono: UN mensaje por número, no por fila
'    Si cliente tiene 3 facturas → 1 mensaje consolidado, no 3 mensajes

' 4. Límite de intentos: si Col.S >= 5 y sigue VENCIDO → pintar rojo, escalar a humano
If wsOps.Cells(fila, colIntentos).Value >= 5 Then
    wsOps.Cells(fila, colIntentos).Interior.Color = RGB(255, 0, 0)
    LogEvento "WA_ESCALADO", "EnvioMasivo", cliente & " — 5+ intentos sin respuesta"
    GoTo SiguienteCliente
End If

' 5. Respetar EXCLUIR (Col.Q): si tiene cualquier valor, saltar la fila
If wsOps.Cells(fila, colExcluir).Value <> "" Then GoTo SiguienteCliente
```

---

## Post-envío — actualizar celdas

```vba
' Después de cada envío exitoso:
Application.EnableEvents = False
wsOps.Cells(fila, colIntentos).Value = wsOps.Cells(fila, colIntentos).Value + 1
wsOps.Cells(fila, colUltimoEnvio).Value = Now()
wsOps.Cells(fila, colAccionWA).Font.Bold = True  ' Visual: indica que se despachó
Application.EnableEvents = True

LogEvento "WA_ENVIADO", "EnvioMasivo", cliente & " | " & telDestino & " | " & variante & _
          " | " & IIf(modoPrueba, "PRUEBA", "PRODUCCIÓN")
```

---

## Modo PRUEBA vs PRODUCCIÓN

```vba
' Determinar destino ANTES de enviar:
Dim telDestino As String
Dim modoPrueba As Boolean

modoPrueba = (modo = "PRUEBA")

If modoPrueba Then
    telDestino = cfg("tel_prueba")  ' CONFIGURACION.B14 — todos los mensajes aquí
Else
    telDestino = Util_FormatPhoneForWA(wsOps.Cells(fila, colTel).Value)
    ' Validar que tenga teléfono:
    If Len(telDestino) <> 12 Then
        LogEvento "WA_SKIP", "EnvioMasivo", cliente & " — teléfono inválido: " & telDestino
        GoTo SiguienteCliente
    End If
End If

' Confirmar antes de envío masivo en PRODUCCIÓN:
If Not modoPrueba Then
    If MsgBox("⚠️ PRODUCCIÓN activa. ¿Confirma envío masivo a " & _
              totalClientes & " clientes?", vbYesNo + vbCritical) = vbNo Then
        Exit Sub
    End If
End If
```
