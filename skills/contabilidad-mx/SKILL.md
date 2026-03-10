---
name: contabilidad-mx
description: >
  Usar esta skill para cualquier tarea relacionada con datos fiscales mexicanos en
  BajaTax. Activar cuando el usuario mencione RFC, CLABE, régimen fiscal, SAT,
  persona física, persona moral, RESICO, validación de datos de clientes, importación
  de archivos de Contalink, mapeo de columnas fiscales, o cuando el código necesite
  validar o clasificar datos de contribuyentes mexicanos. También activar ante dudas
  sobre formato RFC (12 vs 13 chars), cómo detectar si es empresa o persona,
  o qué hacer con regímenes desconocidos. Esta skill es la referencia fiscal MX
  para todo el sistema BajaTax — úsala siempre que haya datos del SAT involucrados.
---

# Contabilidad MX — Contexto Fiscal BajaTax

## RFC — Identificador central del sistema

Todo el sistema de deduplicación y vinculación entre hojas depende del RFC.

```vba
' Persona Física:  13 chars — XXXX000000XXX (4 letras + fecha AAMMDD + 3 homoclave)
' Persona Moral:   12 chars — XXX000000XXX  (3 letras + fecha AAMMDD + 3 homoclave)

Function Util_ValidateRFC(rfc As String) As Boolean
    Dim r As String: r = UCase(Trim(rfc))
    ' RFCs genéricos del SAT — rechazar
    If r = "XAXX010101000" Or r = "XEXX010101000" Then
        Util_ValidateRFC = False: Exit Function
    End If
    Select Case Len(r)
        Case 12
            Util_ValidateRFC = (r Like "[A-Z][A-Z][A-Z][0-9][0-9][0-9][0-9][0-9][0-9][A-Z0-9][A-Z0-9][A-Z0-9]")
        Case 13
            Util_ValidateRFC = (r Like "[A-Z][A-Z][A-Z][A-Z][0-9][0-9][0-9][0-9][0-9][0-9][A-Z0-9][A-Z0-9][A-Z0-9]")
        Case Else
            Util_ValidateRFC = False
    End Select
End Function

Function Util_GetTipoPersona(rfc As String) As String
    Select Case Len(Trim(rfc))
        Case 12: Util_GetTipoPersona = "PM"
        Case 13: Util_GetTipoPersona = "PF"
        Case Else: Util_GetTipoPersona = "DESCONOCIDO"
    End Select
End Function
```

---

## Regímenes Fiscales SAT

| Código | Nombre | Tipo | Abreviación interna |
|--------|--------|------|---------------------|
| 601 | General de Ley Personas Morales | PM | PM |
| 605 | Sueldos y Salarios | PF | PF |
| 606 | Arrendamiento | PF | PF |
| 612 | Actividades Empresariales y Profesionales | PF | PF |
| 625 | Plataformas Tecnológicas (Uber, Airbnb) | PF | PF |
| 626 | RESICO — Régimen Simplificado de Confianza | PF | RESICO |
| 616 | Sin obligaciones fiscales | PF | PF |

```vba
Function Util_NormalizeRegimen(valor As String) As String
    Dim v As String: v = Trim(UCase(valor))
    Select Case v
        Case "601":                          Util_NormalizeRegimen = "PM"
        Case "605","606","612","625","616":   Util_NormalizeRegimen = "PF"
        Case "626":                          Util_NormalizeRegimen = "RESICO"
        Case "PF","PM","AC","AC FNL","AC PM","RESICO": Util_NormalizeRegimen = v
        Case Else: Util_NormalizeRegimen = valor  ' Guardar como viene, no bloquear
    End Select
End Function
```

---

## CLABE — 18 dígitos

```vba
Function Util_ValidateCLABE(clabe As String) As Boolean
    Dim c As String: c = Replace(Replace(Trim(clabe)," ",""),"-","")
    If Len(c) <> 18 Then Util_ValidateCLABE = False: Exit Function
    Dim i As Integer
    For i = 1 To 18
        If Not IsNumeric(Mid(c,i,1)) Then Util_ValidateCLABE = False: Exit Function
    Next i
    Util_ValidateCLABE = True
End Function
```

---

## Teléfonos mexicanos

```vba
Function Util_NormalizePhone(tel As String) As String
    ' Extraer solo dígitos
    Dim clean As String, i As Integer
    For i = 1 To Len(tel)
        If IsNumeric(Mid(tel,i,1)) Then clean = clean & Mid(tel,i,1)
    Next i
    ' Quitar prefijo +52 o 52 si ya tiene 12 dígitos
    If Left(clean,2) = "52" And Len(clean) = 12 Then clean = Mid(clean,3)
    If Len(clean) = 10 Then Util_NormalizePhone = clean Else Util_NormalizePhone = tel
End Function

Function Util_FormatPhoneForWA(tel As String) As String
    ' WhatsApp Evolution API: formato 52XXXXXXXXXX (sin + ni espacios)
    Dim clean As String: clean = Util_NormalizePhone(tel)
    If Len(clean) = 10 Then Util_FormatPhoneForWA = "52" & clean Else Util_FormatPhoneForWA = clean
End Function
```

---

## Normalizar headers para importación

```vba
Function Util_NormalizeHeader(header As String) As String
    Dim h As String: h = LCase(Trim(header))
    h = Replace(h,"á","a"): h = Replace(h,"é","e")
    h = Replace(h,"í","i"): h = Replace(h,"ó","o")
    h = Replace(h,"ú","u"): h = Replace(h,"ñ","n")
    h = Replace(h,".",""):  h = Replace(h,"[","")
    h = Replace(h,"]",""):  h = Replace(h,"-"," ")
    Util_NormalizeHeader = Trim(h)
End Function
```

## Aliases de headers reconocidos

| Campo BajaTax | Aliases (después de normalizar) |
|---|---|
| NOMBRE | nombre del contribuyente, cliente, razon social, contribuyente, nombre completo |
| RFC | rfc, r f c, rfc del contribuyente, registro federal |
| EMAIL | email, e mail, correo, correo electronico, mail |
| TELEFONO | telefono, tel, celular, numero, whatsapp |
| CONCEPTO | concepto, descripcion, servicio, comentarios |
| MONTO | monto, monto base, total, importe, honorarios, cantidad |
| FACTURA | factura, id factura, folio, folio de factura, numero de factura |
| REGIMEN | regimen, regimen fiscal, tipo regimen |
| VENCIMIENTO | vencimiento, fecha vencimiento, fecha de vencimiento, fecha limite, vigencia impuestos |

---

## Columnas obligatorias vs opcionales

**Obligatorias** — sin al menos una, importación NO procede:
- RFC o NOMBRE DEL CONTRIBUYENTE

**Opcionales** — se importan si existen, vacías si no:
- EMAIL, TELEFONO, CONCEPTO, MONTO, FACTURA, REGIMEN, VENCIMIENTO

**Ignorar siempre** (datos del SAT irrelevantes para BajaTax):
- Contraseña, Fiel, Sellos, Vigencia, DIOT, Cedula, Rep. Legal

> Columna desconocida → preguntar al usuario, nunca descartar silenciosamente.
> Para el algoritmo completo de importación, ver skill **excel-importacion**.
