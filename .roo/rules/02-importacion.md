# 02-importacion.md — Reglas del Módulo de Importación

> Roo Code: lee estas reglas cuando trabajes en 02_Mod_ImportarArchivos.bas o cualquier lógica de importación.

## Principio Fundamental

El sistema detecta datos por su CONTENIDO (patrones), no por posición ni nombre de columna. Debe funcionar con archivos sin headers, con headers en cualquier fila, con columnas en cualquier orden, y con datos mezclados con basura.

## Detección por Contenido (método primario)

Clasificar cada columna del archivo de origen analizando sus celdas:

| Dato | Patrón | Umbral |
|------|--------|--------|
| RFC | 3-4 letras + 6 dígitos + 2-3 alfanuméricos (12-13 chars total) | >50% de celdas coinciden |
| Email | Contiene `@` y al menos un `.` después | >50% coinciden |
| Teléfono | 10 dígitos (con o sin +52, 044, 045, paréntesis, guiones) | >50% coinciden |
| Monto | Numérico con decimales, rango 100-500,000 | >50% coinciden |
| Fecha | Objeto datetime o formato reconocible (dd/mm/yyyy, etc.) | >50% coinciden |
| Nombre | String >10 chars, no matchea otros patrones | >50% coinciden |
| Régimen | Matchea catálogo SAT (601,605,606,612,625,626,616,PF,PM,AC) | >30% coinciden |

## Detección de Headers (método complementario)

Solo aplica cuando el archivo tiene headers reconocibles:
1. Escanear primeras 20 filas
2. La fila donde hay más texto tipo-etiqueta (strings cortos, sin números) = fila de headers
3. Comparar contra diccionario de aliases (ver agent.md sección 3.3)
4. Los datos empiezan en la fila siguiente

## Normalización de Headers

Antes de comparar un header contra el diccionario:
```
1. Trim()
2. LCase()
3. Quitar acentos: á→a, é→e, í→i, ó→o, ú→u, ñ→n
4. Quitar: puntos, guiones, corchetes, paréntesis
5. Quitar palabras vacías: "de", "del", "la", "el"
```

## Estrategia de Dos Fases

**Fase 1 — Escaneo**: Analizar archivo, proponer mapeo con % de confianza por columna.
**Fase 2 — Confirmación**: Ventana al usuario mostrando: "Columna C → RFC (87%), Columna E → Email (94%)". El usuario confirma o corrige.

## Destino: REGISTROS (A-K)

Todo dato importado va a REGISTROS primero. Mapeo al esquema interno:

| REGISTROS | Contenido |
|-----------|-----------|
| A | RESPONSABLE (asignar o preguntar) |
| B | NOMBRE DEL CONTRIBUYENTE |
| C | RFC |
| D | EMAIL |
| E | TELÉFONO |
| F | FECHA (cobranza/emisión) |
| G | CONCEPTO |
| H | MONTO |
| I | FACTURA (opcional — no todos tienen) |
| J | REGIMEN |
| K | VENCIMIENTO |

Cols L, M, N son trackers internos — NO importar, el sistema los gestiona.

## Deduplicación

Clave compuesta para detectar si un registro ya existe en OPERACIONES:
```
DUPLICADO = mismo RFC + mismo Concepto + misma Fecha Cobranza + mismo Monto
```
- Si los 4 coinciden → preguntar: "¿Re-procesar o saltar?"
- ID Factura es verificación adicional cuando existe, pero NO es requisito
- Hash MD5 del archivo completo para detectar si el mismo archivo se importa dos veces

## Columnas a Ignorar (por ahora)

Si se detectan columnas con estos contenidos, no importar:
- Contraseña, Fiel, Sellos, Vigencia, DIOT, Cédula, Rep. Legal
- Lista extensible — si columna es desconocida → preguntar al usuario

## Manejo de Errores Específicos

- SIEMPRE cerrar Workbooks abiertos en caso de error (`wb.Close SaveChanges:=False`)
- Si un archivo falla → registrar error, continuar con el siguiente
- Timeout: si un archivo tiene >50,000 filas → advertir al usuario antes de procesar
- Log de cada importación: archivo, filas leídas, registros nuevos, duplicados saltados, errores

## Situaciones Especiales

- **Múltiples hojas**: analizar cada una, preguntar cuál importar o importar todas
- **Datos horizontales** (meses como columnas): detectar y ofrecer transponer
- **Celdas merged**: desmerge antes de leer
- **Filas vacías intermedias**: saltar, no tratar como fin de datos
- **Encoding**: intentar UTF-8, si falla probar Latin-1 / Windows-1252
