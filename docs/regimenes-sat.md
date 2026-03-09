# regimenes-sat.md — Tabla de Regímenes Fiscales SAT

> Referencia para validación del campo REGIMEN en REGISTROS (Col.J), OPERACIONES (Col.C) y DIRECTORIO (Col.E).

---

## Regímenes Principales (usados en BajaTax)

| Código | Nombre Técnico SAT | Explicación Sencilla | Tipo | Abreviación |
|--------|--------------------|-----------------------|------|-------------|
| 601 | General de Ley Personas Morales | Empresas y sociedades con fines de lucro | PM | PM |
| 605 | Sueldos y Salarios e Ingresos Asimilados a Salarios | Empleados que cobran por nómina | PF | PF |
| 606 | Arrendamiento | Personas que rentan casas, locales o bodegas | PF | PF |
| 612 | Personas Físicas con Actividades Empresariales y Profesionales | Freelancers, médicos, abogados, dueños de negocio | PF | PF |
| 625 | Régimen de las Actividades Empresariales con ingresos a través de Plataformas Tecnológicas | Ventas o servicios por apps (Amazon, Uber, Airbnb) | PF | PF |
| 626 | Régimen Simplificado de Confianza (RESICO) | Pequeños negocios con pago de impuestos reducido | PF | RESICO |
| 616 | Sin obligaciones fiscales | Personas sin ingresos propios (estudiantes, etc.) | PF | PF |

---

## Clasificación por Tipo de Contribuyente

### Persona Física (PF) — RFC de 13 caracteres
Formato: 4 letras + 6 dígitos (fecha nacimiento) + 3 homoclave

Regímenes aplicables: 605, 606, 612, 625, 626, 616

### Persona Moral (PM) — RFC de 12 caracteres
Formato: 3 letras + 6 dígitos (fecha constitución) + 3 homoclave

Regímenes aplicables: 601

### Asociación Civil (AC) — Variante de PM
- **AC** = Asociación Civil (con fines de lucro)
- **AC FNL** = Asociación Civil sin Fines de Lucro

---

## Abreviaciones Usadas en el Sistema

El campo REGIMEN en los archivos de origen puede venir como código numérico o como abreviación. El sistema debe reconocer ambos:

| Valor en archivo | Interpretación |
|------------------|----------------|
| 601 | Persona Moral |
| 605 | Persona Física — Sueldos |
| 606 | Persona Física — Arrendamiento |
| 612 | Persona Física — Actividades Empresariales |
| 625 | Persona Física — Plataformas |
| 626 | RESICO |
| 616 | Sin obligaciones |
| PF | Persona Física (genérico) |
| PM | Persona Moral (genérico) |
| AC | Asociación Civil |
| AC FNL | Asociación Civil sin Fines de Lucro |
| AC PM | Asociación Civil (variante PM) |
| RESICO | Régimen Simplificado de Confianza (626) |

---

## Validación de RFC

### Persona Física (13 caracteres)
```
Posición 1-4:  Letras (iniciales del nombre)
Posición 5-10: Dígitos (fecha nacimiento AAMMDD)
Posición 11-13: Homoclave alfanumérica
Ejemplo: AALR930211MX2
```

### Persona Moral (12 caracteres)
```
Posición 1-3:  Letras (iniciales de la razón social)
Posición 4-9:  Dígitos (fecha constitución AAMMDD)
Posición 10-12: Homoclave alfanumérica
Ejemplo: CAT931006E16
```

### RFCs Genéricos del SAT (rechazar en importación)
- XAXX010101000 — Público en general
- XEXX010101000 — Operaciones con extranjeros

---

## Notas de Implementación

1. Al importar, normalizar el campo régimen: si viene como texto ("PF", "AC FNL"), almacenar tal cual. Si viene como número (601, 626), almacenar como texto.
2. La validación de régimen es INFORMATIVA — no bloquear la importación si el valor no coincide con la tabla, solo advertir.
3. El régimen se usa para clasificación interna y puede aparecer en los PDFs de estado de cuenta.
