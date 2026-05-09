# CLAUDE.md — Agente: Carga de Real Consumido

## Descripción general

Este agente lee facturas publicitarias en PDF (Google DV360/CM360, Meta, TikTok) y carga los subtotales de cada factura en la columna **"REAL CONSUMIDO"** del archivo Excel correspondiente al anunciante.

---

## Estructura de carpetas esperada

```
📁 proyecto/
├── 📁 facturas_entrada/              ← PDFs de facturas a procesar
├── 📁 Archivos Real Consumido/       ← Archivos Excel a completar
│   ├── Real Consumido_Fargo_*.xlsx
│   ├── Real Consumido_Oroweat_*.xlsx
│   ├── Real Consumido_Rapiditas_*.xlsx
│   ├── Real Consumido_Artesano_*.xlsx
│   ├── Real Consumido_Salmas_*.xlsx
│   ├── Real Consumido_Bimbo_*.xlsx   ← o Back To School, Barcelona
│   ├── Real Consumido_Takis_*.xlsx
│   └── Real Consumido_TiaRosa_*.xlsx
└── CLAUDE_REAL_CONSUMIDO.md
```

---

## PROCESO 1 — Extraer datos de las facturas

Para cada PDF en `facturas_entrada/`, extraé los siguientes campos:

| Campo | Descripción |
|---|---|
| **Plataforma** | Google DV360, Google CM360, Meta, TikTok |
| **Anunciante** | Oroweat, Fargo, Rapiditas, Artesano, Salmas, Bimbo, Takis, Tia Rosa, etc. |
| **Mes de facturación** | Del campo `Billing Period` (Meta) o `Resumen del X mes` (Google) o `Periodo` (TikTok) |
| **Subtotal en ARS** | Campo `Subtotal en ARS` (Google), `Subtotal:` (Meta), `Subtotal ARS` (TikTok). **Nunca usar el Total con IVA.** |

### 1.1 Identificar la plataforma

| Indicador en el PDF | Plataforma |
|---|---|
| Google Argentina SRL + "Display and Video 360" | Google DV360 |
| Google Argentina SRL + "Campaign Manager 360" | Google CM360 |
| Meta Platforms Ireland Limited | Meta |
| IMS Argentina / Aleph / TikTok | TikTok |

### 1.2 Identificar el anunciante

Buscalo en el nombre de campaña o en el campo "Anunciante" del PDF:

| Patrón en el PDF | Anunciante |
|---|---|
| `Oroweat_ARG` / `OROWEAT` | Oroweat |
| `Artesano_ARG` / `ARTESANO` | Artesano |
| `Rapiditas_ARG` / `RAPIDITAS` | Rapiditas |
| `Salmas_ARG` / `SALMAS` / `SLMS` | Salmas |
| `_FAR_` / `FARGO` / `FRGP` | Fargo |
| `Takis_ARG` / `TAKIS` / `_TAK_` | Takis |
| `Tia Rosa_ARG` / `TIAROSA` | Tia Rosa |
| `_BIM_BIM_` / `BIMBO` / `Back to school` / `Barcelona` | Bimbo |

---

## PROCESO 2 — Completar el Excel de Real Consumido

### 2.1 Encontrar el archivo Excel del anunciante

Buscá en `Archivos Real Consumido/` el archivo cuyo nombre contenga el anunciante de la factura.  
Ejemplos:
- Factura de **Fargo** → `Real Consumido_Fargo_*.xlsx`
- Factura de **Oroweat** → `Real Consumido_Oroweat_*.xlsx`

> Si hay más de un archivo para el mismo anunciante, usá el que tenga la versión más alta (V4, V5, etc.) o el más reciente.

> Si no existe ningún archivo para el anunciante, registrá el caso y continuá con el siguiente.

---

### 2.2 Identificar la hoja correcta

El archivo puede tener varias hojas (CALENDARIO, OOH, Audiencias, etc.).  
Usá la hoja llamada **"Digital"**.  
Si hay varias hojas "Digital" (Digital E1, Digital E2, Digital Q1...), elegí la que corresponda al período de la factura según su nombre.

---

### 2.3 Localizar el mes en el Excel

La fila **13** contiene los encabezados de mes como celdas mergeadas.  
Buscá la celda (o rango mergeado) que contenga el nombre del mes de la factura.

| Mes de factura | Buscar en fila 13 |
|---|---|
| Enero | celda que contenga "ENERO" o "ENE" |
| Febrero | celda que contenga "FEBRERO" o "FEB" |
| Marzo | celda que contenga "MARZO" o "MAR" |
| Abril | celda que contenga "ABRIL" o "ABR" |
| Mayo | celda que contenga "MAYO" o "MAY" |
| ... | ... |

La columna de esa celda (o la primera columna del rango mergeado) es el **inicio del bloque del mes**.

---

### 2.4 Encontrar la columna "INVERSION" del mes

En la fila **15**, dentro del bloque del mes (desde la columna de inicio hasta la siguiente celda de mes), buscá la celda que contenga exactamente **"INVERSION"** (puede estar en mayúsculas).

Esa es la columna de inversión planificada para el mes.

---

### 2.5 Verificar o crear la columna "REAL CONSUMIDO"

Mirá la celda **inmediatamente a la derecha** de la columna INVERSION en la fila 15:

- **Si ya dice "REAL CONSUMIDO"** → esa es la columna donde escribir. No crear nada nuevo.
- **Si está vacía o tiene otro nombre** → escribir `REAL CONSUMIDO` en esa celda y usarla.

> **Nunca insertar una nueva columna** (eso desplazaría datos). Solo escribir el encabezado en la celda vacía.

---

### 2.6 Identificar la fila de la plataforma

En la columna **C** (Medio) y columna **E** (Formato) de las filas de datos (fila 16 en adelante), buscá la fila que corresponde a la plataforma de la factura:

| Plataforma de la factura | Buscar en columna C o E |
|---|---|
| **Meta** | C contiene "Facebook" |
| **Google DV360** | C contiene "Programmatic" o "DV360" |
| **Google CM360 — Display** | E contiene "Display" (y C es "Programmatic" o "Google") |
| **Google CM360 — Video/Bumper** | E contiene "Bumper" o "Video" (y C es "Programmatic" o "Google") |
| **TikTok** | C contiene "TikTok" |

> Si hay múltiples filas candidatas para Google (ej: una de Display y una de Bumper), mapear las líneas de detalle de la factura CM360 a sus filas:
> - Líneas con formato `Adserving Display` → fila de Display
> - Líneas con formato `Adserving Video` → fila de Bumper

> **Si una fila de medio no tiene ninguna factura correspondiente** (ej: "Mercado Libre / Meli Play") → ignorar esa fila completamente.

---

### 2.7 Escribir el valor en la celda

Una vez identificada la fila de la plataforma y la columna "REAL CONSUMIDO" del mes, escribí el **Subtotal en ARS** de la factura en la intersección.

Formato del número: `#,##0.00` (mismo formato que la columna INVERSION).

**Resumen del mapeo factura → celda:**

```
Factura Meta Fargo, Abril, subtotal = 9,209,265.69
→ Archivo: Real Consumido_Fargo_*.xlsx
→ Hoja: Digital
→ Mes: Fila 13 contiene "ABRIL" en columna BE (ejemplo)
→ INVERSION del mes: BE15 = "INVERSION"
→ REAL CONSUMIDO del mes: BF15 (crear si no existe)
→ Fila de Meta: fila donde C = "Facebook & Instagram" (ej: fila 16)
→ Celda a escribir: BF16 = 9,209,265.69
```

---

### 2.8 Guardar el archivo

Guardar el archivo Excel **con el mismo nombre y en la misma ubicación** (sobreescribir).  
No cambiar el nombre ni mover el archivo.

---

## Reglas generales

- El monto a cargar es siempre el **Subtotal en ARS** de la factura, **sin IVA, sin IIBB, sin percepciones**.
- Si la celda de REAL CONSUMIDO ya tiene un valor cargado, **sobreescribir** con el nuevo valor (o advertir al usuario si prefiere sumar).
- Si para un anunciante hay **DV360 + CM360** del mismo mes, son facturas distintas → cada una va a su fila correspondiente.
- Si no existe archivo Excel para algún anunciante (ej: Tia Rosa, Takis), registrarlo en un log y continuar.
- Al finalizar, imprimir en consola un resumen:
  ```
  ✅ Facturas procesadas: N
  📋 Celdas actualizadas: N
  ⚠️  Sin archivo Real Consumido: [lista de anunciantes]
  💾 Archivos modificados: [lista de archivos]
  ```

---

## Estructura del Excel (referencia técnica)

Basada en los archivos analizados (Fargo, Oroweat):

| Fila | Contenido |
|---|---|
| 7 | Nombre de la campaña / anunciante |
| 8-9 | Fecha inicio y fin |
| 13 | **Encabezados de mes** (celdas mergeadas: "FEBRERO", "MARZO", "ABRIL"…) |
| 14 | Sub-encabezados opcionales (TECH FEE, AD SERVING, etc.) |
| 15 | **Encabezados de columna** (INVERSION, REAL CONSUMIDO, IMPRESIONES…) |
| 16+ | **Filas de plataformas / formatos** |

**Columnas clave (pueden variar por archivo):**

| Columna | Contenido |
|---|---|
| C | Medio (Facebook & Instagram, Programmatic, TikTok, etc.) |
| E | Formato (Display, Bumper 6', Post/Video/Reels, etc.) |
| H | Ubicación (Facebook + Instagram, Youtube, RON, etc.) |
| AD | Inversión total (suma de todos los meses) |
| AG, AS, BE... | INVERSION por mes (varía según el archivo) |
| AH, AT, BF... | REAL CONSUMIDO por mes (inmediatamente a la derecha de INVERSION) |

---

*Última actualización: Mayo 2026*
