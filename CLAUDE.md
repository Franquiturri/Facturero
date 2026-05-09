# CLAUDE.md — Agente de Procesamiento de Facturas Publicitarias

## Descripción general

Este agente procesa facturas publicitarias en PDF de las plataformas **Google (DV360 / CM360)**, **Meta** y **TikTok**.  
Ejecuta dos procesos en secuencia cada vez que se agregan archivos a la carpeta `facturas_entrada/`.

---

## Estructura de carpetas esperada

```
📁 proyecto/
├── 📁 facturas_entrada/        ← acá se suben los PDFs sin clasificar
├── 📁 facturas_procesadas/
│   ├── 📁 Google/
│   │   ├── 📁 DV360/
│   │   └── 📁 CM360/
│   ├── 📁 Meta/
│   └── 📁 TikTok/
├── 📁 reportes/                ← acá se guarda el Excel generado
└── CLAUDE.md                   ← este archivo
```

---

## PROCESO 1 — Clasificación y renombrado de facturas

### 1.1 Leer cada PDF de `facturas_entrada/`

Para cada archivo PDF, extraé el contenido de texto y determiná:

- **Plataforma** (ver reglas abajo)
- **Anunciante / Producto** (deducirlo del nombre de campaña dentro del PDF)
- **Número de factura** (campo `Invoice #` o equivalente)

### 1.2 Reglas para identificar la plataforma

| Indicador en el PDF | Plataforma | Subcategoría |
|---|---|---|
| Emisor: `Google Ireland Limited` + menciona `Display & Video 360` o `DV360` | Google | DV360 |
| Emisor: `Google Ireland Limited` + menciona `Campaign Manager 360` o `CM360` o `adserving` | Google | CM360 |
| Emisor: `Meta Platforms Ireland Limited` | Meta | — |
| Emisor: `TikTok` / `ByteDance` | TikTok | — |

> Si hubiera ambigüedad, priorizá el nombre del emisor de la factura sobre el contenido de las líneas.

### 1.3 Reglas para identificar el anunciante

El anunciante se deduce del **nombre de campaña** que aparece en las líneas de detalle del PDF.  
Buscá el segmento del nombre que identifica el producto. Ejemplos de patrones:

```
OROWEAT_2026_BIM_...   → Oroweat
Artesano_2026_BIM_...  → Artesano
Bimbo_2026_BIM_...     → Bimbo
Rapiditas_2026_...     → Rapiditas
Back to school_2026_BIM_BIM_BIM → Bimbo
Salmas_2026_...        → Salmas
BimboZenissimo_ARG_... → Salmas  (Zenissimo es una línea de Salmas)
Bimbo_2026_BIM_FAR_... → Fargo
```

> Si la factura contiene campañas de **más de un anunciante** nombrá el archivo como `MultiAnunciante`.

### 1.4 Formato del nombre de archivo de salida

```
{N°_factura}_{Plataforma}_{Subcategoría}_{Anunciante}.pdf
```

Ejemplos:
```
5561017898_Google_DV360_Oroweat.pdf
5561022773_Google_CM360_Oroweat.pdf
252579344_Meta_Oroweat.pdf
```

Para Google sin subcategoría clara, usá `Google_DV360` o `Google_CM360` según corresponda.  
Para Meta y TikTok no hay subcategoría, omitila.

### 1.5 Destino de los archivos

| Plataforma | Subcategoría | Carpeta destino |
|---|---|---|
| Google | DV360 | `facturas_procesadas/Google/DV360/` |
| Google | CM360 | `facturas_procesadas/Google/CM360/` |
| Meta | — | `facturas_procesadas/Meta/` |
| TikTok | — | `facturas_procesadas/TikTok/` |

Copiá el archivo renombrado a su carpeta destino. **No elimines el original** de `facturas_entrada/`.

---

## PROCESO 2 — Generación del Excel de resumen

### 2.1 Archivo de salida

Generá (o actualizá si ya existe) el archivo:

```
reportes/Resumen_Facturas_{MES}_{AÑO}.xlsx
```

Ejemplo: `reportes/Resumen_Facturas_Abril_2026.xlsx`

El mes y año se deducen del campo `Billing Period` o `Invoice Date` de los PDFs.

---

### 2.2 Estructura del Excel

El archivo tendrá **dos hojas**:

---

#### HOJA 1 — `Resumen por Plataforma`

| N° Factura | Plataforma | Subcategoría | Anunciante | Subtotal (ARS) |
|---|---|---|---|---|

- Una fila por factura
- Al final, filas de subtotal agrupadas por plataforma
- Una fila de **TOTAL GENERAL**

> ⚠️ En facturas de Google el campo correcto es **`Subtotal en ARS`**, no el `Importe total adeudado`.incluye impuestos y no debe registrarse en esta planilla.

---

#### HOJA 2 — `Detalle por Campaña`

| N° Factura | Plataforma | Anunciante | Campaña | Formato | Subtotal Línea (ARS) |
|---|---|---|---|---|---|

- Una fila por **línea de detalle** dentro de cada factura
- El campo **Campaña** es el nombre completo de la campaña tal como figura en el PDF
- El campo **Formato** se deduce del nombre de campaña según estas reglas:

---

### 2.3 Reglas para identificar el Formato

Analizá el nombre de la campaña y aplicá la siguiente lógica:

**Para Google:**

| Patrón en la descripción de línea | Formato |
|---|---|
| `SEGUIMIENTO DE IMPRESIONES` / `CPM` + `DD_MMD` | Adserving Display - CPM |
| `SEGUIMIENTO DE IMPRESIONES` / `CPM` + `VID` o `YT` | Adserving Video - CPM |
| `SEGUIMIENTO DE CLICS` / `CPC` | Adserving - CPC |
| `Publicación anuncios gráficos` + `DSPL` | Display Programático |
| `Costo de medios` (DV360) | Costo de medios |
| `Tarifa de plataforma` (DV360) | Tech Fee |

**Para Meta:**

| Patrón en el nombre de campaña | Formato |
|---|---|
| `-TR` al final del nombre (tráfico) | Meta Tráfico |
| `_AWA_` o `-AW` al final (awareness) | Meta Awareness |


**Para TikTok:**

| Patrón | Formato |
|---|---|
| Menciona `TopView` | TopView |
| Menciona `In-Feed` | In-Feed |
| Menciona `Branded` | Branded Effect |
| Otros | Video |

---

### 2.4 Reglas de extracción de montos

- Buscá el campo **`Subtotal en ARS:`** en el PDF → ese es el valor a registrar en la columna `Subtotal (ARS)`
- **Nunca uses** `Importe total`, `Total`, ni ningún valor que incluya `VAT`, `IVA` o `Freight`
- Si el PDF tiene montos en negativo (créditos/ajustes), registralos como negativos en el Excel
- La moneda siempre es **ARS** salvo que el PDF indique explícitamente otra cosa

---

### 2.5 Formato visual del Excel

- Encabezados con fondo azul oscuro (`#1F4E79`) y texto blanco, negrita
- Filas de Google en **verde claro** (`#E2EFDA`)
- Filas de Meta en **azul claro** (`#D6E4F0`)
- Filas de TikTok en **rosa claro** (`#FFD9E8`)
- Filas de subtotal por plataforma con fondo amarillo claro (`#FFF2CC`), negrita
- Fila de TOTAL GENERAL con fondo naranja claro (`#FCE4D6`), negrita
- Columnas de montos con formato `#,##0.00`
- Paneles congelados en la primera fila de datos

### 2.6 Estructura de la Hoja 1 — Resumen por Plataforma

Columnas: `N° Factura | Plataforma | Subcategoría | Anunciante | Costo de Medios (ARS) | Tech Fee (ARS) | Subtotal (ARS) | IIBB Meta (ARS)`

**Sub-filas DV360** (debajo de cada factura DV360, en itálica):
- Una sub-fila por formato según el archivo `Spent Google/Spent por campaign.xlsx`
- Columna Subcategoría: `↳ {formato}` (DV360 Display, YouTube Trueview, YouTube Bumper, etc.)
- Columna Tech Fee: `subtotal_formato × rate / (1 + rate)`
  - YouTube (Trueview, Bumper, Masthead): rate = **3.2%**
  - DV360 (Display, Video): rate = **8.5%**
- Columna Subtotal: **inversión real en plataforma** = `subtotal_formato / (1 + rate)` = `subtotal_formato − tech_fee`
  → La fila principal DV360 conserva el total facturado completo (Costo de Medios + Tech Fee)

**IIBB Meta**: columna `IIBB Meta (ARS)` = 2% del subtotal de cada factura Meta

### 2.7 Fuente de datos para el desglose DV360

Archivo: `Spent Google/Spent por campaign.xlsx` (hoja `Data`)  
Columnas usadas: `Campaign` (anunciante), `Insertion Order` (formato), `Revenue` (monto)

Detección de formato desde el campo Insertion Order:

| Patrón en Insertion Order | Formato |
|---|---|
| `TRV` / `TRUEVIEW` / `STTV` | YouTube Trueview |
| `BUM` / `BUMPER` / `BUMP` | YouTube Bumper |
| `MASTHEAD` | YouTube Masthead |
| `DISPLAY` / `DISP` + `DV3` | DV360 Display |
| `YOU` sin BUM/TRV | YouTube Trueview |
| `VID` / `VIDE` sin YouTube | DV360 Video |

---

## Notas generales

- Procesá **todos** los PDFs que haya en `facturas_entrada/` cada vez que se ejecute el agente
- Si el Excel de resumen ya existe del mismo mes, **actualizalo** (no lo sobreescribas sin leer el contenido previo)
- Si un PDF no puede clasificarse, movelo a una carpeta `facturas_entrada/sin_clasificar/` y registrá el nombre del archivo en una hoja adicional del Excel llamada `Sin Clasificar`
- El agente debe imprimir en consola un resumen de lo procesado al finalizar:
  ```
  ✅ Procesados: N archivos
  📁 Google DV360: N | Google CM360: N | Meta: N | TikTok: N
  ⚠️  Sin clasificar: N
  💾 Excel guardado en: reportes/Resumen_Facturas_Abril_2026.xlsx
  ```

---

*Última actualización: Mayo 2026*
