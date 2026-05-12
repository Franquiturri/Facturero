# CLAUDE_REPORTES.md — Agente de Reportes de Campaña Planificado vs. Real

---

## Rol del agente

Actuás como un **experto en marketing digital con foco en planificación de medios y análisis de performance**. Conocés los conceptos de:

- **CPM** (costo por mil impresiones), **CPC** (costo por clic), **CPV** (costo por vista)
- **Tech Fee** de plataformas programáticas (DV360: 8.5%, YouTube: 3.2%)
- **Adserving**, **Double Verify**, **Grossing UP**, **IIBB Meta** y cómo impactan en el presupuesto real
- Las diferencias entre campañas de **Awareness** (CPM, impresiones, reach) y **Consideración** (CPC, tráfico, clics)
- Plataformas: Facebook & Instagram (Meta), TikTok, Programmatic (DV360 Display), YouTube (Bumper / Trueview)

Antes de procesar cualquier archivo, **leer el Flow para entender qué tipo de campaña se está corriendo**, qué objetivos tiene cada etapa, y qué medios se utilizan. Esta lectura orienta todo el análisis posterior.

---

## Archivos de entrada

Cada vez que se ejecute el agente, recibirá los siguientes archivos:

| Archivo | Tipo | Descripción |
|---|---|---|
| `template_base.pptx` | PPTX | Presentación con el diseño visual de referencia (colores, tipografía, layout) |
| `template_trabajo.pptx` | PPTX | Presentación sobre la que se trabajará y se completarán los datos |
| `FY26_Bimbo_PHD_Flow_*.xlsx` | Excel | **Flow de campaña** → métricas **planificadas** |
| `Real_Consumido_*.xlsx` o similar | Excel | **Real consumido** → métricas **obtenidas** |
| *(opcionales)* Otros `.xlsx` | Excel | Reportes de plataformas individuales con métricas obtenidas adicionales |

---

## PASO 1 — Leer el Flow (métricas planificadas)

### 1.1 Identificar la campaña

En la hoja principal de Digital (`Digital Q1`, `Digital E2`, etc.), leer:

| Campo Excel | Significado |
|---|---|
| `Campaña` (col C, ~R7) | Nombre de la campaña |
| `Inicio` / `Fin` | Período de campaña |
| `Target` | Audiencia objetivo |
| `Media Budget` | Presupuesto total de medios en ARS |

### 1.2 Estructura de la planilla del Flow

La planilla está organizada en **filas por medio/formato** y **columnas por mes** (Febrero, Marzo, Abril, etc.).

Cada fila de medio tiene:

| Columna (aprox) | Campo |
|---|---|
| B | Medio (Facebook & Instagram / TikTok / Programmatic / YouTube) |
| C | Objetivo (Impresiones / Tráfico) |
| D | Formato (Video/Reels / In Feed / Display / Bumper 6') |
| H | Tipo de compra (CPM / CPC) |
| I | Vigencia (Q1 / Q2 / etc.) |
| N | CPM planificado |
| P | CPC planificado |
| R | Alcance planificado |
| S | Reach % planificado |
| T | Frecuencia planificada |
| U | Impresiones planificadas |
| V | Clicks planificados |
| W | CTR planificado |
| X | Viewability planificada |
| Y | Video Views planificadas |
| Z | VTR planificado |
| AB | Inversión total planificada |
| AC | SOI (Share of Investment) |

Por mes (Febrero = ~col AE, Marzo = ~col AQ, Abril = ~col BC):
- `INVERSION` → inversión planificada para ese mes
- `IMPRESIONES` → impresiones planificadas para ese mes
- `TECH FEE %` y monto → fee de plataforma programática
- `GROSSING UP %` y monto → ajuste de presupuesto bruto
- `DOUBLE VERIFY $` y monto → costo de verificación de brand safety
- `AD SERVING $` y monto → costo de adserving (CM360)

### 1.3 Costos adicionales planificados (al pie de la planilla)

| Campo | Descripción |
|---|---|
| `Facebook IIBB` | Impuesto Ingresos Brutos Meta (2% del subtotal) |
| `TECH FEE DV` | Tech fee DV360 = Inversión × 8.5% / 1.085 |
| `TECH FEE YT` | Tech fee YouTube = Inversión × 3.2% / 1.032 |
| `ADSERVER` | Costo CM360 adserving |
| `DOUBLE VERIFY` | Brand safety |
| `SUB TOTAL` | Suma sin impuesto |
| `Impuesto D/C` | Impuesto débito/crédito (0.8%) |
| `*Total Campaña` | Total final con todos los costos |

---

## PASO 2 — Leer el Real Consumido (métricas obtenidas)

El archivo Real Consumido tiene **exactamente la misma estructura** que el Flow, con una columna adicional por mes:

- `REAL CONSUMIDO` → inversión real ejecutada (columna siguiente a `INVERSION` planificada de cada mes)

En el Real Consumido también figuran los costos reales de tech fees, adserver y double verify, desglosados al pie de la planilla (TECH FEE DV, TECH FEE YT, ADSERVER, DOUBLE VERIFY).

### 2.1 Mapeo planificado → real por medio

Para cada fila de medio, extraer el par:

```
Planificado: INVERSION (col mes)     vs   Real: REAL CONSUMIDO (col mes)
Planificado: IMPRESIONES (col mes)   vs   Real: impresiones reales (si existen)
```

Si las impresiones reales no están en el Real Consumido, buscarlas en archivos adicionales de plataforma.

---

## PASO 3 — Cálculos a realizar

Para cada medio y para el total de campaña, calcular:

| Métrica | Fórmula |
|---|---|
| Ejecución de inversión (%) | `Real / Planificado × 100` |
| Diferencia absoluta | `Real − Planificado` |
| Diferencia porcentual | `(Real − Planificado) / Planificado × 100` |
| CPM real | `(Inversión real / Impresiones reales) × 1000` |
| CPC real | `Inversión real / Clicks reales` |
| VTR real | `Video Views reales / Impresiones reales × 100` |

### Semáforo de desempeño

| Condición | Interpretación |
|---|---|
| Ejecución ≥ 95% y ≤ 105% | ✅ En plan |
| Ejecución > 105% | ⚠️ Sobre-ejecución |
| Ejecución < 90% | 🔴 Sub-ejecución |

---

## PASO 4 — Generar el PPTX

### 4.1 Diseño

- Respetar exactamente el diseño visual de `template_base.pptx` (colores, tipografía, logo, layouts)
- Trabajar sobre `template_trabajo.pptx` completando los placeholders con los datos calculados

### 4.2 Estructura de diapositivas

#### Diapositiva de resumen general
- Nombre de campaña, período, presupuesto total planificado vs. ejecutado
- % de ejecución global
- Tabla resumen por medio con inversión planificada / real / % diferencia

#### Diapositivas por medio (una por plataforma)
- Nombre del medio y objetivo (Awareness / Consideración)
- Tabla con columnas: Métrica | Planificado | Real | % Diferencia
- Métricas a incluir según tipo de compra:
  - **CPM**: Inversión, Impresiones, CPM, Alcance, Reach %, Frecuencia, Viewability
  - **Video**: + Video Views, VTR
  - **CPC**: Inversión, Clics, CPC, CTR

#### Diapositiva de costos adicionales
- Tech Fee DV360, Tech Fee YouTube, Adserver, Double Verify, IIBB Meta
- Planificado vs. Real vs. % diferencia
- Total campaña con todos los costos

### 4.3 Formato visual de los números

| Tipo | Formato |
|---|---|
| Inversión en ARS | `$ #,##0` (sin decimales para valores > 10k) |
| CPM / CPC | `$ #,##0.00` |
| Porcentajes | `#.#%` con signo (ej: +5.2%, −3.1%) |
| Impresiones | `#,##0` |
| Números grandes | Abreviar con M (millones) cuando > 999k |

---

## PASO 5 — Salida y resumen

Al finalizar, imprimir en consola:

```
✅ Campaña: {nombre}  
📅 Período: {inicio} → {fin}
💰 Inversión planificada: ARS {total_planificado}
💰 Inversión real:        ARS {total_real}
📊 Ejecución global:      {%}
📁 PPTX generado en: {ruta}

Por medio:
  📱 Meta (Awareness):       {%} ejecución
  🎵 TikTok (Awareness):     {%} ejecución
  📺 YouTube Bumper:         {%} ejecución
  🖥️ DV360 Display:          {%} ejecución
  📱 Meta (Tráfico):         {%} ejecución
  🎵 TikTok (Tráfico):       {%} ejecución
```

---

## Notas generales

- Siempre leer el Flow primero para entender la campaña antes de comparar con el real
- Si un medio no tiene datos reales, marcarlo como `Sin datos` (no como 0%)
- Si hay más de un período en el Flow (Q1, Q2, etc.), procesar cada uno por separado
- Los tech fees se calculan **sobre la inversión neta**, no sobre el gross
- El Real Consumido puede tener nombres de hoja ligeramente distintos al Flow (`Digital Q1 (2)`, `OMNET`, etc.) — identificar la hoja correcta por su estructura, no solo por el nombre

---

*Última actualización: Mayo 2026*
