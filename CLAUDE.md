# CLAUDE.md — Felipe Caballero
## Quién soy
Felipe Caballero — Product Manager en UnionX. Email: felipe@unionx.cl

---

# Proyecto: Automatización Reportes Semanales UnionX

## Contexto general
Script Python que actualiza automáticamente cada lunes a las 13:00 el archivo Excel
`FORECAST FINAL SKU 26-27.xlsx` con datos desde 3 fuentes de Google.

---

## Rutas clave
| Archivo | Ruta |
|---|---|
| Script principal | `C:\Users\felip\Desktop\UnionX Cloude\actualizar_reportes.py` |
| Script análisis planificación | `C:\Users\felip\Desktop\UnionX Cloude\analisis_stock_critico.py` |
| Script reportes CST/UNID | `C:\Users\felip\Desktop\UnionX Cloude\format_td_reportes_cst.py` |
| Script partial refresh | `C:\Users\felip\Desktop\UnionX Cloude\partial_refresh.py` |
| Credenciales Google | `C:\Users\felip\Desktop\UnionX Cloude\credentials.json` |
| Log de ejecución | `C:\Users\felip\Desktop\UnionX Cloude\log.txt` |
| Excel objetivo | `C:\Users\felip\Desktop\UNIONX\FORECAST FINAL SKU\FORECAST FINAL SKU 26-27 V2.xlsx` |
| Output análisis | `C:\Users\felip\Desktop\UNIONX\FORECAST FINAL SKU\Analisis Planificacion\analisis_planificacion_ABR26.xlsx` |
| Python | `C:\Users\felip\AppData\Local\Programs\Python\Python312\python.exe` |

---

## Fuentes de datos
| # | Fuente | ID | Destino Excel |
|---|---|---|---|
| 1 | Google Sheets "Matriz stock" / pestaña `BaseStk` | `1N5TpIQrFCJwzxyxtueyNi2UoASONd_TbiLcz17a3fag` | Pestaña `STOCK` |
| 2 | Google Sheets "Importaciones UnionX" / pestaña 0 | `1RpxZ69Wnfcots006Hp5fzawYxhUsscW03O_hD3psjHA` | Pestaña `BASE TRANSITOS` |
| 3 | Google Drive "Raw ventas Y.xlsx" / hoja `RAW` | `1K11y6icDm9M3X3glGUVCOe4HsbpWpEBm` | Pestaña `Raw` |

---

## Decisiones de diseño importantes

### STOCK
- Se actualizan solo las **columnas cuyos encabezados coinciden** entre Sheets y Excel
- Columna vacía inicial del Drive se excluye (no tiene encabezado)
- Pegado desde **fila 2** (justo bajo el encabezado)

### BASE TRANSITOS
- Misma lógica de columnas coincidentes
- Las columnas MES, STOCK ACTUAL, Tipo Categoria, Valor USD TOTAL, MARCA son **fórmulas propias** — no se tocan aunque existan en Google Sheets con el mismo nombre
- Protección implementada vía `TRANSITOS_FORMULA_COLS` en `build_column_mapping(skip_excel_names=...)` — activado cuando `data_start_row==14`
- Pegado desde **fila 14** — las filas 1-13 las usa Felipe para otra cosa y NO deben modificarse
- `parse_values=False` — los números se pegan **tal como vienen de Google Sheets** (sin convertir comas/puntos). Si se convirtieran, los valores se inflan porque el formato chileno los malinterpreta

#### Función `set_transitos_column_formats()` — se llama ANTES de escribir
Detecta 4 tipos de columnas y aplica formato previo para evitar conversiones automáticas de Excel:
| Tipo | Detección | Formato aplicado |
|---|---|---|
| Fechas | patrón `DD/MM/YY` o `DD/MM/YYYY` | `DD/MM/YY` |
| Moneda | patrón `$X,YYY` (ej. `$2,200`) | sin cambio (conserva `$#.##0,00` existente) |
| **Numérico entero** | patrón `^\d{1,3}(\.\d{3})*$` (ej. `"1.250"` o `"42"`) | `0` (número entero, sin decimales) |
| Resto | cualquier otra columna | `@` (texto puro, sin auto-conversión) |

**Columnas numéricas (SKU, cantidades)** se convierten a `int` en el path `parse_values=False`:
- `"1.250"` (miles chilenos) → `int(1250)` con formato `0` — las tablas dinámicas pueden agregar correctamente
- `"42"` (entero simple) → `int(42)` con formato `0`
- Decisión: si el campo tiene formato `@` (texto), la tabla dinámica no puede sumar ni agrupar por SKU numérico

#### Conversiones en el else branch (parse_values=False)
1. `.strip()` — elimina espacios sobrantes (ej. `"  100 "` → `"100"`)
2. `parse_sheet_dates_only()` — fechas `DD/MM/YY` o `DD/MM/YYYY` → `datetime.date`
3. **Moneda `$X,YYY`** — patrón `^\$(\d+),(\d+)$` → float: `"$2,200"` → `2.2`, `"$0,550"` → `0.55`
   - Razón: Excel interpreta `"$2,200"` como $2.200 (dos mil doscientos) si no se convierte a float
   - Con formato `$#.##0,00` muestra `$2,20` (correcto: dos dólares veinte centavos)
4. **Decimal con punto** (1-2 dígitos) — `"1.5"` → `"1,5"` (normaliza si Google Sheets devuelve punto)
   - NO afecta miles chilenos: `"1.234"` (3 dígitos) queda intacto

### Raw (ventas)
- El archivo Drive tiene 10 hojas; la correcta es **`RAW`** (mayúsculas), con 40 columnas y ~340k filas
- Las otras hojas (Black, Cyber, Hoja9, Fcst, Web, Resumen, Facturado, Full, Comisiones) no corresponden
- Copia completa: se limpian todas las filas desde fila 2 y se pegan todas las del Drive
- Los valores `datetime.time` de la columna "Hora Venta" se convierten a string `"HH:MM:SS"` porque xlwings no puede escribirlos como COM VARIANT

### Números desde Google Sheets
- Google Sheets usa **punto como separador de miles** (formato chileno): `"3.480"` = 3480
- La función `parse_sheet_val()` convierte automáticamente:
  - `"3.480"` → `3480` (int)
  - `"1.234.567"` → `1234567` (int)
  - `"1.234,56"` → `1234.56` (float)
  - `"3,14"` → `3.14` (float)
  - `"27/2/2026"` → `date(2026, 2, 27)`

---

## Tablas dinámicas

### TDs que dependen de Raw (se expanden cada semana)
| TD | Hoja | Nombre COM |
|---|---|---|
| TablaDinámica4 | NUEVO VENTA SKU | TablaDinámica4 |
| TablaDinámica5 | **TD VENTAS N** | TablaDinámica5 |
| TablaDinámica2 | **TD MG** | TablaDinámica2 |

- Las 3 comparten el mismo PivotCache (Cache [9])
- Se actualiza usando `PivotCaches().Create()` + `ChangePivotCache()`
- El primer `ChangePivotCache` exitoso propaga a las otras TDs del grupo (E_INVALIDARG en las siguientes = comportamiento normal, no es error real)
- **OBLIGATORIO cada actualización**: ampliar el rango de origen (SourceData) de las TDs **TD VENTAS N** y **TD MG** al nuevo número de filas de la hoja Raw — si no se hace, las TDs siguen apuntando al rango anterior y no reflejan los datos nuevos
- Este paso está implementado en el flujo como paso 10: `Expandir PivotCaches que apuntan a Raw`

### TD VENTAS N — orden manual
- Hoja: `TD VENTAS N`, pivot: `TablaDinámica5`
- Todos los campos tienen orden `-4135` (xlManual)
- El script **captura el orden antes** del `RefreshAll()` y lo **restaura después**
- Constantes: `VENTAS_N_SHEET = 'TD VENTAS N'`, `VENTAS_N_PIVOT = 'TablaDinámica5'`

---

## Flujo de ejecución
1. Descargar datos de las 3 fuentes Google
2. Abrir Excel con xlwings (visible=False)
3. **Limpiar filtros activos** en STOCK, BASE TRANSITOS y Raw (`clear_autofilter`) — si quedó un filtro activo, `used_range` devuelve rango incompleto y el write se desordena
4. Leer datos anteriores para comparación
5. Generar reporte de cambios (SKUs nuevos / eliminados / modificados)
6. Actualizar STOCK (columnas coincidentes, fila 2+, `parse_values=True`)
7. Actualizar BASE TRANSITOS: `set_transitos_column_formats()` (preformato) → `update_matched_columns()` (parse_values=False) → `update_transitos_mes_formula()` (fórmula col N)
8. Actualizar Raw (copia completa, fila 2+)
9. Actualizar VTA X marca meta desde TD VENTAS N
10. Expandir PivotCaches que apuntan a Raw al nuevo número de filas
11. Capturar orden de TD VENTAS N
12. Refrescar **cada TD individualmente** (síncrono) + `CalculateUntilAsyncQueriesDone()` — reemplaza `RefreshAll()` asíncrono que podía guardar antes de que el cambio de cache se confirmara
13. Restaurar orden de TD VENTAS N
14. Guardar y cerrar Excel (`wb.save()` + `wb.close()`)
15. `run_analisis(log_fn=log)` — genera `analisis_planificacion_ABR26.xlsx` con 8 hojas (críticos + sobrestock + inquietantes)

---

## VTA X marca meta — automatización VENTA ACUM.

### Lógica implementada (`update_vta_x_marca`)
- Lee **TD VENTAS N** (pivot por marca × mes): fila 4 = años, fila 5 = meses (números), datos desde fila 6
- 2026 empieza en col 16 (índice 15 base-0); meses ABR=4, MAY=5, JUN=6 en cols 16-19
- Construye dict `{marca_norm: {mes_num: valor}}`
- Lee **VTA X marca meta**: última tabla, cabeceras en fila 77
  - VENTA ACUM. ABR 26 → col 5 (E)
  - VENTA ACUM. MAY 26 → col 10 (J)
  - VENTA ACUM. JUN 26 → col 15 (O)
- Normalización de marca: `_norm_brand()` pasa a minúsculas y elimina sufijos `(...)` — ej. "UMA (Mattel)" → "uma"
- **"Otras Marcas" en TD VENTAS N = Proveedores Nacionales en su completitud** — se mapea a `'p. nacionales'` antes de buscar en VTA X marca meta
- Filas omitidas (tienen fórmulas propias): Total general (fila 87), P. Nacionales (fila 89), TOTAL EMPRESA (fila 91)
- **ORDEN CRÍTICO**: `update_vta_x_marca()` debe ejecutarse DESPUÉS de `expand_and_refresh_pivots()` — si va antes, lee datos viejos de TD VENTAS N
- Constantes: `VTA_SHEET`, `TD_VENTAS_N`, `VTA_ACUM_COLS = [5,10,15]`, `VTA_SKIP_LABELS`, `MES_ABBR`

### Columnas VENTA ACUM. (VTA X marca meta)
Se detectan automáticamente buscando la última fila que contenga "VENTA ACUM." en cualquier celda.
Usar `MES_ABBR` para mapear la abreviación del mes (ENE…DIC) al número de mes.

---

## BASE TRANSITOS — Fórmula columna N (MES)

```
=SI(M14="";"";SI(DIA(M14)<=10;MES(M14);SI(MES(M14)=12;1;MES(M14)+1)))
```
- Día ≤ 10 del mes → mes de la fecha ETA bodega (col M)
- Día > 10 → mes siguiente (diciembre → 1 = enero)
- Usa **punto y coma** (`;`) porque Excel está en español
- Se escribe con `ws.range(...).formula_local = ...` (no `.formula`)
- Constante: `TRANSITOS_MES_FORMULA` con placeholder `{fecha}`
- Se extiende automáticamente a todas las filas de datos; limpia sobrantes si la base achicó

---

## Tarea programada Windows
- Instalada vía PowerShell (`crear_tarea_programada.ps1`)
- Ejecuta cada **lunes a las 13:00**
- RunLevel: Limited (no requiere admin)
- Log en `C:\Users\felip\Desktop\UnionX Cloude\log.txt`

## Hook automático CLAUDE.md

- Script: `update_claudemd_hook.py`
- Se ejecuta en cada evento **Stop** de Claude Code (al terminar sesión)
- Configurado en `C:\Users\felip\.claude\settings.json`
- Actualiza la línea `Ultima actualizacion:` con fecha/hora actual
- Integra sección `## Pendiente` al cuerpo principal si existe

---

## Email semanal — Borrador Gmail

### Estructura del email (replicar cada lunes)
- **Asunto:** `Reporte Planificación y Transito Embarques Semana {N}`
- **Número de semana:** usar `datetime.date.today().isocalendar()[1]`
- **Destinatarios** (extraer del último enviado en Gmail #sent, buscar thread "planificaci"):
  `felipe@unionx.cl, nicolas@unionx.cl, martin@grupoeter.cl, andres@grupoeter.cl, sguzman@grupoeter.cl, nicole@grupoeter.cl, ignacia@melollevo.cl, claudia@melollevo.cl, michela@unionx.cl, soledad@grupoeter.cl, trinidad@melollevo.cl`
- **URL compose:** `https://mail.google.com/mail/u/0/?view=cm&fs=1&to={destinatarios}&su={asunto}&tf=1`

### Formato del cuerpo (estilo del reporte anterior)
1. Saludo: *"Buenas Tardes Equipo,"*
2. Por cada sección: **título en negrita** → **tabla** → **punteo debajo** (ul/li)
3. Secciones en orden:
   - **Venta MES AÑO** — tabla marcas vs lineal + punteo: total empresa, propias/nacionales, mejores/peores
   - **Forecast Inventarios** — tabla CST x Marca ABR+MAY + punteo: stock total empresa, meses cobertura, alertas
   - **Stock Crítico** — tabla críticos por marca + punteo: total SKUs, sin stock, próximas llegadas, inquietantes
   - **Sobrestock** — tabla top categorías + punteo: total SKUs/monto, focos principales
4. **Línea capital inmovilizado** (párrafo destacado, fondo morado claro `#f0e6ff`):
   - Mencionar stock total en sobrestock Y el excedente real sobre 4 meses óptimos
   - Valor exacto: leer título de hoja `Capital Inmovilizado` del analisis (`Total exceso sobre 4 meses óptimos: $XXX`)
   - Formato: *"El stock en sobrestock asciende a $XMM, sin embargo el excedente real sobre 4 meses óptimos corresponde a $YMM"*
5. Cierre: *"Quedo atento a cualquier consulta."* + Saludos Felipe

### Datos clave a leer de analisis_planificacion_ABR26.xlsx
| Dato | Fuente |
|---|---|
| Venta por marca, % vs Lineal | Hoja `VTA x Marca MES AÑO` |
| Stock crítico por marca | Hoja `Critico x Marca` |
| Sobrestock por marca/cat | Hoja `Sobrestock x Marca y Cat` |
| Stock+Cobertura ABR/MAY | Hoja `CST x Marca` |
| Capital inmovilizado ($) | Título de hoja `Capital Inmovilizado` (fila 1) |

### Inyección en Gmail
- Usar `trustedTypes.createPolicy('inject', {createHTML: s => s})` + `el.innerHTML = policy.createHTML(html)`
- El borrador se guarda automáticamente — no hace falta Ctrl+S explícito
- Adjuntar Excel manualmente antes de enviar

---

## partial_refresh.py — Re-run parcial sin descargar Google

Cuando las bases (STOCK, BASE TRANSITOS, Raw) ya están OK y solo se necesita:
- Refrescar pivots y expandir rangos
- Actualizar VTA X marca meta
- Regenerar REPORTE CST FLAT / REPORTE UNID FLAT
- Re-ejecutar análisis de planificación

```bash
cd C:\Users\felip\Desktop\UnionX Cloude
python partial_refresh.py
```

- No descarga nada de Google
- Exit code 0 = éxito total
- Usa las mismas funciones de `actualizar_reportes.py` + `format_td_reportes_cst.py`

---

## Backup antes de ejecución manual
```bash
cp "...FORECAST FINAL SKU 26-27.xlsx" "...FORECAST FINAL SKU 26-27_BACKUP_$(date +%Y%m%d_%H%M).xlsx"
```

---

## Errores conocidos y sus soluciones
| Error | Causa | Solución aplicada |
|---|---|---|
| `datetime.time` COM VARIANT | Columna "Hora Venta" en Raw | Convertir a string `HH:MM:SS` en `sanitize_raw_rows()` |
| Fechas invertidas | Excel interpreta DD/MM como MM/DD | Convertir a `datetime.date`; en BASE TRANSITOS usar `parse_sheet_dates_only()` |
| Números mal (3.480 → 3.48) | Punto = miles en Chile, Excel lo lee como decimal | `parse_sheet_val()` para STOCK; BASE TRANSITOS usa `parse_values=False` (no tocar) |
| Valores inflados en BASE TRANSITOS | `parse_sheet_val` interpretaba comas/puntos del formato chileno | `parse_values=False` en BASE TRANSITOS; solo se convierten fechas |
| `$2,200` se infla a $2.200 (×1000) | Excel lee `"$2,200"` como $2200 (miles US) en celda con formato `$#.##0,00` | Convertir a float: `"$2,200"` → `2.2`; conservar formato `$#.##0,00` existente |
| Valores con coma muestran punto | Celda tenía formato antiguo con punto decimal; Excel auto-convierte string | `set_transitos_column_formats()` fija formato `@` ANTES de escribir |
| Fechas en BASE TRANSITOS invertidas | Strings `"DD/MM/YY"` interpretados como `MM/DD/YY` por Excel | `parse_sheet_dates_only()` convierte a `datetime.date` + formato `DD/MM/YY` |
| Filtros activos en Excel | `used_range` incompleto, filas ocultas no se escriben | `clear_autofilter()` antes de leer y escribir |
| PivotTables no guardan nuevo origen | `RefreshAll()` asíncrono: save ocurría antes de confirmar | Refresh individual síncrono + `CalculateUntilAsyncQueriesDone()` |
| openpyxl pivot cache error | `CalculatedItem.formula` NoneType | Usar xlwings en vez de openpyxl para el Excel principal |
| PivotCache SourceData error 1004 | No se puede asignar string directamente | Usar `PivotCaches().Create()` + `ChangePivotCache()` |
| TablaDinámica5/TablaDinámica2 no se actualizan (ADVERTENCIA en log) | `ChangePivotCache()` falla con E_INVALIDARG porque tienen cache propio no compatible con el nuevo cache COM | Fallback: `pt.PivotCache().SourceData = build_new_source(src)` — edita el cache existente en vez de reemplazarlo |
| Tarea programada acceso denegado | RunLevel Highest requiere admin | Usar RunLevel Limited |
| `app.quit()` — RPC server not available | Excel ya se había cerrado antes de que se llamara `app.quit()` en el `finally` | **Benigno** — todo ya fue guardado correctamente; ignorar este error |
| VTA X marca meta no se actualiza (lee datos viejos) | `update_vta_x_marca()` se ejecutaba ANTES del refresh de pivots; leía TD VENTAS N sin refrescar | Mover `update_vta_x_marca()` para DESPUÉS de `expand_and_refresh_pivots()` en `main()` |
| "Otras Marcas" en TD VENTAS N no se escribe en P. Nacionales | La fila "Otras Marcas" del pivot se saltaba en vez de mapearse | En `update_vta_x_marca()`: si `marca_n == 'otras marcas'` → usar `'p. nacionales'` como clave |
| SKU aparece como NUEVO cuando en Excel tiene otra categoría | Google Sheets "Importaciones" tiene columna "Tipo Categoria" que sobreescribía la fórmula Excel | `TRANSITOS_FORMULA_COLS` en `build_column_mapping()` excluye: MES, STOCK ACTUAL, Tipo Categoria, Valor USD TOTAL, MARCA cuando `data_start_row==14` |

---

## Reportes generados por format_td_reportes_cst.py

### Entry point
`create_reporte_cst_formato(wb_xw)` — llamado desde `actualizar_reportes.py` cada lunes.
Internamente llama a `create_reporte_cst_flat(wb_xw)` y `create_reporte_unid_flat(wb_xw)`.
Elimina las hojas antiguas `REPORTE CST FORMATO` y `REPORTE CST PIVOT` si existen.

### Hojas generadas
| Hoja | Cols | Filas de datos |
|---|---|---|
| `REPORTE CST FLAT` | 44 (A–AR) | 816 (espejo de FCST BASE SKU MACRO) |
| `REPORTE UNID FLAT` | 47 (A–AU) | 816 |

### REPORTE CST FLAT — layout de columnas
| Cols | Contenido |
|---|---|
| A | (índice) |
| B–H | Marca, Cat Padre, Cat Hijo, SKU, Descripción, Cat Comercial, Cat Comercial2 |
| (col fija) | StockHoyCst, Cobert.ACT26 |
| I–L (ABR, 4 cols) | Llegadas, Stk+Ped, VentaCst, **Cobert** |
| M–R (MAY, 6 cols) | StkIni, Compra, Llegadas, Stk+Ped, VentaCst, **Cobert** |
| S–X (JUN, 6 cols) | ídem |
| Y–AD (JUL, 6 cols) | ídem |
| AE–AJ (AGO, 6 cols) | ídem |
| AK–AN (SEP, 4 cols) | StkIni, Llegadas, Stk+Ped, VentaCst (**sin Cobert, sin Compra**) |
| AO–AR (OCT, 4 cols) | ídem |

### REPORTE UNID FLAT — layout de columnas
| Cols | Contenido |
|---|---|
| A | (índice) |
| B–H | misma jerarquía que CST |
| (col fija) | StockHoyUnid, Cobert.ACT26 |
| I–M (ABR, 5 cols) | Compra, Llegadas, Stk+PedUnid, VentaPPTO, **Cobert** |
| N–S (MAY, 6 cols) | StkIniUnid, Compra, Llegadas, Stk+PedUnid, VentaPPTO, **Cobert** |
| T–Y (JUN, 6 cols) | ídem |
| Z–AE (JUL, 6 cols) | ídem |
| AF–AK (AGO, 6 cols) | ídem |
| AL–AP (SEP, 5 cols) | StkIniUnid, Compra, Llegadas, Stk+PedUnid, VentaPPTO (**sin Cobert**) |
| AQ–AU (OCT, 5 cols) | ídem |
| ABR Stock Inicial Unid | = Stock Hoy Unid (col L de FCST BASE) — no existe col StkIni ABR separada |

### Fórmula Cobertura (rolling 3 meses)
```
=IFERROR(IF((V1+V2+V3)=0,"",SP/((V1+V2+V3)/3),"")
```
- `SP` = Stk+Ped (o Stk+PedUnid) del mes en cuestión
- `V1/V2/V3` = VentaCst (o VentaPPTO) del mes actual + 2 meses siguientes
- Solo se calcula para **ABR–AGO** (últimos meses con 3 meses completos hacia adelante)
- SEP y OCT **no tienen columna Cobert**

### Cobert.ACT26
```
StockHoyCst / ((VentaCst ABR + VentaCst MAY + VentaCst JUN) / 3)
```
(Para UNID: StockHoyUnid / promedio VentaPPTO ABR+MAY+JUN)

### Fuente de datos
- Todas las fórmulas apuntan a **FCST BASE SKU MACRO** (213 cols reales, max col HE, headers fila 3 — el `used_range` puede reportar hasta ~15k cols por artefactos de formato vacío, pero los datos reales terminan en HE=213)
- Los encabezados pueden tener espacios iniciales (ej. `' Venta PPTO ABR26'`) — **no usar `.strip()`**
- Columnas duplicadas en FCST BASE (ej. 'Llegadas Cst AGO' en 2 cols) no afectan porque se referencian por letra de columna, no por nombre
- **Columnas de tránsito FCST BASE V2** (0-based): `CW`(100)=Transito ABR26, `DJ`(113)=Embarcado Mayo, `DV`(125)=Transito JUN26, `EH`(137)=JUL, `ET`(149)=AGO, `FF`(161)=SEP, `FR`(173)=OCT — NO usar columnas "PEDIDO X LLEGADAS Y" (índice −1, son compra, no tránsito real). FCST BASE V2 tiene 213 cols reales (max col HE).
- Stock Hoy CST en V2: col **O** (índice 14); Stock Hoy Unid: col **K** (índice 10)

### Mapeo correcto de bloques CST en format_td_reportes_cst.py
Cada mes tiene patrón: [UNID cols] [CST cols] [Stk+Ped Unid | Stk+Ped Cst] [Venta Neta]
**FCST BASE V2 — max col HE (213). Columnas más allá de HE NO existen.**
| Bloque | Constante | Columnas (llegadas, stk_ped_cst, venta_cst) |
|---|---|---|
| ABR_COLS | CST ABR | Llegadas=DA, Stk+Ped=DC, Venta=CY |
| MAY26 MAYO_AGO | CST MAY | StkIni=DK, Compra=DM, Llegadas=DN, Stk+Ped=DP, Venta=DL |
| JUN26 | CST JUN | StkIni=DW, Compra=DY, Llegadas=DZ, Stk+Ped=EB, Venta=DX |
| JUL26 | CST JUL | StkIni=EI, Compra=EK, Llegadas=EL, Stk+Ped=EN, Venta=EJ |
| AGO26 | CST AGO | StkIni=EU, Compra=EW, Llegadas=EX, Stk+Ped=EZ, Venta=EV |
| SEP26 SEP_OCT | CST SEP | StkIni=FG, Llegadas=FJ, Stk+Ped=FL, Venta=FH |
| OCT26 | CST OCT | StkIni=FS, Llegadas=FU, Stk+Ped=FW, Venta=FT |

### Mapeo correcto de bloques UNID en format_td_reportes_cst.py
| Bloque | Constante | Columnas (ppto, compra/pedido, llegadas/transito, stk_ped_unid) |
|---|---|---|
| UNID_ABR | UNID ABR | Ppto=CU, Pedido=CV, Transito=CW, Stk+Ped=DB |
| MAY26 UNID_MAY_AGO | UNID MAY | StkIni=DF, Ppto=DG, Pedido=DH, Embarcado=DJ, Stk+Ped=DO |
| JUN26 | UNID JUN | StkIni=DS, Ppto=DT, Pedido=DU, Transito=DV, Stk+Ped=EA |
| JUL26 | UNID JUL | StkIni=EE, Ppto=EF, Pedido=EG, Transito=EH, Stk+Ped=EM |
| AGO26 | UNID AGO | StkIni=EQ, Ppto=ER, Pedido=ES, Transito=ET, Stk+Ped=EY |
| SEP26 UNID_SEP_OCT | UNID SEP | StkIni=FC, Ppto=FD, Pedido=FE, Transito=FF, Stk+Ped=FK |
| OCT26 | UNID OCT | StkIni=FO, Ppto=FP, Pedido=FQ, Transito=FR, Stk+Ped=FV |

---

## Análisis de planificación — analisis_stock_critico.py

### Archivo
`C:\Users\felip\Desktop\UnionX Cloude\analisis_stock_critico.py`

### Entry point
`run_analisis(log_fn=print)` — llamado automáticamente desde `actualizar_reportes.py` cada lunes, después de guardar y cerrar el Excel.

### Salida
`C:\Users\felip\Desktop\UNIONX\FORECAST FINAL SKU\Analisis Planificacion\analisis_planificacion_ABR26.xlsx`

### Hojas generadas (se regeneran automáticamente cada lunes)
| Hoja | Contenido | Color |
|---|---|---|
| `Critico x Marca` | Resumen solo por Marca, cobertura < 1 mes | Rojo |
| `Critico x Marca y Cat` | Resumen por Marca × CatComercial, cobertura < 1 mes | Rojo |
| `Detalle Critico` | SKUs con cobertura < 1 mes, ordenados por urgencia — **22 columnas** | Rojo |
| `Sobrestock x Marca` | Resumen solo por Marca, cobertura > 6 meses | Morado |
| `Sobrestock x Marca y Cat` | Resumen por Marca × CatComercial, cobertura > 6 meses | Morado |
| `Detalle Sobrestock` | SKUs con cobertura > 6 meses, ordenados de mayor a menor cobertura — **21 columnas** | Morado |
| `Inquietante x Cat Padre` | Resumen por **Marca × Cat. Padre**, cobertura < 2 meses | Amarillo |
| `Inquietante x Cat Hijo` | Resumen por **Marca × Cat. Hijo**, cobertura < 2 meses | Amarillo |
| `Tránsitos por Embarque` | UNIFICADA — resumen por PI con detalle SKU expandible 3 niveles, cobertura proyectada al mes anterior a la llegada | Azul marino |
| `Nuevos en Tránsito` | SKUs con Tipo Categoría = NUEVO, agrupados por mes de llegada | Verde teal |
| `VTA x Marca MES AÑO` | Linealidad vs Meta del mes en curso por marca — 7 cols, % vs Lineal como métrica principal (ej. `VTA x Marca ABR 26`) | Negro/dinámico |
| `CST x Marca` | Resumen de costos por Marca con cobertura y flujo ABR→SEP — 37 cols en $M CLP, coberturas coloreadas | Negro/multicolor |
| `Capital Inmovilizado` | Sobrestock jerárquico desplegable: Marca → Cat.Padre → Cat.Hijo → SKU, capital excedente sobre 4 meses óptimo | Morado |
| `Sobrestock c-Llegada` | SKUs sobrestock con embarques encima | Morado |
| `Fecha de Quiebre` | Críticos con mes proyectado de quiebre (proyección mes a mes con Stk+Ped FLAT) + estado OC | Rojo |

### Hojas de Tránsito — `write_transitos_embarque()` y `write_nuevos_transito()`
Generadas al final de `run_analisis()` a partir de `trans_rows_full`.

**Tránsitos por Embarque** — UNIFICADA (antes eran 2 hojas separadas) — 11 columnas, una fila por PI + filas SKU colapsadas:
- Columnas: PI | ETA Chile/Desc | ETA Bodega | Mes | Marcas | SKUs | Críticos<1m | Inquiet.1-2m | Unidades | Valor USD | Nivel Riesgo
- Cobertura mostrada: **proyectada al mes ANTERIOR al de llegada** (ej. llega JUL → muestra cobertura JUN)
  - Razón: el stock tiene que aguantar hasta que llegue el embarque; el mes anterior es el cuello de botella
- `cob_map_full`: dict `{sku: {'tipo': 'cobert'|'sin_venta', 'actual': val, 4: cob_abr, 5: cob_may, 6: cob_jun, 7: cob_jul, 8: cob_ago}}`
  - `tipo='sin_venta'` solo si NO hay ventas en NINGÚN mes ABR–AGO (no solo ABR)
  - SKUs descontinuados muestran `—` sin alerta
  - SKUs en tránsito no encontrados en CST FLAT (no descontinuados) → ALERTA en log
- Filas de detalle SKU: `outline_level=1, hidden=True` — se expanden con `[+]` junto a cada fila PI
- Ordenamiento dentro del PI: críticos (`cob<1`) primero, luego inquietantes, luego resto
- `cob=0` con ventas → `'Sin stock'` en rojo (más crítico que <1m)
- Botones `1`/`2` arriba a la izquierda colapsan/expanden todos los grupos a la vez
- `ws.sheet_properties = WorksheetProperties(outlinePr=Outline(summaryBelow=False))`
- SKUs inválidos (con espacios, `\n`, o >25 chars) se excluyen de filas de detalle

**VTA x Marca MES AÑO** — hoja de cumplimiento mensual con **linealidad** como métrica principal:
- Nombre de hoja se genera automáticamente según el mes en curso: `f'VTA x Marca {MES_ABBR} {YY}'` (ej. `VTA x Marca ABR 26`)
- Fuente: hoja `VTA X marca meta` del Excel principal — la función `update_vta_x_marca()` en `actualizar_reportes.py` ya actualiza la columna `VENTA ACUM.` antes de llamar a `run_analisis()`
- Detecta dinámicamente las columnas PPTO, VENTA ACUM., Cumpl. de la última tabla que tenga "VENTA ACUM." en la cabecera

**Linealidad (lógica central):**
- `ayer = today - timedelta(days=1)` → `dia_ayer = ayer.day`
- `total_dias = calendar.monthrange(año, mes_num)[1]`
- `linealidad = dia_ayer / total_dias` (ej. día 20 de 30 = 66.7%)
- `meta_lineal = ppto * linealidad` — lo que se esperaría haber vendido hasta ayer a ritmo constante

**7 columnas:**
| Col | Nombre | Descripción |
|---|---|---|
| 1 | Marca | Nombre de marca |
| 2 | Meta Total MES AÑO | Presupuesto completo del mes (PPTO) |
| 3 | Meta al día XX (NN%) | Meta Total × linealidad — objetivo proporcional al día de ayer |
| 4 | Venta Acum. MES AÑO | Venta acumulada real |
| 5 | vs Lineal ($) | Acum - Meta Lineal (verde si positivo, rojo si negativo) |
| 6 | **% vs Lineal** | ⭐ Métrica principal — Acum / Meta Lineal — **con color** |
| 7 | % vs Meta | Acum / Meta Total — referencia secundaria (gris neutro) |

- Ordenado por % vs Lineal descendente (mejor arriba)
- Colores en % vs Lineal: 🟣 morado=≥110%, 🟢 verde=90-110%, 🟡 amarillo=70-90%, 🟠 naranja=50-70%, 🔴 rojo=<50%
- Título incluye: `Linealidad: día {dia_ayer} de {total_dias} = {linealidad:.1%}`
- Subtítulo explicativo en fila 2 (gris, itálica)
- Freeze panes en A4 (encabezados en fila 3, datos desde fila 4)
- Filas omitidas (VTA_SKIP): `total general`, `total empresa`, `p. nacionales`
- Fila TOTAL PROPIA al final (fondo negro)
- `import calendar` requerido (ya agregado al módulo)

**Nuevos en Tránsito** — 6 columnas, filtrado a `tipo_cat` que contenga "nuevo":
- Columnas: SKU | Descripción | Marca | Mes Llegada | Fecha ETA Bodega | Cantidad
- Separadores visuales por mes (fila verde oscuro con `▶ Llegada MES`)
- Ordenado por mes → ETA Bodega → SKU

**CST x Marca** — hoja de costos por marca con flujo mensual ABR→SEP:
- Fuente: hoja `TD REPORTES CST MP+N SKU` del Excel principal (736 filas × 46 cols)
- Función: `write_cst_mp_n_sku(owb, data_mp)`; datos leídos en `run_analisis` → `data_mp`
- **37 columnas**: Marca | Stock Hoy CST | Cobert.ACT | [ABR: Tránsito, Stk+Ped, Venta, Cobert] | [MAY-SEP: StkIni, Compra, Llegadas, Stk+Ped, Venta, Cobert] × 5
- Valores numéricos en **$M CLP** (formato `$ #,##0.0,,`)
- Fila 1: título negro; Fila 2: encabezados de grupo por mes (colores distintos por mes); Fila 3: sub-encabezados; freeze panes en B5
- Coberturas coloreadas: 🔴<1m | 🟠 1-2m | 🟢 2-4m | 🔵 4-6m | 🟣>6m
- Filas de totales al final en fondo negro: TOTAL PROPIA / PROV. NACIONALES / TOTAL EMPRESA
- Filas omitidas (SKIP): `total general`, `total empresa`, `proveedores nacionales` → se muestran como totales
- Nota al pie con leyenda de colores
- Colores de grupo por mes: ABR=azul oscuro, MAY=verde, JUN=morado, JUL=naranja, AGO=azul acero, SEP=gris

**Descubrimiento dinámico de columnas BASE TRANSITOS:**
- Se lee fila 1 del sheet (`_trans_hdrs`) para mapear nombres → índices
- `_tcol(nombre, fallback)` — usa encabezado si existe, sino fallback hardcodeado
- Fallbacks confirmados: SKU=0, Desc=1, PI=2, Cantidad=6, ETA Bodega=12, MES=13, Stock Actual=14, Tipo Cat=15, Valor USD=16, MARCA=17
- Solo se incluyen en `trans_rows_full` las filas que tienen PI o ETA Bodega no vacíos

### Layout Detalle Critico (22 columnas)
| Col | Contenido | Fuente |
|---|---|---|
| 1 | Marca | REPORTE CST FLAT col A |
| 2 | Cat. Comercial | REPORTE CST FLAT col D |
| 3 | Cat. Padre | REPORTE CST FLAT col B |
| 4 | Cat. Hijo | REPORTE CST FLAT col C |
| 5 | Ranking Comercial | FCST BASE col J (índice 9) |
| 6 | SKU | REPORTE CST FLAT col E |
| 7 | Descripcion | REPORTE CST FLAT col F |
| 8 | Cobert. ACT 26 | REPORTE CST FLAT col H |
| 9 | Stock Unid | REPORTE UNID FLAT |
| 10 | Venta PPTO ABR Unid | REPORTE UNID FLAT |
| 11 | Stock CST | REPORTE CST FLAT |
| 12 | Venta CST ABR | REPORTE CST FLAT |
| 13–18 | Llegadas ABR–SEP Unid | FCST BASE transit cols; si 0 → usa qty de BASE TRANSITOS |
| 19 | Prox. Llegada | BASE TRANSITOS (override) o primer mes válido con llegadas > 0 |
| 20 | PI Embarque | BASE TRANSITOS |
| 21 | ETA Bodega | BASE TRANSITOS |
| **22** | **Fecha Máx. de Carga** | Calculada — solo si SKU **sin OC** en BASE TRANSITOS |

### Layout Detalle Sobrestock (21 columnas)
Igual que Detalle Critico pero **sin col 22** (Fecha Máx. de Carga).

### Filtros aplicados
- Excluir categorías vacías, `(sin clasificar)` y que contengan `descontinua`
- Excluir filas cuya Marca empiece con `Total ` (subtotales del FCST BASE)

### Fuentes de datos y lógica
- **Cobertura**: REPORTE CST FLAT col H (Cobert. ACT 26)
- **Stock/Venta CST**: REPORTE CST FLAT cols G, K
- **Stock/Venta UNID**: REPORTE UNID FLAT cols G, L
- **Cat. Padre / Cat. Hijo**: REPORTE CST FLAT cols B y C (índices 1 y 2)
- **Ranking Comercial**: FCST BASE SKU MACRO col J (índice 9)
- **Puerto Origen**: FCST BASE col S (índice 18) — usado para calcular Fecha Máx. de Carga
- **Llegadas por mes (unidades)**: FCST BASE columnas de tránsito (ver tabla abajo). Si FCST BASE tiene 0 → se usa la qty de BASE TRANSITOS para ese mes
- **PI y ETA bodega**: BASE TRANSITOS (fuente de verdad para mes real de llegada)
- Si BASE TRANSITOS tiene mes distinto al FCST BASE → prevalece BASE TRANSITOS para `Próx. Llegada`, PI y ETA

### Columnas tránsito FCST BASE — V2 (0-based) — TRANSITO_FCST
**FCST BASE V2 max col = HE (213 cols reales). Índices 0-based correctos:**
| Mes | Idx | Col | Nombre header |
|---|---|---|---|
| ABR26 | 100 | CW | Transito ABR 26 |
| MAY26 | 113 | DJ | Embarcado Mayo |
| JUN26 | 125 | DV | Transito JUN 26 |
| JUL26 | 137 | EH | Transito JUL 26 |
| AGO26 | 149 | ET | Transito AGO 26 |
| SEP26 | 161 | FF | Transito SEP 26 |
| OCT26 | 173 | FR | Transito OCT 26 |

**IMPORTANTE**: La columna inmediatamente anterior a cada una es "PEDIDO X\nLLEGADAS Y" — compra/pedido, NO tránsito real. No usar esas.
**data_fcst read range**: `ws_fcst.range((4,1),(lrf,213)).value` — usar 213, NO 433.

### Lógica de meses válidos — `_is_valid_prox(mes_str)`
Un mes es válido como próxima llegada solo si su **día 10** (cutoff según regla BASE TRANSITOS) aún no ha pasado.
- Ejemplo hoy=20-ABR: ABR26 cutoff=10-ABR < hoy → inválido → se omite
- Aplica tanto al cálculo de `prox` desde FCST BASE como al override de BASE TRANSITOS
- Garantiza que nunca aparezca un mes cuyas llegadas ya debieron haberse recibido

### Fecha Máx. de Carga — `fecha_max_carga(prox, puerto)`
- **Solo se muestra en Detalle Critico** (col 22), **solo si el SKU no tiene OC** en BASE TRANSITOS (`has_emb=False`)
- Lógica: último día del mes anterior al mes de llegada − días de tránsito según puerto
  - **Shenzhen** → 70 días
  - **Ningbo** → 55 días
- Ejemplo: prox=JUL26, Ningbo → 30/06/2026 − 55 días = **06/05/2026**
- Color: 🔴 rojo si ya pasó | 🟠 naranja si quedan ≤7 días | 🟢 verde si hay margen

### Hojas Inquietantes — `write_resumen_cat()`
- Agrupan por **Marca × Categoría** (cat_padre o cat_hijo) — 15 columnas
- `cob_prom`: promedio de TODOS los SKUs de esa (marca × cat), excluyendo sobrestock (>6m) para evitar distorsión bimodal
- Filtro: solo aparecen filas con `cob_prom < 2` — excluye categorías donde el problema es aislado
- Ordenadas por `cob_prom` ascendente (más urgente primero)
- Columnas: Marca | Cat | SKUs | Cob.Prom | SKUs<0.5 | SKUs0.5-1 | SKUs1-2 | StkUnid | VentaUnid | StkCST | VentaCST | LlegABR | SinEmb | ConLleg | DetalleMeses

### Lógica de colores
- **Críticos**: rojo <0.5 mes, naranja 0.5–1, verde=llega próximo mes, naranja=llegada futura, rojo=sin embarque
- **Sobrestock**: naranja 6–9 meses, rojo 9–12, morado >12; verde=sin embarque (bueno), rojo=llega próximo (peor)
- **Inquietantes**: rojo <0.5, naranja 0.5–1, amarillo 1–2

### Columnas diferenciales resumen
| Columna | Crítico | Sobrestock |
|---|---|---|
| Col 5 | SKUs sin Stock (rojo) | Con Llegadas — más mercadería encima (naranja) |
| Col 11 | Sin Embarque (rojo) | Sin Llegadas (verde — bueno) |

---

### Capital Inmovilizado — `write_capital_inmovilizado()`
Jerarquía desplegable de 3 niveles con openpyxl outline grouping:
- **Nivel 0** (siempre visible): Marca — fondo morado oscuro `4A235A`, totales de capital/stock/venta
- **Nivel 1** (colapsado): Cat. Padre — fondo morado medio `7D3C98`
- **Nivel 2** (colapsado): Cat. Hijo — fondo morado claro `D7BDE2`
- **Nivel 3** (colapsado): SKU — alternado blanco/`F4ECF7`, cobertura coloreada individual
- Botones `1` `2` `3` `4` en Excel colapsan/expanden todos los niveles
- `OPTIMO_MESES = 4` — capital excedente = stock - (venta × 4)
- Ordenado por capital_exceso descendente en cada nivel
- Columnas (10): Marca/Cat/SKU | Descripción | SKUs | Cobert.ACT | Meses Exceso | Stock CST | Venta CST | Stock Óptimo | Capital Inmovil. | Tiene Llegadas

---

### Fecha de Quiebre — `write_fecha_quiebre()`
Proyecta en qué **mes** se agota el stock para cada SKU crítico.

**Lógica de proyección (`_proyectar_quiebre_mes`):**
- **Stock = 0** → `⛔ Quebrado` (en col Mes de Quiebre)
- **ABR** (mes actual): `stock_hoy < VentaABR × (días_restantes / días_totales)` → quiebra ABR
- **MAY–AGO**: usa `Stk+Ped` (rc[15], rc[21], rc[27], rc[33]) = StkIni + Llegadas confirmadas (tránsitos)
  - Si `Stk+Ped <= 0` → quiebra ese mes
  - Si `Stk+Ped < VentaMes` → quiebra ese mes
  - Si no quiebra en ningún mes → `∞ sin quiebre`
- **IMPORTANTE**: usar `Stk+Ped` no `StkIni` — la columna de tránsito/llegadas ya está incorporada

**Columnas (12):** Marca | Cat.Comercial | Cat.Padre | Cat.Hijo | SKU | Descripción | Cobert.ACT | Stock CST | Venta CST ABR | **Mes de Quiebre** | ETA OC | **Estado**

**Colores Mes de Quiebre:** `⛔ Quebrado`=rojo oscuro | ABR=rojo | MAY=naranja | JUN=amarillo | JUL=azul claro | AGO=verde claro | `∞`=verde

**Estado (col 12):**
- `✔ OC a tiempo — ETA DD/MM` — ETA ≤ primer día del mes de quiebre
- `⚠ Quiebre MES — OC DD/MM` — OC llega después del inicio del mes de quiebre
- `⚠ Quebrado — OC ETA DD/MM` — ya sin stock pero hay OC
- `✘ Sin OC — llegada proy. JUN26` — sin OC en BASE TRANSITOS pero FCST BASE tiene llegada proyectada
- `✘ Sin OC abierta` — sin OC y sin proyección en horizonte

**Ordenamiento:** quebrados primero, luego por mes de quiebre ascendente (ABR→AGO→sin quiebre)

**Índices REPORTE CST FLAT usados:**
| Campo | Índice (0-based) |
|---|---|
| Stk+Ped ABR | rc[9] |
| Venta ABR | rc[10] |
| Stk+Ped MAY | rc[15] |
| Venta MAY | rc[16] |
| Stk+Ped JUN | rc[21] |
| Venta JUN | rc[22] |
| Stk+Ped JUL | rc[27] |
| Venta JUL | rc[28] |
| Stk+Ped AGO | rc[33] |
| Venta AGO | rc[34] |
