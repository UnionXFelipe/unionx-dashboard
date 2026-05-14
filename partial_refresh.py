"""
partial_refresh.py
Refresca pivots, regenera REPORTE CST FLAT / REPORTE UNID FLAT
y re-ejecuta el análisis de planificación.
NO descarga datos de Google — las hojas STOCK, BASE TRANSITOS y Raw ya deben estar OK.
"""
import sys
import os
import datetime

# ── Logging ───────────────────────────────────────────────────────────────────
LOG_FILE = os.path.join(os.path.dirname(__file__), 'log.txt')

def log(msg, indent=0):
    ts  = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    pfx = '   ' * indent
    line = f"[{ts}] {pfx}{msg}"
    print(line, flush=True)
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(line + '\n')

# ── Importar funciones del script principal ────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from actualizar_reportes import (
    EXCEL_PATH, EXCEL_TAB_RAW,
    expand_and_refresh_pivots,
    update_vta_x_marca,
)
from format_td_reportes_cst import create_reporte_cst_formato

import xlwings as xw

def main():
    log('=' * 60)
    log(f'Partial refresh — {datetime.datetime.now().strftime("%A %d/%m/%Y %H:%M")}')

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        log('Abriendo Excel...')
        wb = app.books.open(EXCEL_PATH)

        # Contar filas actuales en Raw para saber el rango correcto
        ws_raw = wb.sheets[EXCEL_TAB_RAW]
        raw_last_row = ws_raw.used_range.last_cell.row - 1   # sin header
        raw_last_col = ws_raw.used_range.last_cell.column

        log(f'Raw: {raw_last_row} filas × {raw_last_col} columnas')

        # 1. Expandir PivotCaches + refrescar todas las TDs
        log('Expandiendo y refrescando tablas dinámicas...')
        n = expand_and_refresh_pivots(wb, EXCEL_TAB_RAW, raw_last_row, raw_last_col)
        log(f'   {n} tabla(s) de Raw expandidas.', 1)

        # 2. Actualizar VTA X marca meta DESPUÉS del refresh
        log("Actualizando 'VTA X marca meta'...")
        update_vta_x_marca(wb)

        # 3. Regenerar REPORTE CST FLAT y REPORTE UNID FLAT
        log("Regenerando 'REPORTE CST FLAT' y 'REPORTE UNID FLAT'...")
        create_reporte_cst_formato(wb)

        # 4. Forzar recálculo completo antes de guardar
        log('Recalculando fórmulas antes de guardar...')
        try:
            app.api.CalculateFull()
        except Exception:
            try:
                app.api.CalculateUntilAsyncQueriesDone()
            except Exception:
                app.calculate()

        # 5. Guardar y cerrar
        log('Guardando Excel...')
        wb.save()
        wb.close()
        log('Excel guardado y cerrado.')

    except Exception as e:
        log(f'ERROR: {e}')
        import traceback
        log(traceback.format_exc())
        sys.exit(1)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    # 5. Re-ejecutar análisis de planificación
    log('Generando análisis de planificación...')
    try:
        from analisis_stock_critico import run_analisis
        run_analisis(log_fn=log)
    except Exception as e:
        log(f'   ADVERTENCIA análisis: {e}')
        import traceback
        log(traceback.format_exc())

    log('✅ Partial refresh completado.')
    log('=' * 60)

if __name__ == '__main__':
    main()
