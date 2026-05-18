# -*- coding: utf-8 -*-
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

FOLDER_ID   = "1UXbSH9EytHJWLkao1bKJ2vjAXIzpHoUj"
ANALISIS    = r"C:\Users\felip\Desktop\UNIONX\FORECAST FINAL SKU\Analisis Planificacion\analisis_planificacion_MAY26.xlsx"
METAS       = r"C:\Users\felip\Desktop\UNIONX\PPTO 2026\Metas oficiales 1SEM Nuevo.xlsx"
THIS_DIR    = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(THIS_DIR, "drive_config.py")

def main():
    from drive_utils import get_service, upload_or_update, find_file

    print("Conectando con Google Drive...")
    svc = get_service(rw=True)

    # Upload analisis
    fname = os.path.basename(ANALISIS)
    existing = find_file(svc, FOLDER_ID, fname)
    analisis_id = upload_or_update(svc, ANALISIS, fname, FOLDER_ID, existing)
    print(f"OK Analisis: {fname} -> {analisis_id}")

    # Upload Metas
    metas_name = os.path.basename(METAS)
    existing_m = find_file(svc, FOLDER_ID, metas_name)
    metas_id = upload_or_update(svc, METAS, metas_name, FOLDER_ID, existing_m)
    print(f"OK Metas: {metas_name} -> {metas_id}")

    # Write drive_config.py
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        f.write('"""drive_config.py -- Generado automaticamente (NO subir a git)"""\n')
        f.write(f'DRIVE_FOLDER_ID = "{FOLDER_ID}"\n')
        f.write(f'METAS_FILE_ID   = "{metas_id}"\n')
    print(f"\nOK drive_config.py guardado en: {CONFIG_PATH}")

    print("\n" + "=" * 60)
    print("COPIA ESTO EN Streamlit Cloud -> App settings -> Secrets:\n")
    print("[dashboard]")
    print(f'drive_folder_id = "{FOLDER_ID}"')
    print(f'metas_file_id   = "{metas_id}"')
    print()
    print("[google_credentials]")
    print("# Pega el contenido de credentials.json como pares clave=valor")
    print("=" * 60)

if __name__ == "__main__":
    main()
