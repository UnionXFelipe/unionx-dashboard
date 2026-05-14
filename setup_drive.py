# -*- coding: utf-8 -*-
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
"""
setup_drive.py — Ejecutar UNA SOLA VEZ para:
  1. Crear carpeta "UnionX Dashboard" en Google Drive
  2. Subir análisis más reciente y Metas oficiales
  3. Compartir la carpeta con los emails indicados
  4. Generar drive_config.py con los IDs (para uso local)
  5. Imprimir el bloque a copiar en Streamlit Cloud secrets

Uso:
  python setup_drive.py
  python setup_drive.py felipe@unionx.cl nicolas@unionx.cl
"""
import sys, glob, os

ANALISIS_DIR = (
    r"C:\Users\felip\Desktop\UNIONX\FORECAST FINAL SKU\Analisis Planificacion"
)
METAS_PATH = (
    r"C:\Users\felip\Desktop\UNIONX\PPTO 2026\Metas oficiales 1SEM Nuevo.xlsx"
)
FOLDER_NAME  = "UnionX Dashboard"
THIS_DIR     = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH  = os.path.join(THIS_DIR, "drive_config.py")
SHARE_EMAILS = sys.argv[1:]   # emails opcionales como argumentos


def main():
    from drive_utils import get_service, upload_or_update, share_folder

    print("Conectando con Google Drive...")
    svc = get_service(rw=True)

    # ── 1. Crear carpeta ──────────────────────────────────────────────────────
    folder = svc.files().create(
        body={
            "name": FOLDER_NAME,
            "mimeType": "application/vnd.google-apps.folder",
        },
        fields="id,webViewLink",
    ).execute()
    folder_id = folder["id"]
    print(f"✅ Carpeta creada: '{FOLDER_NAME}'")
    print(f"   ID  : {folder_id}")
    print(f"   URL : {folder.get('webViewLink', '—')}")

    # ── 2. Compartir con emails ───────────────────────────────────────────────
    for email in SHARE_EMAILS:
        share_folder(svc, folder_id, email, role="reader")
        print(f"   Compartido con: {email}")

    # ── 3. Subir análisis más reciente ────────────────────────────────────────
    analisis_id = None
    analisis_files = sorted(
        glob.glob(os.path.join(ANALISIS_DIR, "analisis_planificacion_*.xlsx"))
    )
    if analisis_files:
        latest = analisis_files[-1]
        fname  = os.path.basename(latest)
        analisis_id = upload_or_update(svc, latest, fname, folder_id)
        print(f"✅ Análisis subido : {fname}")
        print(f"   ID: {analisis_id}")
    else:
        print("⚠️  No se encontró ningún archivo analisis_planificacion_*.xlsx")
        print(f"   Ruta buscada: {ANALISIS_DIR}")

    # ── 4. Subir Metas oficiales ──────────────────────────────────────────────
    metas_id = None
    if os.path.exists(METAS_PATH):
        metas_id = upload_or_update(
            svc, METAS_PATH, "Metas oficiales 1SEM Nuevo.xlsx", folder_id
        )
        print(f"✅ Metas oficiales subido")
        print(f"   ID: {metas_id}")
    else:
        print(f"⚠️  Metas oficiales no encontrado: {METAS_PATH}")

    # ── 5. Escribir drive_config.py ───────────────────────────────────────────
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        f.write('"""drive_config.py — Generado por setup_drive.py (NO subir a git)"""\n')
        f.write(f'DRIVE_FOLDER_ID = "{folder_id}"\n')
        f.write(f'METAS_FILE_ID   = "{metas_id or ""}"\n')
    print(f"\n✅ Configuración guardada en: {CONFIG_PATH}")

    # ── 6. Imprimir bloque para Streamlit Cloud ───────────────────────────────
    print("\n" + "=" * 60)
    print("📋 COPIA ESTO EN Streamlit Cloud → App settings → Secrets:\n")
    print("[dashboard]")
    print(f'drive_folder_id = "{folder_id}"')
    print(f'metas_file_id   = "{metas_id or ""}"')
    print()
    print("[google_credentials]")
    print("# Pega aquí el contenido de credentials.json como pares clave=valor")
    print("# Ver .streamlit/secrets.toml.example para el formato exacto")
    print("=" * 60)


if __name__ == "__main__":
    main()
