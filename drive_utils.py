"""
drive_utils.py — Utilidades Google Drive para UnionX
Usado por: dashboard_stock.py, actualizar_reportes.py, setup_drive.py
"""
import io
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

SCOPES_RO = ["https://www.googleapis.com/auth/drive.readonly"]
SCOPES_RW = ["https://www.googleapis.com/auth/drive"]
CREDS_LOCAL = r"C:\Users\felip\Desktop\UnionX Cloude\credentials.json"


def _creds(rw=False, secrets=None):
    scopes = SCOPES_RW if rw else SCOPES_RO
    if secrets is not None:
        info = {k: v for k, v in secrets.items()}
        # Streamlit puede escapar \n en private_key → desescapar
        if "private_key" in info:
            info["private_key"] = str(info["private_key"]).replace("\\n", "\n")
        return service_account.Credentials.from_service_account_info(
            info, scopes=scopes
        )
    return service_account.Credentials.from_service_account_file(
        CREDS_LOCAL, scopes=scopes
    )


def get_service(rw=False, secrets=None):
    """Retorna cliente autenticado de Drive API v3."""
    return build(
        "drive", "v3",
        credentials=_creds(rw=rw, secrets=secrets),
        cache_discovery=False,
    )


def download_bytes(service, file_id: str) -> bytes:
    """Descarga un archivo de Drive y retorna sus bytes crudos."""
    req = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    return buf.getvalue()


def list_folder(service, folder_id: str) -> list:
    """Lista archivos en carpeta (no recursivo).
    Retorna [{id, name, modifiedTime}] ordenado por modifiedTime desc."""
    resp = service.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields="files(id,name,modifiedTime)",
        orderBy="modifiedTime desc",
        pageSize=100,
    ).execute()
    return resp.get("files", [])


def find_file(service, folder_id: str, name: str) -> str | None:
    """Busca un archivo por nombre exacto en la carpeta. Retorna file_id o None."""
    resp = service.files().list(
        q=f"'{folder_id}' in parents and name='{name}' and trashed=false",
        fields="files(id)",
        pageSize=1,
    ).execute()
    files = resp.get("files", [])
    return files[0]["id"] if files else None


def upload_or_update(
    service, local_path: str, name: str,
    folder_id: str, existing_id: str | None = None
) -> str:
    """Sube archivo nuevo o actualiza uno existente. Retorna file_id."""
    mime = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    media = MediaFileUpload(local_path, mimetype=mime, resumable=True)
    if existing_id:
        f = service.files().update(
            fileId=existing_id,
            media_body=media,
            fields="id",
        ).execute()
    else:
        f = service.files().create(
            body={"name": name, "parents": [folder_id]},
            media_body=media,
            fields="id",
        ).execute()
    return f["id"]


def share_folder(service, folder_id: str, email: str, role: str = "reader"):
    """Comparte una carpeta con un email (role: reader|writer|commenter)."""
    service.permissions().create(
        fileId=folder_id,
        body={"type": "user", "role": role, "emailAddress": email},
        sendNotificationEmail=False,
    ).execute()
