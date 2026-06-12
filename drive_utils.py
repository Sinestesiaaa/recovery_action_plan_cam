from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
import os
import tempfile
from typing import Any
from urllib.parse import urlparse


DRIVE_SCOPES = ("https://www.googleapis.com/auth/drive",)


@dataclass(frozen=True)
class DriveFileMeta:
    file_id: str
    name: str
    modified_time: str | None = None
    size: int | None = None
    parents: tuple[str, ...] = ()


def discover_service_account_file(preferred_path: str | Path | None = None) -> Path:
    if preferred_path:
        candidate = Path(preferred_path)
        if candidate.exists():
            return candidate

    env_value = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "").strip()
    env_path = Path(env_value) if env_value else None
    if env_path and env_path.exists():
        return env_path

    env_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if env_json:
        temp_file = Path(tempfile.gettempdir()) / "google_service_account.json"
        temp_file.write_text(env_json, encoding="utf-8")
        return temp_file

    secret_folder = Path("secret")
    if secret_folder.exists():
        json_files = sorted(secret_folder.glob("*.json"))
        if json_files:
            return json_files[0]

    raise FileNotFoundError(
        "Service account JSON tidak ditemukan. Simpan file JSON di folder `secret/` "
        "atau set `GOOGLE_APPLICATION_CREDENTIALS`."
    )


def extract_drive_folder_id(value: str | None) -> str | None:
    if not value:
        return None

    text = value.strip()
    if not text:
        return None

    if "drive.google.com" not in text:
        return text

    parsed = urlparse(text)
    segments = [segment for segment in parsed.path.split("/") if segment]
    if "folders" in segments:
        index = segments.index("folders")
        if index + 1 < len(segments):
            return segments[index + 1]

    return text


def _build_drive_service(credentials_path: str | Path) -> Any:
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise ImportError(
            "Dependency Google Drive belum terpasang. Install `google-api-python-client` "
            "dan `google-auth`."
        ) from exc

    credentials = service_account.Credentials.from_service_account_file(
        str(credentials_path),
        scopes=list(DRIVE_SCOPES),
    )
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def resolve_latest_file(
    file_name: str,
    credentials_path: str | Path,
    file_id: str | None = None,
    folder_id: str | None = None,
) -> DriveFileMeta:
    service = _build_drive_service(credentials_path)

    if file_id:
        metadata = (
            service.files()
            .get(
                fileId=file_id,
                fields="id,name,modifiedTime,size",
                supportsAllDrives=True,
            )
            .execute()
        )
        return DriveFileMeta(
            file_id=metadata["id"],
            name=metadata.get("name", file_name),
            modified_time=metadata.get("modifiedTime"),
            size=int(metadata["size"]) if metadata.get("size") else None,
            parents=tuple(metadata.get("parents", [])),
        )

    folder_ids_to_search = [folder_id] if folder_id else [None]
    visited_folders: set[str] = set()
    metadata = None

    while folder_ids_to_search and metadata is None:
        current_folder_id = folder_ids_to_search.pop(0)
        if current_folder_id and current_folder_id in visited_folders:
            continue
        if current_folder_id:
            visited_folders.add(current_folder_id)

        escaped_name = file_name.replace("'", "\\'")
        query_parts = [f"name = '{escaped_name}'", "trashed = false"]
        if current_folder_id:
            query_parts.append(f"'{current_folder_id}' in parents")
        query = " and ".join(query_parts)

        response = (
            service.files()
            .list(
                q=query,
                spaces="drive",
                fields="files(id,name,modifiedTime,size,parents,mimeType)",
                orderBy="modifiedTime desc",
                pageSize=10,
                includeItemsFromAllDrives=True,
                supportsAllDrives=True,
            )
            .execute()
        )
        files = response.get("files", [])
        if files:
            metadata = files[0]
            break

        if current_folder_id:
            child_response = (
                service.files()
                .list(
                    q=f"'{current_folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
                    spaces="drive",
                    fields="files(id,name)",
                    includeItemsFromAllDrives=True,
                    supportsAllDrives=True,
                    pageSize=100,
                )
                .execute()
            )
            for child in child_response.get("files", []):
                folder_ids_to_search.append(child["id"])

    if metadata is None:
        raise FileNotFoundError(
            f"File `{file_name}` tidak ditemukan di Google Drive yang dapat diakses service account."
        )

    return DriveFileMeta(
        file_id=metadata["id"],
        name=metadata.get("name", file_name),
        modified_time=metadata.get("modifiedTime"),
        size=int(metadata["size"]) if metadata.get("size") else None,
        parents=tuple(metadata.get("parents", [])),
    )


def find_child_folder_id(
    credentials_path: str | Path,
    parent_folder_id: str,
    child_folder_name: str,
) -> str | None:
    service = _build_drive_service(credentials_path)
    response = (
        service.files()
        .list(
            q=(
                f"'{parent_folder_id}' in parents and "
                f"mimeType = 'application/vnd.google-apps.folder' and trashed = false"
            ),
            spaces="drive",
            fields="files(id,name)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
            pageSize=100,
        )
        .execute()
    )
    for item in response.get("files", []):
        if item.get("name", "").strip().lower() == child_folder_name.strip().lower():
            return item["id"]
    return None


def list_files_in_folder_tree(
    credentials_path: str | Path,
    folder_id: str,
    suffixes: tuple[str, ...] = (".xlsx",),
) -> list[DriveFileMeta]:
    service = _build_drive_service(credentials_path)
    pending_folders = [folder_id]
    visited_folders: set[str] = set()
    files: list[DriveFileMeta] = []

    while pending_folders:
        current_folder_id = pending_folders.pop(0)
        if current_folder_id in visited_folders:
            continue
        visited_folders.add(current_folder_id)

        response = (
            service.files()
            .list(
                q=f"'{current_folder_id}' in parents and trashed = false",
                spaces="drive",
                fields="files(id,name,modifiedTime,size,parents,mimeType)",
                includeItemsFromAllDrives=True,
                supportsAllDrives=True,
                pageSize=1000,
            )
            .execute()
        )

        for item in response.get("files", []):
            mime_type = item.get("mimeType", "")
            if mime_type == "application/vnd.google-apps.folder":
                pending_folders.append(item["id"])
                continue

            name = item.get("name", "")
            if suffixes and not name.lower().endswith(suffixes):
                continue

            files.append(
                DriveFileMeta(
                    file_id=item["id"],
                    name=name,
                    modified_time=item.get("modifiedTime"),
                    size=int(item["size"]) if item.get("size") else None,
                    parents=tuple(item.get("parents", [])),
                )
            )

    return files


def download_file_bytes(credentials_path: str | Path, file_id: str) -> bytes:
    try:
        from googleapiclient.http import MediaIoBaseDownload
    except ImportError as exc:
        raise ImportError(
            "Dependency Google Drive belum terpasang. Install `google-api-python-client` "
            "dan `google-auth`."
        ) from exc

    service = _build_drive_service(credentials_path)
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    buffer = BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)

    done = False
    while not done:
        _, done = downloader.next_chunk()

    return buffer.getvalue()


def upload_or_replace_file(
    credentials_path: str | Path,
    folder_id: str,
    file_name: str,
    file_bytes: bytes,
    mime_type: str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
) -> DriveFileMeta:
    try:
        from googleapiclient.http import MediaIoBaseUpload
    except ImportError as exc:
        raise ImportError(
            "Dependency Google Drive belum terpasang. Install `google-api-python-client` "
            "dan `google-auth`."
        ) from exc

    service = _build_drive_service(credentials_path)
    escaped_name = file_name.replace("'", "\\'")
    existing = (
        service.files()
        .list(
            q=(
                f"'{folder_id}' in parents and name = '{escaped_name}' and trashed = false"
            ),
            spaces="drive",
            fields="files(id,name,modifiedTime,size,parents)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
            pageSize=10,
        )
        .execute()
    ).get("files", [])

    media = MediaIoBaseUpload(BytesIO(file_bytes), mimetype=mime_type, resumable=False)
    metadata = {"name": file_name, "parents": [folder_id]}

    if existing:
        result = (
            service.files()
            .update(
                fileId=existing[0]["id"],
                media_body=media,
                body=metadata,
                supportsAllDrives=True,
            )
            .execute()
        )
    else:
        result = (
            service.files()
            .create(
                media_body=media,
                body=metadata,
                fields="id,name,modifiedTime,size,parents",
                supportsAllDrives=True,
            )
            .execute()
        )

    return DriveFileMeta(
        file_id=result["id"],
        name=result.get("name", file_name),
        modified_time=result.get("modifiedTime"),
        size=int(result["size"]) if result.get("size") else None,
        parents=tuple(result.get("parents", [])),
    )
