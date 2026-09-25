"""Local application settings; the authentication token uses Windows keyring."""

from __future__ import annotations

import json
import os
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import keyring


SERVICE = "ReceiptLedger-LMStudio"
ACCOUNT = "server-token"


def application_dir() -> Path:
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent


def default_data_dir() -> Path:
    # Isolate development/test instances from the installed application's data.
    override = os.environ.get("RECEIPT_LEDGER_DATA_DIR")
    return Path(override).resolve() if override else application_dir() / "data"


def resolve_config_path(value: str) -> str:
    if not value:
        return ""
    path = Path(value).expanduser()
    return str((application_dir() / path).resolve() if not path.is_absolute() else path)


def portable_path(value: str) -> str:
    if not value:
        return value
    try:
        return Path(value).resolve().relative_to(application_dir()).as_posix()
    except ValueError:
        return str(value)


def map_session_paths(value, transform):
    path_keys = {"workbook", "image", "path", "active", "filter_pinned", "saved_to", "stored_image"}
    if isinstance(value, list):
        return [map_session_paths(item, transform) for item in value]
    if isinstance(value, dict):
        return {key: transform(item) if key in path_keys and isinstance(item, str) and item
                else map_session_paths(item, transform) for key, item in value.items()}
    return value


def app_data_dir() -> Path:
    override = os.environ.get("RECEIPT_LEDGER_DATA_DIR")
    if override:
        return Path(override)
    return Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData" / "Local"))) / "ReceiptLedger"


@dataclass
class Settings:
    server_url: str = "http://localhost:1234/v1"
    model: str = ""
    model_choices: list[str] = field(default_factory=list)
    workbook_path: str = ""
    keep_images: bool = True
    theme: str = "system"
    image_folder: str = ""
    backup_generations: int = 20

    @classmethod
    def load(cls) -> "Settings":
        path = app_data_dir() / "settings.json"
        if not path.exists():
            return cls()
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                return cls()
            result = cls(**{key: data[key] for key in cls.__dataclass_fields__ if key in data})
            defaults = cls()
            for key in ("server_url", "model", "workbook_path", "image_folder", "theme"):
                if not isinstance(getattr(result, key), str):
                    setattr(result, key, getattr(defaults, key))
            if not isinstance(result.keep_images, bool):
                result.keep_images = defaults.keep_images
            result.model_choices = list(dict.fromkeys(x for x in result.model_choices if isinstance(x, str))) if isinstance(result.model_choices, list) else []
            # Older versions always kept images here. Never silently relocate existing images.
            if "image_folder" not in data:
                result.image_folder = str(app_data_dir() / "images")
            result.workbook_path = resolve_config_path(result.workbook_path)
            result.image_folder = resolve_config_path(result.image_folder)
            if type(result.backup_generations) is not int or not 1 <= result.backup_generations <= 999:
                result.backup_generations = 20
            return result
        except (OSError, ValueError, TypeError):
            return cls()

    def save(self) -> None:
        folder = app_data_dir()
        folder.mkdir(parents=True, exist_ok=True)
        fd, name = tempfile.mkstemp(prefix=".settings-", suffix=".json", dir=folder)
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as handle:
                data = dict(self.__dict__)
                for key in ("workbook_path", "image_folder"):
                    data[key] = portable_path(data[key])
                json.dump(data, handle, ensure_ascii=False, indent=2)
            os.replace(name, folder / "settings.json")
        finally:
            if os.path.exists(name):
                os.unlink(name)


def get_token() -> str:
    try:
        return keyring.get_password(SERVICE, ACCOUNT) or ""
    except Exception:
        return ""


def save_token(token: str) -> None:
    if token:
        keyring.set_password(SERVICE, ACCOUNT, token)
    else:
        try:
            keyring.delete_password(SERVICE, ACCOUNT)
        except keyring.errors.PasswordDeleteError:
            pass
