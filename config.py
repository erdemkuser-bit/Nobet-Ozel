# -*- coding: utf-8 -*-
"""
config.py — MUH-NET uygulama yapılandırması
Hardcoded yol veya şifre içermez; her şey config.json üzerinden yönetilir.
"""

from __future__ import annotations

import json
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent
_CONFIG_FILE = PROJECT_DIR / "config.json"
APP_VERSION = "1.0.0"

_DEFAULTS: dict = {
    "db_path":      str(PROJECT_DIR / "secure_data.db"),
    "excel_path":   str(PROJECT_DIR / "personel_listesi.xlsx"),
    "app_version":  APP_VERSION,
    "first_run":    True,
}


class AppConfig:
    """
    Uygulama genelinde tek örnek (singleton) yapılandırma nesnesi.
    config.json yoksa varsayılan değerlerle oluşturulur.
    """

    def __init__(self) -> None:
        self._data: dict = {}
        self._load()

    def _load(self) -> None:
        if _CONFIG_FILE.exists():
            try:
                with open(_CONFIG_FILE, encoding="utf-8") as f:
                    saved = json.load(f)
                # Yeni anahtarlar için varsayılanlarla birleştir
                self._data = {**_DEFAULTS, **saved}
            except (json.JSONDecodeError, OSError):
                self._data = _DEFAULTS.copy()
        else:
            self._data = _DEFAULTS.copy()
            self._save()

    def _save(self) -> None:
        with open(_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self._data, f, ensure_ascii=False, indent=2)

    def get(self, key: str, default=None):
        return self._data.get(key, default)

    def set(self, key: str, value) -> None:
        self._data[key] = value
        self._save()

    @property
    def db_path(self) -> Path:
        return Path(self._data["db_path"])

    @property
    def excel_path(self) -> Path:
        return Path(self._data["excel_path"])

    @property
    def first_run(self) -> bool:
        return bool(self._data.get("first_run", True))

    def mark_configured(self) -> None:
        self.set("first_run", False)


# Modül düzeyinde tek örnek — tüm dosyalar bu nesneyi import eder
config = AppConfig()
