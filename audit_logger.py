# -*- coding: utf-8 -*-
"""
audit_logger.py — MUH-NET Denetim ve İz Takip Modülü

Her kritik işlem SHA-256 imzasıyla mühürlenerek audit_logs tablosuna yazılır.
İmza; timestamp, user_action, target_id, details, admin_id ve severity
alanlarının birleşiminden hesaplanır. Tabloda herhangi bir satırın dışarıdan
değiştirilmesi imza uyuşmazlığı olarak tespit edilir.
"""

from __future__ import annotations

import hashlib
import sqlite3
from datetime import datetime
from typing import Optional

from config import config

# ─── Şema ─────────────────────────────────────────────────────────────────────

_SCHEMA = """
CREATE TABLE IF NOT EXISTS audit_logs (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp   TEXT    NOT NULL,
    user_action TEXT    NOT NULL,
    target_id   TEXT    NOT NULL DEFAULT '',
    details     TEXT    NOT NULL DEFAULT '',
    admin_id    TEXT    NOT NULL DEFAULT 'YÖNETİCİ',
    severity    TEXT    NOT NULL DEFAULT 'BİLGİ',
    imza        TEXT    NOT NULL DEFAULT ''
);
"""

# Şiddet seviyeleri (artan önem sırası)
SEV_INFO     = "BİLGİ"
SEV_WARNING  = "UYARI"
SEV_CRITICAL = "KRİTİK HATA"


class AuditLogger:
    """
    Thread-safe olmayan basit audit kayıt yöneticisi.
    Her kayıt bir SHA-256 parmak izi (imza) içerir.
    """

    def __init__(self) -> None:
        self._path = config.db_path
        self._ensure_schema()

    # ── İç yardımcılar ────────────────────────────────────────────────────────

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self._path))
        conn.row_factory = sqlite3.Row
        return conn

    def _ensure_schema(self) -> None:
        with self._connect() as conn:
            conn.executescript(_SCHEMA)

    @staticmethod
    def _compute_imza(ts: str, action: str, target: str,
                      details: str, admin: str, severity: str) -> str:
        """Kayıt içeriğinden deterministik SHA-256 imza üretir."""
        content = f"{ts}|{action}|{target}|{details}|{admin}|{severity}"
        return hashlib.sha256(content.encode("utf-8")).hexdigest()

    # ── Kayıt API'si ──────────────────────────────────────────────────────────

    def log(
        self,
        user_action: str,
        target_id:   str = "",
        details:     str = "",
        admin_id:    str = "YÖNETİCİ",
        severity:    str = SEV_INFO,
    ) -> None:
        """
        Yeni bir denetim kaydı oluşturur.
        İmza otomatik hesaplanarak tabloya yazılır.
        """
        ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        imza = self._compute_imza(ts, user_action, target_id, details, admin_id, severity)
        with self._connect() as conn:
            conn.execute(
                """INSERT INTO audit_logs
                   (timestamp, user_action, target_id, details,
                    admin_id, severity, imza)
                   VALUES (?,?,?,?,?,?,?)""",
                (ts, user_action, target_id, details, admin_id, severity, imza),
            )

    def log_warning(
        self,
        user_action: str,
        target_id:   str = "",
        details:     str = "",
        admin_id:    str = "SİSTEM",
    ) -> None:
        self.log(user_action, target_id, details, admin_id, SEV_WARNING)

    def log_critical(
        self,
        user_action: str,
        target_id:   str = "",
        details:     str = "",
        admin_id:    str = "SİSTEM",
    ) -> None:
        self.log(user_action, target_id, details, admin_id, SEV_CRITICAL)

    # ── Sorgulama API'si ──────────────────────────────────────────────────────

    def get_logs(
        self,
        limit:           int = 500,
        severity_filter: Optional[str] = None,
        keyword:         Optional[str] = None,
    ) -> list[dict]:
        """
        Kayıtları en yeniden en eskiye doğru döner.
        severity_filter ve/veya keyword ile filtreler.
        """
        query  = "SELECT * FROM audit_logs WHERE 1=1"
        params: list = []

        if severity_filter:
            query += " AND severity = ?"
            params.append(severity_filter)

        if keyword:
            kw = f"%{keyword}%"
            query += (
                " AND (user_action LIKE ? OR target_id LIKE ?"
                " OR details LIKE ? OR admin_id LIKE ?)"
            )
            params.extend([kw, kw, kw, kw])

        query += " ORDER BY id DESC LIMIT ?"
        params.append(limit)

        with self._connect() as conn:
            rows = conn.execute(query, params).fetchall()
        return [dict(r) for r in rows]

    def get_stats(self) -> dict:
        """Toplam, seviye bazlı sayım ve son kayıt zamanı döner."""
        with self._connect() as conn:
            total = conn.execute("SELECT COUNT(*) FROM audit_logs").fetchone()[0]
            kritik = conn.execute(
                "SELECT COUNT(*) FROM audit_logs WHERE severity=?",
                (SEV_CRITICAL,),
            ).fetchone()[0]
            uyari = conn.execute(
                "SELECT COUNT(*) FROM audit_logs WHERE severity=?",
                (SEV_WARNING,),
            ).fetchone()[0]
            last = conn.execute(
                "SELECT timestamp FROM audit_logs ORDER BY id DESC LIMIT 1"
            ).fetchone()
        return {
            "toplam":  total,
            "kritik":  kritik,
            "uyari":   uyari,
            "bilgi":   total - kritik - uyari,
            "son_kayit": last["timestamp"] if last else "—",
        }

    # ── Bütünlük Doğrulama ────────────────────────────────────────────────────

    def verify_integrity(self) -> tuple[int, int, list[int]]:
        """
        Tüm kayıtların SHA-256 imzasını yeniden hesaplayarak doğrular.
        Döner: (toplam_kayıt, bozuk_kayıt_sayısı, bozuk_id_listesi)
        """
        with self._connect() as conn:
            rows = conn.execute("SELECT * FROM audit_logs").fetchall()

        total  = len(rows)
        broken_ids: list[int] = []

        for r in rows:
            expected = self._compute_imza(
                r["timestamp"], r["user_action"], r["target_id"],
                r["details"],   r["admin_id"],    r["severity"],
            )
            if expected != r["imza"]:
                broken_ids.append(r["id"])

        return total, len(broken_ids), broken_ids


# Modül düzeyinde tek örnek
audit = AuditLogger()
