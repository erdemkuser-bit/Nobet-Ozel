#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
reset_lockout.py — MUH-NET Acil Sistem Sıfırlama Aracı
═══════════════════════════════════════════════════════

AMAÇ   : Kilitlenmiş hesabı açmak ve/veya yönetici şifresini sıfırlamak.
BAĞIMLI: main.py, gui kütüphanesi veya harici bağımlılık GEREKTİRMEZ.
GÜVENLİK:
  • Bu script yalnızca secure_data.db ile AYNI KLASÖRDE çalışır.
  • Başka bir klasörden çalıştırılırsa veritabanını bulamaz ve çıkar.
  • Her başarılı işlem audit_logs tablosuna KRİTİK HATA seviyesinde yazılır.

KULLANIM:
  python reset_lockout.py

"""

from __future__ import annotations

import getpass
import hashlib
import sqlite3
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

# ─── Güvenlik: DB yalnızca bu scriptin bulunduğu klasörde aranır ─────────────

_SCRIPT_DIR: Path = Path(__file__).resolve().parent
_DB_PATH:    Path = _SCRIPT_DIR / "secure_data.db"

# Ayar anahtarları (auth.py ile aynı)
_KEY_HASH   = "admin_password_hash"
_KEY_FAILED = "failed_attempts"
_KEY_LOCKED = "is_locked"


# ─── Renkli terminal çıktısı (Windows/ANSI) ──────────────────────────────────

class _C:
    RESET  = "\033[0m"
    BOLD   = "\033[1m"
    RED    = "\033[91m"
    YELLOW = "\033[93m"
    GREEN  = "\033[92m"
    CYAN   = "\033[96m"
    DIM    = "\033[2m"

    @classmethod
    def r(cls, text: str, color: str) -> str:
        return f"{color}{text}{cls.RESET}"


def _enable_ansi() -> None:
    """Windows'ta ANSI kaçış kodlarını etkinleştirir."""
    if sys.platform == "win32":
        import ctypes
        kernel32 = ctypes.windll.kernel32
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)


# ─── Veritabanı yardımcıları (bağımsız — import yok) ─────────────────────────

def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(str(_DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _get_setting(key: str) -> Optional[str]:
    with _connect() as conn:
        row = conn.execute(
            "SELECT deger FROM ayarlar WHERE anahtar = ?", (key,)
        ).fetchone()
        return row["deger"] if row else None


def _set_setting(key: str, value: str) -> None:
    with _connect() as conn:
        conn.execute(
            "INSERT INTO ayarlar(anahtar, deger) VALUES(?,?) "
            "ON CONFLICT(anahtar) DO UPDATE SET deger = excluded.deger",
            (key, value),
        )


def _sha256(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _audit_log(action: str, details: str,
               severity: str = "KRİTİK HATA",
               admin_id: str = "RESET_SCRIPT") -> None:
    """
    audit_logs tablosuna doğrudan yazar (audit_logger.py import etmeden).
    İmzayı audit_logger.py ile aynı algoritmayla hesaplar —
    böylece ana uygulama imza doğrulamasında bu kayıtlar da geçer.
    """
    ts    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    target = "RESET"
    content = f"{ts}|{action}|{target}|{details}|{admin_id}|{severity}"
    imza  = hashlib.sha256(content.encode("utf-8")).hexdigest()

    with _connect() as conn:
        # Tablo yoksa oluştur (çok nadir durum: DB sıfırdan oluşturulmuş)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS audit_logs (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp   TEXT    NOT NULL,
                user_action TEXT    NOT NULL,
                target_id   TEXT    NOT NULL DEFAULT '',
                details     TEXT    NOT NULL DEFAULT '',
                admin_id    TEXT    NOT NULL DEFAULT 'SİSTEM',
                severity    TEXT    NOT NULL DEFAULT 'BİLGİ',
                imza        TEXT    NOT NULL DEFAULT ''
            )
        """)
        conn.execute(
            """INSERT INTO audit_logs
               (timestamp, user_action, target_id, details,
                admin_id, severity, imza)
               VALUES (?,?,?,?,?,?,?)""",
            (ts, action, target, details, admin_id, severity, imza),
        )


# ─── Ekran yardımcıları ───────────────────────────────────────────────────────

def _banner() -> None:
    c = _C
    print()
    print(c.r("═" * 62, c.CYAN))
    print(c.r("  MUH-NET — Acil Sistem Sıfırlama Aracı", c.BOLD + c.CYAN))
    print(c.r("  reset_lockout.py  |  Tüm işlemler denetim tablosuna loglanır", c.DIM))
    print(c.r("═" * 62, c.CYAN))
    print()


def _show_status() -> dict:
    """Mevcut durumu okur, terminale yazar ve dict döner."""
    failed  = _get_setting(_KEY_FAILED) or "0"
    locked  = _get_setting(_KEY_LOCKED) or "0"
    has_pw  = bool(_get_setting(_KEY_HASH))

    c = _C
    print(c.r("  ── Mevcut Sistem Durumu ──", c.BOLD))
    print(f"  Şifre tanımlı     : "
          + (c.r("Evet", c.GREEN) if has_pw else c.r("Hayır (kurulum gerekli)", c.YELLOW)))
    print(f"  Başarısız deneme  : "
          + (c.r(failed, c.YELLOW) if int(failed) > 0 else c.r(failed, c.GREEN)))
    print(f"  Hesap kilitli     : "
          + (c.r("EVET  ⚠", c.RED) if locked == "1" else c.r("Hayır", c.GREEN)))
    print()
    return {"has_pw": has_pw, "failed": int(failed), "locked": locked == "1"}


# ─── İşlem fonksiyonları ──────────────────────────────────────────────────────

def unlock_account() -> bool:
    """
    failed_attempts → 0, is_locked → 0 olarak ayarlar.
    Denetim tablosuna KRİTİK kayıt atar.
    """
    _set_setting(_KEY_FAILED, "0")
    _set_setting(_KEY_LOCKED, "0")

    print(_C.r("  ✔ Başarısız deneme sayacı sıfırlandı (failed_attempts = 0).", _C.GREEN))
    print(_C.r("  ✔ Hesap kilidi kaldırıldı (is_locked = 0).", _C.GREEN))

    _audit_log(
        "YÖNETİCİ MÜDAHALESİ: Hesap Kilidi Kaldırıldı",
        "failed_attempts=0, is_locked=0 olarak sıfırlandı. "
        "İşlem: reset_lockout.py",
    )
    return True


def reset_password() -> bool:
    """
    Yeni şifre belirlemek veya hash'i temizlemek için iki seçenek sunar.
    Her iki durumda da denetim tablosuna KRİTİK kayıt atılır.
    """
    c = _C
    print()
    print(c.r("  ── Şifre Sıfırlama ──", c.BOLD))
    print("  [1] Yeni şifre belirle")
    print("  [2] Şifreyi tamamen sil  (sonraki girişte kurulum ekranı açılır)")
    print("  [0] Bu adımı atla")
    print()

    choice = input("  Seçiminiz (0/1/2): ").strip()

    if choice == "0":
        print(c.r("  ↩ Şifre adımı atlandı.", c.DIM))
        return False

    if choice == "1":
        while True:
            try:
                pw1 = getpass.getpass("  Yeni şifre (en az 6 karakter, görünmez): ")
            except (EOFError, KeyboardInterrupt):
                print()
                print(c.r("  İptal edildi.", c.YELLOW))
                return False

            pw1 = pw1.strip()
            if len(pw1) < 6:
                print(c.r("  [!] Şifre en az 6 karakter olmalıdır.", c.YELLOW))
                continue

            try:
                pw2 = getpass.getpass("  Yeni şifre (tekrar):                    ")
            except (EOFError, KeyboardInterrupt):
                print()
                print(c.r("  İptal edildi.", c.YELLOW))
                return False

            if pw1 != pw2:
                print(c.r("  [!] Şifreler eşleşmiyor. Tekrar deneyin.", c.YELLOW))
                continue

            _set_setting(_KEY_HASH, _sha256(pw1))
            _set_setting(_KEY_FAILED, "0")
            _set_setting(_KEY_LOCKED, "0")
            print(c.r("  ✔ Yeni şifre kaydedildi. Kilit ve sayaç da sıfırlandı.", c.GREEN))
            _audit_log(
                "YÖNETİCİ MÜDAHALESİ: Şifre Terminelden Sıfırlandı",
                "Yeni şifre hash'i güncellendi. failed_attempts=0, is_locked=0. "
                "İşlem: reset_lockout.py",
            )
            return True

    elif choice == "2":
        confirm = input(
            c.r("  Şifreyi tamamen silmek istediğinizden emin misiniz? (E/H): ", c.YELLOW)
        ).strip().upper()
        if confirm == "E":
            _set_setting(_KEY_HASH, "")
            _set_setting(_KEY_FAILED, "0")
            _set_setting(_KEY_LOCKED, "0")
            print(c.r(
                "  ✔ Şifre hash'i silindi. Sonraki girişte kurulum ekranı açılacak.",
                c.GREEN,
            ))
            _audit_log(
                "YÖNETİCİ MÜDAHALESİ: Şifre Hash'i Silindi",
                "admin_password_hash boşaltıldı — sonraki girişte setup ekranı gelecek. "
                "İşlem: reset_lockout.py",
            )
            return True
        else:
            print(c.r("  ↩ Şifre silme işlemi iptal edildi.", c.DIM))
            return False
    else:
        print(c.r("  [!] Geçersiz seçim. Bu adım atlandı.", c.YELLOW))
        return False


# ─── Ana akış ────────────────────────────────────────────────────────────────

def main() -> None:
    _enable_ansi()
    _banner()

    # ── Güvenlik: DB bu scriptle aynı klasörde mi? ─────────────────────────
    if not _DB_PATH.exists():
        print(_C.r(
            f"  [HATA] Veritabanı bulunamadı:\n  {_DB_PATH}\n\n"
            "  Bu script yalnızca secure_data.db ile AYNI KLASÖRDE çalışır.\n"
            "  Lütfen scripti proje kök dizininden çalıştırın.",
            _C.RED,
        ))
        sys.exit(1)

    # ── Mevcut durum ───────────────────────────────────────────────────────
    status = _show_status()

    # ── Kullanıcı onayı ────────────────────────────────────────────────────
    print(_C.r(
        "  ⚠  UYARI: Bu araç güvenlik mekanizmalarını atlatır.\n"
        "     Tüm işlemler denetim tablosuna KRİTİK HATA olarak kaydedilir.",
        _C.YELLOW,
    ))
    print()
    try:
        confirm = input(
            _C.r("  Sistemi sıfırlamak istediğinizden emin misiniz? (E/H): ", _C.BOLD)
        ).strip().upper()
    except (EOFError, KeyboardInterrupt):
        print()
        print(_C.r("\n  İptal edildi. Çıkılıyor...\n", _C.DIM))
        sys.exit(0)

    if confirm != "E":
        print(_C.r("\n  İşlem iptal edildi. Çıkılıyor...\n", _C.DIM))
        sys.exit(0)

    # ── İşlem menüsü ──────────────────────────────────────────────────────
    print()
    print(_C.r("  ── Yapılacak İşlem ──", _C.BOLD))
    print("  [1] Hesap kilidini kaldır                (failed_attempts + is_locked sıfırla)")
    print("  [2] Şifreyi sıfırla                      (yeni şifre veya hash silme)")
    print("  [3] Her ikisini yap                      (kilit kaldır + şifre sıfırla)")
    print("  [0] İptal")
    print()

    try:
        action = input("  Seçiminiz (0/1/2/3): ").strip()
    except (EOFError, KeyboardInterrupt):
        print()
        print(_C.r("\n  İptal edildi.\n", _C.DIM))
        sys.exit(0)

    if action == "0":
        print(_C.r("\n  İşlem iptal edildi. Çıkılıyor...\n", _C.DIM))
        sys.exit(0)

    if action not in ("1", "2", "3"):
        print(_C.r(f"\n  [HATA] Geçersiz seçim: '{action}'. Çıkılıyor.\n", _C.RED))
        sys.exit(1)

    print()
    performed: list[str] = []

    if action in ("1", "3"):
        if unlock_account():
            performed.append("hesap kilidi kaldırıldı")

    if action in ("2", "3"):
        if reset_password():
            performed.append("şifre sıfırlandı")

    # ── Özet KRİTİK denetim kaydı ─────────────────────────────────────────
    if performed:
        summary = ", ".join(performed)
        _audit_log(
            "YÖNETİCİ MÜDAHALESİ: Sistem Resetleme ve Kilit Kaldırma işlemi gerçekleştirildi.",
            f"Yapılan işlemler: {summary}. Script: reset_lockout.py",
        )

        print()
        print(_C.r("  " + "─" * 58, _C.CYAN))
        print(_C.r(f"  ✔ Tamamlandı: {summary}", _C.GREEN + _C.BOLD))
        print(_C.r(
            "  ✔ Denetim tablosuna KRİTİK HATA kaydı eklendi.",
            _C.GREEN,
        ))
        print(_C.r("  " + "─" * 58, _C.CYAN))
    else:
        print(_C.r("\n  Hiçbir işlem yapılmadı.\n", _C.YELLOW))

    print()


if __name__ == "__main__":
    main()
