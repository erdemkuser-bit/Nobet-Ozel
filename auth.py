# -*- coding: utf-8 -*-
"""
auth.py — MUH-NET yönetici kimlik doğrulama katmanı
SHA-256 tabanlı şifre yönetimi; hashlenmiş değer secure_data.db'de tutulur.
Başarısız giriş sayısı ve kilit durumu da DB'ye kalıcı olarak yazılır.
"""

from __future__ import annotations

import hashlib
import tkinter as tk
from tkinter import messagebox, simpledialog

from audit_logger import audit
# data_manager geç import edilir (dairesel import önlemi)

_MAX_ATTEMPTS = 3
_KEY_HASH     = "admin_password_hash"
_KEY_FAILED   = "failed_attempts"
_KEY_LOCKED   = "is_locked"


def _sha256(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _get_db():
    from data_manager import db
    return db


# ── Şifre yönetimi ────────────────────────────────────────────────────────────

def is_password_set() -> bool:
    """DB'de yönetici şifresi hash'i kayıtlı ve dolu mu?"""
    val = _get_db().get_setting(_KEY_HASH)
    return bool(val)


# ── Kilit durumu (kalıcı) ─────────────────────────────────────────────────────

def is_account_locked() -> bool:
    """DB'deki is_locked değeri '1' ise True döner."""
    return _get_db().get_setting(_KEY_LOCKED) == "1"


def get_failed_attempts() -> int:
    """Kümülatif başarısız deneme sayısını döner."""
    val = _get_db().get_setting(_KEY_FAILED)
    try:
        return int(val) if val else 0
    except (ValueError, TypeError):
        return 0


def _increment_failed() -> int:
    """Başarısız deneme sayacını artırır; max dolunca hesabı kilitler."""
    db = _get_db()
    new_count = get_failed_attempts() + 1
    db.set_setting(_KEY_FAILED, str(new_count))
    if new_count >= _MAX_ATTEMPTS:
        db.set_setting(_KEY_LOCKED, "1")
    return new_count


def _reset_lockout() -> None:
    """Başarılı girişte sayacı ve kilidi sıfırlar."""
    db = _get_db()
    db.set_setting(_KEY_FAILED, "0")
    db.set_setting(_KEY_LOCKED, "0")


# ── Kullanıcıya açılan diyaloglar ─────────────────────────────────────────────

def setup_password(parent: tk.Widget) -> bool:
    """
    İlk kurulum veya şifre sıfırlama.
    Kullanıcı iki kez aynı şifreyi girerse hash DB'ye yazılır.
    Başarıda True, iptal/hata durumunda False döner.
    """
    while True:
        pw1 = simpledialog.askstring(
            "Yönetici Şifresi Oluştur",
            "Yeni yönetici şifresini belirleyin\n(en az 6 karakter):",
            parent=parent,
            show="*",
        )
        if pw1 is None:
            return False
        pw1 = pw1.strip()
        if len(pw1) < 6:
            messagebox.showwarning(
                "Geçersiz Şifre",
                "Şifre en az 6 karakter olmalıdır.",
                parent=parent,
            )
            continue

        pw2 = simpledialog.askstring(
            "Şifreyi Onayla",
            "Şifreyi tekrar girin:",
            parent=parent,
            show="*",
        )
        if pw2 is None:
            return False

        if pw1 != pw2:
            messagebox.showwarning(
                "Eşleşmiyor",
                "Girilen şifreler uyuşmuyor. Lütfen tekrar deneyin.",
                parent=parent,
            )
            continue

        _get_db().set_setting(_KEY_HASH, _sha256(pw1))
        _reset_lockout()   # Yeni şifre sonrası kilidi de kaldır
        messagebox.showinfo(
            "Başarılı",
            "Yönetici şifresi başarıyla oluşturuldu.",
            parent=parent,
        )
        audit.log("Yönetici Şifresi Oluşturuldu / Değiştirildi",
                  details="Şifre hash'i güncellendi.")
        return True


def authenticate(
    parent: tk.Widget,
    prompt: str = "Devam etmek için yönetici şifresini girin:",
) -> bool:
    """
    Şifre doğrulama penceresi gösterir.
    Doğru şifrede True, yanlış/iptal durumunda False döner.
    DB'de şifre yoksa setup_password() çağrılır.
    Hesap kilitliyse (is_locked=1) giriş denemesine izin vermez.
    """
    if not is_password_set():
        messagebox.showinfo(
            "İlk Kurulum",
            "Henüz yönetici şifresi tanımlanmamış.\n"
            "Lütfen şimdi bir şifre oluşturun.",
            parent=parent,
        )
        return setup_password(parent)

    # Kalıcı kilit kontrolü (reset_lockout.py ile kaldırılabilir)
    if is_account_locked():
        messagebox.showerror(
            "Hesap Kilitli",
            "Çok fazla hatalı giriş denemesi nedeniyle hesap kilitlenmiştir.\n\n"
            "Kilidi kaldırmak için:\n"
            "  python reset_lockout.py\n"
            "komutunu proje klasöründe çalıştırın.",
            parent=parent,
        )
        audit.log_critical(
            "Kilitli Hesaba Giriş Denemesi",
            details="is_locked=1 iken giriş denemesi engellendi.",
            admin_id="BİLİNMEYEN",
        )
        return False

    stored_hash = _get_db().get_setting(_KEY_HASH)

    for attempt in range(1, _MAX_ATTEMPTS + 1):
        pw = simpledialog.askstring(
            "Yönetici Yetkilendirmesi",
            f"{prompt}\n(Deneme {attempt}/{_MAX_ATTEMPTS})",
            parent=parent,
            show="*",
        )
        if pw is None:
            return False  # Kullanıcı iptal etti

        if _sha256(pw.strip()) == stored_hash:
            _reset_lockout()   # Başarılı girişte sayacı sıfırla
            return True

        # Başarısız deneme — DB'ye yaz ve logla
        new_count = _increment_failed()
        remaining = _MAX_ATTEMPTS - attempt
        is_last   = remaining == 0

        if is_last:
            audit.log_critical(
                "Başarısız Şifre Denemesi — Hesap Kilitlendi",
                details=(
                    f"Deneme {attempt}/{_MAX_ATTEMPTS}. "
                    f"failed_attempts={new_count}. ERİŞİM KİLİTLENDİ."
                ),
                admin_id="BİLİNMEYEN",
            )
        else:
            audit.log_warning(
                "Hatalı Şifre Girişi",
                details=(
                    f"Deneme {attempt}/{_MAX_ATTEMPTS}. "
                    f"failed_attempts={new_count}. Kalan hak: {remaining}."
                ),
                admin_id="BİLİNMEYEN",
            )

        if remaining > 0:
            messagebox.showwarning(
                "Hatalı Şifre",
                f"Şifre yanlış. {remaining} deneme hakkınız kaldı.",
                parent=parent,
            )

    messagebox.showerror(
        "Hesap Kilitlendi",
        f"{_MAX_ATTEMPTS} başarısız giriş denemesi.\n"
        "Hesabınız kilitlendi.\n\n"
        "Kilidi kaldırmak için:\n"
        "  python reset_lockout.py\n"
        "komutunu proje klasöründe çalıştırın.",
        parent=parent,
    )
    return False


def change_password(parent: tk.Widget) -> bool:
    """
    Mevcut şifreyi doğruladıktan sonra yeni şifre belirlenir.
    Başarıda True döner.
    """
    if not authenticate(parent, "Şifre değiştirmek için mevcut şifreyi girin:"):
        return False
    return setup_password(parent)
