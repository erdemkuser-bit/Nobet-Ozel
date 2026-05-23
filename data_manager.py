# -*- coding: utf-8 -*-
"""
data_manager.py — MUH-NET güvenli veri katmanı
SQLite tabanlı CRUD, Excel senkronizasyonu ve bütünlük kontrolü.
"""

from __future__ import annotations

import calendar
import sqlite3
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

from config import config
from scheduler import Personel, _to_date
from audit_logger import audit

# ─── Veritabanı şeması ────────────────────────────────────────────────────────

_SCHEMA = """
PRAGMA journal_mode = WAL;
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS personel (
    sicil            TEXT PRIMARY KEY,
    rutbe            TEXT NOT NULL DEFAULT 'Bilinmiyor',
    ad_soyad         TEXT NOT NULL DEFAULT '',
    kidem_yili       INTEGER NOT NULL DEFAULT 0,
    muafiyet_turu    TEXT NOT NULL DEFAULT 'Yok',
    grup_adi         TEXT NOT NULL DEFAULT 'Genel',
    toplam_puan      REAL NOT NULL DEFAULT 0,
    hafta_ici        INTEGER NOT NULL DEFAULT 0,
    cuma             INTEGER NOT NULL DEFAULT 0,
    cumartesi        INTEGER NOT NULL DEFAULT 0,
    pazar            INTEGER NOT NULL DEFAULT 0,
    ozel_nobet       INTEGER NOT NULL DEFAULT 0,
    son_nobet_tarihi TEXT,
    aktif            INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS mazeretler (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    sicil         TEXT NOT NULL,
    ad_soyad      TEXT NOT NULL DEFAULT '',
    baslangic     TEXT NOT NULL,
    bitis         TEXT NOT NULL,
    mazeret_turu  TEXT NOT NULL DEFAULT 'İzin',
    mazeret_kodu  TEXT,
    aciklama      TEXT,
    FOREIGN KEY(sicil) REFERENCES personel(sicil) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS ayarlar (
    anahtar TEXT PRIMARY KEY,
    deger   TEXT
);
"""


class DatabaseManager:
    """
    SQLite tabanlı güvenli veri yöneticisi.
    Tüm erişim bu sınıf üzerinden yapılır; doğrudan bağlantı açılmaz.
    """

    def __init__(self) -> None:
        self._path: Path = config.db_path
        self._ensure_schema()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self._path))
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def _ensure_schema(self) -> None:
        with self._connect() as conn:
            conn.executescript(_SCHEMA)

    # ── Ayarlar ───────────────────────────────────────────────────────────────

    def get_setting(self, key: str) -> Optional[str]:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT deger FROM ayarlar WHERE anahtar = ?", (key,)
            ).fetchone()
            return row["deger"] if row else None

    def set_setting(self, key: str, value: str) -> None:
        with self._connect() as conn:
            conn.execute(
                "INSERT INTO ayarlar(anahtar, deger) VALUES(?,?) "
                "ON CONFLICT(anahtar) DO UPDATE SET deger = excluded.deger",
                (key, value),
            )

    # ── Personel CRUD ─────────────────────────────────────────────────────────

    def get_all_personel(self) -> list[dict]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT * FROM personel WHERE aktif = 1 ORDER BY ad_soyad"
            ).fetchall()
            return [dict(r) for r in rows]

    def get_personel_by_sicil(self, sicil: str) -> Optional[dict]:
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM personel WHERE sicil = ?", (sicil,)
            ).fetchone()
            return dict(row) if row else None

    def add_personel(self, data: dict) -> str:
        """Yeni personel ekler. Hata varsa hata mesajı, başarıysa '' döner."""
        sicil = str(data.get("sicil", "")).strip()
        if not sicil:
            return "Sicil numarası boş olamaz."
        if len(sicil) > 20:
            return "Sicil numarası 20 karakteri geçemez."
        if self.get_personel_by_sicil(sicil):
            return f"'{sicil}' sicil numarası zaten kayıtlı."
        with self._connect() as conn:
            conn.execute(
                """INSERT INTO personel
                   (sicil, rutbe, ad_soyad, kidem_yili, muafiyet_turu, grup_adi,
                    toplam_puan, hafta_ici, cuma, cumartesi, pazar, ozel_nobet,
                    son_nobet_tarihi, aktif)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,1)""",
                (
                    sicil,
                    data.get("rutbe", "Bilinmiyor"),
                    data.get("ad_soyad", ""),
                    int(data.get("kidem_yili", 0)),
                    data.get("muafiyet_turu", "Yok"),
                    data.get("grup_adi", "Genel"),
                    float(data.get("toplam_puan", 0)),
                    int(data.get("hafta_ici", 0)),
                    int(data.get("cuma", 0)),
                    int(data.get("cumartesi", 0)),
                    int(data.get("pazar", 0)),
                    int(data.get("ozel_nobet", 0)),
                    data.get("son_nobet_tarihi"),
                ),
            )
        audit.log(
            "Personel Eklendi",
            target_id=sicil,
            details=f"{data.get('ad_soyad', '')} | {data.get('rutbe', '')} | "
                    f"Kıdem: {data.get('kidem_yili', 0)} | Muafiyet: {data.get('muafiyet_turu', 'Yok')}",
        )
        return ""

    def update_personel(self, sicil: str, data: dict) -> str:
        """Var olan personel kaydını günceller."""
        if not self.get_personel_by_sicil(sicil):
            return f"'{sicil}' sicil numarası bulunamadı."
        with self._connect() as conn:
            conn.execute(
                """UPDATE personel SET
                   rutbe=?, ad_soyad=?, kidem_yili=?, muafiyet_turu=?,
                   grup_adi=?, toplam_puan=?, hafta_ici=?, cuma=?,
                   cumartesi=?, pazar=?, ozel_nobet=?, son_nobet_tarihi=?
                   WHERE sicil=?""",
                (
                    data.get("rutbe", "Bilinmiyor"),
                    data.get("ad_soyad", ""),
                    int(data.get("kidem_yili", 0)),
                    data.get("muafiyet_turu", "Yok"),
                    data.get("grup_adi", "Genel"),
                    float(data.get("toplam_puan", 0)),
                    int(data.get("hafta_ici", 0)),
                    int(data.get("cuma", 0)),
                    int(data.get("cumartesi", 0)),
                    int(data.get("pazar", 0)),
                    int(data.get("ozel_nobet", 0)),
                    data.get("son_nobet_tarihi"),
                    sicil,
                ),
            )
        audit.log(
            "Personel Güncellendi",
            target_id=sicil,
            details=f"{data.get('ad_soyad', '')} | Rütbe: {data.get('rutbe', '')} | "
                    f"Puan: {data.get('toplam_puan', 0)}",
        )
        return ""

    def delete_personel(self, sicil: str) -> str:
        """Soft-delete: aktif=0 yapar."""
        row = self.get_personel_by_sicil(sicil)
        if not row:
            return f"'{sicil}' sicil numarası bulunamadı."
        with self._connect() as conn:
            conn.execute(
                "UPDATE personel SET aktif = 0 WHERE sicil = ?", (sicil,)
            )
        audit.log_warning(
            "Personel Silindi (Pasif)",
            target_id=sicil,
            details=f"{row['ad_soyad']} | {row['rutbe']}",
        )
        return ""

    def update_personel_counters(self, personel_list: list[Personel]) -> None:
        """Nöbet dağıtımı sonrası puan/sayaç güncelleme (toplu)."""
        with self._connect() as conn:
            for p in personel_list:
                conn.execute(
                    """UPDATE personel SET
                       toplam_puan=?, hafta_ici=?, cuma=?, cumartesi=?,
                       pazar=?, ozel_nobet=?, son_nobet_tarihi=?
                       WHERE sicil=?""",
                    (
                        p.toplam_puan,
                        p.hafta_ici,
                        p.cuma,
                        p.cumartesi,
                        p.pazar,
                        p.ozel_nobet,
                        p.son_nobet_tarihi.strftime("%d.%m.%Y")
                        if p.son_nobet_tarihi else None,
                        p.sicil,
                    ),
                )
        guncellenen = [p.sicil for p in personel_list if p.ay_ici_nobet > 0]
        audit.log(
            "Nöbet Puanları Güncellendi",
            target_id="TOPLU",
            details=f"{len(guncellenen)} personelin puan/sayaçları güncellendi. "
                    f"Siciller: {', '.join(guncellenen[:10])}"
                    + (" ..." if len(guncellenen) > 10 else ""),
        )

    # ── Mazeret CRUD ──────────────────────────────────────────────────────────

    def get_all_mazeretler(self) -> list[dict]:
        with self._connect() as conn:
            rows = conn.execute(
                "SELECT m.*, p.ad_soyad AS personel_adi "
                "FROM mazeretler m LEFT JOIN personel p ON m.sicil = p.sicil "
                "ORDER BY m.baslangic DESC"
            ).fetchall()
            return [dict(r) for r in rows]

    def validate_mazeret(self, sicil: str, baslangic: date, bitis: date,
                         exclude_id: Optional[int] = None) -> str:
        """Girdi doğrulama: hata varsa mesaj, geçerliyse '' döner."""
        if bitis < baslangic:
            return "Bitiş tarihi başlangıç tarihinden önce olamaz."

        # Aynı personel için tarih çakışması kontrolü
        with self._connect() as conn:
            query = (
                "SELECT id, baslangic, bitis FROM mazeretler "
                "WHERE sicil = ? AND NOT (bitis < ? OR baslangic > ?)"
            )
            params: list = [sicil, baslangic.isoformat(), bitis.isoformat()]
            if exclude_id is not None:
                query += " AND id != ?"
                params.append(exclude_id)
            rows = conn.execute(query, params).fetchall()

        if rows:
            c = rows[0]
            return (
                f"Bu personel için {c['baslangic']}–{c['bitis']} "
                "tarihleriyle çakışan bir mazeret kaydı zaten mevcut."
            )
        return ""

    def add_mazeret(self, data: dict) -> str:
        """Yeni mazeret/izin kaydı ekler."""
        sicil = str(data.get("sicil", "")).strip()
        if not sicil:
            return "Sicil numarası boş olamaz."

        baslangic = _to_date(data.get("baslangic"))
        bitis = _to_date(data.get("bitis"))
        if not baslangic or not bitis:
            return "Geçerli tarih giriniz (GG.AA.YYYY)."

        err = self.validate_mazeret(sicil, baslangic, bitis)
        if err:
            return err

        with self._connect() as conn:
            conn.execute(
                """INSERT INTO mazeretler
                   (sicil, ad_soyad, baslangic, bitis,
                    mazeret_turu, mazeret_kodu, aciklama)
                   VALUES(?,?,?,?,?,?,?)""",
                (
                    sicil,
                    data.get("ad_soyad", ""),
                    baslangic.isoformat(),
                    bitis.isoformat(),
                    data.get("mazeret_turu", "İzin"),
                    data.get("mazeret_kodu", ""),
                    data.get("aciklama", ""),
                ),
            )
        audit.log(
            "Mazeret / İzin Eklendi",
            target_id=sicil,
            details=(
                f"{data.get('mazeret_turu', 'İzin')} | "
                f"{baslangic.strftime('%d.%m.%Y')} → {bitis.strftime('%d.%m.%Y')} | "
                f"Kod: {data.get('mazeret_kodu', '')} | "
                f"Açıklama: {data.get('aciklama', '') or '—'}"
            ),
        )
        return ""

    def delete_mazeret(self, mazeret_id: int) -> None:
        # Silmeden önce kaydı oku (log için)
        with self._connect() as conn:
            row = conn.execute(
                "SELECT * FROM mazeretler WHERE id = ?", (mazeret_id,)
            ).fetchone()
            conn.execute("DELETE FROM mazeretler WHERE id = ?", (mazeret_id,))
        if row:
            audit.log_warning(
                "Mazeret / İzin Silindi",
                target_id=row["sicil"],
                details=(
                    f"ID#{mazeret_id} | {row['mazeret_turu']} | "
                    f"{row['baslangic']} → {row['bitis']}"
                ),
            )

    # ── Scheduler API ─────────────────────────────────────────────────────────

    def get_scheduler_personel(self) -> list[Personel]:
        """Veritabanındaki aktif personeli scheduler.Personel listesine dönüştürür."""
        result = []
        for r in self.get_all_personel():
            result.append(Personel(
                sicil=r["sicil"],
                rutbe=r["rutbe"],
                ad_soyad=r["ad_soyad"],
                kidem_yili=int(r["kidem_yili"]),
                muafiyet_turu=r["muafiyet_turu"],
                grup_adi=r["grup_adi"],
                toplam_puan=float(r["toplam_puan"]),
                hafta_ici=int(r["hafta_ici"]),
                cuma=int(r["cuma"]),
                cumartesi=int(r["cumartesi"]),
                pazar=int(r["pazar"]),
                ozel_nobet=int(r["ozel_nobet"]),
                son_nobet_tarihi=_to_date(r.get("son_nobet_tarihi")),
            ))
        return result

    def get_scheduler_mazeretler(self, yil: int, ay: int) -> dict[str, list[date]]:
        """Verilen aya ait mazeretleri {sicil: [date, ...]} formatında döner."""
        _, days_in_month = calendar.monthrange(yil, ay)
        ay_bas = date(yil, ay, 1)
        ay_bit = date(yil, ay, days_in_month)

        with self._connect() as conn:
            rows = conn.execute(
                "SELECT sicil, baslangic, bitis FROM mazeretler "
                "WHERE NOT (bitis < ? OR baslangic > ?)",
                (ay_bas.isoformat(), ay_bit.isoformat()),
            ).fetchall()

        result: dict[str, list[date]] = {}
        for row in rows:
            bas = _to_date(row["baslangic"])
            bit = _to_date(row["bitis"])
            if not bas or not bit:
                continue
            # Ay sınırlarına kırp
            current = max(bas, ay_bas)
            end = min(bit, ay_bit)
            while current <= end:
                result.setdefault(row["sicil"], []).append(current)
                current += timedelta(days=1)
        return result

    # ── Excel Senkronizasyonu ─────────────────────────────────────────────────

    def sync_from_excel(self, excel_path: str) -> tuple[int, int, int]:
        """
        Excel Personel_Havuzu'ndan DB'ye içe aktarım (yıkıcı olmayan).
        Döner: (eklendi, atlandı, hata)
        """
        from scheduler import load_personel_havuzu
        try:
            personel_list = load_personel_havuzu(excel_path)
        except Exception:
            return 0, 0, 1

        added = skipped = errors = 0
        for p in personel_list:
            if self.get_personel_by_sicil(p.sicil):
                skipped += 1
                continue
            err = self.add_personel({
                "sicil":           p.sicil,
                "rutbe":           p.rutbe,
                "ad_soyad":        p.ad_soyad,
                "kidem_yili":      p.kidem_yili,
                "muafiyet_turu":   p.muafiyet_turu,
                "grup_adi":        p.grup_adi,
                "toplam_puan":     p.toplam_puan,
                "hafta_ici":       p.hafta_ici,
                "cuma":            p.cuma,
                "cumartesi":       p.cumartesi,
                "pazar":           p.pazar,
                "ozel_nobet":      p.ozel_nobet,
                "son_nobet_tarihi": (
                    p.son_nobet_tarihi.strftime("%d.%m.%Y")
                    if p.son_nobet_tarihi else None
                ),
            })
            if err:
                errors += 1
            else:
                added += 1
        audit.log(
            "Excel'den İçe Aktarım",
            target_id="TOPLU",
            details=f"Eklendi: {added} | Atlandı: {skipped} | Hata: {errors} | Kaynak: {excel_path}",
            severity="BİLGİ" if errors == 0 else "UYARI",
        )
        return added, skipped, errors

    def integrity_check(self, excel_path: str) -> list[str]:
        """
        DB ile Excel arasındaki personel uyumsuzluklarını listeler.
        Uygulama açılışında çalıştırılır; boş liste = sorun yok.
        """
        from scheduler import load_personel_havuzu
        warnings: list[str] = []
        try:
            excel_p = {p.sicil: p for p in load_personel_havuzu(excel_path)}
        except Exception as exc:
            return [f"Excel okunamadı: {exc}"]

        db_p = {r["sicil"]: r for r in self.get_all_personel()}

        for sicil, p in excel_p.items():
            if sicil not in db_p:
                warnings.append(
                    f"Excel'de bulunan '{p.ad_soyad}' ({sicil}) veritabanında yok."
                )
        for sicil, r in db_p.items():
            if sicil not in excel_p:
                warnings.append(
                    f"Veritabanındaki '{r['ad_soyad']}' ({sicil}) Excel'de bulunamadı."
                )
        return warnings


# Modül düzeyinde tek örnek
db = DatabaseManager()
