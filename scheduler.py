# -*- coding: utf-8 -*-
"""
scheduler.py — Nöbet Dağıtım Motoru
JGY 64-4 (B) Jandarma Genel Komutanlığı Nöbet/Vardiyalı Nöbet Hizmetleri Yönergesi

Bağımlılıklar: openpyxl (standart kurulum dışında başka kütüphane gerekmez)

Kullanım:
    py -3.12 scheduler.py                      # Haziran 2026 (varsayılan)
    py -3.12 scheduler.py 2026 7               # Temmuz 2026
    py -3.12 scheduler.py 2026 6 path/to/file  # Özel Excel yolu
"""

from __future__ import annotations

import sys
import os
import calendar
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from typing import Optional

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("HATA: openpyxl kurulu değil.  >  py -3.12 -m pip install openpyxl")


# ═══════════════════════════════════════════════════════════════════════════════
# BLOK 1 — VERİ MODELLERİ
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class Personel:
    """
    Hem Excel'den okunan birikimli sayaçları hem de algoritmanın
    runtime sırasında güncellediği anlık durumu barındırır.
    """
    sicil:            str
    rutbe:            str
    ad_soyad:         str
    kidem_yili:       int
    muafiyet_turu:    str   # Daimi | Gece | Hamile | Kadın-Ayda1 | Yok
    grup_adi:         str

    # ── Birikimli sayaçlar (Excel Personel_Havuzu'ndan yüklenir) ─────────────
    toplam_puan:      float
    hafta_ici:        int
    cuma:             int
    cumartesi:        int
    pazar:            int
    ozel_nobet:       int   # Excel'deki birleşik Özel sütunu (Özel-1+2+3)
    son_nobet_tarihi: Optional[date]

    # ── Alt özel sayaçlar (runtime; Excel'de ayrı sütun yoktur) ──────────────
    ozel1: int = 0
    ozel2: int = 0
    ozel3: int = 0
    yedek: int = 0

    # ── Runtime anlık durum (her schedule_month() çağrısında sıfırdan) ───────
    ay_ici_nobet:  int        = 0
    atanan_gunler: list[date] = field(default_factory=list)

    def __repr__(self) -> str:
        return f"<Personel {self.sicil} {self.rutbe} {self.ad_soyad}>"


@dataclass
class NobetAtama:
    """Bir güne atanan tek nöbet kaydı."""
    tarih:         date
    kategori:      str         # Ozel3 | Ozel2 | Ozel1 | Cmt | Pzr | Cuma | HI
    puan:          int
    atanan:        Personel
    yedekler:      list[Personel]   # min 2
    gap_kullanilan: int             # fiilen uygulanan gün aralığı


# ═══════════════════════════════════════════════════════════════════════════════
# BLOK 2 — EXCEL OKUYUCU
# ═══════════════════════════════════════════════════════════════════════════════

def _to_date(value) -> Optional[date]:
    """
    Hücredeki her türlü değeri güvenle date nesnesine çevirir.

    Desteklenen giriş tipleri:
      • datetime  — openpyxl'in otomatik dönüştürdüğü format (data_only=True)
      • date      — saf tarih nesnesi
      • str       — "DD.MM.YYYY", "YYYY-MM-DD", "DD/MM/YYYY" formatları
      • int/float — Excel seri numarası (openpyxl dönüştüremezse fallback)

    Önemli: Python'da datetime, date'in alt sınıfıdır.
    isinstance(dt_obj, date) → True döner; bu yüzden datetime kontrolü önce gelmelidir.
    """
    if value is None:
        return None

    # datetime önce kontrol edilmeli (date'in alt sınıfı olduğundan sıra kritik)
    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if isinstance(value, str):
        metin = value.strip()
        if not metin:
            return None
        for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y.%m.%d"):
            try:
                return datetime.strptime(metin, fmt).date()
            except ValueError:
                continue
        return None

    # int veya float → Excel seri numarası (openpyxl read_only=False durumunda nadir görülür)
    if isinstance(value, (int, float)):
        try:
            from openpyxl.utils.datetime import from_excel
            result = from_excel(int(value))
            if isinstance(result, datetime):
                return result.date()
            if isinstance(result, date):
                return result
        except Exception:
            pass

    return None


def load_personel_havuzu(path: str) -> list[Personel]:
    """
    Personel_Havuzu sayfasını okur.
    Satır 1 = ana başlık, Satır 2 = sütun başlıkları, Satır 3+ = veri.
    Sütun sırası (create_excel_template.py ile birebir eşleşir):
    A=Sicil  B=Rütbe  C=AdSoyad  D=KıdemYılı  E=MuafiyetTürü  F=GrupAdı
    G=ToplamPuan  H=Haftaİçi  I=Cuma  J=Cumartesi  K=Pazar  L=ÖzelNöbet
    M=SonNöbetTarihi
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Personel_Havuzu"]
    personeller: list[Personel] = []

    for row in ws.iter_rows(min_row=3, values_only=True):
        sicil = row[0]
        if not sicil:
            continue
        sicil = str(sicil).strip()
        # Açıklama / boşluk satırlarını atla:
        # Gerçek sicil numaraları kısa ve alfanümeriktir (≤20 karakter).
        # Açıklama satırlarındaki metin (📌 PUAN AĞIRLIKLARI: …) çok uzundur.
        if not sicil or len(sicil) > 20:
            continue
        rutbe         = str(row[1]).strip()  if row[1]  else "Bilinmiyor"
        ad_soyad      = str(row[2]).strip()  if row[2]  else ""
        kidem_yili    = int(row[3])          if row[3]  else 0
        muafiyet_turu = str(row[4]).strip()  if row[4]  else "Yok"
        grup_adi      = str(row[5]).strip()  if row[5]  else "Genel"
        toplam_puan   = float(row[6])        if row[6]  else 0.0
        hafta_ici     = int(row[7])          if row[7]  else 0
        cuma          = int(row[8])          if row[8]  else 0
        cumartesi     = int(row[9])          if row[9]  else 0
        pazar         = int(row[10])         if row[10] else 0
        ozel_nobet    = int(row[11])         if row[11] else 0
        son_nobet     = _to_date(row[12])

        personeller.append(Personel(
            sicil=sicil,
            rutbe=rutbe,
            ad_soyad=ad_soyad,
            kidem_yili=kidem_yili,
            muafiyet_turu=muafiyet_turu,
            grup_adi=grup_adi,
            toplam_puan=toplam_puan,
            hafta_ici=hafta_ici,
            cuma=cuma,
            cumartesi=cumartesi,
            pazar=pazar,
            ozel_nobet=ozel_nobet,
            son_nobet_tarihi=son_nobet,
        ))

    wb.close()
    return personeller


def load_aylik_mazeretler(path: str) -> dict[str, list[date]]:
    """
    Aylık_Mazeretler sayfasını okur.
    Döner: {sicil: [mazeret_tarihi, ...]}

    Sütun düzeni (create_excel_template.py ile eşleşir):
      A=Sicil  B=AdSoyad  C=BaşlangıçTarihi  D=BitişTarihi
      E=MazeretTürü  F=MazeretKodu  G=Açıklama

    C ile D arasındaki tüm günler (her iki uç dahil) mazeretli sayılır.
    D sütunu boşsa sadece C günü mazeretlidir (tek günlük kayıtlara geriye dönük uyumluluk).
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Aylık_Mazeretler"]
    mazeretler: dict[str, list[date]] = {}

    for row in ws.iter_rows(min_row=4, values_only=True):
        sicil = row[0]
        baslangic_ham = row[2]
        if not sicil or not baslangic_ham:
            continue
        sicil = str(sicil).strip()
        baslangic = _to_date(baslangic_ham)
        if not baslangic:
            continue

        # D sütunu (index 3) dolu → aralık; boş → tek gün
        bitis_ham = row[3] if len(row) > 3 else None
        bitis = _to_date(bitis_ham)
        if bitis is None or bitis < baslangic:
            bitis = baslangic

        gun = baslangic
        while gun <= bitis:
            mazeretler.setdefault(sicil, []).append(gun)
            gun += timedelta(days=1)

    wb.close()
    return mazeretler


def load_istisnai_tatiller(path: str) -> dict[date, tuple[str, int]]:
    """
    İstisnai_Tatiller sayfasını okur.
    Döner: {tarih: (kategori_kodu, puan_degeri)}
    Örnek: {date(2026,5,25): ("Ozel1", 5)}
    Sütun A=Tarih  B=Açıklama  C=Kategori  D=PuanDeğeri
    Kategori eşlemesi: "Özel-1"→"Ozel1", "Özel-2"→"Ozel2", "Özel-3"→"Ozel3"
    """
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["İstisnai_Tatiller"]
    tatiller: dict[date, tuple[str, int]] = {}

    _kat_map = {
        "özel-1": "Ozel1", "ozel-1": "Ozel1", "ozel1": "Ozel1",
        "özel-2": "Ozel2", "ozel-2": "Ozel2", "ozel2": "Ozel2",
        "özel-3": "Ozel3", "ozel-3": "Ozel3", "ozel3": "Ozel3",
    }

    for row in ws.iter_rows(min_row=4, values_only=True):
        tarih_ham = row[0]
        kategori_ham = row[2]
        puan_ham = row[3]
        if not tarih_ham or not kategori_ham:
            continue
        tarih = _to_date(tarih_ham)
        if not tarih:
            continue
        kat_str = str(kategori_ham).strip().lower()
        kategori = _kat_map.get(kat_str, "Ozel1")
        puan = int(puan_ham) if puan_ham else 5
        tatiller[tarih] = (kategori, puan)

    wb.close()
    return tatiller


# ═══════════════════════════════════════════════════════════════════════════════
# BLOK 3 — TAKVİM MOTORU
# ═══════════════════════════════════════════════════════════════════════════════

# Nöbet puan değerleri (kurallar.txt §6)
PUAN_HARITASI: dict[str, int] = {
    "Ozel3": 7,
    "Ozel2": 6,
    "Ozel1": 5,
    "Cmt":   4,
    "Pzr":   3,
    "Cuma":  2,
    "HI":    1,
}

# Yazım öncelik sırası — yüksek puanlıdan başla (kurallar.txt §7)
YAZIM_SIRASI: list[str] = ["Ozel3", "Ozel2", "Ozel1", "Cmt", "Pzr", "Cuma", "HI"]

# Türkçe ay adları (loglarda kullanılır)
AY_ADLARI: dict[int, str] = {
    1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan",
    5: "Mayıs", 6: "Haziran", 7: "Temmuz", 8: "Ağustos",
    9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık",
}


def gun_kategorisi(
    tarih: date,
    ozel_tatiller: dict[date, tuple[str, int]],
) -> str:
    """
    Bir tarihin nöbet kategorisini döner.
    Öncelik sırası:
      1. İstisnai_Tatiller sayfasında kayıtlıysa → orada belirtilen kategori
      2. weekday() kontrolü: 4=Cuma, 5=Cmt, 6=Pzr, 0-3=HI
    """
    if tarih in ozel_tatiller:
        return ozel_tatiller[tarih][0]
    wd = tarih.weekday()
    if wd == 4:
        return "Cuma"
    if wd == 5:
        return "Cmt"
    if wd == 6:
        return "Pzr"
    return "HI"


def ay_takvimi_olustur(
    yil: int,
    ay: int,
    ozel_tatiller: dict[date, tuple[str, int]],
) -> dict[str, list[date]]:
    """
    Belirtilen ay için her kategoriye ait gün listelerini döner.
    Örnek: {"Ozel3": [], "Cmt": [date(2026,6,6), date(2026,6,13), ...], ...}
    """
    takvim: dict[str, list[date]] = {k: [] for k in YAZIM_SIRASI}
    _, gun_sayisi = calendar.monthrange(yil, ay)
    for gun_no in range(1, gun_sayisi + 1):
        tarih = date(yil, ay, gun_no)
        kat = gun_kategorisi(tarih, ozel_tatiller)
        takvim[kat].append(tarih)
    return takvim


def _onceki_ay_son_gun(yil: int, ay: int) -> date:
    """Verilen ayın bir önceki ayının son günü."""
    birinci = date(yil, ay, 1)
    return birinci - timedelta(days=1)


def onceki_ay_son_3_gun(yil: int, ay: int) -> list[date]:
    """Önceki ayın son 3 gününü döner (ay geçişi kısıt kontrolünde kullanılır)."""
    son = _onceki_ay_son_gun(yil, ay)
    return [son - timedelta(days=i) for i in range(2, -1, -1)]


# ═══════════════════════════════════════════════════════════════════════════════
# BLOK 4 — KISIT DENETÇİSİ
# ═══════════════════════════════════════════════════════════════════════════════

def _bayram_gun_mu(tarih: date, ozel_tatiller: dict[date, tuple[str, int]]) -> bool:
    """Tarih, Özel-2 veya Özel-3 kategorisinde bayram günü mü?"""
    if tarih in ozel_tatiller:
        return ozel_tatiller[tarih][0] in ("Ozel2", "Ozel3")
    return False


def is_eligible(
    p: Personel,
    tarih: date,
    mazeretler: dict[str, list[date]],
    ozel_tatiller: dict[date, tuple[str, int]],
    min_gap: int,
    yil: int,
    ay: int,
) -> bool:
    """
    Bir personelin belirtilen tarih için nöbet tutup tutamayacağını kontrol eder.
    Tüm 7 kısıt sırasıyla denetlenir; herhangi biri False döndürürse personel elenir.

    K1 — Kalıcı muafiyet (Daimi veya Hamile)
    K2 — Mazeret çizelgesinde o tarih var
    K3 — 48 saat kuralı (son nöbetten 2 gün geçmeden yeni nöbet yok)
    K4 — Ay geçişi koruması (önceki ayın son 3 günü nöbet → bu ayın ilk 3 günü kısıt)
    K5 — Gün aralığı (min_gap günden daha kısa süre önce nöbet tutmuşsa)
    K6 — Kadın-Ayda1 kısıtı (o ay içinde zaten 1 nöbet varsa)
    K7 — Bayram geçiş kuralı (önceki bayramda nöbet → bu bayramda yazılamaz)
    """
    # K1 — Kalıcı muafiyet
    if p.muafiyet_turu in ("Daimi", "Hamile"):
        return False

    # K2 — Mazeret
    kisi_mazeretler = mazeretler.get(p.sicil, [])
    if tarih in kisi_mazeretler:
        return False

    # K3 — 48 saat kuralı (yalnızca geriye dönük):
    # Algoritma kategorileri kronolojik sırayla DEĞİL öncelik sırasıyla yazar
    # (Cmt→Pzr→Cuma→HI). Bu nedenle atanan_gunler ve son_nobet_tarihi içinde
    # tarih'ten büyük (gelecek) değerler olabilir; bunlar 48h kısıtını tetiklememeli.
    # Sadece tarih'ten KÜÇÜK (geçmiş) en yakın göreve bakarız.
    geriye_atanmis = [d for d in p.atanan_gunler if d < tarih]
    if geriye_atanmis:
        en_son_gecmis = max(geriye_atanmis)
        if (tarih - en_son_gecmis).days < 2:
            return False
    elif p.son_nobet_tarihi is not None and p.son_nobet_tarihi < tarih:
        # Bu ay içinde tarih'ten önceki atama yok; Excel'deki son tarihe bak
        if (tarih - p.son_nobet_tarihi).days < 2:
            return False

    # K4 — Ay geçişi koruması (kurallar.txt §8):
    # "Son 3 gün nöbet tutan personele bir sonraki ayın ilk 3 gününe otomatik kısıt"
    # Sadece son_nobet_tarihi ÖNCEKİ aya ait ve o ayın son 3 günündeyse geçerli.
    # Runtime'da son_nobet_tarihi bu ayın tarihlerine güncellendiğinden ay kontrolü zorunludur.
    if tarih.day <= 3 and p.son_nobet_tarihi is not None:
        onceki_ay_sonu = date(yil, ay, 1) - timedelta(days=1)   # Önceki ayın son günü
        if (p.son_nobet_tarihi.year  == onceki_ay_sonu.year
                and p.son_nobet_tarihi.month == onceki_ay_sonu.month
                and (onceki_ay_sonu - p.son_nobet_tarihi).days <= 2):
            return False

    # K5 — Gün aralığı kısıtı (9→2 arası kademeli gap)
    for atanmis in p.atanan_gunler:
        if abs((tarih - atanmis).days) < min_gap:
            return False

    # K6 — Kadın-Ayda1: aylık maksimum 1 nöbet
    if p.muafiyet_turu == "Kadın-Ayda1" and p.ay_ici_nobet >= 1:
        return False

    # K7 — Bayram geçiş kuralı:
    # Hedef gün bir bayram günüyse, personelin son_nobet_tarihi de bir bayram günüyse → yasak
    if _bayram_gun_mu(tarih, ozel_tatiller) and p.son_nobet_tarihi is not None:
        if _bayram_gun_mu(p.son_nobet_tarihi, ozel_tatiller):
            return False

    return True


# ═══════════════════════════════════════════════════════════════════════════════
# BLOK 5 — ADAY SEÇİCİ
# ═══════════════════════════════════════════════════════════════════════════════

# Rütbe kıymet tablosu — yüksek sayı = kıdemli (kurallar.txt §10)
RUTBE_KIYMETI: dict[str, int] = {
    "Orgeneral":          18,
    "Korgeneral":         17,
    "Tümgeneral":         16,
    "Tuğgeneral":         15,
    "Albay":              14,
    "Yarbay":             13,
    "Binbaşı":            12,
    "Yüzbaşı":            11,
    "Üsteğmen":           10,
    "Teğmen":              9,
    "Asteğmen":            8,
    "Kd.Başçavuş":         7,
    "Başçavuş":            6,
    "Üstçavuş":            5,
    "Kıdemli Üstçavuş":    5,
    "Çavuş":               4,
    "Uzman Jandarma":      4,
    "Uzman Erbaş":         3,
    "Sözleşmeli Erbaş/Er": 2,
    "Erbaş/Er":            1,
    "Memur":               2,
    "Bilinmiyor":          0,
}

def _rutbe_kiymeti(p: Personel) -> int:
    """Bilinmeyen rütbeler için varsayılan 0 döner."""
    return RUTBE_KIYMETI.get(p.rutbe, 0)


def sort_key_factory(kategori: str):
    """
    Kategori başına doğru sıralama anahtarı döndürür.
    Tüm kriterler ascending: küçük = daha önce atanır.
    Eşitlik bozucu: (-rutbe_kiymeti, -kidem_yili) — kıdemlinin aleyhine (kurallar.txt §7 son kural).
    """
    def _tiebreak(p: Personel):
        return (-_rutbe_kiymeti(p), -p.kidem_yili)

    if kategori in ("Ozel3", "Ozel2", "Ozel1"):
        def key_ozel(p: Personel):
            return (
                p.ozel1 + p.ozel2 + p.ozel3,   # toplam özel nöbet (asc)
                p.ozel3,                          # Özel-3 alt kırılım (asc)
                p.ozel2,                          # Özel-2 alt kırılım (asc)
                p.ozel1,                          # Özel-1 alt kırılım (asc)
                p.yedek,                          # yedek sayısı (asc)
                p.ay_ici_nobet,                   # ay içi toplam (asc)
                *_tiebreak(p),
            )
        return key_ozel

    if kategori == "Cmt":
        def key_cmt(p: Personel):
            return (
                p.cumartesi,      # 1. öncelik: Cumartesi nöbet sayısı
                p.toplam_puan,    # 2. öncelik: toplam puan (tie-breaker)
                *_tiebreak(p),
            )
        return key_cmt

    if kategori == "Pzr":
        def key_pzr(p: Personel):
            return (
                p.ay_ici_nobet,
                p.pazar,
                p.cumartesi,
                p.yedek,
                p.toplam_puan,
                *_tiebreak(p),
            )
        return key_pzr

    if kategori == "Cuma":
        def key_cuma(p: Personel):
            return (
                p.ay_ici_nobet,
                p.cuma,
                p.cumartesi,
                p.pazar,
                p.yedek,
                p.toplam_puan,
                *_tiebreak(p),
            )
        return key_cuma

    # Hafta içi (HI) — varsayılan
    def key_hi(p: Personel):
        return (
            p.hafta_ici,      # 1. öncelik: Hafta içi nöbet sayısı
            p.toplam_puan,    # 2. öncelik: toplam puan (tie-breaker)
            *_tiebreak(p),
        )
    return key_hi


# ═══════════════════════════════════════════════════════════════════════════════
# BLOK 6 — SIKIŞTRIMA DÖNGÜsÜ (Gap Controller)
# ═══════════════════════════════════════════════════════════════════════════════

def _nobet_yaz_gun(
    personel_listesi: list[Personel],
    tarih: date,
    kategori: str,
    mazeretler: dict[str, list[date]],
    ozel_tatiller: dict[date, tuple[str, int]],
    min_gap: int,
    yil: int,
    ay: int,
    hariç_tut: Optional[Personel] = None,
) -> Optional[Personel]:
    """
    Belirtilen gün için uygun adayları filtreler, sıralar ve birinci adayı döner.
    hariç_tut: yedek seçiminde asıl nöbetçiyi listeden çıkarmak için kullanılır.
    """
    adaylar = [
        p for p in personel_listesi
        if p is not hariç_tut
        and is_eligible(p, tarih, mazeretler, ozel_tatiller, min_gap, yil, ay)
    ]
    if not adaylar:
        return None
    return sorted(adaylar, key=sort_key_factory(kategori))[0]


def gap_ile_yaz(
    personel_listesi: list[Personel],
    tarih: date,
    kategori: str,
    mazeretler: dict[str, list[date]],
    ozel_tatiller: dict[date, tuple[str, int]],
    yil: int,
    ay: int,
) -> tuple[Optional[Personel], int]:
    """
    Gün aralığını 9'dan 2'ye kadar kademeli olarak azaltarak uygun aday arar.
    Hiçbir gap'te bulunamazsa gap=1 ile son çare denemesi yapar.
    Döner: (atanan_personel_veya_None, kullanilan_gap)
    """
    for gap in range(9, 1, -1):      # 9 → 8 → 7 → … → 2
        atanan = _nobet_yaz_gun(
            personel_listesi, tarih, kategori,
            mazeretler, ozel_tatiller, gap, yil, ay,
        )
        if atanan:
            return atanan, gap

    # Gap=1: en son çare — sadece 48h (K3) ve kalıcı muafiyet (K1) korunur,
    # gün aralığı kısıtı fiilen kaldırılmış olur
    atanan = _nobet_yaz_gun(
        personel_listesi, tarih, kategori,
        mazeretler, ozel_tatiller, 1, yil, ay,
    )
    return atanan, 1


# ═══════════════════════════════════════════════════════════════════════════════
# BLOK 7 — ANA ZAMANLAYICI
# ═══════════════════════════════════════════════════════════════════════════════

def _guncelle_puan_sayaclari(p: Personel, kategori: str) -> None:
    """
    Atama gerçekleşince personelin ilgili kategori sayacını ve toplam puanını artırır.
    ay_ici_nobet'i burada YÜKSELTMİYOR — schedule_month() bunu ayrıca yönetir.
    """
    puan = PUAN_HARITASI[kategori]
    p.toplam_puan += puan

    if kategori == "HI":
        p.hafta_ici += 1
    elif kategori == "Cuma":
        p.cuma += 1
    elif kategori == "Cmt":
        p.cumartesi += 1
    elif kategori == "Pzr":
        p.pazar += 1
    elif kategori == "Ozel1":
        p.ozel1 += 1
        p.ozel_nobet += 1
    elif kategori == "Ozel2":
        p.ozel2 += 1
        p.ozel_nobet += 1
    elif kategori == "Ozel3":
        p.ozel3 += 1
        p.ozel_nobet += 1


def yedek_sec(
    personel_listesi: list[Personel],
    tarih: date,
    asil_nobetci: Personel,
    mazeretler: dict[str, list[date]],
    ozel_tatiller: dict[date, tuple[str, int]],
    yil: int,
    ay: int,
    min_yedek: int = 2,
) -> list[Personel]:
    """
    Asıl nöbetçi dışında, kısıt K1 (muafiyet) ve K2 (mazeret) geçenleri
    yedek sıralamasına göre seçer. Kural: en az yedek puanı olan önce.
    Asgari 2 yedek personel gerekir (kurallar.txt §9).
    """
    def yedek_sort_key(p: Personel):
        return (
            p.yedek,
            _rutbe_kiymeti(p),    # yedekte kıdemce geride olan önce (asc)
            p.kidem_yili,
        )

    # Birinci geçiş: tam is_eligible (gap=2) + aylık 3 yedek sınırı
    adaylar = [
        p for p in personel_listesi
        if p is not asil_nobetci
        and p.yedek < 3
        and is_eligible(p, tarih, mazeretler, ozel_tatiller, 2, yil, ay)
    ]

    # Yeterli aday çıkmazsa fallback: sadece K1 (muafiyet) + K2 (mazeret) kontrolü
    if len(adaylar) < min_yedek:
        adaylar = [
            p for p in personel_listesi
            if p is not asil_nobetci
            and p.muafiyet_turu not in ("Daimi", "Hamile")
            and tarih not in mazeretler.get(p.sicil, [])
        ]

    adaylar_sirali = sorted(adaylar, key=yedek_sort_key)
    secilen = adaylar_sirali[:max(min_yedek, 2)]

    # Seçilen yedeklerin sayacını güncelle
    for yed in secilen:
        yed.yedek += 1

    return secilen


def schedule_month(
    yil: int,
    ay: int,
    excel_path: str,
) -> tuple[list[NobetAtama], list[Personel], list[tuple[date, str]]]:
    """
    Hedef ay için nöbet listesini oluşturur.
    Tüm YAZIM_SIRASI kategorilerini sırayla işler.
    Döner: Tarih sırasına göre sıralanmış NobetAtama listesi.
    """
    # ── Verileri yükle ────────────────────────────────────────────────────────
    personeller  = load_personel_havuzu(excel_path)
    mazeretler   = load_aylik_mazeretler(excel_path)
    ozel_tatiller = load_istisnai_tatiller(excel_path)
    takvim       = ay_takvimi_olustur(yil, ay, ozel_tatiller)

    aktif_personel = [p for p in personeller if p.muafiyet_turu != "Daimi"]

    if not aktif_personel:
        print("UYARI: Aktif (muafiyet_turu=Daimi olmayan) personel bulunamadı.")
        return []

    # ── Yeni personel için ortalama puan başlangıcı ───────────────────────────
    # toplam_puan == 0 olan personelin başlangıç puanı, mevcut personellerin
    # ortalamasına eşitlenerek üst üste nöbet yığılması engellenir.
    puanli = [p for p in aktif_personel if p.toplam_puan > 0]
    if puanli:
        ortalama_puan = sum(p.toplam_puan for p in puanli) / len(puanli)
        for p in aktif_personel:
            if p.toplam_puan == 0:
                p.toplam_puan = ortalama_puan

    sonuclar: list[NobetAtama] = []
    atanmayan_gunler: list[tuple[date, str]] = []

    # ── Kategori bazlı yazım (öncelik sırasına göre) ──────────────────────────
    for kategori in YAZIM_SIRASI:
        gunler = sorted(takvim[kategori])
        for tarih in gunler:
            atanan, gap = gap_ile_yaz(
                aktif_personel, tarih, kategori,
                mazeretler, ozel_tatiller, yil, ay,
            )

            if atanan is None:
                atanmayan_gunler.append((tarih, kategori))
                continue

            # Sayaçları güncelle
            atanan.ay_ici_nobet += 1
            atanan.atanan_gunler.append(tarih)
            atanan.son_nobet_tarihi = tarih
            _guncelle_puan_sayaclari(atanan, kategori)

            yedekler = yedek_sec(
                aktif_personel, tarih, atanan,
                mazeretler, ozel_tatiller, yil, ay,
            )

            sonuclar.append(NobetAtama(
                tarih=tarih,
                kategori=kategori,
                puan=PUAN_HARITASI[kategori],
                atanan=atanan,
                yedekler=yedekler,
                gap_kullanilan=gap,
            ))

    # Tarihe göre sırala
    sonuclar.sort(key=lambda a: a.tarih)

    return sonuclar, aktif_personel, atanmayan_gunler


def schedule_month_from_data(
    yil: int,
    ay: int,
    personel_list: list[Personel],
    mazeretler: dict[str, list[date]],
    ozel_tatiller: dict[date, tuple[str, int]],
) -> tuple[list[NobetAtama], list[Personel], list[tuple[date, str]]]:
    """
    schedule_month() ile aynı mantığı uygular; ancak veriyi Excel'den okumak
    yerine çağıran taraftan (örn. data_manager.py) hazır alır.
    Bu sayede SQLite tabanlı veri akışı desteklenir.
    """
    takvim = ay_takvimi_olustur(yil, ay, ozel_tatiller)
    aktif_personel = [p for p in personel_list if p.muafiyet_turu != "Daimi"]

    if not aktif_personel:
        print("UYARI: Aktif personel bulunamadı (muafiyet_turu=Daimi olmayan).")
        return [], [], []

    # ── Yeni personel için ortalama puan başlangıcı ───────────────────────────
    puanli = [p for p in aktif_personel if p.toplam_puan > 0]
    if puanli:
        ortalama_puan = sum(p.toplam_puan for p in puanli) / len(puanli)
        for p in aktif_personel:
            if p.toplam_puan == 0:
                p.toplam_puan = ortalama_puan

    sonuclar: list[NobetAtama] = []
    atanmayan_gunler: list[tuple[date, str]] = []

    for kategori in YAZIM_SIRASI:
        gunler = sorted(takvim[kategori])
        for tarih in gunler:
            atanan, gap = gap_ile_yaz(
                aktif_personel, tarih, kategori,
                mazeretler, ozel_tatiller, yil, ay,
            )
            if atanan is None:
                atanmayan_gunler.append((tarih, kategori))
                continue

            atanan.ay_ici_nobet += 1
            atanan.atanan_gunler.append(tarih)
            atanan.son_nobet_tarihi = tarih
            _guncelle_puan_sayaclari(atanan, kategori)

            yedekler = yedek_sec(
                aktif_personel, tarih, atanan,
                mazeretler, ozel_tatiller, yil, ay,
            )
            sonuclar.append(NobetAtama(
                tarih=tarih,
                kategori=kategori,
                puan=PUAN_HARITASI[kategori],
                atanan=atanan,
                yedekler=yedekler,
                gap_kullanilan=gap,
            ))

    sonuclar.sort(key=lambda a: a.tarih)
    return sonuclar, aktif_personel, atanmayan_gunler


# ═══════════════════════════════════════════════════════════════════════════════
# BLOK 8 — TERMİNAL LOGGER
# ═══════════════════════════════════════════════════════════════════════════════

_KAT_ETIKET: dict[str, str] = {
    "Ozel3": "Özel-3",
    "Ozel2": "Özel-2",
    "Ozel1": "Özel-1",
    "Cmt":   "Cumartesi",
    "Pzr":   "Pazar",
    "Cuma":  "Cuma",
    "HI":    "Hafta İçi",
}

def log_nobet_listesi(
    atamalar: list[NobetAtama],
    aktif_personel: list[Personel],
    atanmayan_gunler: list[tuple[date, str]],
    yil: int,
    ay: int,
) -> None:
    """
    Nöbet listesini terminale düzenli tablo formatında basar.
    Özet istatistikleri, gap uyarılarını ve atanmayan günleri gösterir.
    """
    SEP   = "=" * 80
    SSEP  = "-" * 80
    ay_ad = AY_ADLARI.get(ay, str(ay))
    bos   = ""

    print(bos)
    print(SEP)
    print(f"  NÖBET LİSTESİ — {ay_ad} {yil}")
    print(f"  JGY 64-4 (B) — Nöbet Dağıtım Motoru")
    print(SEP)

    # ── Başlık satırı ──────────────────────────────────────────────────────────
    print(f"  {'TARİH':<14} {'KAT.':<12} {'P':>2}  {'PERSONEL':<24} {'RÜTBE':<20}  {'GAP':>3}  YEDEKLER")
    print(SSEP)

    # ── Atama satırları (tarih sırası) ─────────────────────────────────────────
    gap1_gunler: list[date] = []

    for a in atamalar:
        tarih_str  = a.tarih.strftime("%d.%m.%Y")
        kat_str    = _KAT_ETIKET.get(a.kategori, a.kategori)
        puan_str   = str(a.puan)
        ad_str     = a.atanan.ad_soyad[:24]
        rutbe_str  = a.atanan.rutbe[:20]
        gap_str    = str(a.gap_kullanilan)
        yed_str    = ", ".join(
            (y.ad_soyad.split() or ["?"])[-1] for y in a.yedekler
        ) if a.yedekler else "—"

        gap_uyari  = " (!)" if a.gap_kullanilan == 1 else "   "
        print(
            f"  {tarih_str:<14} {kat_str:<12} {puan_str:>2}  "
            f"{ad_str:<24} {rutbe_str:<20}  {gap_str:>3}{gap_uyari}  {yed_str}"
        )

        if a.gap_kullanilan == 1:
            gap1_gunler.append(a.tarih)

    print(SSEP)

    # ── Özet istatistikler ─────────────────────────────────────────────────────
    if atamalar:
        toplam_nobet = len(atamalar)
        nobetci_sayisi = len({a.atanan.sicil for a in atamalar})
        ort = toplam_nobet / nobetci_sayisi if nobetci_sayisi else 0

        # Kategori dağılımı
        kat_sayim: dict[str, int] = {}
        for a in atamalar:
            kat_sayim[a.kategori] = kat_sayim.get(a.kategori, 0) + 1

        print(f"\n  ÖZET:")
        print(f"    Toplam nöbet:      {toplam_nobet}")
        print(f"    Nöbet tutan kişi:  {nobetci_sayisi} / {len(aktif_personel)}")
        print(f"    Kişi başı ortalama: {ort:.2f}")
        print(f"\n    Kategori dağılımı:")
        for kat in YAZIM_SIRASI:
            if kat in kat_sayim:
                print(f"      {_KAT_ETIKET[kat]:<14}  {kat_sayim[kat]:>3} nöbet")

        # Personel başına nöbet sayısı tablosu
        print(f"\n    Personel bazlı dağılım:")
        print(f"    {'AD SOYAD':<26} {'RÜTBE':<20} {'AY İÇİ':>7} {'TOPLAM PUAN':>12}")
        print(f"    {'-'*26} {'-'*20} {'-'*7} {'-'*12}")

        # Sadece bu ay nöbet tutanları sırala
        nobetci_map: dict[str, Personel] = {a.atanan.sicil: a.atanan for a in atamalar}
        for p in sorted(nobetci_map.values(), key=lambda x: -x.ay_ici_nobet):
            print(
                f"    {p.ad_soyad:<26} {p.rutbe:<20} "
                f"{p.ay_ici_nobet:>7} {p.toplam_puan:>12.0f}"
            )

    # ── Uyarılar ───────────────────────────────────────────────────────────────
    print(bos)
    if gap1_gunler:
        tarih_listesi = ", ".join(g.strftime("%d.%m") for g in gap1_gunler)
        print(f"  ⚠  UYARI: gap=1 kullanıldı (asgari aralık sıkıştırması) → {len(gap1_gunler)} günde:")
        print(f"     {tarih_listesi}")

    if atanmayan_gunler:
        print(f"  ✗  HATA: {len(atanmayan_gunler)} güne nöbet ATANAMADI (uygun personel yok):")
        for t, k in atanmayan_gunler:
            print(f"     {t.strftime('%d.%m.%Y')}  [{_KAT_ETIKET.get(k, k)}]")

    if not gap1_gunler and not atanmayan_gunler:
        print("  ✓  Tüm günlere nöbet başarıyla atandı. Hiçbir kısıt ihlali yok.")

    print(SEP)
    print(bos)


# ═══════════════════════════════════════════════════════════════════════════════
# GİRİŞ NOKTASI
# ═══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    # ── Komut satırı argümanları ───────────────────────────────────────────────
    args = sys.argv[1:]

    try:
        yil = int(args[0]) if len(args) >= 1 else date.today().year
        ay  = int(args[1]) if len(args) >= 2 else date.today().month + 1
        if ay > 12:
            ay = 1
            yil += 1
    except (ValueError, IndexError):
        print("Kullanım: py scheduler.py [yil] [ay] [excel_yolu]")
        sys.exit(1)

    excel_path = args[2] if len(args) >= 3 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "personel_listesi.xlsx",
    )

    if not os.path.isfile(excel_path):
        print(f"HATA: Excel dosyası bulunamadı → {excel_path}")
        sys.exit(1)

    ay_ad = AY_ADLARI.get(ay, str(ay))
    print(f"\n  Nöbet listesi oluşturuluyor: {ay_ad} {yil}")
    print(f"  Kaynak: {excel_path}\n")

    # ── Zamanlayıcıyı çalıştır ────────────────────────────────────────────────
    sonuclar, aktif_personel, atanmayan = schedule_month(yil, ay, excel_path)

    # ── Sonuçları logla ───────────────────────────────────────────────────────
    log_nobet_listesi(sonuclar, aktif_personel, atanmayan, yil, ay)


if __name__ == "__main__":
    main()
