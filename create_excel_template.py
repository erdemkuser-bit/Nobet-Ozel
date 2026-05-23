# -*- coding: utf-8 -*-
"""
Nöbet Dağıtım Otomasyonu — Excel Şablonu Oluşturucu
JGY 64-4 (B) Jandarma Genel Komutanlığı Nöbet/Vardiyalı Nöbet Hizmetleri Yönergesine uygun
"""

from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.styles.numbers import FORMAT_DATE_DDMMYY
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import date
import os


# ─── Renk Paleti ─────────────────────────────────────────────────────────────
CLR = {
    "header_bg":    "1F3864",   # koyu lacivert — başlık arka planı
    "header_font":  "FFFFFF",   # beyaz — başlık yazısı
    "subheader_bg": "2E75B6",   # orta mavi — alt başlık
    "freeze_bg":    "D6E4F0",   # açık mavi — dondurulmuş ilk satır
    "row_odd":      "EBF3FB",   # çok açık mavi — tek satır
    "row_even":     "FFFFFF",   # beyaz — çift satır
    "accent_gold":  "F4B942",   # altın — özel vurgu
    "accent_red":   "C0392B",   # kırmızı — uyarı
    "accent_green": "1E8449",   # yeşil — onay
    "border_dark":  "1F3864",   # koyu kenarlık
    "border_light": "AED6F1",   # açık kenarlık
    "tab_blue":     "2196F3",
    "tab_orange":   "FF9800",
    "tab_green":    "4CAF50",
}

# ─── Yardımcı: Kenarlık ────────────────────────────────────────────────────
def thin_border(color="AED6F1"):
    s = Side(border_style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def thick_border(color="1F3864"):
    s = Side(border_style="medium", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

# ─── Yardımcı: Hücre Uygulayıcı ──────────────────────────────────────────
def style_header_cell(cell, text, bg=CLR["header_bg"], font_color=CLR["header_font"],
                       font_size=11, bold=True, wrap=True):
    cell.value = text
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.font = Font(name="Calibri", bold=bold, size=font_size, color=font_color)
    cell.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=wrap)
    cell.border = thick_border()

def style_data_cell(cell, value=None, row_idx=1, align="left", number_format=None):
    if value is not None:
        cell.value = value
    bg = CLR["row_odd"] if row_idx % 2 == 1 else CLR["row_even"]
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.font = Font(name="Calibri", size=10)
    cell.alignment = Alignment(horizontal=align, vertical="center")
    cell.border = thin_border()
    if number_format:
        cell.number_format = number_format

def set_col_width(ws, col_letter, width):
    ws.column_dimensions[col_letter].width = width

def set_row_height(ws, row, height):
    ws.row_dimensions[row].height = height


# ═══════════════════════════════════════════════════════════════════════════════
# SAYFA 1 — Personel_Havuzu
# ═══════════════════════════════════════════════════════════════════════════════
def build_personel_havuzu(wb: Workbook):
    ws = wb.create_sheet("Personel_Havuzu")
    ws.sheet_properties.tabColor = CLR["tab_blue"]

    # ── Başlık Satırı (Satır 1) ───────────────────────────────────────────────
    baslik_metni = "PERSONEL NÖBET HAVUZU — JGY 64-4 (B) Yönergesine Uygun"
    ws.merge_cells("A1:M1")
    cell = ws["A1"]
    cell.value = baslik_metni
    cell.fill = PatternFill("solid", fgColor=CLR["header_bg"])
    cell.font = Font(name="Calibri", bold=True, size=13, color=CLR["header_font"])
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thick_border()
    set_row_height(ws, 1, 36)

    # ── Sütun Başlıkları (Satır 2) ────────────────────────────────────────────
    sutunlar = [
        # (başlık,              genişlik, açıklama_tooltip)
        ("Sicil / ID",          14, "Personel sicil numarası veya benzersiz kimlik"),
        ("Rütbe",               16, "Subay/Astsubay/Uzman Erbaş vb."),
        ("Ad Soyad",            22, "Personelin tam adı"),
        ("Kıdem Yılı",          12, "Fiili hizmet yılı (muafiyet hesabında kullanılır)"),
        ("Muafiyet Türü",       17, "Daimi / Gece / Hamile / Yok"),
        ("Grup Adı",            18, "Dahil olduğu nöbet grubu (Normal, Vardiya-1 vb.)"),
        ("Toplam Puan",         13, "Toplam birikimli nöbet puanı"),
        ("Hafta İçi Nöbet",     16, "Toplam hafta içi nöbet sayısı (Puan=1)"),
        ("Cuma Nöbet",          13, "Toplam Cuma nöbet sayısı (Puan=2)"),
        ("Cumartesi Nöbet",     16, "Toplam Cumartesi nöbet sayısı (Puan=4)"),
        ("Pazar Nöbet",         13, "Toplam Pazar nöbet sayısı (Puan=3)"),
        ("Özel Nöbet",          13, "Özel-1+2+3 toplam nöbet sayısı (Puan=5/6/7)"),
        ("Son Nöbet Tarihi",    18, "Personelin en son nöbet tuttuğu tarih"),
    ]

    for col_idx, (baslik, genislik, _) in enumerate(sutunlar, start=1):
        col_letter = get_column_letter(col_idx)
        cell = ws.cell(row=2, column=col_idx)
        style_header_cell(cell, baslik, bg=CLR["subheader_bg"])
        set_col_width(ws, col_letter, genislik)

    set_row_height(ws, 2, 40)

    # ── Veri Doğrulama (Data Validation) ──────────────────────────────────────
    # Rütbe listesi
    rutbe_dv = DataValidation(
        type="list",
        formula1='"Orgeneral,Korgeneral,Tümgeneral,Tuğgeneral,Albay,Yarbay,Binbaşı,'
                 'Yüzbaşı,Üsteğmen,Teğmen,Asteğmen,Kd.Başçavuş,Başçavuş,Üstçavuş,'
                 'Kıdemli Üstçavuş,Çavuş,Uzman Jandarma,Uzman Erbaş,Sözleşmeli Erbaş/Er,'
                 'Erbaş/Er,Memur"',
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Geçersiz Rütbe",
        error="Lütfen listeden geçerli bir rütbe seçin."
    )
    ws.add_data_validation(rutbe_dv)
    rutbe_dv.sqref = "B3:B1000"

    # Muafiyet türü listesi
    muafiyet_dv = DataValidation(
        type="list",
        formula1='"Daimi,Gece,Hamile,Kadın-Ayda1,Yok"',
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Geçersiz Muafiyet",
        error="Lütfen listeden bir muafiyet türü seçin."
    )
    ws.add_data_validation(muafiyet_dv)
    muafiyet_dv.sqref = "E3:E1000"

    # ── Örnek Veri Satırları ──────────────────────────────────────────────────
    ornek_veri = [
        # Sicil, Rütbe,          Ad Soyad,          Kıdem, Muafiyet,        Grup,       T.Puan, HI,  Cu, Cmt, Pzr, Özel, SonNöbet
        ("10001", "Yarbay",      "Ahmet YILMAZ",     22, "Yok",             "Nöb.Amir",   48,   18,   6,   8,   9,   7,  date(2026, 5, 10)),
        ("10002", "Binbaşı",     "Mehmet KAYA",      18, "Yok",             "Nöb.Amir",   42,   16,   5,   7,   8,   6,  date(2026, 5, 14)),
        ("10003", "Yüzbaşı",     "Ali DEMİR",        14, "Yok",             "Nöb.Subay",  36,   14,   4,   6,   7,   5,  date(2026, 5, 12)),
        ("10004", "Teğmen",      "Ayşe ÇELIK",       3,  "Kadın-Ayda1",    "Nöb.Subay",  12,    5,   2,   2,   2,   1,  date(2026, 5, 18)),
        ("10005", "Başçavuş",    "Hasan ŞAHIN",      16, "Yok",             "Nöb.Astsb",  44,   17,   6,   8,   8,   5,  date(2026, 5, 11)),
        ("10006", "Üstçavuş",    "Zeynep ARSLAN",    9,  "Gece",            "Nöb.Astsb",  20,    9,   3,   3,   3,   2,  date(2026, 5, 16)),
        ("10007", "Uzman Jandarma","Fatih KOÇAK",    7,  "Yok",             "Nöb.Astsb",  26,   10,   4,   5,   5,   2,  date(2026, 5, 9)),
        ("10008", "Uzman Erbaş", "Elif ÖZTÜRK",      5,  "Hamile",          "—",           0,    0,   0,   0,   0,   0,  None),
    ]

    for row_idx, satir in enumerate(ornek_veri, start=3):
        row_num = row_idx
        for col_idx, deger in enumerate(satir, start=1):
            cell = ws.cell(row=row_num, column=col_idx)
            if isinstance(deger, date):
                style_data_cell(cell, deger, row_idx, align="center",
                                number_format="DD.MM.YYYY")
            elif isinstance(deger, int):
                style_data_cell(cell, deger, row_idx, align="center")
            else:
                style_data_cell(cell, deger, row_idx,
                                align="center" if col_idx in (1, 2, 4, 5) else "left")
        set_row_height(ws, row_num, 22)

    # ── Boş Veri Satırları (50 adet) ──────────────────────────────────────────
    son_ornek = 3 + len(ornek_veri)
    for row_num in range(son_ornek, son_ornek + 50):
        for col_idx in range(1, 14):
            cell = ws.cell(row=row_num, column=col_idx)
            style_data_cell(cell, row_idx=row_num)
        set_row_height(ws, row_num, 20)

    # ── Puan Ağırlıkları Açıklama Kutusu ──────────────────────────────────────
    aciklama_satir = son_ornek + 52
    ws.merge_cells(f"A{aciklama_satir}:M{aciklama_satir}")
    hucre = ws[f"A{aciklama_satir}"]
    hucre.value = (
        "📌 PUAN AĞIRLIKLARI:  "
        "Hafta İçi=1  |  Cuma=2  |  Pazar=3  |  Cumartesi=4  |  "
        "Özel-1 (23 Nis, 1 May, 19 May, 15 Tem, 30 Ağu, 28 Eki)=5  |  "
        "Özel-2 (1 Oca, 29 Eki, Arife/Son Gün)=6  |  "
        "Özel-3 (31 Ara, R.Bay. 1-2. Gün, K.Bay. 1-2-3. Gün)=7"
    )
    hucre.fill = PatternFill("solid", fgColor="FFF3CD")
    hucre.font = Font(name="Calibri", bold=True, size=9, color="856404")
    hucre.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    hucre.border = thin_border(CLR["accent_gold"])
    set_row_height(ws, aciklama_satir, 32)

    # ── Dondur (Freeze) ────────────────────────────────────────────────────────
    ws.freeze_panes = "A3"

    # ── Yazdırma Alanı / Sayfa Düzeni ─────────────────────────────────────────
    ws.print_title_rows = "1:2"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1


# ═══════════════════════════════════════════════════════════════════════════════
# SAYFA 2 — Aylık_Mazeretler
# ═══════════════════════════════════════════════════════════════════════════════
def build_aylik_mazeretler(wb: Workbook):
    ws = wb.create_sheet("Aylık_Mazeretler")
    ws.sheet_properties.tabColor = CLR["tab_orange"]

    # ── Başlık ────────────────────────────────────────────────────────────────
    ws.merge_cells("A1:G1")
    cell = ws["A1"]
    cell.value = "AYLIK NÖBET MAZERET ÇİZELGESİ — Md.5/(15) ve Md.27 Gereği"
    cell.fill = PatternFill("solid", fgColor="7E3517")
    cell.font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thick_border("7E3517")
    set_row_height(ws, 1, 36)

    # ── Alt Açıklama (Satır 2) ────────────────────────────────────────────────
    ws.merge_cells("A2:G2")
    hucre2 = ws["A2"]
    hucre2.value = (
        "⚠  Her ayın en geç 22'sine kadar bir sonraki aya ait mazeretler bu çizelgeye işlenmelidir.  "
        "Son 3 günde nöbet tutan personele, sonraki ayın ilk 3 günü otomatik kısıt uygulanır."
    )
    hucre2.fill = PatternFill("solid", fgColor="FEF3E2")
    hucre2.font = Font(name="Calibri", italic=True, size=9, color="7E3517")
    hucre2.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    hucre2.border = thin_border(CLR["accent_gold"])
    set_row_height(ws, 2, 30)

    # ── Sütun Başlıkları (Satır 3) ────────────────────────────────────────────
    # A=Sicil  B=AdSoyad  C=BaşlangıçTarihi  D=BitişTarihi
    # E=MazeretTürü  F=MazeretKodu  G=Açıklama
    sutunlar = [
        ("Sicil / ID",         14),
        ("Ad Soyad",           25),
        ("Başlangıç Tarihi",   18),
        ("Bitiş Tarihi",       18),
        ("Mazeret Türü",       22),
        ("Mazeret Kodu",       14),
        ("Açıklama / Not",     35),
    ]

    for col_idx, (baslik, genislik) in enumerate(sutunlar, start=1):
        col_letter = get_column_letter(col_idx)
        cell = ws.cell(row=3, column=col_idx)
        style_header_cell(cell, baslik, bg="7E3517")
        set_col_width(ws, col_letter, genislik)

    set_row_height(ws, 3, 38)

    # ── Veri Doğrulama: Mazeret Türü ──────────────────────────────────────────
    mazeret_tur_dv = DataValidation(
        type="list",
        formula1='"İzin,İstirahat,Hastane Yatış,Hava Değişimi,Kurs,Görevli,'
                 'Tutuklu/Görevden Uzak,Mahkeme,Cenazeye Katılım,Özel Mazeret"',
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Geçersiz Mazeret Türü",
        error="Listeden bir mazeret türü seçin."
    )
    ws.add_data_validation(mazeret_tur_dv)
    mazeret_tur_dv.sqref = "E4:E2000"

    # ── Mazeret Kodları Açıklaması: Md.31'e göre ──────────────────────────────
    mazeret_kod_dv = DataValidation(
        type="list",
        formula1='"1-Görevli,2-İstirahat/Hastane,3-Hava Değişimi,4-İzin,'
                 '5-Kurs,6-Pazar Kısıt,7-Tatil+Yedek Kısıt,8-Diğer"',
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Geçersiz Mazeret Kodu",
        error="Yönerge Md.31'e göre 1-8 arası kod seçin."
    )
    ws.add_data_validation(mazeret_kod_dv)
    mazeret_kod_dv.sqref = "F4:F2000"

    # ── Örnek Mazeret Satırları ────────────────────────────────────────────────
    # Yapı: (Sicil, AdSoyad, BaşlangıçTarihi, BitişTarihi, MazeretTürü, MazeretKodu, Açıklama)
    ornek_mazeretler = [
        ("10001", "Ahmet YILMAZ", date(2026, 6, 5),  date(2026, 6, 12), "İzin",      "4-İzin",              "Yıllık izin 05-12 Haziran"),
        ("10003", "Ali DEMİR",    date(2026, 6, 14), date(2026, 6, 21), "Kurs",      "5-Kurs",              "Yabancı dil kursu 14-21 Haziran"),
        ("10005", "Hasan ŞAHIN",  date(2026, 6, 20), date(2026, 6, 25), "İstirahat", "2-İstirahat/Hastane", "Sağlık raporu 20-25 Haziran"),
        ("10002", "Mehmet KAYA",  date(2026, 6, 27), date(2026, 6, 27), "Mahkeme",   "1-Görevli",           "Tanıklık — Adli yükümlülük"),
    ]

    for row_idx, satir in enumerate(ornek_mazeretler, start=4):
        for col_idx, deger in enumerate(satir, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if isinstance(deger, date):
                # C ve D sütunları tarih (Başlangıç=3, Bitiş=4)
                style_data_cell(cell, deger, row_idx, align="center",
                                number_format="DD.MM.YYYY")
            else:
                style_data_cell(cell, deger, row_idx,
                                align="center" if col_idx in (1, 5, 6) else "left")
        set_row_height(ws, row_idx, 22)

    # ── Boş Satırlar (100 adet) ────────────────────────────────────────────────
    son_ornek = 4 + len(ornek_mazeretler)
    for row_num in range(son_ornek, son_ornek + 100):
        for col_idx in range(1, 8):   # A–G (7 sütun)
            cell = ws.cell(row=row_num, column=col_idx)
            style_data_cell(cell, row_idx=row_num)
        set_row_height(ws, row_num, 20)

    # ── Mazeret Kodu Referans Kutusu ──────────────────────────────────────────
    ref_satir = son_ornek + 102
    ws.merge_cells(f"A{ref_satir}:G{ref_satir}")
    ref_hucre = ws[f"A{ref_satir}"]
    ref_hucre.value = (
        "📋 MAZERET KODU REFERANSI (Md.31):  "
        "1=Görevli  |  2=İstirahat/Hastane  |  3=Hava Değişimi  |  4=İzin  |  "
        "5=Kurs  |  6=Pazar Günü Kısıt  |  7=Tatil+Yedek Kısıt  |  8=Diğer"
    )
    ref_hucre.fill = PatternFill("solid", fgColor="E8F5E9")
    ref_hucre.font = Font(name="Calibri", bold=True, size=9, color="1E8449")
    ref_hucre.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ref_hucre.border = thin_border(CLR["accent_green"])
    set_row_height(ws, ref_satir, 28)

    ws.freeze_panes = "A4"
    ws.print_title_rows = "1:3"
    ws.page_setup.orientation = "landscape"


# ═══════════════════════════════════════════════════════════════════════════════
# SAYFA 3 — İstisnai_Tatiller
# ═══════════════════════════════════════════════════════════════════════════════
def build_istisnai_tatiller(wb: Workbook):
    ws = wb.create_sheet("İstisnai_Tatiller")
    ws.sheet_properties.tabColor = CLR["tab_green"]

    # ── Başlık ────────────────────────────────────────────────────────────────
    ws.merge_cells("A1:E1")
    cell = ws["A1"]
    cell.value = "İSTİSNAİ TATİL GÜNLERİ — Nöbet Puan Kategorileri (Md.27 / Md.31)"
    cell.fill = PatternFill("solid", fgColor="145A32")
    cell.font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thick_border("145A32")
    set_row_height(ws, 1, 36)

    # ── Alt Açıklama (Satır 2) ────────────────────────────────────────────────
    ws.merge_cells("A2:E2")
    hucre2 = ws["A2"]
    hucre2.value = (
        "⚠  Bakanlar Kurulu kararıyla idari izne dönüştürülen günler 'Özel-1' olarak işaretlenir. "
        "Nöbetler yayımlandıktan SONRA ortaya çıkan Özel-1 günlerinde nöbet puan farkı bir sonraki aya aktarılır."
    )
    hucre2.fill = PatternFill("solid", fgColor="D5F5E3")
    hucre2.font = Font(name="Calibri", italic=True, size=9, color="145A32")
    hucre2.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    hucre2.border = thin_border(CLR["accent_green"])
    set_row_height(ws, 2, 30)

    # ── Sütun Başlıkları (Satır 3) ────────────────────────────────────────────
    sutunlar = [
        ("Tatil Tarihi",   16),
        ("Açıklama",       35),
        ("Kategori",       16),
        ("Puan Değeri",    13),
        ("Yasal Dayanak / Not", 40),
    ]

    for col_idx, (baslik, genislik) in enumerate(sutunlar, start=1):
        col_letter = get_column_letter(col_idx)
        cell = ws.cell(row=3, column=col_idx)
        style_header_cell(cell, baslik, bg="145A32")
        set_col_width(ws, col_letter, genislik)

    set_row_height(ws, 3, 38)

    # ── Veri Doğrulama: Kategori ──────────────────────────────────────────────
    kategori_dv = DataValidation(
        type="list",
        formula1='"Özel-1,Özel-2,Özel-3"',
        allow_blank=False,
        showErrorMessage=True,
        errorTitle="Geçersiz Kategori",
        error="Sadece Özel-1, Özel-2 veya Özel-3 girebilirsiniz."
    )
    ws.add_data_validation(kategori_dv)
    kategori_dv.sqref = "C4:C500"

    # ── Sabit Yönerge Tatilleri (Md.31'den türetilen tam liste) ───────────────
    yonerge_tatilleri = [
        # (Tarih,                   Açıklama,                            Kategori, Puan, Not)
        (date(2026,  1,  1),  "Yılbaşı",                               "Özel-2",    6,  "Md.31/(5) — 01 Ocak"),
        (date(2026,  4, 23),  "Ulusal Egemenlik ve Çocuk Bayramı",     "Özel-1",    5,  "Md.31/(5) — 23 Nisan"),
        (date(2026,  5,  1),  "Emek ve Dayanışma Günü",                "Özel-1",    5,  "Md.31/(5) — 01 Mayıs"),
        (date(2026,  5, 19),  "Atatürk'ü Anma, Gençlik ve Spor Bayramı","Özel-1",  5,  "Md.31/(5) — 19 Mayıs"),
        (date(2026,  5, 25),  "İdari İzin",                            "Özel-1",    5,  "Bakanlar Kurulu kararı — örnek veri"),
        (date(2026,  7, 15),  "Demokrasi ve Millî Birlik Günü",        "Özel-1",    5,  "Md.31/(5) — 15 Temmuz"),
        (date(2026,  8, 30),  "Zafer Bayramı",                         "Özel-1",    5,  "Md.31/(5) — 30 Ağustos"),
        (date(2026, 10, 28),  "Cumhuriyet Bayramı Arifesi",            "Özel-2",    6,  "Md.31/(5) — 28 Ekim (Arife)"),
        (date(2026, 10, 29),  "Cumhuriyet Bayramı",                    "Özel-1",    5,  "Md.31/(5) — 29 Ekim"),
        (date(2026, 12, 31),  "Yılbaşı Gecesi",                        "Özel-3",    7,  "Md.31/(5) — 31 Aralık"),
        # ── Ramazan Bayramı (örnek — her yıl değişir) ───────────────────────
        (date(2026,  3, 20),  "Ramazan Bayramı Arifesi",               "Özel-2",    6,  "Md.31/(5) — Arife günü"),
        (date(2026,  3, 21),  "Ramazan Bayramı 1. Gün",                "Özel-3",    7,  "Md.31/(5) — Ramazan 1. Gün"),
        (date(2026,  3, 22),  "Ramazan Bayramı 2. Gün",                "Özel-3",    7,  "Md.31/(5) — Ramazan 2. Gün"),
        (date(2026,  3, 23),  "Ramazan Bayramı Son Günü",              "Özel-2",    6,  "Md.31/(5) — Son gün"),
        # ── Kurban Bayramı (örnek — her yıl değişir) ────────────────────────
        (date(2026,  5, 27),  "Kurban Bayramı Arifesi",                "Özel-2",    6,  "Md.31/(5) — Arife günü"),
        (date(2026,  5, 28),  "Kurban Bayramı 1. Gün",                 "Özel-3",    7,  "Md.31/(5) — Kurban 1. Gün"),
        (date(2026,  5, 29),  "Kurban Bayramı 2. Gün",                 "Özel-3",    7,  "Md.31/(5) — Kurban 2. Gün"),
        (date(2026,  5, 30),  "Kurban Bayramı 3. Gün",                 "Özel-3",    7,  "Md.31/(5) — Kurban 3. Gün"),
        (date(2026,  5, 31),  "Kurban Bayramı Son Günü",               "Özel-2",    6,  "Md.31/(5) — Son gün"),
    ]

    # Kategori renk haritası
    kategori_renkler = {
        "Özel-1": ("FFF3CD", "856404"),   # sarı zemin, koyu sarı yazı
        "Özel-2": ("FFE0B2", "E65100"),   # turuncu zemin, koyu turuncu yazı
        "Özel-3": ("FFCDD2", "B71C1C"),   # kırmızı zemin, koyu kırmızı yazı
    }

    for row_idx, (tarih, aciklama, kategori, puan, not_) in enumerate(yonerge_tatilleri, start=4):
        row_bg, row_font = kategori_renkler.get(kategori, (CLR["row_even"], "000000"))

        veriler = [tarih, aciklama, kategori, puan, not_]
        for col_idx, deger in enumerate(veriler, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if isinstance(deger, date):
                cell.value = deger
                cell.number_format = "DD.MM.YYYY"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 4:  # Puan
                cell.value = deger
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(name="Calibri", bold=True, size=11,
                                 color="FFFFFF")
                puan_bg = {"Özel-1": "856404", "Özel-2": "E65100",
                           "Özel-3": "B71C1C"}.get(kategori, "000000")
                cell.fill = PatternFill("solid", fgColor=puan_bg)
                cell.border = thin_border()
                set_row_height(ws, row_idx, 22)
                continue
            else:
                cell.value = deger
                cell.alignment = Alignment(
                    horizontal="center" if col_idx == 3 else "left",
                    vertical="center"
                )
            cell.fill = PatternFill("solid", fgColor=row_bg)
            cell.font = Font(name="Calibri", size=10, color=row_font,
                             bold=(col_idx == 3))
            cell.border = thin_border()
        set_row_height(ws, row_idx, 22)

    # ── Boş Satırlar (30 adet, ek özel günler için) ───────────────────────────
    son_satir = 4 + len(yonerge_tatilleri)
    for row_num in range(son_satir, son_satir + 30):
        for col_idx in range(1, 6):
            cell = ws.cell(row=row_num, column=col_idx)
            style_data_cell(cell, row_idx=row_num)
        set_row_height(ws, row_num, 20)

    # ── Kategori Renk Açıklaması ──────────────────────────────────────────────
    acik_satir = son_satir + 32
    aciklamalar = [
        ("Özel-1 (Puan=5)", "FFF3CD", "856404",
         "23 Nis, 1 May, 19 May, 15 Tem, 30 Ağu, 28 Eki + İdari izin günleri"),
        ("Özel-2 (Puan=6)", "FFE0B2", "E65100",
         "1 Oca, 29 Eki + Dini bayramların arife ve son günleri"),
        ("Özel-3 (Puan=7)", "FFCDD2", "B71C1C",
         "31 Ara + Ramazan Bayramı 1-2. Gün + Kurban Bayramı 1-2-3. Gün"),
    ]
    for i, (etiket, bg, fg, acik) in enumerate(aciklamalar):
        r = acik_satir + i
        ws.merge_cells(f"A{r}:B{r}")
        h1 = ws[f"A{r}"]
        h1.value = etiket
        h1.fill = PatternFill("solid", fgColor=bg)
        h1.font = Font(name="Calibri", bold=True, size=10, color=fg)
        h1.alignment = Alignment(horizontal="center", vertical="center")
        h1.border = thin_border(fg)

        ws.merge_cells(f"C{r}:E{r}")
        h2 = ws[f"C{r}"]
        h2.value = acik
        h2.fill = PatternFill("solid", fgColor=bg)
        h2.font = Font(name="Calibri", size=9, color=fg)
        h2.alignment = Alignment(horizontal="left", vertical="center")
        h2.border = thin_border(fg)
        set_row_height(ws, r, 22)

    ws.freeze_panes = "A4"
    ws.print_title_rows = "1:3"
    ws.page_setup.orientation = "landscape"


# ═══════════════════════════════════════════════════════════════════════════════
# SAYFA 0 — Anasayfa / README
# ═══════════════════════════════════════════════════════════════════════════════
def build_anasayfa(wb: Workbook):
    ws = wb.active
    ws.title = "📋 Kılavuz"
    ws.sheet_properties.tabColor = "607D8B"

    # Başlık
    ws.merge_cells("B2:I2")
    hucre = ws["B2"]
    hucre.value = "NÖBET DAĞITIM OTOMASYON SİSTEMİ"
    hucre.fill = PatternFill("solid", fgColor=CLR["header_bg"])
    hucre.font = Font(name="Calibri", bold=True, size=18, color="FFFFFF")
    hucre.alignment = Alignment(horizontal="center", vertical="center")
    set_row_height(ws, 2, 50)

    ws.merge_cells("B3:I3")
    alt = ws["B3"]
    alt.value = "JGY 64-4 (B) — Jandarma Genel Komutanlığı Nöbet/Vardiyalı Nöbet Hizmetleri Yönergesi"
    alt.fill = PatternFill("solid", fgColor=CLR["subheader_bg"])
    alt.font = Font(name="Calibri", italic=True, size=11, color="FFFFFF")
    alt.alignment = Alignment(horizontal="center", vertical="center")
    set_row_height(ws, 3, 28)

    # Sayfa açıklamaları
    satirlar = [
        (5,  "📁  SAYFA REHBERİ",             CLR["header_bg"],  "FFFFFF", True, 13, 32),
        (7,  "1.  Personel_Havuzu",            CLR["tab_blue"],   "FFFFFF", True, 11, 26),
        (8,  "    Tüm personeli, rütbelerini, kümülatif nöbet puanlarını ve muafiyet durumlarını içerir.", CLR["row_odd"], "000000", False, 10, 22),
        (9,  "    ➜  Yeni personel eklemek için tablonun altındaki boş satırlara veri girin.", CLR["row_even"], "1A5276", False, 10, 20),
        (11, "2.  Aylık_Mazeretler",           "7E3517",          "FFFFFF", True, 11, 26),
        (12, "    Her ay en geç 22'sine kadar personelin nöbet tutamayacağı günler işlenmelidir.", CLR["row_odd"], "000000", False, 10, 22),
        (13, "    ➜  Mazeret kodu Md.31'deki 1-8 skalasına göre seçilir.", CLR["row_even"], "7E3517", False, 10, 20),
        (15, "3.  İstisnai_Tatiller",          "145A32",          "FFFFFF", True, 11, 26),
        (16, "    Özel-1/2/3 kategorisindeki tatil günleri ve nöbet puan değerleri burada tutulur.", CLR["row_odd"], "000000", False, 10, 22),
        (17, "    ➜  İdari izin kararları çıktıkça bu sayfaya eklenir.", CLR["row_even"], "145A32", False, 10, 20),
        (19, "⚙️  ÖNEMLİ NOTLAR",              CLR["header_bg"],  "FFFFFF", True, 13, 32),
        (21, "    •  48 SAAT KURALI: Nöbeti biten personele 48 saat geçmeden yeni nöbet yazılamaz.", CLR["row_odd"], "B71C1C", True, 10, 22),
        (22, "    •  ARDIŞIK BAYRAM KURALI: Önceki bayramda nöbet tutan personele sonraki bayramda nöbet yazılmaz.", CLR["row_even"], "7E3517", True, 10, 22),
        (23, "    •  KADIN PERSONEL: Tabur ve altında ayda max 1 nöbet; alay ve üstünde evliyse 1, bekarsâ normal.", CLR["row_odd"], "1A5276", True, 10, 22),
        (24, "    •  24 HİZMET YILI: Tamamlayan personel kalıcı muafiyete girer (Subaylar için 23 yıl).", CLR["row_even"], "145A32", True, 10, 22),
        (25, "    •  YENİ PERSONEL PUANI: Grubun min-maks puan ortalamasından başlatılır.", CLR["row_odd"], "000000", False, 10, 22),
    ]

    for satir_no, metin, bg, fg, bold, size, yukseklik in satirlar:
        ws.merge_cells(f"B{satir_no}:I{satir_no}")
        hucre = ws[f"B{satir_no}"]
        hucre.value = metin
        hucre.fill = PatternFill("solid", fgColor=bg)
        hucre.font = Font(name="Calibri", bold=bold, size=size, color=fg)
        hucre.alignment = Alignment(horizontal="left", vertical="center",
                                    wrap_text=True)
        hucre.border = thin_border()
        set_row_height(ws, satir_no, yukseklik)

    for col_letter, width in [("A", 3), ("B", 80), ("J", 3)]:
        ws.column_dimensions[col_letter].width = width

    ws.sheet_view.showGridLines = False


# ═══════════════════════════════════════════════════════════════════════════════
# ANA FONKSİYON
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    cikti_dosya = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "personel_listesi.xlsx"
    )

    wb = Workbook()

    print("📋  Kılavuz sayfası oluşturuluyor...")
    build_anasayfa(wb)

    print("👥  Personel_Havuzu sayfası oluşturuluyor...")
    build_personel_havuzu(wb)

    print("📅  Aylık_Mazeretler sayfası oluşturuluyor...")
    build_aylik_mazeretler(wb)

    print("🗓   İstisnai_Tatiller sayfası oluşturuluyor...")
    build_istisnai_tatiller(wb)

    # İlk sayfayı aktif yap
    wb.active = wb["📋 Kılavuz"]

    wb.save(cikti_dosya)
    print(f"\n✅  Excel şablonu başarıyla oluşturuldu:")
    print(f"    {cikti_dosya}")
    print(f"\n📊  Sayfalar:")
    for sh in wb.sheetnames:
        print(f"    • {sh}")


if __name__ == "__main__":
    main()
