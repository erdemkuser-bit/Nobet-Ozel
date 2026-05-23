# -*- coding: utf-8 -*-
"""
gui_manager.py — MUH-NET İdari Yönetim Paneli
Personel CRUD, İzin/Mazeret yönetimi ve sistem ayarları.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable, Optional

from audit_logger import audit, SEV_CRITICAL, SEV_WARNING, SEV_INFO
from auth import change_password
from config import config
from data_manager import db

# ─── Renk ve font sabitleri ───────────────────────────────────────────────────
BG        = "#1a2035"
BG_CARD   = "#1e2d42"
BG_ENTRY  = "#253450"
FG        = "#e8eaf6"
FG_DIM    = "#8899aa"
ACCENT    = "#3a7bd5"
ACCENT2   = "#00c6ff"
RED       = "#e74c3c"
GREEN     = "#2ecc71"
FONT_BASE = ("Segoe UI", 10)
FONT_BOLD = ("Segoe UI", 10, "bold")
FONT_HEAD = ("Segoe UI", 11, "bold")
FONT_TINY = ("Segoe UI", 8)

_RUTBE_LIST = [
    "Orgeneral", "Korgeneral", "Tümgeneral", "Tuğgeneral",
    "Albay", "Yarbay", "Binbaşı", "Yüzbaşı", "Üsteğmen", "Teğmen", "Asteğmen",
    "Kd.Başçavuş", "Başçavuş", "Üstçavuş", "Kıdemli Üstçavuş", "Çavuş",
    "Uzman Jandarma", "Uzman Erbaş", "Sözleşmeli Erbaş/Er", "Erbaş/Er",
    "Memur", "Bilinmiyor",
]
_MUAFIYET_LIST = ["Yok", "Daimi", "Gece", "Hamile", "Kadın-Ayda1"]
_MAZERET_TURU  = [
    "İzin", "İstirahat", "Hastane Yatış", "Hava Değişimi",
    "Kurs", "Görevli", "Tutuklu/Görevden Uzak", "Mahkeme",
    "Cenazeye Katılım", "Özel Mazeret",
]
_MAZERET_KODU  = [
    "1-Görevli", "2-İstirahat/Hastane", "3-Hava Değişimi",
    "4-İzin", "5-Kurs", "6-Pazar Kısıt", "7-Tatil+Yedek Kısıt", "8-Diğer",
]


# ─── Yardımcı widget fabrikaları ──────────────────────────────────────────────

def _lbl(parent, text, **kw):
    return tk.Label(
        parent, text=text,
        bg=kw.pop("bg", BG_CARD),
        fg=kw.pop("fg", FG),
        font=kw.pop("font", FONT_BASE),   # font=X geçilirse önce pop et, yoksa FONT_BASE
        **kw,
    )


def _entry(parent, textvariable=None, width=24, state="normal"):
    return ttk.Entry(parent, textvariable=textvariable,
                     width=width, state=state)


def _btn(parent, text, command, fg=FG, bg=ACCENT, **kw):
    return tk.Button(
        parent, text=text, command=command,
        bg=bg, fg=fg, activebackground=ACCENT2, activeforeground="#fff",
        relief="flat", padx=10, pady=4, cursor="hand2",
        font=FONT_BOLD, **kw,
    )


def _combo(parent, values, textvariable=None, width=22):
    c = ttk.Combobox(parent, values=values, textvariable=textvariable,
                     width=width, state="readonly")
    if values:
        c.current(0)
    return c


# ─── Personel Ekleme / Düzenleme Diyalogu ────────────────────────────────────

class PersonelDialog(tk.Toplevel):
    """
    Tek bir personeli eklemek veya düzenlemek için modal form.
    on_save(data_dict) başarılı kayıtta çağrılır.
    """

    def __init__(self, parent: tk.Widget,
                 on_save: Callable[[dict], None],
                 initial: Optional[dict] = None):
        super().__init__(parent)
        self._on_save = on_save
        self._initial = initial or {}
        self._is_edit = bool(initial)

        title = "Personel Düzenle" if self._is_edit else "Yeni Personel Ekle"
        self.title(title)
        self.resizable(False, False)
        self.configure(bg=BG_CARD)
        self.grab_set()

        self._build_ui()
        self._populate()
        self.transient(parent)
        self.wait_window()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 4}
        frm = tk.Frame(self, bg=BG_CARD)
        frm.pack(fill="both", expand=True, padx=16, pady=12)

        def row(lbl, widget, r):
            _lbl(frm, lbl).grid(row=r, column=0, sticky="w", **pad)
            widget.grid(row=r, column=1, sticky="ew", **pad)

        self.v_sicil   = tk.StringVar()
        self.v_rutbe   = tk.StringVar()
        self.v_adsoyad = tk.StringVar()
        self.v_kidem   = tk.StringVar(value="0")
        self.v_muaf    = tk.StringVar(value="Yok")
        self.v_grup    = tk.StringVar(value="Genel")
        self.v_puan    = tk.StringVar(value="0")
        self.v_hi      = tk.StringVar(value="0")
        self.v_cuma    = tk.StringVar(value="0")
        self.v_cmt     = tk.StringVar(value="0")
        self.v_pzr     = tk.StringVar(value="0")
        self.v_ozel    = tk.StringVar(value="0")
        self.v_snt     = tk.StringVar()

        sicil_state = "readonly" if self._is_edit else "normal"
        self._e_sicil = ttk.Entry(frm, textvariable=self.v_sicil,
                                  width=24, state=sicil_state)

        row("Sicil / ID",      self._e_sicil, 0)
        row("Ad Soyad",        _entry(frm, self.v_adsoyad), 1)
        row("Rütbe",           _combo(frm, _RUTBE_LIST, self.v_rutbe), 2)
        row("Kıdem Yılı",      _entry(frm, self.v_kidem, width=8), 3)
        row("Muafiyet Türü",   _combo(frm, _MUAFIYET_LIST, self.v_muaf), 4)
        row("Grup Adı",        _entry(frm, self.v_grup), 5)
        row("Toplam Puan",     _entry(frm, self.v_puan, width=10), 6)
        row("Hafta İçi Nöbet", _entry(frm, self.v_hi,   width=8), 7)
        row("Cuma Nöbet",      _entry(frm, self.v_cuma, width=8), 8)
        row("Cumartesi Nöbet", _entry(frm, self.v_cmt,  width=8), 9)
        row("Pazar Nöbet",     _entry(frm, self.v_pzr,  width=8), 10)
        row("Özel Nöbet",      _entry(frm, self.v_ozel, width=8), 11)
        row("Son Nöbet Tarihi\n(GG.AA.YYYY)",
            _entry(frm, self.v_snt), 12)

        frm.columnconfigure(1, weight=1)

        btn_frm = tk.Frame(self, bg=BG_CARD)
        btn_frm.pack(fill="x", padx=16, pady=(0, 12))
        _btn(btn_frm, "Kaydet", self._save, bg=GREEN).pack(side="right", padx=4)
        _btn(btn_frm, "İptal",  self.destroy, bg="#555").pack(side="right", padx=4)

    def _populate(self):
        d = self._initial
        if not d:
            return
        self.v_sicil.set(d.get("sicil", ""))
        self.v_adsoyad.set(d.get("ad_soyad", ""))
        self.v_rutbe.set(d.get("rutbe", "Bilinmiyor"))
        self.v_kidem.set(str(d.get("kidem_yili", 0)))
        self.v_muaf.set(d.get("muafiyet_turu", "Yok"))
        self.v_grup.set(d.get("grup_adi", "Genel"))
        self.v_puan.set(str(d.get("toplam_puan", 0)))
        self.v_hi.set(str(d.get("hafta_ici", 0)))
        self.v_cuma.set(str(d.get("cuma", 0)))
        self.v_cmt.set(str(d.get("cumartesi", 0)))
        self.v_pzr.set(str(d.get("pazar", 0)))
        self.v_ozel.set(str(d.get("ozel_nobet", 0)))
        self.v_snt.set(d.get("son_nobet_tarihi") or "")

    def _save(self):
        sicil = self.v_sicil.get().strip()
        adsoyad = self.v_adsoyad.get().strip()

        if not sicil:
            messagebox.showwarning("Eksik Alan", "Sicil numarası girilmeli.", parent=self)
            return
        if not adsoyad:
            messagebox.showwarning("Eksik Alan", "Ad Soyad girilmeli.", parent=self)
            return

        try:
            kidem = int(self.v_kidem.get())
            puan  = float(self.v_puan.get())
            hi    = int(self.v_hi.get())
            cuma  = int(self.v_cuma.get())
            cmt   = int(self.v_cmt.get())
            pzr   = int(self.v_pzr.get())
            ozel  = int(self.v_ozel.get())
        except ValueError:
            messagebox.showwarning(
                "Geçersiz Veri",
                "Sayısal alanlara yalnızca rakam girilebilir.",
                parent=self,
            )
            return

        # Son nöbet tarihi format kontrolü
        snt_str = self.v_snt.get().strip()
        if snt_str:
            import re
            if not re.match(r"^\d{2}\.\d{2}\.\d{4}$", snt_str):
                messagebox.showwarning(
                    "Geçersiz Tarih",
                    "Son Nöbet Tarihi GG.AA.YYYY formatında girilmeli.\n"
                    "Örnek: 15.05.2026",
                    parent=self,
                )
                return

        self._on_save({
            "sicil":           sicil,
            "rutbe":           self.v_rutbe.get(),
            "ad_soyad":        adsoyad,
            "kidem_yili":      kidem,
            "muafiyet_turu":   self.v_muaf.get(),
            "grup_adi":        self.v_grup.get().strip() or "Genel",
            "toplam_puan":     puan,
            "hafta_ici":       hi,
            "cuma":            cuma,
            "cumartesi":       cmt,
            "pazar":           pzr,
            "ozel_nobet":      ozel,
            "son_nobet_tarihi": snt_str or None,
        })
        self.destroy()


# ─── Mazeret Ekleme Diyalogu ──────────────────────────────────────────────────

class MazeretDialog(tk.Toplevel):
    """Yeni izin/mazeret kaydı için modal form."""

    def __init__(self, parent: tk.Widget, on_save: Callable[[dict], None]):
        super().__init__(parent)
        self._on_save = on_save
        self.title("İzin / Mazeret Ekle")
        self.resizable(False, False)
        self.configure(bg=BG_CARD)
        self.grab_set()
        self._build_ui()
        self.transient(parent)
        self.wait_window()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}
        frm = tk.Frame(self, bg=BG_CARD)
        frm.pack(fill="both", expand=True, padx=16, pady=12)

        def row(lbl, widget, r):
            _lbl(frm, lbl).grid(row=r, column=0, sticky="w", **pad)
            widget.grid(row=r, column=1, sticky="ew", **pad)

        personel_rows = db.get_all_personel()
        self._sicil_map: dict[str, str] = {
            f"{r['ad_soyad']} ({r['sicil']})": r["sicil"]
            for r in personel_rows
        }
        self._personel_display = list(self._sicil_map.keys())

        self.v_personel  = tk.StringVar()
        self.v_baslangic = tk.StringVar()
        self.v_bitis     = tk.StringVar()
        self.v_tur       = tk.StringVar(value=_MAZERET_TURU[0])
        self.v_kod       = tk.StringVar(value=_MAZERET_KODU[0])
        self.v_aciklama  = tk.StringVar()

        combo_personel = ttk.Combobox(
            frm, values=self._personel_display,
            textvariable=self.v_personel, width=30,
        )
        combo_personel.state(["!readonly"])
        if self._personel_display:
            combo_personel.current(0)

        row("Personel",            combo_personel, 0)
        row("Başlangıç (GG.AA.YYYY)", _entry(frm, self.v_baslangic, width=14), 1)
        row("Bitiş (GG.AA.YYYY)",     _entry(frm, self.v_bitis,     width=14), 2)
        row("Mazeret Türü",         _combo(frm, _MAZERET_TURU, self.v_tur, width=26), 3)
        row("Mazeret Kodu",         _combo(frm, _MAZERET_KODU, self.v_kod, width=26), 4)
        row("Açıklama",             _entry(frm, self.v_aciklama, width=28), 5)

        frm.columnconfigure(1, weight=1)

        btn_frm = tk.Frame(self, bg=BG_CARD)
        btn_frm.pack(fill="x", padx=16, pady=(0, 12))
        _btn(btn_frm, "Kaydet", self._save, bg=GREEN).pack(side="right", padx=4)
        _btn(btn_frm, "İptal",  self.destroy, bg="#555").pack(side="right", padx=4)

        _lbl(frm,
             "⚠ Geçmişe yönelik tarihler de kabul edilir ancak uyarı verilir.",
             fg=FG_DIM, font=FONT_TINY).grid(
                 row=6, column=0, columnspan=2, sticky="w", padx=10, pady=(4, 0))

    def _save(self):
        import re
        from datetime import date as dt_date

        display = self.v_personel.get().strip()
        sicil = self._sicil_map.get(display, display)

        bas_str = self.v_baslangic.get().strip()
        bit_str = self.v_bitis.get().strip()

        date_re = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
        if not date_re.match(bas_str) or not date_re.match(bit_str):
            messagebox.showwarning(
                "Geçersiz Tarih",
                "Lütfen GG.AA.YYYY formatında tarih girin.\nÖrnek: 10.06.2026",
                parent=self,
            )
            return

        try:
            d, m, y = bas_str.split(".")
            bas = dt_date(int(y), int(m), int(d))
            d, m, y = bit_str.split(".")
            bit = dt_date(int(y), int(m), int(d))
        except ValueError:
            messagebox.showwarning("Geçersiz Tarih",
                                   "Tarih değerleri hatalı.", parent=self)
            return

        if bit < bas:
            messagebox.showwarning(
                "Tarih Hatası",
                "Bitiş tarihi başlangıç tarihinden önce olamaz.",
                parent=self,
            )
            return

        # Geçmiş tarih uyarısı (engel değil)
        today = dt_date.today()
        if bit < today:
            if not messagebox.askyesno(
                "Geçmiş Tarih Uyarısı",
                f"Seçilen tarih aralığı ({bas_str} – {bit_str}) geçmişte.\n"
                "Yine de kaydetmek istiyor musunuz?",
                parent=self,
            ):
                return

        self._on_save({
            "sicil":        sicil,
            "ad_soyad":     display.split("(")[0].strip(),
            "baslangic":    bas_str,
            "bitis":        bit_str,
            "mazeret_turu": self.v_tur.get(),
            "mazeret_kodu": self.v_kod.get(),
            "aciklama":     self.v_aciklama.get().strip(),
        })
        self.destroy()


# ─── Personel Sekmesi ─────────────────────────────────────────────────────────

class PersonelTab(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        try:
            self._build()
        except Exception:
            import traceback
            traceback.print_exc()
            tk.Label(self, text="[PersonelTab] Sekme yüklenemedi.\nDetay için terminal çıktısına bakın.",
                     bg=BG, fg="#e74c3c", font=FONT_BASE, justify="left").pack(padx=20, pady=20)
            return
        try:
            self._refresh()
        except Exception:
            import traceback
            traceback.print_exc()

    def _build(self):
        # Araç çubuğu
        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=8, pady=(8, 4))
        _btn(bar, "+ Ekle",    self._ekle,   bg=GREEN).pack(side="left",  padx=3)
        _btn(bar, "✎ Düzenle", self._duzenle).pack(side="left",  padx=3)
        _btn(bar, "✖ Sil",     self._sil,    bg=RED).pack(side="left",  padx=3)
        _btn(bar, "⟳ Yenile",  self._refresh, bg="#555").pack(side="right", padx=3)

        _lbl(bar, "Filtre:", bg=BG).pack(side="left", padx=(12, 2))
        self.v_filter = tk.StringVar()
        self.v_filter.trace_add("write", lambda *_: self._refresh())
        ttk.Entry(bar, textvariable=self.v_filter, width=18).pack(side="left")

        # Treeview
        cols = ("sicil", "rutbe", "ad_soyad", "kidem", "muaf", "grup",
                "puan", "hi", "cuma", "cmt", "pzr", "ozel", "snt")
        headers = ("Sicil", "Rütbe", "Ad Soyad", "Kıdem",
                   "Muafiyet", "Grup", "Puan",
                   "Hİ", "Cu", "Cmt", "Pzr", "Özel", "Son Nöbet")
        widths = (80, 120, 170, 50, 80, 70, 50, 35, 35, 35, 35, 35, 90)

        frm = tk.Frame(self, bg=BG)
        frm.pack(fill="both", expand=True, padx=8, pady=4)

        vsb = ttk.Scrollbar(frm, orient="vertical")
        hsb = ttk.Scrollbar(frm, orient="horizontal")
        self.tree = ttk.Treeview(
            frm, columns=cols, show="headings",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
            selectmode="browse",
        )
        vsb.configure(command=self.tree.yview)
        hsb.configure(command=self.tree.xview)

        for col, hdr, w in zip(cols, headers, widths):
            self.tree.heading(col, text=hdr)
            self.tree.column(col, width=w, minwidth=30, anchor="center")
        self.tree.column("ad_soyad", anchor="w")

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)

        self._status = _lbl(self, "", fg=FG_DIM, font=FONT_TINY, bg=BG)
        self._status.pack(anchor="w", padx=10, pady=2)

    def _refresh(self, *_):
        flt = self.v_filter.get().lower()
        self.tree.delete(*self.tree.get_children())
        rows = db.get_all_personel()
        for r in rows:
            if flt and flt not in r["ad_soyad"].lower() \
                    and flt not in r["sicil"].lower():
                continue
            self.tree.insert("", "end", iid=r["sicil"], values=(
                r["sicil"], r["rutbe"], r["ad_soyad"], r["kidem_yili"],
                r["muafiyet_turu"], r["grup_adi"], r["toplam_puan"],
                r["hafta_ici"], r["cuma"], r["cumartesi"],
                r["pazar"], r["ozel_nobet"],
                r["son_nobet_tarihi"] or "",
            ))
        total = len(db.get_all_personel())
        shown = len(self.tree.get_children())
        self._status.config(
            text=f"Toplam {total} personel  |  Gösterilen: {shown}")

    def _selected_sicil(self) -> Optional[str]:
        sel = self.tree.selection()
        return sel[0] if sel else None

    def _ekle(self):
        def on_save(data):
            err = db.add_personel(data)
            if err:
                messagebox.showerror("Hata", err, parent=self)
            else:
                messagebox.showinfo("Başarılı",
                                    f"'{data['ad_soyad']}' eklendi.", parent=self)
                self._refresh()
        PersonelDialog(self.winfo_toplevel(), on_save)

    def _duzenle(self):
        sicil = self._selected_sicil()
        if not sicil:
            messagebox.showwarning("Seçim Yok",
                                   "Lütfen düzenlenecek personeli seçin.", parent=self)
            return
        row = db.get_personel_by_sicil(sicil)
        if not row:
            return

        def on_save(data):
            err = db.update_personel(sicil, data)
            if err:
                messagebox.showerror("Hata", err, parent=self)
            else:
                messagebox.showinfo("Güncellendi",
                                    f"'{data['ad_soyad']}' güncellendi.", parent=self)
                self._refresh()
        PersonelDialog(self.winfo_toplevel(), on_save, initial=row)

    def _sil(self):
        sicil = self._selected_sicil()
        if not sicil:
            messagebox.showwarning("Seçim Yok",
                                   "Lütfen silinecek personeli seçin.", parent=self)
            return
        row = db.get_personel_by_sicil(sicil)
        if not row:
            return
        if messagebox.askyesno(
            "Silme Onayı",
            f"'{row['ad_soyad']}' adlı personel silinsin mi?\n"
            "(Kayıt pasif hale getirilecek, kalıcı silinmeyecek)",
            parent=self,
        ):
            err = db.delete_personel(sicil)
            if err:
                messagebox.showerror("Hata", err, parent=self)
            else:
                self._refresh()


# ─── İzin/Mazeret Sekmesi ─────────────────────────────────────────────────────

class MazeretTab(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        try:
            self._build()
        except Exception:
            import traceback
            traceback.print_exc()
            tk.Label(self, text="[MazeretTab] Sekme yüklenemedi.\nDetay için terminal çıktısına bakın.",
                     bg=BG, fg="#e74c3c", font=FONT_BASE, justify="left").pack(padx=20, pady=20)
            return
        try:
            self._refresh()
        except Exception:
            import traceback
            traceback.print_exc()

    def _build(self):
        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=8, pady=(8, 4))
        _btn(bar, "+ Ekle", self._ekle,  bg=GREEN).pack(side="left", padx=3)
        _btn(bar, "✖ Sil",  self._sil,   bg=RED).pack(side="left",   padx=3)
        _btn(bar, "⟳ Yenile", self._refresh, bg="#555").pack(side="right", padx=3)

        cols   = ("id", "sicil", "ad_soyad", "baslangic", "bitis",
                  "tur", "kod", "aciklama")
        hdrs   = ("ID", "Sicil", "Ad Soyad", "Başlangıç", "Bitiş",
                  "Tür", "Kod", "Açıklama")
        widths = (40, 70, 160, 90, 90, 110, 120, 180)

        frm = tk.Frame(self, bg=BG)
        frm.pack(fill="both", expand=True, padx=8, pady=4)

        vsb = ttk.Scrollbar(frm, orient="vertical")
        hsb = ttk.Scrollbar(frm, orient="horizontal")
        self.tree = ttk.Treeview(
            frm, columns=cols, show="headings",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
        )
        vsb.configure(command=self.tree.yview)
        hsb.configure(command=self.tree.xview)

        for col, hdr, w in zip(cols, hdrs, widths):
            self.tree.heading(col, text=hdr)
            self.tree.column(col, width=w, minwidth=30, anchor="center")
        self.tree.column("ad_soyad", anchor="w")
        self.tree.column("aciklama", anchor="w")

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)

        self._status = _lbl(self, "", fg=FG_DIM, font=FONT_TINY, bg=BG)
        self._status.pack(anchor="w", padx=10, pady=2)

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        rows = db.get_all_mazeretler()
        for r in rows:
            self.tree.insert("", "end", iid=str(r["id"]), values=(
                r["id"], r["sicil"],
                r.get("personel_adi") or r.get("ad_soyad", ""),
                r["baslangic"], r["bitis"],
                r["mazeret_turu"], r.get("mazeret_kodu", ""),
                r.get("aciklama", ""),
            ))
        self._status.config(text=f"Toplam {len(rows)} mazeret kaydı")

    def _ekle(self):
        def on_save(data):
            err = db.add_mazeret(data)
            if err:
                messagebox.showerror("Hata", err, parent=self)
            else:
                messagebox.showinfo("Başarılı", "Mazeret kaydı eklendi.", parent=self)
                self._refresh()
        MazeretDialog(self.winfo_toplevel(), on_save)

    def _sil(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Seçim Yok",
                                   "Lütfen silinecek kaydı seçin.", parent=self)
            return
        mid = int(sel[0])
        if messagebox.askyesno("Silme Onayı",
                               "Seçili mazeret kaydı silinsin mi?", parent=self):
            db.delete_mazeret(mid)
            self._refresh()


# ─── Sistem Ayarları Sekmesi ──────────────────────────────────────────────────

class AyarlarTab(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        try:
            self._build()
        except Exception:
            import traceback
            traceback.print_exc()
            tk.Label(self, text="[AyarlarTab] Sekme yüklenemedi.\nDetay için terminal çıktısına bakın.",
                     bg=BG, fg="#e74c3c", font=FONT_BASE, justify="left").pack(padx=20, pady=20)

    def _build(self):
        pad = dict(padx=16, pady=8)

        tk.Label(
            self, text="Sistem Ayarları", bg=BG, fg=ACCENT2,
            font=("Segoe UI", 13, "bold"),
        ).pack(anchor="w", **pad)

        # Şifre değiştir
        frm1 = tk.LabelFrame(
            self, text=" Güvenlik ", bg=BG_CARD, fg=FG,
            font=FONT_BOLD, bd=1, relief="groove",
        )
        frm1.pack(fill="x", **pad)
        _lbl(frm1, "Yönetici şifresini değiştir:", bg=BG_CARD).pack(
            anchor="w", padx=10, pady=(8, 4))
        _btn(frm1, "🔑 Şifreyi Değiştir",
             self._change_pw).pack(anchor="w", padx=10, pady=(0, 8))

        # Bütünlük kontrolü
        frm2 = tk.LabelFrame(
            self, text=" Veri Bütünlüğü ", bg=BG_CARD, fg=FG,
            font=FONT_BOLD, bd=1, relief="groove",
        )
        frm2.pack(fill="x", **pad)
        _lbl(frm2,
             "Excel dosyası ile veritabanı arasındaki farkları kontrol eder.",
             bg=BG_CARD).pack(anchor="w", padx=10, pady=(8, 4))
        _btn(frm2, "🔍 Bütünlük Kontrolü Çalıştır",
             self._run_integrity).pack(anchor="w", padx=10, pady=(0, 8))

        # Excel'den içe aktar
        frm3 = tk.LabelFrame(
            self, text=" Veri Aktarımı ", bg=BG_CARD, fg=FG,
            font=FONT_BOLD, bd=1, relief="groove",
        )
        frm3.pack(fill="x", **pad)
        _lbl(frm3,
             "Excel Personel_Havuzu'ndan veritabanına personel içe aktar (yıkıcı olmayan).",
             bg=BG_CARD).pack(anchor="w", padx=10, pady=(8, 4))
        _btn(frm3, "📥 Excel'den İçe Aktar",
             self._sync_excel).pack(anchor="w", padx=10, pady=(0, 8))

        # Durum log alanı
        self.log = tk.Text(
            self, height=8, bg=BG_ENTRY, fg=FG, font=FONT_TINY,
            state="disabled", relief="flat", bd=0,
        )
        self.log.pack(fill="x", padx=16, pady=(4, 12))

    def _log(self, msg: str, tag: str = ""):
        self.log.config(state="normal")
        self.log.tag_config("ok",   foreground=GREEN)
        self.log.tag_config("warn", foreground="#f39c12")
        self.log.tag_config("err",  foreground=RED)
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.config(state="disabled")

    def _change_pw(self):
        change_password(self.winfo_toplevel())

    def _run_integrity(self):
        exc_path = str(config.excel_path)
        if not config.excel_path.exists():
            self._log(f"Excel dosyası bulunamadı: {exc_path}", "err")
            audit.log_critical(
                "Bütünlük Kontrolü Başarısız",
                details=f"Excel dosyası bulunamadı: {exc_path}",
            )
            return
        warns = db.integrity_check(exc_path)
        self._log("── Bütünlük Kontrolü ──")
        if not warns:
            self._log("✔ Sorun bulunamadı. DB ile Excel uyumlu.", "ok")
            audit.log("Bütünlük Kontrolü Çalıştırıldı",
                      details="DB ile Excel arasında fark bulunamadı — temiz.")
        else:
            for w in warns:
                self._log(f"⚠ {w}", "warn")
            audit.log_critical(
                "Bütünlük Kontrolü — Uyumsuzluk Tespit Edildi",
                details=f"{len(warns)} uyarı: " + " | ".join(warns[:5])
                        + (" ..." if len(warns) > 5 else ""),
            )

    def _sync_excel(self):
        if not config.excel_path.exists():
            self._log(f"Excel dosyası bulunamadı: {config.excel_path}", "err")
            return
        added, skipped, errors = db.sync_from_excel(str(config.excel_path))
        self._log("── Excel'den İçe Aktarım ──")
        self._log(f"✔ Eklendi: {added}   Atlandı: {skipped}   Hata: {errors}",
                  "ok" if errors == 0 else "warn")


# ─── Sistem Denetim (Log) Sekmesi ────────────────────────────────────────────

_SEV_COLORS = {
    SEV_INFO:     ("#e8eaf6", "#1e2d42"),   # fg, bg
    SEV_WARNING:  ("#f39c12", "#1e2d42"),
    SEV_CRITICAL: ("#ff4757", "#2d1a1a"),
}


class DenetimTab(tk.Frame):
    """
    Tüm audit_logs kayıtlarını kronolojik olarak gösterir.
    Filtreleme, imza doğrulama ve istatistik özeti içerir.
    """

    def __init__(self, parent):
        super().__init__(parent, bg=BG)
        try:
            self._build()
        except Exception:
            import traceback
            traceback.print_exc()
            tk.Label(self, text="[DenetimTab] Sekme yüklenemedi.\nDetay için terminal çıktısına bakın.",
                     bg=BG, fg="#e74c3c", font=FONT_BASE, justify="left").pack(padx=20, pady=20)
            return
        try:
            self._refresh()
        except Exception:
            import traceback
            traceback.print_exc()

    def _build(self):
        # ── Araç çubuğu ───────────────────────────────────────────────────────
        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=8, pady=(8, 4))

        _lbl(bar, "Filtre:", bg=BG).pack(side="left", padx=(0, 4))

        self.v_sev = tk.StringVar(value="Tümü")
        sev_combo = ttk.Combobox(
            bar,
            values=["Tümü", SEV_INFO, SEV_WARNING, SEV_CRITICAL],
            textvariable=self.v_sev,
            width=14,
            state="readonly",
        )
        sev_combo.pack(side="left", padx=(0, 8))
        sev_combo.bind("<<ComboboxSelected>>", lambda _: self._refresh())

        _lbl(bar, "Ara:", bg=BG).pack(side="left", padx=(0, 4))
        self.v_search = tk.StringVar()
        self.v_search.trace_add("write", lambda *_: self._refresh())
        ttk.Entry(bar, textvariable=self.v_search, width=20).pack(side="left", padx=(0, 8))

        _btn(bar, "⟳ Yenile",        self._refresh,          bg="#555").pack(side="left", padx=2)
        _btn(bar, "🔒 İmza Doğrula", self._verify_integrity, bg=ACCENT).pack(side="left", padx=2)

        # ── İstatistik bandı ──────────────────────────────────────────────────
        self._stat_lbl = _lbl(self, "", fg=FG_DIM, font=FONT_TINY, bg=BG)
        self._stat_lbl.pack(anchor="w", padx=10, pady=(0, 2))

        # ── Treeview ──────────────────────────────────────────────────────────
        cols   = ("id", "timestamp", "severity", "user_action",
                  "target_id", "details", "admin_id", "imza_kisa")
        hdrs   = ("ID", "Zaman Damgası", "Önem", "İşlem",
                  "Hedef", "Detay", "Yönetici", "İmza")
        widths = (40, 130, 90, 170, 80, 280, 90, 80)

        frm = tk.Frame(self, bg=BG)
        frm.pack(fill="both", expand=True, padx=8, pady=4)

        vsb = ttk.Scrollbar(frm, orient="vertical")
        hsb = ttk.Scrollbar(frm, orient="horizontal")
        self.tree = ttk.Treeview(
            frm, columns=cols, show="headings",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
        )
        vsb.configure(command=self.tree.yview)
        hsb.configure(command=self.tree.xview)

        for col, hdr, w in zip(cols, hdrs, widths):
            self.tree.heading(col, text=hdr)
            self.tree.column(col, width=w, minwidth=30, anchor="center")
        self.tree.column("user_action", anchor="w")
        self.tree.column("details",     anchor="w")

        # Renk etiketleri
        self.tree.tag_configure(SEV_INFO,     foreground="#e8eaf6")
        self.tree.tag_configure(SEV_WARNING,  foreground="#f39c12")
        self.tree.tag_configure(SEV_CRITICAL, foreground="#ff4757",
                                background="#3d1515")

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)

    def _refresh(self, *_):
        sev = self.v_sev.get()
        kw  = self.v_search.get().strip() or None

        logs = audit.get_logs(
            limit=500,
            severity_filter=None if sev == "Tümü" else sev,
            keyword=kw,
        )

        self.tree.delete(*self.tree.get_children())
        for r in logs:
            imza_kisa = r["imza"][:12] + "…" if r["imza"] else "?"
            self.tree.insert(
                "", "end",
                values=(
                    r["id"],
                    r["timestamp"],
                    r["severity"],
                    r["user_action"],
                    r["target_id"],
                    r["details"],
                    r["admin_id"],
                    imza_kisa,
                ),
                tags=(r["severity"],),
            )

        # İstatistik güncelle
        stats = audit.get_stats()
        self._stat_lbl.config(
            text=(
                f"Toplam: {stats['toplam']}  |  "
                f"BİLGİ: {stats['bilgi']}  |  "
                f"UYARI: {stats['uyari']}  |  "
                f"KRİTİK: {stats['kritik']}  |  "
                f"Son kayıt: {stats['son_kayit']}  |  "
                f"Gösterilen: {len(logs)}"
            )
        )

    def _verify_integrity(self):
        total, broken, broken_ids = audit.verify_integrity()
        if broken == 0:
            messagebox.showinfo(
                "İmza Doğrulama",
                f"✔ {total} kaydın tamamı doğrulandı.\nHiçbir kayıtta tahrifat tespit edilmedi.",
                parent=self,
            )
            audit.log("Denetim İmza Doğrulaması", details=f"{total} kayıt kontrol edildi — temiz.")
        else:
            ids_str = ", ".join(f"#{i}" for i in broken_ids[:20])
            messagebox.showerror(
                "⚠ Tahrifat Tespit Edildi!",
                f"{total} kayıt içinden {broken} kayıtta imza uyuşmazlığı!\n\n"
                f"Etkilenen ID'ler: {ids_str}"
                + (" ..." if len(broken_ids) > 20 else ""),
                parent=self,
            )
            audit.log_critical(
                "Denetim İmza Uyuşmazlığı",
                details=f"{broken}/{total} kayıt bozuk! ID'ler: {ids_str}",
            )
        self._refresh()


# ─── Admin Panel (Ana Pencere) ────────────────────────────────────────────────

class AdminPanel(tk.Toplevel):
    """
    İdari Yönetim Paneli — Tkinter Toplevel.
    auth.authenticate() ile doğrulama yapıldıktan sonra çağrılmalıdır.
    """

    def __init__(self, parent: tk.Widget):
        super().__init__(parent)
        self.title("MUH-NET — İdari Yönetim Paneli")
        self.geometry("980x620")
        self.minsize(820, 520)
        self.configure(bg=BG)
        self.resizable(True, True)
        self.transient(parent)

        # Style'ı notebook oluşturulmadan ÖNCE ve grab_set'ten ÖNCE kur
        self._configure_style()

        try:
            self._build()
        except Exception:
            import traceback
            traceback.print_exc()
            tk.Label(
                self,
                text="Panel yüklenirken beklenmeyen bir hata oluştu.\n"
                     "Detaylı hata mesajı için terminale bakın.",
                bg=BG, fg="#e74c3c",
                font=("Segoe UI", 11), justify="center",
            ).pack(expand=True)

        # Widget'ların ekrana basılması için idletask flush + render zorla
        self.update_idletasks()
        self.update()

        # grab_set EN SONA — build tamamlandıktan sonra
        self.grab_set()
        self.lift()
        self.focus_force()

    def _configure_style(self) -> None:
        """
        ttk.Style GLOBAL'dır; Toplevel'e bağımlı değildir.
        Notebook pane'i (TNotebook) ve sekme başlıklarını (TNotebook.Tab) ayrı ayrı ayarla.
        """
        style = ttk.Style()
        # Mevcut temayı koru, sadece özel adlı stili ekle
        style.configure("Admin.TNotebook",
                        background=BG,
                        borderwidth=0,
                        tabmargins=[0, 0, 0, 0])
        style.configure("Admin.TNotebook.Tab",
                        background=BG_CARD,
                        foreground=FG,
                        padding=[16, 7],
                        font=FONT_BOLD)
        style.map("Admin.TNotebook.Tab",
                  background=[("selected", ACCENT), ("active", "#2d4a7a")],
                  foreground=[("selected", "#ffffff"), ("active", "#ffffff")])
        # Notebook pane arka planı — çoğu temada TNotebook.Tab'ın altındaki alan
        try:
            style.configure("TNotebook", background=BG)
        except Exception:
            pass

    def _build(self) -> None:
        # ── Başlık bandı ──────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg="#0d1628", height=46)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)

        tk.Label(
            hdr,
            text="⚙  İDARİ YÖNETİM PANELİ",
            bg="#0d1628", fg=ACCENT2,
            font=("Segoe UI", 12, "bold"),
            padx=16,
        ).pack(side="left", fill="y")

        _btn(hdr, "✖  Kapat", self.destroy,
             bg="#c0392b", fg="white").pack(side="right", padx=12, pady=8)

        # ── İçerik alanı (notebook'u sarar) ───────────────────────────────────
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, side="top")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)

        # ── Notebook ──────────────────────────────────────────────────────────
        nb = ttk.Notebook(body, style="Admin.TNotebook")
        nb.grid(row=0, column=0, sticky="nsew")

        # ── Sekmeler — her biri kendi try-except'ine sahip ────────────────────
        def _safe_tab(cls, label):
            try:
                tab = cls(nb)
                nb.add(tab, text=label)
            except Exception:
                import traceback
                traceback.print_exc()
                fallback = tk.Frame(nb, bg=BG)
                tk.Label(
                    fallback,
                    text=f"[{cls.__name__}] Yüklenemedi — terminale bakın.",
                    bg=BG, fg="#e74c3c", font=FONT_BASE,
                ).pack(expand=True, padx=20, pady=20)
                nb.add(fallback, text=label)

        _safe_tab(PersonelTab, "  👤  Personel  ")
        _safe_tab(MazeretTab,  "  📅  İzin / Mazeret  ")
        _safe_tab(AyarlarTab,  "  ⚙  Sistem Ayarları  ")
        _safe_tab(DenetimTab,  "  🔍  Sistem Denetim (Log)  ")

        nb.select(0)   # İlk sekmeyi seçili göster
