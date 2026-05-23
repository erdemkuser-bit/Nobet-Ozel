from datetime import date
from scheduler import schedule_month_from_data, Personel


def simulator_calistir(personel_listesi: list[Personel], baslangic_yil: int = 2026, ay_sayisi: int = 6):
    print("=== NÖBET SİMÜLASYONU BAŞLADI ===")

    su_anki_yil = baslangic_yil
    su_anki_ay = 1

    for i in range(ay_sayisi):
        print(f"➔ Dönem: {su_anki_yil} - Ay: {su_anki_ay}")

        # Runtime sayaçları her ay başında sıfırla
        for p in personel_listesi:
            p.ay_ici_nobet = 0
            p.yedek = 0
            p.atanan_gunler = []

        try:
            # schedule_month_from_data(yil, ay, personel_list, mazeretler, ozel_tatiller)
            cizelge, guncel_personel, atanmayan = schedule_month_from_data(
                su_anki_yil,
                su_anki_ay,
                personel_listesi,
                mazeretler={},
                ozel_tatiller={},
            )

            # Birikimli sayaçları bir sonraki aya taşı (son_nobet_tarihi zaten güncellendi)
            personel_listesi = guncel_personel

            print(f"  ✔ {su_anki_ay}. Ay başarıyla simüle edildi. "
                  f"({len(cizelge)} atama, {len(atanmayan)} atanmayan gün)")

        except Exception as e:
            print(f"  ❌ Simülasyon bu ayda tıkandı: {e}")
            import traceback
            traceback.print_exc()
            break

        su_anki_ay += 1
        if su_anki_ay > 12:
            su_anki_ay = 1
            su_anki_yil += 1

    print("\n" + "=" * 60)
    print("📊 SİMÜLASYON SONUNDA PERSONEL NÖBET VE PUAN DAĞILIMI")
    print("=" * 60)
    print(f"{'Ad Soyad':<25} | {'Toplam Puan':<12} | {'Hafta İçi':<10} | {'Cumartesi':<10} | {'Pazar':<6}")
    print("-" * 72)
    for p in sorted(personel_listesi, key=lambda x: x.toplam_puan, reverse=True):
        print(f"{p.ad_soyad:<25} | {p.toplam_puan:<12.1f} | {p.hafta_ici:<10} | {p.cumartesi:<10} | {p.pazar:<6}")


# ── Test personel listesi ─────────────────────────────────────────────────────
if __name__ == "__main__":
    test_personelleri = [
        Personel(sicil="2022-6125", rutbe="Uzm.Çvş.", ad_soyad="Erdem KÜSER",        kidem_yili=4,  muafiyet_turu="Yok", grup_adi="A_Grubu", toplam_puan=0,  hafta_ici=0, cuma=0, cumartesi=0, pazar=0, ozel_nobet=0, son_nobet_tarihi=None),
        Personel(sicil="2004-831",  rutbe="Uzm.J.",   ad_soyad="Alinas BAŞIAÇIK",     kidem_yili=22, muafiyet_turu="Yok", grup_adi="A_Grubu", toplam_puan=15, hafta_ici=2, cuma=1, cumartesi=0, pazar=1, ozel_nobet=0, son_nobet_tarihi=None),
        Personel(sicil="2019-9509", rutbe="Uzm.Çvş.", ad_soyad="Fikri BOYRAZ",        kidem_yili=7,  muafiyet_turu="Yok", grup_adi="A_Grubu", toplam_puan=0,  hafta_ici=0, cuma=0, cumartesi=0, pazar=0, ozel_nobet=0, son_nobet_tarihi=None),
        Personel(sicil="2023-3314", rutbe="Uzm.Çvş.", ad_soyad="M. Furkan ÇALIŞKAN",  kidem_yili=3,  muafiyet_turu="Yok", grup_adi="A_Grubu", toplam_puan=5,  hafta_ici=1, cuma=0, cumartesi=0, pazar=0, ozel_nobet=0, son_nobet_tarihi=None),
        Personel(sicil="2010-135",  rutbe="Uzm.Çvş.", ad_soyad="Muhammet ŞİMŞEK",     kidem_yili=16, muafiyet_turu="Yok", grup_adi="A_Grubu", toplam_puan=8,  hafta_ici=2, cuma=0, cumartesi=0, pazar=0, ozel_nobet=0, son_nobet_tarihi=None),
        Personel(sicil="2015-6248", rutbe="Uzm.Çvş.", ad_soyad="Murat TAŞDEMİR",      kidem_yili=11, muafiyet_turu="Yok", grup_adi="A_Grubu", toplam_puan=0,  hafta_ici=0, cuma=0, cumartesi=0, pazar=0, ozel_nobet=0, son_nobet_tarihi=None),
        Personel(sicil="2018-0215", rutbe="Uzm.Çvş.", ad_soyad="Mustafa HUBA",         kidem_yili=8,  muafiyet_turu="Yok", grup_adi="A_Grubu", toplam_puan=0,  hafta_ici=0, cuma=0, cumartesi=0, pazar=0, ozel_nobet=0, son_nobet_tarihi=None),
        Personel(sicil="2020-4412", rutbe="Uzm.Çvş.", ad_soyad="Salih ALPARSLAN",      kidem_yili=6,  muafiyet_turu="Yok", grup_adi="A_Grubu", toplam_puan=0,  hafta_ici=0, cuma=0, cumartesi=0, pazar=0, ozel_nobet=0, son_nobet_tarihi=None),
        Personel(sicil="2021-1105", rutbe="Uzm.Çvş.", ad_soyad="Ümit BAYAZIT",         kidem_yili=5,  muafiyet_turu="Yok", grup_adi="A_Grubu", toplam_puan=0,  hafta_ici=0, cuma=0, cumartesi=0, pazar=0, ozel_nobet=0, son_nobet_tarihi=None),
    ]

    simulator_calistir(test_personelleri, baslangic_yil=2026, ay_sayisi=6)
