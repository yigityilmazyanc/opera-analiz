#!/usr/bin/env python3
"""
Opera Analiz raporlarını e-posta ile gönderir:
  1) "Aylık Rapor" + "Günlük Rapor" sekmeleri Excel üzerinden TEK PDF'e çevrilir
     (bilgisayarda Excel kurulu olmalı; dosya açıkken de çalışır).
  2) Aynı klasördeki "Piyasa Analiz <Yıl>.xlsx" bulunursa olduğu gibi eklenir.
  3) Hepsi tek e-postayla gönderilir.

Mail ayarları ayarlar.json'a yazılır (gitignore'da — GitHub'a çıkmaz):
  "mail_gonderen": "gonderen@gmail.com",
  "mail_sifre":    "GMAIL UYGULAMA ŞİFRESİ",   ← normal Gmail şifresi DEĞİL!
        myaccount.google.com → Güvenlik → 2 Adımlı Doğrulama (açık olmalı) →
        "Uygulama şifreleri" → yeni şifre üret → buraya yapıştır.
  "mail_kime":     "alici@ornek.com"           (virgülle birden çok alıcı yazılabilir)
İsteğe bağlı: "mail_smtp" (vars: smtp.gmail.com), "mail_port" (vars: 465),
              "piyasa_dosya" (Piyasa Excel'inin tam yolu; boşsa otomatik aranır).

Kullanım:  python mail_gonder.py               # PDF üret + gönder
           python mail_gonder.py --sadece-pdf  # göndermeden PDF'i üret (deneme)
"""
import datetime as dt, mimetypes, os, smtplib, sys
from email.message import EmailMessage
from pathlib import Path

from opera_guncelle import HERE, ayarlari_yukle, dosya_bul, yol_cevir

RAPOR_SEKMELER = ["Aylık Rapor", "Günlük Rapor"]
PDF_YOL = HERE / "Opera Rapor.pdf"          # gitignore'da


def pdf_uret(xlsx):
    """İki rapor sekmesini Excel COM ile tek PDF'e bas. Excel dosya açıkken de çalışır
    (ayrı bir Excel örneği salt-okunur açar)."""
    if os.name != "nt":
        sys.exit("HATA: PDF dönüşümü Excel gerektirir — bu adım Windows'ta çalışır (gonder.bat kullanın).")
    import win32com.client
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(xlsx), ReadOnly=True, UpdateLinks=0)
        wb.Worksheets(RAPOR_SEKMELER).Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, str(PDF_YOL))
        wb.Close(False)
    finally:
        excel.Quit()
    print(f"PDF üretildi: {PDF_YOL.name} ({PDF_YOL.stat().st_size // 1024} KB)")


def mail_ayar_dogrula(ayar):
    eksik = [k for k in ("mail_gonderen", "mail_sifre", "mail_kime")
             if not ayar.get(k) or "BURAYA" in str(ayar[k]).upper() or "SIFRESI" in str(ayar[k]).upper()]
    if eksik:
        sys.exit("HATA: ayarlar.json içinde mail ayarları eksik: " + ", ".join(eksik) +
                 "\n  mail_gonderen = gönderen Gmail adresi"
                 "\n  mail_sifre    = Gmail UYGULAMA şifresi (myaccount.google.com → Güvenlik →"
                 "\n                  2 Adımlı Doğrulama → Uygulama şifreleri)"
                 "\n  mail_kime     = alıcı adres(ler), virgülle ayrılır")


def main():
    ayar = ayarlari_yukle()
    bugun = dt.date.today()
    if "--sadece-pdf" not in sys.argv:
        mail_ayar_dogrula(ayar)

    xlsx = dosya_bul(ayar, bugun.year, bugun.month)
    pdf_uret(xlsx)

    # Piyasa Analiz dosyası: ayarla verilmişse o, yoksa Opera dosyasının klasöründe ara
    piyasa = yol_cevir(ayar["piyasa_dosya"]) if ayar.get("piyasa_dosya") \
        else xlsx.parent / f"Piyasa Analiz {bugun.year}.xlsx"
    ekler = [(PDF_YOL, f"Opera Rapor {bugun:%d.%m.%Y}.pdf")]
    if piyasa.exists():
        ekler.append((piyasa, piyasa.name))
    else:
        print(f"uyarı: {piyasa.name} bulunamadı — mail yalnız PDF ile gönderilecek.")

    if "--sadece-pdf" in sys.argv:
        print("--sadece-pdf: mail atlanıyor."); return

    msg = EmailMessage()
    msg["From"] = ayar["mail_gonderen"]
    msg["To"] = ", ".join(a.strip() for a in str(ayar["mail_kime"]).split(","))
    msg["Subject"] = f"Opera Analiz Raporu — {bugun:%d.%m.%Y}"
    msg.set_content("Ekte:\n" + "\n".join(f"- {ad}" for _, ad in ekler) +
                    "\n\nBu e-posta otomatik gönderilmiştir.")
    for yol, ad in ekler:
        tur, _ = mimetypes.guess_type(ad)
        ana, alt = (tur or "application/octet-stream").split("/", 1)
        msg.add_attachment(yol.read_bytes(), maintype=ana, subtype=alt, filename=ad)

    smtp = ayar.get("mail_smtp", "smtp.gmail.com")
    port = int(ayar.get("mail_port", 465))
    with smtplib.SMTP_SSL(smtp, port, timeout=60) as s:
        s.login(ayar["mail_gonderen"], ayar["mail_sifre"])
        s.send_message(msg)
    print(f"Gönderildi → {msg['To']}  ({len(ekler)} ek)")


if __name__ == "__main__":
    main()
