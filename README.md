# Opera Analiz — EPİAŞ Otomatik Veri Dolumu

`Opera Analiz <Ay> <Yıl>.xlsx` dosyasının **Opera Enerji Veri** sekmesini EPİAŞ Şeffaflık
Platformu'ndan otomatik doldurur: fiyatlar, sistem yönü, Opera'nın GÖP/İA miktarları ve
üç HES'in üretim + KGÜP verileri. Ayın başından bugüne (+ açıklandıysa yarına) çalışır,
her gün güncellenebilir.

---

## 🔧 Kurulum (Windows) — 3 adım, tek sefer

### 1) Python kur
- https://www.python.org/downloads/ adresinden **Python 3.11+** indir.
- Kurulumda **"Add python.exe to PATH"** kutusunu MUTLAKA işaretle.

### 2) Bu repoyu indir
- Git varsa: `git clone https://github.com/<kullanici>/opera-analiz.git`
- Git yoksa: GitHub'da **Code → Download ZIP** → bir klasöre çıkart (ör. `C:\opera-analiz`).

### 3) `kurulum.bat`'a çift tıkla  ← **gerisini bu yapar**
- Gerekli paketleri kurar (`eptr2 pandas openpyxl pywin32`),
- `ayarlar.json`'u oluşturur ve Not Defteri'nde açar — alanları doldur, kaydet:

```json
{
  "api_kullanici": "EPOSTANIZI-BURAYA-YAZIN",      ← EPİAŞ Şeffaflık e-postası
  "api_sifre":     "SIFRENIZI-BURAYA-YAZIN",       ← EPİAŞ şifresi
  "mail_gonderen": "GONDEREN-GMAIL-ADRESI",        ← e-posta göndermek istersen
  "mail_sifre":    "GMAIL-UYGULAMA-SIFRESI",
  "mail_kime":     "ALICI-ADRES"
}
```

- EPİAŞ hesabı yoksa kayit.epias.com.tr adresinden ücretsiz açılır.
- Excel dosyası **Downloads/İndirilenler** klasöründe `Opera Analiz <Ay> <Yıl>.xlsx`
  adıyla aranır (ör. `Opera Analiz Agustos 2026.xlsx`; Türkçe yazım "Ağustos" da olur).
  **Ay değişince ayar değişmez** — yeni ayın dosyasını Downloads'a koymak yeterli.
- `ayarlar.json` **gitignore'dadır** — şifreler GitHub'a asla çıkmaz. Yine de kimseyle paylaşma.
- İsteğe bağlı gelişmiş alanlar (gerekirse eklenir): `"dosya_klasoru": "D:\\baska\\klasor"`,
  `"dosya_adi_kalibi": "..."`, `"api_sifre_dosyasi": "~/.epias_pw"` (şifreyi dosyadan okur;
  `EPIAS_PW` ortam değişkeni hepsini ezer), `"piyasa_dosya"` (Piyasa Excel'inin tam yolu).

Kurulum bitti — günlük kullanım aşağıda.

### 4) Çalıştır
Repo klasöründeki **`calistir.bat`** dosyasına çift tıkla — ya da `cmd`'de:
```bat
cd C:\opera-analiz
python opera_guncelle.py
```
Excel dosyası **kapalı** olmalı (açıkken kaydedilemez). Diğer seçenekler:
```bat
python opera_guncelle.py --sifirdan          & rem tabloyu komple silip baştan doldur
python opera_guncelle.py --ay 2026-06        & rem geçmiş bir ay (dosyası klasörde olmalı)
python opera_guncelle.py --push              & rem + kopyayı repoya koy, git commit + push (git ister)
```

### 5) Her gün otomatik çalışsın (isteğe bağlı — Görev Zamanlayıcı)
`cmd`'de tek satır (her gün 15:30'da **sessizce** çalıştırır — pencere açılmaz):
```bat
schtasks /Create /SC DAILY /ST 15:30 /TN "OperaAnaliz" /TR "C:\Windows\pyw.exe C:\opera-analiz\opera_guncelle.py"
```
15:30 seçilme sebebi: yarının PTF'si 14:00'te açıklanır, dünün santral üretimi de yayınlanmış olur.
O saatte çalışamadıysa (bilgisayar kapalıydı, Excel açıktı) tekrar denemez —
**`calistir.bat`'a çift tıklamak yeterli**. Görevi silmek için: `schtasks /Delete /TN "OperaAnaliz" /F`

### 6) Raporu e-posta ile gönder — `gonder.bat` (çift tıkla)
**`gonder.bat`** üç adımı arka arkaya yapar:
1. Opera Analiz verilerini günceller,
2. (yan klasörde `piyasa-analiz` varsa) Piyasa Analiz verilerini günceller,
3. **Aylık Rapor + Günlük Rapor** sekmelerini **tek PDF**'e çevirir (Excel kurulu olmalı),
   yanına **Piyasa Analiz `<Yıl>`.xlsx**'i ekler ve tek e-postayla gönderir.

Sadece veri güncellemek istiyorsan `calistir.bat` yeterli — mail atmaz.

Mail için `ayarlar.json`'a üç alan eklenir:
```json
  "mail_gonderen": "gonderen@gmail.com",
  "mail_sifre":    "GMAIL-UYGULAMA-SIFRESI",
  "mail_kime":     "alici@ornek.com"
```
- `mail_sifre` normal Gmail şifresi DEĞİLDİR: myaccount.google.com → **Güvenlik** →
  **2 Adımlı Doğrulama** (açık olmalı) → **Uygulama şifreleri** → üretilen 16 haneli şifre.
- `mail_kime` virgülle birden çok alıcı alabilir.
- Excel dosyaları **açıkken de** PDF üretilebilir; ama veriler güncellenemez —
  `gonder.bat` bu durumda uyarı gösterip mevcut haliyle gönderir.

---

## Doldurulan sütunlar
| Sütun | İçerik | Kaynak |
|---|---|---|
| A, B | Tarih, Saat | script yazar |
| E | Sistem Yönü (Enerji Açığı/Fazlası) | `mcp-smp-imb` |
| F | PTF (TL/MWh) | `mcp` |
| G | SMF (TL/MWh) | `smp` (uzlaşan saatler) |
| H | PTF (USD/MWh) | `mcp.priceUsd` |
| I | AOF (TL/MWh) | `wap` (GİP ağırlıklı ort.) |
| J / L | GÖP SSM / SAM (MWh) | `dam-clearing` org=104782 (matchedOffers / matchedBids) |
| N / Q | İA Satış / Alış Miktarı (MWh) | `bi-short` / `bi-long` org=104782 |
| AB / AC / AF | ARPACIK HES: UEVM / İlk KGÜP / Son KGÜP | `rt-gen pp_id=1942` / `kgup-v1` / `kgup` (uevcb 3194132) |
| AI / AJ / AM | YAVUZ HES: UEVM / İlk KGÜP / Son KGÜP | `rt-gen pp_id=1263` / `kgup-v1` / `kgup` (uevcb 4138) |
| AP / AQ / AT | MİDİLLİ HES: UEVM / İlk KGÜP / Son KGÜP | `rt-gen pp_id=1265` / `kgup-v1` / `kgup` (uevcb 117864) |

Org: OPERA ENERJİ A.Ş. (TOPLAYICI), org_id 104782 (40X000000104782K).
K, M, O, P, R, S gibi tutar sütunları Excel formülü olarak korunur ve bu verilerden hesaplanır.

**"GES Üretim" sekmesi** de doldurulur: Midilli HES'in hibrit güneş üretimi
(`rt-gen pp_id=1265` → `sun` kolonu) E sütununa, PTF TL/USD F/G'ye yazılır;
Tutar (H/I) ve günlük özet (N/O/P) sekmenin kendi formülleriyle hesaplanır.

## Veri zamanlaması (bilinçli boşluklar)
- **PTF**: yarının 24 saati bir gün önce 14:00'te tek seferde açıklanır → hep eksiksiz.
- **SMF + Sistem Yönü**: uzlaştırma ~4-8 saat geriden, bazı saatler (özellikle 00:00) daha geç →
  gün içinde boşluk NORMALDİR, akşam/ertesi sabah tamamlanır.
- **UEVM (santral üretimi)**: EPİAŞ yalnız DÜNE kadar yayınlar → bugünün saatleri yarın dolar.
- **GÖP eşleşme**: hedef günün ihalesi 14:00'te sonuçlanınca gelir.
- Henüz yayınlanmamış hücreye DOKUNULMAZ; veri gelince sonraki çalıştırmada yazılır.
- Verisi hiç olmayan satırlar boş tutulur — tablo her gün kendiliğinden bir gün uzar.
