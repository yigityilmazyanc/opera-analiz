# Opera Analiz — EPİAŞ Otomatik Veri Dolumu

`Opera Analiz Temmuz 2026.xlsx` → **Opera Enerji Veri** sekmesini EPİAŞ Şeffaflık
Platformu'ndan (eptr2) doldurur. Ayın başından bugüne, her gün güncellenebilir.

## Doldurulan sütunlar
| Sütun | İçerik | Kaynak |
|---|---|---|
| A, B | Tarih, Saat | (kopuk dış-bağlantı formülleri yerine gerçek değer) |
| F | PTF (TL/MWh) | `mcp` |
| G | SMF (TL/MWh) | `smp` (uzlaşan saatler) |
| H | PTF (USD/MWh) | `mcp.priceUsd` |
| I | AOF (TL/MWh) | `wap` (GİP ağırlıklı ort.) |
| AB / AC / AF | ARPACIK HES: UEVM / İlk KGÜP / Son KGÜP | `rt-gen pp_id=1942` / `kgup-v1` / `kgup` (uevcb 3194132) |
| AI / AJ / AM | YAVUZ HES: UEVM / İlk KGÜP / Son KGÜP | `rt-gen pp_id=1263` / `kgup-v1` / `kgup` (uevcb 4138) |
| AP / AQ / AT | MİDİLLİ HES: UEVM / İlk KGÜP / Son KGÜP | `rt-gen pp_id=1265` / `kgup-v1` / `kgup` (uevcb 117864) |

Org: OPERA ENERJİ A.Ş. (TOPLAYICI), org_id 104782 (40X000000104782K).

## Kullanım
```bash
# Enerji Terminali venv'i kullanılır (eptr2 + openpyxl kurulu):
"/mnt/c/Users/yigit/Desktop/EnerjiPiyasasi/📋 API Ayarları/.venv/bin/python" opera_guncelle.py          # Downloads'taki dosyayı günceller
... opera_guncelle.py --push                                                                            # + repo kopyası + git push
... opera_guncelle.py --ay 2026-08 --dosya "/mnt/c/.../Opera Analiz Ağustos 2026.xlsx"                  # yeni ay
```

## Otomasyon
Enerji Terminali sunucusu (`terminal_server.py`) açıkken **her gün 15:00'ten sonraki ilk
saatlik turda** bu script `--push` ile otomatik çalışır (yarının PTF'si 14:00'te açıklandıktan,
dünün UEVM'i yayınlandıktan sonra). Sonuç sağlık şeridinde görünür.

## Veri zamanlaması (bilinçli boşluklar)
- **UEVM (santral rt-gen)**: yalnız DÜNE kadar yayınlanır → bugünün saatleri ertesi gün dolar.
- **SMF**: uzlaştırma ~birkaç saat geriden gelir → son saatler bir sonraki turda dolar.
- **Yarının PTF/KGÜP'ü**: 14:00 / ~17:00'te açıklanır.
- Henüz yayınlanmamış hücrelere DOKUNULMAZ (eski formül/değer korunur), veri gelince yazılır.

Şifre `~/.epias_pw` içinde — bu repoya asla girmez. Excel dosyası AÇIKKEN kaydedilemez.
