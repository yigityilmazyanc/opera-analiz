#!/usr/bin/env python3
"""
Opera Analiz — EPİAŞ Şeffaflık verilerini Excel'e işler.
"Opera Enerji Veri" sekmesinde ayın başından bugüne (yarının PTF'si açıklandıysa yarına):
  F: PTF (TL/MWh)   G: SMF   H: PTF (USD/MWh)   I: AOF (GİP AOF)
  Santral blokları (UEVM ← gerçek zamanlı üretim, İLK KGÜP ← kgup-v1, SON KGÜP ← kgup):
    ARPACIK HES  → AB / AC / AF     YAVUZ HES → AI / AJ / AM     MİDİLLİ HES → AP / AQ / AT
Ayrıca Tarih (A) ve Saat (B) hücreleri gerçek değerle yazılır (kopuk dış-bağlantı formülleri yerine).

Kullanım:
  python opera_guncelle.py                 # sadece Excel'i güncelle
  python opera_guncelle.py --push          # + kopyasını repoya koy, commit'le, GitHub'a pushla
  python opera_guncelle.py --ay 2026-08    # başka ay (varsayılan: içinde bulunulan ay)
  python opera_guncelle.py --dosya "C:\\...\\baska.xlsx"

Şifre: ~/.epias_pw (bu dosya ve repo şifre İÇERMEZ).
NOT: Excel dosyası açıkken kaydedilemez — Excel'i kapatıp yeniden çalıştır.
"""
import argparse, datetime as dt, shutil, subprocess, sys
from pathlib import Path
import pandas as pd
import openpyxl
from eptr2 import EPTR2

HERE = Path(__file__).resolve().parent
VARSAYILAN_XLSX = Path("/mnt/c/Users/yigit/Downloads/Opera Analiz Temmuz 2026.xlsx")
SEKME = "Opera Enerji Veri"
ILK_SATIR = 3                      # veri 3. satırdan başlar (başlık 2. satır)

ORG_ID = "104782"                  # OPERA ENERJİ A.Ş. (TOPLAYICI) — 40X000000104782K
# ad → (rt-gen pp_id, kgup uevcb_id, [UEVM, İLK KGÜP, SON KGÜP] sütunları)
SANTRALLER = {
    "ARPACIK HES": ("1942", "3194132", ["AB", "AC", "AF"]),
    "YAVUZ HES":   ("1263", "4138",    ["AI", "AJ", "AM"]),
    "MİDİLLİ HES": ("1265", "117864",  ["AP", "AQ", "AT"]),
}

def ts_dict(df, kolon):
    """DataFrame → {(gün, saat): değer} — date kolonu ISO+03:00 gelir."""
    if kolon not in df.columns: return {}
    t = pd.to_datetime(df["date"], utc=True).dt.tz_convert("Etc/GMT-3")
    return {(g.day, g.hour): v for g, v in zip(t, df[kolon]) if pd.notna(v)}

def cek(e, key, bas, bit, **kw):
    """Aralığı çek; uç tarih reddedilirse günü geri çekerek dene.
    (santral bazlı rt-gen yalnız DÜNE kadar verir; mcp yarını 14:00'ten sonra verir)"""
    for geri in range(4):
        son = bit - dt.timedelta(days=geri)
        if son < bas: break
        try:
            df = e.call(key, start_date=bas.isoformat(), end_date=son.isoformat(), **kw)
            if len(df): return df
        except Exception:
            continue
    print(f"   uyarı: {key} verisi alınamadı ({bas} → {bit})")
    return pd.DataFrame(columns=["date"])

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ay", default=dt.date.today().strftime("%Y-%m"))
    ap.add_argument("--dosya", default=str(VARSAYILAN_XLSX))
    ap.add_argument("--push", action="store_true", help="kopyayı repoya koy + git push")
    a = ap.parse_args()

    yil, ay = map(int, a.ay.split("-"))
    bas = dt.date(yil, ay, 1)
    bugun = dt.date.today()
    bit = min(bugun + dt.timedelta(days=1),                       # yarının PTF'si 14:00'ten sonra var
              (bas + dt.timedelta(days=40)).replace(day=1) - dt.timedelta(days=1))
    if bit < bas: sys.exit(f"HATA: {a.ay} henüz başlamadı.")
    xlsx = Path(a.dosya)
    if not xlsx.exists(): sys.exit(f"HATA: dosya yok: {xlsx}")

    pw = (Path.home() / ".epias_pw").read_text(encoding="utf-8").splitlines()[0].strip()
    e = EPTR2(username="yigityilmazyanc@outlook.com", password=pw)
    print(f"Aralık: {bas} → {bit}  |  Dosya: {xlsx.name}")

    print("-> fiyatlar (mcp/smp/wap) + sistem yönü (mcp-smp-imb)")
    mcp = cek(e, "mcp", bas, bit); smp = cek(e, "smp", bas, bit); wap = cek(e, "wap", bas, bit)
    imb = cek(e, "mcp-smp-imb", bas, bit)
    seriler = {"E": ts_dict(imb, "systemStatus"),
               "F": ts_dict(mcp, "price"), "G": ts_dict(smp, "systemMarginalPrice"),
               "H": ts_dict(mcp, "priceUsd"), "I": ts_dict(wap, "wap")}
    for ad, (pp_id, uevcb_id, kolonlar) in SANTRALLER.items():
        print(f"-> {ad} (uevm + ilk/son kgüp)")
        seriler[kolonlar[0]] = ts_dict(cek(e, "rt-gen", bas, bit, pp_id=pp_id), "total")
        seriler[kolonlar[1]] = ts_dict(cek(e, "kgup-v1", bas, bit, org_id=ORG_ID, uevcb_id=uevcb_id), "toplam")
        seriler[kolonlar[2]] = ts_dict(cek(e, "kgup", bas, bit, org_id=ORG_ID, uevcb_id=uevcb_id), "toplam")

    from openpyxl.formula.translate import Translator
    wb = openpyxl.load_workbook(xlsx)
    ws = wb[SEKME]
    gun_sayisi = ((bas + dt.timedelta(days=40)).replace(day=1) - bas).days
    son_satir = ILK_SATIR + gun_sayisi * 24 - 1

    # kopuk dış-bağlantı formüllerini temizle: '[1]...' başka bilgisayardaki kaynak dosyaya
    # bakıyordu, bir daha hesaplanamaz — Excel'de 0/#REF çöpü olarak görünüyordu.
    temiz = 0
    for r in range(ILK_SATIR, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.startswith("=") and "[1]" in v:
                ws.cell(r, c).value = None; temiz += 1
    if temiz: print(f"Temizlenen kopuk formül hücresi: {temiz}")

    # satır iskeleti (İA sabitleri, K=F*J gibi şirket-içi formüller): yeni gün dolarken boş
    # hücreler DÜNÜN AYNI SAATİNDEN kopyalanır (formüller satıra çevrilir) — İA miktarı gibi
    # elle girilen değerler gün gün değişebildiği için en yakın kaynak dün; ilk gün için satır 3.
    yonetilen_c = [openpyxl.utils.column_index_from_string(k) for k in ({"A", "B"} | set(seriler))]
    def iskelet(r):
        kaynak = r - 24 if r - 24 >= ILK_SATIR else ILK_SATIR
        if kaynak == r: return
        for c in range(1, ws.max_column + 1):
            if c in yonetilen_c or ws.cell(r, c).value is not None: continue
            v = ws.cell(kaynak, c).value
            if v is None: continue
            kol = openpyxl.utils.get_column_letter(c)
            ws.cell(r, c).value = (Translator(v, origin=f"{kol}{kaynak}").translate_formula(f"{kol}{r}")
                                   if isinstance(v, str) and v.startswith("=") else v)

    yazilan = {k: 0 for k in seriler}
    for gun in range(1, gun_sayisi + 1):
        for saat in range(24):
            r = ILK_SATIR + (gun - 1) * 24 + saat
            if any((gun, saat) in s for s in seriler.values()):
                ws.cell(r, 1).value = dt.datetime(yil, ay, gun, saat)
                ws.cell(r, 2).value = f"{saat:02d}:00"
                for kolon, s in seriler.items():
                    if (gun, saat) in s:
                        v = s[(gun, saat)]
                        ws[f"{kolon}{r}"].value = v if isinstance(v, str) else float(v)
                        yazilan[kolon] += 1
                iskelet(r)                   # boş kalan iskelet hücrelerini dünden tamamla
            else:                            # verisi olmayan satır KOMPLE boş kalsın
                for c in range(1, ws.max_column + 1):
                    if ws.cell(r, c).value is not None: ws.cell(r, c).value = None
    for r in range(son_satir + 1, ws.max_row + 1):   # ay bloğu sonrası artık satırlar
        for c in range(1, ws.max_column + 1):
            if ws.cell(r, c).value is not None: ws.cell(r, c).value = None
    try:
        wb.save(xlsx)
    except PermissionError:
        sys.exit("HATA: Excel dosyası AÇIK — kapatıp yeniden çalıştırın (kaydedilemedi).")
    ozet = ", ".join(f"{k}:{n}" for k, n in yazilan.items())
    print(f"Yazıldı ({xlsx.name}) — hücre sayıları: {ozet}")

    if a.push:
        kopya = HERE / xlsx.name
        if xlsx.resolve() != kopya.resolve() and \
           (not kopya.exists() or xlsx.stat().st_mtime > kopya.stat().st_mtime):
            shutil.copy2(xlsx, kopya)   # yalnız daha yeniyse kopyala (bayat, tazeyi ezmesin)
        for cmd in (["git", "add", "-A"],
                    ["git", "commit", "-m", f"Veri güncellemesi {dt.datetime.now():%Y-%m-%d %H:%M}"],
                    ["git", "push"]):
            r = subprocess.run(cmd, cwd=HERE, capture_output=True, text=True)
            if r.returncode != 0 and "nothing to commit" not in (r.stdout + r.stderr):
                print(f"git uyarı ({' '.join(cmd)}): {(r.stderr or r.stdout).strip()[:200]}")
        print("GitHub'a pushlandı.")

if __name__ == "__main__":
    main()
