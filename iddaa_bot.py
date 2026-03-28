"""
═══════════════════════════════════════════════════════════
  İddaa Telegram Botu
  - Excel'den veri okur (iddaa_analiz.xlsx)
  - Maç analizi, alt/üst tahmini, kupon önerisi, handikap kuponu
  - Her gün 08:00'de Excel otomatik güncellenir
═══════════════════════════════════════════════════════════
"""
import os, re, logging
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (Application, CommandHandler, MessageHandler,
                           CallbackQueryHandler, ContextTypes, filters)
import openpyxl

# ── Ayarlar ───────────────────────────────────────────────────────────────────
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "BURAYA_TOKEN")
SCRIPT_DIR     = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH     = os.path.join(SCRIPT_DIR, "iddaa_analiz.xlsx")

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")
log = logging.getLogger(__name__)

LIG_ISIMLERI = {
    "PL": "Premier League", "LL": "La Liga", "BL": "Bundesliga",
    "SA": "Serie A", "L1": "Ligue 1", "SL": "Süper Lig",
}
ORANLAR_SAYFALAR    = ["PL - Oranlar", "LL - Oranlar", "BL - Oranlar",
                       "SA - Oranlar", "L1 - Oranlar", "SL - Oranlar"]
ISTATISTIK_SAYFALAR = ["PL - Istatistik", "LL - Istatistik", "BL - Istatistik",
                       "SA - Istatistik", "L1 - Istatistik", "SL - Istatistik"]
PUAN_SAYFALAR       = ["PL - Puan Tablosu", "LL - Puan Tablosu", "BL - Puan Tablosu",
                       "SA - Puan Tablosu", "L1 - Puan Tablosu", "SL - Puan Tablosu"]

# ── Excel Okuyucu ─────────────────────────────────────────────────────────────
def excel_oku():
    if not os.path.exists(EXCEL_PATH):
        return None
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    veri = {"oranlar": [], "istatistik": {}, "puan": [], "guncelleme": ""}

    try:
        ozet = wb["OZET"]
        veri["guncelleme"] = str(ozet.cell(1, 1).value or "")
    except Exception:
        pass

    # Oranlar — 22 sütun (genişletilmiş)
    for sh in ORANLAR_SAYFALAR:
        if sh not in wb.sheetnames:
            continue
        lig = sh.split(" - ")[0]
        ws = wb[sh]
        for row in ws.iter_rows(min_row=3, values_only=True):
            if not row[0] or "Veri" in str(row[0]):
                continue
            try:
                def _f(v):
                    try: return float(v)
                    except (TypeError, ValueError): return 0.0

                veri["oranlar"].append({
                    "lig": lig,
                    "tarih": str(row[0]),
                    "saat":  str(row[1] or ""),
                    "ev":    str(row[2] or ""),
                    "dep":   str(row[3] or ""),
                    "o1":  _f(row[4]),  "ox":  _f(row[5]),  "o2":  _f(row[6]),
                    "p1":  _f(row[7]),  "px":  _f(row[8]),  "p2":  _f(row[9]),
                    "marj": _f(row[10]),
                    "favori": str(row[11] or ""),
                    # Yeni alanlar
                    "over25":   _f(row[12]), "under25":  _f(row[13]),
                    "kg_evet":  _f(row[14]), "kg_hayir": _f(row[15]),
                    "iy1":  _f(row[16]), "iyx":  _f(row[17]), "iy2":  _f(row[18]),
                    "hcp_ev":  _f(row[19]), "hcp_dep": _f(row[20]),
                    "hcp_val": str(row[21] or "-"),
                })
            except Exception:
                pass

    # İstatistik
    for sh in ISTATISTIK_SAYFALAR:
        if sh not in wb.sheetnames:
            continue
        lig = sh.split(" - ")[0]
        ws = wb[sh]
        stat = {}
        for row in ws.iter_rows(min_row=3, max_row=12, values_only=True):
            if row[0] and row[1]:
                stat[str(row[0])] = {"deger": row[1], "pct": row[2]}
        veri["istatistik"][lig] = stat

    # Puan tablosu
    for sh in PUAN_SAYFALAR:
        if sh not in wb.sheetnames:
            continue
        lig = sh.split(" - ")[0]
        ws = wb[sh]
        for row in ws.iter_rows(min_row=3, values_only=True):
            if row[0] and row[1]:
                veri["puan"].append({
                    "lig": lig, "sira": row[0], "takim": str(row[1]),
                    "oyun": row[2], "galibiyet": row[3], "beraberlik": row[4],
                    "maglubiyet": row[5], "puan": row[9], "form": str(row[10] or ""),
                })
    return veri

# ── Maç Ara ───────────────────────────────────────────────────────────────────
def mac_ara(veri, sorgu):
    sorgu_lower = sorgu.lower()
    return [m for m in veri["oranlar"]
            if sorgu_lower in m["ev"].lower() or sorgu_lower in m["dep"].lower()]

# ── Maç Analizi ──────────────────────────────────────────────────────────────
def mac_analiz_metni(m, veri):
    lig  = m["lig"]
    stat = veri["istatistik"].get(lig, {})

    ev_gal_pct = (stat.get("Ev Sahibi Galibiyeti", {}).get("pct") or 0)
    ust_pct    = (stat.get("2.5 Üst (3+ gol)",     {}).get("pct") or 0)
    kg_pct     = (stat.get("Karşılıklı Gol (KG)",  {}).get("pct") or 0)
    ort_gol    = (stat.get("Maç Başı Ort. Gol",     {}).get("deger") or 0)

    # 1X2 öneri
    max_p = max(m["p1"], m["px"], m["p2"])
    if max_p == m["p1"]:   oneri = "1 (Ev Sahibi)"; oneri_oran = m["o1"]
    elif max_p == m["px"]: oneri = "X (Beraberlik)"; oneri_oran = m["ox"]
    else:                  oneri = "2 (Deplasman)";  oneri_oran = m["o2"]

    # Alt/Üst — gerçek oran varsa onu öncelikle kullan
    if m["over25"] and m["under25"]:
        if m["over25"] < m["under25"]:
            au_oneri = f"2.5 ÜST @ `{m['over25']}`"
        else:
            au_oneri = f"2.5 ALT @ `{m['under25']}`"
    else:
        # Gerçek oran yoksa lig istatistiğine bak
        if ust_pct >= 0.55:   au_oneri = "2.5 ÜST (lig bazlı)"
        elif ust_pct <= 0.45: au_oneri = "2.5 ALT (lig bazlı)"
        else:                  au_oneri = "Belirsiz"

    # İlk Yarı önerisi
    if m["iy1"] and m["iyx"] and m["iy2"]:
        min_iy = min(m["iy1"], m["iyx"], m["iy2"])
        if min_iy == m["iy1"]:   iy_oneri = f"İY-1 @ `{m['iy1']}`"
        elif min_iy == m["iyx"]: iy_oneri = f"İY-X @ `{m['iyx']}`"
        else:                     iy_oneri = f"İY-2 @ `{m['iy2']}`"
    else:
        iy_oneri = "-"

    # Güven
    if max_p >= 65:   guven = "🟢 Yüksek"
    elif max_p >= 50: guven = "🟡 Orta"
    else:             guven = "🔴 Düşük"

    lig_tam = LIG_ISIMLERI.get(lig, lig)

    metin = (
        f"⚽ *{m['ev']} vs {m['dep']}*\n"
        f"🏆 {lig_tam}  |  📅 {m['tarih']}  {m['saat']}\n"
        f"{'─'*32}\n"
        f"💹 *1X2 Oranları*\n"
        f"  1️⃣ Ev:  `{m['o1']:.2f}`  (%{m['p1']:.0f})\n"
        f"  ✖️ X:   `{m['ox']:.2f}`  (%{m['px']:.0f})\n"
        f"  2️⃣ Dep: `{m['o2']:.2f}`  (%{m['p2']:.0f})\n"
        f"  📊 Bahisçi Marjı: %{m['marj']:.1f}\n"
    )

    # İlk Yarı
    if m["iy1"] and m["iyx"] and m["iy2"]:
        metin += (
            f"{'─'*32}\n"
            f"⏱️ *İlk Yarı 1X2*\n"
            f"  İY-1: `{m['iy1']:.2f}`  |  İY-X: `{m['iyx']:.2f}`  |  İY-2: `{m['iy2']:.2f}`\n"
        )

    # Handikap
    if m["hcp_ev"] and m["hcp_dep"]:
        metin += (
            f"{'─'*32}\n"
            f"⚖️ *Asian Handicap* ({m['hcp_val']})\n"
            f"  Ev: `{m['hcp_ev']:.2f}`  |  Dep: `{m['hcp_dep']:.2f}`\n"
        )

    # Alt/Üst ve KG
    metin += f"{'─'*32}\n🎯 *Alt/Üst & KG (Gerçek Oranlar)*\n"
    if m["over25"] and m["under25"]:
        metin += f"  Over 2.5: `{m['over25']:.2f}`  |  Under 2.5: `{m['under25']:.2f}`\n"
    else:
        metin += f"  2.5 Üst lig oranı: %{ust_pct*100:.0f}  |  Ort. gol: {ort_gol:.2f}\n"
    if m["kg_evet"] and m["kg_hayir"]:
        metin += f"  KG Evet: `{m['kg_evet']:.2f}`  |  KG Hayır: `{m['kg_hayir']:.2f}`\n"
    else:
        metin += f"  KG lig oranı: %{kg_pct*100:.0f}\n"

    # Lig istatistikleri özeti
    metin += (
        f"{'─'*32}\n"
        f"📈 *{lig_tam} Sezon İstatistikleri*\n"
        f"  🏠 Ev galibiyet oranı: %{ev_gal_pct*100:.0f}\n"
        f"  ⚽ Maç başı ort. gol: {ort_gol:.2f}\n"
        f"  📊 2.5 Üst oranı: %{ust_pct*100:.0f}\n"
        f"  🔄 KG oranı: %{kg_pct*100:.0f}\n"
        f"{'─'*32}\n"
        f"🎯 *Öneri*\n"
        f"  Seçim: *{oneri}* @ `{oneri_oran:.2f}`\n"
        f"  Alt/Üst: *{au_oneri}*\n"
        f"  İlk Yarı: *{iy_oneri}*\n"
        f"  Güven: {guven} (%{max_p:.0f})\n"
    )
    return metin

# ── Kupon Önerisi (1X2) ───────────────────────────────────────────────────────
def kupon_oneri(veri, adet=3):
    adaylar = []
    for m in veri["oranlar"]:
        max_p = max(m["p1"], m["px"], m["p2"])
        if max_p == m["p1"]:   sec = "1"; oran = m["o1"]
        elif max_p == m["px"]: sec = "X"; oran = m["ox"]
        else:                  sec = "2"; oran = m["o2"]
        if max_p >= 55 and oran >= 1.20:
            adaylar.append({**m, "sec": sec, "oran": oran, "ihtimal": max_p})

    adaylar.sort(key=lambda x: (x["ihtimal"] - x["marj"]), reverse=True)
    secilen = adaylar[:adet]

    if not secilen:
        return "⚠️ Yeterli güvenilir maç bulunamadı."

    toplam_oran = round(
        __import__("functools").reduce(lambda a, b: a * b, [s["oran"] for s in secilen]), 2)
    birlesik = round(
        __import__("functools").reduce(lambda a, b: a * b, [s["ihtimal"]/100 for s in secilen]) * 100, 1)

    metin = f"🎫 *{adet} Maçlık Kupon Önerisi*\n{'─'*32}\n"
    for i, s in enumerate(secilen, 1):
        lig_tam = LIG_ISIMLERI.get(s["lig"], s["lig"])
        metin += (f"{i}. *{s['ev']} vs {s['dep']}*\n"
                  f"   📅 {s['tarih']}  |  🏆 {lig_tam}\n"
                  f"   Seçim: *{s['sec']}* @ `{s['oran']:.2f}`  (%{s['ihtimal']:.0f})\n\n")

    metin += (f"{'─'*32}\n"
              f"💹 Toplam Oran: *{toplam_oran}*\n"
              f"🎯 Birleşik İhtimal: *%{birlesik}*\n\n"
              f"💵 100 TL → *{round(100*toplam_oran)} TL*\n"
              f"💵 200 TL → *{round(200*toplam_oran)} TL*\n")
    return metin

# ── Alt/Üst Kupon ─────────────────────────────────────────────────────────────
def altust_kupon(veri, adet=3):
    adaylar = []
    for m in veri["oranlar"]:
        # Gerçek oran varsa kullan
        if m["over25"] and m["under25"] and m["over25"] > 1 and m["under25"] > 1:
            if m["over25"] < m["under25"]:
                adaylar.append({**m, "au": "ÜST", "au_oran": m["over25"],
                                "au_pct": round(1/m["over25"]*100, 1)})
            else:
                adaylar.append({**m, "au": "ALT", "au_oran": m["under25"],
                                "au_pct": round(1/m["under25"]*100, 1)})
        else:
            # Lig istatistiğine dön
            lig  = m["lig"]
            stat = veri["istatistik"].get(lig, {})
            ust_pct = (stat.get("2.5 Üst (3+ gol)", {}).get("pct") or 0) * 100
            ort_gol = stat.get("Maç Başı Ort. Gol", {}).get("deger") or 0
            if ust_pct >= 55:
                adaylar.append({**m, "au": "ÜST", "au_oran": None, "au_pct": ust_pct})
            elif ust_pct <= 42:
                adaylar.append({**m, "au": "ALT", "au_oran": None, "au_pct": 100 - ust_pct})

    adaylar.sort(key=lambda x: x["au_pct"], reverse=True)
    secilen = adaylar[:adet]

    if not secilen:
        return "⚠️ Yeterli alt/üst adayı bulunamadı."

    metin = f"📊 *{adet} Maçlık Alt/Üst Kuponu*\n{'─'*32}\n"
    for i, s in enumerate(secilen, 1):
        emoji = "🔼" if s["au"] == "ÜST" else "🔽"
        lig_tam = LIG_ISIMLERI.get(s["lig"], s["lig"])
        oran_str = f"@ `{s['au_oran']:.2f}`" if s["au_oran"] else "(lig bazlı)"
        metin += (f"{i}. *{s['ev']} vs {s['dep']}*\n"
                  f"   📅 {s['tarih']}  |  🏆 {lig_tam}\n"
                  f"   {emoji} 2.5 *{s['au']}* {oran_str} — %{s['au_pct']:.0f} ihtimal\n\n")

    metin += f"{'─'*32}\n⚠️ Bahis oranlarını bahisçiden kontrol et."
    return metin

# ── Handikap Kuponu ───────────────────────────────────────────────────────────
def handicap_kupon(veri, adet=3):
    adaylar = []
    for m in veri["oranlar"]:
        if not m["hcp_ev"] or not m["hcp_dep"]:
            continue
        # Daha düşük oran → daha muhtemel taraf
        if m["hcp_ev"] < m["hcp_dep"]:
            sec = "Ev Handicap"; oran = m["hcp_ev"]
        else:
            sec = "Dep Handicap"; oran = m["hcp_dep"]
        pct = round(1 / oran * 100, 1) if oran > 1 else 0
        if pct >= 52 and oran >= 1.70:  # Handikap bahislerinde 1.70+ anlamlı
            adaylar.append({**m, "sec": sec, "oran": oran, "pct": pct})

    adaylar.sort(key=lambda x: (x["pct"] - x["marj"]), reverse=True)
    secilen = adaylar[:adet]

    if not secilen:
        return "⚠️ Yeterli handikap adayı bulunamadı."

    toplam_oran = 1.0
    for s in secilen:
        toplam_oran *= s["oran"]
    toplam_oran = round(toplam_oran, 2)

    metin = f"🔀 *{adet} Maçlık Handikap Kuponu*\n{'─'*32}\n"
    for i, s in enumerate(secilen, 1):
        lig_tam = LIG_ISIMLERI.get(s["lig"], s["lig"])
        metin += (f"{i}. *{s['ev']} vs {s['dep']}*\n"
                  f"   📅 {s['tarih']}  |  🏆 {lig_tam}\n"
                  f"   Handicap ({s['hcp_val']}): *{s['sec']}* @ `{s['oran']:.2f}`  (%{s['pct']:.0f})\n\n")

    metin += (f"{'─'*32}\n"
              f"💹 Toplam Oran: *{toplam_oran}*\n\n"
              f"💵 100 TL → *{round(100*toplam_oran)} TL*\n"
              f"💵 200 TL → *{round(200*toplam_oran)} TL*\n"
              f"⚠️ Handikap bahisleri yüksek risk içerir.")
    return metin

# ── Ana Menü ─────────────────────────────────────────────────────────────────
def ana_menu_kb():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("⚽ Maç Analizi",    callback_data="mac_analizi"),
         InlineKeyboardButton("🎫 Kupon Öner",     callback_data="kupon_3")],
        [InlineKeyboardButton("📊 Alt/Üst Kupon",  callback_data="altust"),
         InlineKeyboardButton("🔀 Handikap",       callback_data="handicap")],
        [InlineKeyboardButton("🏆 Puan Tablosu",   callback_data="puan"),
         InlineKeyboardButton("🔄 Veri Güncelle",  callback_data="guncelle")],
    ])

# ── Handler'lar ───────────────────────────────────────────────────────────────
async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    veri = excel_oku()
    gun  = veri["guncelleme"] if veri else "Excel bulunamadı"
    metin = (f"👋 *İddaa Analiz Botuna Hoş Geldin!*\n\n"
             f"📅 Son güncelleme: {gun}\n\n"
             f"Ne yapmak istiyorsun?")
    await update.message.reply_text(metin, reply_markup=ana_menu_kb(), parse_mode="Markdown")

async def button_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    veri = excel_oku()

    if not veri:
        await q.edit_message_text("❌ Excel dosyası bulunamadı. Önce `iddaa_analiz.py` çalıştır.")
        return

    if q.data == "kupon_3":
        metin = kupon_oneri(veri, 3)
        kb = [[InlineKeyboardButton("5 Maçlık Kupon", callback_data="kupon_5"),
               InlineKeyboardButton("🏠 Ana Menü",    callback_data="ana_menu")]]
        await q.edit_message_text(metin, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif q.data == "kupon_5":
        metin = kupon_oneri(veri, 5)
        kb = [[InlineKeyboardButton("3 Maçlık Kupon", callback_data="kupon_3"),
               InlineKeyboardButton("🏠 Ana Menü",    callback_data="ana_menu")]]
        await q.edit_message_text(metin, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif q.data == "altust":
        metin = altust_kupon(veri, 3)
        kb = [[InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")]]
        await q.edit_message_text(metin, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif q.data == "handicap":
        metin = handicap_kupon(veri, 3)
        kb = [[InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")]]
        await q.edit_message_text(metin, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif q.data == "puan":
        metin = "🏆 *Puan Tabloları (İlk 5)*\n\n"
        ligden_takimlar = {}
        for t in veri["puan"]:
            ligden_takimlar.setdefault(t["lig"], []).append(t)
        for lig_kisa, lig_tam in LIG_ISIMLERI.items():
            takimlar = ligden_takimlar.get(lig_kisa, [])
            if not takimlar:
                continue
            metin += f"*{lig_tam}*\n"
            for t in takimlar[:5]:
                metin += f"  {t['sira']}. {t['takim'][:20]:<20} {t['puan']} puan  {t['form']}\n"
            metin += "\n"
        kb = [[InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")]]
        await q.edit_message_text(metin, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif q.data == "mac_analizi":
        metin = ("⚽ *Maç Analizi*\n\n"
                 "Takım adını yaz, analiz edeyim.\n"
                 "Örnek: `Galatasaray`, `Arsenal`, `Bayern`")
        kb = [[InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")]]
        await q.edit_message_text(metin, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif q.data == "guncelle":
        await q.edit_message_text("🔄 Excel güncelleniyor... (birkaç dakika sürebilir)")
        try:
            import subprocess
            script = os.path.join(SCRIPT_DIR, "iddaa_analiz.py")
            subprocess.run(["python3", script], timeout=300)
            kb = [[InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")]]
            await q.edit_message_text("✅ Excel güncellendi!", reply_markup=InlineKeyboardMarkup(kb))
        except Exception as e:
            await q.edit_message_text(f"❌ Güncelleme hatası: {e}")

    elif q.data == "ana_menu":
        gun = veri["guncelleme"]
        metin = f"👋 *Ana Menü*\n📅 Son güncelleme: {gun}\n\nNe yapmak istiyorsun?"
        await q.edit_message_text(metin, reply_markup=ana_menu_kb(), parse_mode="Markdown")

    elif q.data.startswith("mac_"):
        idx = int(q.data.split("_")[1])
        veri2 = excel_oku()
        m = veri2["oranlar"][idx]
        metin = mac_analiz_metni(m, veri2)
        kb = [[InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")]]
        await q.edit_message_text(metin, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

async def mesaj_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    sorgu = update.message.text.strip()
    if len(sorgu) < 2:
        return

    veri = excel_oku()
    if not veri:
        await update.message.reply_text("❌ Excel bulunamadı.")
        return

    bulunanlar = mac_ara(veri, sorgu)

    if not bulunanlar:
        await update.message.reply_text(
            f"🔍 `{sorgu}` için yaklaşan maç bulunamadı.\n\n"
            "Takım adını yaz. Örnek:\n"
            "`Galatasaray`, `Fenerbahce`, `Arsenal`, `Bayern`, `Inter`, `PSG`",
            parse_mode="Markdown"
        )
        return

    if len(bulunanlar) == 1:
        metin = mac_analiz_metni(bulunanlar[0], veri)
        kb = [[InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")]]
        await update.message.reply_text(metin, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")
    else:
        metin = f"🔍 *'{sorgu}' için {len(bulunanlar)} maç bulundu:*\n\n"
        kb_rows = []
        for i, m in enumerate(bulunanlar[:8]):
            idx = veri["oranlar"].index(m)
            lig_tam = LIG_ISIMLERI.get(m["lig"], m["lig"])
            metin += f"{i+1}. {m['ev']} vs {m['dep']} ({m['tarih']}) — {lig_tam}\n"
            kb_rows.append([InlineKeyboardButton(
                f"{m['ev']} vs {m['dep']}", callback_data=f"mac_{idx}")])
        kb_rows.append([InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")])
        await update.message.reply_text(metin, reply_markup=InlineKeyboardMarkup(kb_rows), parse_mode="Markdown")

# ── Otomatik Güncelleme (Her gün 08:00) ──────────────────────────────────────
async def otomatik_guncelle(ctx: ContextTypes.DEFAULT_TYPE):
    log.info("Otomatik Excel güncellemesi başlatılıyor...")
    try:
        import subprocess
        script = os.path.join(SCRIPT_DIR, "iddaa_analiz.py")
        subprocess.run(["python3", script], timeout=300)
        log.info("Excel güncellendi!")
    except Exception as e:
        log.error(f"Otomatik güncelleme hatası: {e}")

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("=" * 50)
    print("  İddaa Telegram Botu Başlatılıyor...")
    print("=" * 50)

    if not os.path.exists(EXCEL_PATH):
        print("⚠️  Excel bulunamadı, otomatik oluşturuluyor...")
        try:
            import subprocess
            script = os.path.join(SCRIPT_DIR, "iddaa_analiz.py")
            subprocess.run(["python", script], timeout=300)
            print("✅ Excel oluşturuldu!")
        except Exception as e:
            print(f"❌ Excel oluşturulamadı: {e}")

    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("menu",  start))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, mesaj_handler))

    app.job_queue.run_daily(otomatik_guncelle,
                            time=datetime.strptime("08:00", "%H:%M").time())

    print("✅ Bot çalışıyor! Telegram'da /start yaz.")
    app.run_polling()

if __name__ == "__main__":
    main()
