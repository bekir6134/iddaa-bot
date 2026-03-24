"""
═══════════════════════════════════════════════════════════
  İddaa Telegram Botu
  - Excel'den veri okur (iddaa_analiz.xlsx)
  - Maç analizi, alt/üst tahmini, kupon önerisi
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

# ── Excel Okuyucu ─────────────────────────────────────────────────────────────
def excel_oku():
    """Excel'i okuyup tüm veriyi döndür"""
    if not os.path.exists(EXCEL_PATH):
        return None
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    veri = {"oranlar": [], "istatistik": {}, "puan": [], "guncelleme": ""}

    # Güncelleme tarihi
    try:
        ozet = wb["OZET"]
        veri["guncelleme"] = str(ozet.cell(1,1).value or "")
    except: pass

    # Oranlar (tüm ligler)
    for sh in ["PL - Oranlar", "LL - Oranlar", "BL - Oranlar"]:
        if sh not in wb.sheetnames: continue
        lig = sh.split(" - ")[0]
        ws = wb[sh]
        for row in ws.iter_rows(min_row=3, values_only=True):
            if not row[0] or "Veri" in str(row[0]): continue
            try:
                veri["oranlar"].append({
                    "lig": lig, "tarih": str(row[0]), "saat": str(row[1] or ""),
                    "ev": str(row[2] or ""), "dep": str(row[3] or ""),
                    "o1": float(row[4] or 0), "ox": float(row[5] or 0),
                    "o2": float(row[6] or 0),
                    "p1": float(row[7] or 0), "px": float(row[8] or 0),
                    "p2": float(row[9] or 0), "marj": float(row[10] or 0),
                    "favori": str(row[11] or ""),
                })
            except: pass

    # İstatistik (tüm ligler)
    for sh in ["PL - Istatistik", "LL - Istatistik", "BL - Istatistik"]:
        if sh not in wb.sheetnames: continue
        lig = sh.split(" - ")[0]
        ws = wb[sh]
        stat = {}
        for row in ws.iter_rows(min_row=3, max_row=12, values_only=True):
            if row[0] and row[1]:
                stat[str(row[0])] = {"deger": row[1], "pct": row[2]}
        veri["istatistik"][lig] = stat

    # Puan tablosu
    for sh in ["PL - Puan Tablosu", "LL - Puan Tablosu", "BL - Puan Tablosu"]:
        if sh not in wb.sheetnames: continue
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
    """Sorguya göre maç bul (kısmi isim eşleştirme)"""
    sorgu_lower = sorgu.lower()
    bulunanlar = []
    for m in veri["oranlar"]:
        if (sorgu_lower in m["ev"].lower() or
            sorgu_lower in m["dep"].lower()):
            bulunanlar.append(m)
    return bulunanlar

# ── Maç Analizi ──────────────────────────────────────────────────────────────
def mac_analiz_metni(m, veri):
    """Tek maç için detaylı analiz metni üret"""
    lig = m["lig"]
    stat = veri["istatistik"].get(lig, {})

    # Lig bazlı istatistikler
    ev_gal_pct = stat.get("Ev Sahibi Galibiyeti", {}).get("pct", 0) or 0
    ust_pct    = stat.get("2.5 Üst (3+ gol)", {}).get("pct", 0) or 0
    kg_pct     = stat.get("Karşılıklı Gol (KG)", {}).get("pct", 0) or 0
    ort_gol    = stat.get("Maç Başı Ort. Gol", {}).get("deger", 0) or 0

    # Öneri hesapla
    max_p = max(m["p1"], m["px"], m["p2"])
    if max_p == m["p1"]:   oneri = "1 (Ev Sahibi)"; oneri_oran = m["o1"]
    elif max_p == m["px"]: oneri = "X (Beraberlik)"; oneri_oran = m["ox"]
    else:                  oneri = "2 (Deplasman)";  oneri_oran = m["o2"]

    # Alt/Üst tahmini
    if ust_pct >= 0.55:    au_oneri = "2.5 ÜST"
    elif ust_pct <= 0.45:  au_oneri = "2.5 ALT"
    else:                  au_oneri = "Belirsiz"

    # Güven seviyesi
    if max_p >= 65:   guven = "🟢 Yüksek"
    elif max_p >= 50: guven = "🟡 Orta"
    else:             guven = "🔴 Düşük"

    metin = (
        f"⚽ *{m['ev']} vs {m['dep']}*\n"
        f"🏆 {m['lig']}  |  📅 {m['tarih']}  {m['saat']}\n"
        f"{'─'*30}\n"
        f"💹 *Oranlar*\n"
        f"  1️⃣ Ev: `{m['o1']}`  (%{m['p1']:.0f} ihtimal)\n"
        f"  ✖️ X:  `{m['ox']}`  (%{m['px']:.0f} ihtimal)\n"
        f"  2️⃣ Dep: `{m['o2']}`  (%{m['p2']:.0f} ihtimal)\n"
        f"  📊 Bahisçi Marjı: %{m['marj']}\n"
        f"{'─'*30}\n"
        f"📈 *{m['lig']} Sezon İstatistikleri*\n"
        f"  🏠 Ev galibiyet oranı: %{ev_gal_pct*100:.0f}\n"
        f"  ⚽ Maç başı ort. gol: {ort_gol:.2f}\n"
        f"  📊 2.5 Üst oranı: %{ust_pct*100:.0f}\n"
        f"  🔄 KG oranı: %{kg_pct*100:.0f}\n"
        f"{'─'*30}\n"
        f"🎯 *Öneri*\n"
        f"  Seçim: *{oneri}* @ `{oneri_oran}`\n"
        f"  Alt/Üst: *{au_oneri}*\n"
        f"  Güven: {guven} (%{max_p:.0f})\n"
    )
    return metin

# ── Kupon Önerisi ─────────────────────────────────────────────────────────────
def kupon_oneri(veri, adet=3):
    """En güvenli N maçı seç ve kupon oluştur"""
    # İhtimale göre sırala (en yüksek güven)
    adaylar = []
    for m in veri["oranlar"]:
        max_p = max(m["p1"], m["px"], m["p2"])
        if max_p == m["p1"]:   sec="1"; oran=m["o1"]
        elif max_p == m["px"]: sec="X"; oran=m["ox"]
        else:                  sec="2"; oran=m["o2"]
        # Minimum güven %55, minimum oran 1.20
        if max_p >= 55 and oran >= 1.20:
            adaylar.append({**m, "sec": sec, "oran": oran, "ihtimal": max_p})

    # Marj düşük + ihtimal yüksek = en iyi seçim
    adaylar.sort(key=lambda x: (x["ihtimal"] - x["marj"]), reverse=True)
    secilen = adaylar[:adet]

    if not secilen:
        return "⚠️ Yeterli güvenilir maç bulunamadı."

    toplam_oran = 1.0
    for s in secilen: toplam_oran *= s["oran"]
    toplam_oran = round(toplam_oran, 2)

    # Birleşik ihtimal
    birlesik = 1.0
    for s in secilen: birlesik *= s["ihtimal"]/100
    birlesik = round(birlesik*100, 1)

    metin = f"🎫 *{adet} Maçlık Kupon Önerisi*\n{'─'*30}\n"
    for i, s in enumerate(secilen, 1):
        metin += (f"{i}. *{s['ev']} vs {s['dep']}*\n"
                  f"   📅 {s['tarih']}  |  🏆 {s['lig']}\n"
                  f"   Seçim: *{s['sec']}* @ `{s['oran']}`  (%{s['ihtimal']:.0f})\n\n")

    metin += (f"{'─'*30}\n"
              f"💹 Toplam Oran: *{toplam_oran}*\n"
              f"🎯 Birleşik İhtimal: *%{birlesik}*\n\n"
              f"💵 100 TL → *{round(100*toplam_oran)} TL*\n"
              f"💵 200 TL → *{round(200*toplam_oran)} TL*\n")
    return metin

def altust_kupon(veri, adet=3):
    """Alt/Üst kupon önerisi"""
    adaylar = []
    for m in veri["oranlar"]:
        lig  = m["lig"]
        stat = veri["istatistik"].get(lig, {})
        ust_pct  = (stat.get("2.5 Üst (3+ gol)", {}).get("pct", 0) or 0) * 100
        ort_gol  = stat.get("Maç Başı Ort. Gol", {}).get("deger", 0) or 0

        if ust_pct >= 55:
            adaylar.append({**m, "au": "ÜST", "au_pct": ust_pct, "ort_gol": ort_gol})
        elif ust_pct <= 42:
            adaylar.append({**m, "au": "ALT", "au_pct": 100-ust_pct, "ort_gol": ort_gol})

    adaylar.sort(key=lambda x: x["au_pct"], reverse=True)
    secilen = adaylar[:adet]

    if not secilen:
        return "⚠️ Yeterli alt/üst adayı bulunamadı."

    metin = f"📊 *{adet} Maçlık Alt/Üst Kuponu*\n{'─'*30}\n"
    for i, s in enumerate(secilen, 1):
        emoji = "🔼" if s["au"] == "ÜST" else "🔽"
        metin += (f"{i}. *{s['ev']} vs {s['dep']}*\n"
                  f"   📅 {s['tarih']}  |  🏆 {s['lig']}\n"
                  f"   {emoji} 2.5 *{s['au']}* — Lig oranı: %{s['au_pct']:.0f}  Ort.gol: {s['ort_gol']:.2f}\n\n")

    metin += f"{'─'*30}\n⚠️ Oranları bahisçiden kontrol et."
    return metin

# ── Ana Menü ─────────────────────────────────────────────────────────────────
def ana_menu_kb():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("⚽ Maç Analizi",    callback_data="mac_analizi"),
         InlineKeyboardButton("🎫 Kupon Öner",     callback_data="kupon_3")],
        [InlineKeyboardButton("📊 Alt/Üst Kupon",  callback_data="altust"),
         InlineKeyboardButton("🏆 Puan Tablosu",   callback_data="puan")],
        [InlineKeyboardButton("🔄 Veri Güncelle",  callback_data="guncelle")],
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
    q = update.callback_query; await q.answer()
    veri = excel_oku()

    if not veri:
        await q.edit_message_text("❌ Excel dosyası bulunamadı. Önce `iddaa_analiz.py` çalıştır.")
        return

    if q.data == "kupon_3":
        metin = kupon_oneri(veri, 3)
        kb = [[InlineKeyboardButton("5 Maçlık Kupon", callback_data="kupon_5"),
               InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")]]
        await q.edit_message_text(metin, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif q.data == "kupon_5":
        metin = kupon_oneri(veri, 5)
        kb = [[InlineKeyboardButton("3 Maçlık Kupon", callback_data="kupon_3"),
               InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")]]
        await q.edit_message_text(metin, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif q.data == "altust":
        metin = altust_kupon(veri, 3)
        kb = [[InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")]]
        await q.edit_message_text(metin, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif q.data == "puan":
        ligler = {"PL": [], "LL": [], "BL": []}
        for t in veri["puan"]:
            if t["lig"] in ligler:
                ligler[t["lig"]].append(t)
        metin = "🏆 *Puan Tabloları (İlk 5)*\n\n"
        isimler = {"PL": "Premier League", "LL": "La Liga", "BL": "Bundesliga"}
        for lig, takimlar in ligler.items():
            metin += f"*{isimler[lig]}*\n"
            for t in takimlar[:5]:
                metin += f"  {t['sira']}. {t['takim'][:20]:<20} {t['puan']} puan\n"
            metin += "\n"
        kb = [[InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")]]
        await q.edit_message_text(metin, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif q.data == "mac_analizi":
        metin = ("⚽ *Maç Analizi*\n\n"
                 "Takım adını yaz, analiz edeyim.\n"
                 "Örnek: `Arsenal` veya `Bayern`")
        kb = [[InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")]]
        await q.edit_message_text(metin, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

    elif q.data == "guncelle":
        await q.edit_message_text("🔄 Excel güncelleniyor...")
        try:
            import subprocess
            script = os.path.join(SCRIPT_DIR, "iddaa_analiz.py")
            subprocess.run(["python3", script], timeout=120)
            kb = [[InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")]]
            await q.edit_message_text("✅ Excel güncellendi!", reply_markup=InlineKeyboardMarkup(kb))
        except Exception as e:
            await q.edit_message_text(f"❌ Güncelleme hatası: {e}")

    elif q.data == "ana_menu":
        gun = veri["guncelleme"]
        metin = f"👋 *Ana Menü*\n📅 Son güncelleme: {gun}\n\nNe yapmak istiyorsun?"
        await q.edit_message_text(metin, reply_markup=ana_menu_kb(), parse_mode="Markdown")

    elif q.data.startswith("mac_"):
        # Maç detay
        idx = int(q.data.split("_")[1])
        veri2 = excel_oku()
        m = veri2["oranlar"][idx]
        metin = mac_analiz_metni(m, veri2)
        kb = [[InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")]]
        await q.edit_message_text(metin, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")

async def mesaj_handler(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Kullanıcının yazdığı takım adını ara"""
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
            "Takım adını İngilizce yaz. Örnek:\n"
            "`Arsenal`, `Bayern`, `Barcelona`, `Liverpool`",
            parse_mode="Markdown"
        )
        return

    if len(bulunanlar) == 1:
        # Tek maç bulunduysa direkt analiz
        metin = mac_analiz_metni(bulunanlar[0], veri)
        kb = [[InlineKeyboardButton("🏠 Ana Menü", callback_data="ana_menu")]]
        await update.message.reply_text(metin, reply_markup=InlineKeyboardMarkup(kb), parse_mode="Markdown")
    else:
        # Birden fazla maç bulunduysa listele
        metin = f"🔍 *'{sorgu}' için {len(bulunanlar)} maç bulundu:*\n\n"
        kb_rows = []
        for i, m in enumerate(bulunanlar[:8]):
            idx = veri["oranlar"].index(m)
            metin += f"{i+1}. {m['ev']} vs {m['dep']} ({m['tarih']})\n"
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
        subprocess.run(["python3", script], timeout=120)
        log.info("Excel güncellendi!")
    except Exception as e:
        log.error(f"Otomatik güncelleme hatası: {e}")

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("="*50)
    print("  İddaa Telegram Botu Başlatılıyor...")
    print("="*50)

    if not os.path.exists(EXCEL_PATH):
        print("⚠️  Excel bulunamadı, otomatik oluşturuluyor...")
        try:
            import subprocess
            script = os.path.join(SCRIPT_DIR, "iddaa_analiz.py")
            subprocess.run(["python", script], timeout=120)
            print("✅ Excel oluşturuldu!")
        except Exception as e:
            print(f"❌ Excel oluşturulamadı: {e}")

    app = Application.builder().token(TELEGRAM_TOKEN).build()

    # Handler'lar
    app.add_handler(CommandHandler("start",  start))
    app.add_handler(CommandHandler("menu",   start))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, mesaj_handler))

    # Her gün 08:00'de güncelle
    app.job_queue.run_daily(otomatik_guncelle,
                            time=datetime.strptime("08:00", "%H:%M").time())

    print("✅ Bot çalışıyor! Telegram'da /start yaz.")
    app.run_polling()

if __name__ == "__main__":
    main()
