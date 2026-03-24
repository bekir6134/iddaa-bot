import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

FD_KEY   = "0fb7930eaa3744f4b6175315a4b734ce"
ODDS_KEY = "ae7313d30cccf1080f66e93c57d9da7a"

FD_BASE   = "https://api.football-data.org/v4"
ODDS_BASE = "https://api.the-odds-api.com/v4"

LIGLER = {
    "Premier League": {"fd_code": "PL",   "odds_key": "soccer_epl"},
    "La Liga":        {"fd_code": "PD",   "odds_key": "soccer_spain_la_liga"},
    "Bundesliga":     {"fd_code": "BL1",  "odds_key": "soccer_germany_bundesliga"},
    "Super Lig":      {"fd_code": "TSL",  "odds_key": "soccer_turkey_super_league"},
}

RENKLER = {
    "baslik_bg":   "1F4E79",  # koyu mavi
    "baslik_yazi": "FFFFFF",  # beyaz
    "alt_baslik":  "2E75B6",  # orta mavi
    "satir1":      "DEEAF1",  # açık mavi
    "satir2":      "FFFFFF",  # beyaz
    "kazandi":     "C6EFCE",  # yeşil
    "kaybetti":    "FFC7CE",  # kırmızı
    "beraberlik":  "FFEB9C",  # sarı
    "pozitif":     "375623",  # koyu yeşil yazı
    "negatif":     "9C0006",  # koyu kırmızı yazı
}

def fd_get(endpoint, params=None):
    r = requests.get(f"{FD_BASE}/{endpoint}",
                     headers={"X-Auth-Token": FD_KEY},
                     params=params, timeout=15)
    if r.status_code == 200:
        return r.json()
    print(f"  FD API hata {r.status_code}: {endpoint}")
    return None

def odds_get(sport_key):
    r = requests.get(f"{ODDS_BASE}/sports/{sport_key}/odds",
                     params={"apiKey": ODDS_KEY, "regions": "eu",
                             "markets": "h2h", "oddsFormat": "decimal"},
                     timeout=15)
    if r.status_code == 200:
        return r.json()
    print(f"  Odds API hata {r.status_code}: {sport_key}")
    return []

# ── Stil yardımcıları ─────────────────────────────────────────────────────────
def stil_baslik(ws, row, col, deger, bg=None, yazi=None, bold=True, boyut=11, hizala="center"):
    c = ws.cell(row=row, column=col, value=deger)
    c.font = Font(name="Arial", bold=bold, color=yazi or RENKLER["baslik_yazi"], size=boyut)
    if bg:
        c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal=hizala, vertical="center", wrap_text=True)
    return c

def stil_veri(ws, row, col, deger, bg=None, yazi="000000", bold=False, hizala="center", sayi_fmt=None):
    c = ws.cell(row=row, column=col, value=deger)
    c.font = Font(name="Arial", bold=bold, color=yazi, size=10)
    if bg:
        c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal=hizala, vertical="center")
    if sayi_fmt:
        c.number_format = sayi_fmt
    return c

def ince_kenar(ws, min_row, max_row, min_col, max_col):
    ince = Side(style="thin", color="BFBFBF")
    for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                             min_col=min_col, max_col=max_col):
        for c in row:
            c.border = Border(left=ince, right=ince, top=ince, bottom=ince)

def col_gen(ws, sutunlar):
    for i, en in enumerate(sutunlar, 1):
        ws.column_dimensions[get_column_letter(i)].width = en

# ── Sayfa: Puan Tablosu ───────────────────────────────────────────────────────
def yaz_puan_tablosu(ws, lig_adi, fd_code):
    print(f"  Puan tablosu çekiliyor: {lig_adi}")
    data = fd_get(f"competitions/{fd_code}/standings")
    if not data:
        ws["A1"] = "Veri alınamadı"
        return

    ws.freeze_panes = "A3"
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 20

    # Başlık
    ws.merge_cells("A1:M1")
    stil_baslik(ws, 1, 1, f"{lig_adi} — Puan Tablosu ({datetime.now().strftime('%d.%m.%Y')})",
                bg=RENKLER["baslik_bg"], boyut=13)

    # Sütun başlıkları
    basliklar = ["Sıra", "Takım", "O", "G", "B", "M", "AG", "YG", "A", "Puan",
                 "Form", "Ev G%", "Dep G%"]
    for i, b in enumerate(basliklar, 1):
        stil_baslik(ws, 2, i, b, bg=RENKLER["alt_baslik"], boyut=10)

    col_gen(ws, [6, 22, 5, 5, 5, 5, 6, 6, 6, 7, 12, 8, 8])

    standings = data.get("standings", [{}])[0].get("table", [])
    for idx, t in enumerate(standings):
        row = idx + 3
        bg = RENKLER["satir1"] if idx % 2 == 0 else RENKLER["satir2"]
        pos = t.get("position", "")

        # Sıra rengi (şampiyon, Avrupa, küme düşme)
        if pos <= 1:   bg_pos = "FFD700"  # altın
        elif pos <= 4: bg_pos = "C6EFCE"  # yeşil
        elif pos >= len(standings) - 2: bg_pos = "FFC7CE"  # kırmızı
        else:          bg_pos = bg

        stil_veri(ws, row, 1,  pos,                              bg=bg_pos, bold=True)
        stil_veri(ws, row, 2,  t.get("team",{}).get("name",""),  bg=bg,     hizala="left", bold=True)
        stil_veri(ws, row, 3,  t.get("playedGames",""),          bg=bg)
        stil_veri(ws, row, 4,  t.get("won",""),                  bg=bg)
        stil_veri(ws, row, 5,  t.get("draw",""),                 bg=bg)
        stil_veri(ws, row, 6,  t.get("lost",""),                 bg=bg)
        stil_veri(ws, row, 7,  t.get("goalsFor",""),             bg=bg)
        stil_veri(ws, row, 8,  t.get("goalsAgainst",""),         bg=bg)
        stil_veri(ws, row, 9,  t.get("goalDifference",""),       bg=bg)
        stil_veri(ws, row, 10, t.get("points",""),               bg=bg, bold=True)
        stil_veri(ws, row, 11, t.get("form","") or "-",          bg=bg)

        oyun = t.get("playedGames", 1) or 1
        kazanma = t.get("won", 0)
        ws.cell(row=row, column=12, value=f"=D{row}/C{row}").number_format = "0%"
        ws.cell(row=row, column=13, value=f"=D{row}/C{row}").number_format = "0%"

    ince_kenar(ws, 2, len(standings)+2, 1, 13)

# ── Sayfa: Maç Sonuçları ──────────────────────────────────────────────────────
def yaz_mac_sonuclari(ws, lig_adi, fd_code):
    print(f"  Maç sonuçları çekiliyor: {lig_adi}")
    data = fd_get(f"competitions/{fd_code}/matches",
                  params={"status": "FINISHED", "limit": 50})
    if not data:
        ws["A1"] = "Veri alınamadı"; return

    ws.freeze_panes = "A3"
    ws.row_dimensions[1].height = 30; ws.row_dimensions[2].height = 20

    ws.merge_cells("A1:J1")
    stil_baslik(ws, 1, 1, f"{lig_adi} — Son 50 Maç Sonucu",
                bg=RENKLER["baslik_bg"], boyut=13)

    basliklar = ["Tarih", "Hafta", "Ev Sahibi", "Skor", "Deplasman",
                 "Sonuç", "İY Skor", "Toplam Gol", "KG", "Durum"]
    for i, b in enumerate(basliklar, 1):
        stil_baslik(ws, 2, i, b, bg=RENKLER["alt_baslik"], boyut=10)

    col_gen(ws, [12, 7, 22, 8, 22, 8, 9, 11, 6, 10])

    maclar = data.get("matches", [])
    for idx, m in enumerate(maclar):
        row = idx + 3
        bg = RENKLER["satir1"] if idx % 2 == 0 else RENKLER["satir2"]

        tarih = m.get("utcDate","")[:10]
        hafta = m.get("matchday","")
        ev    = m.get("homeTeam",{}).get("name","")
        dep   = m.get("awayTeam",{}).get("name","")
        full  = m.get("score",{}).get("fullTime",{})
        ht    = m.get("score",{}).get("halfTime",{})
        ev_g  = full.get("home"); dep_g = full.get("away")
        iy_ev = ht.get("home");   iy_dep= ht.get("away")

        if ev_g is not None and dep_g is not None:
            skor = f"{ev_g} - {dep_g}"
            iy   = f"{iy_ev} - {iy_dep}" if iy_ev is not None else "-"
            toplam = ev_g + dep_g
            kg   = "Var" if ev_g > 0 and dep_g > 0 else "Yok"
            if ev_g > dep_g:   sonuc="1"; bg_s=RENKLER["kazandi"]
            elif ev_g == dep_g: sonuc="X"; bg_s=RENKLER["beraberlik"]
            else:               sonuc="2"; bg_s=RENKLER["kaybetti"]
        else:
            skor=iy="-"; toplam=""; kg=""; sonuc="-"; bg_s=bg

        stil_veri(ws, row, 1,  tarih,  bg=bg, hizala="center")
        stil_veri(ws, row, 2,  hafta,  bg=bg)
        stil_veri(ws, row, 3,  ev,     bg=bg, hizala="left")
        stil_veri(ws, row, 4,  skor,   bg=bg, bold=True)
        stil_veri(ws, row, 5,  dep,    bg=bg, hizala="left")
        stil_veri(ws, row, 6,  sonuc,  bg=bg_s, bold=True)
        stil_veri(ws, row, 7,  iy,     bg=bg)
        stil_veri(ws, row, 8,  toplam, bg=bg)
        stil_veri(ws, row, 9,  kg,     bg=bg)
        stil_veri(ws, row, 10, m.get("status",""), bg=bg)

    ince_kenar(ws, 2, len(maclar)+2, 1, 10)

# ── Sayfa: Bahis Oranları ─────────────────────────────────────────────────────
def yaz_oranlar(ws, lig_adi, odds_key):
    print(f"  Bahis oranları çekiliyor: {lig_adi}")
    maclar = odds_get(odds_key)

    ws.freeze_panes = "A3"
    ws.row_dimensions[1].height = 30; ws.row_dimensions[2].height = 20

    ws.merge_cells("A1:L1")
    stil_baslik(ws, 1, 1, f"{lig_adi} — Yaklaşan Maç Oranları",
                bg=RENKLER["baslik_bg"], boyut=13)

    basliklar = ["Tarih", "Saat", "Ev Sahibi", "Deplasman",
                 "1 Oran", "X Oran", "2 Oran",
                 "1 İhtimal%", "X İhtimal%", "2 İhtimal%",
                 "Marj%", "Favori"]
    for i, b in enumerate(basliklar, 1):
        stil_baslik(ws, 2, i, b, bg=RENKLER["alt_baslik"], boyut=10)

    col_gen(ws, [11, 7, 22, 22, 8, 8, 8, 10, 10, 10, 8, 12])

    if not maclar:
        ws.cell(row=3, column=1, value="Veri bulunamadı (yaklaşan maç yok veya API limiti)")
        return

    for idx, m in enumerate(maclar):
        row = idx + 3
        bg = RENKLER["satir1"] if idx % 2 == 0 else RENKLER["satir2"]

        dt = m.get("commence_time","")
        tarih = dt[:10] if dt else ""
        saat  = dt[11:16] if len(dt) > 15 else ""
        ev    = m.get("home_team","")
        dep   = m.get("away_team","")

        # Ortalama oran hesapla (tüm bookmaker'lardan)
        o1 = ox = o2 = None
        bookmakers = m.get("bookmakers",[])
        if bookmakers:
            vals1=[]; valsx=[]; vals2=[]
            for bm in bookmakers:
                for mkt in bm.get("markets",[]):
                    if mkt.get("key")=="h2h":
                        for o in mkt.get("outcomes",[]):
                            if o["name"]==ev:        vals1.append(o["price"])
                            elif o["name"]=="Draw":  valsx.append(o["price"])
                            else:                    vals2.append(o["price"])
            if vals1: o1 = round(sum(vals1)/len(vals1), 2)
            if valsx: ox = round(sum(valsx)/len(valsx), 2)
            if vals2: o2 = round(sum(vals2)/len(vals2), 2)

        stil_veri(ws, row, 1, tarih, bg=bg)
        stil_veri(ws, row, 2, saat,  bg=bg)
        stil_veri(ws, row, 3, ev,    bg=bg, hizala="left")
        stil_veri(ws, row, 4, dep,   bg=bg, hizala="left")

        if o1 and ox and o2:
            stil_veri(ws, row, 5, o1, bg="C6EFCE" if o1==min(o1,ox,o2) else bg, bold=o1==min(o1,ox,o2), sayi_fmt="0.00")
            stil_veri(ws, row, 6, ox, bg="C6EFCE" if ox==min(o1,ox,o2) else bg, bold=ox==min(o1,ox,o2), sayi_fmt="0.00")
            stil_veri(ws, row, 7, o2, bg="C6EFCE" if o2==min(o1,ox,o2) else bg, bold=o2==min(o1,ox,o2), sayi_fmt="0.00")
            # İhtimal % (vig temizlenmiş)
            p1=1/o1; px=1/ox; p2=1/o2; ptop=p1+px+p2
            stil_veri(ws, row, 8,  round(p1/ptop*100,1), bg=bg, sayi_fmt="0.0")
            stil_veri(ws, row, 9,  round(px/ptop*100,1), bg=bg, sayi_fmt="0.0")
            stil_veri(ws, row, 10, round(p2/ptop*100,1), bg=bg, sayi_fmt="0.0")
            marj = round((ptop-1)*100,1)
            stil_veri(ws, row, 11, marj, bg=bg, sayi_fmt="0.0")
            favori = ev if o1==min(o1,ox,o2) else ("Beraberlik" if ox==min(o1,ox,o2) else dep)
            stil_veri(ws, row, 12, favori, bg=bg, bold=True)
        else:
            for c in range(5,13): ws.cell(row=row,column=c,value="-").alignment=Alignment(horizontal="center")

    ince_kenar(ws, 2, len(maclar)+2, 1, 12)

# ── Sayfa: İstatistik Özet ────────────────────────────────────────────────────
def yaz_istatistik(ws, lig_adi, fd_code):
    print(f"  İstatistikler hesaplanıyor: {lig_adi}")
    data = fd_get(f"competitions/{fd_code}/matches",
                  params={"status": "FINISHED", "limit": 100})
    if not data:
        ws["A1"] = "Veri alınamadı"; return

    maclar = [m for m in data.get("matches",[])
              if m.get("score",{}).get("fullTime",{}).get("home") is not None]

    if not maclar:
        ws["A1"] = "Tamamlanan maç yok"; return

    # Hesaplamalar
    n = len(maclar)
    ev_g=dep_g=beraberlik=toplam_gol=kg_var=iki_yas=uc_yas=0
    for m in maclar:
        h = m["score"]["fullTime"]["home"]
        a = m["score"]["fullTime"]["away"]
        if h > a:  ev_g += 1
        elif h==a: beraberlik += 1
        else:      dep_g += 1
        toplam_gol += h+a
        if h > 0 and a > 0: kg_var += 1
        if h+a > 2: iki_yas += 1
        if h+a > 3: uc_yas += 1

    ws.freeze_panes = "A3"
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A1:D1")
    stil_baslik(ws, 1, 1, f"{lig_adi} — Sezon İstatistikleri ({n} maç)",
                bg=RENKLER["baslik_bg"], boyut=13)

    # Başlıklar
    for i,b in enumerate(["İstatistik","Değer","Yüzde","Yorum"],1):
        stil_baslik(ws, 2, i, b, bg=RENKLER["alt_baslik"], boyut=10)

    col_gen(ws, [28, 10, 10, 25])

    satirlar = [
        ("Toplam Maç",            n,      None,   ""),
        ("Ev Sahibi Galibiyeti",  ev_g,   ev_g/n, "1 oynamak için baz oran"),
        ("Beraberlik",            beraberlik, beraberlik/n, "X için baz oran"),
        ("Deplasman Galibiyeti",  dep_g,  dep_g/n,"2 oynamak için baz oran"),
        ("Maç Başı Ort. Gol",     round(toplam_gol/n,2), None, ""),
        ("2.5 Üst (3+ gol)",      iki_yas, iki_yas/n, "2.5 üst bahisi için"),
        ("2.5 Alt (0-2 gol)",     n-iki_yas, (n-iki_yas)/n, "2.5 alt bahisi için"),
        ("3.5 Üst (4+ gol)",      uc_yas,  uc_yas/n, ""),
        ("Karşılıklı Gol (KG)",   kg_var,  kg_var/n, "Her iki takım da gol attı"),
        ("KG Yok",                n-kg_var,(n-kg_var)/n, ""),
    ]

    for idx, (isim, deger, pct, yorum) in enumerate(satirlar):
        row = idx + 3
        bg = RENKLER["satir1"] if idx % 2 == 0 else RENKLER["satir2"]
        stil_veri(ws, row, 1, isim,  bg=bg, hizala="left", bold=True)
        stil_veri(ws, row, 2, deger, bg=bg)
        if pct is not None:
            stil_veri(ws, row, 3, pct, bg=bg, sayi_fmt="0.0%")
        else:
            ws.cell(row=row, column=3, value="-").alignment=Alignment(horizontal="center")
        stil_veri(ws, row, 4, yorum, bg=bg, hizala="left")

    ince_kenar(ws, 2, len(satirlar)+2, 1, 4)

    # Takım bazlı gol istatistikleri
    takim_stats = {}
    for m in maclar:
        ev_t  = m.get("homeTeam",{}).get("name","")
        dep_t = m.get("awayTeam",{}).get("name","")
        h = m["score"]["fullTime"]["home"]
        a = m["score"]["fullTime"]["away"]
        for t,gol_at,gol_ye,ev_mi in [(ev_t,h,a,True),(dep_t,a,h,False)]:
            if t not in takim_stats:
                takim_stats[t]={"gol_at":0,"gol_ye":0,"mac":0,"galibiyet":0}
            takim_stats[t]["gol_at"]+=gol_at
            takim_stats[t]["gol_ye"]+=gol_ye
            takim_stats[t]["mac"]+=1
            if (ev_mi and h>a) or (not ev_mi and a>h):
                takim_stats[t]["galibiyet"]+=1

    # Takım tablosu
    bslk_row = len(satirlar) + 5
    ws.merge_cells(f"A{bslk_row}:F{bslk_row}")
    stil_baslik(ws, bslk_row, 1, "Takım Gol İstatistikleri",
                bg=RENKLER["alt_baslik"], boyut=11)
    for i,b in enumerate(["Takım","Maç","Gol Attı","Gol Yedi","Avg Att","Avg Yedi"],1):
        stil_baslik(ws, bslk_row+1, i, b, bg=RENKLER["alt_baslik"], boyut=10)

    sirali = sorted(takim_stats.items(), key=lambda x:-x[1]["gol_at"])
    for idx,(takim,s) in enumerate(sirali):
        row = bslk_row+2+idx
        bg = RENKLER["satir1"] if idx%2==0 else RENKLER["satir2"]
        mac = s["mac"] or 1
        stil_veri(ws,row,1,takim,           bg=bg,hizala="left")
        stil_veri(ws,row,2,s["mac"],        bg=bg)
        stil_veri(ws,row,3,s["gol_at"],     bg=bg)
        stil_veri(ws,row,4,s["gol_ye"],     bg=bg)
        stil_veri(ws,row,5,round(s["gol_at"]/mac,2),bg=bg,sayi_fmt="0.00")
        stil_veri(ws,row,6,round(s["gol_ye"]/mac,2),bg=bg,sayi_fmt="0.00")
    ince_kenar(ws, bslk_row, bslk_row+1+len(sirali), 1, 6)

# ── Özet Sayfası ─────────────────────────────────────────────────────────────
def yaz_ozet(ws):
    ws.merge_cells("A1:F1")
    stil_baslik(ws, 1, 1, f"İddaa Analiz Sistemi — Güncelleme: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
                bg=RENKLER["baslik_bg"], boyut=14)
    ws.row_dimensions[1].height = 35

    basliklar = ["Lig","Sekme","İçerik","Son Güncelleme"]
    for i,b in enumerate(basliklar,1):
        stil_baslik(ws, 2, i, b, bg=RENKLER["alt_baslik"])

    col_gen(ws, [20,25,40,20])

    sekmeler = [
        ("Premier League","PL - Puan Tablosu","Sıralama, form, gol farkı",""),
        ("Premier League","PL - Maçlar","Son 50 maç, skor, 1X2, İY",""),
        ("Premier League","PL - Oranlar","Yaklaşan maç oranları + ihtimal",""),
        ("Premier League","PL - İstatistik","1X2%, 2.5 üst/alt, KG oranları",""),
        ("La Liga","LL - Puan Tablosu","",""),
        ("La Liga","LL - Maçlar","",""),
        ("La Liga","LL - Oranlar","",""),
        ("La Liga","LL - İstatistik","",""),
        ("Bundesliga","BL - Puan Tablosu","",""),
        ("Bundesliga","BL - Maçlar","",""),
        ("Bundesliga","BL - Oranlar","",""),
        ("Bundesliga","BL - İstatistik","",""),
        ("Süper Lig","SL - Puan Tablosu","",""),
        ("Süper Lig","SL - Maçlar","",""),
        ("Süper Lig","SL - Oranlar","",""),
        ("Süper Lig","SL - İstatistik","",""),
    ]
    guncelleme = datetime.now().strftime("%d.%m.%Y %H:%M")
    for idx,(lig,sekme,icerik,_) in enumerate(sekmeler):
        row=idx+3
        bg=RENKLER["satir1"] if idx%2==0 else RENKLER["satir2"]
        stil_veri(ws,row,1,lig,    bg=bg,hizala="left",bold=True)
        stil_veri(ws,row,2,sekme,  bg=bg,hizala="left")
        stil_veri(ws,row,3,icerik, bg=bg,hizala="left")
        stil_veri(ws,row,4,guncelleme,bg=bg)

    ince_kenar(ws,2,len(sekmeler)+2,1,4)

    # Notlar
    not_row = len(sekmeler)+5
    ws.merge_cells(f"A{not_row}:F{not_row}")
    stil_baslik(ws,not_row,1,"NOTLAR",bg=RENKLER["alt_baslik"],boyut=11)
    notlar=[
        "• Bahis oranları birden fazla bookmaker'ın ortalamasıdır",
        "• İhtimal% = vig (marj) temizlenmiş gerçek olasılık tahmini",
        "• Marj% = bahisçinin karı (düşük marj = daha adil oran)",
        "• İstatistikler mevcut sezon maçlarına dayanmaktadır",
        "• Veriyi güncellemek için scripti tekrar çalıştırın",
    ]
    for i,n in enumerate(notlar):
        c = ws.cell(row=not_row+1+i, column=1, value=n)
        c.font=Font(name="Arial",size=10,italic=True)
        ws.merge_cells(f"A{not_row+1+i}:F{not_row+1+i}")

# ── ANA FONKSİYON ─────────────────────────────────────────────────────────────
def main():
    print("="*55)
    print("  İddaa Analiz Sistemi — Excel Raporu Oluşturuluyor")
    print("="*55)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # boş sayfayı sil

    # Özet sayfası
    ws_ozet = wb.create_sheet("OZET")
    yaz_ozet(ws_ozet)

    kisaltmalar = {"Premier League":"PL","La Liga":"LL","Bundesliga":"BL","Super Lig":"SL"}

    for lig_adi, bilgi in LIGLER.items():
        kisa = kisaltmalar.get(lig_adi, lig_adi[:2])
        print(f"\n[{lig_adi}]")

        ws1 = wb.create_sheet(f"{kisa} - Puan Tablosu")
        yaz_puan_tablosu(ws1, lig_adi, bilgi["fd_code"])

        ws2 = wb.create_sheet(f"{kisa} - Maclar")
        yaz_mac_sonuclari(ws2, lig_adi, bilgi["fd_code"])

        ws3 = wb.create_sheet(f"{kisa} - Oranlar")
        yaz_oranlar(ws3, lig_adi, bilgi["odds_key"])

        ws4 = wb.create_sheet(f"{kisa} - Istatistik")
        yaz_istatistik(ws4, lig_adi, bilgi["fd_code"])

    path = os.path.join(SCRIPT_DIR, "iddaa_analiz.xlsx")
    wb.save(path)
    print(f"\n{'='*55}")
    print(f"  Excel kaydedildi: {path}")
    print(f"  Toplam sayfa: {len(wb.sheetnames)}")
    print(f"{'='*55}")

if __name__ == "__main__":
    main()
