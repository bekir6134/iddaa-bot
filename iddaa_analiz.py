import requests
import openpyxl
import time
import re
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
    "Premier League": {"fd_code": "PL",  "odds_key": "soccer_epl",                "fd_destekli": True},
    "La Liga":        {"fd_code": "PD",  "odds_key": "soccer_spain_la_liga",       "fd_destekli": True},
    "Bundesliga":     {"fd_code": "BL1", "odds_key": "soccer_germany_bundesliga",  "fd_destekli": True},
    "Serie A":        {"fd_code": "SA",  "odds_key": "soccer_italy_serie_a",       "fd_destekli": True},
    "Ligue 1":        {"fd_code": "FL1", "odds_key": "soccer_france_ligue_one",    "fd_destekli": True},
    "Süper Lig":      {"fd_code": "TSL", "odds_key": "soccer_turkey_super_league", "fd_destekli": False},
}

RENKLER = {
    "baslik_bg":   "1F4E79",
    "baslik_yazi": "FFFFFF",
    "alt_baslik":  "2E75B6",
    "satir1":      "DEEAF1",
    "satir2":      "FFFFFF",
    "kazandi":     "C6EFCE",
    "kaybetti":    "FFC7CE",
    "beraberlik":  "FFEB9C",
}

_fd_call_count = 0

def fd_get(endpoint, params=None):
    global _fd_call_count
    if _fd_call_count > 0:
        time.sleep(7)
    _fd_call_count += 1
    r = requests.get(f"{FD_BASE}/{endpoint}",
                     headers={"X-Auth-Token": FD_KEY},
                     params=params, timeout=20)
    if r.status_code == 200:
        return r.json()
    print(f"  FD API hata {r.status_code}: {endpoint}")
    return None

def odds_get(sport_key):
    r = requests.get(f"{ODDS_BASE}/sports/{sport_key}/odds",
                     params={
                         "apiKey": ODDS_KEY, "regions": "eu",
                         "markets": "h2h,totals,btts,h2h_h1,asian_handicap",
                         "oddsFormat": "decimal",
                     }, timeout=20)
    if r.status_code == 200:
        return r.json()
    print(f"  Odds API hata {r.status_code}: {sport_key}")
    return []

# ── Takım adı normalize ───────────────────────────────────────────────────────
def normalize_takim(isim):
    """FD ve Odds API arasında takım adı eşleştirme için normalize et."""
    isim = re.sub(r'\b(FC|AFC|SC|CF|FK|SK|1\.)\b\.?', '', isim)
    return isim.strip().lower()

# ── Market yardımcıları ───────────────────────────────────────────────────────
def _ort(bookmakers, market_key, outcome_name):
    vals = []
    for bm in bookmakers:
        for mkt in bm.get("markets", []):
            if mkt.get("key") == market_key:
                for o in mkt.get("outcomes", []):
                    if o.get("name", "").lower() == outcome_name.lower():
                        vals.append(o["price"])
    return round(sum(vals) / len(vals), 2) if vals else None

def _ort_totals(bookmakers, point, over_under):
    vals = []
    for bm in bookmakers:
        for mkt in bm.get("markets", []):
            if mkt.get("key") == "totals":
                for o in mkt.get("outcomes", []):
                    if (o.get("name", "").lower() == over_under.lower() and
                            abs(float(o.get("point", 0)) - point) < 0.01):
                        vals.append(o["price"])
    return round(sum(vals) / len(vals), 2) if vals else None

def _ort_handicap(bookmakers, home_team):
    ev_vals = []; dep_vals = []; hcp_vals = []
    for bm in bookmakers:
        for mkt in bm.get("markets", []):
            if mkt.get("key") == "asian_handicap":
                for o in mkt.get("outcomes", []):
                    if o.get("name", "") == home_team:
                        ev_vals.append(o["price"])
                        hcp_vals.append(float(o.get("point", 0)))
                    else:
                        dep_vals.append(o["price"])
    ev  = round(sum(ev_vals)  / len(ev_vals),  2) if ev_vals  else None
    dep = round(sum(dep_vals) / len(dep_vals), 2) if dep_vals else None
    hcp = round(sum(hcp_vals) / len(hcp_vals), 2) if hcp_vals else None
    return ev, dep, hcp

# ── Form hesapla ──────────────────────────────────────────────────────────────
def takim_form_hesapla(maclar, takim, son_n=5):
    ilgili = []
    for m in sorted(maclar, key=lambda x: x.get("utcDate", "")):
        ht = m.get("homeTeam", {}).get("name", "")
        at = m.get("awayTeam", {}).get("name", "")
        if takim.lower() not in ht.lower() and takim.lower() not in at.lower():
            continue
        full = m.get("score", {}).get("fullTime", {})
        hg = full.get("home"); ag = full.get("away")
        if hg is None:
            continue
        ev_mi = takim.lower() in ht.lower()
        att  = hg if ev_mi else ag
        yedi = ag if ev_mi else hg
        sonuc = "G" if att > yedi else ("B" if att == yedi else "M")
        ilgili.append({"sonuc": sonuc, "att": att, "yedi": yedi, "ev_mi": ev_mi})

    son = ilgili[-son_n:]
    if not son:
        return {"form": "-", "g": 0, "b": 0, "m": 0,
                "avg_att": 0.0, "avg_yedi": 0.0, "ev_form": "-", "dep_form": "-"}
    return {
        "form": "-".join(x["sonuc"] for x in son),
        "g": sum(1 for x in son if x["sonuc"] == "G"),
        "b": sum(1 for x in son if x["sonuc"] == "B"),
        "m": sum(1 for x in son if x["sonuc"] == "M"),
        "avg_att":  round(sum(x["att"]  for x in son) / len(son), 1),
        "avg_yedi": round(sum(x["yedi"] for x in son) / len(son), 1),
        "ev_form":  "-".join(x["sonuc"] for x in son if x["ev_mi"])      or "-",
        "dep_form": "-".join(x["sonuc"] for x in son if not x["ev_mi"])  or "-",
    }

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
def yaz_puan_tablosu(ws, lig_adi, fd_code, mac_verisi=None):
    print(f"  Puan tablosu: {lig_adi}")
    data = fd_get(f"competitions/{fd_code}/standings")
    if not data:
        ws["A1"] = "Veri alınamadı"; return

    ws.freeze_panes = "A3"
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 20
    ws.merge_cells("A1:M1")
    stil_baslik(ws, 1, 1, f"{lig_adi} — Puan Tablosu ({datetime.now().strftime('%d.%m.%Y')})",
                bg=RENKLER["baslik_bg"], boyut=13)
    for i, b in enumerate(["Sıra","Takım","O","G","B","M","AG","YG","A","Puan","Form","Ev G%","Dep G%"], 1):
        stil_baslik(ws, 2, i, b, bg=RENKLER["alt_baslik"], boyut=10)
    col_gen(ws, [6, 22, 5, 5, 5, 5, 6, 6, 6, 7, 12, 8, 8])

    standings = data.get("standings", [{}])[0].get("table", [])

    ev_dep_stats = {}
    if mac_verisi:
        for m in mac_verisi:
            full = m.get("score", {}).get("fullTime", {})
            hg = full.get("home"); ag = full.get("away")
            if hg is None: continue
            ht = m.get("homeTeam", {}).get("name", "")
            at = m.get("awayTeam", {}).get("name", "")
            for t in [ht, at]:
                if t not in ev_dep_stats:
                    ev_dep_stats[t] = {"ev_g":0,"ev_mac":0,"dep_g":0,"dep_mac":0}
            ev_dep_stats[ht]["ev_mac"] += 1
            ev_dep_stats[at]["dep_mac"] += 1
            if hg > ag:  ev_dep_stats[ht]["ev_g"] += 1
            elif ag > hg: ev_dep_stats[at]["dep_g"] += 1

    for idx, t in enumerate(standings):
        row = idx + 3
        bg  = RENKLER["satir1"] if idx % 2 == 0 else RENKLER["satir2"]
        pos = t.get("position", "")
        if pos <= 1:                       bg_pos = "FFD700"
        elif pos <= 4:                     bg_pos = "C6EFCE"
        elif pos >= len(standings) - 2:    bg_pos = "FFC7CE"
        else:                              bg_pos = bg

        takim_adi = t.get("team", {}).get("name", "")
        stil_veri(ws, row, 1,  pos,                         bg=bg_pos, bold=True)
        stil_veri(ws, row, 2,  takim_adi,                   bg=bg, hizala="left", bold=True)
        stil_veri(ws, row, 3,  t.get("playedGames", ""),    bg=bg)
        stil_veri(ws, row, 4,  t.get("won", ""),            bg=bg)
        stil_veri(ws, row, 5,  t.get("draw", ""),           bg=bg)
        stil_veri(ws, row, 6,  t.get("lost", ""),           bg=bg)
        stil_veri(ws, row, 7,  t.get("goalsFor", ""),       bg=bg)
        stil_veri(ws, row, 8,  t.get("goalsAgainst", ""),   bg=bg)
        stil_veri(ws, row, 9,  t.get("goalDifference", ""), bg=bg)
        stil_veri(ws, row, 10, t.get("points", ""),         bg=bg, bold=True)
        stil_veri(ws, row, 11, t.get("form", "") or "-",    bg=bg)

        s = ev_dep_stats.get(takim_adi, {})
        ev_pct  = round(s.get("ev_g",0)  / (s.get("ev_mac",0)  or 1) * 100, 1) if s else None
        dep_pct = round(s.get("dep_g",0) / (s.get("dep_mac",0) or 1) * 100, 1) if s else None
        stil_veri(ws, row, 12, f"%{ev_pct}"  if ev_pct  is not None else "-", bg=bg)
        stil_veri(ws, row, 13, f"%{dep_pct}" if dep_pct is not None else "-", bg=bg)

    ince_kenar(ws, 2, len(standings) + 2, 1, 13)

# ── Sayfa: Maç Sonuçları ──────────────────────────────────────────────────────
def yaz_mac_sonuclari(ws, lig_adi, maclar):
    print(f"  Maç sonuçları: {lig_adi} ({len(maclar)} maç)")
    ws.freeze_panes = "A3"
    ws.row_dimensions[1].height = 30; ws.row_dimensions[2].height = 20
    ws.merge_cells("A1:J1")
    stil_baslik(ws, 1, 1, f"{lig_adi} — Son Maç Sonuçları",
                bg=RENKLER["baslik_bg"], boyut=13)
    for i, b in enumerate(["Tarih","Hafta","Ev Sahibi","Skor","Deplasman",
                            "Sonuç","İY Skor","Toplam Gol","KG","Durum"], 1):
        stil_baslik(ws, 2, i, b, bg=RENKLER["alt_baslik"], boyut=10)
    col_gen(ws, [12,7,22,8,22,8,9,11,6,10])

    for idx, m in enumerate(maclar):
        row = idx + 3
        bg  = RENKLER["satir1"] if idx % 2 == 0 else RENKLER["satir2"]
        full = m.get("score",{}).get("fullTime",{})
        ht_s = m.get("score",{}).get("halfTime",{})
        hg = full.get("home"); ag = full.get("away")
        ih = ht_s.get("home"); ia = ht_s.get("away")

        if hg is not None and ag is not None:
            skor   = f"{hg} - {ag}"
            iy     = f"{ih} - {ia}" if ih is not None else "-"
            toplam = hg + ag
            kg     = "Var" if hg > 0 and ag > 0 else "Yok"
            if hg > ag:    sonuc="1"; bg_s=RENKLER["kazandi"]
            elif hg == ag: sonuc="X"; bg_s=RENKLER["beraberlik"]
            else:          sonuc="2"; bg_s=RENKLER["kaybetti"]
        else:
            skor=iy="-"; toplam=""; kg=""; sonuc="-"; bg_s=bg

        stil_veri(ws, row, 1,  m.get("utcDate","")[:10],   bg=bg)
        stil_veri(ws, row, 2,  m.get("matchday",""),        bg=bg)
        stil_veri(ws, row, 3,  m.get("homeTeam",{}).get("name",""), bg=bg, hizala="left")
        stil_veri(ws, row, 4,  skor,                        bg=bg, bold=True)
        stil_veri(ws, row, 5,  m.get("awayTeam",{}).get("name",""), bg=bg, hizala="left")
        stil_veri(ws, row, 6,  sonuc,                       bg=bg_s, bold=True)
        stil_veri(ws, row, 7,  iy,                          bg=bg)
        stil_veri(ws, row, 8,  toplam,                      bg=bg)
        stil_veri(ws, row, 9,  kg,                          bg=bg)
        stil_veri(ws, row, 10, m.get("status",""),          bg=bg)

    ince_kenar(ws, 2, len(maclar)+2, 1, 10)

# ── Sayfa: Bahis Oranları ─────────────────────────────────────────────────────
def yaz_oranlar(ws, lig_adi, odds_key):
    print(f"  Bahis oranları: {lig_adi}")
    maclar = odds_get(odds_key)

    ws.freeze_panes = "A3"
    ws.row_dimensions[1].height = 30; ws.row_dimensions[2].height = 20

    basliklar = ["Tarih","Saat","Ev Sahibi","Deplasman",
                 "1","X","2","1%","X%","2%","Marj%","Favori",
                 "Over2.5","Under2.5","KG Evet","KG Hayır",
                 "İY-1","İY-X","İY-2",
                 "Hcp Ev","Hcp Dep","Hcp Değer"]
    N = len(basliklar)
    ws.merge_cells(f"A1:{get_column_letter(N)}1")
    stil_baslik(ws, 1, 1, f"{lig_adi} — Yaklaşan Maç Oranları",
                bg=RENKLER["baslik_bg"], boyut=13)
    for i, b in enumerate(basliklar, 1):
        stil_baslik(ws, 2, i, b, bg=RENKLER["alt_baslik"], boyut=10)
    col_gen(ws, [11,7,22,22, 7,7,7,8,8,8,7,14, 9,9,8,8, 7,7,7, 9,9,8])

    if not maclar:
        ws.cell(row=3, column=1, value="Yaklaşan maç bulunamadı")
        return

    for idx, m in enumerate(maclar):
        row = idx + 3
        bg  = RENKLER["satir1"] if idx % 2 == 0 else RENKLER["satir2"]
        dt  = m.get("commence_time","")
        ev  = m.get("home_team","")
        dep = m.get("away_team","")
        bms = m.get("bookmakers",[])

        o1 = _ort(bms,"h2h",ev); ox=_ort(bms,"h2h","Draw"); o2=_ort(bms,"h2h",dep)
        over25  = _ort_totals(bms, 2.5, "Over")
        under25 = _ort_totals(bms, 2.5, "Under")
        kg_evet  = _ort(bms,"btts","Yes")
        kg_hayir = _ort(bms,"btts","No")
        iy1=_ort(bms,"h2h_h1",ev); iyx=_ort(bms,"h2h_h1","Draw"); iy2=_ort(bms,"h2h_h1",dep)
        hcp_ev, hcp_dep, hcp_val = _ort_handicap(bms, ev)

        stil_veri(ws, row, 1, dt[:10] if dt else "",        bg=bg)
        stil_veri(ws, row, 2, dt[11:16] if len(dt)>15 else "",bg=bg)
        stil_veri(ws, row, 3, ev,  bg=bg, hizala="left")
        stil_veri(ws, row, 4, dep, bg=bg, hizala="left")

        if o1 and ox and o2:
            min_o = min(o1, ox, o2)
            for col, val in [(5,o1),(6,ox),(7,o2)]:
                stil_veri(ws,row,col,val, bg="C6EFCE" if val==min_o else bg,
                          bold=val==min_o, sayi_fmt="0.00")
            p1=1/o1; px=1/ox; p2=1/o2; pt=p1+px+p2
            stil_veri(ws,row,8,  round(p1/pt*100,1), bg=bg, sayi_fmt="0.0")
            stil_veri(ws,row,9,  round(px/pt*100,1), bg=bg, sayi_fmt="0.0")
            stil_veri(ws,row,10, round(p2/pt*100,1), bg=bg, sayi_fmt="0.0")
            stil_veri(ws,row,11, round((pt-1)*100,1),bg=bg, sayi_fmt="0.0")
            favori = ev if o1==min_o else ("Beraberlik" if ox==min_o else dep)
            stil_veri(ws,row,12, favori, bg=bg, bold=True)
        else:
            for c in range(5,13):
                ws.cell(row=row,column=c,value="-").alignment=Alignment(horizontal="center")

        for col, val in [(13,over25),(14,under25),(15,kg_evet),(16,kg_hayir),
                         (17,iy1),(18,iyx),(19,iy2),(20,hcp_ev),(21,hcp_dep)]:
            stil_veri(ws,row,col, val or "-", bg=bg, sayi_fmt="0.00" if val else None)
        stil_veri(ws,row,22, f"{hcp_val:+.1f}" if hcp_val is not None else "-", bg=bg)

    ince_kenar(ws, 2, len(maclar)+2, 1, N)

# ── Sayfa: Lig İstatistikleri (genişletilmiş) ─────────────────────────────────
def yaz_istatistik(ws, lig_adi, maclar):
    print(f"  Lig istatistikleri: {lig_adi}")
    tamam = [m for m in maclar
             if m.get("score",{}).get("fullTime",{}).get("home") is not None]
    if not tamam:
        ws["A1"] = "Tamamlanan maç yok"; return

    n = len(tamam)
    ev_g=dep_g=ber=toplam_gol=kg=ust15=ust25=ust35=0
    iy_ev=iy_ber=iy_dep=y2_ev=y2_ber=y2_dep=0

    for m in tamam:
        h  = m["score"]["fullTime"]["home"]
        a  = m["score"]["fullTime"]["away"]
        ih = m["score"].get("halfTime",{}).get("home") or 0
        ia = m["score"].get("halfTime",{}).get("away") or 0
        h2 = h-ih; a2 = a-ia

        if h>a:   ev_g+=1
        elif h==a: ber+=1
        else:      dep_g+=1
        toplam_gol += h+a
        if h>0 and a>0: kg+=1
        if h+a>1: ust15+=1
        if h+a>2: ust25+=1
        if h+a>3: ust35+=1
        if ih>ia:   iy_ev+=1
        elif ih==ia: iy_ber+=1
        else:        iy_dep+=1
        if h2>a2:   y2_ev+=1
        elif h2==a2: y2_ber+=1
        else:        y2_dep+=1

    ws.freeze_panes = "A3"
    ws.row_dimensions[1].height = 30
    ws.merge_cells("A1:D1")
    stil_baslik(ws, 1, 1, f"{lig_adi} — Sezon İstatistikleri ({n} maç)",
                bg=RENKLER["baslik_bg"], boyut=13)
    for i, b in enumerate(["İstatistik","Değer","Yüzde","Yorum"], 1):
        stil_baslik(ws, 2, i, b, bg=RENKLER["alt_baslik"], boyut=10)
    col_gen(ws, [30, 10, 10, 30])

    def ayrac(ws, row, metin):
        ws.merge_cells(f"A{row}:D{row}")
        c = ws.cell(row=row, column=1, value=metin)
        c.font = Font(name="Arial", bold=True, color=RENKLER["baslik_yazi"], size=9)
        c.fill = PatternFill("solid", start_color=RENKLER["alt_baslik"])
        c.alignment = Alignment(horizontal="center", vertical="center")

    satirlar = [
        ("toplam", "Toplam Maç", n, None, ""),
        ("ayrac",  "─── TAM SKOR ───", None, None, None),
        ("veri",   "Ev Sahibi Galibiyeti", ev_g,   ev_g/n,        "1 için baz oran"),
        ("veri",   "Beraberlik",           ber,    ber/n,         "X için baz oran"),
        ("veri",   "Deplasman Galibiyeti", dep_g,  dep_g/n,       "2 için baz oran"),
        ("veri",   "Maç Başı Ort. Gol",   round(toplam_gol/n,2), None, ""),
        ("veri",   "1.5 Üst (2+ gol)",    ust15,  ust15/n,       ""),
        ("veri",   "2.5 Üst (3+ gol)",    ust25,  ust25/n,       "2.5 üst bahisi"),
        ("veri",   "3.5 Üst (4+ gol)",    ust35,  ust35/n,       ""),
        ("veri",   "Karşılıklı Gol (KG)", kg,     kg/n,          "Her iki takım gol attı"),
        ("veri",   "KG Yok",              n-kg,   (n-kg)/n,      ""),
        ("ayrac",  "─── İLK YARI ───", None, None, None),
        ("veri",   "İY Ev Galibiyeti",        iy_ev,  iy_ev/n,   ""),
        ("veri",   "İY Beraberlik",           iy_ber, iy_ber/n,  ""),
        ("veri",   "İY Deplasman Galibiyeti", iy_dep, iy_dep/n,  ""),
        ("ayrac",  "─── İKİNCİ YARI (FT-HT) ───", None, None, None),
        ("veri",   "2Y Ev Galibiyeti",        y2_ev,  y2_ev/n,   ""),
        ("veri",   "2Y Beraberlik",           y2_ber, y2_ber/n,  ""),
        ("veri",   "2Y Deplasman Galibiyeti", y2_dep, y2_dep/n,  ""),
    ]

    for idx, satir in enumerate(satirlar):
        row = idx + 3
        tip = satir[0]
        if tip == "ayrac":
            ayrac(ws, row, satir[1])
            continue
        isim, deger, pct, yorum = satir[1], satir[2], satir[3], satir[4]
        bg = RENKLER["satir1"] if idx % 2 == 0 else RENKLER["satir2"]
        stil_veri(ws, row, 1, isim,  bg=bg, hizala="left", bold=True)
        stil_veri(ws, row, 2, deger, bg=bg)
        if pct is not None:
            stil_veri(ws, row, 3, pct, bg=bg, sayi_fmt="0.0%")
        else:
            ws.cell(row=row, column=3, value="-").alignment = Alignment(horizontal="center")
        stil_veri(ws, row, 4, yorum, bg=bg, hizala="left")

    ince_kenar(ws, 2, len(satirlar)+2, 1, 4)

# ── Sayfa: Takım Bazlı Detaylı İstatistik ────────────────────────────────────
def yaz_takim_istatistik(ws, lig_adi, maclar):
    """Bot bu sayfayı okuyarak maç analizine takım bazlı veri ekler."""
    print(f"  Takım istatistikleri: {lig_adi}")
    tamam = [m for m in sorted(maclar, key=lambda x: x.get("utcDate",""))
             if m.get("score",{}).get("fullTime",{}).get("home") is not None]
    if not tamam:
        ws["A1"] = "Tamamlanan maç yok"; return

    ws.freeze_panes = "A3"
    ws.row_dimensions[1].height = 30; ws.row_dimensions[2].height = 20

    basliklar = [
        "Takım (Normalize)",
        "Ev Maç","Ev G%","Ev Att","Ev Yedi","Ev Üst15%","Ev Üst25%","Ev KG%",
        "Dep Maç","Dep G%","Dep Att","Dep Yedi","Dep Üst15%","Dep Üst25%","Dep KG%",
        "Son5 Form","Son5 Att","Son5 Yedi",
    ]
    N = len(basliklar)
    ws.merge_cells(f"A1:{get_column_letter(N)}1")
    stil_baslik(ws, 1, 1, f"{lig_adi} — Takım Bazlı İstatistikler",
                bg=RENKLER["baslik_bg"], boyut=13)
    for i, b in enumerate(basliklar, 1):
        stil_baslik(ws, 2, i, b, bg=RENKLER["alt_baslik"], boyut=10)
    col_gen(ws, [22, 8,8,8,8,9,9,8, 8,8,8,8,9,9,8, 14,8,8])

    stats = {}
    for m in tamam:
        ht = m.get("homeTeam",{}).get("name","")
        at = m.get("awayTeam",{}).get("name","")
        h  = m["score"]["fullTime"]["home"]
        a  = m["score"]["fullTime"]["away"]
        for t in [ht, at]:
            if t not in stats:
                stats[t] = {
                    "ev_mac":0,"ev_g":0,"ev_att":0,"ev_yedi":0,"ev_ust15":0,"ev_ust25":0,"ev_kg":0,
                    "dep_mac":0,"dep_g":0,"dep_att":0,"dep_yedi":0,"dep_ust15":0,"dep_ust25":0,"dep_kg":0,
                }
        stats[ht]["ev_mac"]+=1; stats[ht]["ev_att"]+=h; stats[ht]["ev_yedi"]+=a
        if h>a:       stats[ht]["ev_g"]+=1
        if h+a>1:     stats[ht]["ev_ust15"]+=1
        if h+a>2:     stats[ht]["ev_ust25"]+=1
        if h>0 and a>0: stats[ht]["ev_kg"]+=1

        stats[at]["dep_mac"]+=1; stats[at]["dep_att"]+=a; stats[at]["dep_yedi"]+=h
        if a>h:       stats[at]["dep_g"]+=1
        if h+a>1:     stats[at]["dep_ust15"]+=1
        if h+a>2:     stats[at]["dep_ust25"]+=1
        if h>0 and a>0: stats[at]["dep_kg"]+=1

    for idx, (takim, s) in enumerate(sorted(stats.items())):
        row = idx + 3
        bg  = RENKLER["satir1"] if idx % 2 == 0 else RENKLER["satir2"]
        em  = s["ev_mac"]  or 1
        dm  = s["dep_mac"] or 1
        form = takim_form_hesapla(tamam, takim, son_n=5)

        stil_veri(ws, row, 1,  normalize_takim(takim),                   bg=bg, hizala="left", bold=True)
        stil_veri(ws, row, 2,  s["ev_mac"],                              bg=bg)
        stil_veri(ws, row, 3,  f"%{round(s['ev_g']/em*100,0):.0f}",      bg=bg)
        stil_veri(ws, row, 4,  round(s["ev_att"]/em, 2),                 bg=bg, sayi_fmt="0.00")
        stil_veri(ws, row, 5,  round(s["ev_yedi"]/em, 2),                bg=bg, sayi_fmt="0.00")
        stil_veri(ws, row, 6,  f"%{round(s['ev_ust15']/em*100,0):.0f}",  bg=bg)
        stil_veri(ws, row, 7,  f"%{round(s['ev_ust25']/em*100,0):.0f}",  bg=bg)
        stil_veri(ws, row, 8,  f"%{round(s['ev_kg']/em*100,0):.0f}",     bg=bg)
        stil_veri(ws, row, 9,  s["dep_mac"],                             bg=bg)
        stil_veri(ws, row, 10, f"%{round(s['dep_g']/dm*100,0):.0f}",     bg=bg)
        stil_veri(ws, row, 11, round(s["dep_att"]/dm, 2),                bg=bg, sayi_fmt="0.00")
        stil_veri(ws, row, 12, round(s["dep_yedi"]/dm, 2),               bg=bg, sayi_fmt="0.00")
        stil_veri(ws, row, 13, f"%{round(s['dep_ust15']/dm*100,0):.0f}", bg=bg)
        stil_veri(ws, row, 14, f"%{round(s['dep_ust25']/dm*100,0):.0f}", bg=bg)
        stil_veri(ws, row, 15, f"%{round(s['dep_kg']/dm*100,0):.0f}",    bg=bg)
        stil_veri(ws, row, 16, form["form"],                              bg=bg)
        stil_veri(ws, row, 17, form["avg_att"],                           bg=bg, sayi_fmt="0.0")
        stil_veri(ws, row, 18, form["avg_yedi"],                          bg=bg, sayi_fmt="0.0")

    ince_kenar(ws, 2, len(stats)+2, 1, N)

# ── Özet Sayfası ─────────────────────────────────────────────────────────────
def yaz_ozet(ws):
    ws.merge_cells("A1:F1")
    stil_baslik(ws, 1, 1,
                f"İddaa Analiz Sistemi — Güncelleme: {datetime.now().strftime('%d.%m.%Y %H:%M')}",
                bg=RENKLER["baslik_bg"], boyut=14)
    ws.row_dimensions[1].height = 35
    for i, b in enumerate(["Lig","Sekme","İçerik","Son Güncelleme"], 1):
        stil_baslik(ws, 2, i, b, bg=RENKLER["alt_baslik"])
    col_gen(ws, [20, 25, 55, 20])

    kisaltmalar = {"Premier League":"PL","La Liga":"LL","Bundesliga":"BL",
                   "Serie A":"SA","Ligue 1":"L1","Süper Lig":"SL"}
    sekmeler = []
    for lig, bilgi in LIGLER.items():
        kisa = kisaltmalar.get(lig, lig[:2])
        if bilgi["fd_destekli"]:
            sekmeler += [
                (lig, f"{kisa} - Puan Tablosu", "Sıralama, form, gol farkı, ev/dep %"),
                (lig, f"{kisa} - Maçlar",       "Son maç sonuçları"),
                (lig, f"{kisa} - Oranlar",      "1X2, Alt/Üst, KG, İY, Handikap"),
                (lig, f"{kisa} - Istatistik",   "1X2%, 1.5/2.5üst, KG, İY, 2Y"),
                (lig, f"{kisa} - Takim",        "Ev/Dep ayrımlı takım bazlı istatistik"),
            ]
        else:
            sekmeler += [
                (lig, f"{kisa} - Oranlar", "Bahis oranları (detaylı veri API kısıtı nedeniyle mevcut değil)"),
            ]

    guncelleme = datetime.now().strftime("%d.%m.%Y %H:%M")
    for idx, (lig, sekme, icerik) in enumerate(sekmeler):
        row = idx + 3
        bg  = RENKLER["satir1"] if idx % 2 == 0 else RENKLER["satir2"]
        stil_veri(ws, row, 1, lig,        bg=bg, hizala="left", bold=True)
        stil_veri(ws, row, 2, sekme,      bg=bg, hizala="left")
        stil_veri(ws, row, 3, icerik,     bg=bg, hizala="left")
        stil_veri(ws, row, 4, guncelleme, bg=bg)

    ince_kenar(ws, 2, len(sekmeler)+2, 1, 4)

    not_row = len(sekmeler) + 5
    ws.merge_cells(f"A{not_row}:F{not_row}")
    stil_baslik(ws, not_row, 1, "NOTLAR", bg=RENKLER["alt_baslik"], boyut=11)
    for i, n in enumerate([
        "• Süper Lig için puan tablosu/istatistik mevcut değil (API kısıtı), sadece oranlar gösterilir",
        "• Bahis oranları birden fazla bookmaker'ın ortalamasıdır",
        "• İhtimal% = vig temizlenmiş gerçek olasılık tahmini",
        "• Alt/Üst ve KG oranları gerçek bahis marketinden çekilmektedir",
        "• Takım istatistikleri mevcut sezon maçlarına dayanır",
    ]):
        c = ws.cell(row=not_row+1+i, column=1, value=n)
        c.font = Font(name="Arial", size=10, italic=True)
        ws.merge_cells(f"A{not_row+1+i}:F{not_row+1+i}")

# ── ANA FONKSİYON ─────────────────────────────────────────────────────────────
def main():
    print("="*58)
    print("  İddaa Analiz Sistemi — Excel Raporu Oluşturuluyor")
    print("="*58)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws_ozet = wb.create_sheet("OZET")
    yaz_ozet(ws_ozet)

    kisaltmalar = {"Premier League":"PL","La Liga":"LL","Bundesliga":"BL",
                   "Serie A":"SA","Ligue 1":"L1","Süper Lig":"SL"}

    for lig_adi, bilgi in LIGLER.items():
        kisa = kisaltmalar.get(lig_adi, lig_adi[:2])
        print(f"\n[{lig_adi}]")

        if bilgi["fd_destekli"]:
            print(f"  Maçlar çekiliyor...")
            mac_data = fd_get(f"competitions/{bilgi['fd_code']}/matches",
                              params={"status":"FINISHED","limit":100})
            mac_listesi = []
            if mac_data:
                mac_listesi = [m for m in mac_data.get("matches",[])
                               if m.get("score",{}).get("fullTime",{}).get("home") is not None]

            ws1 = wb.create_sheet(f"{kisa} - Puan Tablosu")
            yaz_puan_tablosu(ws1, lig_adi, bilgi["fd_code"], mac_verisi=mac_listesi)

            ws2 = wb.create_sheet(f"{kisa} - Maclar")
            yaz_mac_sonuclari(ws2, lig_adi, mac_listesi[:50])

            ws4 = wb.create_sheet(f"{kisa} - Istatistik")
            yaz_istatistik(ws4, lig_adi, mac_listesi)

            ws5 = wb.create_sheet(f"{kisa} - Takim")
            yaz_takim_istatistik(ws5, lig_adi, mac_listesi)
        else:
            print(f"  FD destekli değil, atlanıyor...")
            for suffix in ["Puan Tablosu","Maclar","Istatistik","Takim"]:
                ws = wb.create_sheet(f"{kisa} - {suffix}")
                ws["A1"] = f"{lig_adi} için detaylı veri mevcut değil (API kısıtı)"

        ws3 = wb.create_sheet(f"{kisa} - Oranlar")
        yaz_oranlar(ws3, lig_adi, bilgi["odds_key"])

    path = os.path.join(SCRIPT_DIR, "iddaa_analiz.xlsx")
    wb.save(path)
    print(f"\n{'='*58}")
    print(f"  Excel kaydedildi: {path}")
    print(f"  Toplam sayfa: {len(wb.sheetnames)}")
    print(f"{'='*58}")

if __name__ == "__main__":
    main()
