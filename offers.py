"""
Персональна пропозиція для кожного ліда: макет сайту 1920x1080 (JPG) + лист чеською.

Запуск:  python offers.py output/leady_2026-10-18_1156_ceny.xlsx
         python offers.py output/leady_….xlsx --limit 20      # тільки перші 20
Результат: папка offers/ з картинками + offers/nabidky.xlsx (тема, текст листа, WhatsApp, файл).
Свої дані (ім'я, телефон) – у config.json.
"""
import argparse
import base64
import hashlib
import html
import io
import json
import os
import re
import unicodedata
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image

import lead_finder as lf
from reprice import load

HERE = Path(__file__).resolve().parent
CONFIG = HERE / "config.json"
DEFAULT_CONFIG = {
    "sender_name": "Tvoje Jméno",
    "sender_phone": "+420 000 000 000",
    "sender_email": "tvuj@email.cz",
    "sender_web": "",
    "delivery_days": 7,
    "smtp_host": "smtp.gmail.com",
    "smtp_port": 587,
    "smtp_user": "tvuj@email.cz",
    "smtp_password": "APP-PASSWORD",
    "daily_limit": 25,
    "delay_seconds": [90, 240],
}

# ---------------------------------------------------------------- вміст по нішах
# tagline, CTA, 4 послуги (назва, опис), палітра, віджет у hero
NICHES = {
    "kadeřnictví": ("Účes, se kterým se budete cítit skvěle", "Objednat se online",
                    [("Dámské střihy", "Střih na míru tvaru obličeje"), ("Barvení & melír", "Balayage, ombré, tónování"),
                     ("Pánské střihy", "Klasika i moderní styl"), ("Společenské účesy", "Svatby, plesy, focení")],
                    "beauty", "booking"),
    "kosmetický salon": ("Péče o pleť s viditelným výsledkem", "Rezervovat termín",
                         [("Ošetření pleti", "Čištění, hydratace, anti-age"), ("Lifting řas", "Přirozeně zvýrazněný pohled"),
                          ("Úprava obočí", "Tvarování a barvení"), ("Masáž obličeje", "Relax a regenerace")],
                         "beauty", "booking"),
    "manikúra": ("Krásné nehty bez čekání", "Rezervovat termín",
                 [("Gelové nehty", "Modeláž a doplnění"), ("Manikúra", "Klasická i japonská"),
                  ("Pedikúra", "Mokrá i suchá"), ("Nail art", "Zdobení dle přání")], "beauty", "booking"),
    "masáže": ("Odpočinek pro tělo i mysl", "Rezervovat masáž",
               [("Klasická masáž", "Uvolnění zad a šíje"), ("Relaxační masáž", "Aromaterapie a klid"),
                ("Sportovní masáž", "Regenerace po výkonu"), ("Dárkové poukazy", "Radost pro vaše blízké")],
               "spa", "booking"),
    "barbershop": ("Střih, vousy a styl", "Objednat se",
                   [("Pánský střih", "Fade, klasika, nůžky"), ("Úprava vousů", "Tvarování a péče"),
                    ("Hot towel", "Holení břitvou"), ("Dětský střih", "Pro malé gentlemany")], "dark", "booking"),
    "autoservis": ("Spolehlivý servis vašeho auta", "Objednat servis",
                   [("Servisní prohlídky", "Dle předpisu výrobce"), ("Diagnostika", "Rychlé odhalení závad"),
                    ("Brzdy a podvozek", "Opravy a výměny"), ("Příprava na STK", "Bez zbytečných starostí")],
                   "auto", "quote"),
    "pneuservis": ("Přezutí bez front a čekání", "Rezervovat přezutí",
                   [("Přezutí pneu", "Osobní i SUV"), ("Uskladnění", "Pneu i kola v suchu"),
                    ("Vyvážení kol", "Moderní technologie"), ("Prodej pneu", "Výhodné ceny")], "auto", "booking"),
    "autoškola": ("Řidičák snadno a v klidu", "Přihlásit se",
                  [("Skupina B", "Osobní automobil"), ("Kondiční jízdy", "Pro jistotu za volantem"),
                   ("Skupina A", "Motocykly"), ("Školení řidičů", "Profesní průkazy")], "auto", "booking"),
    "zubní ordinace": ("Zdravý a krásný úsměv", "Objednat se",
                       [("Preventivní prohlídky", "Pro děti i dospělé"), ("Dentální hygiena", "Šetrně a důkladně"),
                        ("Bělení zubů", "Zářivý úsměv"), ("Implantáty", "Moderní náhrady")], "medical", "booking"),
    "veterinář": ("Péče o vaše mazlíčky", "Objednat se",
                  [("Očkování", "Psi, kočky, králíci"), ("Preventivní prohlídky", "Zdravý mazlíček"),
                   ("Chirurgie", "Kastrace a zákroky"), ("Pohotovost", "Když spěcháte")], "medical", "booking"),
    "restaurace": ("Poctivá kuchyně a skvělá atmosféra", "Rezervovat stůl",
                   [("Polední menu", "Každý všední den"), ("Stálý lístek", "Sezónní suroviny"),
                    ("Rezervace stolu", "Online za pár vteřin"), ("Oslavy a akce", "Prostor pro vaše chvíle")],
                   "food", "menu"),
    "kavárna": ("Výběrová káva a domácí dorty", "Prohlédnout menu",
                [("Výběrová káva", "Z lokálních pražíren"), ("Domácí dorty", "Pečeme každý den"),
                 ("Snídaně", "Celý den"), ("Káva s sebou", "I do vlastního hrnku")], "food", "menu"),
    "pekárna": ("Čerstvé pečivo každé ráno", "Objednat online",
                [("Chléb", "Kváskový a poctivý"), ("Sladké pečivo", "Koláče, buchty, rohlíčky"),
                 ("Dorty na objednávku", "K oslavám i svatbám"), ("Catering", "Pro firmy i akce")], "food", "menu"),
    "květinářství": ("Květiny pro každou příležitost", "Objednat kytici",
                     [("Řezané květiny", "Čerstvé každý den"), ("Svatební kytice", "Na míru vašemu dni"),
                      ("Smuteční vazby", "Citlivě a důstojně"), ("Rozvoz květin", "Ještě dnes")], "flowers", "quote"),
    "fotograf": ("Fotky, které vyprávějí příběh", "Domluvit focení",
                 [("Svatby", "Celý den i obřad"), ("Portréty", "Osobní i firemní"),
                  ("Rodinné focení", "V ateliéru i venku"), ("Produktové foto", "Pro e-shopy")], "dark", "quote"),
    "instalatér": ("Voda, topení, plyn – rychle a spolehlivě", "Poptat opravu",
                   [("Havarijní služba", "Rychlý výjezd"), ("Koupelny", "Rekonstrukce na klíč"),
                    ("Topení", "Kotle a radiátory"), ("Odpady", "Čištění a opravy")], "crafts", "quote"),
    "elektrikář": ("Elektroinstalace bez starostí", "Poptat práci",
                   [("Elektroinstalace", "Byty, domy, firmy"), ("Revize", "Včetně zprávy"),
                    ("Opravy poruch", "Rychlý výjezd"), ("Chytrá domácnost", "Moderní řešení")], "crafts", "quote"),
    "truhlářství": ("Nábytek na míru z poctivého dřeva", "Nezávazná poptávka",
                    [("Kuchyně na míru", "Návrh i montáž"), ("Vestavěné skříně", "Každý centimetr využitý"),
                     ("Dveře a schody", "Masiv i dýha"), ("Renovace", "Nový život nábytku")], "crafts", "quote"),
    "malíř pokojů": ("Čisté malování bez stresu", "Nezávazná kalkulace",
                     [("Malování interiérů", "Rychle a čistě"), ("Stěrky", "Hladké stěny"),
                      ("Tapetování", "Moderní tapety"), ("Fasády", "Nátěry a opravy")], "crafts", "quote"),
    "stavební firma": ("Stavíme kvalitně a v termínu", "Poptat stavbu",
                       [("Rodinné domy", "Na klíč"), ("Rekonstrukce", "Byty i domy"),
                        ("Zateplení", "Úspora energie"), ("Projekce", "Od návrhu po kolaudaci")], "pro", "quote"),
    "účetní": ("Účetnictví, o které se nemusíte starat", "Nezávazná konzultace",
               [("Vedení účetnictví", "Pro OSVČ i firmy"), ("Daňová přiznání", "Včas a správně"),
                ("Mzdy", "Kompletní agenda"), ("Poradenství", "Daně bez stresu")], "pro", "quote"),
    "fitness": ("Tvoje forma začíná tady", "Vyzkoušet zdarma",
                [("Posilovna", "Moderní stroje"), ("Skupinové lekce", "Jóga, kruháč, spinning"),
                 ("Osobní trenér", "Plán na míru"), ("Permanentky", "Výhodné ceny")], "fitness", "booking"),
    "realitní kancelář": ("Prodáme vaši nemovitost rychle a výhodně", "Ocenit nemovitost",
                          [("Prodej", "Za nejlepší cenu"), ("Pronájem", "Spolehliví nájemníci"),
                           ("Odhad ceny", "Zdarma a do 24 h"), ("Právní servis", "Bezpečný převod")], "pro", "quote"),
    "úklidové služby": ("Čisto bez námahy", "Spočítat cenu",
                        [("Úklid domácností", "Pravidelně i jednorázově"), ("Úklid kanceláří", "Mimo pracovní dobu"),
                         ("Mytí oken", "Bez šmouh"), ("Po rekonstrukci", "Kompletní úklid")], "medical", "quote"),
    "zámečnictví": ("Zámečník, když ho potřebujete", "Zavolat ihned",
                    [("Otevírání dveří", "Nonstop výjezd"), ("Výměna zámků", "Bezpečnostní vložky"),
                     ("Kovové konstrukce", "Zábradlí, brány"), ("Bezpečnostní dveře", "Montáž a servis")],
                    "crafts", "quote"),
}
FALLBACK = ("Profesionální služby ve vašem okolí", "Kontaktovat",
            [("Služby", "Kvalitně a včas"), ("Poradenství", "Rádi poradíme"),
             ("Reference", "Spokojení zákazníci"), ("Kontakt", "Ozvěte se nám")], "pro", "quote")

PALETTES = {  # bg, surface, text, muted, accent, accent2
    "beauty":  ("#FBF6F3", "#FFFFFF", "#2A1E22", "#7A6A6F", "#B4536F", "#E8B4A6"),
    "spa":     ("#F4F1EA", "#FFFFFF", "#23302A", "#6C7A72", "#5E8B6F", "#C9B79C"),
    "dark":    ("#0F0F10", "#1A1A1C", "#F4F1EA", "#A19D96", "#C9A063", "#5A4A32"),
    "auto":    ("#0E1624", "#16213A", "#F1F5F9", "#94A3B8", "#F05A28", "#2B3A55"),
    "medical": ("#F3F8FA", "#FFFFFF", "#0F2A3A", "#5C7382", "#0E8F96", "#9ED8DB"),
    "food":    ("#FAF4EA", "#FFFFFF", "#2B1D14", "#7B6655", "#B4532A", "#E9C79A"),
    "flowers": ("#FCF5F7", "#FFFFFF", "#2D1F2A", "#7D6B77", "#C2477A", "#A8C69F"),
    "crafts":  ("#F5F3EE", "#FFFFFF", "#1F2A24", "#66726B", "#2F6B4F", "#D7B98E"),
    "pro":     ("#F4F6FA", "#FFFFFF", "#0E1B33", "#5B6782", "#1F4FD1", "#B9C8F2"),
    "fitness": ("#0B0C0E", "#16181C", "#F5F7F2", "#9EA39A", "#A3E635", "#2C3320"),
}

ICONS = [  # прості лінійні іконки (24x24, stroke)
    '<path d="M20 6 9 17l-5-5"/>',
    '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/>',
    '<path d="m12 3 2.7 5.6 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.9 1-6.1-4.4-4.3 6.1-.9z"/>',
    '<path d="M12 21s-7-6.2-7-11.5A7 7 0 0 1 19 9.5C19 14.8 12 21 12 21z"/><circle cx="12" cy="9.5" r="2.5"/>',
]


def slugify(s):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().lower()
    s = re.sub(r"\b(s\.?r\.?o|spol|a\.?s|v\.?o\.?s|k\.?s)\b\.?", "", s)
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-")[:40] or "firma"


def domain_for(lead):
    w = lead.get("website") or ""
    host = re.sub(r"^https?://", "", w).split("/")[0].lower()
    if host and not any(x in host for x in lf.SOCIAL + lf.FREE_BUILDERS):
        return host.replace("www.", "")
    return slugify(short_name(lead["name"])) + ".cz"


def short_name(name):
    n = re.sub(r",?\s*(s\.\s?r\.\s?o\.|spol\. s r\.\s?o\.|a\.\s?s\.|v\.\s?o\.\s?s\.)\s*$", "", name, flags=re.I)
    n = re.split(r"\s+[-–|]\s+", n)[0].strip()
    return n if len(n) <= 34 else n[:32].rsplit(" ", 1)[0] + "…"


def esc(s):
    return html.escape(str(s or ""))


def fetch_photo(lead):
    """Фото з профілю Firmy.cz як data-URI (або None)."""
    url = lead.get("image")
    if not url and lead.get("firmy_url"):
        try:
            h = lf.firmy_get(lead["firmy_url"]).text
            for ld in lf.LD_RE.findall(h):
                d = json.loads(ld)
                if d.get("@type") == "LocalBusiness":
                    url = d.get("image") or ""
                    lead["description"] = lead.get("description") or d.get("description", "")
        except Exception:
            url = ""
    if not url:
        return None
    try:
        r = lf.get(url + ("&" if "?" in url else "?") + "fl=res,1600,,3", tries=2, timeout=20)
        if r.status_code != 200 or len(r.content) < 8000:  # дрібне = швидше за все лого
            return None
        mime = r.headers.get("content-type", "image/jpeg").split(";")[0]
        lead["_photo_w"] = Image.open(io.BytesIO(r.content)).size[0]
        return f"data:{mime};base64,{base64.b64encode(r.content).decode()}"
    except Exception:
        return None


# ---------------------------------------------------------------- HTML макет
def widget(kind, pal, cta, lead):
    bg, surf, txt, mut, acc, acc2 = pal
    if kind == "booking":
        slots = "".join(f'<span class="slot{" on" if i == 1 else ""}">{t}</span>'
                        for i, t in enumerate(["9:00", "10:30", "13:00", "14:30", "16:00", "17:30"]))
        return f'''<div class="wcard"><div class="wt">Online rezervace</div>
          <div class="wsub">Volné termíny – zítra</div><div class="slots">{slots}</div>
          <div class="wbtn">{esc(cta)}</div></div>'''
    if kind == "menu":
        rows = [("Polévka dne", "65 Kč"), ("Hlavní jídlo dne", "169 Kč"), ("Domácí dezert", "89 Kč")]
        r = "".join(f'<div class="mrow"><span>{a}</span><b>{b}</b></div>' for a, b in rows)
        return f'''<div class="wcard"><div class="wt">Dnešní nabídka</div>{r}
          <div class="wbtn">{esc(cta)}</div></div>'''
    return f'''<div class="wcard"><div class="wt">Nezávazná poptávka</div>
      <div class="field">Jméno</div><div class="field">Telefon</div><div class="field tall">Co potřebujete?</div>
      <div class="wbtn">Odeslat poptávku</div></div>'''


def render_html(lead, photo):
    tagline, cta, services, pal_key, wkind = NICHES.get(lead.get("niche"), FALLBACK)
    bg, surf, txt, mut, acc, acc2 = PALETTES[pal_key]
    dark = pal_key in ("dark", "auto", "fitness")
    name = short_name(lead["name"])
    city = (lead.get("city") or "").split(",")
    district = city[-1].strip() if city else ""
    town = city[0].strip() if city else ""
    phone = lead.get("phone") or ""
    # фото на весь екран тільки якщо воно достатньо чітке, інакше – фото в рамці
    sharp = photo and lead.get("_photo_w", 0) >= 1500
    variant = "full" if sharp and int(hashlib.md5(lead["name"].encode()).hexdigest(), 16) % 2 else "split"
    reviews = int(lead.get("reviews") or 0)
    rating = lead.get("rating")
    badge = (f'<div class="badge"><svg viewBox="0 0 24 24" class="star">{ICONS[2]}</svg>'
             f'{esc(rating)} · {reviews} hodnocení na Firmy.cz</div>' if reviews >= 3 and float(rating or 0) >= 4.3
             else f'<div class="badge">{esc(lead.get("niche", "").capitalize())} · {esc(town)}</div>')
    menu = "".join(f"<span>{m}</span>" for m in ["Služby", "Ceník", "O nás", "Galerie", "Kontakt"])
    cards = "".join(f'''<div class="card"><div class="ic"><svg viewBox="0 0 24 24">{ICONS[i % 4]}</svg></div>
        <div class="ct">{esc(t)}</div><div class="cd">{esc(d)}</div></div>''' for i, (t, d) in enumerate(services))
    if photo:
        visual = f'<div class="photo" style="background-image:url({photo})"></div>'
    else:
        initials = "".join(w[0] for w in re.findall(r"[A-Za-zÀ-ž]+", name)[:2]).upper()
        visual = f'<div class="photo art"><div class="mono">{esc(initials)}</div></div>'
    hero_bg = (f"linear-gradient(90deg,{bg}F2 0%,{bg}CC 45%,{bg}33 100%),url({photo}) center/cover"
               if variant == "full" else bg)
    sub = f"{esc(tagline)}. {esc(lead.get('niche','').capitalize())} v lokalitě {esc(district or town)}."

    return f'''<!doctype html><html lang="cs"><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{width:1920px;height:1080px;font-family:Inter,"Inter Display",Arial,sans-serif;background:#D9DCE1;color:{txt};overflow:hidden}}
.chrome{{height:56px;background:#E9EBEE;display:flex;align-items:center;padding:0 22px;gap:18px;border-bottom:1px solid #D0D3D8}}
.dots span{{display:inline-block;width:13px;height:13px;border-radius:50%;margin-right:8px}}
.url{{flex:1;max-width:760px;margin:0 auto;height:34px;border-radius:17px;background:#fff;display:flex;align-items:center;
  padding:0 18px;font-size:15px;color:#3C4043;gap:10px}}
.url b{{font-weight:500}} .lock{{width:14px;height:14px;stroke:#188038;fill:none;stroke-width:2.4}}
.site{{height:1024px;background:{bg};position:relative;overflow:hidden}}
nav{{height:92px;display:flex;align-items:center;padding:0 96px;gap:48px;position:relative;z-index:3}}
nav.full{{background:{bg}EE;box-shadow:0 1px 0 {txt}14}}
.logo{{font-weight:800;font-size:26px;letter-spacing:-.5px;display:flex;align-items:center;gap:12px;white-space:nowrap}}
.logo i{{width:38px;height:38px;border-radius:11px;background:{acc};display:inline-block;box-shadow:inset -10px -10px 0 {acc2}}}
.menu{{display:flex;gap:36px;font-size:17px;font-weight:500;color:{mut};margin-left:auto}}
.navcta{{background:{acc};color:{"#111" if pal_key=="fitness" else "#fff"};padding:14px 24px;border-radius:12px;font-weight:600;font-size:17px;white-space:nowrap}}
.hero{{display:flex;align-items:center;height:620px;padding:0 96px;gap:72px;position:relative;background:{hero_bg}}}
.hero.full{{height:1024px;margin-top:-92px;padding:190px 96px 0;align-items:flex-start}}
.copy{{flex:0 0 {"860px" if variant=="full" else "820px"};position:relative;z-index:2}}
.badge{{display:inline-flex;align-items:center;gap:10px;background:{surf};border:1px solid {acc2}88;padding:10px 18px;border-radius:999px;
  font-size:16px;font-weight:600;color:{txt};box-shadow:0 6px 20px #0000000d}}
.star{{width:18px;height:18px;fill:#F5B400;stroke:#F5B400}}
h1{{font-size:{86 if len(name) <= 18 else 70 if len(name) <= 26 else 58}px;line-height:1.02;letter-spacing:-2.5px;font-weight:800;margin:30px 0 26px}}
h1 em{{font-style:normal;color:{acc}}}
.sub{{font-size:23px;line-height:1.5;color:{mut};max-width:680px}}
.btns{{display:flex;gap:16px;margin-top:40px}}
.b1{{background:{acc};color:{"#111" if pal_key=="fitness" else "#fff"};padding:22px 34px;border-radius:14px;font-size:19px;font-weight:700;box-shadow:0 14px 30px {acc}55}}
.b2{{border:2px solid {txt}22;padding:20px 30px;border-radius:14px;font-size:19px;font-weight:600;display:flex;gap:10px;align-items:center;background:{surf}AA}}
.visual{{flex:1;height:540px;position:relative}}
.photo{{position:absolute;inset:0 0 0 40px;border-radius:32px;background-size:cover;background-position:center;box-shadow:0 30px 70px #0000002e}}
.photo.art{{background:radial-gradient(circle at 30% 25%,{acc2} 0,transparent 55%),radial-gradient(circle at 80% 80%,{acc}cc 0,transparent 60%),{surf};
  display:flex;align-items:center;justify-content:flex-end;padding-right:80px}}
.mono{{font-size:220px;font-weight:800;color:{"#ffffff" if dark else "#ffffffee"};letter-spacing:-10px;text-shadow:0 20px 60px #00000033}}
.wcard{{position:absolute;{"right:0;top:30px" if variant=="full" else "left:-10px;bottom:-30px"};width:360px;background:{surf};color:{txt};border-radius:22px;padding:26px;
  box-shadow:0 24px 60px #00000033;z-index:4;border:1px solid {txt}12}}
.wt{{font-weight:700;font-size:20px}} .wsub{{color:{mut};font-size:15px;margin:6px 0 16px}}
.slots{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}}
.slot{{border:1.5px solid {txt}22;border-radius:10px;padding:10px 0;text-align:center;font-weight:600;font-size:16px}}
.slot.on{{background:{acc};border-color:{acc};color:{"#111" if pal_key=="fitness" else "#fff"}}}
.wbtn{{margin-top:18px;background:{txt};color:{bg};text-align:center;padding:15px;border-radius:12px;font-weight:700;font-size:16px}}
.mrow{{display:flex;justify-content:space-between;padding:13px 0;border-bottom:1px dashed {txt}22;font-size:17px;margin-top:4px}}
.field{{border:1.5px solid {txt}1f;border-radius:10px;padding:13px 14px;color:{mut};font-size:15px;margin-top:10px}}
.field.tall{{height:70px}}
.services{{position:absolute;left:96px;right:96px;bottom:44px;display:grid;grid-template-columns:repeat(4,1fr);gap:24px;z-index:3}}
.card{{background:{surf};border-radius:20px;padding:28px 28px 26px;border:1px solid {txt}10;box-shadow:0 10px 30px #00000010}}
.ic{{width:52px;height:52px;border-radius:14px;background:{acc}1f;display:flex;align-items:center;justify-content:center;margin-bottom:18px}}
.ic svg{{width:26px;height:26px;fill:none;stroke:{acc};stroke-width:2.2;stroke-linecap:round;stroke-linejoin:round}}
.ct{{font-size:21px;font-weight:700}} .cd{{font-size:16px;color:{mut};margin-top:6px}}
.ph svg{{width:20px;height:20px;fill:none;stroke:currentColor;stroke-width:2.2}}
</style></head><body>
<div class="chrome"><div class="dots"><span style="background:#FF5F57"></span><span style="background:#FEBC2E"></span><span style="background:#28C840"></span></div>
  <div class="url"><svg class="lock" viewBox="0 0 24 24"><rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/></svg>
  <b>{esc(domain_for(lead))}</b></div><div style="width:80px"></div></div>
<div class="site">
  <nav class="{variant}"><div class="logo"><i></i>{esc(name)}</div><div class="menu">{menu}</div><div class="navcta">{esc(phone) or esc(cta)}</div></nav>
  <section class="hero {variant}">
    <div class="copy">{badge}<h1>{esc(name)}</h1><p class="sub">{sub}</p>
      <div class="btns"><div class="b1">{esc(cta)}</div>
      <div class="b2 ph"><svg viewBox="0 0 24 24"><path d="M5 4h4l2 5-2.5 1.5a11 11 0 0 0 5 5L15 13l5 2v4a2 2 0 0 1-2 2A16 16 0 0 1 3 6a2 2 0 0 1 2-2"/></svg>{esc(phone) or "Zavolejte nám"}</div></div></div>
    <div class="visual">{visual if variant == "split" else ""}{widget(wkind, PALETTES[pal_key], cta, lead)}</div>
  </section>
  <div class="services">{cards}</div>
</div></body></html>'''


# ---------------------------------------------------------------- лист
PROBLEM_CZ = {
    "БЕЗ САЙТУ": "Všiml jsem si ale, že nemáte vlastní webové stránky – zákazníci, kteří vás hledají na Googlu, "
                 "tak často skončí u konkurence.",
    "SOCIAL": "Všiml jsem si ale, že místo webu máte jen profil na sociální síti – na Googlu vás tak "
              "zákazníci najdou mnohem hůř.",
    "САЙТ НЕ ПРАЦЮЄ": "Zkoušel jsem otevřít váš web, ale nefunguje – zákazníci místo vaší nabídky vidí chybu.",
}
REASON_CZ = [
    ("телефон", "na mobilu se špatně zobrazuje (a většina zákazníků dnes hledá z telefonu)"),
    ("HTTPS", "prohlížeč u něj ukazuje „Nezabezpečeno“"),
    ("SSL", "prohlížeč hlásí problém s certifikátem"),
    ("©", "působí, že nebyl delší dobu aktualizován"),
    ("конструктор", "běží na bezplatném konstruktoru s cizí doménou"),
    ("CMS", "běží na zastaralém systému"),
    ("jQuery", "používá zastaralé technologie"),
    ("2000-х", "používá zastaralé technologie"),
    ("повільний", "načítá se pomalu"),
    ("порожній", "obsahuje velmi málo informací"),
]
PACKAGE_CZ = {
    "booking": "moderní web s online rezervací, ceníkem a galerií",
    "menu": "moderní web s menu, rezervací stolu a galerií",
    "quote": "moderní web s přehledem služeb, galerií prací a poptávkovým formulářem",
}


def email_text(lead, cfg):
    _, _, _, _, wkind = NICHES.get(lead.get("niche"), FALLBACK)
    name = short_name(lead["name"])
    status = lead["status"]
    reasons = lead.get("reasons") or ""
    if status == "БЕЗ САЙТУ":
        problem = PROBLEM_CZ["SOCIAL" if "лише" in reasons else "БЕЗ САЙТУ"]
    elif status == "САЙТ НЕ ПРАЦЮЄ":
        problem = PROBLEM_CZ[status]
    else:
        found = list(dict.fromkeys(cz for key, cz in REASON_CZ if key in reasons))[:2]
        problem = ("Prohlížel jsem si váš web a všiml jsem si, že " + " a ".join(found) + "."
                   if found else "Prohlížel jsem si váš web a myslím, že by si zasloužil modernější podobu.")
    reviews = int(lead.get("reviews") or 0)
    praise = f" Máte tam {reviews} hodnocení – to je skvělá vizitka." if reviews >= 10 else ""
    price = f"{int(lead['price']):,}".replace(",", " ")
    support = int(lead.get("support") or 0)
    web = f" | {cfg['sender_web']}" if cfg.get("sender_web") else ""
    dom = domain_for(lead)
    own = lead.get("website") and dom in (lead.get("website") or "")
    domain_line = (f"web na vaší doméně {dom} se zabezpečením HTTPS" if own
                   else f"vlastní doménu {dom} a zabezpečení HTTPS")
    subject = f"Návrh nového webu pro {name}"
    body = f"""Dobrý den,

našel jsem {name} na Firmy.cz.{praise} {problem}

Připravil jsem pro vás ukázku, jak by mohl vypadat váš nový web – najdete ji v příloze.

Co dostanete:
• {PACKAGE_CZ[wkind]}
• perfektní zobrazení na mobilu
• {domain_line}
• základní nastavení pro Google, aby vás zákazníci našli

Cena: {price} Kč jednorázově, hotovo do {cfg['delivery_days']} dnů.
Správa a drobné úpravy potom za {support} Kč měsíčně (nepovinné).

Pokud vás to zaujme, stačí odpovědět na tento e-mail nebo zavolat na {cfg['sender_phone']}.

S pozdravem
{cfg['sender_name']}
{cfg['sender_phone']} | {cfg['sender_email']}{web}

P.S. Pokud o web nemáte zájem, stačí odepsat „ne“ a už se vám nebudu ozývat."""
    wa = (f"Dobrý den, tady {cfg['sender_name']}. Připravil jsem ukázku nového webu pro {name} "
          f"(posílám obrázek). Kompletní web za {price} Kč, hotovo do {cfg['delivery_days']} dnů. "
          f"Mám vám poslat podrobnosti?")
    return subject, body, wa


# ---------------------------------------------------------------- головне
def write_offers_xlsx(rows, path):
    cols = [("Пріоритет", "status", 15), ("Назва", "name", 32), ("Email", "email", 28),
            ("Телефон", "phone", 17), ("Ціна, Kč", "price", 10), ("Тема листа", "subject", 32),
            ("Текст листа", "body", 70), ("WhatsApp / SMS", "wa", 45), ("Картинка", "image_file", 30),
            ("Відправлено", "sent", 12)]
    wb = Workbook()
    ws = wb.active
    ws.title = "Пропозиції"
    ws.append([c[0] for c in cols])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="305496")
    for r in rows:
        ws.append([r.get(k, "") for _, k, _ in cols])
        row = ws.max_row
        for col in range(1, len(cols) + 1):
            ws.cell(row, col).alignment = Alignment(vertical="top", wrap_text=col in (7, 8))
        ws.cell(row, 9).hyperlink = r["image_file"]
        ws.cell(row, 9).font = Font(color="0563C1", underline="single")
        ws.row_dimensions[row].height = 120
    for i, (_, _, w) in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def load_config():
    if not CONFIG.exists():
        CONFIG.write_text(json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=2), "utf8")
        print(f"! Створив {CONFIG.name} – впиши туди своє ім'я, телефон, email.")
    return {**DEFAULT_CONFIG, **json.loads(CONFIG.read_text("utf8"))}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("leads_xlsx")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--out", default=str(HERE / "offers"))
    ap.add_argument("--no-photos", action="store_true", help="не тягнути фото з Firmy.cz")
    args = ap.parse_args()

    from playwright.sync_api import sync_playwright

    cfg = load_config()
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    leads = load(args.leads_xlsx, "Ліди")
    if args.limit:
        leads = leads[:args.limit]
    print(f"{len(leads)} лідів → {out}")

    rows = []
    with sync_playwright() as p:
        exe = os.environ.get("CHROMIUM_PATH")
        browser = p.chromium.launch(executable_path=exe) if exe else p.chromium.launch()
        page = browser.new_page(viewport={"width": 1920, "height": 1080})
        for i, lead in enumerate(leads, 1):
            fname = out / f"{i:03d}_{slugify(lead['name'])}.jpg"
            if not fname.exists():
                photo = None if args.no_photos else fetch_photo(lead)
                page.set_content(render_html(lead, photo), wait_until="networkidle", timeout=60000)
                page.screenshot(path=str(fname), type="jpeg", quality=88)
            subject, body, wa = email_text(lead, cfg)
            rows.append({**lead, "subject": subject, "body": body, "wa": wa,
                         "image_file": str(fname.relative_to(out))})
            if i % 20 == 0:
                print(f"   {i}/{len(leads)}")
        browser.close()

    write_offers_xlsx(rows, out / "nabidky.xlsx")
    print(f"Готово: {out / 'nabidky.xlsx'} + {len(rows)} картинок")


if __name__ == "__main__":
    main()
