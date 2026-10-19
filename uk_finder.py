"""
Гарячі ліди у Великій Британії: Thomson Local -> аудит сайту -> email + перевірка Ltd -> JSON/Excel.

Гарячий лід = живий бізнес (відгуки) зі старим / слабким / мертвим сайтом.
Канал:  email   – компанія Ltd і є адреса (B2B email корпоративним абонентам дозволений PECR);
        call    – решта (sole trader, без email): тільки дзвінок, email/SMS без згоди не можна.

Запуск:  python uk_finder.py --top 200
"""
import argparse
import json
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import urljoin, urlparse

import lead_finder as lf

NICHES = ["plumbers", "electricians", "builders", "roofers", "landscape-gardeners", "removals",
          "cleaning-services", "hairdressers", "beauty-salons", "barbers", "dentists", "restaurants",
          "driving-schools", "accountants", "painters-decorators", "carpet-cleaners", "car-repairs",
          "florists", "photographers", "tree-surgeons", "dog-grooming", "fencing-contractors", "locksmiths"]
CITIES = ["london", "manchester", "birmingham", "leeds", "liverpool", "bristol", "sheffield",
          "nottingham", "leicester", "glasgow", "edinburgh", "cardiff", "newcastle-upon-tyne"]

LTD_RE = re.compile(r"\b(ltd|limited)\b|company\s*(no|number|reg)|registered in (england|scotland|wales)", re.I)
CHAINS = ("quooker", "dyno", "pimlico", "british gas", "anytime", "timpson", "toni&guy", "rush hair", "supercuts")


def tl_search(niche, city, page):
    url = f"https://www.thomsonlocal.com/search/{niche}/{city}" + (f"?page={page}" if page > 1 else "")
    try:
        h = lf.firmy_get(url).text  # той самий ввічливий rate-limit
    except Exception:
        return []
    out = []
    for it in h.split('itemtype="http://schema.org/LocalBusiness"')[1:]:
        g = lambda rx: (m.group(1).strip() if (m := re.search(rx, it)) else "")
        name = lf.htmllib.unescape(g(r'itemprop="name"[^>]*>([^<]+)'))
        if not name:
            continue
        out.append({
            "id": g(r'data-pvd-c="(\d+)"') or name,
            "name": name,
            "phone": g(r'href="tel:([^"]+)"'),
            "website": g(r'itemprop="sameAs" href="([^"]+)"'),
            "street": g(r'itemprop="streetAddress"[^>]*>([^<]+)'),
            "city": g(r'itemprop="addressLocality"[^>]*>([^<]+)') or city.replace("-", " ").title(),
            "postcode": g(r'itemprop="postalCode"[^>]*>([^<]+)'),
            "reviews": int(g(r"Reviews \((\d+)\)") or 0),
            "stars": int(g(r"stars stars(\d)") or 0),
            "niche": niche, "search_city": city,
            "tl_url": "https://www.thomsonlocal.com" + g(r'href="(/search/[^"#]+/\d+/[^"#]+)"'),
        })
    return out


def uk_phone(p):
    d = re.sub(r"\D", "", p or "")
    if d.startswith("44"):
        d = "0" + d[2:]
    kind = "mobile" if d.startswith("07") else "landline"
    pretty = f"{d[:5]} {d[5:]}" if d.startswith("07") or len(d) != 11 else f"{d[:4]} {d[4:7]} {d[7:]}"
    return pretty, "44" + d[1:], kind


def site_text_and_emails(url):
    """Текст головної + контактних сторінок і email-и (з пріоритетом домену сайту)."""
    if not url.startswith("http"):
        url = "http://" + url
    host = urlparse(url).netloc.replace("www.", "")
    texts, emails = [], []
    for path in ("", "/contact", "/contact-us", "/about", "/about-us"):
        try:
            r = lf.get(urljoin(url, path) if path else url, tries=1, timeout=12, verify=False)
        except Exception:
            continue
        if r.status_code != 200:
            continue
        texts.append(r.text)
        emails += [lf.htmllib.unescape(e).lower() for e in re.findall(r'mailto:([^"\'?>\s]+)', r.text) + lf.EMAIL_RE.findall(r.text)
                   if lf.good_email(e)]
        if emails and path:
            break
    emails = list(dict.fromkeys(emails))
    emails.sort(key=lambda e: (host not in e, len(e)))
    return " ".join(texts), emails


def process(c):
    status, score, reasons, _ = lf.audit_site(c["website"])
    c["status"], c["score"], c["reasons"] = status, score, reasons
    c["email"], c["ltd"] = "", bool(LTD_RE.search(c["name"]))
    if c["website"] and status in ("СТАРИЙ САЙТ", "СЛАБКИЙ САЙТ"):
        text, emails = site_text_and_emails(c["website"])
        c["email"] = emails[0] if emails else ""
        c["ltd"] = c["ltd"] or bool(LTD_RE.search(re.sub(r"<[^>]+>", " ", text)))
    c["phone"], c["wa"], c["phone_type"] = uk_phone(c["phone"])
    c["channel"] = "email" if c["email"] and c["ltd"] else "call"
    # гарячість: проблема сайту + живість бізнесу + легальний email
    c["hot"] = score + min(c["reviews"], 40) + c["stars"] * 3 + (25 if c["channel"] == "email" else 0)
    return c


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--top", type=int, default=200)
    ap.add_argument("--pages", type=int, default=1)
    ap.add_argument("--out", default="output/uk_leads.json")
    args = ap.parse_args()

    cards = {}
    for city in CITIES:
        for niche in NICHES:
            for page in range(1, args.pages + 1):
                for c in tl_search(niche, city, page):
                    cards.setdefault(c["id"], c)
        print(f"[пошук] {city}: усього {len(cards)}", flush=True)

    with_site = [c for c in cards.values() if c["website"] and not any(x in c["name"].lower() for x in CHAINS)]
    print(f"[аудит] {len(with_site)} сайтів …", flush=True)
    done = []
    with ThreadPoolExecutor(10) as ex:
        for i, f in enumerate(as_completed([ex.submit(process, c) for c in with_site]), 1):
            try:
                done.append(f.result())
            except Exception:
                pass
            if i % 200 == 0:
                print(f"   {i}/{len(with_site)}", flush=True)

    # дублікати (одна фірма в кількох нішах/містах) – за доменом
    uniq, seen = [], set()
    for c in sorted(done, key=lambda c: -c["hot"]):
        k = urlparse(c["website"] if c["website"].startswith("http") else "http://" + c["website"]).netloc.replace("www.", "")
        if k in seen:
            continue
        seen.add(k)
        uniq.append(c)
    Path("output").mkdir(exist_ok=True)
    Path("output/uk_all.json").write_text(json.dumps(uniq, ensure_ascii=False), "utf8")  # усі перевірені, для добору
    hot = [c for c in uniq if c["status"] in ("СТАРИЙ САЙТ", "СЛАБКИЙ САЙТ", "САЙТ НЕ ПРАЦЮЄ") and c["phone"]]
    hot.sort(key=lambda c: (c["channel"] != "email", -c["hot"]))
    top = hot[:args.top]
    Path(args.out).parent.mkdir(exist_ok=True)
    Path(args.out).write_text(json.dumps(top, ensure_ascii=False, indent=1), "utf8")
    stat = lambda k, v: sum(1 for c in top if c[k] == v)
    print(f"\nГотово: {args.out} — {len(top)} гарячих (з {len(uniq)} сайтів)")
    print(f"  email (Ltd): {stat('channel', 'email')}, дзвінок: {stat('channel', 'call')}")
    for s in ("САЙТ НЕ ПРАЦЮЄ", "СТАРИЙ САЙТ", "СЛАБКИЙ САЙТ"):
        print(f"  {s}: {stat('status', s)}")


if __name__ == "__main__":
    main()
