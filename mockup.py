"""
Унікальний макет сайту 1920x1080 для кожного ліда.

Кожна фірма отримує свою комбінацію: арт-напрям (6 розкладок) × палітра × пара шрифтів ×
тематичні фото (Unsplash, вільна ліцензія) × слоган. Вибір детермінований від назви фірми,
тож повторний рендер дає той самий результат.
"""
import hashlib
import html
import re
from pathlib import Path
from string import Template

import requests

PHOTO_DIR = Path(__file__).resolve().parent / "assets" / "photos"

# ---------------------------------------------------------------- фото по нішах (Unsplash id)
P = {
    "hair": "1560066984-138dadb4c035 1522337360788-8b13dee7a37e 1562322140-8baeececf3df 1521590832167-7bcbfaa6381f "
            "1633681926022-84c23e8cb2d6 1600948836101-f9ffda59d250 1559599101-f09722fb4948 1580618672591-eb180b1a973f "
            "1595476108010-b4d1f102b1b1 1519699047748-de8e457a634e",
    "barber": "1503951914875-452162b0f3f1 1585747860715-2ba37e788b70 1599351431202-1e0f0137899a 1621605815971-fbc98d665033 "
              "1622286342621-4bd786c2447c 1512690459411-b9245aed614b 1593702295094-aea22597af65 1534297635766-a262cdcb8ee4 "
              "1517832606299-7ae9b720a186",
    "cosmetic": "1570172619644-dfd03ed5d881 1616394584738-fc6e612e71b9 1596462502278-27bfdc403348 1522335789203-aabd1fc54bc9 "
                "1487412947147-5cebf100ffc2",
    "nails": "1604654894610-df63bc536371 1610992015732-2449b76344bc 1519014816548-bf5fe059798b 1607779097040-26e80aa78e66",
    "spa": "1544161515-4ab6ce6db874 1600334089648-b0d9d3028eb2 1540555700478-4be289fbecef 1515377905703-c4788e51af15 "
           "1519823551278-64ac92734fb1 1591343395082-e120087004b4 1600334129128-685c5582fd35",
    "auto": "1486262715619-67b85e0b08d3 1487754180451-c456f719a1fc 1625047509248-ec889cbff17f 1619642751034-765dfdf7c58e "
            "1530046339160-ce3e530c7d2f 1632823469850-2f77dd9c7f93 1492144534655-ae79c964c9d7",
    "cars": "1503376780353-7e6692767b70 1580273916550-e323be2ae537 1492144534655-ae79c964c9d7 1494976388531-d1058494cdd8 "
            "1625047509248-ec889cbff17f",
    "drive": "1449965408869-eaa3f722e40d 1485463611174-f302f6a5c1c9 1503376780353-7e6692767b70",
    "dental": "1606811971618-4486d14f3f99 1629909613654-28e377c37b09 1588776814546-1ffcf47267a5 1609840114035-3c981b782dfe "
              "1598256989800-fe5f95da9787",
    "vet": "1548199973-03cce0bbc87b 1450778869180-41d0601e046e 1587300003388-59208cc962cb 1583337130417-3346a1be7dee "
           "1628009368231-7bb7cfcb0def 1543466835-00a7907e9de1 1514888286974-6c03e2ca1dba",
    "food": "1517248135467-4c7edcad34c4 1414235077428-338989a2e8c0 1555396273-367ea4eb4db5 1504674900247-0877df9cc836 "
            "1559339352-11d035aa65de 1424847651672-bf20a4b0982b 1550966871-3ed3cdb5ed0c 1592861956120-e524fc739696 "
            "1466978913421-dad2ebd01d17 1540189549336-e6e99c3679fe",
    "cafe": "1495474472287-4d71bcdd2085 1501339847302-ac426a4a7cbb 1554118811-1e0d58224f24 1511920170033-f8396924c348 "
            "1445116572660-236099ec97a0 1559925393-8be0ec4767c8 1442512595331-e89e73853f31 1509042239860-f550ce710b93 "
            "1521017432531-fbd92d768814",
    "bakery": "1509440159596-0249088772ff 1486427944299-d1955d23e34d 1555507036-ab1f4038808a 1517433670267-08bbd4be890f "
              "1608198093002-ad4e005484ec 1568254183919-78a4f43a2877 1549931319-a545dcf3bc73 1612240498936-65f5101365d2",
    "flowers": "1487530811176-3780de880c2d 1490750967868-88aa4486c946 1455659817273-f96807779a8a 1561181286-d3fee7d55364 "
               "1563241527-3004b7be0ffd 1526047932273-341f2a7631f9 1470509037663-253afd7f0f51 1519378058457-4c29a0a2efac",
    "photo": "1452587925148-ce544e77e70d 1502982720700-bfff97f2ecac 1516035069371-29a1b244cc32 1519741497674-611481863552 "
             "1511285560929-80b456fea0bc 1554080353-a576cf803bda 1471341971476-ae15ff5dd4ea",
    "plumb": "1585704032915-c3400ca199e7 1607472586893-edb57bdc0e39 1620626011761-996317b8d101 1584622650111-993a426fbf0a "
             "1600566752355-35792bedcfea 1604709177225-055f99402ea3",
    "electric": "1621905251189-08b45d6a269e 1621905252507-b35492cc74b4 1544724569-5f546fd6f2b5 1592833159155-c62df1b65634",
    "interior": "1556228453-efd6c1ff04f6 1595428774223-ef52624120d2 1611269154421-4e27233ac5c7 1600585152220-90363fe7e115 "
                "1586023492125-27b2c045efd7 1618221195710-dd6b41faaea6",
    "carpentry": "1589939705384-5185137a7f0f 1504148455328-c376907d081c 1600585152220-90363fe7e115 1595428774223-ef52624120d2 "
                 "1556228453-efd6c1ff04f6",
    "paint": "1562259949-e8e7689d7828 1595814433015-e6f5ce69614e 1618221195710-dd6b41faaea6 1513519245088-0e12902e5a38 "
             "1586023492125-27b2c045efd7",
    "build": "1504307651254-35680f356dfd 1541888946425-d81bb19240f5 1503387762-592deb58ef4e 1600585154340-be6161a56a0c "
             "1564013799919-ab600027ffc6 1589939705384-5185137a7f0f",
    "realty": "1560518883-ce09059eeffa 1600596542815-ffad4c1539a9 1600585154340-be6161a56a0c 1512917774080-9991f1c4c750 "
              "1600607687939-ce8a6c25118c 1502672260266-1c1ef2d93688 1522708323590-d24dbb6b0267 1493809842364-78817add7ffb",
    "office": "1554224155-6726b3ff858f 1450101499163-c8848c66ca85 1497366216548-37526070297c 1497366811353-6870744d04b2 "
              "1600880292203-757bb62b4baf 1460925895917-afdab827c52f",
    "fitness": "1534438327276-14e5300c3a48 1571019613454-1cb2f99b2d8b 1517836357463-d25dfeac3438 1540497077202-7c8a3999166f "
               "1581009146145-b5ef050c2e1e 1574680096145-d05b474e2155 1518611012118-696072aa579a",
    "clean": "1581578731548-c64695cc6952 1527515637462-cff94eecc1ac 1563453392212-326f5e854473 1585421514738-01798e348b17 "
             "1600585154084-4e5fe7c39198",
    "lock": "1558002038-1055907df827 1582139329536-e7284fece509 1600585154084-4e5fe7c39198",
}
P.update({
    "garden": "1558904541-efa843a96f01 1416879595882-3373a0480b5b 1585320806297-9794b3e4eeae 1523348837708-15d4a09cfac2 "
              "1557429287-b2e26467fc2b 1611843467160-25afb8df1074",
    "trees": "1542273917363-3b1817f69a2d 1513836279014-a89f7a76ae86 1502082553048-f009c37129b9 1473773508845-188df298d2d1",
    "removals": "1600518464441-9154a4dea21b 1520038410233-7141be7e6f97 1600585152220-90363fe7e115 1586023492125-27b2c045efd7",
    "fence": "1558036117-15d82a90b9b1 1557429287-b2e26467fc2b 1585320806297-9794b3e4eeae 1558904541-efa843a96f01",
})
P = {k: v.split() for k, v in P.items()}

# ---------------------------------------------------------------- ніші: фото, настрій, слогани
# mood -> які розкладки й палітри підходять
NICHE = {
    "kadeřnictví":      ("hair", "beauty", ["Vlasy, které <em>vám sluší</em>", "Krása začíná <em>u nás</em>", "Váš styl, <em>naše péče</em>"]),
    "kosmetický salon": ("cosmetic", "beauty", ["Pleť, která <em>září</em>", "Čas jen <em>pro vás</em>", "Přirozená krása, <em>skutečná péče</em>"]),
    "manikúra":         ("nails", "beauty", ["Nehty, které <em>zaujmou</em>", "Detaily dělají <em>krásu</em>", "Péče o ruce <em>s láskou</em>"]),
    "masáže":           ("spa", "spa", ["Zpomalte. <em>Nadechněte se.</em>", "Odpočinek, který <em>cítíte</em>", "Uvolnění pro <em>tělo i mysl</em>"]),
    "barbershop":       ("barber", "barber", ["Střih. Vousy. <em>Styl.</em>", "Gentleman <em>začíná tady</em>", "Ostrý střih, <em>čistá linie</em>"]),
    "autoservis":       ("auto", "auto", ["Vaše auto v <em>dobrých rukou</em>", "Servis bez <em>kompromisů</em>", "Opravíme, <em>co jiní nezvládnou</em>"]),
    "pneuservis":       ("cars", "auto", ["Přezujeme vás <em>bez čekání</em>", "Bezpečně <em>v každé sezóně</em>", "Kola v <em>perfektní kondici</em>"]),
    "autoškola":        ("drive", "auto", ["Řidičák <em>s klidem</em>", "Za volant <em>s jistotou</em>", "Naučíme vás <em>jezdit</em>"]),
    "zubní ordinace":   ("dental", "medical", ["Úsměv, který <em>vydrží</em>", "Zdravé zuby <em>bez strachu</em>", "Péče o váš <em>úsměv</em>"]),
    "veterinář":        ("vet", "medical", ["Péče o vaše <em>nejbližší</em>", "Zdravý mazlíček, <em>šťastný domov</em>", "S láskou <em>ke zvířatům</em>"]),
    "restaurace":       ("food", "food", ["Chuť, která <em>se vrací</em>", "Poctivé jídlo, <em>skvělá společnost</em>", "Stůl pro vás je <em>připraven</em>"]),
    "kavárna":          ("cafe", "cafe", ["Káva, kvůli které <em>se vrátíte</em>", "Místo pro vaše <em>chvíle</em>", "Výběrová káva <em>s duší</em>"]),
    "pekárna":          ("bakery", "cafe", ["Voní to <em>domovem</em>", "Pečeno <em>každé ráno</em>", "Poctivé pečivo <em>s láskou</em>"]),
    "květinářství":     ("flowers", "flowers", ["Květiny, které <em>mluví za vás</em>", "Krása <em>v každé kytici</em>", "Radost <em>v květech</em>"]),
    "fotograf":         ("photo", "photo", ["Okamžiky, které <em>zůstanou</em>", "Příběhy <em>ve fotkách</em>", "Světlo. Emoce. <em>Vzpomínky.</em>"]),
    "instalatér":       ("plumb", "crafts", ["Voda, topení, plyn <em>bez starostí</em>", "Opravíme to <em>rychle a čistě</em>", "Koupelna <em>vašich snů</em>"]),
    "elektrikář":       ("electric", "crafts", ["Elektřina <em>v bezpečí</em>", "Spolehlivě <em>pod proudem</em>", "Elektro práce <em>bez starostí</em>"]),
    "truhlářství":      ("carpentry", "crafts", ["Nábytek <em>na míru</em>", "Ze dřeva, <em>s citem</em>", "Poctivé řemeslo <em>pro váš domov</em>"]),
    "malíř pokojů":     ("paint", "crafts", ["Nové barvy <em>vašeho domova</em>", "Malujeme <em>čistě a rychle</em>", "Stěny, které <em>dýchají</em>"]),
    "stavební firma":   ("build", "pro", ["Stavíme <em>na jistotu</em>", "Od základů <em>po klíč</em>", "Kvalita, která <em>vydrží</em>"]),
    "účetní":           ("office", "pro", ["Účetnictví <em>bez stresu</em>", "Čísla <em>v pořádku</em>", "Vaše finance <em>v dobrých rukou</em>"]),
    "fitness":          ("fitness", "fitness", ["Silnější <em>každý den</em>", "Tvoje forma <em>začíná tady</em>", "Překonej <em>sám sebe</em>"]),
    "realitní kancelář": ("realty", "realty", ["Domov, <em>který hledáte</em>", "Prodáme <em>za nejlepší cenu</em>", "Nemovitosti <em>s jistotou</em>"]),
    "úklidové služby":  ("clean", "medical", ["Čisto <em>bez námahy</em>", "Uklidíme, <em>vy odpočívejte</em>", "Zářivě čistý <em>prostor</em>"]),
    "zámečnictví":      ("lock", "auto", ["Otevřeme <em>každé dveře</em>", "Bezpečí <em>vašeho domova</em>", "Zámečník <em>nonstop</em>"]),
}
FALLBACK_NICHE = ("office", "pro", ["Profesionálně <em>a spolehlivě</em>", "Služby <em>s jistotou</em>", "Rádi vám <em>pomůžeme</em>"])

STYLES_BY_MOOD = {
    "beauty": ["arch", "editorial", "magazine"], "spa": ["arch", "editorial", "split"],
    "barber": ["bold", "editorial", "split"], "auto": ["bold", "split", "corporate"],
    "medical": ["split", "corporate", "arch"], "food": ["editorial", "magazine", "arch"],
    "cafe": ["magazine", "arch", "editorial"], "flowers": ["arch", "magazine", "editorial"],
    "crafts": ["split", "bold", "magazine"], "pro": ["corporate", "split", "editorial"],
    "fitness": ["bold", "editorial", "split"], "photo": ["editorial", "magazine", "bold"],
    "realty": ["corporate", "editorial", "split"],
}

# bg, surface, ink, muted, accent, on_accent, line
PALETTES = {
    "beauty":  [("#F7F1EC", "#FFFFFF", "#2B1F22", "#86737A", "#B0697A", "#FFFFFF", "#E7D8D2"),
                ("#F4EEE7", "#FFFDFB", "#1F1A17", "#7C6F66", "#A8885A", "#FFFFFF", "#E2D6C6"),
                ("#EEF0EA", "#FFFFFF", "#22281F", "#6E7768", "#7D8C6A", "#FFFFFF", "#D6DCCE"),
                ("#14110F", "#1E1A17", "#F4ECE4", "#A99C90", "#C9A27A", "#14110F", "#3A322B")],
    "spa":     [("#F1EEE7", "#FFFFFF", "#23302A", "#6C7A72", "#5E8B6F", "#FFFFFF", "#DAD7CC"),
                ("#F5F0EA", "#FFFFFF", "#2E2622", "#857870", "#B08968", "#FFFFFF", "#E5DBD1")],
    "barber":  [("#0F0F10", "#19191B", "#F4F1EA", "#A19D96", "#C9A063", "#0F0F10", "#2C2B29"),
                ("#121619", "#1B2125", "#EEF1F2", "#97A1A6", "#D0753F", "#FFFFFF", "#2A3136"),
                ("#151311", "#201D1A", "#F2EBE1", "#A69A8B", "#B5462E", "#FFFFFF", "#332E29")],
    "auto":    [("#0D1117", "#161C24", "#F1F5F9", "#94A3B8", "#F05A28", "#FFFFFF", "#253041"),
                ("#0E0E0E", "#181818", "#F5F5F5", "#9A9A9A", "#E5322D", "#FFFFFF", "#2A2A2A"),
                ("#111418", "#1A1F25", "#F2F4F7", "#98A2B3", "#F5B700", "#111418", "#2A313B")],
    "medical": [("#F4F8FA", "#FFFFFF", "#0F2A3A", "#5C7382", "#0E8F96", "#FFFFFF", "#DCE7EC"),
                ("#F5F7FB", "#FFFFFF", "#101E3C", "#5D6B86", "#2F6FEB", "#FFFFFF", "#DFE5F0"),
                ("#F3F8F5", "#FFFFFF", "#12291E", "#5E7568", "#2E9A6A", "#FFFFFF", "#D9E7DF")],
    "food":    [("#F6EFE4", "#FFFDF8", "#2A1C13", "#7B6655", "#B4532A", "#FFFFFF", "#E6D8C4"),
                ("#14201A", "#1C2A23", "#F3EEE3", "#A9B0A2", "#D4A64A", "#14201A", "#2C3B33"),
                ("#120F0D", "#1C1814", "#F5EDE3", "#A89A8A", "#D98A3A", "#120F0D", "#332B24")],
    "cafe":    [("#F3ECE3", "#FFFCF8", "#2B2019", "#7E6B5D", "#8B5A3C", "#FFFFFF", "#E2D5C6"),
                ("#F1EFE6", "#FFFFFF", "#232619", "#717563", "#6B7A3A", "#FFFFFF", "#DEDCCD"),
                ("#1A1411", "#241C17", "#F3E9DE", "#AC9C8C", "#C98B4E", "#1A1411", "#3A2F27")],
    "flowers": [("#FBF3F4", "#FFFFFF", "#2D1F2A", "#7D6B77", "#C2477A", "#FFFFFF", "#F0DDE2"),
                ("#F4F3EC", "#FFFFFF", "#202A20", "#6B7768", "#6E8F5E", "#FFFFFF", "#DCE0D2"),
                ("#F5F0F7", "#FFFFFF", "#271D2E", "#776B80", "#7F4E9C", "#FFFFFF", "#E3D9EA")],
    "crafts":  [("#F4F1EA", "#FFFFFF", "#1F2A24", "#66726B", "#2F6B4F", "#FFFFFF", "#DED9CC"),
                ("#F3EEE6", "#FFFFFF", "#2A211B", "#7C6E62", "#B3572E", "#FFFFFF", "#E2D8CA"),
                ("#15181B", "#1E2226", "#F0F2F3", "#9AA3AA", "#F2A541", "#15181B", "#2D3338")],
    "pro":     [("#F5F7FB", "#FFFFFF", "#0E1B33", "#5B6782", "#1F4FD1", "#FFFFFF", "#DFE4EE"),
                ("#F4F6F3", "#FFFFFF", "#13261F", "#5D6F66", "#1E7A55", "#FFFFFF", "#DCE3DD"),
                ("#F6F5F2", "#FFFFFF", "#1C1B2E", "#6A6880", "#5B4BDB", "#FFFFFF", "#E3E1EA")],
    "fitness": [("#0B0C0E", "#16181C", "#F5F7F2", "#9EA39A", "#A3E635", "#0B0C0E", "#262A2F"),
                ("#0C0C0C", "#171717", "#F7F4F0", "#A09A93", "#FF6B1A", "#0C0C0C", "#2A2A2A"),
                ("#0B0E14", "#151A23", "#F1F4FA", "#97A0B3", "#3B82F6", "#FFFFFF", "#252C3A")],
    "photo":   [("#0E0E0E", "#181818", "#F2F2F2", "#9B9B9B", "#E8E1D5", "#0E0E0E", "#2A2A2A"),
                ("#F2EFEA", "#FFFFFF", "#151515", "#6F6B66", "#151515", "#FFFFFF", "#DDD8D0"),
                ("#191716", "#23201E", "#F1ECE6", "#A39C94", "#C66B4A", "#FFFFFF", "#38332F")],
    "realty":  [("#F5F6F8", "#FFFFFF", "#0F1A2B", "#5E6878", "#B08D57", "#FFFFFF", "#E1E4E9"),
                ("#0F1114", "#191C21", "#F3F1EC", "#A3A09A", "#C8A96A", "#0F1114", "#2B2F36"),
                ("#F3F4EF", "#FFFFFF", "#18261E", "#647069", "#2F5D46", "#FFFFFF", "#DDE1D8")],
}

FONTS = {  # (display, body, display_weight, uppercase)
    "editorial": [("Playfair Display", "Inter", 500, False), ("Cormorant Garamond", "Manrope", 500, False),
                  ("DM Serif Display", "Outfit", 400, False)],
    "arch":      [("Cormorant Garamond", "Manrope", 500, False), ("Fraunces", "Inter", 400, False),
                  ("Playfair Display", "Outfit", 500, False)],
    "bold":      [("Bebas Neue", "Inter", 400, True), ("Anton", "Manrope", 400, True), ("Oswald", "Inter", 600, True)],
    "split":     [("Space Grotesk", "Inter", 600, False), ("Manrope", "Manrope", 700, False),
                  ("Fraunces", "Inter", 500, False)],
    "magazine":  [("Fraunces", "Inter", 400, False), ("DM Serif Display", "Manrope", 400, False),
                  ("Syne", "Inter", 700, False)],
    "corporate": [("Manrope", "Manrope", 800, False), ("Outfit", "Inter", 700, False),
                  ("Space Grotesk", "Inter", 700, False)],
}
GFONTS = ("https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Manrope:wght@400;500;600;700;800"
          "&family=Outfit:wght@400;500;600;700&family=Playfair+Display:ital,wght@0,400;0,500;0,600;1,400;1,500"
          "&family=Cormorant+Garamond:ital,wght@0,500;0,600;1,500&family=DM+Serif+Display:ital@0;1"
          "&family=Fraunces:ital,opsz,wght@0,9..144,400;0,9..144,500;1,9..144,400&family=Bebas+Neue&family=Anton"
          "&family=Oswald:wght@500;600&family=Space+Grotesk:wght@500;600;700&family=Syne:wght@600;700;800&display=block")

ICON = {
    "phone": '<path d="M5 4h4l2 5-2.5 1.5a11 11 0 0 0 5 5L15 13l5 2v4a2 2 0 0 1-2 2A16 16 0 0 1 3 6a2 2 0 0 1 2-2"/>',
    "pin": '<path d="M12 21s-7-6.2-7-11.5A7 7 0 0 1 19 9.5C19 14.8 12 21 12 21z"/><circle cx="12" cy="9.5" r="2.5"/>',
    "star": '<path d="m12 3 2.7 5.6 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.9 1-6.1-4.4-4.3 6.1-.9z"/>',
    "arrow": '<path d="M5 12h14M13 6l6 6-6 6"/>',
    "check": '<path d="M20 6 9 17l-5-5"/>',
    "cal": '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 10h18"/>',
    "clock": '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/>',
}


UI_CS = {
    "menu1": ["Služby", "Ceník", "O nás", "Kontakt"], "menu2": ["Služby", "Ceník", "Galerie", "Kontakt"],
    "menu3": ["Služby", "Reference", "O nás", "Kontakt"], "reviews": "hodnocení", "call_us": "Zavolejte nám", "call": "Zavolejte",
    "sub_ed": "{head}. {niche} v lokalitě {place}.", "sub_arch": "{tagline}. {niche} v lokalitě {place} – objednejte se online během pár vteřin.",
    "sub_line": "Osobní přístup, férové ceny a rychlé termíny v lokalitě {place}.",
    "sub_mag": "{tagline}. Navštivte nás v lokalitě {place}.", "sub_corp": "{name} – {tagline}. Ozvěte se nám a do 24 hodin máte odpověď.",
    "trust": ["Nezávazně", "Rychlá odpověď", "Férové ceny"],
    "w_book": "Online rezervace", "days": ["Po", "Út", "St", "Čt", "Pá"], "w_menu": "Dnes v nabídce",
    "menu_rows": [("Polévka dne", "65 Kč"), ("Hlavní chod", "189 Kč"), ("Dezert", "95 Kč")],
    "w_quote": "Nezávazná poptávka", "f_name": "Vaše jméno", "f_phone": "Telefon", "f_need": "Co potřebujete?", "w_send": "Odeslat poptávku",
}


def menu(t, k):
    return "".join(f"<span>{x}</span>" for x in t[k])


def svg(name, cls="i"):
    return f'<svg class="{cls}" viewBox="0 0 24 24">{ICON[name]}</svg>'


def esc(s):
    return html.escape(str(s or ""))


def seed(lead, salt=""):
    return int(hashlib.sha256((lead["name"] + "|" + salt).encode()).hexdigest(), 16)


def photo_path(pid):
    PHOTO_DIR.mkdir(parents=True, exist_ok=True)
    f = PHOTO_DIR / f"{pid}.jpg"
    if not f.exists():
        r = requests.get(f"https://images.unsplash.com/photo-{pid}?w=1920&q=82&fm=jpg&fit=max", timeout=60)
        r.raise_for_status()
        f.write_bytes(r.content)
    return f.as_uri()


# ---------------------------------------------------------------- вибір комбінації
def pick(lead, niche_info):
    from uk_content import NICHE_EN
    key, mood, heads = NICHE.get(lead.get("niche")) or NICHE_EN.get(lead.get("niche")) or FALLBACK_NICHE
    pool = P[key]
    s = seed(lead)
    style = STYLES_BY_MOOD[mood][s % 3]
    palettes = PALETTES[mood]
    pal = palettes[(s // 3) % len(palettes)]
    font = FONTS[style][(s // 17) % len(FONTS[style])]
    start = (s // 101) % len(pool)
    photos = [pool[(start + i) % len(pool)] for i in range(3)]
    headline = heads[(s // 1009) % len(heads)]
    return {"style": style, "pal": pal, "font": font, "photos": photos, "headline": headline, "mood": mood,
            "variant": (s // 7) % 2}


# ---------------------------------------------------------------- спільний CSS
BASE_CSS = Template("""
*{box-sizing:border-box;margin:0;padding:0}
html,body{width:1920px;height:1080px;overflow:hidden}
body{font-family:'$body',Inter,sans-serif;background:$bg;color:$ink;-webkit-font-smoothing:antialiased}
.chrome{height:48px;background:#ECEDEF;display:flex;align-items:center;padding:0 20px;border-bottom:1px solid #D6D8DC}
.chrome .d{display:flex;gap:8px}.chrome .d i{width:12px;height:12px;border-radius:50%;display:block}
.chrome .u{margin:0 auto;width:640px;height:30px;border-radius:15px;background:#fff;display:flex;align-items:center;
  justify-content:center;gap:8px;font:500 14px Inter,sans-serif;color:#3C4043}
.chrome .u svg{width:13px;height:13px;fill:none;stroke:#188038;stroke-width:2.6}
.site{position:relative;height:1032px;overflow:hidden;background:$bg}
.disp{font-family:'$disp',serif;font-weight:$dw;$upper;text-wrap:balance}
.sub{text-wrap:pretty}
em{font-style:italic;color:$accent}
.i{width:20px;height:20px;fill:none;stroke:currentColor;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;flex:none}
.btn{display:inline-flex;align-items:center;gap:12px;padding:22px 36px;font-weight:600;font-size:18px;border-radius:$radius;white-space:nowrap}
.btn.a{background:$accent;color:$on_accent}
.btn.o{border:1.5px solid currentColor}
.kick{font-size:15px;letter-spacing:.24em;text-transform:uppercase;font-weight:600;color:$muted;display:flex;align-items:center;gap:16px}
.kick:before{content:"";width:44px;height:1.5px;background:$accent}
.ph{background-size:cover;background-position:center}
""")


def chrome(domain):
    return (f'<div class="chrome"><div class="d"><i style="background:#FF5F57"></i><i style="background:#FEBC2E"></i>'
            f'<i style="background:#28C840"></i></div><div class="u"><svg viewBox="0 0 24 24"><rect x="5" y="11" '
            f'width="14" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/></svg>{esc(domain)}</div>'
            f'<div style="width:60px"></div></div>')


def rating_html(lead, t=UI_CS):
    r, n = lead.get("rating"), int(lead.get("reviews") or 0)
    try:
        good = float(r) >= 4.3 and n >= 3
    except (TypeError, ValueError):
        good = False
    return (f"{svg('star', 'i st')} <b>{esc(r)}</b>&nbsp;·&nbsp;{n} {t['reviews']}") if good else ""


def ctx(lead, c, extra):
    """Спільні дані для шаблонів."""
    from offers import NICHES, FALLBACK, short_name  # послуги й CTA з offers.py
    from uk_content import NICHES_EN, UI_EN
    t = UI_EN if lead.get("lang") == "en" else UI_CS
    if lead.get("niche") in NICHES_EN:
        label, tagline, cta, services, wkind = NICHES_EN[lead["niche"]]
    else:
        tagline, cta, services, _, wkind = NICHES.get(lead.get("niche"), FALLBACK)
        label = (lead.get("niche") or "").capitalize()
    name = short_name(lead["name"])
    city = [x.strip() for x in (lead.get("city") or "").split(",")]
    place = city[-1] if len(city) > 1 else (city[0] if city else "")
    town = city[0] if city else ""
    return {"name": name, "cta": cta, "services": services, "tagline": tagline, "wkind": wkind, "t": t,
            "place": place, "town": town, "phone": lead.get("phone") or "", "niche": label,
            "rating": rating_html(lead, t), **extra}


def name_size(name, base, step=6, per=4):
    """Менший кегль для довгих назв."""
    over = max(0, len(name) - 14)
    return max(base - (over // per + (1 if over else 0)) * step, int(base * .55))


def logo(d, c, upper=False):
    initials = "".join(w[0] for w in re.findall(r"[A-Za-zÀ-ž]+", d["name"])[:2]).upper()
    n = d["name"].upper() if upper else d["name"]
    return f'<div class="logo"><span class="mk">{esc(initials)}</span><span class="ln">{esc(n)}</span></div>'


def widget(d, c):
    pal = c["pal"]
    if d["wkind"] == "booking":
        days = "".join(f'<span class="{"on" if i == 2 else ""}"><small>{w}</small>{n}</span>'
                       for i, (w, n) in enumerate(zip(d["t"]["days"], [14, 15, 16, 17, 18])))
        slots = "".join(f'<span class="{"on" if i == 1 else ""}">{t}</span>' for i, t in enumerate(["9:30", "11:00", "14:30", "16:00"]))
        return f'<div class="wg"><div class="wh">{svg("cal")} {d["t"]["w_book"]}</div><div class="days">{days}</div><div class="slots">{slots}</div><div class="wb">{esc(d["cta"])}</div></div>'
    if d["wkind"] == "menu":
        rows = "".join(f'<div class="mr"><span>{a}</span><i></i><b>{b}</b></div>' for a, b in
                       d["t"]["menu_rows"])
        return f'<div class="wg"><div class="wh">{svg("clock")} {d["t"]["w_menu"]}</div>{rows}<div class="wb">{esc(d["cta"])}</div></div>'
    t = d["t"]
    return (f'<div class="wg"><div class="wh">{svg("check")} {t["w_quote"]}</div><div class="f">{t["f_name"]}</div>'
            f'<div class="f">{t["f_phone"]}</div><div class="f t">{t["f_need"]}</div><div class="wb">{t["w_send"]}</div></div>')


WIDGET_CSS = Template("""
.wg{background:$surface;color:$ink;border-radius:22px;padding:26px;width:380px;box-shadow:0 30px 70px rgba(0,0,0,.22);border:1px solid $line}
.wh{display:flex;align-items:center;gap:10px;font-weight:700;font-size:18px;margin-bottom:16px}.wh .i{color:$accent}
.days{display:grid;grid-template-columns:repeat(5,1fr);gap:8px}.days span,.slots span{border:1px solid $line;border-radius:12px;
 text-align:center;padding:9px 0;font-weight:700;font-size:17px}.days small{display:block;font-size:11px;color:$muted;font-weight:600}
.slots{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-top:10px}.slots span{font-size:15px;padding:10px 0}
.days .on,.slots .on{background:$accent;border-color:$accent;color:$on_accent}.days .on small{color:inherit;opacity:.8}
.wb{margin-top:16px;background:$ink;color:$bg;text-align:center;padding:15px;border-radius:12px;font-weight:700;font-size:16px}
.mr{display:flex;align-items:baseline;gap:10px;padding:10px 0;font-size:17px}.mr i{flex:1;border-bottom:1.5px dotted $line}
.f{border:1px solid $line;border-radius:10px;padding:13px 14px;color:$muted;font-size:15px;margin-top:10px}.f.t{height:68px}
""")


# ---------------------------------------------------------------- 6 арт-напрямів
def editorial(d, c, ph):
    nm = d["name"]
    big = name_size(nm, 150, 10)
    services = "".join(f'<div class="sv"><span>0{i + 1}</span><b>{esc(t)}</b><small>{esc(s)}</small></div>'
                       for i, (t, s) in enumerate(d["services"]))
    rating = f'<div class="rt">{d["rating"]}</div>' if d["rating"] else ""
    css = """
.hero{position:absolute;inset:0;background:linear-gradient(180deg,rgba(0,0,0,.55) 0%,rgba(0,0,0,.08) 30%,rgba(0,0,0,.25) 55%,rgba(0,0,0,.86) 100%),url(PH0) center/cover}
nav{position:relative;z-index:2;display:flex;align-items:center;padding:34px 80px;color:#fff;gap:48px}
.logo{display:flex;align-items:center;gap:14px}.mk{width:46px;height:46px;border:1.5px solid rgba(255,255,255,.7);border-radius:50%;display:flex;
 align-items:center;justify-content:center;font:600 15px 'BODY';letter-spacing:.06em}.ln{font-family:'DISP';font-size:28px;letter-spacing:.01em}
.mn{display:flex;gap:40px;margin-left:auto;font-size:15px;letter-spacing:.16em;text-transform:uppercase;font-weight:500;opacity:.9}
nav .btn{padding:15px 26px;font-size:15px;border:1.5px solid rgba(255,255,255,.8);color:#fff;border-radius:RADIUS}
.cp{position:absolute;left:80px;bottom:250px;color:#fff;max-width:1250px;z-index:2}
.cp .kick{color:rgba(255,255,255,.85)}.cp .kick:before{background:#fff}
h1{font-size:BIGpx;line-height:.98;letter-spacing:-.02em;margin:26px 0 24px}
h1 em{color:#fff;opacity:.92}
.sub{font-size:24px;line-height:1.5;opacity:.88;max-width:760px}
.bt{display:flex;gap:16px;margin-top:38px}.bt .o{color:#fff}
.rt{position:absolute;right:80px;bottom:270px;z-index:2;color:#fff;font-size:18px;display:flex;align-items:center;gap:6px;
 background:rgba(255,255,255,.12);backdrop-filter:blur(14px);border:1px solid rgba(255,255,255,.25);padding:14px 22px;border-radius:999px}
.st{fill:#F5C451;stroke:#F5C451}
.svs{position:absolute;left:80px;right:80px;bottom:56px;display:grid;grid-template-columns:repeat(4,1fr);border-top:1px solid rgba(255,255,255,.35);z-index:2;color:#fff}
.sv{padding:26px 28px 0 0}.sv span{font-family:'DISP';font-size:22px;color:ACCENT}.sv b{display:block;font-size:22px;font-weight:600;margin:8px 0 4px}
.sv small{font-size:16px;opacity:.75}
"""
    headline = f'{esc(nm)}' if len(nm) <= 26 else d["headline"]
    sub = esc(d["t"]["sub_ed"].format(head=re.sub("<[^>]+>", "", d["headline"]), niche=d["niche"], place=d["place"])) \
        if len(nm) <= 26 else f'{esc(d["niche"])} {esc(nm)} · {esc(d["place"])}'
    body = f"""<div class="hero"></div>
<nav>{logo(d, c)}<div class="mn">{menu(d["t"], "menu1")}</div><span class="btn">{esc(d["phone"] or d["cta"])}</span></nav>
<div class="cp"><div class="kick">{esc(d["niche"])} · {esc(d["town"])}</div><h1 class="disp">{headline}</h1><p class="sub">{sub}</p>
<div class="bt"><span class="btn a">{esc(d["cta"])} {svg("arrow")}</span><span class="btn o">{svg("phone")} {esc(d["phone"] or d["t"]["call_us"])}</span></div></div>
{rating}<div class="svs">{services}</div>"""
    return css.replace("BIG", str(big)), body


def arch(d, c, ph):
    hl = d["headline"]
    services = "".join(f'<div class="sv"><span>0{i + 1}</span>{esc(t)}</div>' for i, (t, _) in enumerate(d["services"]))
    ring = f'{d["cta"].upper()} • {d["name"].upper()[:22]} • '
    css = """
nav{display:flex;align-items:center;padding:34px 96px;gap:48px;position:relative;z-index:3}
.logo{display:flex;align-items:center;gap:14px}.mk{width:44px;height:44px;border-radius:50% 50% 6px 6px;background:ACCENT;color:ONACC;display:flex;
 align-items:center;justify-content:center;font:600 15px 'BODY'}.ln{font-family:'DISP';font-size:30px}
.mn{display:flex;gap:40px;margin-left:auto;font-size:17px;color:MUTED;font-weight:500}
nav .btn{padding:15px 26px;font-size:16px;background:INK;color:BG}
.cp{position:absolute;left:96px;top:210px;width:900px}
h1{font-size:118px;line-height:1.0;letter-spacing:-.015em;margin:30px 0 30px}
.sub{font-size:23px;line-height:1.6;color:MUTED;max-width:640px}
.bt{display:flex;gap:16px;margin-top:42px;align-items:center}.bt .o{color:INK}
.meta{display:flex;gap:34px;margin-top:44px;color:MUTED;font-size:17px}.meta span{display:flex;gap:10px;align-items:center}.meta .i{color:ACCENT}
.ar{position:absolute;right:120px;top:120px;width:600px;height:830px;border-radius:300px 300px 28px 28px;background:url(PH0) center/cover;box-shadow:0 40px 90px rgba(0,0,0,.18)}
.ar2{position:absolute;right:640px;bottom:70px;width:290px;height:290px;border-radius:50%;background:url(PH1) center/cover;border:12px solid BG;box-shadow:0 20px 50px rgba(0,0,0,.15)}
.ring{position:absolute;right:600px;top:130px;width:190px;height:190px;animation:none}
.ring text{font:600 13.5px 'BODY';letter-spacing:.18em;fill:INK}
.ring .c{fill:BG;stroke:LINE}
.dot{position:absolute;right:672px;top:202px;width:46px;height:46px;border-radius:50%;background:ACCENT;display:flex;align-items:center;justify-content:center;color:ONACC}
.ln2{position:absolute;right:90px;top:90px;width:660px;height:890px;border:1.5px solid LINE;border-radius:330px 330px 34px 34px}
.svs{position:absolute;left:96px;bottom:64px;display:flex;gap:44px;font-size:19px;font-weight:600}
.sv span{font-family:'DISP';font-style:italic;color:ACCENT;margin-right:10px;font-size:22px}
.rt{display:inline-flex;align-items:center;gap:6px;font-size:17px;color:INK}.st{fill:#E0A93B;stroke:#E0A93B}
"""
    rating = f'<span class="rt">{d["rating"]}</span>' if d["rating"] else ""
    body = f"""<div class="ln2"></div><div class="ar"></div><div class="ar2"></div>
<svg class="ring" viewBox="0 0 200 200"><circle class="c" cx="100" cy="100" r="99"/><defs><path id="cp" d="M100,100 m-72,0 a72,72 0 1,1 144,0 a72,72 0 1,1 -144,0"/></defs>
<text><textPath href="#cp" textLength="448" lengthAdjust="spacing">{esc(ring)}</textPath></text></svg><div class="dot">{svg("arrow")}</div>
<nav>{logo(d, c)}<div class="mn">{menu(d["t"], "menu2")}</div><span class="btn">{esc(d["cta"])}</span></nav>
<div class="cp"><div class="kick">{esc(d["name"])}</div><h1 class="disp">{hl}</h1>
<p class="sub">{esc(d["t"]["sub_arch"].format(tagline=d["tagline"], niche=d["niche"], place=d["place"]))}</p>
<div class="bt"><span class="btn a">{esc(d["cta"])} {svg("arrow")}</span>{rating}</div>
<div class="meta"><span>{svg("phone")}{esc(d["phone"])}</span><span>{svg("pin")}{esc(d["place"] or d["town"])}</span></div></div>
<div class="svs">{services}</div>"""
    return css, body


def bold(d, c, ph):
    nm = d["name"].upper()
    big = name_size(nm, 190, 14, 3)
    band = " ✦ ".join(t.upper() for t, _ in d["services"] * 3)
    words = nm.split()
    if len(words) > 1:
        half = (len(words) + 1) // 2
        lines = [" ".join(words[:half]), " ".join(words[half:])]
        title = f'{esc(lines[0])}<br><span class="out">{esc(lines[1])}</span>'
    else:
        lines = [nm]
        title = esc(nm)
    # кегль під найдовший рядок, щоб назва не налазила на фото й віджет (~0.5em на літеру)
    big = min(big, int(1100 / (0.5 * max(len(x) for x in lines))))
    css = """
nav{display:flex;align-items:center;padding:30px 72px;gap:48px;position:relative;z-index:3}
.logo{display:flex;align-items:center;gap:14px}.mk{width:48px;height:48px;background:ACCENT;color:ONACC;display:flex;align-items:center;justify-content:center;
 font:400 24px 'DISP';clip-path:polygon(0 0,100% 0,100% 75%,75% 100%,0 100%)}.ln{font:400 26px 'DISP';letter-spacing:.06em}
.mn{display:flex;gap:40px;margin-left:auto;font-size:15px;letter-spacing:.14em;text-transform:uppercase;font-weight:600;color:#fff}
nav .btn{padding:16px 26px;font-size:16px;background:ACCENT;color:ONACC}
.pic{position:absolute;right:0;top:0;width:1000px;height:1032px;background:linear-gradient(180deg,rgba(0,0,0,.72) 0%,rgba(0,0,0,0) 22%),linear-gradient(90deg,BG 0%,rgba(0,0,0,0) 30%),url(PH0) center/cover;
 clip-path:polygon(22% 0,100% 0,100% 100%,0 100%);filter:saturate(.9) contrast(1.05)}
.cp{position:absolute;left:72px;top:180px;z-index:2;width:1180px}
.tag{display:inline-flex;gap:10px;align-items:center;font-weight:700;letter-spacing:.2em;text-transform:uppercase;font-size:14px;color:ACCENT}
.tag:before{content:"";width:34px;height:3px;background:ACCENT}
h1{font-size:BIGpx;line-height:1.0;letter-spacing:.005em;margin:40px 0 30px}
.out{color:transparent;-webkit-text-stroke:2px INK}
.sub{font-size:22px;line-height:1.55;color:MUTED;max-width:600px}
.bt{display:flex;gap:16px;margin-top:40px}.bt .o{color:INK}
.ph2{position:absolute;right:70px;bottom:190px;z-index:3}
.band{position:absolute;left:-40px;right:-40px;bottom:46px;background:ACCENT;color:ONACC;transform:rotate(-1.6deg);padding:20px 0;white-space:nowrap;
 font:400 34px 'DISP';letter-spacing:.06em;z-index:4;overflow:hidden}
.rt{margin-top:28px;display:flex;align-items:center;gap:6px;color:INK;font-size:18px}.st{fill:#F5C451;stroke:#F5C451}
"""
    rating = f'<div class="rt">{d["rating"]}</div>' if d["rating"] else ""
    body = f"""<div class="pic"></div>
<nav>{logo(d, c, upper=True)}<div class="mn">{menu(d["t"], "menu2")}</div><span class="btn">{esc(d["phone"] or d["cta"])}</span></nav>
<div class="cp"><div class="tag">{esc(d["niche"])} · {esc(d["place"] or d["town"])}</div><h1 class="disp">{title}</h1>
<p class="sub">{esc(sub_line(d))}</p>
<div class="bt"><span class="btn a">{esc(d["cta"])} {svg("arrow")}</span><span class="btn o">{svg("phone")} {esc(d["phone"] or d["t"]["call"])}</span></div>{rating}</div>
<div class="ph2">{widget(d, c)}</div>
<div class="band">{esc(band)} ✦ {esc(band)}</div>"""
    return css.replace("BIG", str(big)), body


def sub_line(d):
    """Підзаголовок, що не повторює слоган."""
    head = re.sub("<[^>]+>", "", d["headline"]).lower()
    first = d["tagline"].lower().split()[0].strip(",")
    lead_in = "" if first in head else d["tagline"] + ". "
    return lead_in + d["t"]["sub_line"].format(place=d["place"] or d["town"])


def split(d, c, ph):
    rows = "".join(f'<div class="rw"><span class="n">0{i + 1}</span><div><b>{esc(t)}</b><small>{esc(s)}</small></div>{svg("arrow")}</div>'
                   for i, (t, s) in enumerate(d["services"]))
    left = c["variant"] == 0
    css = """
nav{display:flex;align-items:center;padding:28px 64px;gap:48px;position:relative;z-index:3;background:BG;border-bottom:1px solid LINE}
.logo{display:flex;align-items:center;gap:14px}.mk{width:44px;height:44px;border-radius:12px;background:ACCENT;color:ONACC;display:flex;align-items:center;
 justify-content:center;font:700 16px 'BODY'}.ln{font:DW 25px 'DISP';letter-spacing:-.01em}
.mn{display:flex;gap:38px;margin-left:auto;font-size:17px;color:MUTED;font-weight:500}
nav .btn{padding:15px 24px;font-size:16px;background:ACCENT;color:ONACC}
.pic{position:absolute;top:101px;bottom:0;SIDE:0;width:880px;background:url(PH0) center/cover}
.pic .bd{position:absolute;OTHER:-1px;bottom:56px;background:SURFACE;padding:22px 28px;border-radius:RADIUS 0 0 RADIUS;display:flex;gap:14px;align-items:center;
 font-size:17px;color:INK;box-shadow:0 20px 40px rgba(0,0,0,.15)}
.bd .i{color:ACCENT;width:26px;height:26px}
.cp{position:absolute;top:150px;CPSIDE:88px;width:860px}
h1{font-size:90px;line-height:1.02;letter-spacing:-.025em;margin:26px 0 22px}
.sub{font-size:21px;line-height:1.6;color:MUTED;max-width:720px}
.rws{margin-top:34px;border-top:1px solid LINE}
.rw{display:flex;align-items:center;gap:26px;padding:19px 0;border-bottom:1px solid LINE}
.rw .n{font:DW 20px 'DISP';color:ACCENT;width:34px}.rw b{font-size:22px;font-weight:600;display:block}.rw small{font-size:16px;color:MUTED}
.rw div{flex:1}.rw .i{color:MUTED}
.bt{display:flex;gap:16px;margin-top:34px}.bt .o{color:INK}
.st{fill:#E0A93B;stroke:#E0A93B}
"""
    css = css.replace("CPSIDE", "right" if left else "left").replace("OTHER", "right" if left else "left")
    css = css.replace("SIDE", "left" if left else "right")
    if not left:
        css = css.replace("border-radius:RADIUS 0 0 RADIUS", "border-radius:0 RADIUS RADIUS 0")
    badge = d["rating"] or f'{svg("pin")} {esc(d["niche"])} · {esc(d["place"] or d["town"])}'
    body = f"""<nav>{logo(d, c)}<div class="mn">{menu(d["t"], "menu3")}</div><span class="btn">{svg("phone")} {esc(d["phone"] or d["cta"])}</span></nav>
<div class="pic"><div class="bd">{badge}</div></div>
<div class="cp"><div class="kick">{esc(d["name"])}</div><h1 class="disp">{d["headline"]}</h1>
<p class="sub">{esc(sub_line(d))}</p>
<div class="rws">{rows}</div>
<div class="bt"><span class="btn a">{esc(d["cta"])} {svg("arrow")}</span><span class="btn o">{svg("phone")} {esc(d["phone"] or d["t"]["call"])}</span></div></div>"""
    return css, body


def magazine(d, c, ph):
    band = " ✦ ".join(t for t, _ in d["services"] * 3)
    css = """
nav{display:grid;grid-template-columns:1fr auto 1fr;align-items:center;padding:30px 80px;position:relative;z-index:3;border-bottom:1px solid LINE}
.mn{display:flex;gap:34px;font-size:15px;letter-spacing:.14em;text-transform:uppercase;font-weight:600;color:MUTED}.mn.r{justify-content:flex-end;align-items:center}
.logo{display:flex;flex-direction:column;align-items:center;gap:2px}.mk{display:none}.ln{font:DW 34px 'DISP';letter-spacing:-.01em}
nav .btn{padding:13px 22px;font-size:15px;background:ACCENT;color:ONACC}
.cp{position:absolute;left:80px;top:170px;width:820px}
.no{font:DW 20px 'DISP';color:ACCENT;letter-spacing:.06em}
h1{font-size:104px;line-height:1.02;letter-spacing:-.025em;margin:22px 0 28px;max-width:800px}
.sub{font-size:22px;line-height:1.6;color:MUTED;max-width:640px}
.bt{display:flex;gap:16px;margin-top:40px}.bt .o{color:INK}
.meta{margin-top:40px;display:flex;gap:30px;font-size:17px;color:MUTED}.meta span{display:flex;align-items:center;gap:9px}.meta .i{color:ACCENT}
.g{position:absolute;right:80px;top:140px;width:900px;height:720px;display:grid;grid-template-columns:1.15fr 1fr;grid-template-rows:1fr 1fr;gap:20px}
.g div{border-radius:RADIUS;background-size:cover;background-position:center;position:relative}
.g .a{grid-row:1/3;background-image:url(PH0)}.g .b{background-image:url(PH1)}.g .c{background-image:url(PH2)}
.g .cap{position:absolute;left:18px;bottom:18px;background:SURFACE;color:INK;padding:10px 16px;border-radius:999px;font-size:14px;font-weight:600}
.band{position:absolute;left:0;right:0;bottom:0;height:120px;background:INK;color:BG;display:flex;align-items:center;white-space:nowrap;overflow:hidden;
 font:italic DW 40px 'DISP';gap:0;padding-left:80px}
.st{fill:#E0A93B;stroke:#E0A93B}
"""
    rating = f'<span>{d["rating"]}</span>' if d["rating"] else ""
    body = f"""<nav><div class="mn">{"".join(f"<span>{x}</span>" for x in d["t"]["menu2"][:3])}</div>{logo(d, c)}<div class="mn r"><span>{d["t"]["menu2"][3]}</span><span class="btn">{esc(d["cta"])}</span></div></nav>
<div class="cp"><div class="no">— {esc(d["niche"])}, {esc(d["town"])}</div><h1 class="disp">{d["headline"]}</h1>
<p class="sub">{esc(d["t"]["sub_mag"].format(tagline=d["tagline"], place=d["place"] or d["town"]))}</p>
<div class="bt"><span class="btn a">{esc(d["cta"])} {svg("arrow")}</span><span class="btn o">{svg("phone")} {esc(d["phone"] or d["t"]["call"])}</span></div>
<div class="meta">{rating}<span>{svg("pin")}{esc(d["place"] or d["town"])}</span></div></div>
<div class="g"><div class="a"><span class="cap">{esc(d["services"][0][0])}</span></div><div class="b"></div><div class="c"><span class="cap">{esc(d["services"][1][0])}</span></div></div>
<div class="band">{esc(band)}</div>"""
    return css, body


def corporate(d, c, ph):
    cards = "".join(f'<div class="cd"><div class="ic">{svg(["check", "clock", "star", "pin"][i])}</div><b>{esc(t)}</b><small>{esc(s)}</small></div>'
                    for i, (t, s) in enumerate(d["services"]))
    css = """
.bgm{position:absolute;inset:0;background:radial-gradient(900px 600px at 85% 10%,ACCENT22,transparent 60%),radial-gradient(700px 500px at 0% 100%,ACCENT14,transparent 60%),BG}
nav{display:flex;align-items:center;padding:28px 80px;gap:48px;position:relative;z-index:3}
.logo{display:flex;align-items:center;gap:12px}.mk{width:42px;height:42px;border-radius:12px;background:linear-gradient(135deg,ACCENT,INK);color:#fff;display:flex;
 align-items:center;justify-content:center;font:700 15px 'BODY'}.ln{font:DW 25px 'DISP';letter-spacing:-.02em}
.mn{display:flex;gap:38px;margin:0 auto;font-size:17px;color:MUTED;font-weight:500;background:SURFACE;padding:14px 32px;border-radius:999px;border:1px solid LINE}
nav .btn{padding:15px 26px;font-size:16px;background:INK;color:BG;border-radius:999px}
.cp{position:absolute;left:80px;top:170px;width:860px;z-index:2}
.pill{display:inline-flex;align-items:center;gap:8px;background:SURFACE;border:1px solid LINE;padding:10px 18px;border-radius:999px;font-size:16px;font-weight:600}
.pill .i{color:ACCENT}
h1{font-size:92px;line-height:1.02;letter-spacing:-.035em;margin:28px 0 24px}
h1 em{font-style:normal;background:linear-gradient(90deg,ACCENT,INK);-webkit-background-clip:text;color:transparent}
.sub{font-size:22px;line-height:1.6;color:MUTED;max-width:700px}
.bt{display:flex;gap:16px;margin-top:38px}.bt .btn{border-radius:999px}.bt .o{color:INK}
.tr{display:flex;gap:28px;margin-top:38px;font-size:17px;color:MUTED}.tr span{display:flex;align-items:center;gap:8px}.tr .i{color:ACCENT}
.pic{position:absolute;right:80px;top:130px;width:820px;height:600px;border-radius:32px;background:url(PH0) center/cover;box-shadow:0 40px 80px rgba(15,25,50,.18)}
.fl{position:absolute;right:760px;top:470px;z-index:3}
.fl2{position:absolute;right:110px;top:100px;z-index:3;background:SURFACE;border-radius:18px;padding:18px 22px;box-shadow:0 20px 50px rgba(0,0,0,.14);
 display:flex;gap:14px;align-items:center;font-size:16px;border:1px solid LINE}
.fl2 .ic2{width:44px;height:44px;border-radius:12px;background:ACCENT;color:ONACC;display:flex;align-items:center;justify-content:center}
.fl2 b{display:block;font-size:19px}
.cds{position:absolute;left:80px;right:80px;bottom:52px;display:grid;grid-template-columns:repeat(4,1fr);gap:22px;z-index:2}
.cd{background:SURFACE;border:1px solid LINE;border-radius:22px;padding:24px 26px}
.ic{width:46px;height:46px;border-radius:12px;background:ACCENT1A;color:ACCENT;display:flex;align-items:center;justify-content:center;margin-bottom:14px}
.cd b{font-size:20px;display:block}.cd small{font-size:15px;color:MUTED}
.st{fill:#F5B400;stroke:#F5B400}
"""
    pill = d["rating"] or f'{svg("pin")} {esc(d["niche"])} · {esc(d["place"] or d["town"])}'
    body = f"""<div class="bgm"></div>
<nav>{logo(d, c)}<div class="mn">{menu(d["t"], "menu3")}</div><span class="btn">{esc(d["cta"])}</span></nav>
<div class="cp"><span class="pill">{pill}</span><h1 class="disp">{d["headline"]}</h1>
<p class="sub">{esc(d["t"]["sub_corp"].format(name=d["name"], tagline=d["tagline"][:1].lower() + d["tagline"][1:]))}</p>
<div class="bt"><span class="btn a">{esc(d["cta"])} {svg("arrow")}</span><span class="btn o">{svg("phone")} {esc(d["phone"] or d["t"]["call"])}</span></div>
<div class="tr">{"".join(f'<span>{svg("check")} {x}</span>' for x in d["t"]["trust"])}</div></div>
<div class="pic"></div><div class="fl2"><div class="ic2">{svg("phone")}</div><div><small>{d["t"]["call_us"]}</small><b>{esc(d["phone"])}</b></div></div>
<div class="fl">{widget(d, c)}</div>
<div class="cds">{cards}</div>"""
    return css, body


STYLE_FN = {"editorial": editorial, "arch": arch, "bold": bold, "split": split, "magazine": magazine, "corporate": corporate}


def render(lead, domain, force_style=None):
    c = pick(lead, None)
    if force_style:
        c["style"] = force_style
        c["font"] = FONTS[force_style][seed(lead, "f") % len(FONTS[force_style])]
    bg, surface, ink, muted, accent, on_accent, line = c["pal"]
    disp, body, dw, upper = c["font"]
    d = ctx(lead, c, {"headline": c["headline"]})
    photos = [photo_path(p) for p in c["photos"]]
    css_style, body_html = STYLE_FN[c["style"]](d, c, photos)
    radius = {"bold": "4px", "editorial": "999px", "arch": "999px", "split": "14px", "magazine": "6px",
              "corporate": "14px"}[c["style"]]
    rep = {"PH0": photos[0], "PH1": photos[1], "PH2": photos[2], "ACCENT": accent, "ONACC": on_accent, "INK": ink,
           "MUTED": muted, "SURFACE": surface, "LINE": line, "BG": bg, "RADIUS": radius, "DISP": disp, "BODY": body,
           "DW": str(dw)}
    for k in sorted(rep, key=len, reverse=True):
        css_style = css_style.replace(k, rep[k])
    base = BASE_CSS.substitute(body=body, disp=disp, dw=dw, upper="text-transform:uppercase;" if upper else "",
                               bg=bg, ink=ink, muted=muted, accent=accent, on_accent=on_accent, radius=radius)
    wcss = WIDGET_CSS.substitute(surface=surface, ink=ink, line=line, accent=accent, on_accent=on_accent, muted=muted, bg=bg)
    page = f"""<!doctype html><html lang="cs"><head><meta charset="utf-8"><link rel="stylesheet" href="{GFONTS}">
<style>{base}{wcss}{css_style}</style></head><body>{chrome(domain)}<div class="site">{body_html}</div></body></html>"""
    return page, c
