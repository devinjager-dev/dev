"""Bouwt een Excel-overzicht van alle charms uit
https://nl.charmsoff.com/collections/customize-your-bracelet

Gebruik:
    python build_charms_overzicht.py <map-met-page_*.json> <uitvoer.xlsx>

De page_*.json bestanden zijn de publieke Shopify-feed van de collectie:
    https://nl.charmsoff.com/collections/customize-your-bracelet/products.json?limit=250&page=N

Het bestand wordt gemaakt met XlsxWriter, omdat alleen die bibliotheek echte
Excel-checkboxes (Microsoft 365) kan schrijven. Niet opnieuw opslaan via
LibreOffice of openpyxl: dan verdwijnen de checkboxes.
"""
import collections
import datetime
import glob
import json
import os
import re
import sys

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

SHOP = "https://nl.charmsoff.com"
COLLECTION_URL = SHOP + "/collections/customize-your-bracelet"

# Volgorde en Nederlandse labels van de thema's (tags in de shop).
THEMES = [
    ("Countries & Travel", "Landen & Reizen"),
    ("Letters & Numbers", "Letters & Cijfers"),
    ("Love & Hearts", "Liefde & Hartjes"),
    ("Animals", "Dieren"),
    ("Zodiac", "Sterrenbeelden"),
    ("Quotes & Positivity", "Quotes & Positiviteit"),
    ("Sun Moon & Stars", "Zon, Maan & Sterren"),
    ("Nature & Flowers", "Natuur & Bloemen"),
    ("Food & Drink", "Eten & Drinken"),
    ("Luck & Spirituality", "Geluk & Spiritualiteit"),
    ("Family & Friends", "Familie & Vrienden"),
    ("Seasonal", "Seizoensgebonden"),
    ("Beauty & Girls", "Beauty & Girls"),
    ("Sport & Music", "Sport & Muziek"),
    ("Birthstones & Months", "Geboortestenen & Maanden"),
    ("Money & Casino", "Geld & Casino"),
    ("Classic & Plain", "Klassiek & Effen"),
]
THEME_ORDER = [nl for _, nl in THEMES] + ["Mysteriebox"]

COLORS = ["Zilver", "Goud", "Roségoud", "Zwart", "Gemengd"]
TYPES = ["Klassieke link-bedel", "Hangende bedel", "Dubbele bedel", "Dubbel verbonden bedel"]

METAL_WORD = {"zilver": "Zilver", "goud": "Goud", "roségoud": "Roségoud"}


def kleur(p):
    """Metaalkleur van de schakel.
    1. Expliciet in de naam: de laatste 'op/aan zilver|goud|roségoud'
       ('Letter A in goud met stenen, op Zilver' -> Zilver).
    2. Shop-tags, zoals de armbandbouwer op de site: Rose Gold, dan Silver, dan gold.
    3. Producttype / bestandsnaam van de foto. Wat overblijft zijn de zwarte
       Halloween-charms (zwarte schakel, gecontroleerd op de productfoto's)."""
    tags = set(p["tags"])
    title = p["title"].lower()
    img = (p["images"][0]["src"].lower() if p["images"] else "")
    if "mystery" in tags:
        return "Gemengd"
    said = re.findall(r"\b(?:op|aan)\s+(zilver|goud|roségoud)\b", title)
    if said:
        return METAL_WORD[said[-1]]
    if "Rose Gold" in tags or p["product_type"] == "Roségoud":
        return "Roségoud"
    if "Silver" in tags:
        return "Zilver"
    if "gold" in tags:
        return "Goud"
    if p["product_type"] == "Zilver" or "on-silver" in img:
        return "Zilver"
    if p["product_type"] == "Goud" or "on-gold" in img:
        return "Goud"
    return "Zwart"


def charm_type(p):
    tags = set(p["tags"])
    if "linked" in tags:
        return "Dubbel verbonden bedel"
    if "double" in tags:
        return "Dubbele bedel"
    if "hanging" in tags:
        return "Hangende bedel"
    return "Klassieke link-bedel"


def themas(p):
    tags = set(p["tags"])
    if "mystery" in tags:
        return ["Mysteriebox"]
    found = [nl for en, nl in THEMES if en in tags]
    return found or ["(geen thema)"]


def load(folder):
    prods = []
    for f in sorted(glob.glob(os.path.join(folder, "page_*.json"))):
        with open(f, encoding="utf-8") as fh:
            prods += json.load(fh)["products"]
    seen, out = set(), []
    for p in prods:
        if p["id"] in seen:
            continue
        seen.add(p["id"])
        v = p["variants"][0]
        out.append({
            "naam": p["title"].strip().rstrip(","),
            "kleur": kleur(p),
            "type": charm_type(p),
            "themas": themas(p),
            "sku": v.get("sku") or "",
            "prijs": float(v["price"]),
            "van_prijs": float(v["compare_at_price"]) if v.get("compare_at_price") else None,
            "bestseller": "Ja" if "best-seller-top60" in p["tags"] else "",
            "nieuw": "Ja" if "new" in p["tags"] else "",
            "voorraad": "Ja" if v["available"] else "Nee",
            "link": f"{SHOP}/products/{p['handle']}",
            "afbeelding": p["images"][0]["src"] if p["images"] else "",
        })
    # De link is de sleutel waarmee afvinken over tabbladen heen wordt
    # vergeleken (COUNTIF/MATCH); die mag dus geen jokertekens bevatten.
    assert len({r["link"] for r in out}) == len(out), "links niet uniek"
    assert not any(ch in r["link"] for r in out for ch in "*?~"), "jokerteken in link"
    return out


def sort_key(r):
    t = r["themas"][0]
    return (COLORS.index(r["kleur"]), TYPES.index(r["type"]),
            THEME_ORDER.index(t) if t in THEME_ORDER else 99, r["naam"].lower())


# ---------- opmaak ----------
# Villa for You thema: hoofdkleur #3CC3AF. Witte tekst op #3CC3AF haalt maar
# 2,18:1 contrast, dus koppen krijgen donkere tekst (6,73:1). Titels, links en
# de kopjes van de invulkolommen gebruiken een donkere tint (wit erop: 6,08:1).
VFY = "#3CC3AF"        # hoofdkleur
VFY_DARK = "#1A6E63"   # donkere tint: titels, links, kop invulkolommen
VFY_TINT = "#E6F7F4"   # lichte tint: totalen
VFY_ZEBRA = "#F4FBFA"  # zeer lichte tint: om-en-om regels
VFY_PICKED = "#C9EEE7"  # regel die je zelf hebt afgevinkt
INK = "#1F2A2E"        # tekstkleur
GRID = "#CFE3E0"       # randen
COLOR_FILL = {"Zilver": "#E6E8EB", "Goud": "#F6E7B8", "Roségoud": "#F4D6CC",
              "Zwart": "#BFBFBF", "Gemengd": "#DDEBF7"}
EURO = "€ #,##0.00"


class Formats:
    """Maakt XlsxWriter-formats aan en hergebruikt identieke."""

    BASE = {"font_name": "Arial", "font_size": 10, "font_color": INK, "valign": "vcenter"}

    def __init__(self, wb):
        self.wb = wb
        self.cache = {}

    def __call__(self, **props):
        full = dict(self.BASE, **props)
        key = tuple(sorted(full.items()))
        if key not in self.cache:
            self.cache[key] = self.wb.add_format(full)
        return self.cache[key]

    def cell(self, zebra=False, **props):
        p = {"border": 1, "border_color": GRID}
        if zebra:
            p["bg_color"] = VFY_ZEBRA
        p.update(props)
        return self(**p)


def q(sheet):
    """Bladnaam veilig in een formule."""
    return "'" + sheet.replace("'", "''") + "'"


# ---------- lijstbladen ----------
# Elke lijst begint met dezelfde drie kolommen, daarna de charmgegevens en
# aan het eind drie verborgen hulpkolommen voor de status-formule.
INPUT_COLS = [("Afvinken", 10), ("Gereserveerd door", 22), ("Status", 30)]
HELPER_COLS = ["Sleutel (hulp)", "Aantal keer afgevinkt (hulp)", "Naam elders (hulp)"]

COLUMNS = {
    "kleur": ("Kleur", 11), "type": ("Type", 21), "themas": ("Thema('s)", 32),
    "thema": ("Thema", 24), "naam": ("Naam", 50), "sku": ("SKU", 13),
    "prijs": ("Prijs", 10), "van_prijs": ("Van-prijs", 10), "bestseller": ("Bestseller", 11),
    "nieuw": ("Nieuw", 8), "voorraad": ("Op voorraad", 11), "link": ("Link", 58),
    "afbeelding": ("Afbeelding (URL)", 58),
}


class ListSheet:
    def __init__(self, title, rows, fields):
        self.title = title
        self.rows = rows            # lijst van (thema of None, charm-dict)
        self.fields = fields
        self.n = len(rows)
        first_data = len(INPUT_COLS)
        self.col = {f: first_data + i for i, f in enumerate(fields)}
        self.last_visible = first_data + len(fields) - 1
        self.key_col = self.last_visible + 1
        self.tot_col = self.key_col + 1
        self.name_col = self.tot_col + 1

    def rng(self, col):
        c = xl_col_to_name(col)
        return f"{q(self.title)}!${c}$2:${c}${self.n + 1}"


def build(folder, out_path):
    data = load(folder)
    all_rows = sorted(data, key=sort_key)

    sheets = [ListSheet("Alle charms", [(None, r) for r in all_rows],
                        ["kleur", "type", "themas", "naam", "sku", "prijs", "van_prijs",
                         "bestseller", "nieuw", "voorraad", "link", "afbeelding"])]
    color_fields = ["type", "themas", "naam", "sku", "prijs", "bestseller", "voorraad", "link"]
    for tab, kleuren in [("Zilver", ["Zilver"]), ("Goud", ["Goud"]), ("Roségoud", ["Roségoud"]),
                         ("Zwart & mysteriebox", ["Zwart", "Gemengd"])]:
        sheets.append(ListSheet(tab, [(None, r) for r in all_rows if r["kleur"] in kleuren],
                                color_fields))
    long_rows = [(t, r) for r in data for t in r["themas"]]
    long_rows.sort(key=lambda tr: (THEME_ORDER.index(tr[0]) if tr[0] in THEME_ORDER else 99,
                                   COLORS.index(tr[1]["kleur"]), TYPES.index(tr[1]["type"]),
                                   tr[1]["naam"].lower()))
    sheets.append(ListSheet("Per thema", long_rows,
                            ["thema", "kleur", "type", "naam", "sku", "prijs", "voorraad", "link"]))
    by_title = {s.title: s for s in sheets}

    wb = xlsxwriter.Workbook(out_path, {"strings_to_urls": False})
    F = Formats(wb)
    ws_ov = wb.add_worksheet("Overzicht")  # eerst aanmaken zodat het vooraan staat

    for s in sheets:
        write_list_sheet(wb, F, s, sheets)

    write_overview(F, ws_ov, data, by_title["Alle charms"], by_title["Per thema"], long_rows)
    wb.close()
    return data


def write_list_sheet(wb, F, s, sheets):
    ws = wb.add_worksheet(s.title)
    ws.set_tab_color(VFY)
    ws.hide_gridlines(2)

    head = F(bold=True, bg_color=VFY, text_wrap=True, border=1, border_color=GRID)
    head_input = F(bold=True, bg_color=VFY_DARK, font_color="#FFFFFF", text_wrap=True,
                   border=1, border_color=GRID)
    head_helper = F(italic=True, font_color="#777777", bg_color="#EEEEEE", text_wrap=True)
    ws.set_row(0, 30)
    for c, (label, width) in enumerate(INPUT_COLS):
        ws.write_string(0, c, label, head_input if c < 2 else head)
        ws.set_column(c, c, width)
    for f, c in s.col.items():
        label, width = COLUMNS[f]
        ws.write_string(0, c, label, head)
        ws.set_column(c, c, width)
    for i, label in enumerate(HELPER_COLS):
        c = s.key_col + i
        ws.write_string(0, c, label, head_helper)
        ws.set_column(c, c, 14, None, {"hidden": True})

    link_c = xl_col_to_name(s.col["link"])
    key_c, tot_c, name_c = (xl_col_to_name(c) for c in (s.key_col, s.tot_col, s.name_col))

    for i, (thema, r) in enumerate(s.rows):
        row = i + 1          # 0-based rij in XlsxWriter
        R = row + 1          # Excel-rijnummer
        zebra = (R % 2 == 1)

        # Invulkolommen
        ws.insert_checkbox(row, 0, False,
                           F.cell(zebra, checkbox=True, align="center"))
        ws.write_blank(row, 1, None, F.cell(zebra))

        # Gegevens
        values = dict(r, thema=thema, themas=" / ".join(r["themas"]))
        for f, c in s.col.items():
            v = values[f]
            if f in ("prijs", "van_prijs"):
                if v is None:
                    ws.write_blank(row, c, None, F.cell(zebra, num_format=EURO))
                else:
                    ws.write_number(row, c, v, F.cell(zebra, num_format=EURO))
            elif f == "link":
                ws.write_url(row, c, v, F.cell(zebra, font_color=VFY_DARK, underline=1), v)
            elif f == "kleur":
                ws.write_string(row, c, v, F.cell(bg_color=COLOR_FILL[v]))
            else:
                ws.write_string(row, c, v, F.cell(zebra))

        # Hulpkolommen:
        #   sleutel = link als deze regel is afgevinkt
        #   aantal  = hoe vaak deze charm op alle bladen samen is afgevinkt
        #   naam    = eerste naam bij een afgevinkte regel van deze charm
        ws.write_formula(row, s.key_col, f'=IF($A{R}=TRUE,${link_c}{R},"")', None, "")
        count = "+".join(f"COUNTIF({o.rng(o.key_col)},${link_c}{R})" for o in sheets)
        ws.write_formula(row, s.tot_col, "=" + count, None, 0)
        lookup = '""'
        for o in reversed(sheets):
            lookup = (f"IFERROR(INDEX({o.rng(1)},MATCH(${link_c}{R},{o.rng(o.key_col)},0))&\"\","
                      f"{lookup})")
        ws.write_formula(row, s.name_col, "=" + lookup, None, "")

        # Status
        status = (f'=IF(${tot_c}{R}-IF($A{R}=TRUE,1,0)<=0,"",'
                  f'IF($A{R}=TRUE,"Let op: ook elders afgevinkt",'
                  f'"Al gereserveerd"&IF(${name_c}{R}<>""," door "&${name_c}{R},"")))')
        ws.write_formula(row, 2, status, F.cell(zebra, italic=True), "")

    last = s.n
    lv = s.last_visible
    # Voorwaardelijke opmaak (eerste regel heeft voorrang)
    ws.conditional_format(1, 2, last, 2, {
        "type": "formula", "criteria": f"=AND($A2=TRUE,${tot_c}2>1)",
        "format": wb.add_format({"bg_color": "#FCE4D6", "font_color": "#9C2A00", "bold": True})})
    ws.conditional_format(1, 0, last, lv, {
        "type": "formula", "criteria": "=$A2=TRUE",
        "format": wb.add_format({"bg_color": VFY_PICKED})})
    ws.conditional_format(1, 0, last, lv, {
        "type": "formula", "criteria": f"=AND($A2<>TRUE,${tot_c}2>0)",
        "format": wb.add_format({"bg_color": "#F2F2F2", "font_color": "#8C8C8C"})})

    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, last, lv)


def write_overview(F, ws, data, s_all, s_theme, long_rows):
    ws.set_tab_color(VFY_DARK)
    ws.hide_gridlines(2)
    widths = [26, 20, 20, 20, 22, 12, 14, 12]
    for c, w in enumerate(widths):
        ws.set_column(c, c, w)

    title = F(bold=True, font_size=16, font_color=VFY_DARK)
    sub = F(bold=True, font_size=11, font_color=VFY_DARK)
    note = F(italic=True, font_size=9, font_color="#555555")
    head = F(bold=True, bg_color=VFY, text_wrap=True, border=1, border_color=GRID)
    num = F.cell()
    bold_total = F.cell(bold=True, bg_color=VFY_TINT)
    label_total = F.cell(bold=True, bg_color=VFY_TINT)

    ws.write_string(0, 0, "Charms overzicht: Maak je armband (nl.charmsoff.com)", title)
    ws.write_string(1, 0, f"Bron: {COLLECTION_URL} (publieke productfeed), opgehaald op "
                          f"{datetime.date.today():%d-%m-%Y}. {len(data)} producten.", note)
    ws.set_row(2, 4)
    for c in range(len(widths)):
        ws.write_blank(2, c, None, F(bg_color=VFY))

    # Zo werkt afvinken
    ws.write_string(3, 0, "Zo werkt afvinken", sub)
    howto = [
        "Zet op een tabblad een vinkje in 'Afvinken' en vul je naam in bij 'Gereserveerd door' "
        "(de donkergroene kopjes zijn de invulkolommen).",
        "Een charm staat op meerdere tabbladen. 'Status' laat automatisch zien als dezelfde charm "
        "al ergens anders is afgevinkt, en door wie. Zo'n regel wordt grijs.",
        "Twee keer afgevinkt? Dan staat er 'Let op: ook elders afgevinkt' in oranje.",
        "Checkboxes werken in Microsoft 365 (Windows, Mac en Excel online). In oudere Excel-versies "
        "staat er ONWAAR; typ dan WAAR om af te vinken.",
    ]
    for i, line in enumerate(howto):
        ws.write_string(4 + i, 0, line, note)

    # Tabel 1: kleur x type, met reserveringen
    kt = collections.Counter((r["kleur"], r["type"]) for r in data)
    kc = collections.Counter(r["kleur"] for r in data)
    r0 = 9
    ws.write_string(r0, 0, "Aantal charms per kleur en type", sub)
    hdr = ["Kleur"] + TYPES + ["Totaal", "Gereserveerd", "Nog vrij"]
    ws.set_row(r0 + 1, 30)
    for c, label in enumerate(hdr):
        ws.write_string(r0 + 1, c, label, head)
    rng_k = s_all.rng(s_all.col["kleur"])
    rng_t = s_all.rng(s_all.col["type"])
    rng_res = s_all.rng(s_all.tot_col)
    nt = len(TYPES)
    for i, k in enumerate(COLORS):
        row = r0 + 2 + i
        R = row + 1
        ws.write_string(row, 0, k, F.cell(bold=True, bg_color=COLOR_FILL[k]))
        for j, t in enumerate(TYPES):
            hc = xl_col_to_name(1 + j)
            ws.write_formula(row, 1 + j, f"=COUNTIFS({rng_k},$A{R},{rng_t},{hc}${r0 + 2})",
                             num, kt[(k, t)])
        last_t = xl_col_to_name(nt)
        tot_c = xl_col_to_name(1 + nt)
        res_c = xl_col_to_name(2 + nt)
        ws.write_formula(row, 1 + nt, f"=SUM(B{R}:{last_t}{R})", bold_total, kc[k])
        ws.write_formula(row, 2 + nt, f'=COUNTIFS({rng_k},$A{R},{rng_res},">0")', num, 0)
        ws.write_formula(row, 3 + nt, f"={tot_c}{R}-{res_c}{R}", bold_total, kc[k])
    tot_row = r0 + 2 + len(COLORS)
    ws.write_string(tot_row, 0, "Totaal", label_total)
    col_sums = [sum(kt[(k, t)] for k in COLORS) for t in TYPES] + [len(data), 0, len(data)]
    for c in range(1, len(hdr)):
        cl = xl_col_to_name(c)
        ws.write_formula(tot_row, c, f"=SUM({cl}{r0 + 3}:{cl}{tot_row})", bold_total,
                         col_sums[c - 1])

    # Tabel 2: thema x kleur
    tk = collections.Counter((t, r["kleur"]) for t, r in long_rows)
    tc = collections.Counter(t for t, _ in long_rows)
    r1 = tot_row + 3
    ws.write_string(r1, 0, "Aantal charms per thema en kleur", sub)
    ws.write_string(r1, 3, "Let op: een charm kan in meer thema's vallen, dus het totaal hier "
                           "is hoger dan het aantal charms.", note)
    hdr2 = ["Thema"] + COLORS + ["Totaal"]
    for c, label in enumerate(hdr2):
        fmt = head
        if label in COLOR_FILL:
            fmt = F(bold=True, bg_color=COLOR_FILL[label], border=1, border_color=GRID)
        ws.write_string(r1 + 1, c, label, fmt)
    rng_th = s_theme.rng(s_theme.col["thema"])
    rng_tk = s_theme.rng(s_theme.col["kleur"])
    themes_present = [t for t in THEME_ORDER if tc[t]]
    for i, t in enumerate(themes_present):
        row = r1 + 2 + i
        R = row + 1
        ws.write_string(row, 0, t, F.cell(bold=True))
        for j, k in enumerate(COLORS):
            hc = xl_col_to_name(1 + j)
            ws.write_formula(row, 1 + j, f"=COUNTIFS({rng_th},$A{R},{rng_tk},{hc}${r1 + 2})",
                             num, tk[(t, k)])
        last_k = xl_col_to_name(len(COLORS))
        ws.write_formula(row, 1 + len(COLORS), f"=SUM(B{R}:{last_k}{R})", bold_total, tc[t])

    # Toelichting
    r2 = r1 + 3 + len(themes_present)
    notes = [
        "Kleur = metaalkleur van de schakel. 'Op zilver/goud/roségoud' in de productnaam weegt "
        "het zwaarst, daarna de shop-tags. Voorbeeld: 'Letter A in goud met stenen, op Zilver' "
        "telt als Zilver.",
        "Zwart = de Halloween-charms op een zwarte schakel (gecontroleerd op de productfoto's). "
        "Gemengd = mysteryboxen.",
        "Type = 'Klassieke link-bedel' of 'Hangende bedel' (zo noemt de site het), plus "
        "'Dubbele bedel' (dubbele breedte) en 'Dubbel verbonden bedel' (twee gekoppelde bedels).",
        "Thema = de categorieën uit de armbandbouwer. Bestseller = tag 'best-seller-top60'.",
        "Gereserveerd = aantal charms dat op minstens één tabblad is afgevinkt (elke charm telt "
        "één keer).",
    ]
    ws.write_string(r2, 0, "Toelichting", sub)
    for i, n in enumerate(notes):
        ws.write_string(r2 + 1 + i, 0, n, note)


if __name__ == "__main__":
    rows = build(sys.argv[1], sys.argv[2])
    print(f"{len(rows)} producten geschreven naar {sys.argv[2]}")
