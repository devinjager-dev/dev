"""Bouwt een Excel-overzicht van alle charms uit
https://nl.charmsoff.com/collections/customize-your-bracelet

Gebruik:
    python build_charms_overzicht.py <map-met-page_*.json> <uitvoer.xlsx>

De page_*.json bestanden zijn de publieke Shopify-feed van de collectie:
    https://nl.charmsoff.com/collections/customize-your-bracelet/products.json?limit=250&page=N
"""
import datetime
import glob
import json
import os
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

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
THEME_NL = dict(THEMES)
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
    # Overige zonder metaaltag: de zwarte Halloween-charms (zwarte schakel,
    # gecontroleerd op de productfoto's).
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
    return out


# ---------- opmaak ----------
FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True)
F_SUB = Font(name=FONT, size=11, bold=True)
F_NOTE = Font(name=FONT, size=9, italic=True, color="555555")
F_LINK = Font(name=FONT, size=10, color="0563C1", underline="single")
F_BOLD = Font(name=FONT, size=10, bold=True)
FILL_HEAD = PatternFill("solid", fgColor="3A3F44")
FILL_TOTAL = PatternFill("solid", fgColor="E8E8E8")
COLOR_FILL = {
    "Zilver": PatternFill("solid", fgColor="E6E8EB"),
    "Goud": PatternFill("solid", fgColor="F6E7B8"),
    "Roségoud": PatternFill("solid", fgColor="F4D6CC"),
    "Zwart": PatternFill("solid", fgColor="BFBFBF"),
    "Gemengd": PatternFill("solid", fgColor="DDEBF7"),
}
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
EURO = '€ #,##0.00'


def header(ws, row, labels):
    for c, label in enumerate(labels, 1):
        cell = ws.cell(row=row, column=c, value=label)
        cell.font = F_HEAD
        cell.fill = FILL_HEAD
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = BORDER


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def sort_key(r, theme=None):
    t = theme if theme else r["themas"][0]
    return (COLORS.index(r["kleur"]), TYPES.index(r["type"]),
            THEME_ORDER.index(t) if t in THEME_ORDER else 99, r["naam"].lower())


def write_list(ws, rows, columns, row_fn, col_widths, color_col=None):
    """Schrijft een tabel met kop, filter en vastgezette kop."""
    header(ws, 1, columns)
    for i, r in enumerate(rows, 2):
        values = row_fn(r)
        for c, (key, val) in enumerate(values, 1):
            cell = ws.cell(row=i, column=c, value=val)
            cell.font = F_BASE
            cell.border = BORDER
            if key in ("prijs", "van_prijs"):
                cell.number_format = EURO
            elif key == "link" and val:
                cell.hyperlink = val
                cell.font = F_LINK
            elif key == "kleur" and val in COLOR_FILL:
                cell.fill = COLOR_FILL[val]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{len(rows) + 1}"
    widths(ws, col_widths)
    ws.row_dimensions[1].height = 28


def build(folder, out_path):
    data = load(folder)
    wb = Workbook()

    # ---------- Alle charms ----------
    ws_all = wb.active
    ws_all.title = "Alle charms"
    all_rows = sorted(data, key=sort_key)
    cols_all = ["Kleur", "Type", "Thema('s)", "Naam", "SKU", "Prijs", "Van-prijs",
                "Bestseller", "Nieuw", "Op voorraad", "Link", "Afbeelding (URL)"]
    write_list(
        ws_all, all_rows, cols_all,
        lambda r: [("kleur", r["kleur"]), ("type", r["type"]),
                   ("themas", " / ".join(r["themas"])), ("naam", r["naam"]),
                   ("sku", r["sku"]), ("prijs", r["prijs"]), ("van_prijs", r["van_prijs"]),
                   ("bestseller", r["bestseller"]), ("nieuw", r["nieuw"]),
                   ("voorraad", r["voorraad"]), ("link", r["link"]),
                   ("afbeelding", r["afbeelding"])],
        [11, 22, 34, 52, 14, 10, 10, 11, 8, 11, 60, 60],
    )

    # ---------- Per kleur ----------
    cols_color = ["Type", "Thema('s)", "Naam", "SKU", "Prijs", "Bestseller",
                  "Op voorraad", "Link"]
    color_tabs = [("Zilver", ["Zilver"]), ("Goud", ["Goud"]), ("Roségoud", ["Roségoud"]),
                  ("Zwart & mysteriebox", ["Zwart", "Gemengd"])]
    for tab, kleuren in color_tabs:
        ws = wb.create_sheet(tab)
        rows = sorted([r for r in data if r["kleur"] in kleuren], key=sort_key)
        write_list(
            ws, rows, cols_color,
            lambda r: [("type", r["type"]), ("themas", " / ".join(r["themas"])),
                       ("naam", r["naam"]), ("sku", r["sku"]), ("prijs", r["prijs"]),
                       ("bestseller", r["bestseller"]), ("voorraad", r["voorraad"]),
                       ("link", r["link"])],
            [22, 34, 52, 14, 10, 11, 11, 60],
        )

    # ---------- Per thema (een regel per charm per thema) ----------
    ws_theme = wb.create_sheet("Per thema")
    long_rows = [(t, r) for r in data for t in r["themas"]]
    long_rows.sort(key=lambda tr: (THEME_ORDER.index(tr[0]) if tr[0] in THEME_ORDER else 99,
                                   COLORS.index(tr[1]["kleur"]), TYPES.index(tr[1]["type"]),
                                   tr[1]["naam"].lower()))
    cols_theme = ["Thema", "Kleur", "Type", "Naam", "SKU", "Prijs", "Op voorraad", "Link"]
    write_list(
        ws_theme, long_rows, cols_theme,
        lambda tr: [("thema", tr[0]), ("kleur", tr[1]["kleur"]), ("type", tr[1]["type"]),
                    ("naam", tr[1]["naam"]), ("sku", tr[1]["sku"]), ("prijs", tr[1]["prijs"]),
                    ("voorraad", tr[1]["voorraad"]), ("link", tr[1]["link"])],
        [26, 11, 22, 52, 14, 10, 11, 60],
    )

    # ---------- Overzicht (aantallen met formules) ----------
    ws = wb.create_sheet("Overzicht", 0)
    n_all = len(all_rows) + 1
    n_long = len(long_rows) + 1
    ws["A1"] = "Charms overzicht: Maak je armband (nl.charmsoff.com)"
    ws["A1"].font = F_TITLE
    ws["A2"] = (f"Bron: {COLLECTION_URL} (publieke productfeed), opgehaald op "
                f"{datetime.date.today():%d-%m-%Y}. {len(data)} producten.")
    ws["A2"].font = F_NOTE

    # Tabel 1: kleur x type
    r0 = 4
    ws.cell(row=r0, column=1, value="Aantal charms per kleur en type").font = F_SUB
    header(ws, r0 + 1, ["Kleur"] + TYPES + ["Totaal"])
    rng_k = f"'Alle charms'!$A$2:$A${n_all}"
    rng_t = f"'Alle charms'!$B$2:$B${n_all}"
    for i, k in enumerate(COLORS):
        row = r0 + 2 + i
        c = ws.cell(row=row, column=1, value=k)
        c.font = F_BOLD
        c.fill = COLOR_FILL[k]
        c.border = BORDER
        for j, _t in enumerate(TYPES):
            col = 2 + j
            tl = get_column_letter(col)
            cell = ws.cell(row=row, column=col,
                           value=f"=COUNTIFS({rng_k},$A{row},{rng_t},{tl}${r0 + 1})")
            cell.font = F_BASE
            cell.border = BORDER
        last = get_column_letter(1 + len(TYPES))
        cell = ws.cell(row=row, column=2 + len(TYPES), value=f"=SUM(B{row}:{last}{row})")
        cell.font = F_BOLD
        cell.border = BORDER
        cell.fill = FILL_TOTAL
    tot_row = r0 + 2 + len(COLORS)
    ws.cell(row=tot_row, column=1, value="Totaal").font = F_BOLD
    for col in range(1, 3 + len(TYPES)):
        cell = ws.cell(row=tot_row, column=col)
        if col > 1:
            cl = get_column_letter(col)
            cell.value = f"=SUM({cl}{r0 + 2}:{cl}{tot_row - 1})"
        cell.font = F_BOLD
        cell.fill = FILL_TOTAL
        cell.border = BORDER

    # Tabel 2: thema x kleur
    r1 = tot_row + 3
    ws.cell(row=r1, column=1, value="Aantal charms per thema en kleur").font = F_SUB
    ws.cell(row=r1, column=4,
            value="Let op: een charm kan in meer thema's vallen, dus het totaal hier is "
                  "hoger dan het aantal charms.").font = F_NOTE
    header(ws, r1 + 1, ["Thema"] + COLORS + ["Totaal"])
    rng_th = f"'Per thema'!$A$2:$A${n_long}"
    rng_tk = f"'Per thema'!$B$2:$B${n_long}"
    themes_present = [t for t in THEME_ORDER if any(t == tr[0] for tr in long_rows)]
    for i, t in enumerate(themes_present):
        row = r1 + 2 + i
        c = ws.cell(row=row, column=1, value=t)
        c.font = F_BOLD
        c.border = BORDER
        for j, _k in enumerate(COLORS):
            col = 2 + j
            kl = get_column_letter(col)
            cell = ws.cell(row=row, column=col,
                           value=f"=COUNTIFS({rng_th},$A{row},{rng_tk},{kl}${r1 + 1})")
            cell.font = F_BASE
            cell.border = BORDER
        last = get_column_letter(1 + len(COLORS))
        cell = ws.cell(row=row, column=2 + len(COLORS), value=f"=SUM(B{row}:{last}{row})")
        cell.font = F_BOLD
        cell.border = BORDER
        cell.fill = FILL_TOTAL
    for j, k in enumerate(COLORS):
        hc = ws.cell(row=r1 + 1, column=2 + j)
        hc.fill = COLOR_FILL[k]
        hc.font = F_BOLD

    # Toelichting
    r2 = r1 + 3 + len(themes_present)
    notes = [
        "Toelichting",
        "Kleur = metaalkleur van de schakel. Bepaald zoals de armbandbouwer op de site doet "
        "(tag 'Silver' gaat voor 'gold'), met Roségoud en Zwart als eigen kleur. Voorbeeld: "
        "'Letter A in goud met stenen, op Zilver' telt als Zilver.",
        "Zwart = de Halloween-charms op een zwarte schakel (geen metaaltag in de shop; "
        "gecontroleerd op de productfoto's). Gemengd = mysteryboxen.",
        "Type = 'Klassieke link-bedel' of 'Hangende bedel' (zo noemt de site het), plus "
        "'Dubbele bedel' (dubbele breedte) en 'Dubbel verbonden bedel' (twee gekoppelde "
        "bedels) op basis van de shop-tags 'double' en 'linked'.",
        "Thema = de categorieën uit de armbandbouwer (Engelse shop-tags vertaald naar het "
        "Nederlands). Bestseller = tag 'best-seller-top60'.",
        "Tabbladen per kleur zijn gesorteerd op type, daarna thema en naam. Het tabblad "
        "'Alle charms' heeft filters op elke kolom.",
    ]
    for i, n in enumerate(notes):
        c = ws.cell(row=r2 + i, column=1, value=n)
        c.font = F_SUB if i == 0 else F_NOTE

    widths(ws, [26, 20, 20, 20, 22, 12, 12])
    for row in ws.iter_rows(min_row=r0 + 1, max_row=r0 + 1):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r0 + 1].height = 30
    ws.sheet_view.showGridLines = False

    wb.save(out_path)
    return data


if __name__ == "__main__":
    rows = build(sys.argv[1], sys.argv[2])
    print(f"{len(rows)} producten geschreven naar {sys.argv[2]}")
