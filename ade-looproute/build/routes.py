import json, sys, time, urllib.request, itertools
OUT = sys.argv[1]
# BAG coordinates (PDOK Locatieserver), lon/lat
LOCS = {
  "froda":   (4.87104, 52.37125),  # De Clercqstraat 40H
  "quattro": (4.88656, 52.37223),  # Hartenstraat 3
  "hertz":   (4.89575, 52.36655),  # Halvemaansteeg 14
  "stmpd":   (4.89717, 52.37764),  # Prins Hendrikkade 33
  "bierfab": (4.89300, 52.37053),  # Rokin 75 (inferred)
  "chenin":  (4.88026, 52.36998),  # Hazenstraat 60C
  "poeder":  (4.93324, 52.33784),  # Duivendrechtsekade 80B
  "zwartgoud": (4.90177, 52.37417),# Geldersekade 89
  "tokyo":   (4.89149, 52.36622),  # Reguliersdwarsstraat 44
}
ids = list(LOCS)
legs = {}
for a, b in itertools.combinations(ids, 2):
    (x1, y1), (x2, y2) = LOCS[a], LOCS[b]
    url = f"https://routing.openstreetmap.de/routed-foot/route/v1/driving/{x1},{y1};{x2},{y2}?overview=full&geometries=geojson"
    for attempt in range(4):
        try:
            with urllib.request.urlopen(url, timeout=60) as r: d = json.load(r)
            break
        except Exception as e:
            print("retry", a, b, e, file=sys.stderr); time.sleep(2 ** (attempt + 1))
    r = d["routes"][0]
    coords = [[round(x, 5), round(y, 5)] for x, y in r["geometry"]["coordinates"]]
    legs[f"{a}|{b}"] = {"m": round(r["distance"]), "s": round(r["duration"]), "g": coords}
    print(a, b, round(r["distance"]), round(r["duration"] / 60, 1), file=sys.stderr)
    time.sleep(0.6)
json.dump(legs, open(OUT, "w"), separators=(",", ":"))
