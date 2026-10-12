import json, sys, time, urllib.request, os
OUT = sys.argv[1]
BBOX = "4.850,52.326,4.952,52.392"
COLS = ["waterdeel_vlak", "wegdeel_hartlijn", "terrein_vlak", "spoorbaandeel_lijn"]
os.makedirs(OUT, exist_ok=True)
for c in COLS:
    url = f"https://api.pdok.nl/brt/top10nl/ogc/v1/collections/{c}/items?f=json&limit=1000&bbox={BBOX}"
    feats = []
    while url:
        for attempt in range(4):
            try:
                with urllib.request.urlopen(url, timeout=120) as r:
                    d = json.load(r)
                break
            except Exception as e:
                print("retry", c, e, file=sys.stderr); time.sleep(2 ** attempt)
        else:
            raise SystemExit("failed " + url)
        feats += d["features"]
        url = next((l["href"] for l in d.get("links", []) if l["rel"] == "next"), None)
        print(c, len(feats), file=sys.stderr)
    json.dump({"type": "FeatureCollection", "features": feats}, open(f"{OUT}/{c}.json", "w"))
