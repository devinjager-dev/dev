import json, sys
from shapely.geometry import shape, box, mapping, MultiPolygon, Polygon, LineString, MultiLineString
from shapely.ops import unary_union, linemerge
SRC, OUT = sys.argv[1], sys.argv[2]
CLIP = box(4.850, 52.326, 4.952, 52.392)
def load(n): return json.load(open(f"{SRC}/{n}.json"))["features"]
def rnd(o):
    if isinstance(o, (list, tuple)):
        if o and isinstance(o[0], (int, float)): return [round(o[0], 5), round(o[1], 5)]
        return [rnd(x) for x in o]
    return o
def polys(geom, tol, min_area):
    g = geom.intersection(CLIP).simplify(tol, preserve_topology=True)
    parts = list(g.geoms) if hasattr(g, "geoms") else [g]
    parts = [p for p in parts if isinstance(p, Polygon) and p.area >= min_area]
    return [rnd(mapping(p)["coordinates"]) for p in parts]
def lines(geoms, tol, min_len=0):
    g = linemerge(unary_union(geoms)).intersection(CLIP).simplify(tol, preserve_topology=False)
    parts = list(g.geoms) if hasattr(g, "geoms") else [g]
    out = []
    for p in parts:
        if isinstance(p, LineString) and p.length >= min_len:
            c = rnd(list(p.coords))
            dedup = [c[0]] + [pt for i, pt in enumerate(c[1:], 1) if pt != c[i-1]]
            if len(dedup) > 1: out.append(dedup)
    return out
res = {}
water = unary_union([shape(f["geometry"]).buffer(0) for f in load("waterdeel_vlak")])
res["water"] = polys(water, 0.000015, 2e-9)
GREEN = {"grasland", "bos: loofbos", "bos: gemengd bos", "dodenakker", "boomgaard"}
green = unary_union([shape(f["geometry"]).buffer(0) for f in load("terrein_vlak") if f["properties"].get("typelandgebruik") in GREEN])
res["green"] = polys(green, 0.00003, 2e-8)
roads = {"major": [], "street": [], "foot": []}
for f in load("wegdeel_hartlijn"):
    p = f["properties"]
    if p.get("hoogteniveau", 0) != 0 or p.get("status") != "in gebruik": continue
    t, u = p.get("typeweg"), p.get("hoofdverkeersgebruik")
    g = shape(f["geometry"])
    if t in ("autosnelweg", "regionale weg", "lokale weg"): roads["major"].append(g)
    elif u == "gemengd verkeer": roads["street"].append(g)
    elif u == "voetgangers": roads["foot"].append(g)
for k, v in roads.items(): res[k] = lines(v, 0.000012)
rail = [shape(f["geometry"]) for f in load("spoorbaandeel_lijn") if f["properties"].get("typespoorbaan") == "trein" and f["properties"].get("hoogteniveau") == 0]
res["rail"] = lines(rail, 0.00002)
s = json.dumps(res, separators=(",", ":"))
open(OUT, "w").write(s)
print({k: len(v) for k, v in res.items()}, "bytes", len(s))
