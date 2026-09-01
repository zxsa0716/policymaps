#!/usr/bin/env python
"""V-World 행정경계 → 프런트 지도용 GeoJSON.

region_geometry(V-World LT_C_ADSIGG_INFO / LT_C_ADSIDO_INFO 실호출분)에서
기초자치단체 227곳을 뽑아 표시용으로 단순화한다. 일반구를 둔 13개 시는
DB 적재 단계에서 이미 하위 일반구를 병합(dissolve)해 두었다.

사용:  python system/make_map_geojson.py
출력:  viz/public/geo/municipalities.geojson
"""
import sys, json, sqlite3, pathlib
sys.stdout.reconfigure(encoding="utf-8")
from shapely.geometry import shape, mapping, MultiPolygon, Polygon
from shapely.validation import make_valid

def polys(g):
    if isinstance(g,(Polygon,MultiPolygon)): return g
    if hasattr(g,"geoms"):
        out=[]
        for x in g.geoms:
            if isinstance(x,MultiPolygon): out.extend(x.geoms)
            elif isinstance(x,Polygon): out.append(x)
        if out: return MultiPolygon(out) if len(out)>1 else out[0]
    return g
def rnd(o,n=5):
    if isinstance(o,(list,tuple)):
        if o and isinstance(o[0],(int,float)): return [round(float(o[0]),n),round(float(o[1]),n)]
        return [rnd(x,n) for x in o]
    return o
def nv(o):
    if isinstance(o,(list,tuple)):
        return 1 if (o and isinstance(o[0],(int,float))) else sum(nv(x) for x in o)
    return 0

DB=r"C:\policy_maps\system\data\policymap.db"
con=sqlite3.connect(f"file:{DB}?mode=ro",uri=True,timeout=300); con.row_factory=sqlite3.Row
rows=con.execute("""SELECT g.region_id,g.geojson,r.name FROM region_geometry g
  JOIN regions r ON r.region_id=g.region_id
  WHERE r.level=2 AND r.status='active' ORDER BY g.region_id""").fetchall()
sido={r["region_id"]:r["name"] for r in con.execute("SELECT region_id,name FROM regions WHERE level=1")}
con.close()

TOL=0.002
feats=[]; tot=0
for r in rows:
    g=shape(json.loads(r["geojson"]))
    if not g.is_valid: g=polys(make_valid(g))
    g=polys(g.simplify(TOL,preserve_topology=True))
    gm=mapping(g)
    co=rnd(gm["coordinates"])
    if gm["type"]=="Polygon": co=[co]
    tot+=nv(co)
    feats.append({"type":"Feature",
      "properties":{"sig_cd":r["region_id"],"name":r["name"],
                    "sido":sido.get(r["region_id"][:2],""),"verified":True},
      "geometry":{"type":"MultiPolygon","coordinates":co}})
out={"type":"FeatureCollection",
 "_meta":{"source":"국토교통부 V-World 공간정보 오픈플랫폼 데이터 API",
   "layer":"LT_C_ADSIGG_INFO(시군구)·LT_C_ADSIDO_INFO(시도)","crs":"EPSG:4326",
   "fetched":"2026-09-01","features":len(feats),
   "note":"일반구를 둔 13개 시는 하위 일반구를 병합(dissolve). 표시용 단순화 "
          f"tolerance {TOL}도(약 {int(TOL*111000)} m), 좌표 소수 5자리",
   "simplify_tolerance_deg":TOL,"coord_precision":5},
 "features":feats}
p=pathlib.Path(r"C:\policy_maps\viz\public\geo\municipalities.geojson")
p.write_text(json.dumps(out,ensure_ascii=False,separators=(",",":")),encoding="utf-8")
print(f"  {len(feats)} features · 정점 {tot:,} · {p.stat().st_size/1024:.0f} KB")
print(f"  (기존 southkorea-maps: 250 features · 정점 32,223 · 643 KB)")
