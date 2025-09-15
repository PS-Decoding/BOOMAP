/*** GeomUtils.gs ***/

// 좌표 문자열 파서 (세미콜론/개행/연속 공백, "lat,lon" / "lon,lat" / "x y" / WKT POLYGON 허용)
function GEOM_parsePoly(str) {
  if (!str) return null;
  let s = String(str).trim();
  s = s.replace(/^MULTIPOLYGON\s*\(\(+/i, "")
       .replace(/^POLYGON\s*\(\(+/i, "")
       .replace(/\)+\s*$/g, "");

  const rawSegs = (s.includes(";") || s.includes("\n"))
    ? s.split(/[\n;]+/)
    : s.split(/\s{2,}/);

  const vertices = [];
  for (let seg of rawSegs) {
    seg = seg.trim();
    if (!seg) continue;

    let a, b;
    if (seg.includes(",")) {
      const parts = seg.split(",").map(t => t.trim());
      if (parts.length < 2) continue;
      a = Number(parts[0]); b = Number(parts[1]);
    } else {
      const parts = seg.split(/\s+/);
      if (parts.length < 2) continue;
      a = Number(parts[0]); b = Number(parts[1]);
    }
    if (!isFinite(a) || !isFinite(b)) continue;

    // lat/lon 자동 보정
    let la = a, lo = b;
    if (Math.abs(la) > 90 || (Math.abs(lo) <= 90 && Math.abs(la) > Math.abs(lo))) {
      const t = la; la = lo; lo = t;
    }
    vertices.push({lat: la, lon: lo});
  }
  return vertices.length ? vertices : null;
}

function GEOM_bbox(vertices) {
  if (!Array.isArray(vertices) || vertices.length === 0) return null;
  let minLat =  Infinity, maxLat = -Infinity;
  let minLon =  Infinity, maxLon = -Infinity;
  for (const v of vertices) {
    const la = Number(v && v.lat), lo = Number(v && v.lon);
    if (!isFinite(la) || !isFinite(lo)) continue;
    if (la < minLat) minLat = la; if (la > maxLat) maxLat = la;
    if (lo < minLon) minLon = lo; if (lo > maxLon) maxLon = lo;
  }
  if (!isFinite(minLat) || !isFinite(maxLat) || !isFinite(minLon) || !isFinite(maxLon)) return null;
  return { minLat, maxLat, minLon, maxLon };
}

function GEOM_pointOnSeg(x0,y0,x1,y1,x2,y2,eps){
  const cross = Math.abs((x2-x1)*(y0-y1) - (y2-y1)*(x0-x1));
  if (cross > eps) return false;
  const withinX = (Math.min(x1,x2)-eps<=x0) && (x0<=Math.max(x1,x2)+eps);
  const withinY = (Math.min(y1,y2)-eps<=y0) && (y0<=Math.max(y1,y2)+eps);
  return withinX && withinY;
}
function GEOM_onBoundary(lat, lon, vertices) {
  const eps = 1e-8;
  if (!Array.isArray(vertices) || vertices.length < 2) return false;
  for (let i = 0, j = vertices.length - 1; i < vertices.length; j = i++) {
    const a = vertices[j], b = vertices[i];
    if (!a || !b) continue;
    const ax = Number(a.lon), ay = Number(a.lat);
    const bx = Number(b.lon), by = Number(b.lat);
    if (!isFinite(ax) || !isFinite(ay) || !isFinite(bx) || !isFinite(by)) continue;
    if (GEOM_pointOnSeg(lon, lat, ax, ay, bx, by, eps)) return true;
  }
  return false;
}
function GEOM_pointInPoly(lat,lon,vertices){
  if (!Array.isArray(vertices)||vertices.length<3) return false;
  if (GEOM_onBoundary(lat,lon,vertices)) return true;
  let inside=false;
  for (let i=0,j=vertices.length-1; i<vertices.length; j=i++){
    const xi=vertices[i].lon, yi=vertices[i].lat;
    const xj=vertices[j].lon, yj=vertices[j].lat;
    const intersect=((yi>lat)!==(yj>lat)) && (lon < (xj-xi)*(lat-yi)/((yj-yi)||1e-16)+xi);
    if (intersect) inside=!inside;
  }
  return inside;
}
