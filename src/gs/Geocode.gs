/***** Geocode.gs — C열(도로명/단지명) → H열(동+지번)
 * - Kakao(Local REST)만 사용: 주소검색 → (실패 시) 키워드→좌표→지번 폴백
 * - 좌표(D/E)는 판정(F/G)용으로 유지
 *****/

// 전역 메모(실행 스코프 내 URL→응답 캐시)
var __KAKAO_MEMO = Object.create(null);

/* (0) Kakao REST Key */
function getKakaoKey_() {
  try {
    if (typeof KAKAO_REST_KEY !== "undefined" && String(KAKAO_REST_KEY).trim()) {
      return String(KAKAO_REST_KEY).trim();
    }
  } catch (_) {}
  const props = PropertiesService.getScriptProperties();
  const fromProp = props.getProperty("KAKAO_REST_KEY");
  return fromProp && String(fromProp).trim();
}

/* (1) onEdit: C열 변경 시 처리 */
function onEditHandler(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (!sh || CFG.TARGET_SHEETS.indexOf(sh.getName()) === -1) return;

    const cS = e.range.getColumn(), cE = e.range.getLastColumn();
    if (cE < CFG.COL_ADDR || cS > CFG.COL_ADDR) return;

    const rS = Math.max(e.range.getRow(), CFG.ROW_START);
    const rE = e.range.getLastRow();
    if (rS > rE) return;

    // 주소 빈칸 → 좌표 비우기
    const addrVals = sh.getRange(rS, CFG.COL_ADDR, rE - rS + 1, 1).getValues();
    for (let i = 0; i < addrVals.length; i++) {
      if (String(addrVals[i][0] || "").trim() === "") {
        const rng = sh.getRange(rS + i, CFG.COL_LAT, 1, 2); // D:E
        try { rng.clearDataValidations(); } catch (_) {}
        rng.clearContent();
      }
    }

    // 지오코딩 → H라벨
    geocodeRowsUnique_(sh, rS, rE, /*onlyEmpty=*/false);
    updateJibeonLabel_(sh, rS, rE);

    // ❗ 여기: 편집된 구간 전부 직접 판정
    const rows = [];
    for (let i = rS; i <= rE; i++) rows.push(i);
    JUDGE_updateSubset_(sh, rows);

  } catch (err) {
    Logger.log("[onEditHandler] " + err);
  }
}

/* (2) 지오코딩(유니크 주소) → 좌표(D/E) */
function geocodeRowsUnique_(sh, rowStart, rowEnd, onlyEmpty) {
  const n = rowEnd - rowStart + 1;
  if (n <= 0) return;

  const addrVals = sh.getRange(rowStart, CFG.COL_ADDR, n, 1).getValues(); // C
  const latVals  = sh.getRange(rowStart, CFG.COL_LAT,  n, 1).getValues(); // D
  const lonVals  = sh.getRange(rowStart, CFG.COL_LON,  n, 1).getValues(); // E

  const outLat = latVals.map(r => [r[0]]);
  const outLon = lonVals.map(r => [r[0]]);
  const needIdx = [];
  const empties = [];

  for (let i = 0; i < n; i++) {
    const a = normForGeocodeInput_(addrVals[i][0]);
    if (!a) { empties.push(i); outLat[i][0] = ""; outLon[i][0] = ""; continue; }
    if (onlyEmpty && outLat[i][0] !== "" && outLon[i][0] !== "") continue;
    needIdx.push(i);
  }

  ensureLatLonFormats_(sh, rowStart, n);

  // 주소 빈칸 → 좌표 클리어
  for (const i of empties) {
    const rng = sh.getRange(rowStart + i, CFG.COL_LAT, 1, 2); // D:E
    try { rng.clearDataValidations(); } catch (_) {}
    rng.clearContent();
  }

  // 유니크 주소 지오코딩(카카오)
  if (needIdx.length) {
    const addrToIdx = new Map();
    for (const i of needIdx) {
      const a = normForGeocodeInput_(addrVals[i][0]);
      if (!addrToIdx.has(a)) addrToIdx.set(a, []);
      addrToIdx.get(a).push(i);
    }
    const runCache = new Map();

    for (const [addr, indices] of addrToIdx.entries()) {
      // 캐시 확인
      let hit = runCache.get(addr) || getCached_(addr);
      if (!hit) {
        hit = kakaoGeocodeToCoords_(addr);           // ← 카카오로 좌표 얻기
        if (hit) { setCached_(addr, hit); runCache.set(addr, hit); }
        else     { runCache.set(addr, null); }
      }
      for (const i of indices) {
        if (hit) { outLat[i][0] = hit.lat; outLon[i][0] = hit.lon; }
        else     { outLat[i][0] = "";      outLon[i][0] = "";      }
      }
    }
  }

  sh.getRange(rowStart, CFG.COL_LAT, n, 1).setValues(outLat);
  sh.getRange(rowStart, CFG.COL_LON, n, 1).setValues(outLon);
}

/* === H열(지번) 업데이트: 좌표 우선 → (없으면) 주소/키워드 폴백 === */
function updateJibeonLabel_(sh, rowStart, rowEnd) {
  __KAKAO_MEMO = Object.create(null); // 배치 시작 시 URL→응답 메모 초기화

  const n = rowEnd - rowStart + 1;
  if (n <= 0) return;

  const key = getKakaoKey_();
  const vals = sh.getRange(rowStart, CFG.COL_ADDR, n, 1).getValues();     // C
  const lat  = sh.getRange(rowStart, CFG.COL_LAT,  n, 1).getValues();     // D
  const lon  = sh.getRange(rowStart, CFG.COL_LON,  n, 1).getValues();     // E

  const out  = new Array(n).fill("").map(_=>[""]);

  // 1) 좌표 중복 최소화를 위한 런타임 캐시
  const coordMemo = new Map(); // "lat,lon" -> label

  for (let i = 0; i < n; i++) {
    const raw = String(vals[i][0] || "").trim();
    const la  = Number(lat[i][0]);
    const lo  = Number(lon[i][0]);

    // (A) 좌표가 있으면: 카카오 역지오코딩으로 지번 생성
    if (isFinite(la) && isFinite(lo)) {
      const k = la.toFixed(7) + "," + lo.toFixed(7);
      let label = coordMemo.get(k);
      if (label === undefined) {
        label = kakaoCoord2Jibeon_(la, lo) || "";
        coordMemo.set(k, label);
      }
      out[i][0] = label || "확인필요";
      continue;
    }

    // (B) 좌표가 없고 키도 없으면 표시
    if (!raw) { out[i][0] = ""; continue; }
    if (!key) { out[i][0] = "(Kakao 키 필요)"; continue; }

    // (C) 주소/키워드로 지번 라벨 폴백
    const queries = buildRoadQueries_(raw);
    let label = kakaoRoadToJibeon_TryList_(queries);
    if (!label) {
      const kw = extractComplexName_(raw) || raw;
      label = kakaoKeywordToJibeon_(kw, raw);
    }
    out[i][0] = label || "확인필요";
  }

  sh.getRange(rowStart, CFG.COL_ADR_LABEL, n, 1).setValues(out);
}

/* (4) Kakao API 래퍼/폴백 */
function kakaoFetch_(url) {
  if (__KAKAO_MEMO[url] !== undefined) return __KAKAO_MEMO[url];
  const key = getKakaoKey_();
  if (!key) return (__KAKAO_MEMO[url] = null);
  const headers = { "Authorization": "KakaoAK " + key };

  const t0 = Date.now();
  let delay = 120;
  for (let i = 0; i < 3; i++) {
    try {
      const res = UrlFetchApp.fetch(url, { method: "get", headers, muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        const data = JSON.parse(res.getContentText() || "{}");
        __KAKAO_MEMO[url] = data;
        return data;
      }
    } catch (_) {}
    if (Date.now() - t0 > 2500) break;
    Utilities.sleep(delay);
    delay = Math.min(delay * 2, 600);
  }
  __KAKAO_MEMO[url] = null;
  return null;
}

function kakaoRoadToJibeon_TryList_(queries) {
  const tried = new Set();
  for (const q of queries) {
    const qq = (q || "").trim();
    if (!qq || tried.has(qq)) continue;
    tried.add(qq);

    const v = kakaoRoadToJibeonOne_(qq, /*similar=*/false) || kakaoRoadToJibeonOne_(qq, /*similar=*/true);
    if (v) return v;
  }
  return "";
}

function kakaoRoadToJibeonOne_(query, similar) {
  const url = "https://dapi.kakao.com/v2/local/search/address.json"
    + "?size=3"
    + (similar ? "&analyze_type=similar" : "&analyze_type=exact")
    + "&query=" + encodeURIComponent(query);

  const data = kakaoFetch_(url);
  if (!data || !Array.isArray(data.documents) || !data.documents.length) return "";

  let dongOnly = "";

  for (const doc of data.documents) {
    const fromStruct = extractDongJibeonFromAddressDoc_(doc);
    if (hasLot_(fromStruct)) return fromStruct;
    if (!dongOnly && fromStruct) dongOnly = fromStruct;

    const s1 = String(doc.address_name || "");
    const s2 = String(doc.road_address && doc.road_address.address_name || "");
    const fromStr = parseDongJibeonFromString_(s1) || parseDongJibeonFromString_(s2);
    if (hasLot_(fromStr)) return fromStr;
    if (!dongOnly && fromStr) dongOnly = fromStr;

    const x = Number(doc.x), y = Number(doc.y);
    if (isFinite(x) && isFinite(y)) {
      const fromCoord = kakaoCoord2Jibeon_(y, x);
      if (hasLot_(fromCoord)) return fromCoord;
      if (!dongOnly && fromCoord) dongOnly = fromCoord;
    }
  }
  return dongOnly;
}

function kakaoKeywordToJibeon_(kw, rawForRegion) {
  const raw = String(rawForRegion || kw || "");
  const { si, gu, dong } = getRegionTokens_(raw);
  const roadHints = buildRoadQueries_(raw).slice(0, 2);

  const queries = _dedupe_([
    `${gu} ${kw}`, `${si} ${gu} ${kw}`, `${dong} ${kw}`,
    kw,
    `${gu} ${kw} ${roadHints[0] || ""}`.trim(),
    `${kw} ${roadHints[0] || ""}`.trim()
  ]);

  let dongOnly = "";

  for (const q of queries) {
    const url = "https://dapi.kakao.com/v2/local/search/keyword.json?size=5&query=" + encodeURIComponent(q);
    const data = kakaoFetch_(url);
    if (!data || !Array.isArray(data.documents) || !data.documents.length) continue;

    for (const d of data.documents) {
      const adrStr = String(d.address_name || "");
      const fromAdr = parseDongJibeonFromString_(adrStr);
      if (hasLot_(fromAdr)) return fromAdr;
      if (!dongOnly && fromAdr) dongOnly = fromAdr;

      const x = Number(d.x), y = Number(d.y);
      if (isFinite(x) && isFinite(y)) {
        const fromCoord = kakaoCoord2Jibeon_(y, x);
        if (hasLot_(fromCoord)) return fromCoord;
        if (!dongOnly && fromCoord) dongOnly = fromCoord;
      }
    }
  }
  return dongOnly;
}

function kakaoCoord2Jibeon_(lat, lon) {
  if (!isFinite(lat) || !isFinite(lon)) return "";
  const url = "https://dapi.kakao.com/v2/local/geo/coord2address.json?input_coord=WGS84&x="
    + encodeURIComponent(lon) + "&y=" + encodeURIComponent(lat);
  const data = kakaoFetch_(url);
  if (!data || !Array.isArray(data.documents) || !data.documents.length) return "";
  const doc = data.documents[0] || {};
  const adr = doc.address || {};
  const dong = normalizeAdminToLegalDong_(adr.region_3depth_name || "");
  const main = String(adr.main_address_no || "").trim();
  const sub  = String(adr.sub_address_no  || "").trim();
  if (dong) return main ? `${dong} ${sub ? main+"-"+sub : main}` : dong;
  return "";
}

/* (5) 도로명 질의 후보/유틸 */
function buildRoadQueries_(raw) {
  const cleaned = _compactSpaces_(String(raw || "")
    .replace(/\b\d+\s*동\s*\d+\s*호?\b/gi, " ")
    .replace(/\b\d+\s*동\b/gi, " ")
    .replace(/\b\d+\s*호\b/gi, " ")
    .replace(/\(.*?\)\s*/g, " "));

  const reRoadBlock = /([가-힣0-9]+(?:대로|로|길|번길)\s*\d+(?:-\d+)?)/g;
  const blocks = [];
  let m; while ((m = reRoadBlock.exec(cleaned)) !== null) blocks.push({ text: m[1], idx: m.index });
  const last = blocks.length ? blocks[blocks.length - 1] : null;

  const { si, gu } = getRegionTokens_(cleaned);
  const qAll = cleaned;
  const qTrimTail = last ? _compactSpaces_(cleaned.slice(0, last.idx + last.text.length)) : cleaned;

  const rb = last ? last.text : "";
  const rbVars = roadBlockVariants_(rb);

  const conservative = rbVars.map(v => _compactSpaces_(((gu||"") + " " + v).trim()))
    .concat(rbVars.map(v => _compactSpaces_(((si||"") + " " + (gu||"") + " " + v).trim())));

  return _dedupe_([qAll, qTrimTail].concat(conservative)).filter(Boolean);
}
function roadBlockVariants_(rb) {
  if (!rb) return [];
  const out = new Set([rb]);
  out.add(rb.replace(/(로)(\d+)/, "$1 $2"));
  out.add(rb.replace(/(번길|길)(\s*)(\d+)/, "$1 $3"));
  out.add(rb.replace(/(\d+)가길/, "$1길"));
  out.add(rb.replace(/(로)\s*(\d+)가길/, "$1 $2길"));
  return Array.from(out);
}
function getRegionTokens_(s) {
  const si = (s.match(/([가-힣]+(?:특별|광역)?시)/) || [])[1] || (s.match(/([가-힣]+시)/) || [])[1] || "";
  const gu = (s.match(/([가-힣]+구)/) || [])[1] || "";
  const dong = (s.match(/([가-힣0-9]+동)/) || [])[1] || "";
  const siNorm = si.replace(/^서울시$/, "서울특별시");
  return { si: siNorm || si, gu, dong };
}
function _compactSpaces_(t){ return String(t).replace(/\s+/g, " ").trim(); }
function _dedupe_(arr){ const s=new Set(); const out=[]; for(const x of arr){ const y=(x||"").trim(); if(!y) continue; if(!s.has(y)){ s.add(y); out.push(y); } } return out; }

/* (6) 문자열 파서/정규화 */
function extractDongJibeonFromAddressDoc_(doc) {
  if (!doc) return "";
  const adr = doc.address || {};
  const dong = normalizeAdminToLegalDong_(adr.region_3depth_name || "");
  const main = String(adr.main_address_no || "").trim();
  const sub  = String(adr.sub_address_no  || "").trim();
  if (dong) return main ? `${dong} ${sub ? main+"-"+sub : main}` : dong;
  return "";
}
function parseDongJibeonFromString_(s) {
  if (!s) return "";
  const str = String(s);
  const trySeg = (seg) => {
    let m = String(seg).match(/([가-힣][가-힣0-9A-Za-z]*동)\s+(산)?\s*(\d{1,5}(?:-\d{1,5})?)/);
    if (m) return `${normalizeAdminToLegalDong_(m[1])} ${ (m[2]?'산 ':'') + m[3] }`;
    m = String(seg).match(/([가-힣][가-힣0-9A-Za-z]*동)\b/);
    if (m) return normalizeAdminToLegalDong_(m[1]);
    return "";
  };
  const parens = str.match(/\([^)]+\)/g) || [];
  for (const p of parens) {
    const seg = p.replace(/[()]/g, " ").replace(/\s+/g, " ").trim();
    const r = trySeg(seg);
    if (r) return r;
  }
  return trySeg(str);
}
function hasLot_(label){ return /\d/.test(String(label||"")); }
function extractComplexName_(s) {
  if (!s) return "";
  let t = String(s);

  // 괄호 내용 제거
  t = t.replace(/\(.*?\)/g, " ");

  // 세대/호수/층수 등만 제거 (단지명에 붙은 숫자는 보존!)
  t = t.replace(/\b\d+\s*동\s*\d+\s*호?\b/gi, " ");
  t = t.replace(/\b\d+\s*동\b/gi, " ");
  t = t.replace(/\b\d+\s*호\b/gi, " ");
  t = t.replace(/\b지하\s*\d+\s*층?\b/gi, " ");
  t = t.replace(/\b\d+\s*층\b/gi, " ");

  // 공백 정리
  t = t.replace(/\s+/g, " ").trim();
  return t;
}
function normalizeAdminToLegalDong_(dong){
  if (!dong) return "";
  return String(dong).trim().replace(/제\d+동$/, "동"); // 면목제5동 → 면목동
}

function normForGeocodeInput_(s) { return String(s || "").trim(); }

/* (8) 좌표 셀 형식 & 주소→좌표 캐시 */
function ensureLatLonFormats_(sh, rowStart, n) {
  const latRange = sh.getRange(rowStart, CFG.COL_LAT, n, 1);
  const lonRange = sh.getRange(rowStart, CFG.COL_LON, n, 1);
  latRange.setNumberFormat("0.########");
  lonRange.setNumberFormat("0.########");
  try { latRange.clearDataValidations(); } catch (_) {}
  try { lonRange.clearDataValidations(); } catch (_) {}
}
function getCached_(addr) {
  try {
    const cache = CacheService.getScriptCache();
    const hit = cache.get(CFG.GEO.KEY_PREFIX + addr);
    if (hit) return JSON.parse(hit);
  } catch (_) {}
  try {
    const props = PropertiesService.getScriptProperties();
    const val = props.getProperty(CFG.GEO.KEY_PREFIX + addr);
    return val ? JSON.parse(val) : null;
  } catch (_) { return null; }
}
function setCached_(addr, result) {
  const str = JSON.stringify(result);
  try { CacheService.getScriptCache().put(CFG.GEO.KEY_PREFIX + addr, str, CFG.GEO.TTL); } catch (_) {}
  try { PropertiesService.getScriptProperties().setProperty(CFG.GEO.KEY_PREFIX + addr, str); } catch (_) {}
}

/* (9) 공용 유틸 */
function isTargetSheet_(sh) {
  return CFG.TARGET_SHEETS && CFG.TARGET_SHEETS.indexOf(sh.getName()) !== -1;
}
function _findHitPoly_(lat, lon, rows) {
  if (!rows || !rows.length) return null;
  for (const r of rows) {
    if (r.bbox) {
      if (lat < r.bbox.minLat || lat > r.bbox.maxLat) continue;
      if (lon < r.bbox.minLon || lon > r.bbox.maxLon) continue;
    }
    if (typeof GEOM_pointInPoly === 'function' ? GEOM_pointInPoly(lat, lon, r.vertices)
        : (typeof pointInPoly_ === 'function' && pointInPoly_(lat, lon, r.vertices))) {
      return r;
    }
  }
  return null;
}
function clearDevCaches_() {
  try { if (typeof MOA_CACHE_CLEAR  === 'function') MOA_CACHE_CLEAR(); } catch(_) {}
  try { if (typeof FAST_CACHE_CLEAR === 'function') FAST_CACHE_CLEAR(); } catch(_) {}
}
function normAddr_(s) {
  return String(s || "").replace(/\s+/g, " ").replace(/\(.*?\)/g, "").trim();
}

/** 카카오 주소검색으로 좌표 1건 선택(구/동 힌트 우선) */
function kakaoAddressToCoords_(q, regionHint) {
  const urlExact   = "https://dapi.kakao.com/v2/local/search/address.json?size=5&analyze_type=exact&query=" + encodeURIComponent(q);
  const urlSimilar = "https://dapi.kakao.com/v2/local/search/address.json?size=5&analyze_type=similar&query=" + encodeURIComponent(q);

  function pickFromDocs(docs) {
    if (!Array.isArray(docs) || !docs.length) return null;
    const si  = regionHint && regionHint.si  || "";
    const gu  = regionHint && regionHint.gu  || "";
    const dong= regionHint && regionHint.dong|| "";

    const hasHint = (s) => {
      const t = String(s || "");
      return (gu && t.includes(gu)) || (dong && t.includes(dong)) || (si && t.includes(si));
    };

    // 1) 구/동(또는 시) 힌트 포함 우선
    let pref = docs.find(d =>
      hasHint(d.address_name) ||
      hasHint(d.road_address && d.road_address.address_name)
    );
    if (!pref) pref = docs[0];

    const x = Number(pref && pref.x), y = Number(pref && pref.y);
    if (isFinite(x) && isFinite(y)) return { lat: y, lon: x };
    return null;
  }

  const d1 = kakaoFetch_(urlExact);
  const hit1 = d1 && pickFromDocs(d1.documents);
  if (hit1) return hit1;

  const d2 = kakaoFetch_(urlSimilar);
  const hit2 = d2 && pickFromDocs(d2.documents);
  return hit2 || null;
}

/** 원문 → (도로명 후보→주소검색) → (지번 라벨→주소검색) → (키워드→좌표) */
function kakaoGeocodeToCoords_(raw) {
  const s = normForGeocodeInput_(raw);
  if (!s) return null;
  const hint = getRegionTokens_(s);                    // si/gu/dong 힌트
  const roadQueries = buildRoadQueries_(s);            // 도로명 블록 후보들

  // (1) 도로명으로 직접 좌표 잡기
  for (const q of roadQueries) {
    const hit = kakaoAddressToCoords_(q, hint);
    if (hit) return hit;
  }

  // (2) 지번 라벨을 먼저 뽑고 → 그걸 주소검색으로 좌표화
  const jibunLabel = kakaoRoadToJibeon_TryList_(roadQueries) || kakaoKeywordToJibeon_(extractComplexName_(s) || s, s);
  if (jibunLabel) {
    // 힌트(시/구)를 붙여 주소검색 성공률 ↑
    const q1 = [hint.si, hint.gu, jibunLabel].filter(Boolean).join(" ");
    const q2 = [hint.gu, jibunLabel].filter(Boolean).join(" ");
    const q3 = jibunLabel;

    const cand = _dedupe_([q1, q2, q3]).filter(Boolean);
    for (const q of cand) {
      const hit = kakaoAddressToCoords_(q, hint);
      if (hit) return hit;
    }
  }

  // (3) 키워드(단지명 등)로 좌표 잡기  — 아파트/주거 카테고리 우선
  const kw = extractComplexName_(s) || s;
  const queries = _dedupe_([
    [hint.gu, kw].filter(Boolean).join(" "),
    [hint.si, hint.gu, kw].filter(Boolean).join(" "),
    [hint.dong, kw].filter(Boolean).join(" "),
    kw
  ]).filter(Boolean);

  for (const q of queries) {
    const url = "https://dapi.kakao.com/v2/local/search/keyword.json?size=5&query=" + encodeURIComponent(q);
    const data = kakaoFetch_(url);
    const docs = data && data.documents;
    if (!Array.isArray(docs) || !docs.length) continue;

    const hasHint = (s) => {
      const t = String(s || "");
      return (hint.gu && t.includes(hint.gu)) || (hint.dong && t.includes(hint.dong)) || (hint.si && t.includes(hint.si));
    };
    const isApt = (d) => /아파트/.test(d.place_name || "") || /아파트/.test(d.category_name || "");

    // ① 지역 힌트 + 아파트
    let pref = docs.find(d => hasHint(d.address_name) && isApt(d));
    // ② 지역 힌트 포함
    if (!pref) pref = docs.find(d => hasHint(d.address_name));
    // ③ 아파트 카테고리
    if (!pref) pref = docs.find(isApt);
    // ④ 첫 문서
    if (!pref) pref = docs[0];

    const x = Number(pref && pref.x), y = Number(pref && pref.y);
    if (isFinite(x) && isFinite(y)) return { lat: y, lon: x };
  }

  return null;
}

