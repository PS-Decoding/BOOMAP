/***** Check_Dev.gs — MOA/FAST 로더 & 캐시 유틸 통합 *****/

/** 모아타운/신속통합 행 캐시 (메모리) */
var __moaRowsCache  = null;
var __fastRowsCache = null;

/** 범용: 2차원 값 범위 해시 (변경 감지) */
function _rangeHash_(vals) {
  const json  = JSON.stringify(vals);
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, json);
  return bytes.map(b => (b + 256).toString(16).slice(-2)).join("");
}

/** 문서 스코프 키 헬퍼 (캐시/속성 충돌 방지) */
function _docScopedKey_(suffix) {
  const docId = SpreadsheetApp.getActive().getId();
  return `${docId}:${suffix}`;
}
/** 캐시 키에 문서 ID 프리픽스 적용 */
function _scopedCacheKey_(rawKey) {
  return _docScopedKey_(rawKey);
}

/* ===================== MOA: 모아타운 ===================== */

function MOA_loadRows_() {
  if (__moaRowsCache) return __moaRowsCache;

  const cache   = CacheService.getScriptCache();
  const cKey    = _scopedCacheKey_(CFG.MOA.CACHE_KEY);
  const hit     = cache.get(cKey);
  const sh      = SpreadsheetApp.getActive().getSheetByName(CFG.MOA.SHEET);
  if (!sh) return [];

  const last = sh.getLastRow();
  if (last < CFG.MOA.START_ROW) return [];

  const num    = last - CFG.MOA.START_ROW + 1;
  const reps   = sh.getRange(CFG.MOA.START_ROW, CFG.MOA.COL_REP,   num, 1).getValues();
  const stages = sh.getRange(CFG.MOA.START_ROW, CFG.MOA.COL_STAGE, num, 1).getValues();
  const polys  = sh.getRange(CFG.MOA.START_ROW, CFG.MOA.COL_POLY,  num, 1).getValues();

  const props    = PropertiesService.getScriptProperties();
  const HASH_KEY = _docScopedKey_("MOA_HASH");
  const hash     = _rangeHash_([reps, stages, polys]);
  const lastHash = props.getProperty(HASH_KEY);

  if (hit && lastHash === hash) {
    try {
      __moaRowsCache = JSON.parse(hit);
      if (Array.isArray(__moaRowsCache)) return __moaRowsCache;
    } catch (_) {}
  }

  const out = [];
  for (let i = 0; i < num; i++) {
    const rep     = String(reps[i][0]   || "").trim();
    const stage   = String(stages[i][0] || "").trim();
    const polyStr = String(polys[i][0]  || "").trim();
    if (!polyStr) continue;

    const vertices = GEOM_parsePoly(polyStr);
    if (!vertices || vertices.length < 3) continue;

    out.push({ rep, stage, vertices, bbox: GEOM_bbox(vertices) });
  }

  __moaRowsCache = out;
  cache.put(cKey, JSON.stringify(out), CFG.MOA.CACHE_TTL);
  props.setProperty(HASH_KEY, hash);
  return out;
}

function MOA_CACHE_CLEAR() {
  __moaRowsCache = null;
  try {
    CacheService.getScriptCache().remove(_scopedCacheKey_(CFG.MOA.CACHE_KEY));
  } catch (_) {}
}

/* ===================== FAST: 신속통합 ===================== */

function FAST_loadRows_() {
  if (__fastRowsCache) return __fastRowsCache;

  const cache   = CacheService.getScriptCache();
  const cKey    = _scopedCacheKey_(CFG.FAST.CACHE_KEY);
  const hit     = cache.get(cKey);

  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.FAST.SHEET);
  if (!sh) return [];

  const last = sh.getLastRow();
  if (last < CFG.FAST.START_ROW) return [];

  const num    = last - CFG.FAST.START_ROW + 1;
  const reps   = sh.getRange(CFG.FAST.START_ROW, CFG.FAST.COL_REP,   num, 1).getValues();
  const stages = sh.getRange(CFG.FAST.START_ROW, CFG.FAST.COL_STAGE, num, 1).getValues();
  const polys  = sh.getRange(CFG.FAST.START_ROW, CFG.FAST.COL_POLY,  num, 1).getValues();

  const props    = PropertiesService.getScriptProperties();
  const HASH_KEY = _docScopedKey_("FAST_HASH");
  const hash     = _rangeHash_([reps, stages, polys]);
  const lastHash = props.getProperty(HASH_KEY);

  if (hit && lastHash === hash) {
    try {
      __fastRowsCache = JSON.parse(hit);
      if (Array.isArray(__fastRowsCache)) return __fastRowsCache;
    } catch (_) {}
  }

  const out = [];
  for (let i = 0; i < num; i++) {
    const rep     = String(reps[i][0]   || "").trim();
    const stage   = String(stages[i][0] || "").trim();
    const polyStr = String(polys[i][0]  || "").trim();
    if (!polyStr) continue;

    const vertices = GEOM_parsePoly(polyStr);
    if (!vertices || vertices.length < 3) continue;

    out.push({ rep, stage, vertices, bbox: GEOM_bbox(vertices) });
  }

  __fastRowsCache = out;
  cache.put(cKey, JSON.stringify(out), CFG.FAST.CACHE_TTL);
  props.setProperty(HASH_KEY, hash);
  return out;
}

function FAST_CACHE_CLEAR() {
  __fastRowsCache = null;
  try {
    CacheService.getScriptCache().remove(_scopedCacheKey_(CFG.FAST.CACHE_KEY));
  } catch (_) {}
}

/* ============== 통합 캐시 클리어 헬퍼 ============== */
/** 다른 스크립트에서 호출 중이면 그대로 동작하도록 이름 유지 */
function clearDevCaches_() {
  try { MOA_CACHE_CLEAR();  } catch (_) {}
  try { FAST_CACHE_CLEAR(); } catch (_) {}
}
