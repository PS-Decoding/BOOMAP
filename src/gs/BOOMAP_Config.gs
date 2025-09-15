/** =========================
 * BOOMAP_Config.gs — 데이터/오버레이 제공 + Kakao JS Key + HTML 열기
 * ========================= */

/* =========[공통/키 로더]========= */
// config.gs 어딘가에 const KAKAO_JS_KEY = '...' 로 선언되어 있어야 함
function getKakaoKey() {
  try {
    // eslint-disable-next-line no-undef
    if (typeof KAKAO_JS_KEY !== 'undefined' && KAKAO_JS_KEY) return KAKAO_JS_KEY;
  } catch (e) {}
  try {
    const props = PropertiesService.getScriptProperties();
    const fromProp = props.getProperty('KAKAO_JS_KEY');
    return fromProp ? String(fromProp).trim() : '';
  } catch (e) {
    return '';
  }
}

/* =========[시트 유틸]========= */
function parseFloor_(floorRaw) {
  if (floorRaw === null || floorRaw === undefined) return '';
  if (typeof floorRaw === 'number' && isFinite(floorRaw)) return floorRaw;

  var s = String(floorRaw).trim();
  if (!s) return '';
  if (/^-?\d+$/.test(s)) return Number(s);

  var bh = s.match(/지하\s*(\d+)/i);
  if (bh) return -Number(bh[1]);

  var b2 = s.match(/^\s*B\s*[-\s]*(\d+)/i);
  if (b2) return -Number(b2[1]);

  var firstPart = s.split(/[\/\\]/)[0];
  var m = String(firstPart).match(/-?\d+/);
  return m ? Number(m[0]) : '';
}

/** RichTextValue에서 첫 링크 안전 추출(부분 링크 대응) */
function getFirstLinkUrlFromRich_(rtv) {
  try {
    if (!rtv) return null;
    if (rtv.getRuns) {
      const runs = rtv.getRuns();
      for (const run of runs) {
        const u = run.getLinkUrl();
        if (u) return u;
      }
    }
    if (rtv.getLinkUrl) return rtv.getLinkUrl();
  } catch (_) {}
  return null;
}

/* =========[Drive 유틸 + 캐시]========= */
/** URL/ID → Drive ID 추출 */
function DRIVE_extractId(urlOrId) {
  if (!urlOrId) return null;
  const s = String(urlOrId);
  let m = s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/); if (m) return m[1];
  m = s.match(/\/folders\/([a-zA-Z0-9_-]{10,})(?:[/?#]|$)/); if (m) return m[1];
  m = s.match(/\/file\/d\/([a-zA-Z0-9_-]{10,})(?:[/?#]|$)/);  if (m) return m[1];
  m = s.match(/[-\w]{25,}/); return m ? m[0] : null;
}

/** ID가 폴더인지 확인 + 파일이면 첫 부모 폴더로 보정 (Drive v3 전용 필드) */
function _ensureFolderId(id) {
  const res = Drive.Files.get(id, {
    fields: 'id,mimeType,parents',
    supportsAllDrives: true,   // v3
    supportsTeamDrives: true   // v2 호환 플래그
  });
  if (!res) return null;
  if (res.mimeType === 'application/vnd.google-apps.folder') return res.id;

  const parents = res.parents || [];
  if (!parents.length) return null;

  const first = parents[0];
  const parentId = (typeof first === 'string') ? first : (first.id || first);
  return parentId || null;
}

/** 폴더 modifiedTime만 가져오기(캐시 키 신선도용) */
function _getFolderModifiedTime_(id){
  try{
    const r = Drive.Files.get(id, {
      fields: 'modifiedTime',
      supportsAllDrives: true,
      supportsTeamDrives: true
    });
    return r && r.modifiedTime ? String(r.modifiedTime) : '';
  }catch(e){ return ''; }
}

/** 이미지 리스트 캐시 TTL(초) */
var IMAGE_LIST_CACHE_TTL = 300; // 5분

/**
 * 폴더(ID/URL) 안의 이미지 파일 나열 (페이징 + 공유드라이브 + 캐시) — v3 전용
 * 반환: [{id,name,mimeType,url,thumb,viewUrl}]
 * @param {string} urlOrId
 * @param {boolean=} forceRefresh  true면 캐시 무시하고 신규 조회
 */
function listImagesInFolder(urlOrId, forceRefresh) {
  const rawId = DRIVE_extractId(urlOrId);
  if (!rawId) throw new Error('유효하지 않은 폴더/파일 URL 또는 ID입니다: ' + urlOrId);

  const folderId = _ensureFolderId(rawId);
  if (!folderId) throw new Error('폴더를 찾을 수 없습니다(파일이거나 접근권한 없음): ' + rawId);

  // 캐시 키 = 폴더ID + modifiedTime
  const mtime = _getFolderModifiedTime_(folderId) || '';
  const cacheKey = 'imglist:' + folderId + ':' + mtime;
  const cache = CacheService.getScriptCache();

  if (!forceRefresh) {
    const cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch(_){}
    }
  }

  // 캐시 미스 → 실제 조회
  const query = `'${folderId}' in parents and trashed = false and mimeType contains 'image/'`;
  const out = [];
  let pageToken = null;

  do {
    const params = {
      q: query,
      pageSize: 200,
      pageToken: pageToken,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
      fields: 'files(id,name,mimeType,webViewLink,thumbnailLink),nextPageToken'
    };

    const res = Drive.Files.list(params);
    const files = (res && res.files) || [];
    pageToken = res && res.nextPageToken;

    for (const f of files) {
      const id    = f.id;
      const name  = f.name || '';
      const mime  = f.mimeType || '';
      const view  = f.webViewLink || '';
      const thumb = f.thumbnailLink ? f.thumbnailLink.replace(/=s\d+$/, '=s800') : '';

      out.push({
        id,
        name,
        mimeType: mime,
        url: 'https://drive.google.com/uc?export=view&id=' + id, // 미리보기
        thumb: thumb || ('https://drive.google.com/thumbnail?id=' + id),
        viewUrl: view
      });
    }
  } while (pageToken);

  // 캐시 저장
  try { cache.put(cacheKey, JSON.stringify(out), IMAGE_LIST_CACHE_TTL); } catch(_){}

  return out;
}

/** (옵션) 시트 C열 링크로 바로 호출 */
function listImagesForListing(folderUrlInSheet) {
  if (!folderUrlInSheet) return [];
  return listImagesInFolder(folderUrlInSheet);
}

/** 프리패치(워밍): 보이는 폴더 몇 개만 미리 캐시에 채우기 */
function warmImageCaches(folderIds){
  if(!Array.isArray(folderIds) || folderIds.length === 0) return 0;
  var count = 0;
  const MAX = 20; // 과도 호출 방지
  for (var i=0; i<folderIds.length && i<MAX; i++){
    var fid = folderIds[i];
    if(!fid) continue;
    try{
      listImagesInFolder(fid, false); // 캐시 미스면 채워짐
      count++;
      Utilities.sleep(80); // Drive API 쿼터 완화
    }catch(e){
      // 권한/빈 폴더 등은 무시
    }
  }
  return count;
}

/* =========[시트 → 매물/오버레이 로드]========= */
function getListings_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targets = (typeof CFG !== 'undefined' && CFG && CFG.TARGET_SHEETS) || [];
  if (!targets.length) return [];

  var COL = {
    type: 2, addr: 3, lat: 4, lng: 5,
    deposit: 9, rent: 10, loan: 14, area: 15, floor: 16,
    rooms: 18, baths: 19, parking: 26, pets: 29, houseType: 35
  };
  var START = (typeof CFG !== 'undefined' && CFG && CFG.ROW_START) ? CFG.ROW_START : 2;

  var rows = [];
  for (var i = 0; i < targets.length; i++) {
    var name = targets[i];
    var sh = ss.getSheetByName(name);
    if (!sh) continue;

    var lastRow = sh.getLastRow();
    if (lastRow < START) continue;

    var width = sh.getLastColumn();
    var range = sh.getRange(START, 1, lastRow - START + 1, width);
    var data = range.getValues();
    var rich = range.getRichTextValues();

    for (var r = 0; r < data.length; r++) {
      var row = data[r];
      var richRow = rich[r];

      var type     = row[COL.type - 1];
      var addr     = row[COL.addr - 1];
      var lat      = row[COL.lat - 1];
      var lng      = row[COL.lng - 1];
      var deposit  = row[COL.deposit - 1];
      var rent     = row[COL.rent - 1];
      var loan     = row[COL.loan - 1];
      var area     = row[COL.area - 1];
      var floorRaw = row[COL.floor - 1];
      var rooms    = row[COL.rooms - 1];
      var baths    = row[COL.baths - 1];
      var parking  = row[COL.parking - 1];
      var pets     = row[COL.pets - 1];
      var house    = row[COL.houseType - 1];

      var addrStr = String(addr || '').trim();
      if (!addrStr) continue;

      var latNum = Number(lat), lngNum = Number(lng);
      if (!isFinite(latNum) || !isFinite(lngNum)) continue;
      if (latNum === 0 && lngNum === 0) continue;

      var floor = parseFloor_(floorRaw);

      // C열 하이퍼링크 → ID 추출(부분 링크 OK)
      var folderUrl = getFirstLinkUrlFromRich_(richRow[COL.addr - 1]);
      var photosFolderId = folderUrl ? DRIVE_extractId(folderUrl) : "";

      rows.push({
        sheet: name,
        type: String(type || '').trim(),
        address: addrStr,
        lat: latNum,
        lng: lngNum,
        depositOrPrice: deposit,
        rent: rent,
        loan: loan,
        area: area,
        floor: floor,
        rooms: rooms,
        baths: baths,
        parking: parking,
        pets: pets,
        houseType: String(house || '').trim(),
        photosFolderId: photosFolderId
      });
    }
  }
  return rows;
}

function parseVertices_(raw) {
  if (raw == null) return [];
  var s = String(raw).trim();
  if (!s) return [];
  try {
    var j = JSON.parse(s);
    if (Array.isArray(j)) {
      return j.map(function (pt) {
        if (Array.isArray(pt) && pt.length >= 2) return [Number(pt[0]), Number(pt[1])];
        if (pt && typeof pt === 'object') {
          var la = pt.lat || pt.latitude;
          var lo = pt.lng || pt.lon || pt.longitude || pt.long;
          return [Number(la), Number(lo)];
        }
        return null;
      }).filter(function (v) { return v && isFinite(v[0]) && isFinite(v[1]); });
    }
  } catch (e) { }
  var parts = s.split(/[\n;|]+/).map(function (t) { return t.trim(); }).filter(Boolean);
  var verts = [];
  for (var i = 0; i < parts.length; i++) {
    var m = parts[i].split(/[,\s]+/).map(function (x) { return x.trim(); }).filter(Boolean);
    if (m.length >= 2) {
      var lat = Number(m[0]), lng = Number(m[1]);
      if (isFinite(lat) && isFinite(lng)) verts.push([lat, lng]);
    }
  }
  return verts;
}

function loadMoaOverlay_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('💚모아타운');
  if (!sh) return [];
  var START = (typeof CFG !== 'undefined' && CFG && CFG.ROW_START) ? CFG.ROW_START : 2;
  var last = sh.getLastRow();
  if (last < START) return [];
  var rng = sh.getRange(START, 1, last - START + 1, sh.getLastColumn()).getValues();
  var out = [];
  for (var i = 0; i < rng.length; i++) {
    var row = rng[i];
    var name = row[1];
    var stage = row[5];
    var vertsRaw = row[21];
    var vertices = parseVertices_(vertsRaw);
    if (vertices.length >= 3) out.push({ name: String(name || '').trim(), stage: String(stage || '').trim(), vertices: vertices });
  }
  return out;
}

function loadFastOverlay_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('❤️신속통합');
  if (!sh) return [];
  var START = (typeof CFG !== 'undefined' && CFG && CFG.ROW_START) ? CFG.ROW_START : 2;
  var last = sh.getLastRow();
  if (last < START) return [];
  var rng = sh.getRange(START, 1, last - START + 1, sh.getLastColumn()).getValues();
  var out = [];
  for (var i = 0; i < rng.length; i++) {
    var row = rng[i];
    var name = row[1];
    var stage = row[5];
    var vertsRaw = row[7];
    var vertices = parseVertices_(vertsRaw);
    if (vertices.length >= 3) out.push({ name: String(name || '').trim(), stage: String(stage || '').trim(), vertices: vertices });
  }
  return out;
}

function getOverlays_() { return { moa: loadMoaOverlay_(), fast: loadFastOverlay_() }; }

/** 클라이언트 호출용 */
function fetchListings() { return getListings_(); }
function fetchOverlays() { return getOverlays_(); }

/** HTML 열기 (파일명 정확히 BooMap) */
function showMap() {
  var html = HtmlService.createHtmlOutputFromFile('BooMap').setWidth(1280).setHeight(860);
  SpreadsheetApp.getUi().showModalDialog(html, 'BOOMAP');
}

/** 웹 앱 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('BooMap')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}