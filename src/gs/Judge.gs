/*** Judge.gs — F/G 판정 유틸 (개선판) ***/

// onEdit(좌표 변경)에 반응하여 해당 행만 판정
function onEditJudge_(e) {
  try {
    if (!e || !e.range) return; // 가드냥
    // 값 변경이 아니면 무시 (서식 붙여넣기 등)냥
    if ('value' in e && 'oldValue' in e && e.value === e.oldValue) return;

    const sh = e.range.getSheet();
    if (!sh || CFG.TARGET_SHEETS.indexOf(sh.getName()) === -1) return;

    const cS = e.range.getColumn(), cE = e.range.getLastColumn();
    const touchesLatLon = !(cE < CFG.COL_LAT || cS > CFG.COL_LON); // D~E와 교차?냥
    if (!touchesLatLon) return;

    const rS = Math.max(e.range.getRow(), CFG.ROW_START);
    const rE = e.range.getLastRow();
    if (rS > rE) return;

    // F/G 체크박스 보장냥
    ensureFlagCheckboxColumns_(sh);

    // 대상 행 배열 구성냥
    const rows = [];
    for (let i = rS; i <= rE; i++) rows.push(i);

    JUDGE_updateSubset_(sh, rows);
  } catch (_) {}
}

// 지오코딩 직후 특정 인덱스(empties/needIdx)만 판정
function JUDGE_updateRowsAfterGeocode_(sh, rowStart, empties, needIdx) {
  const set = new Set();
  for (const i of empties || []) set.add(rowStart + i);
  for (const i of needIdx || []) set.add(rowStart + i);
  const rows = Array.from(set.values()).sort((a, b) => a - b);
  if (!rows.length) return;
  ensureFlagCheckboxColumns_(sh);
  JUDGE_updateSubset_(sh, rows);
}

// 핵심: 주어진 행들만 F/G 계산 — 비연속 행도 안전 처리냥
function JUDGE_updateSubset_(sh, rows) {
  if (!rows || !rows.length) return;

  const latCol = CFG.COL_LAT, lonCol = CFG.COL_LON;
  const fCol = CFG.COL_MOA_FLAG, gCol = CFG.COL_FAST_FLAG;

  const moaRows  = (typeof MOA_loadRows_  === 'function') ? MOA_loadRows_()  : [];
  const fastRows = (typeof FAST_loadRows_ === 'function') ? FAST_loadRows_() : [];

  // 연속 구간으로 묶어서 배치 처리냥
  const ranges = _toContiguousRanges_(rows);

  const writeF = [];
  const writeG = [];
  const writeRows = [];

  for (const [rs, re] of ranges) {
    const len = re - rs + 1;
    const vals = sh.getRange(rs, latCol, len, 2).getValues(); // D:E냥
    for (let i = 0; i < len; i++) {
      const lat = Number(vals[i][0]);
      const lon = Number(vals[i][1]);
      let hitM = false, hitF = false;
      if (isFinite(lat) && isFinite(lon)) {
        const hm = _findHitPoly_(lat, lon, moaRows);
        const hf = _findHitPoly_(lat, lon, fastRows);
        hitM = !!hm; hitF = !!hf;
      }
      writeRows.push(rs + i);
      writeF.push([hitM]);
      writeG.push([hitF]);
    }
  }

  // 다시 연속 구간으로 묶어 나눠 쓰기냥
  _writeBooleansByRows_(sh, fCol, writeRows, writeF);
  _writeBooleansByRows_(sh, gCol, writeRows, writeG);
}

/* ===== 헬퍼들 ===== */

// F/G 칼럼을 체크박스로 세팅 (1회)냥
function ensureFlagCheckboxColumns_(sh) {
  const fCol = CFG.COL_MOA_FLAG, gCol = CFG.COL_FAST_FLAG;
  const start = CFG.ROW_START;
  const last = sh.getMaxRows(); // 충분히 크게 적용냥

  const dv = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  try { sh.getRange(start, fCol, last - start + 1, 1).setDataValidation(dv); } catch (_) {}
  try { sh.getRange(start, gCol, last - start + 1, 1).setDataValidation(dv); } catch (_) {}
}

// 비정렬/비연속 rows → 정렬된 연속 구간 배열 [[s,e], ...]로 변환냥
function _toContiguousRanges_(rows) {
  const a = Array.from(new Set(rows)).filter(n => Number.isInteger(n)).sort((x, y) => x - y);
  if (!a.length) return [];
  const out = [];
  let s = a[0], p = a[0];
  for (let i = 1; i < a.length; i++) {
    if (a[i] === p + 1) { p = a[i]; continue; }
    out.push([s, p]); s = p = a[i];
  }
  out.push([s, p]);
  return out;
}

// 행 인덱스 집합에 대해 같은 순서로 boolean 2차원 배열을 해당 컬럼에 배치 쓰기냥
function _writeBooleansByRows_(sh, col, rows, data2d) {
  if (!rows.length) return;
  const ranges = _toContiguousRanges_(rows);
  let cursor = 0;
  for (const [rs, re] of ranges) {
    const len = re - rs + 1;
    const slice = data2d.slice(cursor, cursor + len);
    sh.getRange(rs, col, len, 1).setValues(slice);
    cursor += len;
  }
}
