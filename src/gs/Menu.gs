/*** Menu.gs — 메뉴 + 통합 갱신/트리거/디버그 ***/

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  addAddRowMenu_(ui);
  addToolsMenu_(ui);
  addAdCopyMenu_(ui);
  addMapMenu_(ui);
}

/** ＋ 행 추가 */
function addAddRowMenu_(ui) {
  ui.createMenu("＋ 행 추가")
    .addItem("1줄 추가",   "addRowsFromTemplate_1")
    .addItem("10줄 추가",  "addRowsFromTemplate_10")
    .addItem("50줄 추가",  "addRowsFromTemplate_50")
    .addItem("100줄 추가", "addRowsFromTemplate_100")
    .addToUi();
}

/** 🛠 도구 */
function addToolsMenu_(ui) {
  ui.createMenu("🛠 도구")
    .addItem("선택 행 갱신",   "TOOL_updateSelectionCombined")
    .addItem("선택 탭 갱신",   "TOOL_updateActiveSheetCombined")
    .addItem("전체 탭 갱신",   "TOOL_updateAllSheetsCombined")
    .addSeparator()
    .addItem("자동 갱신 트리거 설치", "installAutoUpdateTriggers")
    .addItem("자동 갱신 트리거 제거", "removeAutoUpdateTriggers")
    .addSeparator()
    .addItem("디버그(선택 행)", "TOOL_debugCombinedHere")
    .addToUi();
}

/** ✍️ 문구 작성 */
function addAdCopyMenu_(ui) {
  ui.createMenu("✍️ 문구 작성")
    .addItem("광고 문구 작성", "makeAdCopyFromActiveRow")   // ← 기존 함수 그대로 사용
    .addItem("촬영 문구 작성", "makeShootCopyFromActiveRow") // ← 신규
    .addToUi();
}

/** 🗺 매물 지도 */
function addMapMenu_(ui) {
  ui.createMenu("🗺️ BOOMAP")
    .addItem("BOOMAP 실행", "showMap")
    .addToUi();
}

/* ============== 통합 갱신 (지오코딩 → H라벨 → 판정) ============== */

function TOOL_updateSelectionCombined() {
  const START = CFG.ROW_START;
  clearDevCaches_();

  const ss = SpreadsheetApp.getActive();
  const sh = SpreadsheetApp.getActiveSheet();
  if (!isTargetSheet_(sh)) { SpreadsheetApp.getUi().alert("대상 탭이 아닙니다."); return; }

  const r = sh.getActiveRange(); if (!r) return;
  const rS = Math.max(r.getRow(), START);
  const rE = r.getLastRow();
  if (rS > rE) return;

  geocodeRowsUnique_(sh, rS, rE, /*onlyEmpty=*/false);
  updateJibeonLabel_(sh, rS, rE);

  const rows = []; for (let i = rS; i <= rE; i++) rows.push(i);
  JUDGE_updateSubset_(sh, rows);

  SpreadsheetApp.flush();
  ss.toast(`선택 행 갱신 완료: ${sh.getName()} ${rS}~${rE}행`, "갱신 완료", 3);
}

function TOOL_updateActiveSheetCombined() {
  const START = CFG.ROW_START;
  const STEP  = (CFG.BLOCK_ROWS || 400);
  clearDevCaches_();

  const ss = SpreadsheetApp.getActive();
  const sh = SpreadsheetApp.getActiveSheet();
  if (!isTargetSheet_(sh)) { SpreadsheetApp.getUi().alert("대상 탭이 아닙니다."); return; }

  const last = sh.getLastRow();
  if (last < START) { ss.toast(`갱신 대상 없음: ${sh.getName()}`, "정보", 3); return; }

  for (let s = START; s <= last; s += STEP) {
    const e = Math.min(s + STEP - 1, last);
    if (e < s) continue;

    geocodeRowsUnique_(sh, s, e, /*onlyEmpty=*/false);
    updateJibeonLabel_(sh, s, e);

    const rows = []; for (let i = s; i <= e; i++) rows.push(i);
    JUDGE_updateSubset_(sh, rows);
    SpreadsheetApp.flush();
  }
  ss.toast(`선택 탭 갱신 완료: ${sh.getName()}`, "갱신 완료", 3);
}

function TOOL_updateAllSheetsCombined() {
  const START = CFG.ROW_START;
  const STEP  = (CFG.BLOCK_ROWS || 400);
  clearDevCaches_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let countTabs = 0;

  for (const name of CFG.TARGET_SHEETS) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;
    const last = sh.getLastRow();
    if (last < START) continue;

    for (let s = START; s <= last; s += STEP) {
      const e = Math.min(s + STEP - 1, last);
      geocodeRowsUnique_(sh, s, e, /*onlyEmpty=*/true); // 전체는 빈 좌표만 권장
      updateJibeonLabel_(sh, s, e);

      const rows = []; for (let i = s; i <= e; i++) rows.push(i);
      JUDGE_updateSubset_(sh, rows);
      SpreadsheetApp.flush();
    }
    countTabs++;
  }

  SpreadsheetApp.getActive().toast(
    countTabs ? `전체 탭 갱신 완료 (${countTabs}개 탭)` : "갱신할 탭이 없습니다",
    "갱신 완료", 3
  );
}

/* ============== 트리거 ============== */
function installAutoUpdateTriggers() {
  removeAutoUpdateTriggers();
  const id = SpreadsheetApp.getActive().getId();
  ScriptApp.newTrigger("onEditHandler").forSpreadsheet(id).onEdit().create();
  ScriptApp.newTrigger("onEditJudge_").forSpreadsheet(id).onEdit().create();
  SpreadsheetApp.getUi().alert("자동 갱신 트리거(지오코딩/판정)가 설치되었습니다.");
}
function removeAutoUpdateTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (const t of triggers) {
    const fn = t.getHandlerFunction && t.getHandlerFunction();
    if (fn === "onEditHandler" || fn === "onEditJudge_") {
      ScriptApp.deleteTrigger(t); removed++;
    }
  }
  SpreadsheetApp.getUi().alert(removed ? "자동 갱신 트리거 제거 완료" : "제거할 트리거가 없습니다.");
}
