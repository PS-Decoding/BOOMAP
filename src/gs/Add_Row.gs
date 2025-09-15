/***** Add_Row.gs *****/
// Config.gs의 설정 사용
const TEMPLATE_ROW = CFG.TEMPLATE_ROW; // 숨겨둔 템플릿 행 번호

function addRowsFromTemplate_1()   { addRowsFromTemplateBatch_(1); }
function addRowsFromTemplate_10()  { addRowsFromTemplateBatch_(10); }
function addRowsFromTemplate_50()  { addRowsFromTemplateBatch_(50); }
function addRowsFromTemplate_100() { addRowsFromTemplateBatch_(100); }

/**
 * 현재 활성 시트 맨 아래에 count줄 추가:
 *  - 템플릿(3행) 내용을 "그대로" 복사(서식/수식/검증 포함)
 *  - 단, A열(접수일)만 '오늘 날짜'로 값 덮어쓰기 (시각 00:00:00)
 *  - 메뉴 더블클릭 등 중복 실행 방지용 락 적용
 */
function addRowsFromTemplateBatch_(count) {
  if (!Number.isFinite(count) || count <= 0) return;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  // 동시 실행 방지 락
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    ss.toast("다른 추가 작업이 진행 중입니다", "대기 취소", 3);
    return;
  }

  try {
    const cols = sh.getLastColumn();
    const last = sh.getLastRow();

    // 템플릿 행 유효성 검사
    if (!Number.isFinite(TEMPLATE_ROW) || TEMPLATE_ROW < 1 || TEMPLATE_ROW > last) {
      throw new Error("템플릿 행이 없습니다");
    }

    // 1) 맨 아래에 count줄 삽입
    sh.insertRowsAfter(last, count);

    // 2) 템플릿 행 전체를 각 새 행으로 복사 (서식/수식/검증 그대로)
    const src = sh.getRange(TEMPLATE_ROW, 1, 1, cols);
    for (let i = 0; i < count; i++) {
      const dst = sh.getRange(last + 1 + i, 1, 1, cols);
      src.copyTo(dst, { contentsOnly: false });
    }

    // 3) A열(접수일)만 오늘 '날짜'로 덮어쓰기 (시각 제거)
    const firstNewRow = last + 1;
    const rngA = sh.getRange(firstNewRow, 1, count, 1);
    const today = new Date();
    const dateOnly = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    rngA.setValues(Array.from({ length: count }, () => [dateOnly]));
    rngA.setNumberFormat("yyyy. MM. dd");

    // UX
    sh.setActiveRange(sh.getRange(firstNewRow, 1, 1, cols));
    ss.toast(`템플릿에서 ${count}줄 추가 완료 (A열=오늘 날짜)`, "완료", 2);

  } catch (err) {
    ss.toast(`행 추가 중 오류: ${err && err.message ? err.message : err}`, "오류", 5);
    throw err;
  } finally {
    lock.releaseLock();
  }
}
