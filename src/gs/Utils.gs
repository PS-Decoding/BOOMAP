/***** Utils.gs *****/

/** 공용: 선택 행에서 광고문구 본문(거래/가격/대출/주차/반려/옵션/입주 등) 생성 */
function buildAdCopyTextFromRow_(sheet, row) {
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(2, 1, 1, lastCol).getDisplayValues()[0];
  const value  = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];

  const idx = (col) => { let n=0; for (const ch of col.toUpperCase()) n = n*26 + (ch.charCodeAt(0)-64); return n-1; };
  const h = (col) => String(header[idx(col)] || '').trim();
  const v = (col) => String(value[idx(col)]  || '').trim();
  const has = s => s != null && String(s).trim() !== '';

  const lines = [];

  // 소재지
  lines.push(sectionTitle_('소재지 정보'));
  lines.push(line_(h('C'), v('C')));

  // 주택 유형
  lines.push('');
  lines.push(sectionTitle_('주택 유형'));
  lines.push(line_(h('AH'), v('AH')));

  // 거래 정보
  lines.push('');
  lines.push(sectionTitle_('거래 정보'));
  const tradeType = v('B'); // 매매/전세/월세 등
  lines.push(line_(h('B'), tradeType));

  // 가격 정보 (규칙: 1,000만↑ 억 단위, 억 이상은 '##억 ##만 원' 고정)
  lines.push('');
  lines.push(sectionTitle_('가격 정보'));
  const priceMain = fmtPriceKRW_(v('I'));
  lines.push(line_(h('I'), priceMain));
  if (has(v('J'))) { // 보증금 등
    lines.push(line_(h('J'), fmtPriceKRW_(v('J'))));
  }
  // 관리비
  lines.push(line_(h('K'), has(v('K')) ? v('K') : '확인필요'));

  // 대출 (규칙: ✅일 때만 고지 문구 첨부)
  lines.push('');
  lines.push(sectionTitle_('대출 관련'));
  const loanStr = v('N'); // 대출여부
  lines.push(line_(h('M'), v('M'))); // 융자여부 원문
  lines.push(line_(h('N'), loanStr));
  if (/대출✅/.test(loanStr)) {
    lines.push('📢 대출 세부 한도·가능 여부는 금융기관 또는 세무사 확인 필요');
  }

  // 구조
  lines.push('');
  lines.push(sectionTitle_('구조 관련'));
  lines.push(line_(h('O'), v('O')));
  lines.push(line_(h('P'), v('P')));
  {
    const q = v('Q'), r = v('R'), s = v('S');
    const parts = [];
    if (has(q)) parts.push(q);
    const rStr = has(r) ? `방 ${r}` : '';
    const sStr = has(s) ? `욕실 ${s}` : '';
    const mix = joinWithSlash_(rStr, sStr);
    if (mix) parts.push(mix);
    lines.push(line_('구조', parts.filter(has).join(' ')));
  }
  lines.push(line_(h('AJ'), v('AJ')));

  // 주차 (규칙: 주차❌이면 “🅿️ 주차❌”만 표기)
  lines.push('');
  lines.push(sectionTitle_('주차 관련'));
  const parking = v('Z');
  if (/주차❌/.test(parking)) {
    lines.push('🅿️ 주차❌');
  } else {
    lines.push(line_(h('Z'), parking));
    if (has(v('AA'))) lines.push(line_(h('AA'), v('AA')));
    if (has(v('AB'))) lines.push(line_(h('AB'), v('AB')));
  }

  // 반려동물 (규칙: 매매 매물일 경우 표기하지 않음)
  if (!/매매/.test(tradeType || '')) {
    lines.push('');
    lines.push(sectionTitle_('반려동물'));
    lines.push(line_(h('AC'), v('AC')));
  }

  // 옵션
  lines.push('');
  lines.push(sectionTitle_('옵션 관련'));
  lines.push(line_(h('AF'), v('AF')));

  // 입주
  lines.push('');
  lines.push(sectionTitle_('입주 정보'));
  lines.push(line_(h('T'), v('T')));

  return lines.join('\n');
}

function makeAdCopyFromActiveRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const active = sheet.getActiveRange();
  if (!active) return ui.alert('셀을 선택한 뒤 다시 실행해 주세요');

  const row = active.getRow();
  if (row <= 2) return ui.alert('데이터가 있는 행(3행 이상)을 선택해 주세요');

  const text = buildAdCopyTextFromRow_(sheet, row);

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:pretendard,apple sd gothic neo,malgun gothic,system-ui,sans-serif;padding:12px 12px 20px;max-width:720px;">'
    + '<h2 style="margin:0 0 10px;">광고문구 미리보기</h2>'
    + '<pre style="white-space:pre-wrap;word-break:break-word;border:1px solid #ddd;border-radius:8px;padding:12px;background:#fafafa;max-height:70vh;overflow:auto;">'
    + escapeHtml_(text)
    + '</pre>'
    + '<button onclick="copy()" style="margin-top:10px;padding:8px 12px;border-radius:8px;border:1px solid #bbb;background:#fff;cursor:pointer;">복사하기</button>'
    + '<script>function copy(){const t=document.querySelector(\"pre\").innerText;navigator.clipboard.writeText(t).then(()=>{alert(\"복사됨\");});}</script>'
    + '</div>'
  ).setWidth(800).setHeight(600);

  ui.showModalDialog(html, '광고문구 작성');
}

/** ========== 유틸 ========== */
function line_(label, value) {
  const L = String(label || '').trim();
  const V = String(value || '').trim();
  return `- ${L} : ${V}`;
}
function sectionTitle_(t) { return `\n${t}`; }
function joinWithSlash_(a, b) {
  const A = String(a || '').trim();
  const B = String(b || '').trim();
  if (A && B) return `${A} / ${B}`;
  return A || B || '';
}
function escapeHtml_(str) {
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

/** 숫자 문자열 → 규칙에 맞는 ‘억/만 원’ 포맷으로 변환 */
function fmtPriceKRW_(s) {
  if (!s) return s;
  // 이미 '억' 포함 등은 원본 유지
  if (/[억만]\s*원/.test(s)) return s;
  const num = Number(String(s).replace(/[^\d.-]/g, ''));
  if (!isFinite(num) || num <= 0) return s;

  // 단위: ‘만원’ 기준값으로 가정될 수 있어 애매하면 만원 단위로 해석
  // 예: 54800 → 5억 4,800만 원
  // 1000만 미만은 그대로 ‘만 원’ 표기
  if (num < 1000) return `${comma_(num)}만 원`;

  const eok = Math.floor(num / 10000);
  const man = num % 10000;
  if (eok > 0) {
    const manStr = man ? ` ${comma_(man)}만 원` : ' 원';
    return `${comma_(eok)}억${manStr}`;
  }
  return `${comma_(man)}만 원`;
}
function comma_(n){ return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }

/** 촬영 문구 작성: 현재 행에서 텍스트 생성 + 사이드바 오픈 (소재지 정보 제거 + PW/AG 표시) */
function makeShootCopyFromActiveRow() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (!isTargetSheet_(sh)) {
    SpreadsheetApp.getUi().alert("대상 탭이 아닙니다.");
    return;
  }
  const r = sh.getActiveRange();
  if (!r) return;

  const row = r.getRow();
  if (row < (CFG.ROW_START || 2)) {
    SpreadsheetApp.getUi().alert("데이터가 시작되는 행을 선택하세요.");
    return;
  }

  // 필요한 열 인덱스
  const C  = CFG.COL_ADDR; // 주소(C)
  const U  = 21;           // U열 - 담당자/연락처(줄바꿈 포함 가능)
  const X  = 24;           // X열 - 세입자 연락처(있으면 우선)
  const AG = 33;           // AG열 - PW

  // 값 읽기
  const vals = sh.getRange(row, 1, 1, Math.max(C, U, X, AG)).getValues()[0];
  const addr = String(vals[C - 1]  || "").trim();
  const uVal = String(vals[U - 1]  || "").trim();
  const xVal = String(vals[X - 1]  || "").trim();
  const pw   = String(vals[AG - 1] || "").trim();

  // 연락처 라인: X가 있으면 "세입자 : X", 없으면 U(줄바꿈 → 공백)
  let contactLine = "";
  if (xVal) {
    contactLine = "세입자 : " + oneLine_(xVal);
  } else if (uVal) {
    contactLine = oneLine_(uVal);
  } else {
    contactLine = "(연락처 정보 없음)";
  }

  // 광고문구 본문 생성
  let tail = safeBuildAdCopyText_(sh, row);
  tail = String(tail || "");

  // === 촬영 문구에서는 '소재지 정보' 섹션 제거 ===
  tail = dropSojaejiFromAdCopy_(tail);

  // 섹션 간 1줄만 유지
  tail = normalizeParagraphSpacing_(tail);

  // 최종 텍스트
  const lines = [];
  if (addr) lines.push(addr);
  lines.push("────────────────");
  lines.push("연락처 정보");
  lines.push(contactLine);
  lines.push("PW");
  lines.push(pw ? pw : "(PW 없음)");
  lines.push(""); // 한 줄 띄우기
  if (tail) lines.push(tail);

  const finalText = lines.join("\n");

  // 사이드바 표시
  const html = HtmlService.createHtmlOutput(shootCopySidebarTpl_(finalText))
    .setTitle("촬영 문구 작성");
  SpreadsheetApp.getUi().showSidebar(html);
}

/** 여러 줄 → 한 줄(연속 공백/줄바꿈을 공백 1개로) */
function oneLine_(s) {
  return String(s || "").replace(/[\r\n]+/g, " ").replace(/\s+/g, " ").trim();
}

/** 섹션 간 두 줄 이상 → 정확히 한 줄만 남기기 */
function normalizeParagraphSpacing_(s) {
  let t = String(s || "");
  t = t.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  t = t.replace(/\n{3,}/g, "\n\n");
  t = t.replace(/<br\s*\/?>/gi, "\n");
  t = t.replace(/[ \t]+\n/g, "\n").trim();
  return t;
}

/** 기존 ‘광고문구’ 텍스트 생성기를 최대한 재사용 (없으면 "") */
function safeBuildAdCopyText_(sh, row) {
  try {
    if (typeof buildAdCopyTextFromRow_ === "function") {
      return String(buildAdCopyTextFromRow_(sh, row) || "");
    }
    if (typeof AD_COPY_buildTextFromRow === "function") {
      return String(AD_COPY_buildTextFromRow(sh, row) || "");
    }
    if (typeof makeAdCopyFromActiveRow_getText_ === "function") {
      return String(makeAdCopyFromActiveRow_getText_(sh, row) || "");
    }
  } catch (e) {
    Logger.log("[safeBuildAdCopyText_] " + e);
  }
  return "";
}

/** 사이드바 템플릿: textarea + 복사 버튼 (줄바꿈 그대로 복사) */
function shootCopySidebarTpl_(text) {
  const escaped = String(text).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
  return `
    <div style="font-family:system-ui,-apple-system,Segoe UI,Roboto,Apple SD Gothic Neo,Noto Sans KR,sans-serif;padding:10px;">
      <div style="margin-bottom:8px;font-weight:700;">촬영 문구</div>
      <textarea id="t" style="width:100%;height:420px;box-sizing:border-box;font-size:13px;line-height:1.5;"
        spellcheck="false">${escaped}</textarea>
      <div style="margin-top:8px;display:flex;gap:8px;">
        <button onclick="copyTxt()" style="padding:6px 10px;border:1px solid #ddd;border-radius:8px;background:#fff;cursor:pointer;">복사</button>
        <span id="msg" style="font-size:12px;color:#6b7280;align-self:center;"></span>
      </div>
    </div>
    <script>
      function copyTxt(){
        const el = document.getElementById('t');
        el.select(); el.setSelectionRange(0, 999999);
        const ok = document.execCommand('copy');
        document.getElementById('msg').textContent = ok ? "복사됨(줄 간격 1줄)" : "복사 실패";
      }
    </script>
  `;
}

/** 광고문구 텍스트에서 '소재지 정보' 섹션만 제거 */
function dropSojaejiFromAdCopy_(s) {
  let t = String(s || "");
  t = t.replace(/\n소재지 정보\s*\n- [^\n]*\n?/g, "\n"); // 기본 패턴 제거
  t = t.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  t = t.replace(/\n{3,}/g, "\n\n").replace(/[ \t]+\n/g, "\n").trim();
  return t;
}
