/*** Config.gs ***/
var CFG = {
  // 대상 데이터 탭들
  TARGET_SHEETS: ["🏢아파트","🏘️투/쓰리룸","🏠원룸","🏪상가/건물"],

  // 공통 컬럼 (1-based)
  ROW_START: 3,
  COL_ADDR: 3, // C
  COL_LAT:  4, // D
  COL_LON:  5, // E

  // 판정/라벨 컬럼
  COL_MOA_FLAG:  6,  // F
  COL_FAST_FLAG: 7,  // G
  COL_ADR_LABEL: 8,  // H

  // 블록 처리 기본 크기
  BLOCK_ROWS: 800,

  // 모아타운 현황 시트 구성
  MOA: {
    SHEET: "💚모아타운",
    COL_REP:   2,  // B 대표지번
    COL_STAGE: 6,  // F 현재단계
    COL_POLY: 22,  // V 좌표
    START_ROW: 3,
    CACHE_KEY: "moa_rows_v4",
    CACHE_TTL: 600 // 10min
  },

  // 신속통합 현황 시트 구성
  FAST: {
    SHEET: "❤️신속통합",
    COL_REP:   2,  // B 대표지번
    COL_STAGE: 6,  // F 현재단계
    COL_POLY:  8,  // H 좌표
    START_ROW: 3,
    CACHE_KEY: "fast_rows_v1",
    CACHE_TTL: 600
  },

  // 지오코딩 캐시
  GEO: {
    KEY_PREFIX: "geo:",
    TTL: 21600 // 6h
  },

  // 행추가 템플릿
  TEMPLATE_ROW: 3
};

/* =========================
 * Kakao 키
 * ========================= */
const KAKAO_REST_KEY = "f2e2b0e7cb648aef91a2ca6607a4231c";
const KAKAO_JS_KEY = '6b0a27a5811411b86b11906aa46262b8';