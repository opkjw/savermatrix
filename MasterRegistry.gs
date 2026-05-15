// MasterRegistry.gs — 모야모야 팀 중앙 레지스트리
// 별도 Google Apps Script 프로젝트에 배포하세요.
// 배포 후 얻은 URL을 index.html의 MASTER_REGISTRY_URL 상수에 설정하세요.
//
// 시트 구조:
//   teams        : name | url
//   reset_tokens : token | created | expires | used_at

// ── 설정 ────────────────────────────────────────────────────
const MASTER_PIN = 'CHANGE_ME';           // 관리자 PIN (반드시 변경)
const APP_URL    = 'https://opkjw.github.io/svtest/';  // 앱 배포 URL
const TOKEN_TTL  = 24 * 60 * 60 * 1000;  // 토큰 유효시간: 24시간

// ── 시트 이름 ────────────────────────────────────────────────
const MS = { teams: 'teams', tokens: 'reset_tokens' };

// ── 유틸 ────────────────────────────────────────────────────
function mOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function mSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function mInitHeaders(sh, headers) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

// ── GET 핸들러 ───────────────────────────────────────────────
function doGet(e) {
  try {
    const action = (e.parameter.action || '').trim();

    // 팀 목록 반환 (공개)
    if (action === 'teams') {
      const sh = mSheet(MS.teams);
      const vals = sh.getDataRange().getValues();
      if (vals.length < 2 || !vals[0][0]) return mOut({ status: 'ok', teams: [] });
      const hdrs = vals[0].map(String);
      const ni = hdrs.indexOf('name'), ui = hdrs.indexOf('url');
      if (ni < 0 || ui < 0) return mOut({ status: 'ok', teams: [] });
      const teams = vals.slice(1)
        .filter(r => r[ni] && r[ui])
        .map(r => ({ name: String(r[ni]), url: String(r[ui]) }));
      return mOut({ status: 'ok', teams });
    }

    // 토큰 검증 및 소비 (공개 — 토큰 자체가 비밀)
    if (action === 'validateToken') {
      const token = (e.parameter.token || '').trim().toUpperCase();
      if (!token) return mOut({ status: 'error', message: '토큰이 없습니다' });

      const sh = mSheet(MS.tokens);
      const vals = sh.getDataRange().getValues();
      if (vals.length < 2) return mOut({ status: 'error', message: '유효하지 않은 링크입니다' });

      const hdrs = vals[0].map(String);
      const ti = hdrs.indexOf('token');
      const ei = hdrs.indexOf('expires');
      const ui = hdrs.indexOf('used_at');
      if (ti < 0) return mOut({ status: 'error', message: '시트 구조 오류' });

      const now = Date.now();
      for (let i = 1; i < vals.length; i++) {
        if (String(vals[i][ti]).toUpperCase() !== token) continue;
        if (vals[i][ui]) return mOut({ status: 'error', message: '이미 사용된 링크입니다' });
        if (now > Number(vals[i][ei])) return mOut({ status: 'error', message: '링크가 만료되었습니다 (24시간 초과)' });
        // 유효 → 사용 처리
        sh.getRange(i + 1, ui + 1).setValue(now);
        return mOut({ status: 'ok' });
      }
      return mOut({ status: 'error', message: '유효하지 않은 링크입니다' });
    }

    // 토큰 생성 (관리자 전용)
    if (action === 'generateToken') {
      const pin = e.parameter.pin || '';
      if (pin !== MASTER_PIN) return mOut({ status: 'error', message: '인증 실패' });

      const token = Utilities.getUuid().replace(/-/g, '').substring(0, 12).toUpperCase();
      const now = Date.now();
      const expires = now + TOKEN_TTL;

      const sh = mSheet(MS.tokens);
      mInitHeaders(sh, ['token', 'created', 'expires', 'used_at']);
      sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setValues([[token, now, expires, '']]);

      const masterUrl = ScriptApp.getService().getUrl();
      const encoded   = Utilities.base64Encode(masterUrl);
      const link      = APP_URL + '?reset=' + token + '&src=' + encodeURIComponent(encoded);

      return mOut({
        status: 'ok', token, expires,
        link,
        tip: '이 링크를 학부모에게 전송하세요. 24시간 이내 1회만 사용 가능합니다.'
      });
    }

    // 시트 초기화 (관리자 전용 — 최초 1회)
    if (action === 'init') {
      const pin = e.parameter.pin || '';
      if (pin !== MASTER_PIN) return mOut({ status: 'error', message: '인증 실패' });
      const tsh = mSheet(MS.teams);
      mInitHeaders(tsh, ['name', 'url']);
      mSheet(MS.tokens);
      mInitHeaders(mSheet(MS.tokens), ['token', 'created', 'expires', 'used_at']);
      return mOut({ status: 'ok', message: '시트 초기화 완료. teams 시트에 팀 목록을 입력하세요.' });
    }

    return mOut({ status: 'error', message: '알 수 없는 action' });
  } catch (err) {
    return mOut({ status: 'error', message: err.toString() });
  }
}

// ── 만료 토큰 정리 (선택적 — 트리거로 주기 실행 가능) ─────
function cleanExpiredTokens() {
  const sh = mSheet(MS.tokens);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return;
  const hdrs = vals[0].map(String);
  const ei = hdrs.indexOf('expires');
  const now = Date.now();
  for (let i = vals.length - 1; i >= 1; i--) {
    if (Number(vals[i][ei]) < now) sh.deleteRow(i + 1);
  }
}
