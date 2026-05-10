// Code.gs — Baseball Savermatrix Google Apps Script 백엔드
// Google Apps Script 에디터에 전체 복사 후 저장 → 새 배포(웹 앱) → 누구나 접근으로 배포

// ── 인증 설정 ───────────────────────────────────────────────
// AUTH_TOKEN: 빈 문자열('')이면 인증 없이 동작
const AUTH_TOKEN = '';

// PIN 로그인 설정
// ADMIN_PIN:    코치/감독용 (전체 기능 접근)
// RECORDER_PIN: 기록자용 (타석·투수 입력+기록 접근, 통계 없음)
// USER_PIN:     학부모용 (우리 아이만 접근)
// 빈 문자열이면 해당 역할 로그인 비활성화
const ADMIN_PIN    = '1234';
const RECORDER_PIN = '2222';
const USER_PIN     = '0000';

// 팀/학교명 — 학부모 앱 접속 시 자동으로 채워집니다
// 빈 문자열('')이면 학부모가 직접 입력
const TEAM_NAME = '';

// ── 시트 이름 상수 ──────────────────────────────────────────
const SHEETS = {
  games:    'games',
  bat_log:  'bat_log',
  pit_bf:   'pit_bf',
  pit_runs: 'pit_runs',
  roster:   'roster',
  teams:    'teams'   // 팀명 목록 시트 (헤더: name)
};

// ── 유틸: 토큰 인증 검사 ───────────────────────────────────
function checkAuth(token) {
  if (!AUTH_TOKEN) return true;          // 토큰 미설정이면 누구나 허용
  return token === AUTH_TOKEN;
}

// ── 검증 한도 ──────────────────────────────────────────────
const VAL_LIMITS = {
  maxRowsPerSheet: 5000,
  maxStrLen:       500,
  maxDeletedGids:  100,
  maxRosterSize:   200
};

// ── OC 코드 화이트리스트 (타자/투수 결과) ─────────────────
const VALID_OC = {
  bat: ['1B','2B','3B','HR','BB','IBB','HBP','K','KL','GO','FO','LO','PO','DP','SF','SH','E','FC','_RUN','BIP_OC'],
  pit: ['1B','2B','3B','HR','BB','IBB','HBP','K','KL','GO','FO','LO','PO','DP','SF','SH','E','FC','BK','WP']
};

// ── 시트별 검증 스키마 ────────────────────────────────────
// type: 's'=string, 'n'=number, 'b'=boolean, 'd'=date(YYYY-MM-DD), 'oc_bat'/'oc_pit'=OC코드
function getSchema(name) {
  if (name === 'games')    return { req: ['id'], fields: { id:'s', date:'d', opp:'s', type:'s', our:'n', opp_:'n', notes:'s', lastSync:'s' } };
  if (name === 'bat_log')  return { req: ['id','gid'], fields: { id:'s', gid:'s', date:'d', opp:'s', pno:'n', pn:'s', oc:'oc_bat', rbi:'n', run:'n', sb:'n', cs:'n', zone:'s', dir:'s' } };
  if (name === 'pit_bf')   return { req: ['id','gid'], fields: { id:'s', gid:'s', date:'d', opp:'s', pno:'n', pn:'s', oc:'oc_pit' } };
  if (name === 'pit_runs') return { req: ['id','gid'], fields: { id:'s', gid:'s', date:'d', opp:'s', pno:'n', pn:'s', earned:'b' } };
  return null;
}

// ── 단일 값 검증/정규화 (반환: 정규화된 값 또는 SKIP_VALUE) ─
const SKIP_VALUE = Symbol ? Symbol('skip') : '__SKIP__';
function sanitizeValue(v, type) {
  if (v === null || v === undefined || v === '') return null;
  if (type === 's') {
    var s = String(v);
    if (s.length > VAL_LIMITS.maxStrLen) s = s.slice(0, VAL_LIMITS.maxStrLen);
    return s;
  }
  if (type === 'n') {
    var n = Number(v);
    if (!isFinite(n)) return null;
    return n;
  }
  if (type === 'b') return !!v;
  if (type === 'd') {
    var ds = String(v).slice(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(ds)) return null;
    return ds;
  }
  if (type === 'oc_bat') {
    var oc = String(v);
    return VALID_OC.bat.indexOf(oc) >= 0 ? oc : null;
  }
  if (type === 'oc_pit') {
    var op = String(v);
    return VALID_OC.pit.indexOf(op) >= 0 ? op : null;
  }
  return null;
}

// ── 행 검증: 유효하면 정규화 객체, 무효하면 null ───────────
function sanitizeRow(item, schema) {
  if (!item || typeof item !== 'object') return null;
  // 필수 필드 검사
  for (var i = 0; i < schema.req.length; i++) {
    var k = schema.req[i];
    if (item[k] === undefined || item[k] === null || item[k] === '') return null;
  }
  // 화이트리스트 필드만 통과
  var out = {};
  var keys = Object.keys(schema.fields);
  for (var j = 0; j < keys.length; j++) {
    var key = keys[j];
    if (!(key in item)) continue;
    var sanitized = sanitizeValue(item[key], schema.fields[key]);
    if (sanitized !== null) out[key] = sanitized;
  }
  // id는 필수이므로 정규화 후에도 살아있어야 함
  if (schema.req.indexOf('id') >= 0 && !out.id) return null;
  return out;
}

// ── 배열 검증: { valid: [...], skipped: N } ───────────────
function sanitizeRows(items, sheetName) {
  if (!Array.isArray(items)) return { valid: [], skipped: 0 };
  var schema = getSchema(sheetName);
  if (!schema) return { valid: [], skipped: items.length };
  var capped = items.slice(0, VAL_LIMITS.maxRowsPerSheet);
  var skipped = items.length - capped.length;
  var valid = [];
  for (var i = 0; i < capped.length; i++) {
    var row = sanitizeRow(capped[i], schema);
    if (row) valid.push(row);
    else skipped++;
  }
  return { valid: valid, skipped: skipped };
}

// ── 유틸: 시트 가져오기 (없으면 생성) ─────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ── 유틸: 시트 전체를 JSON 배열로 읽기 ────────────────────
function readAll(sheetName) {
  const sh = getSheet(sheetName);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2 || !vals[0][0]) return [];
  const hdrs = vals[0].map(String);
  const tz = Session.getScriptTimeZone();
  return vals.slice(1)
    .filter(r => r[0] !== '' && r[0] !== null && r[0] !== undefined)
    .map(r => {
      const obj = {};
      hdrs.forEach((h, i) => {
        if (!h) return;
        let v = r[i];
        // 구글 시트가 날짜 문자열을 Date 객체로 자동 변환하는 경우 YYYY-MM-DD로 복원
        if (v instanceof Date) {
          v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        } else if (v === '' || v === null) {
          v = null;
        }
        obj[h] = v;
      });
      return obj;
    });
}

// ── 유틸: upsert (id 기준 — 있으면 update, 없으면 append) ─
function upsertRows(sheetName, items) {
  if (!items || !items.length) return;
  const sh = getSheet(sheetName);

  // 현재 헤더 확인
  const allVals = sh.getDataRange().getValues();
  let hdrs;

  if (allVals.length < 1 || !allVals[0][0]) {
    // 시트가 비어있음 → 첫 아이템 기준으로 헤더 생성
    hdrs = Object.keys(items[0]);
    sh.clearContents();
    sh.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
  } else {
    hdrs = allVals[0].map(String);
    // 새 필드가 생겼으면 헤더에 추가
    const newKeys = Object.keys(items[0]).filter(k => k && !hdrs.includes(k));
    if (newKeys.length) {
      hdrs = [...hdrs, ...newKeys];
      sh.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
    }
  }

  // id → 행번호 맵 (헤더 제외, 1-based)
  const idColIdx = hdrs.indexOf('id');
  const freshVals = sh.getDataRange().getValues();
  const idRowMap = {};
  if (idColIdx >= 0 && freshVals.length > 1) {
    freshVals.slice(1).forEach((r, i) => {
      if (r[idColIdx] !== '' && r[idColIdx] !== null) {
        idRowMap[String(r[idColIdx])] = i + 2; // +1 header, +1 0→1-based
      }
    });
  }

  // 날짜처럼 생긴 컬럼을 텍스트 형식으로 고정 (시트 자동 변환 방지)
  const dateLikeCols = hdrs.reduce((acc, h, i) => {
    if (h === 'date' || h === 'lastSync') acc.push(i + 1);
    return acc;
  }, []);
  dateLikeCols.forEach(colNum => {
    sh.getRange(1, colNum, sh.getMaxRows(), 1).setNumberFormat('@STRING@');
  });

  // upsert 처리
  items.forEach(item => {
    const row = hdrs.map(h => (item[h] !== undefined && item[h] !== null) ? item[h] : '');
    const existingRowNum = idColIdx >= 0 ? idRowMap[String(item.id)] : null;
    if (existingRowNum) {
      sh.getRange(existingRowNum, 1, 1, row.length).setValues([row]);
    } else {
      const lastRow = sh.getLastRow();
      sh.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
      if (idColIdx >= 0) idRowMap[String(item.id)] = lastRow + 1;
    }
  });
}

// ── 유틸: gid 기준으로 시트에서 해당 행 삭제 ──────────────
function deleteByGid(sheetName, gids) {
  if (!gids || !gids.length) return;
  const sh = getSheet(sheetName);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return;
  const hdrs = vals[0].map(String);
  const gidCol = hdrs.indexOf('gid');
  const idCol  = hdrs.indexOf('id');
  // games 시트는 id 컬럼 자체가 gid
  const keyCol = gidCol >= 0 ? gidCol : idCol;
  if (keyCol < 0) return;
  const gidSet = new Set(gids.map(String));
  // 뒤에서부터 삭제해야 행 번호 밀림 없음
  for (let i = vals.length - 1; i >= 1; i--) {
    if (gidSet.has(String(vals[i][keyCol]))) {
      sh.deleteRow(i + 1); // 1-based
    }
  }
}

// ── POST 핸들러: 기록 동기화 ───────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (!checkAuth(data.token || '')) {
      return out({ status: 'error', message: '인증 실패: 올바른 토큰이 필요합니다.' });
    }
    if (data.type === 'sync') {
      // 삭제 처리 (새 데이터 upsert 전에) — 한도 적용 + 문자열 변환
      var rawDel = Array.isArray(data.deleted_gids) ? data.deleted_gids : [];
      var delGids = rawDel.slice(0, VAL_LIMITS.maxDeletedGids)
                          .map(function(g){ return String(g); })
                          .filter(function(g){ return g && g.length <= VAL_LIMITS.maxStrLen; });
      var delSkipped = rawDel.length - delGids.length;
      if (delGids.length) {
        deleteByGid(SHEETS.games,    delGids);
        deleteByGid(SHEETS.bat_log,  delGids);
        deleteByGid(SHEETS.pit_bf,   delGids);
        deleteByGid(SHEETS.pit_runs, delGids);
      }
      // 시트별 행 검증 후 upsert
      var vGames = sanitizeRows(data.games,    SHEETS.games);
      var vBat   = sanitizeRows(data.bat_log,  SHEETS.bat_log);
      var vPbf   = sanitizeRows(data.pit_bf,   SHEETS.pit_bf);
      var vPrun  = sanitizeRows(data.pit_runs, SHEETS.pit_runs);
      upsertRows(SHEETS.games,    vGames.valid);
      upsertRows(SHEETS.bat_log,  vBat.valid);
      upsertRows(SHEETS.pit_bf,   vPbf.valid);
      upsertRows(SHEETS.pit_runs, vPrun.valid);
      return out({
        status: 'ok',
        accepted: { games: vGames.valid.length, bat_log: vBat.valid.length, pit_bf: vPbf.valid.length, pit_runs: vPrun.valid.length, deleted_gids: delGids.length },
        skipped:  { games: vGames.skipped,     bat_log: vBat.skipped,     pit_bf: vPbf.skipped,     pit_runs: vPrun.skipped,     deleted_gids: delSkipped }
      });
    }
    return out({ status: 'ok' });
  } catch (err) {
    return out({ status: 'error', message: err.toString() });
  }
}

// ── GET 핸들러: 불러오기 / 선수 명단 저장 ─────────────────
function doGet(e) {
  const action = e.parameter.action;
  try {
    // PIN 로그인은 AUTH_TOKEN 체크 없이 먼저 처리 (이 자체가 인증)
    if (action === 'login') {
      const pin = e.parameter.pin || '';
      if (ADMIN_PIN    && pin === ADMIN_PIN)    return out({ status: 'ok', role: 'admin' });
      if (RECORDER_PIN && pin === RECORDER_PIN) return out({ status: 'ok', role: 'recorder' });
      if (USER_PIN     && pin === USER_PIN)     return out({ status: 'ok', role: 'user' });
      return out({ status: 'error', message: 'PIN이 올바르지 않습니다.' });
    }

    if (!checkAuth(e.parameter.token || '')) {
      return out({ status: 'error', message: '인증 실패: 올바른 토큰이 필요합니다.' });
    }

    // 선수 명단 저장
    if (action === 'saveRoster') {
      var raw = JSON.parse(decodeURIComponent(e.parameter.data));
      if (!Array.isArray(raw)) {
        return out({ status: 'error', message: '선수 명단 형식이 올바르지 않습니다.' });
      }
      // 사이즈 한도 + 기본 검증 (각 항목은 객체이고 no 필드 필수)
      var rosterSkipped = Math.max(0, raw.length - VAL_LIMITS.maxRosterSize);
      const roster = raw.slice(0, VAL_LIMITS.maxRosterSize).filter(function(p){
        return p && typeof p === 'object' && (p.no !== undefined && p.no !== null && p.no !== '');
      });
      rosterSkipped += (Math.min(raw.length, VAL_LIMITS.maxRosterSize) - roster.length);
      // 문자열 필드 길이 제한
      roster.forEach(function(p){
        Object.keys(p).forEach(function(k){
          if (typeof p[k] === 'string' && p[k].length > VAL_LIMITS.maxStrLen) {
            p[k] = p[k].slice(0, VAL_LIMITS.maxStrLen);
          }
        });
      });
      const sh = getSheet(SHEETS.roster);
      sh.clearContents();
      if (roster.length) {
        // 모든 선수 항목의 키를 합산해 헤더 생성 (첫 항목만 보면 siblings 등 누락 가능)
        const keySet = new Set();
        roster.forEach(function(p) { Object.keys(p).forEach(function(k) { keySet.add(k); }); });
        const hdrs = Array.from(keySet);
        sh.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
        const rows = roster.map(function(p) {
          return hdrs.map(function(h) {
            var v = p[h];
            if (Array.isArray(v)) return JSON.stringify(v); // 배열은 JSON 문자열로 저장
            return (v !== undefined && v !== null) ? v : '';
          });
        });
        sh.getRange(2, 1, rows.length, hdrs.length).setValues(rows);
      }
      return out({ status: 'ok', accepted: roster.length, skipped: rosterSkipped });
    }

    // 경기 목록만 빠르게 불러오기 (게임 시작 시 중복 검사용)
    if (action === 'fetchGames') {
      return out({
        status:   'ok',
        games:    readAll(SHEETS.games)
      });
    }

    // 전체 데이터 불러오기
    if (action === 'fetch') {
      return out({
        status:   'ok',
        teamName: TEAM_NAME,
        games:    readAll(SHEETS.games),
        bat_log:  readAll(SHEETS.bat_log),
        pit_bf:   readAll(SHEETS.pit_bf),
        pit_runs: readAll(SHEETS.pit_runs),
        roster:   readAll(SHEETS.roster)
      });
    }

    // 팀명 목록 불러오기
    if (action === 'getTeams') {
      var rows = readAll(SHEETS.teams);
      var teams = rows.map(function(r){ return r.name || ''; }).filter(Boolean);
      return out({ status: 'ok', teams: teams });
    }

    // 연결 테스트
    return out({ status: 'ok', message: 'Baseball Savermatrix API' });
  } catch (err) {
    return out({ status: 'error', message: err.toString() });
  }
}

// ── 유틸: JSON 응답 생성 ────────────────────────────────────
function out(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
