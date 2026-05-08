// Code.gs — Baseball Savermatrix Google Apps Script 백엔드
// Google Apps Script 에디터에 전체 복사 후 저장 → 새 배포(웹 앱) → 누구나 접근으로 배포

// ── 시트 이름 상수 ──────────────────────────────────────────
const SHEETS = {
  games:    'games',
  bat_log:  'bat_log',
  pit_bf:   'pit_bf',
  pit_runs: 'pit_runs',
  roster:   'roster'
};

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
      // 업데이트
      sh.getRange(existingRowNum, 1, 1, row.length).setValues([row]);
    } else {
      // 추가
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
    if (data.type === 'sync') {
      // 삭제 먼저 처리 (새 데이터 upsert 전에)
      const delGids = data.deleted_gids || [];
      if (delGids.length) {
        deleteByGid(SHEETS.games,    delGids);
        deleteByGid(SHEETS.bat_log,  delGids);
        deleteByGid(SHEETS.pit_bf,   delGids);
        deleteByGid(SHEETS.pit_runs, delGids);
      }
      upsertRows(SHEETS.games,    data.games    || []);
      upsertRows(SHEETS.bat_log,  data.bat_log  || []);
      upsertRows(SHEETS.pit_bf,   data.pit_bf   || []);
      upsertRows(SHEETS.pit_runs, data.pit_runs || []);
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
    // 선수 명단 저장
    if (action === 'saveRoster') {
      const roster = JSON.parse(decodeURIComponent(e.parameter.data));
      const sh = getSheet(SHEETS.roster);
      sh.clearContents();
      if (roster.length) {
        const hdrs = Object.keys(roster[0]);
        sh.getRange(1, 1, 1, hdrs.length).setValues([hdrs]);
        const rows = roster.map(p => hdrs.map(h => p[h] !== undefined && p[h] !== null ? p[h] : ''));
        sh.getRange(2, 1, rows.length, hdrs.length).setValues(rows);
      }
      return out({ status: 'ok' });
    }

    // 전체 데이터 불러오기
    if (action === 'fetch') {
      return out({
        status:   'ok',
        games:    readAll(SHEETS.games),
        bat_log:  readAll(SHEETS.bat_log),
        pit_bf:   readAll(SHEETS.pit_bf),
        pit_runs: readAll(SHEETS.pit_runs),
        roster:   readAll(SHEETS.roster)
      });
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
