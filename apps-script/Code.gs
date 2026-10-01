/**
 * 수시모집 합불 조회/입력 API
 * Google Sheets "편집" 탭(A:U)을 읽고 Q:U만 갱신합니다.
 *
 * 배포 전 Apps Script 프로젝트 설정 > 스크립트 속성에
 * ACCESS_PASSWORD와 SESSION_SECRET을 반드시 등록하세요.
 */

const CONFIG = Object.freeze({
  spreadsheetId: '1giYYR1RO7jBVusFqtOaDdUs-KbE_wYmbm5nonsO54uI',
  sheetName: '편집',
  headerRow: 1,
  firstDataRow: 2,
  totalColumns: 21,
  maxSearchResults: 250,
  maxSuggestions: 12,
  lockWaitMs: 20000,
  passwordProperty: 'ACCESS_PASSWORD',
  sessionSecretProperty: 'SESSION_SECRET',
  sessionDurationMs: 3 * 60 * 60 * 1000,
});

const EXPECTED_HEADERS = Object.freeze([
  '대학구분', '재학여부', '재학년도', '반', '번호', '성명', '대학명',
  '모집시기', '전형유형', '선발유형', '계열', '전형명', '모집단위',
  '(원본)평가단위명', '생년월일', '수험번호', '1단계 합격', '최종 합격',
  '불합격 사유', '최초 충원번호', '최종 충원번호',
]);

const COL = Object.freeze({
  universityType: 0, // A 대학구분
  enrollment: 1,     // B 재학여부
  schoolYear: 2,     // C 재학년도
  classNo: 3,        // D 반
  studentNo: 4,      // E 번호
  name: 5,           // F 성명
  university: 6,     // G 대학명
  season: 7,         // H 모집시기
  admissionType: 8,  // I 전형유형
  selectionType: 9,  // J 선발유형
  track: 10,         // K 계열
  admissionName: 11, // L 전형명
  department: 12,    // M 모집단위
  originalUnit: 13,  // N 원본 평가단위명
  birthdate: 14,     // O 생년월일
  examNo: 15,        // P 수험번호
  stage1: 16,        // Q 1단계 합격
  finalResult: 17,   // R 최종 합격
  failureReason: 18, // S 불합격 사유
  firstWait: 19,     // T 최초 충원번호
  finalWait: 20,     // U 최종 충원번호
});

const ALLOWED = Object.freeze({
  stage1: ['', '합격', '불합격'],
  finalResult: ['', '합격', '불합격', '충원합격'],
  failureReason: ['', '미응시', '최저미충족', '불합격', '1단계 불합격'],
});

function doGet() {
  return json_({
    ok: true,
    service: 'admissions-result-sync',
    message: 'POST 요청을 사용하세요.',
    serverTime: new Date().toISOString(),
  });
}

function doPost(e) {
  try {
    const payload = parsePayload_(e);
    if (String(payload.action || '') === 'login') {
      return json_(login_(payload.password));
    }
    requireSession_(payload.sessionToken);

    switch (String(payload.action || '')) {
      case 'health':
        return json_({ ok: true, serverTime: new Date().toISOString() });
      case 'meta':
        return json_(getMeta_());
      case 'universities':
        return json_(getUniversitySuggestions_(payload.query));
      case 'search':
        return json_(search_(payload.filters || {}));
      case 'export':
        return json_(export_(payload.filters || {}));
      case 'save':
        return json_(save_(payload));
      default:
        throw apiError_('UNKNOWN_ACTION', '지원하지 않는 요청입니다.');
    }
  } catch (error) {
    return json_({
      ok: false,
      code: error && error.code ? error.code : 'SERVER_ERROR',
      message: safeErrorMessage_(error),
      details: error && error.details ? error.details : undefined,
      serverTime: new Date().toISOString(),
    });
  }
}

function parsePayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw apiError_('EMPTY_REQUEST', '요청 내용이 없습니다.');
  }
  try {
    return JSON.parse(e.postData.contents);
  } catch (error) {
    throw apiError_('INVALID_JSON', '요청 형식이 올바르지 않습니다.');
  }
}

function login_(providedPassword) {
  const properties = PropertiesService.getScriptProperties();
  const savedPassword = properties.getProperty(CONFIG.passwordProperty);
  const sessionSecret = properties.getProperty(CONFIG.sessionSecretProperty);
  if (!savedPassword || !sessionSecret) {
    throw apiError_('AUTH_NOT_CONFIGURED', '관리자가 로그인 설정을 완료하지 않았습니다.');
  }
  if (!providedPassword || !constantTimeEqual_(String(providedPassword), String(savedPassword))) {
    throw apiError_('INVALID_PASSWORD', '비밀번호가 올바르지 않습니다.');
  }

  const expiresAt = Date.now() + CONFIG.sessionDurationMs;
  const payload = base64UrlEncode_(JSON.stringify({
    exp: expiresAt,
    nonce: Utilities.getUuid(),
  }));
  const signature = sign_(payload, sessionSecret);
  return {
    ok: true,
    sessionToken: payload + '.' + signature,
    expiresAt,
    serverTime: new Date().toISOString(),
  };
}

function requireSession_(token) {
  const sessionSecret = PropertiesService.getScriptProperties().getProperty(CONFIG.sessionSecretProperty);
  if (!sessionSecret) throw apiError_('AUTH_NOT_CONFIGURED', '관리자가 로그인 설정을 완료하지 않았습니다.');
  const parts = String(token || '').split('.');
  if (parts.length !== 2 || !constantTimeEqual_(parts[1], sign_(parts[0], sessionSecret))) {
    throw apiError_('UNAUTHORIZED', '로그인이 필요합니다.');
  }
  let session;
  try {
    session = JSON.parse(base64UrlDecode_(parts[0]));
  } catch (_) {
    throw apiError_('UNAUTHORIZED', '로그인이 필요합니다.');
  }
  if (!session.exp || Number(session.exp) <= Date.now()) {
    throw apiError_('SESSION_EXPIRED', '로그인 시간이 만료되었습니다.');
  }
}

function getMeta_() {
  const rows = getAllRows_();
  return {
    ok: true,
    options: {
      universityTypes: uniqueSorted_(rows.map(r => r[COL.universityType])),
      enrollments: uniqueSorted_(rows.map(r => r[COL.enrollment])),
      classes: uniqueSorted_(rows.map(r => r[COL.classNo]), true),
      tracks: uniqueSorted_(rows.map(r => r[COL.track])),
      admissionTypes: uniqueSorted_(rows.map(r => normalizeAdmissionType_(r[COL.admissionType]))),
    },
    rowCount: rows.length,
    serverTime: new Date().toISOString(),
  };
}

function getUniversitySuggestions_(query) {
  const q = normalize_(query);
  if (q.length < 3) {
    return { ok: true, suggestions: [], serverTime: new Date().toISOString() };
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'universities-v1';
  let universities;
  const cached = cache.get(cacheKey);
  if (cached) {
    universities = JSON.parse(cached);
  } else {
    universities = uniqueSorted_(getAllRows_().map(r => r[COL.university]));
    cache.put(cacheKey, JSON.stringify(universities), 600);
  }

  const suggestions = universities
    .filter(name => normalize_(name).includes(q))
    .slice(0, CONFIG.maxSuggestions);

  return { ok: true, suggestions, serverTime: new Date().toISOString() };
}

function search_(rawFilters) {
  const filters = sanitizeFilters_(rawFilters);
  if (!hasSearchCondition_(filters)) {
    throw apiError_('FILTER_REQUIRED', '검색 조건을 한 가지 이상 입력해 주세요.');
  }

  const rows = getAllRows_();
  const matches = [];
  let totalMatches = 0;

  rows.forEach(row => {
    if (!matchesFilters_(row, filters)) return;
    totalMatches += 1;
    if (matches.length < CONFIG.maxSearchResults) {
      matches.push(toRecord_(row));
    }
  });

  return {
    ok: true,
    records: matches,
    totalMatches,
    truncated: totalMatches > matches.length,
    serverTime: new Date().toISOString(),
  };
}

function export_(rawFilters) {
  const filters = sanitizeFilters_(rawFilters);
  if (!hasSearchCondition_(filters)) {
    throw apiError_('FILTER_REQUIRED', '검색 조건을 한 가지 이상 입력해 주세요.');
  }

  const rows = getAllRows_().filter(row => matchesFilters_(row, filters));
  return {
    ok: true,
    headers: EXPECTED_HEADERS.slice(),
    rows,
    totalMatches: rows.length,
    serverTime: new Date().toISOString(),
  };
}

function save_(payload) {
  const recordId = cleanText_(payload.recordId, 128);
  const expectedVersion = cleanText_(payload.expectedVersion, 128);
  const values = payload.values || {};
  if (!recordId) throw apiError_('RECORD_ID_REQUIRED', '저장할 학생 정보가 없습니다.');

  const lock = LockService.getScriptLock();
  lock.waitLock(CONFIG.lockWaitMs);
  try {
    const sheet = getSheet_();
    assertSchema_(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.firstDataRow) throw apiError_('NOT_FOUND', '저장할 행을 찾지 못했습니다.');

    const range = sheet.getRange(CONFIG.firstDataRow, 1, lastRow - CONFIG.headerRow, CONFIG.totalColumns);
    const rows = range.getDisplayValues();
    let rowIndex = -1;
    for (let i = 0; i < rows.length; i += 1) {
      if (makeRecordId_(rows[i]) === recordId) {
        if (rowIndex !== -1) {
          throw apiError_('DUPLICATE_RECORD', '동일한 학생·대학·전형·수험번호 자료가 중복되어 저장하지 않았습니다.');
        }
        rowIndex = i;
      }
    }
    if (rowIndex === -1) throw apiError_('NOT_FOUND', '자료가 변경되었거나 해당 행을 찾지 못했습니다. 다시 검색해 주세요.');

    const current = rows[rowIndex];
    const currentVersion = makeVersion_(current);
    if (expectedVersion && currentVersion !== expectedVersion) {
      const conflict = apiError_('CONFLICT', '다른 사용자가 먼저 수정했습니다. 최신 값을 불러왔습니다.');
      conflict.details = { current: toRecord_(current) };
      throw conflict;
    }

    const lockedStage1 = normalize_(current[COL.selectionType]) === '일괄합산' ||
      normalize_(current[COL.stage1]) === '일괄합산';
    const stage1 = lockedStage1 ? '일괄합산' : validateEnum_(values.stage1, ALLOWED.stage1, '1단계 합격');
    const finalResult = validateEnum_(values.finalResult, ALLOWED.finalResult, '최종 합격');
    const failureReason = validateEnum_(values.failureReason, ALLOWED.failureReason, '불합격 사유');
    const firstWait = validateRank_(values.firstWait, '최초 충원번호');
    const finalWait = validateRank_(values.finalWait, '최종 충원번호');

    const sheetRow = CONFIG.firstDataRow + rowIndex;
    sheet.getRange(sheetRow, COL.stage1 + 1, 1, 5)
      .setValues([[stage1, finalResult, failureReason, firstWait, finalWait]]);
    SpreadsheetApp.flush();

    current[COL.stage1] = stage1;
    current[COL.finalResult] = finalResult;
    current[COL.failureReason] = failureReason;
    current[COL.firstWait] = firstWait;
    current[COL.finalWait] = finalWait;

    return {
      ok: true,
      record: toRecord_(current),
      serverTime: new Date().toISOString(),
    };
  } finally {
    lock.releaseLock();
  }
}

function getAllRows_() {
  const sheet = getSheet_();
  assertSchema_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.firstDataRow) return [];
  return sheet
    .getRange(CONFIG.firstDataRow, 1, lastRow - CONFIG.headerRow, CONFIG.totalColumns)
    .getDisplayValues();
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const sheet = spreadsheet.getSheetByName(CONFIG.sheetName);
  if (!sheet) throw apiError_('SHEET_NOT_FOUND', `시트 탭 '${CONFIG.sheetName}'을 찾지 못했습니다.`);
  return sheet;
}

function assertSchema_(sheet) {
  const actual = sheet
    .getRange(CONFIG.headerRow, 1, 1, CONFIG.totalColumns)
    .getDisplayValues()[0]
    .map(value => text_(value).trim());
  const mismatch = EXPECTED_HEADERS.findIndex((expected, index) => actual[index] !== expected);
  if (mismatch !== -1) {
    throw apiError_(
      'SCHEMA_MISMATCH',
      `${columnLetter_(mismatch + 1)}열 제목을 확인해 주세요. 저장하지 않았습니다.`
    );
  }
}

function columnLetter_(column) {
  let value = Number(column);
  let result = '';
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function sanitizeFilters_(filters) {
  return {
    university: cleanText_(filters.university, 80),
    name: cleanText_(filters.name, 40),
    admissionName: cleanText_(filters.admissionName, 100),
    universityType: cleanText_(filters.universityType, 40),
    enrollment: cleanText_(filters.enrollment, 40),
    classNo: cleanText_(filters.classNo, 20),
    detailName: cleanText_(filters.detailName, 40),
    detailUniversity: cleanText_(filters.detailUniversity, 80),
    track: cleanText_(filters.track, 40),
    admissionType: cleanText_(filters.admissionType, 40),
    detailAdmissionName: cleanText_(filters.detailAdmissionName, 100),
  };
}

function hasSearchCondition_(filters) {
  return Object.keys(filters).some(key => normalize_(filters[key]).length > 0);
}

function matchesFilters_(row, f) {
  if (!includes_(row[COL.university], f.university)) return false;
  if (!includes_(row[COL.name], f.name)) return false;
  if (!includesAdmissionName_(row[COL.admissionName], f.admissionName)) return false;
  if (!equals_(row[COL.universityType], f.universityType)) return false;
  if (!equals_(row[COL.enrollment], f.enrollment)) return false;
  if (!equals_(row[COL.classNo], f.classNo)) return false;
  if (!includes_(row[COL.name], f.detailName)) return false;
  if (!includes_(row[COL.university], f.detailUniversity)) return false;
  if (!equals_(row[COL.track], f.track)) return false;
  if (!equals_(normalizeAdmissionType_(row[COL.admissionType]), f.admissionType)) return false;
  if (!includesAdmissionName_(row[COL.admissionName], f.detailAdmissionName)) return false;
  return true;
}

function toRecord_(row) {
  const stage1Locked = normalize_(row[COL.selectionType]) === '일괄합산' ||
    normalize_(row[COL.stage1]) === '일괄합산';
  return {
    id: makeRecordId_(row),
    version: makeVersion_(row),
    universityType: text_(row[COL.universityType]),
    enrollment: text_(row[COL.enrollment]),
    classNo: text_(row[COL.classNo]),
    studentNo: text_(row[COL.studentNo]),
    name: text_(row[COL.name]),
    university: text_(row[COL.university]),
    track: text_(row[COL.track]),
    admissionType: normalizeAdmissionType_(row[COL.admissionType]),
    admissionName: normalizeAdmissionName_(row[COL.admissionName]),
    birthdate: text_(row[COL.birthdate]),
    examNo: text_(row[COL.examNo]),
    selectionType: text_(row[COL.selectionType]),
    stage1Locked,
    stage1: stage1Locked ? '일괄합산' : text_(row[COL.stage1]),
    finalResult: text_(row[COL.finalResult]),
    failureReason: text_(row[COL.failureReason]),
    firstWait: text_(row[COL.firstWait]),
    finalWait: text_(row[COL.finalWait]),
  };
}

function makeRecordId_(row) {
  const stable = [
    row[COL.universityType], row[COL.enrollment], row[COL.schoolYear], row[COL.classNo],
    row[COL.studentNo], row[COL.name], row[COL.university], row[COL.season],
    row[COL.admissionType], row[COL.selectionType], row[COL.track], row[COL.admissionName],
    row[COL.department], row[COL.originalUnit], row[COL.birthdate], row[COL.examNo],
  ].map(text_).join('\u001f');
  return digestHex_(stable);
}

function makeVersion_(row) {
  const values = [
    row[COL.stage1], row[COL.finalResult], row[COL.failureReason],
    row[COL.firstWait], row[COL.finalWait],
  ].map(text_).join('\u001f');
  return digestHex_(values);
}

function normalizeAdmissionType_(value) {
  const raw = text_(value).trim();
  const compact = normalize_(raw);
  if (compact.includes('학생부위주(교과)') || compact.includes('학생부교과')) return '교과';
  if (compact.includes('학생부위주(종합)') || compact.includes('학생부종합')) return '종합';
  if (compact.includes('논술')) return '논술';
  return raw;
}

function normalizeAdmissionName_(value) {
  let raw = text_(value).trim();
  const wrappers = ['학생부교과', '학생부종합', '논술위주'];
  wrappers.forEach(prefix => {
    const escaped = prefix.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const full = new RegExp('^' + escaped + '\\s*\\((.*)\\)$');
    const match = raw.match(full);
    if (match) raw = match[1].trim();
  });
  raw = raw
    .replace(/학생부교과/g, '')
    .replace(/학생부종합/g, '')
    .replace(/논술위주/g, '')
    .replace(/^\s*[\-–—:：/|]+\s*/, '')
    .replace(/\(\s*\)/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim();
  return raw || text_(value).trim();
}

function includesAdmissionName_(source, query) {
  if (!normalize_(query)) return true;
  return normalize_(source).includes(normalize_(query)) ||
    normalize_(normalizeAdmissionName_(source)).includes(normalize_(query));
}

function validateEnum_(value, allowed, label) {
  const cleaned = cleanText_(value, 40);
  if (allowed.indexOf(cleaned) === -1) {
    throw apiError_('INVALID_VALUE', `${label} 값이 허용 목록에 없습니다.`);
  }
  return cleaned;
}

function validateRank_(value, label) {
  const cleaned = cleanText_(value, 12);
  if (!cleaned) return '';
  if (!/^\d{1,6}$/.test(cleaned)) {
    throw apiError_('INVALID_VALUE', `${label}는 숫자만 입력해 주세요.`);
  }
  return cleaned;
}

function uniqueSorted_(values, numeric) {
  const seen = {};
  values.forEach(value => {
    const text = text_(value).trim();
    if (text) seen[text] = true;
  });
  return Object.keys(seen).sort((a, b) => numeric
    ? Number(a) - Number(b)
    : a.localeCompare(b, 'ko'));
}

function includes_(source, query) {
  const q = normalize_(query);
  return !q || normalize_(source).includes(q);
}

function equals_(source, query) {
  const q = normalize_(query);
  return !q || normalize_(source) === q;
}

function normalize_(value) {
  return text_(value).trim().toLocaleLowerCase('ko-KR').replace(/\s+/g, '');
}

function cleanText_(value, maxLength) {
  return text_(value)
    .replace(/[\u0000-\u001F\u007F]/g, '')
    .trim()
    .slice(0, maxLength);
}

function text_(value) {
  return value === null || value === undefined ? '' : String(value);
}

function digestHex_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  return bytes.map(byte => {
    const unsigned = byte < 0 ? byte + 256 : byte;
    return ('0' + unsigned.toString(16)).slice(-2);
  }).join('');
}

function sign_(payload, secret) {
  const bytes = Utilities.computeHmacSha256Signature(payload, secret, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/g, '');
}

function base64UrlEncode_(value) {
  return Utilities.base64EncodeWebSafe(value, Utilities.Charset.UTF_8).replace(/=+$/g, '');
}

function base64UrlDecode_(value) {
  const padding = '==='.slice((value.length + 3) % 4);
  return Utilities.newBlob(Utilities.base64DecodeWebSafe(value + padding)).getDataAsString('UTF-8');
}

function constantTimeEqual_(a, b) {
  const left = digestHex_(a);
  const right = digestHex_(b);
  let diff = left.length ^ right.length;
  for (let i = 0; i < Math.max(left.length, right.length); i += 1) {
    diff |= (left.charCodeAt(i) || 0) ^ (right.charCodeAt(i) || 0);
  }
  return diff === 0;
}

function apiError_(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}

function safeErrorMessage_(error) {
  if (error && error.message) return String(error.message).slice(0, 240);
  return '알 수 없는 오류가 발생했습니다.';
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
