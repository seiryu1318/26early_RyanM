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
  rowsCachePrefix: 'admissions-rows-v2',
  rowsCacheTtlSeconds: 60,
  rowsCacheChunkChars: 80000,
  rowsCacheMaxChunks: 24,
  dataRevisionProperty: 'DATA_REVISION',
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
const RESULT_KEYS = Object.freeze(['stage1', 'finalResult', 'failureReason', 'firstWait', 'finalWait']);

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
      case 'revision':
        return json_(getRevision_());
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
  const snapshot = getRowsSnapshot_();
  const rows = snapshot.rows;
  return {
    ok: true,
    options: {
      universityTypes: uniqueSorted_(rows.map(r => r[COL.universityType])),
      enrollments: uniqueSorted_(rows.map(r => r[COL.enrollment])),
      classes: uniqueSorted_(rows.map(r => r[COL.classNo]), true),
      tracks: uniqueSorted_(rows.map(r => r[COL.track])),
      admissionTypes: uniqueSorted_(rows.map(r => normalizeAdmissionType_(r[COL.admissionType]))),
      universities: uniqueSorted_(rows.map(r => r[COL.university])),
    },
    rowCount: rows.length,
    revision: snapshot.revision,
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

  const snapshot = getRowsSnapshot_();
  const rows = snapshot.rows;
  const matches = [];
  let totalMatches = 0;

  rows.forEach((row, rowIndex) => {
    if (!matchesFilters_(row, filters)) return;
    totalMatches += 1;
    if (matches.length < CONFIG.maxSearchResults) {
      matches.push(toRecord_(row, CONFIG.firstDataRow + rowIndex));
    }
  });

  return {
    ok: true,
    records: matches,
    totalMatches,
    truncated: totalMatches > matches.length,
    revision: snapshot.revision,
    serverTime: new Date().toISOString(),
  };
}

function getRevision_() {
  return {
    ok: true,
    revision: getDataRevision_(),
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
  const requestedRowNumber = Number(payload.rowNumber);
  const values = payload.values || {};
  if (!recordId) throw apiError_('RECORD_ID_REQUIRED', '저장할 학생 정보가 없습니다.');
  if (!expectedVersion) {
    throw apiError_('VERSION_REQUIRED', '자료 버전이 없습니다. 다시 검색한 뒤 저장해 주세요.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(CONFIG.lockWaitMs);
  try {
    const sheet = getSheet_();
    assertSchema_(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < CONFIG.firstDataRow) throw apiError_('NOT_FOUND', '저장할 행을 찾지 못했습니다.');

    let rowIndex = -1;
    let current = null;
    if (Number.isInteger(requestedRowNumber) && requestedRowNumber >= CONFIG.firstDataRow && requestedRowNumber <= lastRow) {
      const candidate = sheet.getRange(requestedRowNumber, 1, 1, CONFIG.totalColumns).getDisplayValues()[0];
      if (makeRecordId_(candidate) === recordId) {
        rowIndex = requestedRowNumber - CONFIG.firstDataRow;
        current = candidate;
      }
    }

    if (!current) {
      const rows = sheet
        .getRange(CONFIG.firstDataRow, 1, lastRow - CONFIG.headerRow, CONFIG.totalColumns)
        .getDisplayValues();
      for (let i = 0; i < rows.length; i += 1) {
        if (makeRecordId_(rows[i]) === recordId) {
          if (rowIndex !== -1) {
            throw apiError_('DUPLICATE_RECORD', '동일한 학생·대학·전형·수험번호 자료가 중복되어 저장하지 않았습니다.');
          }
          rowIndex = i;
          current = rows[i];
        }
      }
    }
    if (rowIndex === -1) throw apiError_('NOT_FOUND', '자료가 변경되었거나 해당 행을 찾지 못했습니다. 다시 검색해 주세요.');

    const currentVersion = makeVersion_(current);
    const lockedStage1 = normalize_(current[COL.selectionType]) === '일괄합산' ||
      normalize_(current[COL.stage1]) === '일괄합산';
    const requested = {
      stage1: lockedStage1 ? '일괄합산' : validateEnum_(values.stage1, ALLOWED.stage1, '1단계 합격'),
      finalResult: validateEnum_(values.finalResult, ALLOWED.finalResult, '최종 합격'),
      failureReason: validateEnum_(values.failureReason, ALLOWED.failureReason, '불합격 사유'),
      firstWait: validateRank_(values.firstWait, '최초 충원번호'),
      finalWait: validateRank_(values.finalWait, '최종 충원번호'),
    };
    const currentValues = resultValuesFromRow_(current, lockedStage1);
    const mergeRequest = normalizeMergeRequest_(payload, currentValues, requested, lockedStage1);
    const changedKeys = mergeRequest.changedKeys;
    const baseValues = mergeRequest.baseValues;

    if (currentVersion !== expectedVersion && mergeRequest.supportsFieldMerge) {
      const conflictingFields = changedKeys.filter(key => (
        currentValues[key] !== baseValues[key] && currentValues[key] !== requested[key]
      ));
      if (conflictingFields.length) {
        const conflict = apiError_(
          'FIELD_CONFLICT',
          '같은 항목을 다른 사용자가 먼저 저장했습니다. 먼저 반영된 값을 유지했습니다.'
        );
        conflict.details = {
          current: toRecord_(current, CONFIG.firstDataRow + rowIndex),
          conflictingFields,
        };
        throw conflict;
      }
    } else if (currentVersion !== expectedVersion) {
      const conflict = apiError_('CONFLICT', '다른 사용자가 먼저 수정했습니다. 최신 값을 불러왔습니다.');
      conflict.details = { current: toRecord_(current, CONFIG.firstDataRow + rowIndex) };
      throw conflict;
    }

    const merged = { ...currentValues };
    changedKeys.forEach(key => { merged[key] = requested[key]; });
    const cascaded = cascadeStageResult_(
      current[COL.selectionType],
      merged.stage1,
      merged.finalResult,
      merged.failureReason
    );
    merged.finalResult = cascaded.finalResult;
    merged.failureReason = cascaded.failureReason;

    const sheetRow = CONFIG.firstDataRow + rowIndex;
    sheet.getRange(sheetRow, COL.stage1 + 1, 1, 5)
      .setValues([[
        merged.stage1,
        merged.finalResult,
        merged.failureReason,
        merged.firstWait,
        merged.finalWait,
      ]]);
    SpreadsheetApp.flush();
    const revision = bumpDataRevision_();

    current[COL.stage1] = merged.stage1;
    current[COL.finalResult] = merged.finalResult;
    current[COL.failureReason] = merged.failureReason;
    current[COL.firstWait] = merged.firstWait;
    current[COL.finalWait] = merged.finalWait;

    return {
      ok: true,
      record: toRecord_(current, sheetRow),
      revision,
      serverTime: new Date().toISOString(),
    };
  } finally {
    lock.releaseLock();
  }
}

function getAllRows_() {
  return getRowsSnapshot_().rows;
}

function getRowsSnapshot_() {
  const revision = getDataRevision_();
  const cachedRows = readRowsCache_(revision);
  if (cachedRows) return { rows: cachedRows, revision };

  const sheet = getSheet_();
  assertSchema_(sheet);
  const lastRow = sheet.getLastRow();
  const rows = lastRow < CONFIG.firstDataRow
    ? []
    : sheet
      .getRange(CONFIG.firstDataRow, 1, lastRow - CONFIG.headerRow, CONFIG.totalColumns)
      .getDisplayValues();

  if (getDataRevision_() === revision) writeRowsCache_(rows, revision);
  return { rows, revision };
}

function getDataRevision_() {
  return PropertiesService.getScriptProperties().getProperty(CONFIG.dataRevisionProperty) || '0';
}

function bumpDataRevision_() {
  const revision = String(Date.now()) + '-' + Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty(CONFIG.dataRevisionProperty, revision);
  return revision;
}

function rowsCacheKey_(revision, suffix) {
  return CONFIG.rowsCachePrefix + ':' + revision + ':' + suffix;
}

function readRowsCache_(revision) {
  try {
    const cache = CacheService.getScriptCache();
    const metaRaw = cache.get(rowsCacheKey_(revision, 'meta'));
    if (!metaRaw) return null;
    const meta = JSON.parse(metaRaw);
    const chunkCount = Number(meta.chunks);
    if (!Number.isInteger(chunkCount) || chunkCount < 1 || chunkCount > CONFIG.rowsCacheMaxChunks) return null;

    const keys = [];
    for (let index = 0; index < chunkCount; index += 1) {
      keys.push(rowsCacheKey_(revision, String(index)));
    }
    const stored = cache.getAll(keys);
    const encoded = keys.map(key => stored[key] || '').join('');
    if (!encoded || keys.some(key => !stored[key])) return null;

    const zipped = Utilities.base64DecodeWebSafe(encoded);
    const json = Utilities.ungzip(Utilities.newBlob(zipped)).getDataAsString('UTF-8');
    const rows = JSON.parse(json);
    return Array.isArray(rows) ? rows : null;
  } catch (_) {
    return null;
  }
}

function writeRowsCache_(rows, revision) {
  try {
    const json = JSON.stringify(rows);
    const zipped = Utilities.gzip(Utilities.newBlob(json, 'application/json'));
    const encoded = Utilities.base64EncodeWebSafe(zipped.getBytes());
    const chunks = [];
    for (let offset = 0; offset < encoded.length; offset += CONFIG.rowsCacheChunkChars) {
      chunks.push(encoded.slice(offset, offset + CONFIG.rowsCacheChunkChars));
    }
    if (!chunks.length || chunks.length > CONFIG.rowsCacheMaxChunks) return;

    const cache = CacheService.getScriptCache();
    const entries = {};
    chunks.forEach((chunk, index) => {
      entries[rowsCacheKey_(revision, String(index))] = chunk;
    });
    cache.putAll(entries, CONFIG.rowsCacheTtlSeconds);
    cache.put(
      rowsCacheKey_(revision, 'meta'),
      JSON.stringify({ chunks: chunks.length }),
      CONFIG.rowsCacheTtlSeconds
    );
  } catch (_) {
    // 캐시를 사용할 수 없어도 원본 시트 조회는 정상적으로 계속합니다.
  }
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
  if (!includesAdmissionNameRow_(row, f.admissionName)) return false;
  if (!equals_(row[COL.universityType], f.universityType)) return false;
  if (!equals_(row[COL.enrollment], f.enrollment)) return false;
  if (!equals_(row[COL.classNo], f.classNo)) return false;
  if (!includes_(row[COL.name], f.detailName)) return false;
  if (!includes_(row[COL.university], f.detailUniversity)) return false;
  if (!equals_(row[COL.track], f.track)) return false;
  if (!equals_(normalizeAdmissionType_(row[COL.admissionType]), f.admissionType)) return false;
  if (!includesAdmissionNameRow_(row, f.detailAdmissionName)) return false;
  return true;
}

function resultValuesFromRow_(row, lockedStage1) {
  return {
    stage1: lockedStage1 ? '일괄합산' : text_(row[COL.stage1]),
    finalResult: text_(row[COL.finalResult]),
    failureReason: text_(row[COL.failureReason]),
    firstWait: text_(row[COL.firstWait]),
    finalWait: text_(row[COL.finalWait]),
  };
}

function normalizeMergeRequest_(payload, currentValues, requested, lockedStage1) {
  const supportsFieldMerge = payload.baseValues && typeof payload.baseValues === 'object' &&
    Array.isArray(payload.changedKeys);
  const rawKeys = supportsFieldMerge ? payload.changedKeys : RESULT_KEYS;
  const changedKeys = [...new Set(rawKeys
    .map(value => String(value || ''))
    .filter(key => RESULT_KEYS.indexOf(key) !== -1 && !(lockedStage1 && key === 'stage1')))];
  const rawBase = supportsFieldMerge ? payload.baseValues : currentValues;
  const baseValues = {
    stage1: lockedStage1 ? '일괄합산' : text_(rawBase.stage1),
    finalResult: text_(rawBase.finalResult),
    failureReason: text_(rawBase.failureReason),
    firstWait: text_(rawBase.firstWait),
    finalWait: text_(rawBase.finalWait),
  };
  return { supportsFieldMerge, changedKeys, baseValues, requested };
}

function cascadeStageResult_(selectionType, stage1, finalResult, failureReason) {
  const stepwise = normalize_(selectionType).replace(/\s+/g, '').includes('단계별');
  if (stepwise && text_(stage1) === '불합격') {
    return { finalResult: '불합격', failureReason: '1단계 불합격' };
  }
  return { finalResult: text_(finalResult), failureReason: text_(failureReason) };
}

function toRecord_(row, rowNumber) {
  const stage1Locked = normalize_(row[COL.selectionType]) === '일괄합산' ||
    normalize_(row[COL.stage1]) === '일괄합산';
  return {
    id: makeRecordId_(row),
    version: makeVersion_(row),
    rowNumber: Number(rowNumber) || 0,
    universityType: text_(row[COL.universityType]),
    enrollment: text_(row[COL.enrollment]),
    classNo: text_(row[COL.classNo]),
    studentNo: text_(row[COL.studentNo]),
    name: text_(row[COL.name]),
    university: text_(row[COL.university]),
    track: text_(row[COL.track]),
    admissionType: normalizeAdmissionType_(row[COL.admissionType]),
    admissionName: resolveAdmissionName_(row),
    recruitmentUnit: text_(row[COL.department]),
    birthdate: normalizeBirthdate_(row[COL.birthdate]),
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
  const raw = inlineText_(value);
  const compact = normalize_(raw).replace(/（/g, '(').replace(/）/g, ')');
  const slashless = compact.replace(/[\/·ㆍ]/g, '');
  if (compact === '교과' || compact.includes('학생부위주(교과)') || compact.includes('학생부교과')) return '교과';
  if (compact === '종합' || compact.includes('학생부위주(종합)') || compact.includes('학생부종합')) return '종합';
  if (compact.includes('논술')) return '논술';
  if (slashless === '실기실적위주' || slashless === '실기실적') return '실기·실적';
  if (compact === '학생부위주' || compact === '학생부') return '학생부';
  if (compact === '면접위주' || compact === '면접') return '면접';
  return raw;
}

function normalizeAdmissionName_(value, admissionType) {
  const original = inlineText_(value);
  if (!original) return '';

  // 일부 대학의 L열은 전형유형을 대괄호 태그로 한 번 더 붙입니다.
  // 태그만 제거하고 뒤의 공식 전형명과 내부 괄호는 그대로 둡니다.
  const withoutLeadingTag = original.replace(
    /^\[\s*(?:학생부\s*종합|학생부\s*교과|실기\s*(?:[/·ㆍ]\s*)?실적)\s*\]\s*/,
    ''
  );
  if (withoutLeadingTag !== original) return withoutLeadingTag || original;

  const type = normalizeAdmissionType_(admissionType);
  const prefixesByType = {
    '교과': [
      /^학생부\s*위주\s*[\(（]\s*교과\s*[\)）]/,
      /^학생부\s*교과/,
      /^교과/,
    ],
    '종합': [
      /^학생부\s*위주\s*[\(（]\s*종합\s*[\)）]/,
      /^학생부\s*종합/,
      /^종합/,
    ],
    '논술': [
      /^논술\s*위주/,
      /^논술/,
    ],
    '실기·실적': [
      /^실기\s*(?:[/·ㆍ]\s*)?실적\s*위주/,
      /^실기\s*(?:[/·ㆍ]\s*)?실적/,
      /^실기/,
    ],
    '학생부': [
      /^학생부\s*위주/,
      /^학생부/,
    ],
    '면접': [
      /^면접\s*위주/,
      /^면접/,
    ],
  };
  const prefixes = (prefixesByType[type] || []).slice();
  const rawType = inlineText_(admissionType);
  if (rawType) prefixes.push(new RegExp('^' + escapeRegExp_(rawType)));

  for (let index = 0; index < prefixes.length; index += 1) {
    const match = original.match(prefixes[index]);
    if (!match) continue;

    // 전형유형 바깥쪽 한 겹만 제거합니다. 내부의 "전형", Ⅰ/Ⅱ,
    // 서류형/면접형과 중첩 괄호는 공식 명칭의 일부일 수 있어 보존합니다.
    const rest = original.slice(match[0].length);
    if (!/^\s*(?:[\(（\[【{]|[:：\-–—/|·,，;；])/.test(rest)) continue;

    const withoutPrefix = rest
      .replace(/^\s*(?:[:：\-–—/|·,，;；]+\s*)+/, '')
      .trim();
    const unwrapped = unwrapSingleAdmissionWrapper_(withoutPrefix);
    const beginsWithWrapper = /^[\(（\[【{]/.test(withoutPrefix);
    if (beginsWithWrapper && unwrapped === withoutPrefix) return original;
    return unwrapped || original;
  }

  return original;
}

/**
 * 화면에 표시할 전형명은 다음 우선순위를 따릅니다.
 * 1) 대학어디가/대학 공식 모집요강으로 확인한 교정값
 * 2) L열 전형명(중복된 전형유형 표기만 제거)
 * 3) L열이 비어 있을 때에만 N열의 유효한 전형명 조각
 */
function resolveAdmissionName_(row) {
  const university = inlineText_(row[COL.university]);
  const rawName = inlineText_(row[COL.admissionName]);
  const admissionType = row[COL.admissionType];
  const official = resolveOfficialAdmissionName_(university, rawName);
  if (official) return official;

  const nameFromL = normalizeAdmissionName_(rawName, admissionType);
  const nameFromN = extractAdmissionNameFromOriginalUnit_(
    row[COL.originalUnit],
    row[COL.department],
    admissionType
  );
  return crossCheckAdmissionNames_(nameFromL, nameFromN);
}

function crossCheckAdmissionNames_(nameFromL, nameFromN) {
  const primary = inlineText_(nameFromL);
  const reference = inlineText_(nameFromN);
  if (!primary) return reference;
  if (!reference) return primary;

  const primaryKey = normalize_(primary);
  const referenceKey = normalize_(reference);
  if (primaryKey === referenceKey) return primary;

  // N열은 교차 검토와 검색 보조에만 씁니다. 서로 다른 표기라면
  // 사용자 지정 우선순위에 따라 비어 있지 않은 L열을 유지합니다.
  return primary;
}

function resolveOfficialAdmissionName_(university, rawName) {
  const universityKey = normalize_(university);
  const nameKey = normalize_(rawName)
    .replace(/（/g, '(')
    .replace(/）/g, ')')
    .replace(/[–—]/g, '-');
  if (!universityKey || !nameKey) return '';

  // 아래 교정값은 2027학년도 대학어디가/대학 공식 전형계획에서
  // 명칭을 확인한 경우만 둡니다. 전형유형은 별도 열에 있으므로
  // "학생부교과(...)" 같은 바깥 분류어는 표시값에서 제외합니다.
  if (universityKey === '광운대학교') {
    if (nameKey.includes('광운참빛인재전형ⅰ-면접형') || nameKey.includes('광운참빛인재전형i-면접형')) {
      return '광운참빛인재전형Ⅰ-면접형';
    }
    if (nameKey.includes('광운참빛인재전형ⅱ-서류형') || nameKey.includes('광운참빛인재전형ii-서류형')) {
      return '광운참빛인재전형Ⅱ-서류형';
    }
    if (nameKey.includes('소프트웨어우수인재전형')) return '소프트웨어우수인재전형';
  }

  if (universityKey === '세종대학교') {
    if (nameKey.includes('세종창의인재') || nameKey.includes('세종인재')) {
      if (nameKey.includes('면접형')) return '세종인재 전형(면접형)';
      if (nameKey.includes('서류형')) return '세종인재 전형(서류형)';
    }
  }

  if (universityKey === '건국대학교') {
    if (nameKey.includes('ku자기추천')) return 'KU자기추천';
    if (nameKey.includes('ku지역균형')) return 'KU지역균형';
    if (nameKey.includes('ku논술우수자')) return 'KU논술우수자';
  }

  if (universityKey === '동국대학교') {
    if (nameKey.includes('학교장추천인재')) return '학교장추천인재';
    if (nameKey.includes('dodream')) return 'Do Dream';
    if (nameKey.includes('논술')) return '논술';
  }

  if (universityKey === '청운대학교') {
    if (nameKey.includes('일반전형')) return '일반전형';
    if (nameKey.includes('청운인재전형')) return '청운인재전형';
    if (nameKey.includes('지역인재전형')) return '지역인재전형';
  }

  if (universityKey === '국민대학교' && nameKey.includes('국민프런티어')) {
    return '국민프런티어';
  }

  if (universityKey === '경기대학교' && nameKey.includes('kgu학생부종합')) {
    return 'KGU학생부종합전형';
  }

  if (universityKey === '명지대학교') {
    if (nameKey.includes('명지인재면접')) return '명지인재면접전형';
    if (nameKey.includes('명지인재서류')) return '명지인재서류전형';
  }

  if (universityKey === '한성대학교' && nameKey.includes('한성인재')) {
    return '한성인재';
  }

  return '';
}

function extractAdmissionNameFromOriginalUnit_(value, department, admissionType) {
  const source = inlineText_(value);
  if (!source || isUnusableOriginalUnit_(source)) return '';

  const departmentKey = normalize_(department);
  const candidates = source
    .split(/\s*(?:-|–|—|\||\/|>)\s*/)
    .map(inlineText_)
    .filter(Boolean)
    .filter(candidate => normalize_(candidate) !== departmentKey)
    .filter(candidate => !isUnusableOriginalUnit_(candidate))
    .filter(candidate => /(전형|추천|인재|우수자|논술|실기|특기|교과|종합|균형|면접)/.test(candidate));

  if (!candidates.length) return '';
  return normalizeAdmissionName_(candidates[0], admissionType);
}

function isUnusableOriginalUnit_(value) {
  const compact = normalize_(value);
  if (!compact) return true;
  return /^group\d*$/i.test(compact) ||
    compact === '비블라인드' ||
    compact === '수시전형전체' ||
    compact === '수시모집' ||
    /^20\d{2}학년도.*수시모집$/.test(compact);
}

function escapeRegExp_(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function unwrapSingleAdmissionWrapper_(value) {
  const source = inlineText_(value);
  if (!source) return '';

  const closingFor = {
    '(': ')',
    '（': '）',
    '[': ']',
    '【': '】',
    '{': '}',
  };
  const opening = source.charAt(0);
  const expectedClosing = closingFor[opening];
  if (!expectedClosing) return source;

  const stack = [expectedClosing];
  for (let index = 1; index < source.length; index += 1) {
    const character = source.charAt(index);
    if (closingFor[character]) {
      stack.push(closingFor[character]);
    } else if (character === stack[stack.length - 1]) {
      stack.pop();
      if (stack.length === 0) {
        return index === source.length - 1
          ? source.slice(1, -1).trim()
          : source;
      }
    }
  }

  return source;
}

function normalizeBirthdate_(value) {
  const raw = inlineText_(value);
  if (!raw) return '';

  // Preserve a two-digit source year as YYMMDD (for example 08.06.26 ->
  // 080626), while retaining an explicitly supplied four-digit year.
  const parts = raw.match(/\d+/g) || [];
  if (parts.length >= 3 && /^\d{2}$/.test(parts[0]) && /^\d{1,2}$/.test(parts[1]) && /^\d{1,2}$/.test(parts[2])) {
    return parts[0] + parts[1].padStart(2, '0') + parts[2].padStart(2, '0');
  }
  if (parts.length >= 3 && /^\d{4}$/.test(parts[0]) && /^\d{1,2}$/.test(parts[1]) && /^\d{1,2}$/.test(parts[2])) {
    return parts[0] + parts[1].padStart(2, '0') + parts[2].padStart(2, '0');
  }

  const digits = raw.replace(/\D/g, '');
  if (digits.length === 6 || digits.length === 8) return digits;
  return digits.slice(0, 8);
}

function inlineText_(value) {
  return text_(value)
    .replace(/[\r\n\t\f\v]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function includesAdmissionName_(source, query) {
  if (!normalize_(query)) return true;
  return normalize_(source).includes(normalize_(query)) ||
    normalize_(normalizeAdmissionName_(source)).includes(normalize_(query));
}

function includesAdmissionNameRow_(row, query) {
  if (!normalize_(query)) return true;
  const sources = [
    resolveAdmissionName_(row),
    row[COL.admissionName],
    extractAdmissionNameFromOriginalUnit_(
      row[COL.originalUnit],
      row[COL.department],
      row[COL.admissionType]
    ),
  ];
  return sources.some(source => includesAdmissionName_(source, query));
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
