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
  rowsCachePrefix: 'admissions-rows-v3',
  rowsCacheTtlSeconds: 60,
  rowsCacheChunkChars: 80000,
  rowsCacheMaxChunks: 24,
  searchIndexCachePrefix: 'admissions-search-index-v3',
  searchMatchCachePrefix: 'admissions-search-match-v4',
  metaCachePrefix: 'admissions-meta-v3',
  searchMatchCacheTtlSeconds: 45,
  searchCacheMaxValueBytes: 90000,
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
const SEARCH_FILTER_KEYS = Object.freeze([
  'university', 'name', 'admissionName', 'recruitmentUnit', 'universityType', 'enrollment', 'classNo',
  'detailName', 'detailUniversity', 'track', 'admissionType', 'detailAdmissionName',
  'includeCampuses',
]);
const SEARCH_VALUE_FILTER_KEYS = Object.freeze(
  SEARCH_FILTER_KEYS.filter(key => key !== 'includeCampuses')
);
const EXACT_CANDIDATE_FIELDS = Object.freeze([
  'universityType', 'enrollment', 'classNo', 'track', 'admissionType',
]);

let RUNTIME_SEARCH_INDEX_ = null;
const SEARCH_METRICS_ = {
  runtimeIndexBuilds: 0,
  normalizedRowsBuilt: 0,
  candidateMapBuilds: 0,
  persistentUniversityCacheHits: 0,
  matchCacheHits: 0,
  matchCacheMisses: 0,
  metaCacheHits: 0,
};

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
  const cached = readMetaCache_(snapshot.revision, snapshot.snapshotToken);
  if (cached) {
    SEARCH_METRICS_.metaCacheHits += 1;
    return {
      ok: true,
      options: cached.options,
      rowCount: cached.rowCount,
      revision: snapshot.revision,
      serverTime: new Date().toISOString(),
    };
  }

  const rows = snapshot.rows;
  const options = buildMetaOptions_(rows);

  // 메타데이터를 만들 때 이미 모든 행을 훑으므로 다음 검색을 위한
  // exact 후보 맵도 함께 준비합니다. 대학 맵은 ScriptCache에도 저장되어
  // 다음 요청이 다른 Apps Script 인스턴스에 도착해도 재사용됩니다.
  const searchIndex = getRuntimeSearchIndex_(snapshot);
  ensureCandidateMap_(searchIndex, 'university');
  EXACT_CANDIDATE_FIELDS.forEach(field => ensureCandidateMap_(searchIndex, field));

  const result = {
    options,
    rowCount: rows.length,
  };
  writeMetaCache_(snapshot.revision, snapshot.snapshotToken, result);
  return {
    ok: true,
    ...result,
    revision: snapshot.revision,
    serverTime: new Date().toISOString(),
  };
}

function buildMetaOptions_(rows) {
  const buckets = {
    universityTypes: new Set(),
    enrollments: new Set(),
    classes: new Set(),
    tracks: new Set(),
    admissionTypes: new Set(),
    universities: new Set(),
    names: new Set(),
    admissionNames: new Set(),
    recruitmentUnits: new Set(),
  };
  const add = (bucket, value) => {
    const cleaned = text_(value).trim();
    if (cleaned) bucket.add(cleaned);
  };
  rows.forEach(row => {
    add(buckets.universityTypes, row[COL.universityType]);
    add(buckets.enrollments, row[COL.enrollment]);
    add(buckets.classes, row[COL.classNo]);
    add(buckets.tracks, row[COL.track]);
    add(buckets.admissionTypes, normalizeAdmissionType_(row[COL.admissionType]));
    add(buckets.universities, row[COL.university]);
    add(buckets.names, row[COL.name]);
    add(buckets.admissionNames, resolveAdmissionName_(row));
    add(buckets.recruitmentUnits, row[COL.department]);
  });
  const sorted = bucket => Array.from(bucket).sort((a, b) => a.localeCompare(b, 'ko'));
  return {
    universityTypes: sorted(buckets.universityTypes),
    enrollments: sorted(buckets.enrollments),
    classes: Array.from(buckets.classes).sort((a, b) => Number(a) - Number(b)),
    tracks: sorted(buckets.tracks),
    admissionTypes: sorted(buckets.admissionTypes),
    universities: sorted(buckets.universities),
    names: sorted(buckets.names),
    admissionNames: sorted(buckets.admissionNames),
    recruitmentUnits: sorted(buckets.recruitmentUnits),
  };
}

function getUniversitySuggestions_(query) {
  const q = normalizeUniversity_(query);
  if ([...q].length < 2) {
    return { ok: true, suggestions: [], serverTime: new Date().toISOString() };
  }

  const meta = getMeta_();
  const universities = meta.options.universities || [];

  const suggestions = universities
    .filter(name => normalizeUniversity_(name).includes(q))
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
  const match = findMatchingRowIndexes_(snapshot, filters);
  const visibleIndexes = match.rowIndexes.slice(0, CONFIG.maxSearchResults);
  const matches = visibleIndexes.map(rowIndex => (
    toRecord_(rows[rowIndex], CONFIG.firstDataRow + rowIndex)
  ));
  const totalMatches = match.rowIndexes.length;

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

  const snapshot = getRowsSnapshot_();
  const match = findMatchingRowIndexes_(snapshot, filters);
  const rows = match.rowIndexes.map(rowIndex => snapshot.rows[rowIndex]);
  return {
    ok: true,
    headers: EXPECTED_HEADERS.slice(),
    rows,
    totalMatches: rows.length,
    revision: snapshot.revision,
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
    const effectiveMerge = effectiveMergeRequest_(
      current[COL.selectionType],
      currentValues,
      mergeRequest
    );
    const effectiveChangedKeys = effectiveMerge.changedKeys;
    const effectiveRequested = effectiveMerge.requested;

    if (currentVersion !== expectedVersion && mergeRequest.supportsFieldMerge) {
      const conflictingFields = effectiveChangedKeys.filter(key => (
        currentValues[key] !== baseValues[key] && currentValues[key] !== effectiveRequested[key]
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
  const cached = readRowsCache_(revision);
  if (cached) {
    return {
      rows: cached.rows,
      revision,
      snapshotToken: cached.snapshotToken,
    };
  }

  const sheet = getSheet_();
  assertSchema_(sheet);
  const lastRow = sheet.getLastRow();
  const rows = lastRow < CONFIG.firstDataRow
    ? []
    : sheet
      .getRange(CONFIG.firstDataRow, 1, lastRow - CONFIG.headerRow, CONFIG.totalColumns)
      .getDisplayValues();

  const snapshotToken = createRowsSnapshotToken_(revision);
  if (getDataRevision_() === revision) writeRowsCache_(rows, revision, snapshotToken);
  return { rows, revision, snapshotToken };
}

function getDataRevision_() {
  return PropertiesService.getScriptProperties().getProperty(CONFIG.dataRevisionProperty) || '0';
}

function bumpDataRevision_() {
  const revision = String(Date.now()) + '-' + Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty(CONFIG.dataRevisionProperty, revision);
  return revision;
}

/**
 * Google Sheet를 직접 일괄 갱신한 뒤 관리자가 한 번 실행하는 발행 함수입니다.
 * 공개 웹 API에는 노출하지 않으며 Apps Script 편집기 또는 Execution API에서만
 * 실행해 열린 PC·모바일 화면이 새 데이터 세대를 즉시 감지하게 합니다.
 */
function publishAdmissionsDataRevision() {
  const lock = LockService.getScriptLock();
  lock.waitLock(CONFIG.lockWaitMs);
  try {
    const sheet = getSheet_();
    assertSchema_(sheet);
    SpreadsheetApp.flush();
    const revision = bumpDataRevision_();
    return {
      ok: true,
      revision,
      rowCount: Math.max(0, sheet.getLastRow() - CONFIG.headerRow),
      serverTime: new Date().toISOString(),
    };
  } finally {
    lock.releaseLock();
  }
}

function rowsCacheKey_(revision, suffix) {
  return CONFIG.rowsCachePrefix + ':' + revision + ':' + suffix;
}

function createRowsSnapshotToken_(revision) {
  return digestHex_(
    text_(revision) + '\u001e' + String(Date.now()) + '\u001e' + Utilities.getUuid()
  );
}

function rowsCacheChunkKey_(revision, snapshotToken, index) {
  return rowsCacheKey_(revision, 'snapshot:' + digestHex_(snapshotToken) + ':' + String(index));
}

function readRowsCache_(revision) {
  try {
    const cache = CacheService.getScriptCache();
    const metaRaw = cache.get(rowsCacheKey_(revision, 'meta'));
    if (!metaRaw) return null;
    const meta = JSON.parse(metaRaw);
    const chunkCount = Number(meta.chunks);
    const snapshotToken = text_(meta.snapshotToken);
    if (!Number.isInteger(chunkCount) || chunkCount < 1 || chunkCount > CONFIG.rowsCacheMaxChunks) return null;
    if (!snapshotToken) return null;

    const keys = [];
    for (let index = 0; index < chunkCount; index += 1) {
      keys.push(rowsCacheChunkKey_(revision, snapshotToken, index));
    }
    const stored = cache.getAll(keys);
    const encoded = keys.map(key => stored[key] || '').join('');
    if (!encoded || keys.some(key => !stored[key])) return null;

    const zipped = Utilities.base64DecodeWebSafe(encoded);
    const json = Utilities.ungzip(Utilities.newBlob(zipped)).getDataAsString('UTF-8');
    const rows = JSON.parse(json);
    return Array.isArray(rows) ? { rows, snapshotToken } : null;
  } catch (_) {
    return null;
  }
}

function writeRowsCache_(rows, revision, snapshotToken) {
  try {
    if (!snapshotToken) return;
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
      entries[rowsCacheChunkKey_(revision, snapshotToken, index)] = chunk;
    });
    cache.putAll(entries, CONFIG.rowsCacheTtlSeconds);
    cache.put(
      rowsCacheKey_(revision, 'meta'),
      JSON.stringify({ chunks: chunks.length, snapshotToken }),
      CONFIG.rowsCacheTtlSeconds
    );
  } catch (_) {
    // 캐시를 사용할 수 없어도 원본 시트 조회는 정상적으로 계속합니다.
  }
}

function cacheValueFits_(value) {
  // UTF-8에서 한 UTF-16 code unit은 최대 3바이트이며 surrogate pair는
  // 2 code unit/4바이트이므로 아래 상한은 CacheService 100KB 제한에 보수적입니다.
  return text_(value).length * 3 <= CONFIG.searchCacheMaxValueBytes;
}

function putCacheSafely_(key, value, ttlSeconds) {
  try {
    if (!cacheValueFits_(value)) return false;
    CacheService.getScriptCache().put(key, value, ttlSeconds);
    return true;
  } catch (_) {
    return false;
  }
}

function metaCacheKey_(revision, snapshotToken) {
  return CONFIG.metaCachePrefix + ':' + digestHex_(revision + '\u001e' + snapshotToken);
}

function readMetaCache_(revision, snapshotToken) {
  try {
    const raw = CacheService.getScriptCache().get(metaCacheKey_(revision, snapshotToken));
    if (!raw) return null;
    const zipped = Utilities.base64DecodeWebSafe(raw);
    const json = Utilities.ungzip(Utilities.newBlob(zipped)).getDataAsString('UTF-8');
    const parsed = JSON.parse(json);
    if (!parsed || typeof parsed !== 'object' || !parsed.options || !Number.isInteger(parsed.rowCount)) {
      return null;
    }
    return parsed;
  } catch (_) {
    return null;
  }
}

function writeMetaCache_(revision, snapshotToken, value) {
  const json = JSON.stringify(value);
  const zipped = Utilities.gzip(Utilities.newBlob(json, 'application/json'));
  const encoded = Utilities.base64EncodeWebSafe(zipped.getBytes());
  putCacheSafely_(
    metaCacheKey_(revision, snapshotToken),
    encoded,
    CONFIG.rowsCacheTtlSeconds
  );
}

function universityIndexCacheKey_(revision, snapshotToken) {
  return CONFIG.searchIndexCachePrefix + ':' +
    digestHex_(revision + '\u001e' + snapshotToken) + ':university';
}

function readPersistentUniversityMap_(revision, snapshotToken, rowCount) {
  try {
    const raw = CacheService.getScriptCache().get(
      universityIndexCacheKey_(revision, snapshotToken)
    );
    if (!raw) return null;
    const entries = JSON.parse(raw);
    if (!Array.isArray(entries)) return null;
    const map = new Map();
    for (let entryIndex = 0; entryIndex < entries.length; entryIndex += 1) {
      const entry = entries[entryIndex];
      if (!Array.isArray(entry) || typeof entry[0] !== 'string' || !Array.isArray(entry[1])) return null;
      const indexes = entry[1];
      if (indexes.some(index => !Number.isInteger(index) || index < 0 || index >= rowCount)) return null;
      map.set(entry[0], indexes);
    }
    SEARCH_METRICS_.persistentUniversityCacheHits += 1;
    return map;
  } catch (_) {
    return null;
  }
}

function writePersistentUniversityMap_(revision, snapshotToken, map) {
  putCacheSafely_(
    universityIndexCacheKey_(revision, snapshotToken),
    JSON.stringify(Array.from(map.entries())),
    CONFIG.rowsCacheTtlSeconds
  );
}

function getRuntimeSearchIndex_(snapshot) {
  const now = Date.now();
  const maxAgeMs = CONFIG.rowsCacheTtlSeconds * 1000;
  if (
    RUNTIME_SEARCH_INDEX_ &&
    RUNTIME_SEARCH_INDEX_.revision === snapshot.revision &&
    RUNTIME_SEARCH_INDEX_.snapshotToken === snapshot.snapshotToken &&
    RUNTIME_SEARCH_INDEX_.rows.length === snapshot.rows.length &&
    now - RUNTIME_SEARCH_INDEX_.createdAt < maxAgeMs
  ) {
    return RUNTIME_SEARCH_INDEX_;
  }

  RUNTIME_SEARCH_INDEX_ = {
    revision: snapshot.revision,
    snapshotToken: snapshot.snapshotToken,
    rows: snapshot.rows,
    createdAt: now,
    normalizedRows: new Array(snapshot.rows.length),
    candidateMaps: Object.create(null),
    allRowIndexes: Array.from({ length: snapshot.rows.length }, (_, index) => index),
  };
  SEARCH_METRICS_.runtimeIndexBuilds += 1;
  return RUNTIME_SEARCH_INDEX_;
}

function candidateValue_(row, field) {
  switch (field) {
    case 'university': return normalizeUniversity_(row[COL.university]);
    case 'universityType': return normalize_(row[COL.universityType]);
    case 'enrollment': return normalize_(row[COL.enrollment]);
    case 'classNo': return normalize_(row[COL.classNo]);
    case 'track': return normalize_(row[COL.track]);
    case 'admissionType': return normalize_(normalizeAdmissionType_(row[COL.admissionType]));
    default: return '';
  }
}

function ensureCandidateMap_(searchIndex, field) {
  if (searchIndex.candidateMaps[field]) return searchIndex.candidateMaps[field];

  if (field === 'university') {
    const persistent = readPersistentUniversityMap_(
      searchIndex.revision,
      searchIndex.snapshotToken,
      searchIndex.rows.length
    );
    if (persistent) {
      searchIndex.candidateMaps[field] = persistent;
      return persistent;
    }
  }

  const map = new Map();
  searchIndex.rows.forEach((row, rowIndex) => {
    const value = candidateValue_(row, field);
    if (!value) return;
    if (!map.has(value)) map.set(value, []);
    map.get(value).push(rowIndex);
  });
  searchIndex.candidateMaps[field] = map;
  SEARCH_METRICS_.candidateMapBuilds += 1;
  if (field === 'university') {
    writePersistentUniversityMap_(searchIndex.revision, searchIndex.snapshotToken, map);
  }
  return map;
}

function normalizedSearchFilters_(filters) {
  const normalized = {};
  SEARCH_FILTER_KEYS.forEach(key => {
    if (key === 'includeCampuses') {
      normalized[key] = filters[key] === true;
    } else if (key === 'university' || key === 'detailUniversity') {
      normalized[key] = normalizeUniversity_(filters[key]);
    } else {
      normalized[key] = normalize_(filters[key]);
    }
  });
  return normalized;
}

function normalizedFilterSignature_(filters) {
  return SEARCH_FILTER_KEYS.map(key => (
    key === 'includeCampuses' ? (filters[key] ? '1' : '0') : (filters[key] || '')
  )).join('\u001f');
}

function searchMatchCacheKey_(revision, snapshotToken, normalizedFilters) {
  const signature = revision + '\u001e' + snapshotToken + '\u001e' +
    normalizedFilterSignature_(normalizedFilters);
  return CONFIG.searchMatchCachePrefix + ':' + digestHex_(signature);
}

function readMatchCache_(revision, snapshotToken, normalizedFilters, rowCount) {
  try {
    const raw = CacheService.getScriptCache().get(
      searchMatchCacheKey_(revision, snapshotToken, normalizedFilters)
    );
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.rowIndexes)) return null;
    if (parsed.rowIndexes.some(index => !Number.isInteger(index) || index < 0 || index >= rowCount)) return null;
    return {
      rowIndexes: parsed.rowIndexes,
      candidateCount: Number(parsed.candidateCount) || parsed.rowIndexes.length,
    };
  } catch (_) {
    return null;
  }
}

function writeMatchCache_(revision, snapshotToken, normalizedFilters, match) {
  putCacheSafely_(
    searchMatchCacheKey_(revision, snapshotToken, normalizedFilters),
    JSON.stringify({ rowIndexes: match.rowIndexes, candidateCount: match.candidateCount }),
    CONFIG.searchMatchCacheTtlSeconds
  );
}

function normalizeUniversity_(value) {
  let normalized = text_(value);
  try {
    normalized = normalized.normalize('NFKC');
  } catch (_) {
    // Apps Script V8에서는 지원되지만, 구형 런타임에서도 검색은 계속합니다.
  }
  return normalized
    .trim()
    .toLocaleLowerCase('ko-KR')
    .replace(/[\s\u200B-\u200D\u2060\uFEFF]+/g, '')
    .replace(/[（]/g, '(')
    .replace(/[）]/g, ')');
}

function universityIdentity_(value) {
  const key = normalizeUniversity_(value);
  if (!key) {
    return { key: '', baseKey: '', alias: '', isCampus: false, campusAlias: '' };
  }

  let baseKey = key;
  let isCampus = false;
  let campusName = '';
  const parenthesized = key.match(/^(.+?)\([^()]+\)$/);
  if (parenthesized) {
    baseKey = parenthesized[1];
    campusName = key.slice(parenthesized[1].length + 1, -1);
    isCampus = true;
  } else {
    // 괄호 없이 "미래캠퍼스", "세종캠퍼스"처럼 적힌 분교 표기도
    // 본교명과 같은 묶음으로 처리합니다.
    const namedCampus = key.match(/^(.+?대학교).+캠퍼스$/);
    if (namedCampus) {
      baseKey = namedCampus[1];
      campusName = key.slice(namedCampus[1].length, -'캠퍼스'.length);
      isCampus = true;
    }
  }

  // "건국", "건국대", "건국대학교"를 같은 본교 별칭으로 취급합니다.
  const alias = baseKey.replace(/대학교$/, '').replace(/대학$/, '').replace(/대$/, '');
  const normalizedCampusName = campusName.replace(/캠퍼스$/, '');
  const campusAliases = {
    세: '세',
    세종: '세',
    글: '글',
    글로컬: '글',
    미: '미',
    미래: '미',
  };
  const campusAlias = campusAliases[normalizedCampusName] || normalizedCampusName;
  return { key, baseKey, alias, isCampus, campusAlias };
}

function universityMatchesQuery_(university, query, includeCampuses) {
  const source = universityIdentity_(university);
  const requested = universityIdentity_(query);
  if (!requested.key) return true;
  if (source.key === requested.key) return true;

  // 사용자가 특정 캠퍼스명을 직접 입력했다면 그 캠퍼스만 찾습니다. 시트의
  // (세)/(글)과 사용자가 쓰는 (세종)/(글로컬) 표기도 같은 값으로 봅니다.
  if (requested.isCampus) {
    return source.isCampus && source.alias === requested.alias &&
      source.campusAlias === requested.campusAlias;
  }
  if (!requested.alias || source.alias !== requested.alias) return false;
  return includeCampuses === true || !source.isCampus;
}

function universityCandidateIndexes_(map, query, includeCampuses) {
  const requested = universityIdentity_(query);
  if (!requested.key) return [];

  // 정확한 대학명/캠퍼스명이 들어온 일반 검색은 Map의 단일 bucket을 바로
  // 사용합니다. 대학교를 바꿔 검색할 때 전체 행을 다시 훑지 않습니다.
  if (map.has(requested.key) && (!includeCampuses || requested.isCampus)) {
    return map.get(requested.key).slice();
  }

  const indexes = [];
  map.forEach((rowIndexes, university) => {
    if (universityMatchesQuery_(university, requested.key, includeCampuses)) {
      indexes.push(...rowIndexes);
    }
  });
  indexes.sort((a, b) => a - b);
  return indexes;
}

function candidateRowIndexes_(searchIndex, filters) {
  const lists = [];
  const universityMapRequired = filters.university || filters.detailUniversity;
  if (universityMapRequired) {
    const universityMap = ensureCandidateMap_(searchIndex, 'university');
    if (filters.university) {
      lists.push(universityCandidateIndexes_(universityMap, filters.university, filters.includeCampuses));
    }
    if (filters.detailUniversity) {
      lists.push(universityCandidateIndexes_(universityMap, filters.detailUniversity, filters.includeCampuses));
    }
  }

  EXACT_CANDIDATE_FIELDS.forEach(field => {
    const query = filters[field];
    if (!query) return;
    const map = ensureCandidateMap_(searchIndex, field);
    lists.push(map.get(query) || []);
  });

  if (!lists.length) return searchIndex.allRowIndexes;
  if (lists.some(list => list.length === 0)) return [];
  lists.sort((a, b) => a.length - b.length);
  const remaining = lists.slice(1).map(list => new Set(list));
  return lists[0].filter(rowIndex => remaining.every(set => set.has(rowIndex)));
}

function normalizedAdmissionCandidates_(row) {
  const sources = [
    resolveAdmissionName_(row),
    row[COL.admissionName],
    extractAdmissionNameFromOriginalUnit_(
      row[COL.originalUnit],
      row[COL.department],
      row[COL.admissionType]
    ),
  ];
  const seen = Object.create(null);
  const candidates = [];
  sources.forEach(source => {
    [source, normalizeAdmissionName_(source)].forEach(value => {
      const normalized = normalize_(value);
      if (normalized && !seen[normalized]) {
        seen[normalized] = true;
        candidates.push(normalized);
      }
    });
  });
  return candidates;
}

function getNormalizedSearchRow_(searchIndex, rowIndex) {
  if (searchIndex.normalizedRows[rowIndex]) return searchIndex.normalizedRows[rowIndex];
  const row = searchIndex.rows[rowIndex];
  const normalized = {
    university: normalizeUniversity_(row[COL.university]),
    name: normalize_(row[COL.name]),
    recruitmentUnit: normalize_(row[COL.department]),
    universityType: normalize_(row[COL.universityType]),
    enrollment: normalize_(row[COL.enrollment]),
    classNo: normalize_(row[COL.classNo]),
    track: normalize_(row[COL.track]),
    admissionType: normalize_(normalizeAdmissionType_(row[COL.admissionType])),
    admissionNames: null,
  };
  searchIndex.normalizedRows[rowIndex] = normalized;
  SEARCH_METRICS_.normalizedRowsBuilt += 1;
  return normalized;
}

function matchesNormalizedFilters_(searchIndex, rowIndex, filters) {
  const row = getNormalizedSearchRow_(searchIndex, rowIndex);
  if (filters.university && !universityMatchesQuery_(
    row.university,
    filters.university,
    filters.includeCampuses
  )) return false;
  if (filters.name && !row.name.includes(filters.name)) return false;
  if (filters.recruitmentUnit && !row.recruitmentUnit.includes(filters.recruitmentUnit)) return false;
  if (filters.universityType && row.universityType !== filters.universityType) return false;
  if (filters.enrollment && row.enrollment !== filters.enrollment) return false;
  if (filters.classNo && row.classNo !== filters.classNo) return false;
  if (filters.detailName && !row.name.includes(filters.detailName)) return false;
  if (filters.detailUniversity && !universityMatchesQuery_(
    row.university,
    filters.detailUniversity,
    filters.includeCampuses
  )) return false;
  if (filters.track && row.track !== filters.track) return false;
  if (filters.admissionType && row.admissionType !== filters.admissionType) return false;

  if (filters.admissionName || filters.detailAdmissionName) {
    if (!row.admissionNames) row.admissionNames = normalizedAdmissionCandidates_(searchIndex.rows[rowIndex]);
    if (filters.admissionName && !row.admissionNames.some(value => value.includes(filters.admissionName))) return false;
    if (filters.detailAdmissionName && !row.admissionNames.some(value => value.includes(filters.detailAdmissionName))) return false;
  }
  return true;
}

function findMatchingRowIndexes_(snapshot, filters) {
  const normalizedFilters = normalizedSearchFilters_(filters);
  const cached = readMatchCache_(
    snapshot.revision,
    snapshot.snapshotToken,
    normalizedFilters,
    snapshot.rows.length
  );
  if (cached) {
    SEARCH_METRICS_.matchCacheHits += 1;
    return { ...cached, cacheHit: true };
  }

  SEARCH_METRICS_.matchCacheMisses += 1;
  const searchIndex = getRuntimeSearchIndex_(snapshot);
  const candidates = candidateRowIndexes_(searchIndex, normalizedFilters);
  const rowIndexes = candidates.filter(rowIndex => (
    matchesNormalizedFilters_(searchIndex, rowIndex, normalizedFilters)
  ));
  const result = { rowIndexes, candidateCount: candidates.length, cacheHit: false };
  writeMatchCache_(snapshot.revision, snapshot.snapshotToken, normalizedFilters, result);
  return result;
}

function getSearchMetricsForTest_() {
  return { ...SEARCH_METRICS_ };
}

function resetSearchRuntimeForTest_() {
  RUNTIME_SEARCH_INDEX_ = null;
  Object.keys(SEARCH_METRICS_).forEach(key => { SEARCH_METRICS_[key] = 0; });
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
  const source = filters && typeof filters === 'object' ? filters : {};
  return {
    university: cleanText_(source.university, 80),
    name: cleanText_(source.name, 40),
    admissionName: cleanText_(source.admissionName, 100),
    recruitmentUnit: cleanText_(source.recruitmentUnit, 120),
    universityType: cleanText_(source.universityType, 40),
    enrollment: cleanText_(source.enrollment, 40),
    classNo: cleanText_(source.classNo, 20),
    detailName: cleanText_(source.detailName, 40),
    detailUniversity: cleanText_(source.detailUniversity, 80),
    track: cleanText_(source.track, 40),
    admissionType: cleanText_(source.admissionType, 40),
    detailAdmissionName: cleanText_(source.detailAdmissionName, 100),
    includeCampuses: cleanBoolean_(source.includeCampuses),
  };
}

function hasSearchCondition_(filters) {
  return SEARCH_VALUE_FILTER_KEYS.some(key => normalize_(filters[key]).length > 0);
}

function matchesFilters_(row, f) {
  if (!universityMatchesQuery_(row[COL.university], f.university, f.includeCampuses)) return false;
  if (!includes_(row[COL.name], f.name)) return false;
  if (!includesAdmissionNameRow_(row, f.admissionName)) return false;
  if (!includes_(row[COL.department], f.recruitmentUnit)) return false;
  if (!equals_(row[COL.universityType], f.universityType)) return false;
  if (!equals_(row[COL.enrollment], f.enrollment)) return false;
  if (!equals_(row[COL.classNo], f.classNo)) return false;
  if (!includes_(row[COL.name], f.detailName)) return false;
  if (!universityMatchesQuery_(row[COL.university], f.detailUniversity, f.includeCampuses)) return false;
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

function effectiveMergeRequest_(selectionType, currentValues, mergeRequest) {
  const preview = { ...currentValues };
  mergeRequest.changedKeys.forEach(key => { preview[key] = mergeRequest.requested[key]; });
  const cascaded = cascadeStageResult_(
    selectionType,
    preview.stage1,
    preview.finalResult,
    preview.failureReason
  );
  const changedKeys = mergeRequest.changedKeys.slice();
  if (isStepwiseStageFailure_(selectionType, preview.stage1)) {
    ['finalResult', 'failureReason'].forEach(key => {
      if (changedKeys.indexOf(key) === -1) changedKeys.push(key);
    });
  }
  return {
    changedKeys,
    requested: {
      ...mergeRequest.requested,
      finalResult: cascaded.finalResult,
      failureReason: cascaded.failureReason,
    },
  };
}

function isStepwiseStageFailure_(selectionType, stage1) {
  return normalize_(selectionType).replace(/\s+/g, '').includes('단계별') &&
    text_(stage1) === '불합격';
}

function cascadeStageResult_(selectionType, stage1, finalResult, failureReason) {
  if (isStepwiseStageFailure_(selectionType, stage1)) {
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

function cleanBoolean_(value) {
  if (value === true || value === 1) return true;
  const normalized = normalize_(value);
  return normalized === 'true' || normalized === '1' || normalized === 'on' || normalized === 'yes';
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
