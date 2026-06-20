/**
 * ИУС 2.0 — бесплатная общая база работ
 * Google Таблица + Google Apps Script + Google Drive
 *
 * Что делает:
 * 1) ученик отправляет работу с сайта;
 * 2) строка сразу появляется в листе works;
 * 3) аудио сохраняется в Google Drive;
 * 4) учитель видит работы на expert.html.
 *
 * ВАЖНО:
 * - Apps Script должен быть открыт из Google Таблицы: Расширения → Apps Script.
 * - После вставки кода запустите setupIUS, потом authorizeIUS, потом сделайте Новая версия → Развернуть.
 * - Развертывание веб-приложения: Выполнять от имени: Меня; Доступ: Все.
 */

const IUS_SHEET_NAME = 'works';
const IUS_FOLDER_NAME = 'ИУС 2.0 — аудиозаписи учеников';

const IUS_HEADERS = [
  'id', 'createdAt', 'sentAt', 'status',
  'studentFullName', 'studentClassName', 'studentLogin',
  'taskId', 'taskTitle', 'note',
  'audioFileName', 'audioMimeType', 'audioSize', 'audioFileId', 'audioFileUrl', 'audioDirectUrl',
  'checkedAt', 'teacherFullName', 'total', 'verdict', 'teacherComment', 'selectedJson',
  'rawJson', 'audioError'
];

/**
 * 1. Запустите эту функцию первой.
 * Она создаёт лист works и папку в Drive.
 */
function setupIUS() {
  const sheet = ensureSheet_();
  const folder = getOrCreateAudioFolder_();
  return 'Готово: лист ' + sheet.getName() + ', папка Drive: ' + folder.getName();
}

/**
 * 2. Запустите эту функцию второй.
 * Она принудительно просит разрешение на DriveApp.
 * Без этого аудио не сможет сохраняться в Google Drive.
 */
function authorizeIUS() {
  const sheet = ensureSheet_();
  const folder = getOrCreateAudioFolder_();

  const blob = Utilities.newBlob(
    'Проверка доступа ИУС 2.0 к Google Drive: ' + new Date().toISOString(),
    'text/plain',
    'ius-drive-test.txt'
  );
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = file.getUrl();
  file.setTrashed(true);

  sheet.getRange('A1').setNote('DriveApp разрешён: ' + new Date().toLocaleString('ru-RU'));
  Logger.log('DriveApp разрешён. Тестовый файл создан и удалён: ' + url);
  return 'DriveApp разрешён. Теперь сделайте: Развернуть → Управление развертываниями → Новая версия → Развернуть.';
}

/**
 * 3. Диагностика. Запустите, если аудио не сохраняется.
 */
function diagnoseIUS() {
  const result = {
    ok: true,
    time: new Date().toISOString(),
    spreadsheet: false,
    sheet: false,
    drive: false,
    folderName: IUS_FOLDER_NAME,
    message: ''
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    result.spreadsheet = Boolean(ss);
    result.spreadsheetName = ss ? ss.getName() : '';
    const sheet = ensureSheet_();
    result.sheet = Boolean(sheet);
  } catch (err) {
    result.ok = false;
    result.message = 'Ошибка таблицы: ' + getErrorMessage_(err);
    Logger.log(JSON.stringify(result, null, 2));
    return result;
  }

  try {
    const folder = getOrCreateAudioFolder_();
    const testFile = folder.createFile(Utilities.newBlob('drive ok', 'text/plain', 'ius-diagnose.txt'));
    result.drive = true;
    result.folderUrl = folder.getUrl();
    testFile.setTrashed(true);
  } catch (err) {
    result.ok = false;
    result.drive = false;
    result.message = 'Ошибка DriveApp: ' + getErrorMessage_(err);
    Logger.log(JSON.stringify(result, null, 2));
    return result;
  }

  result.message = 'Всё готово: таблица и DriveApp работают.';
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Тест: добавляет строку без аудио.
 */
function testAddRow() {
  return saveSubmission_({
    app: 'saha-exam-v2',
    type: 'student-submission',
    version: 10,
    id: makeId_('test'),
    createdAt: new Date().toISOString(),
    sentAt: new Date().toISOString(),
    student: { fullName: 'Тестовый ученик', className: '9А', login: 'test' },
    task: { id: 1, title: 'Проверка таблицы', instruction: '' },
    note: 'Тестовая строка без аудио',
    audio: { fileName: '', mimeType: '', size: 0 }
  });
}

/**
 * Тест: добавляет маленький тестовый аудиофайл в Drive.
 */
function testAudioSave() {
  return saveSubmission_({
    app: 'saha-exam-v2',
    type: 'student-submission',
    version: 10,
    id: makeId_('audio-test'),
    createdAt: new Date().toISOString(),
    sentAt: new Date().toISOString(),
    student: { fullName: 'Тест аудио', className: '9А', login: 'audio-test' },
    task: { id: 1, title: 'Проверка сохранения аудио', instruction: '' },
    note: 'Тестовый маленький dataUrl',
    audio: {
      fileName: 'audio-test.webm',
      mimeType: 'audio/webm;codecs=opus',
      size: 32,
      dataUrl: 'data:audio/webm;codecs=opus;base64,AAAAIGZ0eXBpc29tAAACAGlzb21pc28ybXA0MQ=='
    }
  });
}

function doPost(e) {
  try {
    const payload = parsePayload_(e);

    if (payload.action === 'submit') {
      const result = saveSubmission_(payload.submission || {});
      return json_({ ok: true, id: result.id, audioError: result.audioError || '', audioFileUrl: result.audioFileUrl || '' });
    }

    if (payload.action === 'saveResult') {
      saveResult_(payload.submissionId, payload.result || {});
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'Неизвестное действие: ' + (payload.action || '') });
  } catch (err) {
    return json_({ ok: false, error: getErrorMessage_(err) });
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    let data;

    if (p.action === 'ping') {
      data = {
        ok: true,
        message: 'Apps Script работает',
        now: new Date().toISOString(),
        count: getDataRows_().length
      };
    } else if (p.action === 'diag') {
      data = diagnoseForWeb_();
    } else if (p.action === 'list') {
      data = { ok: true, items: listSubmissions_(Number(p.limit || 30)) };
    } else if (p.action === 'get') {
      data = { ok: true, item: getSubmission_(p.id) };
    } else if (p.action === 'audioData') {
      data = getAudioData_(p.id);
    } else {
      data = { ok: false, error: 'Неизвестное действие: ' + (p.action || '') };
    }

    return jsonpOrJson_(data, p.callback);
  } catch (err) {
    return jsonpOrJson_({ ok: false, error: getErrorMessage_(err) }, p.callback);
  }
}

function diagnoseForWeb_() {
  const result = {
    ok: true,
    message: 'Проверка выполнена',
    sheetOk: false,
    driveOk: false,
    count: 0,
    folderUrl: ''
  };

  try {
    ensureSheet_();
    result.sheetOk = true;
    result.count = getDataRows_().length;
  } catch (err) {
    result.ok = false;
    result.message = 'Ошибка таблицы: ' + getErrorMessage_(err);
    return result;
  }

  try {
    const folder = getOrCreateAudioFolder_();
    result.driveOk = true;
    result.folderUrl = folder.getUrl();
  } catch (err) {
    result.ok = false;
    result.driveOk = false;
    result.message = 'Ошибка DriveApp: ' + getErrorMessage_(err) + '. Запустите authorizeIUS вручную в Apps Script.';
    return result;
  }

  result.message = 'Таблица и Google Drive работают. Работ в базе: ' + result.count;
  return result;
}

function parsePayload_(e) {
  if (e && e.parameter && e.parameter.payload) {
    return JSON.parse(e.parameter.payload || '{}');
  }

  if (e && e.postData && e.postData.contents) {
    const contents = e.postData.contents || '';
    if (contents.indexOf('payload=') === 0) {
      const decoded = decodeURIComponent(contents.replace(/^payload=/, '').replace(/\+/g, ' '));
      return JSON.parse(decoded || '{}');
    }
    return JSON.parse(contents || '{}');
  }

  return {};
}

function saveSubmission_(submission) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    if (!submission.id) submission.id = makeId_('work');
    if (!submission.createdAt) submission.createdAt = new Date().toISOString();
    submission.sentAt = submission.sentAt || new Date().toISOString();
    submission.status = 'sent';

    const sheet = ensureSheet_();
    const h = headerMap_();
    const rowIndex = sheet.getLastRow() + 1;

    let safeSubmission = cleanSubmission_(submission);
    let audioInfo = {
      fileName: submission.audio && submission.audio.fileName || '',
      mimeType: normalizeMimeType_(submission.audio && submission.audio.mimeType || ''),
      size: submission.audio && submission.audio.size || '',
      fileId: '',
      fileUrl: '',
      directUrl: ''
    };

    // Сначала сохраняем строку. Так учитель видит работу даже при ошибке Drive.
    sheet.getRange(rowIndex, 1, 1, IUS_HEADERS.length).setValues([makeRow_(safeSubmission, audioInfo, '')]);
    SpreadsheetApp.flush();

    try {
      audioInfo = saveAudio_(submission);
      safeSubmission = cleanSubmission_(submission);
      safeSubmission.audio = Object.assign({}, safeSubmission.audio || {}, {
        fileId: audioInfo.fileId || '',
        fileUrl: audioInfo.fileUrl || '',
        directUrl: audioInfo.directUrl || ''
      });

      sheet.getRange(rowIndex, h.status + 1).setValue('sent');
      sheet.getRange(rowIndex, h.audioFileName + 1).setValue(audioInfo.fileName || '');
      sheet.getRange(rowIndex, h.audioMimeType + 1).setValue(audioInfo.mimeType || '');
      sheet.getRange(rowIndex, h.audioSize + 1).setValue(audioInfo.size || '');
      sheet.getRange(rowIndex, h.audioFileId + 1).setValue(audioInfo.fileId || '');
      sheet.getRange(rowIndex, h.audioFileUrl + 1).setValue(audioInfo.fileUrl || '');
      sheet.getRange(rowIndex, h.audioDirectUrl + 1).setValue(audioInfo.directUrl || '');
      sheet.getRange(rowIndex, h.rawJson + 1).setValue(JSON.stringify(safeSubmission));
      sheet.getRange(rowIndex, h.audioError + 1).setValue('');

      return { id: submission.id, audioError: '', audioFileUrl: audioInfo.fileUrl || '' };
    } catch (err) {
      const audioError = getErrorMessage_(err);
      safeSubmission.audioError = audioError;

      sheet.getRange(rowIndex, h.status + 1).setValue('audio_error');
      sheet.getRange(rowIndex, h.rawJson + 1).setValue(JSON.stringify(safeSubmission));
      sheet.getRange(rowIndex, h.audioError + 1).setValue(audioError);

      return { id: submission.id, audioError };
    }
  } finally {
    lock.releaseLock();
  }
}

function cleanSubmission_(submission) {
  const safe = JSON.parse(JSON.stringify(submission || {}));
  if (safe.audio && safe.audio.dataUrl) delete safe.audio.dataUrl;
  return safe;
}

function makeRow_(submission, audioInfo, audioError) {
  return [
    submission.id || '',
    submission.createdAt || '',
    submission.sentAt || '',
    audioError ? 'audio_error' : (submission.status || 'sent'),
    submission.student && submission.student.fullName || '',
    submission.student && submission.student.className || '',
    submission.student && submission.student.login || '',
    submission.task && submission.task.id || '',
    submission.task && submission.task.title || '',
    submission.note || '',
    audioInfo.fileName || '',
    audioInfo.mimeType || '',
    audioInfo.size || '',
    audioInfo.fileId || '',
    audioInfo.fileUrl || '',
    audioInfo.directUrl || '',
    '', '', '', '', '', '',
    JSON.stringify(submission),
    audioError || ''
  ];
}

function saveAudio_(submission) {
  const audio = submission.audio || {};
  const fileName = sanitizeFileName_(audio.fileName || ('answer-' + (submission.id || Date.now()) + '.webm'));
  const fallbackMimeType = normalizeMimeType_(audio.mimeType || 'audio/webm');
  const size = audio.size || '';

  if (!audio.dataUrl) {
    return { fileName, mimeType: fallbackMimeType, size, fileId: '', fileUrl: '', directUrl: '' };
  }

  const parsed = parseAudioDataUrl_(audio.dataUrl, fallbackMimeType);
  const bytes = Utilities.base64Decode(parsed.base64);
  const blob = Utilities.newBlob(bytes, parsed.mimeType || fallbackMimeType, fileName);
  const folder = getOrCreateAudioFolder_();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    fileName,
    mimeType: parsed.mimeType || fallbackMimeType,
    size,
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    directUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId()
  };
}

function parseAudioDataUrl_(dataUrl, fallbackMimeType) {
  let value = String(dataUrl || '').trim();

  try {
    if (/^data%3A/i.test(value) || value.indexOf('%2C') !== -1 || value.indexOf('%3B') !== -1) {
      value = decodeURIComponent(value);
    }
  } catch (e) {}

  value = value.replace(/[\r\n\t]/g, '').replace(/ /g, '+');

  const commaIndex = value.indexOf(',');
  if (!/^data:/i.test(value) || commaIndex < 0) {
    throw new Error('Неверный формат аудио dataUrl: нет data:...base64,');
  }

  const meta = value.slice(5, commaIndex);
  const base64 = value.slice(commaIndex + 1);
  const parts = meta.split(';').filter(Boolean);
  const mimeType = normalizeMimeType_(parts[0] || fallbackMimeType || 'audio/webm');
  const hasBase64 = parts.some(part => String(part).toLowerCase() === 'base64');

  if (!hasBase64) throw new Error('Неверный формат аудио dataUrl: нет ;base64');
  if (!base64 || base64.length < 20) throw new Error('Неверный формат аудио dataUrl: пустая base64-строка');

  return { mimeType, base64 };
}

function normalizeMimeType_(mimeType) {
  return String(mimeType || 'audio/webm').split(';')[0].trim() || 'audio/webm';
}

function sanitizeFileName_(name) {
  return String(name || 'audio.webm').replace(/[\\/:*?"<>|]+/g, '_').slice(0, 120);
}

function listSubmissions_(limit) {
  const rows = getDataRows_();
  return rows
    .map(rowToItem_)
    .filter(item => item && item.id)
    .reverse()
    .slice(0, Math.max(1, Math.min(limit || 30, 100)));
}

function getSubmission_(id) {
  if (!id) throw new Error('Не указан id работы');
  const rows = getDataRows_();
  for (let i = 0; i < rows.length; i++) {
    const item = rowToItem_(rows[i]);
    if (item && String(item.id) === String(id)) return item;
  }
  throw new Error('Работа не найдена в таблице');
}

function getAudioData_(submissionId) {
  if (!submissionId) throw new Error('Не указан id работы для аудио');

  const item = getSubmission_(submissionId);
  const fileId = item.audio && item.audio.fileId;
  if (!fileId) {
    throw new Error('У этой работы нет audioFileId. Отправьте работу заново или проверьте ошибку аудио в таблице.');
  }

  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const bytes = blob.getBytes();

  // Ограничение нужно, чтобы страница учителя не зависала на слишком больших записях.
  // Для обычной короткой записи webm этого хватает с запасом.
  const maxBytes = 18 * 1024 * 1024;
  if (bytes.length > maxBytes) {
    throw new Error('Аудиофайл слишком большой для встроенного прослушивания. Откройте аудио по ссылке Google Drive.');
  }

  const mimeType = normalizeMimeType_(item.audio.mimeType || blob.getContentType() || 'audio/webm');
  const base64 = Utilities.base64Encode(bytes);

  return {
    ok: true,
    id: item.id,
    fileId: fileId,
    fileName: item.audio.fileName || file.getName(),
    mimeType: mimeType,
    size: bytes.length,
    fileUrl: item.audio.fileUrl || file.getUrl(),
    dataUrl: 'data:' + mimeType + ';base64,' + base64
  };
}

function saveResult_(submissionId, result) {
  if (!submissionId) throw new Error('Не указан id работы');
  const sheet = ensureSheet_();
  const values = sheet.getDataRange().getValues();
  const h = headerMap_();

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][h.id]) === String(submissionId)) {
      sheet.getRange(r + 1, h.status + 1).setValue('checked');
      sheet.getRange(r + 1, h.checkedAt + 1).setValue(result.checkedAt || new Date().toISOString());
      sheet.getRange(r + 1, h.teacherFullName + 1).setValue(result.teacher && result.teacher.fullName || '');
      sheet.getRange(r + 1, h.total + 1).setValue(result.total || 0);
      sheet.getRange(r + 1, h.verdict + 1).setValue(result.verdict || '');
      sheet.getRange(r + 1, h.teacherComment + 1).setValue(result.teacherComment || '');
      sheet.getRange(r + 1, h.selectedJson + 1).setValue(JSON.stringify(result.selected || []));
      return;
    }
  }

  throw new Error('Работа для результата не найдена');
}

function rowToItem_(row) {
  const h = headerMap_();
  const id = row[h.id] || '';
  if (!id) return null;

  let raw = {};
  try { raw = JSON.parse(row[h.rawJson] || '{}'); } catch (e) { raw = {}; }

  const audioError = row[h.audioError] || raw.audioError || '';

  return {
    app: 'saha-exam-v2',
    type: 'student-submission',
    version: raw.version || 10,
    id,
    createdAt: row[h.createdAt] || raw.createdAt || '',
    sentAt: row[h.sentAt] || raw.sentAt || '',
    status: row[h.status] || raw.status || 'sent',
    student: {
      id: raw.student && raw.student.id || '',
      fullName: row[h.studentFullName] || '',
      className: row[h.studentClassName] || '',
      login: row[h.studentLogin] || ''
    },
    task: {
      id: row[h.taskId] || '',
      title: row[h.taskTitle] || '',
      instruction: raw.task && raw.task.instruction || ''
    },
    note: row[h.note] || '',
    audio: {
      fileName: row[h.audioFileName] || '',
      mimeType: row[h.audioMimeType] || '',
      size: row[h.audioSize] || '',
      fileId: row[h.audioFileId] || '',
      fileUrl: row[h.audioFileUrl] || '',
      directUrl: row[h.audioDirectUrl] || ''
    },
    audioError,
    checkedAt: row[h.checkedAt] || '',
    result: {
      total: row[h.total] || '',
      verdict: row[h.verdict] || '',
      teacherComment: row[h.teacherComment] || ''
    }
  };
}

function getDataRows_() {
  const sheet = ensureSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  return values.slice(1).filter(row => row.some(cell => cell !== ''));
}

function ensureSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Откройте Apps Script из Google Таблицы: Расширения → Apps Script');

  let sheet = ss.getSheetByName(IUS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(IUS_SHEET_NAME);

  if (sheet.getMaxColumns() < IUS_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), IUS_HEADERS.length - sheet.getMaxColumns());
  }

  const firstRow = sheet.getRange(1, 1, 1, IUS_HEADERS.length).getValues()[0];
  const mismatch = IUS_HEADERS.some((header, index) => firstRow[index] !== header);
  if (mismatch) {
    sheet.getRange(1, 1, 1, IUS_HEADERS.length).setValues([IUS_HEADERS]);
  }
  sheet.setFrozenRows(1);
  return sheet;
}

function headerMap_() {
  const map = {};
  IUS_HEADERS.forEach((name, index) => map[name] = index);
  return map;
}

function getOrCreateAudioFolder_() {
  const folders = DriveApp.getFoldersByName(IUS_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(IUS_FOLDER_NAME);
}

function makeId_(prefix) {
  return prefix + '-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
}

function getErrorMessage_(err) {
  return String(err && err.message ? err.message : err);
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function jsonpOrJson_(data, callback) {
  if (callback) {
    return ContentService
      .createTextOutput(String(callback) + '(' + JSON.stringify(data) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(data);
}
