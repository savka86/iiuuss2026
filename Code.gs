/**
 * ИУС 2.0 — бесплатная общая база работ
 * Google Таблица + Google Apps Script + Google Drive
 *
 * Важно:
 * 1) Код вставлять в Apps Script, открытый из Google Таблицы: Расширения → Apps Script.
 * 2) Запустить setupIUS один раз и разрешить доступ.
 * 3) Развернуть как веб-приложение:
 *    - Выполнять от имени: Меня
 *    - У кого есть доступ: Все
 * 4) После каждого изменения кода: Управление развертываниями → Изменить → Новая версия → Развернуть.
 */

const SHEET_NAME = 'works';
const FOLDER_NAME = 'ИУС 2.0 — аудиозаписи учеников';

const HEADERS = [
  'id', 'createdAt', 'sentAt', 'status',
  'studentFullName', 'studentClassName', 'studentLogin',
  'taskId', 'taskTitle', 'note',
  'audioFileName', 'audioMimeType', 'audioSize', 'audioFileId', 'audioFileUrl', 'audioDirectUrl',
  'checkedAt', 'teacherFullName', 'total', 'verdict', 'teacherComment', 'selectedJson',
  'rawJson', 'audioError'
];

function setupIUS() {
  ensureSheet_();
  getFolder_();
  return 'ИУС 2.0 готов. Теперь разверните Apps Script как веб-приложение.';
}

function doPost(e) {
  try {
    const payload = parsePayload_(e);

    if (payload.action === 'submit') {
      const result = saveSubmission_(payload.submission || {});
      return json_({ ok: true, id: result.id, audioError: result.audioError || '' });
    }

    if (payload.action === 'saveResult') {
      saveResult_(payload.submissionId, payload.result || {});
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'Unknown action: ' + (payload.action || '') });
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
        sheet: SHEET_NAME,
        count: getDataRows_().length
      };
    } else if (p.action === 'list') {
      data = { ok: true, items: listSubmissions_(Number(p.limit || 30)) };
    } else if (p.action === 'get') {
      data = { ok: true, item: getSubmission_(p.id) };
    } else {
      data = { ok: false, error: 'Unknown action: ' + (p.action || '') };
    }

    return jsonpOrJson_(data, p.callback);
  } catch (err) {
    return jsonpOrJson_({ ok: false, error: getErrorMessage_(err) }, p.callback);
  }
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

  if (e && e.parameter) {
    return e.parameter;
  }

  return {};
}

function saveSubmission_(submission) {
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
    mimeType: submission.audio && submission.audio.mimeType || '',
    size: submission.audio && submission.audio.size || '',
    fileId: '',
    fileUrl: '',
    directUrl: ''
  };

  // Сначала записываем строку в таблицу.
  // Так работа появится у учителя даже если Drive не успеет/не сможет сохранить аудио.
  sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([makeRow_(safeSubmission, audioInfo, '')]);
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
    return { id: submission.id, audioError: '' };
  } catch (err) {
    const audioError = getErrorMessage_(err);
    safeSubmission.audioError = audioError;

    sheet.getRange(rowIndex, h.status + 1).setValue('audio_error');
    sheet.getRange(rowIndex, h.rawJson + 1).setValue(JSON.stringify(safeSubmission));
    sheet.getRange(rowIndex, h.audioError + 1).setValue(audioError);
    return { id: submission.id, audioError };
  }
}

function cleanSubmission_(submission) {
  const safe = JSON.parse(JSON.stringify(submission || {}));
  if (safe.audio && safe.audio.dataUrl) {
    delete safe.audio.dataUrl;
  }
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
  const fileName = audio.fileName || ('answer-' + (submission.id || Date.now()) + '.webm');
  const mimeType = audio.mimeType || 'audio/webm';
  const size = audio.size || '';

  if (!audio.dataUrl) {
    return { fileName, mimeType, size, fileId: '', fileUrl: '', directUrl: '' };
  }

  const match = String(audio.dataUrl).match(/^data:([^;]+);base64,(.+)$/);
  if (!match) throw new Error('Неверный формат аудио dataUrl');

  const bytes = Utilities.base64Decode(match[2]);
  const blob = Utilities.newBlob(bytes, match[1] || mimeType, fileName);
  const folder = getFolder_();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    fileName,
    mimeType: match[1] || mimeType,
    size,
    fileId: file.getId(),
    fileUrl: file.getUrl(),
    directUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId()
  };
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
    version: raw.version || 4,
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

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getMaxColumns() < HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.length - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  return sheet;
}

function headerMap_() {
  const map = {};
  HEADERS.forEach((name, index) => map[name] = index);
  return map;
}

function getFolder_() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOLDER_NAME);
}

function makeId_(prefix) {
  return prefix + '-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
}

function getErrorMessage_(err) {
  return String(err && err.message ? err.message : err);
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpOrJson_(data, callback) {
  if (callback) {
    return ContentService
      .createTextOutput(String(callback) + '(' + JSON.stringify(data) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(data);
}

/**
 * Быстрый тест прямо из Apps Script.
 * После запуска в таблице должна появиться тестовая строка.
 */
function testAddRow() {
  return saveSubmission_({
    app: 'saha-exam-v2',
    type: 'student-submission',
    version: 4,
    id: makeId_('test'),
    createdAt: new Date().toISOString(),
    sentAt: new Date().toISOString(),
    status: 'sent',
    student: {
      fullName: 'Тестовый ученик',
      className: '9А',
      login: 'test'
    },
    task: {
      id: 1,
      title: 'Проверка подключения',
      instruction: ''
    },
    note: 'Тестовая строка из Apps Script',
    audio: {
      fileName: 'no-audio.webm',
      mimeType: 'audio/webm',
      size: 0
    }
  });
}
