// ============================================================
// 360° Evaluation System — Google Apps Script Web App
// v2.0 — Roles: self/manager/peer | Questions: hardcoded
// Deploy: Execute as Me | Access: Anyone
// Run initializeSheets() once after setup.
// ============================================================

const CFG = {
  ROLES: ['self', 'manager', 'peer'],
  MAX_RATING: 10,
  DEFAULT_PIN: '1234',
  // Ваги ролей для підсумкового avg_score (сума = 1.0)
  // На ВП керівник — головний decision-maker, тому 50%.
  ROLE_WEIGHTS: { manager: 0.5, peer: 0.4, self: 0.1 },
  // ID питання-вердикту керівника (див. QUESTION_SETS.manager)
  MANAGER_VERDICT_QID: 'm8',
  VERDICT_CONTINUE:   'Продовжити',
  VERDICT_STOP:       'Не продовжувати',
  VERDICT_EXTEND:     'Потребує додаткового місяця'
};

// ============================================================
// HARDCODED QUESTION SETS PER ROLE
// type: 'rating'  — числова шкала 1-10, враховується у score
// type: 'choice'  — вибір варіанту, НЕ враховується у score
// type: 'open'    — вільний текст, НЕ враховується у score
// ============================================================

const QUESTION_SETS = {
  manager: [
    { id: 'm1', text: 'Оцінка виконання робочих завдань за випробувальний термін', type: 'rating', weight: 5 },
    { id: 'm2', text: 'Відповідність цінностям компанії: Екологічність, Чесність, Гнучкість', type: 'rating', weight: 2 },
    { id: 'm3', text: 'Проактивність у роботі', type: 'rating', weight: 2 },
    { id: 'm4', text: 'Комунікативні навички: чіткість у висловленні, вміння давати та приймати зворотний зв\'язок', type: 'rating', weight: 3 },
    { id: 'm5', text: 'Управління часом та дедлайнами', type: 'rating', weight: 3 },
    { id: 'm6', text: 'Рівень самостійності у роботі', type: 'rating', weight: 2 },
    { id: 'm7', text: 'Потенціал до розвитку в команді', type: 'rating', weight: 2 },
    { id: 'm8', text: 'Чи рекомендуєте ви продовжити роботу з цим співробітником після випробувального терміну?',
      type: 'choice', options: ['Так — ВП пройдено успішно', 'Ні — ВП не пройдено', 'Потребує додаткового місяця для фінального рішення'], weight: 0 },
    { id: 'm9', text: 'Відкритий коментар (необов\'язково)', type: 'open', weight: 0 }
  ],
  peer: [
    { id: 'p1', text: 'Наскільки комфортно та ефективно вам було взаємодіяти по робочих питаннях?', type: 'rating', weight: 2 },
    { id: 'p2', text: 'Чи відповідав співробітник цінностям Екологічності та Чесності у спілкуванні з вами?', type: 'rating', weight: 2 },
    { id: 'p3', text: 'Чи виконував співробітник взяті зобов\'язання перед вами або командою?', type: 'rating', weight: 2 },
    { id: 'p4', text: 'Наскільки цей співробітник вписується в атмосферу команди KEVIN?', type: 'rating', weight: 1 },
    { id: 'p5', text: 'Якби ти міг дати одну щиру пораду цьому співробітнику — що б це було?', type: 'open', weight: 0 }
  ],
  self: [
    { id: 's1', text: 'Наскільки якісно, на вашу думку, ви впорались з робочим функціоналом за цей час?', type: 'rating', weight: 2 },
    { id: 's2', text: 'Наскільки вам вдалось відповідати цінностям компанії: Чесність, Екологічність, Проактивність?', type: 'rating', weight: 1 },
    { id: 's3', text: 'Наскільки комфортно вам було взаємодіяти з командою?', type: 'rating', weight: 1 },
    { id: 's4', text: 'Що, на вашу думку, вам вдалось найкраще за цей період?', type: 'open', weight: 0 },
    { id: 's5', text: 'Що хотіли б покращити або що було найскладнішим?', type: 'open', weight: 0 }
  ]
};

// ============================================================
// HTTP ENTRY POINTS
// ============================================================

function doGet(e) {
  try {
    const action = (e.parameter || {}).action;
    switch (action) {
      case 'getConfigs':     return ok(getConfigs());
      case 'getSubmissions': return ok(getSubmissions());
      case 'getArchive':     return ok(getArchive());
      case 'getSettings':    return ok(getSettings());
      case 'getQuestions':          return ok(QUESTION_SETS);
      case 'getSubmissionDetails':  return ok(getSubmissionDetails(e.parameter.id));
      default:               return ok({ error: 'Unknown GET action: ' + action });
    }
  } catch (err) {
    return ok({ error: err.message });
  }
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (_) {
    return ok({ error: 'Invalid JSON in request body' });
  }
  try {
    const action = body.action;
    switch (action) {
      case 'submitForm':    return ok(submitForm(body));
      case 'updateConfig':  return ok(updateConfig(body));
      case 'deleteConfig':  return ok(deleteConfig(body));
      case 'archiveUser':   return ok(archiveUser(body));
      case 'checkPin':      return ok(checkPin(body));
      case 'setPin':        return ok(setPin(body));
      case 'setBackendUrl': return ok(setBackendUrl(body));
      default:              return ok({ error: 'Unknown POST action: ' + action });
    }
  } catch (err) {
    return ok({ error: err.message });
  }
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SHEET HELPERS
// ============================================================

function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  const s = ss().getSheetByName(name);
  if (!s) throw new Error('Аркуш "' + name + '" не знайдено. Запустіть initializeSheets().');
  return s;
}

function sheetRows(name) {
  const s = getSheet(name);
  if (s.getLastRow() < 1) return [];
  return s.getDataRange().getValues();
}

// ============================================================
// SETTINGS
// ============================================================

function getSettings() {
  const rows = sheetRows('Settings');
  const out = {};
  for (let i = 1; i < rows.length; i++) {
    const k = String(rows[i][0] || '').trim();
    if (k) out[k] = String(rows[i][1] || '');
  }
  return out;
}

function setSetting(key, value) {
  const s = getSheet('Settings');
  const rows = sheetRows('Settings');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) {
      s.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  s.appendRow([key, value]);
}

function checkPin(body) {
  const settings = getSettings();
  const correct = settings['admin_pin'] || CFG.DEFAULT_PIN;
  return { valid: String(body.pin || '') === correct };
}

function setPin(body) {
  const p = String(body.pin || '').trim();
  if (p.length < 4) return { error: 'PIN має містити мінімум 4 символи' };
  setSetting('admin_pin', p);
  return { success: true };
}

function setBackendUrl(body) {
  const url = String(body.url || '').trim();
  if (!url) return { error: 'URL обов\'язковий' };
  setSetting('backend_url', url);
  return { success: true };
}

// ============================================================
// CONFIGS
// Cols: id(0) | name(1) | role(2) | questions_json(3) |
//       weights_json(4) | total(5) | peer_count(6)
// questions_json / weights_json зберігаються порожніми —
// питання hardcoded у QUESTION_SETS
// ============================================================

function getConfigs() {
  const rows = sheetRows('Configs');
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] && !r[1]) continue;
    const peerCount    = Math.max(1, Number(r[6]) || 2);
    const totalRequired = 2 + peerCount; // self + manager + N peers
    out.push({
      id:             String(r[0]),
      name:           String(r[1] || ''),
      role:           String(r[2] || ''),
      peer_count:     peerCount,
      total_required: totalRequired
    });
  }
  return out;
}

function updateConfig(body) {
  const name = String(body.name || '').trim();
  if (!name) return { error: 'Ім\'я обов\'язкове' };

  const role      = String(body.role || '').trim();
  const peerCount = Math.max(1, parseInt(body.peer_count) || 2);

  const s    = getSheet('Configs');
  const rows = sheetRows('Configs');

  if (body.id) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(body.id)) {
        // Cols 1-7: id, name, role, questions_json(empty), weights_json(empty), total(0), peer_count
        s.getRange(i + 1, 1, 1, 7).setValues([[body.id, name, role, '[]', '[]', 0, peerCount]]);
        return { success: true, id: body.id };
      }
    }
    return { error: 'Конфігурацію не знайдено' };
  }

  const newId = Date.now().toString();
  s.appendRow([newId, name, role, '[]', '[]', 0, peerCount]);
  return { success: true, id: newId };
}

function deleteConfig(body) {
  if (!body.id) return { error: 'id обов\'язковий' };
  const s    = getSheet('Configs');
  const rows = sheetRows('Configs');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(body.id)) {
      s.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Конфігурацію не знайдено' };
}

// ============================================================
// SUBMISSIONS
// ============================================================

function getSubmissions() {
  const cfgRows = sheetRows('Configs');
  const subRows = sheetRows('Submissions');
  const out = [];

  for (let i = 1; i < cfgRows.length; i++) {
    const cfg = cfgRows[i];
    if (!cfg[0] && !cfg[1]) continue;

    const id        = String(cfg[0]);
    const peerCount = Math.max(1, Number(cfg[6]) || 2);
    const totalRequired = 2 + peerCount;

    const subs = subRows.slice(1).filter(r => String(r[1]) === id);

    const selfSubs    = subs.filter(r => String(r[3]) === 'self');
    const managerSubs = subs.filter(r => String(r[3]) === 'manager');
    const peerSubs    = subs.filter(r => String(r[3]) === 'peer');

    const isComplete = selfSubs.length >= 1 && managerSubs.length >= 1 && peerSubs.length >= peerCount;

    // Витягуємо вердикт керівника (m8) — робимо ЗАВЖДИ, якщо керівник заповнив,
    // навіть до повного завершення. Це дозволяє показати рішення одразу.
    let managerVerdict = null;
    if (managerSubs.length >= 1) {
      try {
        const answers = JSON.parse(managerSubs[0][4] || '[]');
        const verdictAns = answers.find(a => String(a.qid) === CFG.MANAGER_VERDICT_QID);
        if (verdictAns && verdictAns.value) managerVerdict = String(verdictAns.value);
      } catch (_) {}
    }

    let avgScore   = null;
    let roleScores = {};
    let scoreGap   = null;
    if (isComplete) {
      const avg = arr => arr.length
        ? Math.round(arr.reduce((s, r) => s + (Number(r[5]) || 0), 0) / arr.length)
        : null;
      roleScores.self    = avg(selfSubs);
      roleScores.manager = avg(managerSubs);
      roleScores.peer    = avg(peerSubs);

      // Зважений avg_score за роллю (не просте середнє всіх сабмішенів).
      const w = CFG.ROLE_WEIGHTS;
      avgScore = Math.round(
        roleScores.manager * w.manager +
        roleScores.peer    * w.peer +
        roleScores.self    * w.self
      );

      scoreGap = Math.abs(roleScores.manager - roleScores.peer);
    }

    // ВП-статус — hard override за вердиктом керівника (m8). Бал — довідково.
    let vpStatus = 'pending_verdict';
    if (managerVerdict === CFG.VERDICT_CONTINUE)      vpStatus = 'passed';
    else if (managerVerdict === CFG.VERDICT_STOP)     vpStatus = 'not_passed';
    else if (managerVerdict === CFG.VERDICT_EXTEND)   vpStatus = 'extended';

    out.push({
      evaluated_id:    id,
      name:            String(cfg[1] || ''),
      role:            String(cfg[2] || ''),
      peer_count:      peerCount,
      total_required:  totalRequired,
      submitted_count: subs.length,
      progress: {
        self:    selfSubs.length,
        manager: managerSubs.length,
        peer:    peerSubs.length
      },
      status:           isComplete ? 'complete' : 'incomplete',
      avg_score:        avgScore,
      role_scores:      roleScores,
      role_weights:     CFG.ROLE_WEIGHTS,
      manager_verdict:  managerVerdict,
      vp_status:        vpStatus,
      score_gap:        scoreGap
    });
  }

  return out;
}

function submitForm(body) {
  const { evaluated_id, evaluator_name, role, answers } = body;

  if (!evaluated_id)                              return { error: 'evaluated_id обов\'язковий' };
  if (!String(evaluator_name || '').trim())       return { error: 'Введіть ваше ПІБ' };
  if (!CFG.ROLES.includes(role))                  return { error: 'Невідома роль: ' + role };
  if (!Array.isArray(answers) || !answers.length) return { error: 'Відповіді відсутні' };

  // Find config
  const cfgRows = sheetRows('Configs');
  let cfg = null;
  for (let i = 1; i < cfgRows.length; i++) {
    if (String(cfgRows[i][0]) === String(evaluated_id)) { cfg = cfgRows[i]; break; }
  }
  if (!cfg) return { error: 'Конфігурацію не знайдено' };

  const peerCount = Math.max(1, Number(cfg[6]) || 2);
  const subRows   = sheetRows('Submissions');
  const existing  = subRows.slice(1).filter(r => String(r[1]) === String(evaluated_id));

  // ── Duplicate / capacity checks ───────────────────────────
  if (role === 'self' || role === 'manager') {
    const hasRole = existing.some(r => String(r[3]) === role);
    if (hasRole) return { error: 'Ця роль вже пройшла оцінювання для цього співробітника' };
  } else if (role === 'peer') {
    const peerSubs = existing.filter(r => String(r[3]) === 'peer');
    if (peerSubs.length >= peerCount) {
      return { error: 'Вже зібрано максимальну кількість оцінок від колег (' + peerCount + ')' };
    }
    // Same evaluator cannot submit twice for the same person
    const nameLower = String(evaluator_name).trim().toLowerCase();
    const alreadyIn = peerSubs.some(r => String(r[2]).trim().toLowerCase() === nameLower);
    if (alreadyIn) return { error: 'Ви вже залишили оцінку для цього співробітника' };
  }

  // ── Score calculation (rating questions only) ─────────────
  const questions = QUESTION_SETS[role] || [];
  let score = 0;
  try {
    let weighted = 0, wSum = 0;
    answers.forEach(ans => {
      const q = questions.find(q => q.id === ans.qid);
      if (!q || q.type !== 'rating') return;
      const val = Math.min(Math.max(Number(ans.value) || 1, 1), CFG.MAX_RATING);
      weighted += val * q.weight;
      wSum     += q.weight;
    });
    score = wSum > 0 ? Math.round((weighted / (wSum * CFG.MAX_RATING)) * 100) : 0;
  } catch (_) {}

  // ── Completion check ──────────────────────────────────────
  const selfCount    = existing.filter(r => String(r[3]) === 'self').length    + (role === 'self'    ? 1 : 0);
  const managerCount = existing.filter(r => String(r[3]) === 'manager').length + (role === 'manager' ? 1 : 0);
  const peerActual   = existing.filter(r => String(r[3]) === 'peer').length    + (role === 'peer'    ? 1 : 0);
  const isComplete   = selfCount >= 1 && managerCount >= 1 && peerActual >= peerCount;

  getSheet('Submissions').appendRow([
    new Date().toISOString(),
    String(evaluated_id),
    String(evaluator_name).trim(),
    role,
    JSON.stringify(answers),   // answers = [{qid, value}, ...]
    score,
    isComplete ? 'complete' : 'incomplete'
  ]);

  return { success: true };
}

// ============================================================
// SUBMISSION DETAILS (for admin answer viewer)
// ============================================================

function getSubmissionDetails(evaluatedId) {
  if (!evaluatedId) return { error: 'id required' };

  const cfgRows = sheetRows('Configs');
  const subRows = sheetRows('Submissions');

  let cfg = null;
  for (let i = 1; i < cfgRows.length; i++) {
    if (String(cfgRows[i][0]) === String(evaluatedId)) { cfg = cfgRows[i]; break; }
  }

  const subs = subRows.slice(1).filter(r => String(r[1]) === String(evaluatedId));

  return {
    evaluated_id: evaluatedId,
    name:  cfg ? String(cfg[1] || '') : '',
    role:  cfg ? String(cfg[2] || '') : '',
    submissions: subs.map(r => {
      let answers = [];
      try { answers = JSON.parse(r[4] || '[]'); } catch (_) {}
      return {
        timestamp:      r[0] ? r[0].toString() : '',
        evaluator_name: String(r[2] || ''),
        role:           String(r[3] || ''),
        answers,
        score:          Number(r[5]) || 0
      };
    }),
    // Повертаємо питання щоб фронт міг декодувати qid → текст
    question_sets: QUESTION_SETS
  };
}

// ============================================================
// ARCHIVE
// ============================================================

function getArchive() {
  const rows = sheetRows('Archive');
  const out  = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[1]) continue;
    out.push({
      timestamp:      r[0] ? r[0].toString() : '',
      evaluated_id:   String(r[1]),
      evaluator_name: String(r[2] || ''),
      role:           String(r[3] || ''),
      answers_json:   String(r[4] || ''),
      score:          Number(r[5]) || 0
    });
  }
  return out;
}

function archiveUser(body) {
  if (!body.evaluated_id) return { error: 'evaluated_id обов\'язковий' };

  const subSheet  = getSheet('Submissions');
  const archSheet = getSheet('Archive');
  const rows      = sheetRows('Submissions');

  if (rows.length < 1) return { error: 'Аркуш Submissions порожній' };

  const header    = rows[0];
  const toArchive = [];
  const toKeep    = [];

  for (let i = 1; i < rows.length; i++) {
    (String(rows[i][1]) === String(body.evaluated_id) ? toArchive : toKeep).push(rows[i]);
  }

  if (toArchive.length === 0) return { error: 'Відповіді для цього користувача не знайдено' };

  toArchive.forEach(r => archSheet.appendRow(r));

  subSheet.clearContents();
  subSheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (toKeep.length > 0) {
    subSheet.getRange(2, 1, toKeep.length, header.length).setValues(toKeep);
  }

  return { success: true, archived: toArchive.length };
}

// ============================================================
// INITIALIZATION — Run once manually from Apps Script editor
// ============================================================

function initializeSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const defs = {
    // peer_count додано як 7-а колонка
    'Configs':     ['id', 'name', 'role', 'questions_json', 'weights_json', 'total', 'peer_count'],
    'Submissions': ['timestamp', 'evaluated_id', 'evaluator_name', 'role', 'answers_json', 'score', 'status'],
    'Archive':     ['timestamp', 'evaluated_id', 'evaluator_name', 'role', 'answers_json', 'score'],
    'Settings':    ['key', 'value']
  };

  Object.entries(defs).forEach(([name, headers]) => {
    let s = spreadsheet.getSheetByName(name);
    if (!s) s = spreadsheet.insertSheet(name);
    const hRange = s.getRange(1, 1, 1, headers.length);
    hRange.setValues([headers]);
    hRange.setFontWeight('bold');
    hRange.setBackground('#0d1117');
    hRange.setFontColor('#FFD300');
  });

  // Default settings
  const settSheet = spreadsheet.getSheetByName('Settings');
  const settData  = settSheet.getDataRange().getValues();
  const keys      = settData.map(r => String(r[0]));
  if (!keys.includes('admin_pin'))   settSheet.appendRow(['admin_pin', '1234']);
  if (!keys.includes('backend_url')) settSheet.appendRow(['backend_url', '']);

  SpreadsheetApp.getUi().alert(
    'Аркуші ініціалізовано! (v2.0)\n\n' +
    'Стандартний PIN: 1234\n' +
    'Ролі: self / manager / peer\n' +
    'Питання: hardcoded у QUESTION_SETS\n\n' +
    'Наступний крок: розгорніть скрипт як Web App\n' +
    'і вставте URL у вкладку Налаштування адмінки.'
  );
}
