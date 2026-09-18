/* VZHDO v7 dashboard — see prompt */
(function () {
  const BASE = new URL('./', document.baseURI).pathname;
  const LOGO = BASE + 'assets/vzhdo-logo.png';
  const STRUCT = BASE + 'structure.xlsx?v=' + Date.now();
  const DB = 'vzhdo-v7';
  const $ = (s, e = document) => e.querySelector(s);
  const h = (tag, attrs = {}, kids = []) => {
    const n = document.createElement(tag);
    Object.entries(attrs).forEach(([k, v]) => {
      if (k === 'class') n.className = v;
      else if (k.startsWith('on') && typeof v === 'function') n.addEventListener(k.slice(2).toLowerCase(), v);
      else if (v === true) n.setAttribute(k, '');
      else if (v != null && v !== false) n.setAttribute(k, v);
    });
    (Array.isArray(kids) ? kids : [kids]).forEach((c) => {
      if (c == null) return;
      n.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    });
    return n;
  };
  const startOfDay = (d) => { const x = new Date(d); x.setHours(0,0,0,0); return x; };
  function parseDate(raw) {
    if (raw == null || raw === '') return null;
    if (raw instanceof Date && !isNaN(raw)) return startOfDay(raw);
    if (typeof raw === 'number' && raw > 20000 && raw < 80000) {
      const utc = new Date(Math.round((raw - 25569) * 86400 * 1000));
      return startOfDay(new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate()));
    }
    const s = String(raw).trim();
    let m = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})/);
    if (m) return startOfDay(new Date(+m[3], +m[2] - 1, +m[1]));
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return startOfDay(new Date(+m[1], +m[2] - 1, +m[3]));
    const d = new Date(s); return isNaN(d) ? null : startOfDay(d);
  }
  const daysBetween = (a, b) => Math.round((startOfDay(a) - startOfDay(b)) / 86400000);
  const fmt = (d) => d ? String(d.getDate()).padStart(2,'0') + '.' + String(d.getMonth()+1).padStart(2,'0') + '.' + d.getFullYear() : '—';
  const normSpace = (s) => String(s || '').replace(/\s+/g, ' ').trim();
  const normHeader = (s) => normSpace(s).toLowerCase();
  const normalizeName = (s) => normSpace(s).replace(/ё/gi,'е').replace(/[.]/g,' ').replace(/\s+/g,' ').trim().toLowerCase();
  function nameParts(s) {
    const p = normalizeName(s).split(' ').filter(Boolean);
    return { normalized:p.join(' '), surname:p[0] || '', initials:p.slice(1).map(x => x[0]).join('') };
  }
  function namesMatch(a, b) {
    return normalizeName(a) && normalizeName(a) === normalizeName(b);
  }
  function normStatus(s) {
    const map = {'снят с контроля':'Снят с контроля','рабочий контроль':'Рабочий контроль','в работе':'В работе','просрочен':'Просрочено','просрочено':'Просрочено','на снятии с контроля':'На снятии с контроля','на исполнении':'На исполнении','завершено':'Завершено','выполнен':'Выполнен','выполнено':'Выполнено','не исполнено':'Не исполнено','не исполнен':'Не исполнен'};
    return map[normSpace(s).toLowerCase()] || normSpace(s);
  }
  function classifyDeadline(raw) {
    if (raw == null || String(raw).trim() === '') return { date:null, type:'MISSING', text:null };
    const text = String(raw).trim(), parsed = parseDate(raw);
    return parsed ? { date:parsed, type:'DATE', text } : { date:null, type:'TEXT', text };
  }
  function deadlineDisplay(type, date, text, delta, status, module) {
    if (type === 'TEXT') return text || '—';
    if (!date) return 'Срок не указан';
    const dateStr = fmt(date), completed = /снят с контроля|завершено|выполнен/i.test(status);
    if (delta == null) return dateStr;
    if (completed) return delta < 0 ? dateStr + ' · Срок был ' + Math.abs(delta) + ' дн. назад' : dateStr;
    if (delta < 0) return /просрочено/i.test(status) && module === 'protocol' ? dateStr + ' · Просрочено ' + Math.abs(delta) + ' дн.' : dateStr + ' · Срок истёк ' + Math.abs(delta) + ' дн. назад';
    if (delta === 0) return dateStr + ' · Срок сегодня';
    return dateStr + ' · осталось ' + delta + ' дн.';
  }
  function findHeader(matrix, required) {
    const req = required.map(normHeader);
    for (let i = 0; i < Math.min(matrix.length, 30); i++) {
      const colMap = {};
      (matrix[i] || []).forEach((c, idx) => { const hh = normHeader(String(c ?? '')); if (hh) colMap[hh] = idx; });
      if (req.every(r => Object.keys(colMap).some(k => k === r))) return { rowIndex:i, colMap };
    }
    return null;
  }
  function detectSheet(wb, required) {
    for (const name of wb.SheetNames) {
      const matrix = XLSX.utils.sheet_to_json(wb.Sheets[name], { header:1, raw:true, defval:'' });
      const header = findHeader(matrix, required);
      if (header) return { sheetName:name, matrix, header };
    }
    return null;
  }
  const cell = (row, colMap, header) => { const i = colMap[normHeader(header)]; return i == null ? '' : row[i]; };
  const cellStr = (row, colMap, header) => normSpace(String(cell(row, colMap, header) ?? ''));
  const emptyRow = (row) => !row || row.every(c => c == null || String(c).trim() === '');

  function emptyF() {
    return { q:'', units:[], assigned:[], curators:[], heads:[], spLeaders:[], norDirectors:[], joint:[], statuses:[], types:[], positions:[], number:'', author:'', summary:'', period:'all', dateFrom:'', dateTo:'', dateMonth:'', dateYear:'', dlPreset:'', dlFrom:'', dlTo:'', dlMonth:'', dlYear:'', respFrom:'', respTo:'', respMonth:'', respYear:'', respMissing:false, showAll:false };
  }
  const F = { p: emptyF(), a: emptyF(), i: emptyF() };
  const SORT = { p: {key:'', dir:1}, a: {key:'', dir:1}, i: {key:'', dir:1} };
  let openMs = '';
  const S = { page:'overview', structure:null, structureError:null, notice:'', protocols:[], appeals:[], incoming:[], protoMeta:null, appealMeta:null, inMeta:null, protoKpi:'', appealKpi:'', inKpi:'', tabA:'all', tabI:'all', detail:null, err:'' };
  const refDate = () => startOfDay(new Date());

  function parseStructure(wb) {
    const found = detectSheet(wb, ['Ф.И.О.', 'Структурное подразделение', 'Должность']);
    if (!found) throw new Error('не найдены заголовки structure.xlsx');
    const keys = Object.keys(found.header.colMap);
    const need = [
      ['руководитель-куратор', 'руководитель- куратор', 'руководитель - куратор'],
      ['совместное исполнение'],
      ['руководители сп'],
      ['директора нор']
    ];
    const miss = [];
    if (!need[0].some(a => keys.some(k => k.includes(a) || k === a))) miss.push('Руководитель-куратор');
    if (!need[1].some(a => keys.some(k => k.includes(a)))) miss.push('Совместное исполнение');
    if (!need[2].some(a => keys.some(k => k.includes(a)))) miss.push('Руководители СП');
    if (!need[3].some(a => keys.some(k => k.includes(a)))) miss.push('Директора НОР');
    if (miss.length) throw new Error('structure.xlsx: нет обязательных столбцов: ' + miss.join(', '));
    const { matrix, header } = found, people = [];
    const get = (row, aliases) => { for (const [k, idx] of Object.entries(header.colMap)) { if (aliases.some(a => k.includes(a))) return normSpace(String(row[idx] ?? '')); } return ''; };
    for (let i = header.rowIndex + 1; i < matrix.length; i++) {
      const row = matrix[i]; if (emptyRow(row)) continue;
      const fullName = cellStr(row, header.colMap, 'Ф.И.О.'), structuralUnit = cellStr(row, header.colMap, 'Структурное подразделение');
      if (!fullName && !structuralUnit) continue;
      people.push({ fullName, structuralUnit, position: cellStr(row, header.colMap, 'Должность'), curator: get(row, ['руководитель- куратор','руководитель-куратор']), jointExecution: get(row, ['совместное исполнение']), spLeaderMarker: get(row, ['руководители сп']), norDirectorMarker: get(row, ['директора нор']) });
    }
    const spLeaderByUnit = {}, norDirectorByUnit = {};
    people.forEach(p => {
      if (p.spLeaderMarker && p.fullName) (spLeaderByUnit[p.spLeaderMarker] ||= []).push(p.fullName);
      if (p.norDirectorMarker && p.fullName) (norDirectorByUnit[p.norDirectorMarker] ||= []).push(p.fullName);
    });
    return { people, units: [...new Set(people.map(p => p.structuralUnit).filter(Boolean))].sort(), curators: [...new Set(people.map(p => p.curator).filter(Boolean))].sort(), spLeaderByUnit, norDirectorByUnit, loadedAt: new Date().toISOString() };
  }
  function nameInitials(s) { return nameParts(s).initials; }
  function matchPerson(index, fio) {
    const result = { hits:[], state:'NOT_FOUND' };
    if (!index || !fio) return result;
    const src = nameParts(fio);
    const exact = index.people.filter(p => p.fullName && namesMatch(p.fullName, fio));
    if (exact.length === 1) return { hits:exact, state:'MATCHED' };
    if (exact.length > 1) return { hits:[], state:'AMBIGUOUS' };
    if (!src.surname) return result;
    if (src.initials) {
      const byInitials = index.people.filter(p => { const q=nameParts(p.fullName); return q.surname===src.surname && q.initials===src.initials; });
      if (byInitials.length === 1) return { hits:byInitials, state:'MATCHED' };
      if (byInitials.length > 1) return { hits:[], state:'AMBIGUOUS' };
    }
    const bySurname = index.people.filter(p => nameParts(p.fullName).surname === src.surname);
    if (bySurname.length === 1) {
      const dst = nameParts(bySurname[0].fullName);
      if (src.initials && dst.initials && src.initials !== dst.initials) return result;
      return { hits:bySurname, state:'MATCHED' };
    }
    if (bySurname.length > 1) return { hits:[], state:'AMBIGUOUS' };
    return result;
  }
  function enrichUnits(index, units) {
    const resolvedUnits = [...new Set(units.filter(Boolean))];
    const org = { resolvedUnits, curators:[], jointGroups:[], spLeaders:[], norDirectors:[], departmentHeads:[], positions:[], isJointExecution: resolvedUnits.length > 1, warnings:[] };
    if (!index) { org.warnings.push('structure.xlsx не загружена'); return org; }
    resolvedUnits.forEach(u => {
      index.people.filter(p => p.structuralUnit === u).forEach(p => { if (p.curator) org.curators.push(p.curator); if (p.jointExecution) org.jointGroups.push(p.jointExecution); });
      (index.spLeaderByUnit[u] || []).forEach(n => org.spLeaders.push(n));
      (index.norDirectorByUnit[u] || []).forEach(n => org.norDirectors.push(n));
    });
    org.curators = [...new Set(org.curators)]; org.jointGroups = [...new Set(org.jointGroups)]; org.spLeaders = [...new Set(org.spLeaders)]; org.norDirectors = [...new Set(org.norDirectors)];
    org.departmentHeads = [...new Set(org.spLeaders.concat(org.norDirectors))];
    org.isJointExecution = org.isJointExecution || org.jointGroups.length > 0;
    return org;
  }
  function enrichByPerson(index, fio) {
    const match = matchPerson(index, fio);
    const hits = match.hits;
    const org = enrichUnits(index, hits.map(h => h.structuralUnit).filter(Boolean));
    org.positions = [...new Set(hits.map(h => h.position).filter(Boolean))];
    org.matchedPersons = hits.map(h => h.fullName);
    org.matchState = match.state;
    if (fio && match.state === 'NOT_FOUND') org.warnings.push('Ф.И.О. не найдено в structure.xlsx');
    if (fio && match.state === 'AMBIGUOUS') org.warnings.push('Неоднозначное сопоставление Ф.И.О.; структура не назначена');
    return org;
  }
  function knownUnitList(index) {
    const units = index ? [...index.units] : [];
    units.sort((a,b) => b.length - a.length);
    return units;
  }
  const UNIT_ALIAS = {
    'координатор по внедрению тех.решений': 'Координатор по внедрению технических решений',
    'координатор по внедрению тех решений': 'Координатор по внедрению технических решений',
    'координатор по внедрению тех. решений': 'Координатор по внедрению технических решений'
  };
  function normalizeUnitToken(t) {
    let x = normSpace(String(t || ''));
    x = x.replace(/НОР\s*-\s*/gi, 'НОР-');
    const al = UNIT_ALIAS[x.toLowerCase()];
    return al || x;
  }
  function splitUnits(raw, index) {
    const chunks = String(raw || '').split(/[\n\r;]+/).map(function(p){ return p.trim(); }).filter(Boolean);
    const known = knownUnitList(index);
    const out = [];
    chunks.forEach(function(part) {
      let work = part;
      if (/руководители\s+сп/i.test(work)) { out.push('Руководители СП'); work = work.replace(/руководители\s+сп/ig, ' '); }
      if (/директора\s+нор/i.test(work)) { out.push('Директора НОР'); work = work.replace(/директора\s+нор/ig, ' '); }
      work = normalizeUnitToken(work);
      if (!work) return;
      if (known.includes(work)) { out.push(work); return; }
      const knownNorm = known.map(function(u){ return { raw:u, n: normalizeUnitToken(u) }; });
      const hit = knownNorm.find(function(u){ return u.n.toLowerCase() === work.toLowerCase(); });
      if (hit) { out.push(hit.raw); return; }
      let rest = work;
      knownNorm.forEach(function(u) {
        const re = new RegExp(u.n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i');
        if (re.test(rest)) { out.push(u.raw); rest = rest.replace(re, ' '); }
      });
      rest = normalizeUnitToken(rest.replace(/[,/|+]+/g, ' '));
      if (rest && rest.length > 2) out.push(rest);
    });
    return [...new Set(out.filter(Boolean))];
  }
  function expandSpecial(index, tokens) {
    const out = [];
    tokens.forEach(t => {
      if (/^руководители сп$/i.test(t) && index) out.push(...Object.keys(index.spLeaderByUnit));
      else if (/^директора нор$/i.test(t) && index) out.push(...Object.keys(index.norDirectorByUnit));
      else if (t) out.push(t);
    });
    return [...new Set(out)];
  }
  function protoAction(r) {
    const st = r.executionStatusNormalized;
    if (/просрочено/i.test(st)) return 'Уточнить исполнение и результат';
    if (/снят с контроля/i.test(st)) return 'Действий не требуется';
    if (/на снятии/i.test(st)) return 'Рассмотреть снятие с контроля';
    if (/рабочий контроль/i.test(st)) return 'Продолжить контроль';
    if (r.deadlineType === 'TEXT') return 'Контроль по установленной периодичности';
    if (r.daysDelta === 0) return 'Проверить исполнение сегодня';
    if (r.daysDelta != null && r.daysDelta > 0 && r.daysDelta <= 3) return 'Проверить готовность';
    if (r.daysDelta != null && r.daysDelta < 0) return 'Уточнить исполнение и результат';
    return 'Плановый контроль';
  }
  function appealAction(r) {
    if (/завершено/i.test(r.executionStatusNormalized)) return 'Действий не требуется';
    if (r.daysDelta != null && r.daysDelta < 0) return 'Уточнить подготовку ответа';
    if (r.daysDelta === 0) return 'Проверить готовность ответа';
    if (r.daysDelta != null && r.daysDelta <= 3) return 'Контроль подготовки ответа';
    return 'Плановый контроль';
  }
  function inAction(r) {
    if (/выполнен/i.test(r.executionStatusNormalized)) return 'Действий не требуется';
    if (r.daysDelta != null && r.daysDelta < 0) return 'Уточнить исполнение';
    if (r.daysDelta === 0) return 'Проверить исполнение сегодня';
    if (r.daysDelta != null && r.daysDelta <= 3) return 'Проверить готовность';
    return 'Плановый контроль';
  }
  function finishProto(r) {
    r.daysDelta = r.deadlineDate ? daysBetween(r.deadlineDate, refDate()) : null;
    r.deadlineDisplay = deadlineDisplay(r.deadlineType, r.deadlineDate, r.deadlineText, r.daysDelta, r.executionStatusNormalized, 'protocol');
    r.nextAction = protoAction(r); r.warnings = r.warnings || [];
    const w = [];
    if (!r.executionStatusNormalized) w.push('пустой статус');
    else if (!/в работе|рабочий контроль|просрочено|снят с контроля|на снятии|постоянно|не исполнен/i.test(r.executionStatusNormalized)) w.push('неизвестный статус: ' + r.executionStatusNormalized);
    if (r.deadlineType === 'MISSING' && !/снят с контроля/i.test(r.executionStatusNormalized)) w.push('отсутствующий срок у активной записи');
    if (r.deadlineDate && r.daysDelta != null && r.daysDelta < 0 && !/просрочено|снят с контроля/i.test(r.executionStatusNormalized)) w.push('Срок истёк, но статус источника не «Просрочено»');
    if (/просрочено/i.test(r.executionStatusNormalized) && r.daysDelta != null && r.daysDelta >= 0) w.push('Статус источника «Просрочено», однако срок исполнения ещё не наступил');
    if (S.structure && (r.resolvedUnits||[]).some(u => u !== 'Не определено' && !S.structure.units.some(x => x.toLowerCase() === String(u).toLowerCase()) && !/^руководители сп$|^директора нор$/i.test(u))) w.push('неизвестное структурное подразделение: ' + (r.resolvedUnits||[]).join(', '));
    r.warnings = w;
    return r;
  }
  function finishAppeal(r) {
    r.daysDelta = r.deadlineDate ? daysBetween(r.deadlineDate, refDate()) : null;
    r.deadlineDisplay = deadlineDisplay(r.deadlineType, r.deadlineDate, r.deadlineText, r.daysDelta, r.executionStatusNormalized, 'appeal');
    r.nextAction = appealAction(r);
    const w = [...((r.organization && r.organization.warnings) || [])];
    if (!r.executionStatusNormalized) w.push('Пустой статус исполнения');
    else if (!/на исполнении|завершено/i.test(r.executionStatusNormalized)) w.push('Неизвестный статус: ' + r.executionStatusNormalized);
    if (r.deadlineType === 'MISSING' && !/завершено/i.test(r.executionStatusNormalized)) w.push('Не указан срок исполнения');
    if (r.sourceDepartment && r.resolvedUnit !== 'Не определено' && normSpace(r.sourceDepartment).toLowerCase() !== normSpace(r.resolvedUnit).toLowerCase()) w.push('Исходное подразделение отличается от structure.xlsx');
    r.warnings = [...new Set(w)]; return r;
  }
  function finishIncoming(r) {
    r.daysDelta = r.deadlineDate ? daysBetween(r.deadlineDate, refDate()) : null;
    r.deadlineDisplay = deadlineDisplay(r.deadlineType, r.deadlineDate, r.deadlineText, r.daysDelta, r.executionStatusNormalized, 'incoming');
    r.nextAction = inAction(r);
    const w = [...((r.organization && r.organization.warnings) || [])];
    if (!r.executionStatusNormalized) w.push('Пустой статус исполнения');
    else if (!/не исполнено|не исполнен|на исполнении|выполнен|выполнено/i.test(r.executionStatusNormalized)) w.push('Неизвестный статус: ' + r.executionStatusNormalized);
    if (r.deadlineType === 'MISSING' && !/выполнен|выполнено/i.test(r.executionStatusNormalized)) w.push('Не указан срок исполнения');
    if (r.sourceDepartment && r.resolvedUnit !== 'Не определено' && normSpace(r.sourceDepartment).toLowerCase() !== normSpace(r.resolvedUnit).toLowerCase()) w.push('Исходное подразделение отличается от structure.xlsx');
    r.warnings = [...new Set(w)]; return r;
  }

  const PROTO_REQ = ['№ Протокола','Дата','№ поручения','Содержание поручений','Структурное подразделение','Срок исполнения','Информация о ходе исполнения','Статус исполнения'];
  const APP_REQ = ['Номер обращения','Дата регистрации обращения','Автор обращения','Вид обращения','Краткое содержание','Срок исполнения','Дата предоставления ответа','Статус исполнения','Ответственный исполнитель'];
  const IN_REQ = ['Рег. номер и дата входящего документа','Краткое содержание','Срок исполнения','Статус исполнения'];

  function previewProgress(text) {
    const t = String(text || '').replace(/\s+/g, ' ').trim();
    if (!t) return '';
    const parts = t.split(/(?<=[.!?])\s+/);
    let out = parts[0] || t;
    if (out.length < 90 && parts[1]) out += ' ' + parts[1];
    if (out.length > 220) out = out.slice(0, 217).replace(/\s+\S*$/, '') + '…';
    else if (out.length < t.length) out += '…';
    return out;
  }
  function markDuplicateWarnings(records, keyFn) {
    const counts = new Map();
    records.forEach(r => { const k=normSpace(keyFn(r)); if(k) counts.set(k,(counts.get(k)||0)+1); });
    records.forEach(r => { const k=normSpace(keyFn(r)); if(k && counts.get(k)>1) { r.warnings=r.warnings||[]; r.warnings.push('Возможный дубликат исходного идентификатора'); r.warnings=[...new Set(r.warnings)]; } });
  }
  function parseProtocols(wb, filename) {
    const found = detectSheet(wb, PROTO_REQ);
    if (!found) return { records:[], meta:{ filename, errors:['Не найден лист протоколов'], valid:0, ignored:0, sheet:'', headerRow:0, loadedAt:new Date().toISOString() } };
    const { matrix, header, sheetName } = found, records = []; let ignored = 0;
    for (let i = header.rowIndex + 1; i < matrix.length; i++) {
      const row = matrix[i]; if (emptyRow(row)) continue;
      const protocolNumberRaw = cellStr(row, header.colMap, '№ Протокола');
      const assignmentText = cellStr(row, header.colMap, 'Содержание поручений');
      const assignmentNumberRaw = cellStr(row, header.colMap, '№ поручения');
      const status = cellStr(row, header.colMap, 'Статус исполнения');
      if (!assignmentNumberRaw && !assignmentText) { ignored++; continue; }
      if (/^протокол\s*№/i.test(assignmentText) && !assignmentNumberRaw) { ignored++; continue; }
      if (/^протокол\s*№/i.test(protocolNumberRaw) && /\sот\s/i.test(protocolNumberRaw) && !assignmentNumberRaw && !status) { ignored++; continue; }
      const rawUnit = cellStr(row, header.colMap, 'Структурное подразделение');
      const tokens = splitUnits(rawUnit, S.structure), resolved = expandSpecial(S.structure, tokens);
      const dl = classifyDeadline(cell(row, header.colMap, 'Срок исполнения'));
      const progress = cellStr(row, header.colMap, 'Информация о ходе исполнения');
      records.push(finishProto({ kind:'p', id:'p-'+i+'-'+protocolNumberRaw, protocolNumberRaw, protocolDate: parseDate(cell(row, header.colMap, 'Дата')), assignmentNumberRaw: assignmentNumberRaw, assignmentText, sourceStructuralUnitRaw: rawUnit, sourceStructuralTokens: tokens, resolvedUnits: resolved, deadlineDate: dl.date, deadlineType: dl.type, deadlineText: dl.text, progressInfo: progress, progressPreview: previewProgress(progress), executionStatusNormalized: normStatus(status), organization: enrichUnits(S.structure, resolved), sourceRowNumber: i+1, warnings:[] }));
    }
    markDuplicateWarnings(records, r => (r.protocolNumberRaw||'')+'|'+(r.assignmentNumberRaw||''));
    return { records, meta:{ filename, sheet:sheetName, headerRow:header.rowIndex+1, valid:records.length, ignored, errors:[], loadedAt:new Date().toISOString() } };
  }
  function parseAppeals(wb, filename) {
    const found = detectSheet(wb, APP_REQ);
    if (!found) return { records:[], meta:{ filename, errors:['Не найден лист e-Өтініш'], valid:0, ignored:0, sheet:'', headerRow:0, loadedAt:new Date().toISOString() } };
    const { matrix, header, sheetName } = found, records = []; let ignored = 0;
    for (let i = header.rowIndex + 1; i < matrix.length; i++) {
      const row = matrix[i]; if (emptyRow(row)) continue;
      const number = cellStr(row, header.colMap, 'Номер обращения'); if (!number) { ignored++; continue; }
      const fio = cellStr(row, header.colMap, 'Ответственный исполнитель');
      const org = enrichByPerson(S.structure, fio);
      const dl = classifyDeadline(cell(row, header.colMap, 'Срок исполнения'));
      records.push(finishAppeal({ kind:'a', id:'a-'+number+'-'+i, number, registrationDate: parseDate(cell(row, header.colMap, 'Дата регистрации обращения')), author: cellStr(row, header.colMap, 'Автор обращения'), type: cellStr(row, header.colMap, 'Вид обращения'), summary: cellStr(row, header.colMap, 'Краткое содержание'), deadlineDate: dl.date, deadlineType: dl.type, deadlineText: dl.text, responseDate: parseDate(cell(row, header.colMap, 'Дата предоставления ответа')), executionStatusNormalized: normStatus(cellStr(row, header.colMap, 'Статус исполнения')), responsibleFio: fio, sourceDepartment: cellStr(row, header.colMap, 'Структурное подразделение'), resolvedUnit: org.resolvedUnits.length === 1 ? org.resolvedUnits[0] : 'Не определено', organization: org, sourceRowNumber: i+1, warnings: org.warnings }));
    }
    markDuplicateWarnings(records, r => r.number||'');
    return { records, meta:{ filename, sheet:sheetName, headerRow:header.rowIndex+1, valid:records.length, ignored, errors:[], loadedAt:new Date().toISOString() } };
  }
  function parseIncoming(wb, filename) {
    const found = detectSheet(wb, IN_REQ);
    if (!found) return { records:[], meta:{ filename, errors:['Не найден лист входящих с обязательными заголовками'], valid:0, ignored:0, sheet:'', headerRow:0, loadedAt:new Date().toISOString() } };
    const { matrix, header, sheetName } = found, records = []; let ignored = 0;
    const hasResponsible = header.colMap[normHeader('Ответственный исполнитель')] != null;
    const hasDept = header.colMap[normHeader('Подразделение')] != null || header.colMap[normHeader('Структурное подразделение')] != null;
    for (let i = header.rowIndex + 1; i < matrix.length; i++) {
      const row = matrix[i]; if (emptyRow(row)) continue;
      const reg = cellStr(row, header.colMap, 'Рег. номер и дата входящего документа'); if (!reg) { ignored++; continue; }
      const fio = hasResponsible ? cellStr(row, header.colMap, 'Ответственный исполнитель') : '';
      const org = enrichByPerson(S.structure, fio);
      if (!hasResponsible) org.warnings.push('В файле отсутствует «Ответственный исполнитель»; структура не может быть определена по Ф.И.О.');
      const dl = classifyDeadline(cell(row, header.colMap, 'Срок исполнения'));
      const sourceDepartment = hasDept ? (cellStr(row, header.colMap, 'Подразделение') || cellStr(row, header.colMap, 'Структурное подразделение')) : '';
      const rec = { kind:'i', id:'i-'+reg+'-'+i, regNumberDate: reg, summary: cellStr(row, header.colMap, 'Краткое содержание'), responsibleFio: fio, sourceDepartment, resolvedUnit: org.resolvedUnits.length === 1 ? org.resolvedUnits[0] : 'Не определено', deadlineDate: dl.date, deadlineType: dl.type, deadlineText: dl.text, executionStatusNormalized: normStatus(cellStr(row, header.colMap, 'Статус исполнения')), organization: org, sourceRowNumber: i+1, warnings: org.warnings };
      records.push(finishIncoming(rec));
    }
    markDuplicateWarnings(records, r => r.regNumberDate||'');
    const warnings = [];
    if (!hasResponsible) warnings.push('Нет столбца «Ответственный исполнитель»: официальное СП будет «Не определено».');
    if (!hasDept) warnings.push('Нет исходного столбца подразделения: audit-сравнение недоступно.');
    return { records, meta:{ filename, sheet:sheetName, headerRow:header.rowIndex+1, valid:records.length, ignored, errors:[], warnings, loadedAt:new Date().toISOString() } };
  }
  function reenrich() {
    S.protocols = S.protocols.map(r => { const tokens = splitUnits(r.sourceStructuralUnitRaw, S.structure); r.sourceStructuralTokens = tokens; const resolved = expandSpecial(S.structure, tokens); r.resolvedUnits = resolved; r.organization = enrichUnits(S.structure, resolved); return finishProto(r); });
    S.appeals = S.appeals.map(r => { r.organization = enrichByPerson(S.structure, r.responsibleFio); r.resolvedUnit = r.organization.resolvedUnits.length === 1 ? r.organization.resolvedUnits[0] : 'Не определено'; r.warnings = r.organization.warnings; return finishAppeal(r); });
    S.incoming = S.incoming.map(r => { r.organization = enrichByPerson(S.structure, r.responsibleFio); r.resolvedUnit = r.organization.resolvedUnits.length === 1 ? r.organization.resolvedUnits[0] : 'Не определено'; r.warnings = r.organization.warnings; return finishIncoming(r); });
    sanitizeStructureFilters();
  }
  function sanitizeStructureFilters() {
    let changed = false;
    const validFor = (rows) => ({
      units:new Set(rows.flatMap(r => r.kind==='p' ? (r.resolvedUnits||[]) : [r.resolvedUnit]).filter(Boolean)),
      curators:new Set(rows.flatMap(r => (r.organization&&r.organization.curators)||[])),
      heads:new Set(rows.flatMap(r => (r.organization&&r.organization.departmentHeads)||[])),
      spLeaders:new Set(rows.flatMap(r => (r.organization&&r.organization.spLeaders)||[])),
      norDirectors:new Set(rows.flatMap(r => (r.organization&&r.organization.norDirectors)||[])),
      joint:new Set(rows.flatMap(r => ((r.organization&&r.organization.jointGroups)||[]).concat(r.organization&&r.organization.isJointExecution?['да']:[]))),
      positions:new Set(rows.flatMap(r => (r.organization&&r.organization.positions)||[]))
    });
    [['p',S.protocols],['a',S.appeals],['i',S.incoming]].forEach(([k,rows]) => {
      const valid=validFor(rows), f=F[k];
      ['units','curators','heads','spLeaders','norDirectors','joint','positions'].forEach(field => {
        const before=(f[field]||[]).length; f[field]=(f[field]||[]).filter(v=>valid[field].has(v)); if(f[field].length!==before) changed=true;
      });
    });
    if (changed) S.notice='Организационная структура обновлена. Недействительные значения фильтров сброшены.';
  }
  function serRows(rows) {
    return rows.map(r => Object.assign({}, r, { protocolDate: r.protocolDate && r.protocolDate.toISOString(), registrationDate: r.registrationDate && r.registrationDate.toISOString(), responseDate: r.responseDate && r.responseDate.toISOString(), deadlineDate: r.deadlineDate && r.deadlineDate.toISOString() }));
  }
  function revRows(rows) {
    return (rows || []).map(r => Object.assign({}, r, { protocolDate: r.protocolDate && new Date(r.protocolDate), registrationDate: r.registrationDate && new Date(r.registrationDate), responseDate: r.responseDate && new Date(r.responseDate), deadlineDate: r.deadlineDate && new Date(r.deadlineDate) }));
  }
  function idb() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB, 1);
      req.onupgradeneeded = () => { if (!req.result.objectStoreNames.contains('ops')) req.result.createObjectStore('ops'); };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  async function persist() {
    const payload = { ver:7, protocols:serRows(S.protocols), appeals:serRows(S.appeals), incoming:serRows(S.incoming), protoMeta:S.protoMeta, appealMeta:S.appealMeta, inMeta:S.inMeta };
    try {
      const db = await idb();
      await new Promise((res, rej) => { const tx = db.transaction('ops','readwrite'); tx.objectStore('ops').put(payload, 'data'); tx.oncomplete = res; tx.onerror = () => rej(tx.error); });
    } catch (e) { try { localStorage.setItem(DB, JSON.stringify(payload)); } catch(_){} }
  }
  async function restore() {
    let p = null;
    try {
      const db = await idb();
      p = await new Promise((res) => { const tx = db.transaction('ops','readonly'); const rq = tx.objectStore('ops').get('data'); rq.onsuccess = () => res(rq.result); rq.onerror = () => res(null); });
    } catch (e) {
      try { p = JSON.parse(localStorage.getItem(DB) || 'null'); } catch(_){}
    }
    if (!p || p.ver !== 7) return;
    S.protocols = revRows(p.protocols); S.appeals = revRows(p.appeals); S.incoming = revRows(p.incoming);
    S.protoMeta = p.protoMeta; S.appealMeta = p.appealMeta; S.inMeta = p.inMeta;
  }
  const badgeClass = (s) => {
    if (/на снятии/i.test(s)) return 'b-orange';
    if (/просроч|не исполнен/i.test(s)) return 'b-red';
    if (/снят с контроля|заверш|выполнен/i.test(s)) return 'b-green';
    if (/рабочий контроль/i.test(s)) return 'b-amber';
    if (/в работе|на исполнении/i.test(s)) return 'b-sky';
    return 'b-gray';
  };
  const isFinalStatus = (st) => /снят с контроля|завершено|выполнен/i.test(st || '');
  const isAppealOverdue = (r) => !isFinalStatus(r.executionStatusNormalized) && r.daysDelta != null && r.daysDelta < 0;
  const isInOverdue = (r) => !isFinalStatus(r.executionStatusNormalized) && r.daysDelta != null && r.daysDelta < 0;
  function inRange(d, from, to) {
    if (!d) return false;
    if (from && startOfDay(d) < startOfDay(new Date(from))) return false;
    if (to && startOfDay(d) > startOfDay(new Date(to))) return false;
    return true;
  }
  function matchPeriod(d, f) {
    if (f.period === 'all' && !f.dateFrom && !f.dateTo && !f.dateMonth && !f.dateYear) return true;
    if (!d) return false;
    if (f.period === 'custom' || f.dateFrom || f.dateTo || f.dateMonth || f.dateYear) {
      if (f.dateMonth) { const p = f.dateMonth.split('-'); if (d.getFullYear()!=+p[0] || (d.getMonth()+1)!=+p[1]) return false; }
      if (f.dateYear && String(d.getFullYear()) !== String(f.dateYear)) return false;
      if ((f.dateFrom || f.dateTo) && !inRange(d, f.dateFrom, f.dateTo)) return false;
      return true;
    }
    const ref = refDate();
    if (f.period === 'today') return daysBetween(d, ref) === 0;
    if (f.period === 'week') { const x = new Date(ref); const day = (x.getDay()+6)%7; x.setDate(x.getDate()-day); const e = new Date(x); e.setDate(x.getDate()+6); return d >= x && d <= e; }
    if (f.period === 'month') return d.getMonth()===ref.getMonth() && d.getFullYear()===ref.getFullYear();
    if (f.period === 'quarter') return d.getFullYear()===ref.getFullYear() && Math.floor(d.getMonth()/3)===Math.floor(ref.getMonth()/3);
    if (f.period === 'year') return d.getFullYear()===ref.getFullYear();
    return true;
  }
  function matchDeadline(r, f) {
    const p = f.dlPreset;
    if (!p && !f.dlFrom && !f.dlTo && !f.dlMonth && !f.dlYear) return true;
    const t = (r.deadlineText || '').toLowerCase();
    if (p === 'missing') return r.deadlineType === 'MISSING';
    if (p && p.startsWith('text:')) return t === p.slice(5);
    if (p === 'постоянно' || p === 'еженедельно' || p === 'ежедневно' || p === 'ежемесячно') return t.includes(p);
    if (p === 'text') return r.deadlineType === 'TEXT';
    if (!r.deadlineDate) return false;
    if (p === 'overdue') return r.daysDelta != null && r.daysDelta < 0;
    if (p === 'today') return r.daysDelta === 0;
    if (p === '1-3') return r.daysDelta >= 1 && r.daysDelta <= 3;
    if (p === '4-7') return r.daysDelta >= 4 && r.daysDelta <= 7;
    if (p === '8-30') return r.daysDelta >= 8 && r.daysDelta <= 30;
    if (p === '>30') return r.daysDelta > 30;
    if (f.dlMonth) { const z = f.dlMonth.split('-'); if (r.deadlineDate.getFullYear()!=+z[0] || (r.deadlineDate.getMonth()+1)!=+z[1]) return false; }
    if (f.dlYear && String(r.deadlineDate.getFullYear()) !== String(f.dlYear)) return false;
    if (f.dlFrom || f.dlTo) return inRange(r.deadlineDate, f.dlFrom, f.dlTo);
    return true;
  }
  function hasAny(sel, values) {
    if (!sel || !sel.length) return true;
    const vs = (values || []).map(String);
    return sel.some(x => vs.includes(x));
  }
  function extractRegDate(s) {
    const m = String(s||'').match(/(\d{1,2}[./]\d{1,2}[./]\d{4})/);
    return m ? parseDate(m[1]) : null;
  }
  function compactUnits(units) {
    const u = units || [];
    if (u.length <= 2) return u.join(' + ') || '—';
    return u.slice(0,2).join(' + ') + ' + ещё ' + (u.length-2);
  }


  function protoBase(skip) {
    const f = F.p; skip = skip || {};
    return S.protocols.filter(r => {
      if (f.q) { const hay = [r.protocolNumberRaw, r.assignmentNumberRaw, r.assignmentText, r.sourceStructuralUnitRaw, r.progressInfo, (r.resolvedUnits || []).join(' ')].join(' ').toLowerCase(); if (!hay.includes(f.q.toLowerCase())) return false; }
      if (!skip.units && !hasAny(f.units, r.resolvedUnits)) return false;
      if (!skip.assigned && !hasAny(f.assigned, r.sourceStructuralTokens || splitUnits(r.sourceStructuralUnitRaw, S.structure))) return false;
      if (!skip.curators && !hasAny(f.curators, r.organization.curators)) return false;
      if (!skip.heads && !hasAny(f.heads, r.organization.departmentHeads)) return false;
      if (!skip.spLeaders && !hasAny(f.spLeaders, r.organization.spLeaders)) return false;
      if (!skip.norDirectors && !hasAny(f.norDirectors, r.organization.norDirectors)) return false;
      if (!skip.joint && f.joint && f.joint.length) {
        const ok = hasAny(f.joint, r.organization.jointGroups) || (f.joint.includes('да') && r.organization.isJointExecution);
        if (!ok) return false;
      }
      if (!hasAny(f.statuses, [r.executionStatusNormalized])) return false;
      if (!matchPeriod(r.protocolDate, f)) return false;
      if (!matchDeadline(r, f)) return false;
      return true;
    });
  }
  function protoView(base) {
    let rows = base;
    if (S.protoKpi === 'overdue') rows = rows.filter(r => /просрочено/i.test(r.executionStatusNormalized));
    if (S.protoKpi === 'work') rows = rows.filter(r => r.executionStatusNormalized === 'В работе');
    if (S.protoKpi === 'watch') rows = rows.filter(r => r.executionStatusNormalized === 'Рабочий контроль');
    if (S.protoKpi === 'removing') rows = rows.filter(r => r.executionStatusNormalized === 'На снятии с контроля');
    if (S.protoKpi === 'done') rows = rows.filter(r => r.executionStatusNormalized === 'Снят с контроля');
    return rows;
  }
  function appealBase(skip) {
    const f = F.a; skip = skip || {};
    return S.appeals.filter(r => {
      if (f.q) { const hay = [r.number, r.author, r.type, r.summary, r.resolvedUnit].join(' ').toLowerCase(); if (!hay.includes(f.q.toLowerCase())) return false; }
      if (f.number && !(r.number||'').toLowerCase().includes(f.number.toLowerCase())) return false;
      if (f.author && !(r.author||'').toLowerCase().includes(f.author.toLowerCase())) return false;
      if (f.summary && !(r.summary||'').toLowerCase().includes(f.summary.toLowerCase())) return false;
      if (!hasAny(f.types, [r.type])) return false;
      if (!hasAny(f.units, [r.resolvedUnit])) return false;
      if (!skip.curators && !hasAny(f.curators, r.organization.curators)) return false;
      if (!skip.heads && !hasAny(f.heads, r.organization.departmentHeads)) return false;
      if (!skip.spLeaders && !hasAny(f.spLeaders, r.organization.spLeaders)) return false;
      if (!skip.norDirectors && !hasAny(f.norDirectors, r.organization.norDirectors)) return false;
      if (!skip.joint && f.joint && f.joint.length && !hasAny(f.joint, r.organization.jointGroups) && !(f.joint.includes('да') && r.organization.isJointExecution)) return false;
      if (!hasAny(f.statuses, [r.executionStatusNormalized])) return false;
      if (!matchPeriod(r.registrationDate, f)) return false;
      if (!matchDeadline(r, f)) return false;
      if (f.respMissing) { if (r.responseDate) return false; }
      else if (f.respFrom || f.respTo || f.respMonth || f.respYear) {
        if (!r.responseDate) return false;
        if (f.respMonth) { const z = f.respMonth.split('-'); if (r.responseDate.getFullYear()!=+z[0] || (r.responseDate.getMonth()+1)!=+z[1]) return false; }
        if (f.respYear && String(r.responseDate.getFullYear()) !== String(f.respYear)) return false;
        if ((f.respFrom || f.respTo) && !inRange(r.responseDate, f.respFrom, f.respTo)) return false;
      }
      return true;
    });
  }
  function appealView(base) {
    let rows = base;
    if (S.appealKpi === 'overdue') rows = rows.filter(isAppealOverdue);
    else if (S.appealKpi === 'soon') rows = rows.filter(r => !/завершено/i.test(r.executionStatusNormalized) && r.daysDelta != null && r.daysDelta >= 0 && r.daysDelta <= 7);
    else if (S.appealKpi === 'done') rows = rows.filter(r => /завершено/i.test(r.executionStatusNormalized));
    else if (S.tabA === 'overdue') rows = rows.filter(isAppealOverdue);
    else if (S.tabA === 'done') rows = rows.filter(r => /завершено/i.test(r.executionStatusNormalized));
    else if (S.tabA === 'work') rows = rows.filter(r => /на исполнении/i.test(r.executionStatusNormalized));
    return rows;
  }
  function inBase(skip) {
    const f = F.i; skip = skip || {};
    return S.incoming.filter(r => {
      if (f.q) { const hay = [r.regNumberDate, r.summary, r.resolvedUnit].join(' ').toLowerCase(); if (!hay.includes(f.q.toLowerCase())) return false; }
      if (f.number && !(r.regNumberDate||'').toLowerCase().includes(f.number.toLowerCase())) return false;
      if (f.summary && !(r.summary||'').toLowerCase().includes(f.summary.toLowerCase())) return false;
      if (!hasAny(f.units, [r.resolvedUnit])) return false;
      if (!skip.curators && !hasAny(f.curators, r.organization.curators)) return false;
      if (!skip.heads && !hasAny(f.heads, r.organization.departmentHeads)) return false;
      if (!skip.spLeaders && !hasAny(f.spLeaders, r.organization.spLeaders)) return false;
      if (!skip.norDirectors && !hasAny(f.norDirectors, r.organization.norDirectors)) return false;
      if (!skip.positions && !hasAny(f.positions, r.organization.positions)) return false;
      if (!skip.joint && f.joint && f.joint.length && !hasAny(f.joint, r.organization.jointGroups) && !(f.joint.includes('да') && r.organization.isJointExecution)) return false;
      if (!hasAny(f.statuses, [r.executionStatusNormalized])) return false;
      if (!matchDeadline(r, f)) return false;
      const rd = extractRegDate(r.regNumberDate);
      if ((f.period !== 'all' || f.dateFrom || f.dateTo || f.dateMonth || f.dateYear) && !matchPeriod(rd, f)) return false;
      return true;
    });
  }
  function inView(base) {
    let rows = base;
    if (S.inKpi === 'overdue') rows = rows.filter(isInOverdue);
    else if (S.inKpi === 'soon') rows = rows.filter(r => !/выполнен/i.test(r.executionStatusNormalized) && r.daysDelta != null && r.daysDelta >= 0 && r.daysDelta <= 7);
    else if (S.inKpi === 'done') rows = rows.filter(r => /выполнен/i.test(r.executionStatusNormalized));
    else if (S.tabI === 'overdue') rows = rows.filter(isInOverdue);
    else if (S.tabI === 'done') rows = rows.filter(r => /выполнен/i.test(r.executionStatusNormalized));
    else if (S.tabI === 'work') rows = rows.filter(r => /на исполнении/i.test(r.executionStatusNormalized));
    return rows;
  }
  function exportX(rows, name) {
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'data'); XLSX.writeFile(wb, name);
  }


  function ms(key, label, opts, selected, displayFn) {
    selected = selected || [];
    const box = h('div', { class:'ms' });
    const btn = h('button', { class:'ms-btn', type:'button' }, label + (selected.length ? ' ('+selected.length+')' : ''));
    btn.onclick = (e) => { e.stopPropagation(); openMs = openMs===key ? '' : key; render(); };
    box.appendChild(btn);
    if (openMs === key) {
      const pop = h('div', { class:'ms-pop' });
      pop.addEventListener('click', e => e.stopPropagation());
      const search = h('input', { placeholder:'Поиск' });
      search.addEventListener('input', () => {
        const q = search.value.toLowerCase();
        pop.querySelectorAll('label[data-opt]').forEach(l => { l.style.display = l.getAttribute('data-opt').toLowerCase().includes(q) ? '' : 'none'; });
      });
      pop.appendChild(search);
      const acts = h('div', { class:'ms-actions' }, [h('button',{class:'btn ghost',type:'button'},'Все'), h('button',{class:'btn ghost',type:'button'},'Очистить')]);
      acts.children[0].onclick = () => { selected.splice(0, selected.length, ...opts); render(); };
      acts.children[1].onclick = () => { selected.splice(0, selected.length); render(); };
      pop.appendChild(acts);
      opts.forEach(o => {
        const shown = displayFn ? displayFn(o) : o;
        const lab = h('label', { 'data-opt': shown });
        const cb = h('input', { type:'checkbox' });
        if (selected.includes(o)) cb.checked = true;
        cb.onchange = () => {
          const arr = selected.slice();
          const i = arr.indexOf(o);
          if (cb.checked && i<0) arr.push(o);
          if (!cb.checked && i>=0) arr.splice(i,1);
          selected.splice(0, selected.length, ...arr);
          render();
        };
        lab.appendChild(cb); lab.appendChild(document.createTextNode(' '+shown));
        pop.appendChild(lab);
      });
      box.appendChild(pop);
    }
    return box;
  }
  function resetF(which) { F[which] = emptyF(); render(); }
  function chipBar(f, which, extra) {
    const chips = [];
    const add = (k, arr, clear) => (arr||[]).forEach(v => {
      const c = h('span', { class:'chip' }, [k+': '+v+' ', h('button', { type:'button' }, '×')]);
      c.querySelector('button').onclick = () => { clear(v); render(); };
      chips.push(c);
    });
    add('СП', f.units, v => f.units = f.units.filter(x=>x!==v));
    add('Назначено', f.assigned, v => f.assigned = f.assigned.filter(x=>x!==v));
    add('Куратор', f.curators, v => f.curators = f.curators.filter(x=>x!==v));
    add('Рук. подр.', f.heads, v => f.heads = f.heads.filter(x=>x!==v));
    add('Руководитель СП', f.spLeaders, v => f.spLeaders = f.spLeaders.filter(x=>x!==v));
    add('Директор НОР', f.norDirectors, v => f.norDirectors = f.norDirectors.filter(x=>x!==v));
    add('Совместное', f.joint, v => f.joint = f.joint.filter(x=>x!==v));
    add('Статус', f.statuses, v => f.statuses = f.statuses.filter(x=>x!==v));
    add('Вид', f.types, v => f.types = f.types.filter(x=>x!==v));
    if (f.period && f.period!=='all') add('Период', [f.period], () => f.period='all');
    if (f.dlPreset) add('Срок', [f.dlPreset], () => f.dlPreset='');
    if (f.number) add('Номер', [f.number], () => f.number='');
    const bar = h('div', { class:'chips' }, chips);
    const rst = h('button', { class:'btn ghost', type:'button' }, 'Сбросить все');
    rst.onclick = () => resetF(which);
    bar.appendChild(rst);
    const cnt = ['units','assigned','curators','heads','spLeaders','norDirectors','joint','statuses','types','positions'].reduce((n,k)=>n+(f[k]||[]).length,0) + ['number','author','summary','dateFrom','dateTo','dateMonth','dateYear','dlPreset','dlFrom','dlTo','dlMonth','dlYear','respFrom','respTo','respMonth','respYear'].reduce((n,k)=>n+(f[k]?1:0),0) + (f.period && f.period!=='all'?1:0) + (f.respMissing?1:0);
    bar.appendChild(h('span', { class:'filter-count' }, 'Активных фильтров: ' + cnt));
    bar.appendChild(h('span', { class:'found' }, extra || ''));
    return bar;
  }
  function periodSel(f) {
    const s = h('select', {});
    [['all','Период: все'],['today','Сегодня'],['week','Текущая неделя'],['month','Текущий месяц'],['quarter','Текущий квартал'],['year','Текущий год'],['custom','Произвольный период']].forEach(([v,l]) => s.appendChild(h('option', { value:v }, l)));
    s.value = f.period; s.onchange = () => { f.period = s.value; if (s.value !== 'custom') { f.dateFrom='';f.dateTo='';f.dateMonth='';f.dateYear=''; } render(); }; return s;
  }
  function deadlineSel(f, dynamicText) {
    const s = h('select', {});
    const base = [['','Срок исполнения'],['overdue','Просроченный календарный срок'],['today','Сегодня'],['1-3','1–3 дня'],['4-7','4–7 дней'],['8-30','8–30 дней'],['>30','Более 30 дней'],['постоянно','Постоянно'],['еженедельно','Еженедельно'],['ежемесячно','Ежемесячно'],['ежедневно','Ежедневно'],['missing','Не указан']];
    (dynamicText||[]).forEach(t => { const v='text:'+String(t).toLowerCase(); if (!base.some(x=>x[0]===v) && !base.some(x=>String(t).toLowerCase().includes(x[0]) && x[0])) base.push([v,t]); });
    base.forEach(([v,l]) => s.appendChild(h('option', { value:v }, l)));
    s.value = f.dlPreset; s.onchange = () => { f.dlPreset = s.value; if (s.value) { f.dlFrom='';f.dlTo='';f.dlMonth='';f.dlYear=''; } render(); }; return s;
  }
  function sel(value, on, opts, label) {
    const s = h('select', {}, [h('option', { value:'' }, label)].concat(opts.map(o => h('option', { value:o }, o))));
    s.value = value; s.addEventListener('change', e => { on(e.target.value); render(); }); return s;
  }
  function kpi(cls, title, num, hint, active, fn) {
    const el = h('div', { class:'kpi ' + cls + (active ? ' active' : '') }, [h('div',{class:'label'}, title), h('div',{class:'num'}, String(num)), hint ? h('div',{class:'hint'}, hint) : null]);
    el.onclick = fn; return el;
  }
  function textField(fid, placeholder, value, setter, extra) {
    const i = h('input', Object.assign({ placeholder: placeholder, 'data-fid': fid }, extra || {}));
    i.value = value || '';
    i.addEventListener('input', function(e) {
      setter(e.target.value);
      window.__fid = fid; window.__fpos = e.target.selectionStart;
      clearTimeout(window.__ft);
      window.__ft = setTimeout(function(){ render(); }, 280);
    });
    return i;
  }
  function inputQ(f, fid) { return textField(fid || 'q', 'Поиск', f.q, function(v){ f.q = v; }); }
  function tabs(items, cur, on) {
    return h('div', { class:'tabs' }, items.map(([id, label]) => { const b = h('button', { class:'tab' + (cur === id ? ' on' : '') }, label); b.onclick = () => on(id); return b; }));
  }
  function table(headers, rows, recs, sortId) {
    const wrap = h('div', { class:'table-wrap' }), tbl = h('table');
    tbl.appendChild(h('thead', {}, h('tr', {}, headers.map((hh, idx) => {
      const th = h('th', {}, hh + (sortId && SORT[sortId].key===hh ? (SORT[sortId].dir>0?' ▲':' ▼') : ''));
      th.onclick = () => {
        if (!sortId) return;
        if (SORT[sortId].key === hh) SORT[sortId].dir *= -1; else { SORT[sortId].key = hh; SORT[sortId].dir = 1; }
        render();
      };
      return th;
    }))));
    const tb = h('tbody');
    if (!rows.length) tb.appendChild(h('tr', {}, h('td', { colspan:String(headers.length) }, 'Нет записей. Загрузите файлы на странице «Загрузка».')));
    rows.forEach((cols, i) => {
      const tr = h('tr', { class:'clickable' });
      cols.forEach((c, j) => {
        const td = h('td', String(c || '').length > 80 ? { class:'clamp' } : {}, '');
        if (headers[j] === 'Статус исполнения') td.appendChild(h('span', { class:'badge ' + badgeClass(String(c)) }, String(c ?? '')));
        else td.textContent = c == null ? '' : String(c);
        tr.appendChild(td);
      });
      tr.onclick = () => { S.detail = recs[i]; render(); };
      tb.appendChild(tr);
    });
    tbl.appendChild(tb); wrap.appendChild(tbl); return wrap;
  }
  function uploadCard(title, meta, kind) {
    const inp = h('input', { type:'file', accept:'.xlsx,.xls' });
    inp.addEventListener('change', async (e) => {
      const f = e.target.files && e.target.files[0]; if (!f) return;
      try {
        const wb = XLSX.read(await f.arrayBuffer(), { type:'array', cellDates:true });
        if (kind === 1) { const r = parseProtocols(wb, f.name); S.err = r.meta.errors[0] || ''; if (!S.err) { S.protocols = r.records; S.protoMeta = r.meta; } }
        if (kind === 2) { const r = parseAppeals(wb, f.name); S.err = r.meta.errors[0] || ''; if (!S.err) { S.appeals = r.records; S.appealMeta = r.meta; } }
        if (kind === 3) { const r = parseIncoming(wb, f.name); S.err = r.meta.errors[0] || ''; if (!S.err) { S.incoming = r.records; S.inMeta = r.meta; } }
        persist(); render();
      } catch (err) { S.err = String(err.message || err); render(); }
    });
    const kids = [h('h3', {}, title), inp];
    if (meta) { kids.push(h('div', { class:'meta' }, 'Файл: ' + meta.filename + ' · лист: ' + meta.sheet + ' · строка заголовков: ' + meta.headerRow + ' · записей: ' + meta.valid + ' · пропущено: ' + meta.ignored + ' · загружен: ' + new Date(meta.loadedAt).toLocaleString())); (meta.warnings||[]).forEach(w=>kids.push(h('div',{class:'warn'},w))); }
    else kids.push(h('div', { class:'meta' }, 'Файл не загружен'));
    return h('div', { class:'upload-card' }, kids);
  }

  function sourceDeadline(r) { return r.deadlineDate ? fmt(r.deadlineDate) : (r.deadlineText || '—'); }
  function leaderLabel(name, kind) {
    if (!S.structure) return name;
    const maps = kind === 'sp' ? S.structure.spLeaderByUnit : kind === 'nor' ? S.structure.norDirectorByUnit : null;
    const units = [];
    if (maps) Object.entries(maps).forEach(([u,ns]) => { if ((ns||[]).includes(name)) units.push(u); });
    else { Object.entries(S.structure.spLeaderByUnit).forEach(([u,ns]) => { if ((ns||[]).includes(name)) units.push(u); }); Object.entries(S.structure.norDirectorByUnit).forEach(([u,ns]) => { if ((ns||[]).includes(name)) units.push(u); }); }
    return units.length ? name + ' — ' + [...new Set(units)].join(', ') : name;
  }
  function defaultPriority(r, id) {
    if (id === 'p' && /просрочено/i.test(r.executionStatusNormalized)) return 0;
    if ((id === 'a' && isAppealOverdue(r)) || (id === 'i' && isInOverdue(r)) || (id === 'p' && !isFinalStatus(r.executionStatusNormalized) && r.daysDelta != null && r.daysDelta < 0)) return 1;
    if (r.daysDelta === 0) return 2;
    if (r.daysDelta != null && r.daysDelta >= 1 && r.daysDelta <= 3) return 3;
    if (r.daysDelta != null && r.daysDelta >= 4 && r.daysDelta <= 7) return 4;
    if (r.daysDelta != null && r.daysDelta > 7) return 5;
    if (r.deadlineType === 'TEXT') return 6;
    if (r.deadlineType === 'MISSING') return 7;
    return 8;
  }
  function applySort(rows, id, getters) {
    const st = SORT[id]; if (!st || !st.key || !getters[st.key]) return rows.slice().sort((a,b)=>defaultPriority(a,id)-defaultPriority(b,id) || ((a.daysDelta ?? 99999)-(b.daysDelta ?? 99999)));
    const get = getters[st.key], dir = st.dir || 1;
    return rows.slice().sort(function(a,b){
      const va = get(a), vb = get(b);
      if (va instanceof Date && vb instanceof Date) return ((va?va.getTime():0) - (vb?vb.getTime():0)) * dir;
      if (typeof va === 'number' && typeof vb === 'number') return (va - vb) * dir;
      return String(va||'').localeCompare(String(vb||''), 'ru') * dir;
    });
  }
  function render() {
    const root = $('#root'); root.innerHTML = '';
    const nav = (id, icon, label) => { const b = h('button', { class:'navbtn' + (S.page === id ? ' active' : '') }, [icon + ' ', h('span', { class:'label' }, label)]); b.onclick = () => { S.page = id; render(); }; return b; };
    const pb = protoBase(), ab = appealBase(), ib = inBase();
    const PGET = {'№ Протокола':r=>r.protocolNumberRaw,'Дата':r=>r.protocolDate,'№ поручения':r=>r.assignmentNumberRaw,'Содержание поручений':r=>r.assignmentText,'Структурное подразделение':r=>(r.resolvedUnits||[]).join(', '),'Срок исполнения':r=>r.deadlineDate||r.deadlineText,'Отклонение / осталось':r=>r.daysDelta,'Информация о ходе исполнения':r=>r.progressInfo,'Статус исполнения':r=>r.executionStatusNormalized,'Следующее действие':r=>r.nextAction};
    const AGET = {'Номер обращения':r=>r.number,'Дата регистрации':r=>r.registrationDate,'Автор обращения':r=>r.author,'Вид обращения':r=>r.type,'Краткое содержание':r=>r.summary,'Ответственный исполнитель':r=>r.responsibleFio,'Структурное подразделение':r=>r.resolvedUnit,'Срок исполнения':r=>r.deadlineDate,'Отклонение / осталось':r=>r.daysDelta,'Дата предоставления ответа':r=>r.responseDate,'Статус исполнения':r=>r.executionStatusNormalized,'Следующее действие':r=>r.nextAction};
    const IGET = {'Рег. номер и дата входящего документа':r=>r.regNumberDate,'Краткое содержание':r=>r.summary,'Структурное подразделение':r=>r.resolvedUnit,'Срок исполнения':r=>r.deadlineDate,'Отклонение / осталось':r=>r.daysDelta,'Статус исполнения':r=>r.executionStatusNormalized,'Следующее действие':r=>r.nextAction};
    const pv = applySort(protoView(pb), 'p', PGET);
    const av = applySort(appealView(ab), 'a', AGET);
    const iv = applySort(inView(ib), 'i', IGET);
    const pSrc = S.page === 'overview' ? S.protocols : pb;
    const aSrc = S.page === 'overview' ? S.appeals : ab;
    const iSrc = S.page === 'overview' ? S.incoming : ib;
    const pAll = pSrc.length, pOd = pSrc.filter(r => /просрочено/i.test(r.executionStatusNormalized)).length, pWork = pSrc.filter(r => r.executionStatusNormalized === 'В работе').length;
    const pWatch = pSrc.filter(r => r.executionStatusNormalized === 'Рабочий контроль').length, pRem = pSrc.filter(r => r.executionStatusNormalized === 'На снятии с контроля').length, pDone = pSrc.filter(r => r.executionStatusNormalized === 'Снят с контроля').length;
    const aAll = aSrc.length, aOd = aSrc.filter(isAppealOverdue).length, aSoon = aSrc.filter(r => !/завершено/i.test(r.executionStatusNormalized) && r.daysDelta != null && r.daysDelta >= 0 && r.daysDelta <= 7).length, aDone = aSrc.filter(r => /завершено/i.test(r.executionStatusNormalized)).length;
    const iAll = iSrc.length, iOd = iSrc.filter(isInOverdue).length, iSoon = iSrc.filter(r => !/выполнен/i.test(r.executionStatusNormalized) && r.daysDelta != null && r.daysDelta >= 0 && r.daysDelta <= 7).length, iDone = iSrc.filter(r => /выполнен/i.test(r.executionStatusNormalized)).length;
    const units = S.structure ? S.structure.units : [], curators = S.structure ? S.structure.curators : [], heads = [];
    if (S.structure) {
      Object.entries(S.structure.spLeaderByUnit).forEach(([u, ns]) => ns.forEach(n => heads.push(n + ' — ' + u)));
      Object.entries(S.structure.norDirectorByUnit).forEach(([u, ns]) => ns.forEach(n => heads.push(n + ' — ' + u)));
    }
    const mainKids = [];
    if (S.structureError) mainKids.push(h('div', { class:'warn' }, 'structure.xlsx: ' + S.structureError));
    if (S.notice) mainKids.push(h('div', { class:'notice' }, S.notice));
    if (S.err) mainKids.push(h('div', { class:'warn' }, S.err));

    if (S.page === 'overview') {
      const aWork = aSrc.filter(r=>/на исполнении/i.test(r.executionStatusNormalized)).length;
      const iWork = iSrc.filter(r=>/на исполнении|не исполнен/i.test(r.executionStatusNormalized) && !/выполнен/i.test(r.executionStatusNormalized)).length;
      mainKids.push(h('h1', {}, 'Исполнительская дисциплина'), h('h2', {}, 'Протокольные поручения'));
      mainKids.push(h('div', { class:'kpis' }, [kpi('green','ВСЕГО',pAll,'Подтверждённый результат',false,()=>{S.page='protocols';S.protoKpi='';render();}), kpi('red','ПРОСРОЧЕНО',pOd,'Требует вмешательства',false,()=>{S.page='protocols';S.protoKpi='overdue';render();}), kpi('amber','В РАБОТЕ',pWork,'Предупредить просрочку',false,()=>{S.page='protocols';S.protoKpi='work';render();})]));
      mainKids.push(h('div', { class:'kpis' }, [kpi('gold sec','РАБОЧИЙ КОНТРОЛЬ',pWatch,'Контроль исполнения',false,()=>{S.page='protocols';S.protoKpi='watch';render();}), kpi('orange sec','НА СНЯТИИ С КОНТРОЛЯ',pRem,'Ожидает подтверждения',false,()=>{S.page='protocols';S.protoKpi='removing';render();}), kpi('lime sec','СНЯТ С КОНТРОЛЯ',pDone,'Исполнение подтверждено',false,()=>{S.page='protocols';S.protoKpi='done';render();})]));
      mainKids.push(h('h2', {}, 'Другие документы'));
      const cell = (cls,t,n,fn) => { const d=h('div',{class:'mini-kpi '+cls},[h('div',{class:'t'},t),h('div',{class:'n'},String(n))]); d.onclick=(e)=>{e.stopPropagation();fn();}; return d; };
      const p1 = h('div', { class:'panel' }, [h('h3', {}, 'e-Өтініш'), h('div', { class:'mini' }, [cell('green','ВСЕГО',aAll,()=>{S.page='appeals';S.appealKpi='';S.tabA='all';render();}), cell('red','ПРОСРОЧЕНО',aOd,()=>{S.page='appeals';S.appealKpi='overdue';S.tabA='overdue';render();}), cell('amber','В РАБОТЕ',aWork,()=>{S.page='appeals';S.tabA='work';S.appealKpi='';render();})])]);
      const p2 = h('div', { class:'panel' }, [h('h3', {}, 'Входящие по сроку'), h('div', { class:'mini' }, [cell('green','ВСЕГО',iAll,()=>{S.page='incoming';S.inKpi='';S.tabI='all';render();}), cell('red','ПРОСРОЧЕНО',iOd,()=>{S.page='incoming';S.inKpi='overdue';S.tabI='overdue';render();}), cell('amber','В РАБОТЕ',iWork,()=>{S.page='incoming';S.tabI='work';S.inKpi='';render();})])]);
      mainKids.push(h('div', { class:'panels' }, [p1, p2]));
      const att = S.protocols.filter(r => /просрочено/i.test(r.executionStatusNormalized) || (!/снят с контроля/i.test(r.executionStatusNormalized) && r.daysDelta != null && r.daysDelta <= 7)).slice().sort(function(a,b){
        const rank = function(r){
          if (/просрочено/i.test(r.executionStatusNormalized)) return 0;
          if (r.daysDelta != null && r.daysDelta < 0) return 1;
          if (r.daysDelta === 0) return 2;
          if (r.daysDelta >= 1 && r.daysDelta <= 3) return 3;
          if (r.daysDelta >= 4 && r.daysDelta <= 7) return 4;
          return 5;
        };
        return rank(a) - rank(b) || (a.daysDelta ?? 99) - (b.daysDelta ?? 99);
      }).slice(0, 20);
      const attView = applySort(att, 'p', PGET);
      mainKids.push(h('h3', {}, 'Что требует внимания'));
      mainKids.push(table(['№ Протокола','Дата','№ поручения','Содержание поручений','Структурное подразделение','Срок исполнения','Информация о ходе исполнения','Статус исполнения','Следующее действие'], attView.map(r => [r.protocolNumberRaw, fmt(r.protocolDate), r.assignmentNumberRaw, r.assignmentText, compactUnits(r.resolvedUnits)||r.sourceStructuralUnitRaw, sourceDeadline(r), r.progressPreview, r.executionStatusNormalized, r.nextAction]), attView, 'p'));
    }

    if (S.page === 'protocols') {
      mainKids.push(h('h1', {}, 'Контроль протокольных поручений'));
      
      const assignedOpts = [...new Set(protoBase({assigned:true}).flatMap(r => r.sourceStructuralTokens || []))].sort();
      const jointOpts = [...new Set(protoBase({joint:true}).flatMap(r => (r.organization.jointGroups||[]).concat(r.organization.isJointExecution?['да']:[])))].filter(Boolean).sort();
      const spNames = [...new Set(protoBase({spLeaders:true}).flatMap(r => r.organization.spLeaders||[]))].sort();
      const norNames = [...new Set(protoBase({norDirectors:true}).flatMap(r => r.organization.norDirectors||[]))].sort();
      const headNames = [...new Set(protoBase({heads:true}).flatMap(r => r.organization.departmentHeads||[]))].sort();
      const curCasc = [...new Set(protoBase({curators:true}).flatMap(r => r.organization.curators||[]))].sort();
      const unitCasc = [...new Set(protoBase({units:true}).flatMap(r => r.resolvedUnits||[]))].filter(Boolean).sort();
      const statuses = [...new Set(S.protocols.map(r=>r.executionStatusNormalized).filter(Boolean))];
      mainKids.push(h('div', { class:'filters' }, [
        periodSel(F.p), deadlineSel(F.p, [...new Set(S.protocols.filter(r=>r.deadlineType==='TEXT').map(r=>r.deadlineText).filter(Boolean))]), ms('p-st','Статус', statuses, F.p.statuses), ms('p-u','Структурное подразделение', unitCasc, F.p.units),
        h('button', { class:'btn ghost', type:'button', onClick:()=>{F.p.showAll=!F.p.showAll;render();} }, F.p.showAll?'Скрыть фильтры':'Все фильтры ⚙'), inputQ(F.p,'pq')
      ]));
      if (F.p.showAll) mainKids.push(h('div', { class:'adv' }, [
        ms('p-as','Назначено СП', assignedOpts, F.p.assigned),
        ms('p-cu','Руководитель-куратор', curCasc, F.p.curators),
        ms('p-hd','Руководитель подразделения', headNames, F.p.heads, n=>leaderLabel(n,'head')),
        ms('p-sp','Руководители СП', spNames, F.p.spLeaders, n=>leaderLabel(n,'sp')),
        ms('p-nr','Директора НОР', norNames, F.p.norDirectors, n=>leaderLabel(n,'nor')),
        ms('p-jo','Совместное исполнение', jointOpts, F.p.joint),
        (function(){ const i=h('input',{type:'date'}); i.value=F.p.dateFrom; i.onchange=e=>{F.p.dateFrom=e.target.value;F.p.period='custom';render();}; i.title='Дата с'; return i; })(),
        (function(){ const i=h('input',{type:'date'}); i.value=F.p.dateTo; i.onchange=e=>{F.p.dateTo=e.target.value;F.p.period='custom';render();}; i.title='Дата по'; return i; })(),
        (function(){ const i=h('input',{type:'month'}); i.value=F.p.dateMonth; i.onchange=e=>{F.p.dateMonth=e.target.value;F.p.period='custom';render();}; return i; })(),
        (function(){ const i=h('input',{placeholder:'Год документа'}); i.value=F.p.dateYear; i.setAttribute('data-fid','p-year'); i.oninput=e=>{F.p.dateYear=e.target.value;F.p.period='custom'; window.__fid='p-year'; window.__fpos=e.target.selectionStart; clearTimeout(window.__ft); window.__ft=setTimeout(render,280);}; return i; })(),
        (function(){ const i=h('input',{type:'date'}); i.value=F.p.dlFrom; i.title='Срок с'; i.onchange=e=>{F.p.dlFrom=e.target.value;F.p.dlPreset='';render();}; return i; })(),
        (function(){ const i=h('input',{type:'date'}); i.value=F.p.dlTo; i.title='Срок по'; i.onchange=e=>{F.p.dlTo=e.target.value;F.p.dlPreset='';render();}; return i; })(),
        (function(){ const i=h('input',{type:'month'}); i.value=F.p.dlMonth; i.title='Срок месяц'; i.onchange=e=>{F.p.dlMonth=e.target.value;F.p.dlPreset='';render();}; return i; })(),
        (function(){ const i=h('input',{placeholder:'Год срока'}); i.value=F.p.dlYear; i.setAttribute('data-fid','p-dly'); i.oninput=e=>{F.p.dlYear=e.target.value;F.p.dlPreset=''; window.__fid='p-dly'; window.__fpos=e.target.selectionStart; clearTimeout(window.__ft); window.__ft=setTimeout(render,280);}; return i; })()
      ]));
      mainKids.push(chipBar(F.p,'p','Найдено: '+pv.length));

      mainKids.push(h('div', { class:'kpis' }, [kpi('green','ВСЕГО',pAll,'',S.protoKpi==='all',()=>{S.protoKpi='all';render();}), kpi('red','ПРОСРОЧЕНО',pOd,'',S.protoKpi==='overdue',()=>{S.protoKpi='overdue';render();}), kpi('amber','В РАБОТЕ',pWork,'',S.protoKpi==='work',()=>{S.protoKpi='work';render();})]));
      mainKids.push(h('div', { class:'kpis' }, [kpi('gold sec','РАБОЧИЙ КОНТРОЛЬ',pWatch,'',S.protoKpi==='watch',()=>{S.protoKpi='watch';render();}), kpi('orange sec','НА СНЯТИИ С КОНТРОЛЯ',pRem,'',S.protoKpi==='removing',()=>{S.protoKpi='removing';render();}), kpi('lime sec','СНЯТ С КОНТРОЛЯ',pDone,'',S.protoKpi==='done',()=>{S.protoKpi='done';render();})]));
      const exp = h('button', { class:'btn ghost' }, 'Экспорт');
      exp.onclick = () => exportX(pv.map(r => ({ '№ Протокола':r.protocolNumberRaw, Дата:fmt(r.protocolDate), '№ поручения':r.assignmentNumberRaw, 'Содержание поручений':r.assignmentText, 'Структурное подразделение':(r.resolvedUnits||[]).join('; '), 'Срок исполнения':r.deadlineDisplay, 'Информация о ходе исполнения':r.progressInfo, 'Статус исполнения':r.executionStatusNormalized, 'Следующее действие':r.nextAction })), 'protocols.xlsx');
      mainKids.push(h('div', { class:'toolbar' }, [h('h3', {}, 'Реестр поручений — ' + pv.length), exp]));
      mainKids.push(table(['№ Протокола','Дата','№ поручения','Содержание поручений','Структурное подразделение','Срок исполнения','Отклонение / осталось','Информация о ходе исполнения','Статус исполнения','Следующее действие'], pv.map(r => [r.protocolNumberRaw, fmt(r.protocolDate), r.assignmentNumberRaw, r.assignmentText, compactUnits(r.resolvedUnits)||r.sourceStructuralUnitRaw, sourceDeadline(r), r.deadlineDisplay, r.progressPreview, r.executionStatusNormalized, r.nextAction]), pv, 'p'));
    }
    if (S.page === 'appeals') {
      mainKids.push(h('h1', {}, 'Контроль обращений e-Өтініш'));
      
      const types=[...new Set(S.appeals.map(r=>r.type).filter(Boolean))];
      const ast=[...new Set(S.appeals.map(r=>r.executionStatusNormalized).filter(Boolean))];
      const aJoint=[...new Set(appealBase({joint:true}).flatMap(r=> (r.organization.jointGroups||[]).concat(r.organization.isJointExecution?['да']:[])))].filter(Boolean);
      const aSp=[...new Set(appealBase({spLeaders:true}).flatMap(r=>r.organization.spLeaders||[]))].sort();
      const aNor=[...new Set(appealBase({norDirectors:true}).flatMap(r=>r.organization.norDirectors||[]))].sort();
      const aHead=[...new Set(appealBase({heads:true}).flatMap(r=>r.organization.departmentHeads||[]))].sort();
      const aUnits=[...new Set(appealBase({units:true}).map(r=>r.resolvedUnit).filter(Boolean))].sort();
      const aCur=[...new Set(appealBase({curators:true}).flatMap(r=>r.organization.curators||[]))].sort();
      mainKids.push(h('div', { class:'filters' }, [
        periodSel(F.a), deadlineSel(F.a, [...new Set(S.appeals.filter(r=>r.deadlineType==='TEXT').map(r=>r.deadlineText).filter(Boolean))]), ms('a-st','Статус', ast, F.a.statuses), ms('a-tp','Вид обращения', types, F.a.types),
        h('button', { class:'btn ghost', type:'button', onClick:()=>{F.a.showAll=!F.a.showAll;render();} }, F.a.showAll?'Скрыть фильтры':'Все фильтры ⚙'), inputQ(F.a,'aq')
      ]));
      if (F.a.showAll) mainKids.push(h('div', { class:'adv' }, [
        textField('a-num','Номер обращения', F.a.number, v=>F.a.number=v),
        textField('a-au','Автор обращения', F.a.author, v=>F.a.author=v),
        textField('a-sm','Краткое содержание', F.a.summary, v=>F.a.summary=v),
        ms('a-u','Структурное подразделение', aUnits, F.a.units),
        ms('a-cu','Руководитель-куратор', aCur, F.a.curators),
        ms('a-hd','Руководитель подразделения', aHead, F.a.heads, n=>leaderLabel(n,'head')),
        ms('a-sp','Руководители СП', aSp, F.a.spLeaders, n=>leaderLabel(n,'sp')),
        ms('a-nr','Директора НОР', aNor, F.a.norDirectors, n=>leaderLabel(n,'nor')),
        ms('a-jo','Совместное исполнение', aJoint.concat(['да']), F.a.joint),
        (function(){ const i=h('input',{type:'date'}); i.value=F.a.dateFrom; i.onchange=e=>{F.a.dateFrom=e.target.value;F.a.period='custom';render();}; return i; })(),
        (function(){ const i=h('input',{type:'date'}); i.value=F.a.dateTo; i.onchange=e=>{F.a.dateTo=e.target.value;F.a.period='custom';render();}; return i; })(),
        (function(){ const i=h('input',{type:'month'}); i.value=F.a.dateMonth; i.onchange=e=>{F.a.dateMonth=e.target.value;F.a.period='custom';render();}; return i; })(),
        (function(){ const i=h('input',{placeholder:'Год'}); i.value=F.a.dateYear; i.oninput=e=>{F.a.dateYear=e.target.value;F.a.period='custom';render();}; return i; })(),
        (function(){ const i=h('input',{type:'date'}); i.value=F.a.dlFrom; i.title='Срок с'; i.onchange=e=>{F.a.dlFrom=e.target.value;F.a.dlPreset='';render();}; return i; })(),
        (function(){ const i=h('input',{type:'date'}); i.value=F.a.dlTo; i.title='Срок по'; i.onchange=e=>{F.a.dlTo=e.target.value;F.a.dlPreset='';render();}; return i; })(),
        (function(){ const i=h('input',{type:'month'}); i.value=F.a.dlMonth; i.title='Срок месяц'; i.onchange=e=>{F.a.dlMonth=e.target.value;F.a.dlPreset='';render();}; return i; })(),
        textField('a-dly','Год срока', F.a.dlYear, v=>{F.a.dlYear=v;F.a.dlPreset='';}),
        (function(){ const i=h('input',{type:'date'}); i.value=F.a.respFrom; i.onchange=e=>{F.a.respFrom=e.target.value;render();}; i.title='Дата ответа с'; return i; })(),
        (function(){ const i=h('input',{type:'date'}); i.value=F.a.respTo; i.onchange=e=>{F.a.respTo=e.target.value;render();}; i.title='Дата ответа по'; return i; })(),
        (function(){ const i=h('input',{type:'month'}); i.value=F.a.respMonth; i.title='Ответ месяц'; i.onchange=e=>{F.a.respMonth=e.target.value;render();}; return i; })(),
        textField('a-ry','Год ответа', F.a.respYear, v=>{F.a.respYear=v;F.a.respMissing=false;}),
        (function(){ const l=h('label',{class:'check-filter'}); const c=h('input',{type:'checkbox'}); c.checked=F.a.respMissing; c.onchange=e=>{F.a.respMissing=e.target.checked;if(F.a.respMissing){F.a.respFrom='';F.a.respTo='';F.a.respMonth='';F.a.respYear='';}render();}; l.appendChild(c); l.appendChild(document.createTextNode(' Ответ не предоставлен')); return l; })()
      ]));
      mainKids.push(chipBar(F.a,'a','Найдено: '+av.length));

      mainKids.push(h('div', { class:'kpis' }, [kpi('red','ПРОСРОЧЕНО',aOd,'',S.appealKpi==='overdue',()=>{S.appealKpi='overdue';S.tabA='overdue';render();}), kpi('amber','СРОК ДО 7 ДНЕЙ',aSoon,'',S.appealKpi==='soon',()=>{S.appealKpi='soon';S.tabA='all';render();}), kpi('sky','ЗАВЕРШЕНО',aDone,'',S.appealKpi==='done',()=>{S.appealKpi='done';S.tabA='done';render();})]));
      mainKids.push(tabs([['all','Все '+aAll],['work','На исполнении'],['overdue','Просроченные'],['done','Завершённые']], S.tabA, v => { S.tabA = v; S.appealKpi = ''; render(); }));
      const exp = h('button', { class:'btn ghost' }, 'Экспорт');
      exp.onclick = () => exportX(av.map(r => ({ 'Номер обращения':r.number, 'Дата регистрации':fmt(r.registrationDate), 'Автор обращения':r.author, 'Вид обращения':r.type, 'Краткое содержание':r.summary, 'Ответственный исполнитель':r.responsibleFio, 'Структурное подразделение':r.resolvedUnit, 'Срок исполнения':r.deadlineDisplay, 'Дата предоставления ответа':fmt(r.responseDate), 'Статус исполнения':r.executionStatusNormalized, 'Следующее действие':r.nextAction })), 'eotinish.xlsx');
      mainKids.push(h('div', { class:'toolbar' }, [h('h3', {}, 'Обращения — ' + av.length), exp]));
      mainKids.push(table(['Номер обращения','Дата регистрации','Автор обращения','Вид обращения','Краткое содержание','Ответственный исполнитель','Структурное подразделение','Срок исполнения','Отклонение / осталось','Дата предоставления ответа','Статус исполнения','Следующее действие'], av.map(r => [r.number, fmt(r.registrationDate), r.author, r.type, r.summary, r.responsibleFio, r.resolvedUnit, sourceDeadline(r), r.deadlineDisplay, fmt(r.responseDate), r.executionStatusNormalized, r.nextAction]), av, 'a'));
    }
    if (S.page === 'incoming') {
      mainKids.push(h('h1', {}, 'Контроль входящих документов'));
      
      const ist=[...new Set(S.incoming.map(r=>r.executionStatusNormalized).filter(Boolean))];
      const iJoint=[...new Set(inBase({joint:true}).flatMap(r=>(r.organization.jointGroups||[]).concat(r.organization.isJointExecution?['да']:[])))].filter(Boolean);
      const iSp=[...new Set(inBase({spLeaders:true}).flatMap(r=>r.organization.spLeaders||[]))].sort();
      const iNor=[...new Set(inBase({norDirectors:true}).flatMap(r=>r.organization.norDirectors||[]))].sort();
      const iHead=[...new Set(inBase({heads:true}).flatMap(r=>r.organization.departmentHeads||[]))].sort();
      const iPos=[...new Set(inBase({positions:true}).flatMap(r=>r.organization.positions||[]))].filter(Boolean).sort();
      const iUnits=[...new Set(inBase({units:true}).map(r=>r.resolvedUnit).filter(Boolean))].sort();
      const iCur=[...new Set(inBase({curators:true}).flatMap(r=>r.organization.curators||[]))].sort();
      mainKids.push(h('div', { class:'filters' }, [
        periodSel(F.i), deadlineSel(F.i, [...new Set(S.incoming.filter(r=>r.deadlineType==='TEXT').map(r=>r.deadlineText).filter(Boolean))]), ms('i-st','Статус', ist, F.i.statuses), ms('i-u','Структурное подразделение', iUnits, F.i.units),
        h('button', { class:'btn ghost', type:'button', onClick:()=>{F.i.showAll=!F.i.showAll;render();} }, F.i.showAll?'Скрыть фильтры':'Все фильтры ⚙'), inputQ(F.i,'iq')
      ]));
      if (F.i.showAll) mainKids.push(h('div', { class:'adv' }, [
        textField('i-num','Рег. номер и дата', F.i.number, v=>F.i.number=v),
        textField('i-sm','Краткое содержание', F.i.summary, v=>F.i.summary=v),
        ms('i-cu','Руководитель-куратор', iCur, F.i.curators),
        ms('i-hd','Руководитель подразделения', iHead, F.i.heads, n=>leaderLabel(n,'head')),
        ms('i-sp','Руководители СП', iSp, F.i.spLeaders, n=>leaderLabel(n,'sp')),
        ms('i-nr','Директора НОР', iNor, F.i.norDirectors, n=>leaderLabel(n,'nor')),
        ms('i-jo','Совместное исполнение', iJoint.concat(['да']), F.i.joint),
        ms('i-po','Должность', iPos, F.i.positions),
        (function(){ const i=h('input',{type:'date'}); i.value=F.i.dateFrom; i.title='Дата входящего с'; i.onchange=e=>{F.i.dateFrom=e.target.value;F.i.period='custom';render();}; return i; })(),
        (function(){ const i=h('input',{type:'date'}); i.value=F.i.dateTo; i.title='Дата входящего по'; i.onchange=e=>{F.i.dateTo=e.target.value;F.i.period='custom';render();}; return i; })(),
        (function(){ const i=h('input',{type:'month'}); i.value=F.i.dateMonth; i.title='Месяц входящего'; i.onchange=e=>{F.i.dateMonth=e.target.value;F.i.period='custom';render();}; return i; })(),
        textField('i-year','Год входящего', F.i.dateYear, v=>{F.i.dateYear=v;F.i.period='custom';}),
        (function(){ const i=h('input',{type:'date'}); i.value=F.i.dlFrom; i.title='Срок с'; i.onchange=e=>{F.i.dlFrom=e.target.value;F.i.dlPreset='';render();}; return i; })(),
        (function(){ const i=h('input',{type:'date'}); i.value=F.i.dlTo; i.title='Срок по'; i.onchange=e=>{F.i.dlTo=e.target.value;F.i.dlPreset='';render();}; return i; })(),
        (function(){ const i=h('input',{type:'month'}); i.value=F.i.dlMonth; i.title='Срок месяц'; i.onchange=e=>{F.i.dlMonth=e.target.value;F.i.dlPreset='';render();}; return i; })(),
        textField('i-dly','Год срока', F.i.dlYear, v=>{F.i.dlYear=v;F.i.dlPreset='';})
      ]));
      mainKids.push(chipBar(F.i,'i','Найдено: '+iv.length));

      mainKids.push(h('div', { class:'kpis' }, [kpi('red','ПРОСРОЧЕНО',iOd,'',S.inKpi==='overdue',()=>{S.inKpi='overdue';S.tabI='overdue';render();}), kpi('amber','СРОК ДО 7 ДНЕЙ',iSoon,'',S.inKpi==='soon',()=>{S.inKpi='soon';S.tabI='all';render();}), kpi('sky','ВЫПОЛНЕНО',iDone,'',S.inKpi==='done',()=>{S.inKpi='done';S.tabI='done';render();})]));
      mainKids.push(tabs([['all','Все '+iAll],['work','На исполнении'],['overdue','Просроченные'],['done','Выполненные']], S.tabI, v => { S.tabI = v; S.inKpi = ''; render(); }));
      const exp = h('button', { class:'btn ghost' }, 'Экспорт');
      exp.onclick = () => exportX(iv.map(r => ({ 'Рег. номер и дата':r.regNumberDate, 'Краткое содержание':r.summary, 'Структурное подразделение':r.resolvedUnit, 'Срок исполнения':r.deadlineDisplay, 'Статус исполнения':r.executionStatusNormalized, 'Следующее действие':r.nextAction })), 'incoming.xlsx');
      mainKids.push(h('div', { class:'toolbar' }, [h('h3', {}, 'Входящие документы — ' + iv.length), exp]));
      mainKids.push(table(['Рег. номер и дата входящего документа','Краткое содержание','Структурное подразделение','Срок исполнения','Отклонение / осталось','Статус исполнения','Следующее действие'], iv.map(r => [r.regNumberDate, r.summary, r.resolvedUnit, sourceDeadline(r), r.deadlineDisplay, r.executionStatusNormalized, r.nextAction]), iv, 'i'));
    }
    if (S.page === 'upload') {
      mainKids.push(h('h1', {}, 'Загрузка данных'), h('p', { class:'meta' }, 'Загрузите только три операционных файла. Организационная структура подключается автоматически.'));
      mainKids.push(h('div', { class:'cards3' }, [uploadCard('1.xlsx — Протокольные поручения', S.protoMeta, 1), uploadCard('2.xlsx — e-Өтініш', S.appealMeta, 2), uploadCard('3.xlsx — Входящие документы', S.inMeta, 3)]));
      mainKids.push(h('div', { class:'sys' }, [h('h3', {}, 'Организационная структура'), h('div', { class:'meta' }, 'Файл: structure.xlsx · статус: ' + (S.structure ? 'загружена' : (S.structureError || 'нет')) + ' · записей: ' + (S.structure ? S.structure.people.length : 0) + ' · руководителей СП: ' + (S.structure ? Object.keys(S.structure.spLeaderByUnit).length : 0) + ' · директоров НОР: ' + (S.structure ? Object.keys(S.structure.norDirectorByUnit).length : 0) + ' · логотип: загружен')]));
    }

    if (window.__fid) {
      setTimeout(function(){
        const el = document.querySelector('[data-fid="'+window.__fid+'"]');
        if (el) { el.focus(); try { el.setSelectionRange(window.__fpos||el.value.length, window.__fpos||el.value.length); } catch(e){} }
      }, 0);
    }
    root.appendChild(h('div', { class:'app' }, [h('aside', { class:'sidebar' }, [h('div', { class:'brand' }, [h('img', { src:LOGO, alt:'ВЖДО' }), h('span', {}, 'ТОО «ВЖДО»')]), nav('overview','⌂','Обзор'), nav('protocols','▤','Протоколы'), nav('appeals','▥','e-Өтініш'), nav('incoming','◫','Входящие'), h('div', { class:'nav-spacer' }), nav('upload','⬆','Загрузка')]), h('main', { class:'main' }, mainKids)]));

    if (S.detail) {
      const d = S.detail, kv = [], add = (k, v) => { kv.push(h('b', {}, k)); kv.push(h('span', {}, v == null || v === '' ? '—' : String(v))); };
      if (d.kind === 'p') { add('№ Протокола', d.protocolNumberRaw); add('Дата', fmt(d.protocolDate)); add('№ поручения', d.assignmentNumberRaw); add('Содержание поручения', d.assignmentText); add('Структурное подразделение — источник', d.sourceStructuralUnitRaw); add('Структурные подразделения — resolved', (d.resolvedUnits||[]).join(', ')); add('Срок исполнения', d.deadlineDisplay); add('Информация о ходе исполнения', d.progressInfo); add('Статус исполнения', d.executionStatusNormalized); add('Следующее действие', d.nextAction); add('Руководитель-куратор', (d.organization.curators||[]).join(', ')); add('Совместное исполнение', (d.organization.jointGroups||[]).join(', ') || (d.organization.isJointExecution?'Да':'Нет')); add('Руководитель подразделения', (d.organization.departmentHeads||[]).join(', ')); add('Руководители СП', (d.organization.spLeaders||[]).join(', ')); add('Директора НОР', (d.organization.norDirectors||[]).join(', ')); add('Исходная строка', d.sourceRowNumber); }
      else if (d.kind === 'a') { add('Номер обращения', d.number); add('Дата регистрации обращения', fmt(d.registrationDate)); add('Автор обращения', d.author); add('Вид обращения', d.type); add('Краткое содержание', d.summary); add('Ответственный исполнитель', d.responsibleFio); add('Структурное подразделение', d.resolvedUnit); add('Срок исполнения', d.deadlineDisplay); add('Дата предоставления ответа', fmt(d.responseDate)); add('Статус исполнения', d.executionStatusNormalized); add('Следующее действие', d.nextAction); add('Руководитель-куратор', (d.organization.curators||[]).join(', ')); add('Совместное исполнение', (d.organization.jointGroups||[]).join(', ') || (d.organization.isJointExecution?'Да':'Нет')); add('Руководитель подразделения', (d.organization.departmentHeads||[]).join(', ')); add('Руководитель СП', (d.organization.spLeaders||[]).join(', ')); add('Директор НОР', (d.organization.norDirectors||[]).join(', ')); add('Исходная строка', d.sourceRowNumber); add('Исходное подразделение из выгрузки', d.sourceDepartment); }
      else { add('Рег. номер и дата входящего документа', d.regNumberDate); add('Краткое содержание', d.summary); add('Структурное подразделение', d.resolvedUnit); add('Должность', (d.organization.positions||[]).join(', ')); add('Срок исполнения', d.deadlineDisplay); add('Статус исполнения', d.executionStatusNormalized); add('Следующее действие', d.nextAction); add('Руководитель-куратор', (d.organization.curators||[]).join(', ')); add('Совместное исполнение', (d.organization.jointGroups||[]).join(', ') || (d.organization.isJointExecution?'Да':'Нет')); add('Руководитель подразделения', (d.organization.departmentHeads||[]).join(', ')); add('Руководитель СП', (d.organization.spLeaders||[]).join(', ')); add('Директор НОР', (d.organization.norDirectors||[]).join(', ')); add('Исходная строка', d.sourceRowNumber); add('Исходное подразделение из выгрузки', d.sourceDepartment); }
      root.appendChild(h('div', { class:'drawer-back', onClick:() => { S.detail = null; render(); } }));
      root.appendChild(h('aside', { class:'drawer' }, [h('button', { class:'btn ghost', onClick:() => { S.detail = null; render(); } }, 'Закрыть'), h('h2', {}, d.kind === 'p' ? 'Поручение' : d.kind === 'a' ? 'Обращение' : 'Входящий документ'), h('div', { class:'kv' }, kv)].concat((d.warnings||[]).map(w => h('div', { class:'warn' }, typeof w === 'string' ? w : w.message)))));
    }
  }

  async function boot() {
    await restore();
    try {
      const res = await fetch(STRUCT, { cache:'no-store' });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      S.structure = parseStructure(XLSX.read(await res.arrayBuffer(), { type:'array', cellDates:true }));
      S.structureError = null;
    } catch (e) { S.structureError = e.message || String(e); }
    reenrich(); render();
    let lastDay = refDate().getTime();
    setInterval(() => { const nowDay = refDate().getTime(); if (nowDay !== lastDay) { lastDay = nowDay; reenrich(); render(); } }, 60000);
  }
  boot();
})();
