/* global XLSX */
(() => {
  const $ = (id) => document.getElementById(id);
  const statusEl = $('status');

  const fileInputs = {
    master: $('fileMaster'),
    plan: $('filePlan'),
    works: $('fileWorks'),
    attendance: $('fileAttendance'),
  };

  Object.values(fileInputs).forEach((input) => {
    input.addEventListener('change', () => {
      const box = input.closest('.upload');
      const nameEl = box?.querySelector('.file-name');
      if (nameEl) nameEl.textContent = input.files?.[0]?.name || '파일 선택 전';
      box?.classList.toggle('selected', !!input.files?.[0]);
    });
  });

  $('runBtn').addEventListener('click', runAnalysis);
  $('resetBtn').addEventListener('click', () => location.reload());

  function setStatus(message, type = '') {
    statusEl.className = `status ${type}`.trim();
    statusEl.innerHTML = message;
  }

  function clean(value) {
    if (value === null || value === undefined) return '';
    return String(value).replace(/\u00a0/g, '').replace(/\r?\n/g, ' ').trim();
  }

  function normalizeName(value) {
    return clean(value).replace(/\s+/g, '');
  }

  function normalizeGroup(value) {
    const s = clean(value).toUpperCase();
    if (s.includes('MX')) return 'MX';
    if (s.includes('CE')) return 'CE';
    return s;
  }

  function normalizeStore(value) {
    let s = clean(value);
    if (!s) return '';
    if (s.startsWith('=')) return '';
    s = s
      .replace(/코스트코/g, '')
      .replace(/\s+/g, '')
      .replace(/[()]/g, '')
      .trim();
    if (s === '혁신점' || s === '대구혁신') s = '대구혁신점';
    if (s && !s.endsWith('점') && !s.includes('사무소') && !s.includes('사무실')) s += '점';
    return s;
  }

  function excelSerialToDate(serial) {
    const utcDays = Math.floor(Number(serial) - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date(utcValue * 1000);
    return new Date(dateInfo.getUTCFullYear(), dateInfo.getUTCMonth(), dateInfo.getUTCDate());
  }

  function dateKey(date) {
    if (!date || Number.isNaN(date.getTime())) return '';
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  function parseDate(value, fallbackYear, fallbackMonth) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    if (typeof value === 'number' && Number.isFinite(value)) return excelSerialToDate(value);
    const s = clean(value);
    if (!s) return null;
    if (/^\d+(?:\.\d+)?$/.test(s)) {
      const n = Number(s);
      if (n > 20000 && n < 70000) return excelSerialToDate(n);
    }
    let m = s.match(/^(\d{4})[-/.년\s]*(\d{1,2})[-/.월\s]*(\d{1,2})/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    m = s.match(/^(\d{1,2})\s*월\s*(\d{1,2})\s*일$/);
    if (m && fallbackYear) return new Date(fallbackYear, Number(m[1]) - 1, Number(m[2]));
    m = s.match(/^(\d{1,2})일$/);
    if (m && fallbackYear && fallbackMonth) return new Date(fallbackYear, fallbackMonth - 1, Number(m[1]));
    return null;
  }

  function parseTimeToMinutes(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value.getHours() * 60 + value.getMinutes();
    if (typeof value === 'number' && Number.isFinite(value)) {
      if (value >= 0 && value < 1) return Math.round(value * 24 * 60);
      return null;
    }
    const s = clean(value);
    if (!s) return null;
    if (/^\d+(?:\.\d+)?$/.test(s)) {
      const n = Number(s);
      if (n >= 0 && n < 1) return Math.round(n * 24 * 60);
    }
    let m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
    if (!m) m = s.match(/^(오전|오후)\s*(\d{1,2}):(\d{2})$/);
    if (m && (m[1] === '오전' || m[1] === '오후')) {
      let hh = Number(m[2]);
      const mm = Number(m[3]);
      if (m[1] === '오후' && hh < 12) hh += 12;
      if (m[1] === '오전' && hh === 12) hh = 0;
      return hh * 60 + mm;
    }
    if (!m) return null;
    const hh = Number(m[1]);
    const mm = Number(m[2]);
    if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
    return hh * 60 + mm;
  }

  function minutesToHHMM(minutes) {
    if (minutes === null || minutes === undefined || Number.isNaN(minutes)) return '';
    const h = Math.floor(minutes / 60);
    const m = minutes % 60;
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
  }

  function isOffLike(value) {
    const s = clean(value);
    if (!s) return true;
    return /휴무|휴일|연차|휴가|공가|대체|보상|DIDA|예비군|병가|경조/.test(s);
  }

  function hasEducation(value) { return clean(value).includes('교육'); }
  function hasRotation(value) {
    // '부산순환', '대혁 순환', '순환근무'처럼 어디에 들어가도 순환으로 판단
    return clean(value).replace(/\s+/g, '').includes('순환');
  }

  function extractPlanShift(value) {
    const s = clean(value);
    if (/근무\s*A|근무A|A조/.test(s)) return 'A조';
    if (/근무\s*B|근무B|B조/.test(s)) return 'B조';
    if (/근무\s*C|근무C|C조/.test(s)) return 'C조';
    return '';
  }

  function extractWorksShift(value) {
    const s = clean(value);
    if (/A조/.test(s)) return 'A조';
    if (/B조/.test(s)) return 'B조';
    if (/C조/.test(s)) return 'C조';
    return '';
  }

  function findHeaderRow(rows, requiredWords, maxRows = 10) {
    for (let r = 0; r < Math.min(rows.length, maxRows); r++) {
      const cells = (rows[r] || []).map(clean);
      const hit = requiredWords.every(w => cells.some(c => c.includes(w)));
      if (hit) return r;
    }
    return 0;
  }

  function findCol(row, word, fallback = -1) {
    const idx = (row || []).findIndex(v => clean(v).includes(word));
    return idx >= 0 ? idx : fallback;
  }

  async function readWorkbook(file) {
    const data = await file.arrayBuffer();
    return XLSX.read(data, { type: 'array', cellDates: true, raw: false, cellFormula: false, dateNF: 'yyyy-mm-dd' });
  }

  function sheetRows(wb, preferredNames = []) {
    let sheetName = wb.SheetNames[0];
    for (const name of preferredNames) {
      if (wb.SheetNames.includes(name)) { sheetName = name; break; }
    }
    const ws = wb.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true, blankrows: false });
  }

  async function runAnalysis() {
    try {
      const missingFiles = Object.entries(fileInputs).filter(([, el]) => !el.files || !el.files[0]).map(([key]) => key);
      if (missingFiles.length) { setStatus('엑셀 파일 4개를 모두 업로드해야 합니다.', 'error'); return; }
      setStatus('엑셀 파일을 읽고 있습니다...');
      const [masterWb, planWb, worksWb, attWb] = await Promise.all([
        readWorkbook(fileInputs.master.files[0]),
        readWorkbook(fileInputs.plan.files[0]),
        readWorkbook(fileInputs.works.files[0]),
        readWorkbook(fileInputs.attendance.files[0]),
      ]);
      const peopleRows = sheetRows(masterWb, ['인력DB']);
      const hourRows = sheetRows(masterWb, ['영업시간DB']);
      const planRows = sheetRows(planWb);
      const worksRows = sheetRows(worksWb);
      const attendanceRows = sheetRows(attWb);
      const result = analyze({ peopleRows, hourRows, planRows, worksRows, attendanceRows });
      makeWorkbook(result);
      setStatus(`분석 완료(v9): MX 지각 ${result.lateRows.length}건, 근태 미입력 ${result.noAttendanceRows.length}건, 퇴근 미입력 ${result.noCheckoutRows.length}건, 스케줄 불일치 ${result.mismatchRows.length}건.<br>결과 엑셀이 다운로드됩니다.`, 'ok');
    } catch (err) {
      console.error(err);
      setStatus(`오류가 발생했습니다.<br><b>${escapeHtml(err.message || err)}</b><br>파일 양식이나 시트명이 바뀌었는지 확인하세요.`, 'error');
    }
  }

  function inferYearMonth(attendanceRows, worksRows) {
    const manualMonth = clean($('monthInput').value);
    if (manualMonth) { const [y, m] = manualMonth.split('-').map(Number); return { year: y, month: m }; }

    // 1순위: 웍스스케줄 날짜 헤더. 근태관리에는 전월 말 데이터가 섞일 수 있어서 근태관리 첫 날짜로 월을 잡으면 안 됨.
    let yearFromAttendance = null;
    const attHeader = findHeaderRow(attendanceRows, ['이름', '근무일자']);
    const dateCol = findCol(attendanceRows[attHeader], '근무일자', 1);
    for (let r = attHeader + 1; r < attendanceRows.length; r++) {
      const d = parseDate(attendanceRows[r][dateCol]);
      if (d) { yearFromAttendance = d.getFullYear(); break; }
    }

    const headerRow = findHeaderRow(worksRows, ['성명'], 20);
    const header = worksRows[headerRow] || [];
    for (let c = 0; c < header.length; c++) {
      const d = parseDate(header[c], yearFromAttendance || new Date().getFullYear());
      if (d) return { year: d.getFullYear(), month: d.getMonth() + 1 };
    }

    // 2순위: 웍스스케줄 A1 등에 적힌 월 숫자
    const firstCellMonth = Number(clean(worksRows?.[0]?.[0]));
    if (firstCellMonth >= 1 && firstCellMonth <= 12) return { year: yearFromAttendance || new Date().getFullYear(), month: firstCellMonth };

    // 3순위: 근태관리 날짜
    for (let r = attHeader + 1; r < attendanceRows.length; r++) {
      const d = parseDate(attendanceRows[r][dateCol]);
      if (d) return { year: d.getFullYear(), month: d.getMonth() + 1 };
    }
    const now = new Date();
    return { year: now.getFullYear(), month: now.getMonth() + 1 };
  }

  function analyze({ peopleRows, hourRows, planRows, worksRows, attendanceRows }) {
    const { year, month } = inferYearMonth(attendanceRows, worksRows);
    const manualBaseDate = clean($('baseDateInput').value);

    const people = parsePeople(peopleRows);
    const peopleByName = new Map(people.map(p => [p.name, p]));
    const peopleByEmp = new Map(people.filter(p => p.empNo).map(p => [p.empNo, p]));
    const workHours = parseWorkHours(hourRows);
    const { planByNameDate, dayCols: planDayCols, planReadRows } = parsePlan(planRows, year, month);
    const { worksByNameDate, worksDateCols, worksReadRows } = parseWorks(worksRows, year, month);
    for (const rec of worksReadRows) {
      const plan = planByNameDate.get(`${rec.name}|${rec.date}`);
      rec.planValue = plan?.value || '';
      rec.planShift = plan?.shift || '';
      rec.isRotation = hasRotation(rec.value);
      rec.isEducation = hasEducation(rec.value);
      rec.appliedBasis = rec.isEducation ? '교육 제외' : rec.isRotation ? '순환→매장계획 기준' : isOffLike(rec.value) ? '휴무/제외' : '웍스 기준';
      rec.appliedShift = rec.isRotation ? rec.planShift : rec.shift;
    }
    const { attendanceByNameDate, latestAttendanceDate } = parseAttendance(attendanceRows, peopleByName, peopleByEmp);

    let baseDate = latestAttendanceDate;
    if (manualBaseDate) baseDate = parseDate(manualBaseDate);
    if (!baseDate) throw new Error('근태관리 파일에서 판정 기준일을 찾지 못했습니다.');
    const baseDateKey = dateKey(baseDate);

    const allPeople = people.slice().sort((a, b) => a.group.localeCompare(b.group, 'ko') || a.store.localeCompare(b.store, 'ko') || a.name.localeCompare(b.name, 'ko'));
    const lateRows = [];
    const noAttendanceRows = [];
    const noCheckoutRows = [];
    const mismatchRows = [];
    const ceRows = [];
    const exceptionRows = [];

    // MX 스케줄 불일치: 점수에는 영향 없음. 휴무/근무 차이도 확인용으로 표시.
    for (const person of allPeople) {
      if (person.group !== 'MX') continue;
      for (const dc of worksDateCols) {
        if (dc.dKey > baseDateKey) continue;
        const works = worksByNameDate.get(`${person.name}|${dc.dKey}`);
        const plan = planByNameDate.get(`${person.name}|${dc.dKey}`);
        const worksText = works?.value || '';
        const planText = plan?.value || '';
        if (hasEducation(worksText)) continue;
        const ps = extractPlanShift(planText);
        const ws = extractWorksShift(worksText);
        const planOff = isOffLike(planText);
        const worksOff = isOffLike(worksText);
        if (hasRotation(worksText)) {
          // 순환 문구가 들어가면 웍스 값과 비교하지 않고 매장근무계획의 근무A/B/C를 기준으로 사용한다.
          // 단, 매장근무계획에서 조를 찾지 못하면 지각 판정 기준이 없으므로 확인 대상에 남긴다.
          if (!ps && planText && !planOff) {
            mismatchRows.push({ name: person.name, date: dc.dKey, store: works?.store || person.store, plan: planText, works: worksText, note: '순환근무이나 매장근무계획에서 근무A/B/C 조 확인 필요' });
          }
          continue;
        }
        if ((ps || ws || planText || worksText) && (ps !== ws || planOff !== worksOff)) {
          mismatchRows.push({
            name: person.name, date: dc.dKey, store: works?.store || person.store,
            plan: planText, works: worksText,
            note: `매장근무계획상 ${planText || '공란'} / 웍스스케줄상 ${worksText || '공란'}`,
          });
        }
      }
    }

    // MX 지각/근태 미입력: 반드시 유효 근무조인 경우만 감점. 휴무/대체휴무/보상휴가/연차/교육은 제외.
    for (const person of allPeople) {
      if (person.group !== 'MX') continue;
      for (const dc of worksDateCols) {
        if (dc.dKey > baseDateKey) continue;
        const works = worksByNameDate.get(`${person.name}|${dc.dKey}`);
        const plan = planByNameDate.get(`${person.name}|${dc.dKey}`);
        const worksText = works?.value || '';
        const planText = plan?.value || '';
        const isRotation = hasRotation(worksText);
        if (!worksText || hasEducation(worksText)) continue;
        // 순환 문구가 들어간 날은 웍스 값이 어떤 표현이든 매장근무계획 근무A/B/C를 기준으로 처리한다.
        // 일반 근무일만 휴무/대체/보상/연차 등 제외값을 웍스 기준으로 제외한다.
        if (!isRotation && isOffLike(worksText)) continue;
        let basis = '웍스스케줄';
        let shift = extractWorksShift(worksText);
        let scheduleText = worksText;
        if (isRotation) {
          basis = '매장근무계획관리(순환)';
          if (hasEducation(planText) || isOffLike(planText)) continue;
          shift = extractPlanShift(planText);
          scheduleText = planText || worksText;
        }
        if (!shift) {
          if (isRotation) mismatchRows.push({ name: person.name, date: dc.dKey, store: works?.store || plan?.store || person.store, plan: planText, works: worksText, note: '순환근무이나 매장근무계획에서 근무A/B/C 조 확인 불가' });
          continue;
        }
        const rec = attendanceByNameDate.get(`${person.name}|${dc.dKey}`);
        const baseStore = works?.store || plan?.store || person.store;
        if (!rec || rec.firstIn === null) {
          const row = { group: 'MX', name: person.name, date: dc.dKey, store: baseStore, basis, schedule: scheduleText, reason: '근무 예정이나 출근기록 없음', score: -3 };
          noAttendanceRows.push(row); exceptionRows.push(toExceptionRow(row, '근태미입력'));
          continue;
        }
        const attendStore = rec.store || baseStore;
        let standard = workHours.get(attendStore)?.[shift];
        let standardStore = attendStore;
        if (standard === null || standard === undefined) { standard = workHours.get(baseStore)?.[shift]; standardStore = baseStore; }
        if (standard === null || standard === undefined) { standard = workHours.get(person.store)?.[shift]; standardStore = person.store; }
        if (standard === null || standard === undefined) {
          lateRows.push({ group: 'MX', name: person.name, date: dc.dKey, store: attendStore, standardStore: standardStore || '', shift, basis, schedule: scheduleText, standardTime: '기준없음', actualTime: minutesToHHMM(rec.firstIn), lateMinutes: '', lateType: '기준시간없음', rawScore: 0 });
          continue;
        }
        if (rec.firstIn >= standard) {
          const lateMinutes = rec.firstIn - standard;
          const lateType = lateMinutes <= 10 ? '10분 이내' : lateMinutes < 60 ? '11~59분' : '60분 이상';
          const rawScore = lateType === '60분 이상' ? -2 : -1;
          const row = { group: 'MX', name: person.name, date: dc.dKey, store: attendStore, standardStore, shift, basis, schedule: scheduleText, standardTime: minutesToHHMM(standard), actualTime: minutesToHHMM(rec.firstIn), lateMinutes, lateType, rawScore };
          lateRows.push(row); exceptionRows.push(toExceptionRow(row, '지각'));
        }
      }
    }

    // CE는 매장근무계획관리 기준 근무일의 출근 여부만 확인.
    for (const person of allPeople) {
      if (person.group !== 'CE') continue;
      for (const dc of planDayCols) {
        if (dc.dKey > baseDateKey) continue;
        const plan = planByNameDate.get(`${person.name}|${dc.dKey}`);
        const planText = plan?.value || '';
        if (!planText.includes('근무') || isOffLike(planText)) continue;
        const rec = attendanceByNameDate.get(`${person.name}|${dc.dKey}`);
        const ceRow = { group: 'CE', name: person.name, date: dc.dKey, store: person.store, schedule: planText, status: rec?.firstIn !== null && rec?.firstIn !== undefined ? '출근확인' : '근태미입력' };
        ceRows.push(ceRow);
        if (!rec || rec.firstIn === null) {
          const row = { group: 'CE', name: person.name, date: dc.dKey, store: person.store, basis: '매장근무계획관리', schedule: planText, reason: '근무 예정이나 출근기록 없음', score: -3 };
          noAttendanceRows.push(row); exceptionRows.push(toExceptionRow(row, '근태미입력'));
        }
      }
    }

    // 퇴근 미입력: 출근 기록이 있고 정상 퇴근시간이 없으면 표시. 단, 인력DB에 있는 사람만.
    for (const rec of attendanceByNameDate.values()) {
      if (rec.date > baseDateKey) continue;
      const person = peopleByName.get(rec.name);
      if (!person) continue;
      if (rec.firstIn !== null && !rec.hasOut) {
        const row = { group: person.group, name: person.name, date: rec.date, store: rec.store || person.store, checkIn: minutesToHHMM(rec.firstIn), checkOut: '', reason: '출근은 있으나 정상 퇴근시간 없음', score: 0 };
        noCheckoutRows.push(row); exceptionRows.push(toExceptionRow(row, '퇴근미입력'));
      }
    }

    const autoSummary = buildSummary(allPeople, lateRows, noAttendanceRows, noCheckoutRows);
    return { year, month, baseDate: baseDateKey, people: allPeople, lateRows, noAttendanceRows, noCheckoutRows, mismatchRows, ceRows, exceptionRows, autoSummary, worksReadRows, planReadRows };
  }

  function parsePeople(rows) {
    const h = findHeaderRow(rows, ['성명', '매장명'], 8);
    const header = rows[h] || [];
    const nameCol = findCol(header, '성명', 6);
    const storeCol = findCol(header, '매장명', 4);
    const empCol = findCol(header, '사번', 5);
    let groupCol = findCol(header, '구분', -1);
    if (groupCol < 0) {
      let best = 0, bestCount = -1;
      for (let c = 0; c < Math.min(10, header.length || 10); c++) {
        let cnt = 0;
        for (let r = h + 1; r < rows.length; r++) {
          const g = normalizeGroup(rows[r]?.[c]);
          if (g === 'MX' || g === 'CE') cnt++;
        }
        if (cnt > bestCount) { bestCount = cnt; best = c; }
      }
      groupCol = best;
    }
    const people = [];
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const name = normalizeName(row[nameCol]);
      const group = normalizeGroup(row[groupCol]);
      if (!name || !['MX', 'CE'].includes(group)) continue;
      people.push({ group, name, store: normalizeStore(row[storeCol]), empNo: clean(row[empCol]), manager: clean(row[1]), region: clean(row[2]) });
    }
    return people;
  }

  function parseWorkHours(rows) {
    const h = findHeaderRow(rows, ['점포명', 'A조'], 6);
    const header = rows[h] || [];
    const storeCol = findCol(header, '점포명', 0);
    const aCol = findCol(header, 'A조', 3), bCol = findCol(header, 'B조', 4), cCol = findCol(header, 'C조', 5);
    const map = new Map();
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const store = normalizeStore(row[storeCol]);
      if (!store) continue;
      map.set(store, { 'A조': parseTimeToMinutes(row[aCol]), 'B조': parseTimeToMinutes(row[bCol]), 'C조': parseTimeToMinutes(row[cCol]) });
    }
    return map;
  }

  function parsePlan(rows, year, month) {
    const h = findHeaderRow(rows, ['이름', '01일'], 10);
    const header = rows[h] || [];
    const nameCol = findCol(header, '이름', 3);
    const storeCol = findCol(header, '매장명', 0);
    const dayCols = [];
    for (let c = 0; c < header.length; c++) {
      const s = clean(header[c]);
      const m = s.match(/^(\d{1,2})일$/);
      if (m) {
        const d = new Date(year, month - 1, Number(m[1]));
        dayCols.push({ c, day: Number(m[1]), date: d, dKey: dateKey(d) });
      }
    }
    const planByNameDate = new Map();
    const planReadRows = [];
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const name = normalizeName(row[nameCol]);
      if (!name) continue;
      const store = normalizeStore(row[storeCol]);
      for (const dc of dayCols) {
        const value = clean(row[dc.c]);
        const shift = extractPlanShift(value);
        const rec = { name, date: dc.dKey, value, shift, store, sourceRow: r + 1, sourceCol: dc.c + 1, sourceCell: XLSX.utils.encode_cell({ r, c: dc.c }) };
        planByNameDate.set(`${name}|${dc.dKey}`, rec);
        if (value) planReadRows.push(rec);
      }
    }
    return { planByNameDate, dayCols, planReadRows };
  }

  function parseWorks(rows, year, month) {
    // 웍스스케줄은 행/열 위치가 자주 바뀌고, 숨김 행이 있을 수 있어서
    // "성명" 열을 찾은 뒤 그 오른쪽 최대 해당 월 일수만 날짜 컬럼으로 읽는다.
    const h = findHeaderRow(rows, ['성명'], 20);
    const header = rows[h] || [];
    const nameCol = findCol(header, '성명', 3);
    const storeColHeader = findCol(header, '근무처명', 2);
    const firstDateCol = nameCol + 1;
    const daysInMonth = new Date(year, month, 0).getDate();
    const dateCols = [];
    for (let c = firstDateCol; c < Math.min(header.length, firstDateCol + daysInMonth); c++) {
      let d = parseDate(header[c], year, month);
      if (!d) d = new Date(year, month - 1, c - firstDateCol + 1);
      if (d.getFullYear() === year && d.getMonth() + 1 === month) {
        dateCols.push({ c, date: d, dKey: dateKey(d), headerValue: clean(header[c]), cell: XLSX.utils.encode_cell({ r: h, c }) });
      }
    }
    const worksByNameDate = new Map();
    const worksReadRows = [];
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const name = normalizeName(row[nameCol]);
      if (!name || name === '필터용' || name.includes('합계')) continue;
      let store = normalizeStore(row[storeColHeader]);
      if (!store) store = normalizeStore(row[0]);
      for (const dc of dateCols) {
        const value = clean(row[dc.c]);
        const shift = extractWorksShift(value);
        const rec = { name, date: dc.dKey, value, shift, store, sourceRow: r + 1, sourceCol: dc.c + 1, sourceCell: XLSX.utils.encode_cell({ r, c: dc.c }), headerCell: dc.cell, headerValue: dc.headerValue };
        worksByNameDate.set(`${name}|${dc.dKey}`, rec);
        if (value) worksReadRows.push(rec);
      }
    }
    return { worksByNameDate, worksDateCols: dateCols, worksReadRows };
  }

  function parseAttendance(rows, peopleByName, peopleByEmp) {
    const h = findHeaderRow(rows, ['이름', '근무일자'], 10);
    const header = rows[h] || [];
    const nameCol = findCol(header, '이름', 0);
    const dateCol = findCol(header, '근무일자', 1);
    const inCol = header.findIndex(v => clean(v).includes('출근시간') && clean(v).includes('실제')) >= 0 ? header.findIndex(v => clean(v).includes('출근시간') && clean(v).includes('실제')) : findCol(header, '출근시간', 2);
    const empCol = findCol(header, '사번', 3);
    const outCol = header.findIndex(v => clean(v).includes('퇴근시간') && clean(v).includes('실제')) >= 0 ? header.findIndex(v => clean(v).includes('퇴근시간') && clean(v).includes('실제')) : findCol(header, '퇴근시간', 4);
    const inStoreCol = findCol(header, '출근지점', 7);
    const outStoreCol = findCol(header, '퇴근지점', 8);
    const map = new Map();
    let latest = null;
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const rawName = normalizeName(row[nameCol]);
      const empNo = clean(row[empCol]);
      const person = peopleByEmp.get(empNo) || peopleByName.get(rawName);
      const name = person?.name || rawName;
      const d = parseDate(row[dateCol]);
      if (!name || !d) continue;
      if (!latest || d > latest) latest = d;
      const dKey = dateKey(d);
      const inMin = parseTimeToMinutes(row[inCol]);
      const outMin = parseTimeToMinutes(row[outCol]);
      const inStore = normalizeStore(row[inStoreCol]);
      const outStore = normalizeStore(row[outStoreCol]);
      const key = `${name}|${dKey}`;
      if (!map.has(key)) map.set(key, { name, date: dKey, ins: [], outs: [], stores: [], outStores: [], rows: [] });
      const rec = map.get(key);
      if (inMin !== null) rec.ins.push(inMin);
      if (outMin !== null) rec.outs.push(outMin);
      if (inStore) rec.stores.push({ store: inStore, inMin: inMin ?? 99999 });
      if (outStore) rec.outStores.push(outStore);
      rec.rows.push(row);
    }
    for (const rec of map.values()) {
      rec.firstIn = rec.ins.length ? Math.min(...rec.ins) : null;
      rec.hasOut = rec.outs.length > 0;
      rec.outTime = rec.outs.length ? minutesToHHMM(Math.max(...rec.outs)) : '';
      const sortedStores = rec.stores.slice().sort((a, b) => a.inMin - b.inMin);
      rec.store = sortedStores[0]?.store || '';
    }
    return { attendanceByNameDate: map, latestAttendanceDate: latest };
  }

  function toExceptionRow(row, kind) {
    let judgement = '', lateType = '', lateMinutes = '', baseScore = '';
    if (kind === '지각') { judgement = `${row.actualTime} 출근 / ${row.lateMinutes}분 지각`; lateType = row.lateType; lateMinutes = row.lateMinutes; baseScore = row.lateType === '10분 이내' ? '조건부' : row.rawScore; }
    else if (kind === '근태미입력') { judgement = row.reason; baseScore = -3; }
    else if (kind === '퇴근미입력') { judgement = row.reason; baseScore = 0; }
    return { group: row.group || '', apply: '', kind, name: row.name, date: row.date, autoJudgement: judgement, lateType, lateMinutes, baseScore, result: '', excludeFormula: '', reason: '', approver: '', memo: '' };
  }

  function buildSummary(people, lateRows, noAttendanceRows, noCheckoutRows) {
    const summary = new Map();
    const byName = new Map(people.map(p => [p.name, p]));
    const ensure = (person) => {
      if (!summary.has(person.name)) summary.set(person.name, { group: person.group, name: person.name, store: person.store, late10: 0, late11: 0, late60: 0, lateTotal: 0, lateScore: 0, noAttend: 0, noAttendScore: 0, noCheckout: 0, noCheckoutScore: 0, totalScore: 0 });
      return summary.get(person.name);
    };
    for (const row of lateRows) {
      if (row.lateType === '기준시간없음') continue;
      const s = ensure(byName.get(row.name) || { group: row.group, name: row.name, store: row.store });
      if (row.lateType === '10분 이내') s.late10++;
      else if (row.lateType === '11~59분') s.late11++;
      else if (row.lateType === '60분 이상') s.late60++;
    }
    for (const row of noAttendanceRows) ensure(byName.get(row.name) || { group: row.group, name: row.name, store: row.store }).noAttend++;
    for (const row of noCheckoutRows) ensure(byName.get(row.name) || { group: row.group, name: row.name, store: row.store }).noCheckout++;
    for (const s of summary.values()) {
      s.lateTotal = s.late10 + s.late11 + s.late60;
      s.lateScore = (s.lateTotal >= 3 ? -s.late10 : 0) - s.late11 - (s.late60 * 2);
      s.noAttendScore = -3 * s.noAttend;
      s.noCheckoutScore = s.noCheckout >= 3 ? -(s.noCheckout - 2) : 0;
      s.totalScore = s.lateScore + s.noAttendScore + s.noCheckoutScore;
    }
    return Array.from(summary.values()).sort((a, b) => a.group.localeCompare(b.group, 'ko') || a.store.localeCompare(b.store, 'ko') || a.name.localeCompare(b.name, 'ko'));
  }

  function makeWorkbook(result) {
    const wb = XLSX.utils.book_new();
    wb.Workbook = { Views: [{ RTL: false }], CalcPr: { calcMode: 'auto' } };

    // 요청 시트만 생성합니다. 보고용 파일이 너무 길어지지 않도록 사용방법/자동요약/기타 보조시트는 제외합니다.
    addNoAttendanceSheet(wb, result.noAttendanceRows.filter(x => x.group === 'CE'), 'CE 근태 미입력');
    addMxFinalSummarySheet(wb, result);
    addLateSheet(wb, result.lateRows.filter(x => x.group === 'MX'), 'MX 지각');
    addNoAttendanceSheet(wb, result.noAttendanceRows.filter(x => x.group === 'MX'), 'MX 근태 미입력');
    addNoCheckoutSheet(wb, result.noCheckoutRows.filter(x => x.group === 'MX'), 'MX 퇴근 미입력');
    addMxExceptionSheet(wb, result.exceptionRows.filter(x => x.group === 'MX'));
    addReadWorksSheet(wb, result.worksReadRows, result.baseDate);
    addReadPlanSheet(wb, result.planReadRows, result.baseDate);
    addMismatchSheet(wb, result.mismatchRows);

    const fileName = `근태분석결과_${result.year}${String(result.month).padStart(2, '0')}_${result.baseDate.replace(/-/g, '')}_v9.xlsx`;
    XLSX.writeFile(wb, fileName, { bookType: 'xlsx', cellStyles: true });
  }

  function thinBorder(color = 'D9E5DF') { return { top: { style: 'thin', color: { rgb: color } }, bottom: { style: 'thin', color: { rgb: color } }, left: { style: 'thin', color: { rgb: color } }, right: { style: 'thin', color: { rgb: color } } }; }
  function headerStyle() { return { fill: { fgColor: { rgb: '0B6B43' } }, font: { color: { rgb: 'FFFFFF' }, bold: true, sz: 10.5 }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: thinBorder('0B6B43') }; }
  function bodyStyle(rowIndex) { return { fill: { fgColor: { rgb: rowIndex % 2 === 0 ? 'FFFFFF' : 'F6FAF8' } }, font: { color: { rgb: '1B2A23' }, sz: 10 }, alignment: { vertical: 'center', wrapText: true }, border: thinBorder('E4ECE8') }; }
  function titleStyle() { return { fill: { fgColor: { rgb: '083D2C' } }, font: { color: { rgb: 'FFFFFF' }, bold: true, sz: 18 }, alignment: { horizontal: 'left', vertical: 'center' } }; }
  function subTitleStyle() { return { fill: { fgColor: { rgb: 'EAF6EF' } }, font: { color: { rgb: '0B6B43' }, bold: true, sz: 10.5 }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: thinBorder('BFE4CF') }; }
  function kpiStyle(color = 'EAF6EF') { return { fill: { fgColor: { rgb: color } }, font: { color: { rgb: '0B2F22' }, bold: true, sz: 11 }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: thinBorder('A7D7B8') }; }

  function addSheet(wb, name, rows, options = {}) {
    const safeRows = rows.length ? rows : [['내용 없음']];
    const ws = XLSX.utils.aoa_to_sheet(safeRows);
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
    const headerRow = options.headerRow ?? 0;
    const titleRows = new Set(options.titleRows || []);
    const subtitleRows = new Set(options.subtitleRows || []);
    for (let R = range.s.r; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[addr]) continue;
        ws[addr].s = bodyStyle(R);
        if (titleRows.has(R)) ws[addr].s = titleStyle();
        else if (subtitleRows.has(R)) ws[addr].s = subTitleStyle();
        else if (R === headerRow) ws[addr].s = headerStyle();
      }
    }
    ws['!cols'] = (options.widths || safeRows[0]?.map(() => 16) || [16]).map(wch => ({ wch }));
    ws['!rows'] = safeRows.map((_, i) => ({ hpt: titleRows.has(i) ? 28 : subtitleRows.has(i) ? 22 : i === headerRow ? 24 : 20 }));
    ws['!freeze'] = { xSplit: 0, ySplit: headerRow + 1 };
    ws['!views'] = [{ showGridLines: false, state: 'frozen', ySplit: headerRow + 1 }];
    if (options.autofilter !== false) {
      const refRange = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
      const filterRef = XLSX.utils.encode_range({ s: { r: headerRow, c: refRange.s.c }, e: { r: refRange.e.r, c: refRange.e.c } });
      ws['!autofilter'] = { ref: filterRef };
    }
    XLSX.utils.book_append_sheet(wb, ws, name);
    return ws;
  }

  function addGuideRuleSheet(wb, result) {
    const rows = [
      ['근태 자동 분석 결과 사용방법 & 판정기준', '', ''],
      ['분석월', `${result.year}-${String(result.month).padStart(2, '0')}`, ''],
      ['판정기준일', result.baseDate, '근태관리 파일의 마지막 출근일 기준. 직접 입력한 기준일이 있으면 해당 기준 사용'],
      ['', '', ''],
      ['구분', '기준', '감점/처리'],
      ['업로드 파일', '인력 및 점포별 근무시간 / 매장근무계획관리 / 웍스스케줄 / 근태관리', '엑셀 4개 업로드 후 분석 실행'],
      ['CE 기준', '매장근무계획관리에서 “근무” 포함일만 확인', '출근기록 없으면 CE_근태미입력에 표시'],
      ['MX 기본 기준', '웍스스케줄 A/B/C조 기준', '영업시간DB의 점포별 조 시간 사용'],
      ['MX 순환 기준', '웍스스케줄 값 어디든 “순환” 글자 포함', '매장근무계획관리의 근무A/B/C 기준 사용'],
      ['교육 기준', '웍스스케줄에 “교육” 포함', '지각/근태미입력 판정 제외'],
      ['제외 기준', '휴무/휴일/연차/휴가/공가/대체/보상/DIDA/예비군/병가/경조', '근태미입력 감점 제외'],
      ['지각 판정', '기준시간과 같은 시간에 출근 찍어도 지각', '예: 10:00 기준 / 10:00 출근 = 지각 0분'],
      ['10분 이내 지각', '전체 지각 횟수가 3회 이상일 때만 감점', '1회당 -1점'],
      ['11~59분 지각', '횟수 상관없이 감점', '1회당 -1점'],
      ['60분 이상 지각', '횟수 상관없이 감점', '1회당 -2점'],
      ['근태 미입력', '근무 예정이나 출근 기록 없음', '1회당 -3점'],
      ['퇴근 미입력', '출근은 있으나 정상 퇴근시간 없음', '3회부터 -1점, 이후 1회 추가마다 -1점'],
      ['소명 처리', 'MX_소명처리 시트의 처리결과 칸에 정시인정/출근인정/퇴근인정/교육제외/기타제외 입력', 'MX_최종요약에 자동 반영. “인정/제외/정시” 문구가 있으면 제외 처리'],
      ['확인 시트', '읽기 확인_웍스 / 읽기 확인_매장계획', '프로그램이 실제로 읽은 원본셀과 값 확인 가능'],
    ];
    const ws = addSheet(wb, '사용방법_판정기준', rows, { widths: [18, 62, 62], titleRows: [0], subtitleRows: [4], headerRow: 4, autofilter: false });
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }];
  }

  function mxExclusionCountFormula(kind, nameCell, extraCriteria = []) {
    const base = [`'MX 소명처리'!$A:$A,"${kind}"`, `'MX 소명처리'!$B:$B,${nameCell}`, `'MX 소명처리'!$K:$K,"제외"`];
    return `COUNTIFS(${base.concat(extraCriteria).join(',')})`;
  }

  function addMxFinalSummarySheet(wb, result) {
    const mxPeople = result.people.filter(p => p.group === 'MX');
    const startRow = 9; // 엑셀 실제 행 번호. 1~7행은 보고용 타이틀/KPI 영역, 8행은 표 헤더.
    const rows = [
      ['ZENIEL MX 근태 최종 요약 보고', '', '', '', '', '', '', '', '', '', '', '', ''],
      ['분석월', `${result.year}-${String(result.month).padStart(2, '0')}`, '판정기준일', result.baseDate, 'MX 인원', mxPeople.length, '소명제외', '', '최종 총감점', '', '', '', ''],
      ['지각 대상 인원', '', '근태 미입력 건수', '', '퇴근 미입력 건수', '', '감점 대상자', '', '관리 기준', '소명 입력 시 MX 최종 요약 자동 반영', '', '', ''],
      ['', '', '', '', '', '', '', '', '', '', '', '', ''],
      ['운영 기준', '① 웍스스케줄에 “순환” 포함 시 매장근무계획의 근무A/B/C 기준  ② 교육은 제외  ③ 10분 이내 지각은 최종 지각 총횟수 3회 이상 시 감점', '', '', '', '', '', '', '', '', '', '', ''],
      ['감점 기준', '11~59분: 1회당 -1점 / 60분 이상: 1회당 -2점 / 근태 미입력: 1회당 -3점 / 퇴근 미입력: 3회부터 감점', '', '', '', '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', '', '', '', '', ''],
      ['구분', '이름', '점포', '10분 이내 지각', '11~59분 지각', '60분 이상 지각', '지각 총횟수', '지각감점', '근태미입력', '근태감점', '퇴근미입력', '퇴근감점', '총감점'],
    ];
    for (const p of mxPeople) rows.push([p.group, p.name, p.store, '', '', '', '', '', '', '', '', '', '']);
    const ws = addSheet(wb, 'MX 최종 요약', rows, { widths: [8, 13, 16, 14, 14, 14, 12, 10, 12, 10, 12, 10, 10], titleRows: [0], subtitleRows: [1, 2, 4, 5], headerRow: 7, autofilter: true });
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 12 } },
      { s: { r: 4, c: 1 }, e: { r: 4, c: 12 } },
      { s: { r: 5, c: 1 }, e: { r: 5, c: 12 } },
      { s: { r: 2, c: 9 }, e: { r: 2, c: 12 } },
    ];
    const lastRow = rows.length;

    ws['B2'] = { t: 's', v: `${result.year}-${String(result.month).padStart(2, '0')}`, s: kpiStyle('F7FBF9') };
    ws['D2'] = { t: 's', v: result.baseDate, s: kpiStyle('F7FBF9') };
    ws['F2'] = { t: 'n', v: mxPeople.length, s: kpiStyle('F7FBF9') };
    ws['H2'] = { t: 'n', f: `COUNTIF('MX 소명처리'!$K:$K,"제외")`, s: kpiStyle('EAF6EF') };
    ws['J2'] = { t: 'n', f: `SUM(M${startRow}:M${lastRow})`, s: kpiStyle('FFF4E6') };
    ws['B3'] = { t: 'n', f: `COUNTIF(G${startRow}:G${lastRow},">0")`, s: kpiStyle('EAF6EF') };
    ws['D3'] = { t: 'n', f: `SUM(I${startRow}:I${lastRow})`, s: kpiStyle('FFF4E6') };
    ws['F3'] = { t: 'n', f: `SUM(K${startRow}:K${lastRow})`, s: kpiStyle('FFF4E6') };
    ws['H3'] = { t: 'n', f: `COUNTIF(M${startRow}:M${lastRow},"<0")`, s: kpiStyle('FFECEC') };

    for (let r = startRow; r <= rows.length; r++) {
      const formulaStyle = bodyStyle(r);
      ws[`D${r}`] = { t: 'n', f: `MAX(0,COUNTIFS('MX 지각'!$B:$B,$B${r},'MX 지각'!$J:$J,"10분 이내")-${mxExclusionCountFormula('지각', `$B${r}`, [`'MX 소명처리'!$E:$E,"10분 이내"`])})`, s: formulaStyle };
      ws[`E${r}`] = { t: 'n', f: `MAX(0,COUNTIFS('MX 지각'!$B:$B,$B${r},'MX 지각'!$J:$J,"11~59분")-${mxExclusionCountFormula('지각', `$B${r}`, [`'MX 소명처리'!$E:$E,"11~59분"`])})`, s: formulaStyle };
      ws[`F${r}`] = { t: 'n', f: `MAX(0,COUNTIFS('MX 지각'!$B:$B,$B${r},'MX 지각'!$J:$J,"60분 이상")-${mxExclusionCountFormula('지각', `$B${r}`, [`'MX 소명처리'!$E:$E,"60분 이상"`])})`, s: formulaStyle };
      ws[`G${r}`] = { t: 'n', f: `SUM(D${r}:F${r})`, s: formulaStyle };
      ws[`H${r}`] = { t: 'n', f: `IF(G${r}>=3,-D${r},0)-E${r}-F${r}*2`, s: { ...formulaStyle, font: { color: { rgb: 'B42318' }, bold: true, sz: 10 } } };
      ws[`I${r}`] = { t: 'n', f: `MAX(0,COUNTIFS('MX 근태 미입력'!$B:$B,$B${r})-${mxExclusionCountFormula('근태미입력', `$B${r}`)})`, s: formulaStyle };
      ws[`J${r}`] = { t: 'n', f: `-I${r}*3`, s: { ...formulaStyle, font: { color: { rgb: 'B42318' }, bold: true, sz: 10 } } };
      ws[`K${r}`] = { t: 'n', f: `MAX(0,COUNTIFS('MX 퇴근 미입력'!$B:$B,$B${r})-${mxExclusionCountFormula('퇴근미입력', `$B${r}`)})`, s: formulaStyle };
      ws[`L${r}`] = { t: 'n', f: `IF(K${r}>=3,-(K${r}-2),0)`, s: { ...formulaStyle, font: { color: { rgb: 'B42318' }, bold: true, sz: 10 } } };
      ws[`M${r}`] = { t: 'n', f: `H${r}+J${r}+L${r}`, s: { ...formulaStyle, fill: { fgColor: { rgb: 'FFF4E6' } }, font: { color: { rgb: 'B42318' }, bold: true, sz: 10 } } };
    }
    [8, 10, 12, 13].forEach(c => colorNegative(ws, rows.length, c, startRow));
  }

  function addMxExceptionSheet(wb, items) {
    const rows = [['구분', '이름', '날짜', '자동판정', '지각구분', '지각분', '구간점수(참고)', '처리결과', '처리사유', '승인자', '최종제외', '비고']];
    for (const x of items) rows.push([x.kind, x.name, x.date, x.autoJudgement, x.lateType, x.lateMinutes, x.baseScore, x.result, x.reason, x.approver, '', x.memo]);
    const ws = addSheet(wb, 'MX 소명처리', rows, { widths: [12, 13, 12, 36, 12, 8, 13, 18, 38, 12, 10, 28] });
    for (let r = 2; r <= rows.length; r++) {
      // 처리결과만 입력해도 적용. 처리결과/처리사유/승인자/비고 중 인정·제외·정시 문구가 있으면 제외.
      ws[`K${r}`] = { t: 's', f: `IF(OR(ISNUMBER(SEARCH("인정",$H${r}&$I${r}&$J${r}&$L${r})),ISNUMBER(SEARCH("제외",$H${r}&$I${r}&$J${r}&$L${r})),ISNUMBER(SEARCH("정시",$H${r}&$I${r}&$J${r}&$L${r}))),"제외","반영")`, s: bodyStyle(r) };
    }
  }

  function addSummarySheet(wb, name, summaryRows) {
    const rows = [['구분', '이름', '점포', '10분 이내 지각', '11~59분 지각', '60분 이상 지각', '지각 총횟수', '지각감점', '근태미입력', '근태감점', '퇴근미입력', '퇴근감점', '총감점']];
    for (const s of summaryRows) rows.push([s.group, s.name, s.store, s.late10, s.late11, s.late60, s.lateTotal, s.lateScore, s.noAttend, s.noAttendScore, s.noCheckout, s.noCheckoutScore, s.totalScore]);
    const ws = addSheet(wb, name, rows, { widths: [8, 12, 16, 14, 14, 14, 12, 10, 12, 10, 12, 10, 10] });
    [8, 10, 12, 13].forEach(c => colorNegative(ws, rows.length, c));
  }

  function exclusionCountFormula(kind, nameCell, extraCriteria = []) {
    // 소명처리 J열(최종제외)이 "제외"인 행만 최종요약에서 차감한다.
    // J열은 처리결과(I), 처리사유(K), 승인자(L), 비고(M)에 인정/제외/정시가 들어가면 자동으로 "제외"가 된다.
    const base = [`'소명처리'!$B:$B,"${kind}"`, `'소명처리'!$C:$C,${nameCell}`, `'소명처리'!$J:$J,"제외"`];
    const criteria = base.concat(extraCriteria);
    return `COUNTIFS(${criteria.join(',')})`;
  }

  function addFinalSummarySheet(wb, result, sheetName, sourceRows) {
    const people = sourceRows.length ? sourceRows : result.people.filter(p => sheetName.startsWith('MX_') ? p.group === 'MX' : sheetName.startsWith('CE_') ? p.group === 'CE' : true).map(p => ({ group: p.group, name: p.name, store: p.store }));
    const rows = [['구분', '이름', '점포', '10분 이내 지각', '11~59분 지각', '60분 이상 지각', '지각 총횟수', '지각감점', '근태미입력', '근태감점', '퇴근미입력', '퇴근감점', '총감점']];
    for (const s of people) rows.push([s.group, s.name, s.store, '', '', '', '', '', '', '', '', '', '']);
    const ws = addSheet(wb, sheetName, rows, { widths: [8, 12, 16, 14, 14, 14, 12, 10, 12, 10, 12, 10, 10] });
    for (let r = 2; r <= rows.length; r++) {
      ws[`D${r}`] = { t: 'n', f: `COUNTIFS('자동판정_지각'!$B:$B,$B${r},'자동판정_지각'!$J:$J,"10분 이내")-${exclusionCountFormula('지각', `$B${r}`, [`'소명처리'!$F:$F,"10분 이내"`])}` };
      ws[`E${r}`] = { t: 'n', f: `COUNTIFS('자동판정_지각'!$B:$B,$B${r},'자동판정_지각'!$J:$J,"11~59분")-${exclusionCountFormula('지각', `$B${r}`, [`'소명처리'!$F:$F,"11~59분"`])}` };
      ws[`F${r}`] = { t: 'n', f: `COUNTIFS('자동판정_지각'!$B:$B,$B${r},'자동판정_지각'!$J:$J,"60분 이상")-${exclusionCountFormula('지각', `$B${r}`, [`'소명처리'!$F:$F,"60분 이상"`])}` };
      ws[`G${r}`] = { t: 'n', f: `SUM(D${r}:F${r})` };
      ws[`H${r}`] = { t: 'n', f: `IF(G${r}>=3,-D${r},0)-E${r}-F${r}*2` };
      ws[`I${r}`] = { t: 'n', f: `COUNTIFS('자동판정_근태미입력'!$B:$B,$B${r})-${exclusionCountFormula('근태미입력', `$B${r}`)}` };
      ws[`J${r}`] = { t: 'n', f: `-I${r}*3` };
      ws[`K${r}`] = { t: 'n', f: `COUNTIFS('자동판정_퇴근미입력'!$B:$B,$B${r})-${exclusionCountFormula('퇴근미입력', `$B${r}`)}` };
      ws[`L${r}`] = { t: 'n', f: `IF(K${r}>=3,-(K${r}-2),0)` };
      ws[`M${r}`] = { t: 'n', f: `H${r}+J${r}+L${r}` };
    }
    [8, 10, 12, 13].forEach(c => colorNegative(ws, rows.length, c));
  }

  function addLateSheet(wb, items, name) {
    const rows = [['구분', '이름', '날짜', '출근지점', '기준점포', '조', '기준시트', '스케줄', '기준시간', '지각구분', '실제출근', '지각분', '최종요약 반영점수', '비고']];
    for (const x of items) rows.push([x.group, x.name, x.date, x.store, x.standardStore || x.store, x.shift, x.basis, x.schedule, x.standardTime, x.lateType, x.actualTime, x.lateMinutes, '', '']);
    const ws = addSheet(wb, name, rows, { widths: [8, 12, 12, 16, 16, 8, 22, 16, 10, 12, 10, 8, 16, 34] });
    for (let r = 2; r <= rows.length; r++) {
      const exclusion = `COUNTIFS('MX 소명처리'!$A:$A,"지각",'MX 소명처리'!$B:$B,$B${r},'MX 소명처리'!$C:$C,$C${r},'MX 소명처리'!$E:$E,$J${r},'MX 소명처리'!$K:$K,"제외")`;
      const finalLateTotal = `IFERROR(INDEX('MX 최종 요약'!$G:$G,MATCH($B${r},'MX 최종 요약'!$B:$B,0)),0)`;
      ws[`M${r}`] = { t: 'n', f: `IF(${exclusion}>0,0,IF($J${r}="60분 이상",-2,IF($J${r}="11~59분",-1,IF(${finalLateTotal}>=3,-1,0))))`, s: { ...bodyStyle(r), fill: { fgColor: { rgb: 'FFF8E8' } }, font: { color: { rgb: 'B42318' }, bold: true, sz: 10 } } };
      ws[`N${r}`] = { t: 's', f: `IF(${exclusion}>0,"소명 제외",IF($J${r}="10분 이내",IF(${finalLateTotal}>=3,"지각 총 3회 이상으로 감점 반영","지각 총 3회 미만으로 감점 없음"),"즉시 감점 반영"))`, s: bodyStyle(r) };
    }
  }
  function addNoAttendanceSheet(wb, items, name) {
    const rows = [['구분', '이름', '날짜', '점포', '기준시트', '스케줄', '판정', '기본감점']];
    for (const x of items) rows.push([x.group, x.name, x.date, x.store, x.basis, x.schedule, x.reason, -3]);
    addSheet(wb, name, rows, { widths: [8, 12, 12, 16, 22, 16, 32, 10] });
  }
  function addNoCheckoutSheet(wb, items, name) {
    const rows = [['구분', '이름', '날짜', '출근지점', '출근시간', '퇴근시간', '판정']];
    for (const x of items) rows.push([x.group, x.name, x.date, x.store, x.checkIn, x.checkOut, x.reason]);
    addSheet(wb, name, rows, { widths: [8, 12, 12, 16, 10, 10, 32] });
  }
  function addMismatchSheet(wb, items) {
    const rows = [['이름', '날짜', '점포', '매장근무계획상', '웍스스케줄상', '내용']];
    for (const x of items) rows.push([x.name, x.date, x.store, x.plan, x.works, x.note]);
    addSheet(wb, '스케줄 불일치', rows, { widths: [12, 12, 16, 18, 18, 55] });
  }
  function addCeSheet(wb, items) {
    const rows = [['구분', '이름', '날짜', '점포', '매장근무계획', '출근확인']];
    for (const x of items) rows.push([x.group, x.name, x.date, x.store, x.schedule, x.status]);
    addSheet(wb, 'CE출근확인', rows, { widths: [8, 12, 12, 16, 18, 12] });
  }
  function addExceptionSheet(wb, items) {
    const rows = [['적용여부', '구분', '이름', '날짜', '자동판정', '지각구분', '지각분', '기본감점', '처리결과', '최종제외', '처리사유', '승인자', '비고']];
    for (const x of items) rows.push([x.apply, x.kind, x.name, x.date, x.autoJudgement, x.lateType, x.lateMinutes, x.baseScore, x.result, '', x.reason, x.approver, x.memo]);
    const ws = addSheet(wb, '소명처리', rows, { widths: [10, 12, 12, 12, 35, 12, 8, 10, 18, 10, 38, 12, 28] });
    for (let r = 2; r <= rows.length; r++) {
      // 적용여부는 비워도 적용. 단, A열에 "미적용"을 쓰면 제외하지 않음.
      // 처리결과(I), 처리사유(K), 승인자(L), 비고(M) 어디든 인정/제외/정시 문구가 있으면 제외.
      ws[`J${r}`] = { t: 's', f: `IF($A${r}="미적용","반영",IF(OR(ISNUMBER(SEARCH("인정",$I${r}&$K${r}&$L${r}&$M${r})),ISNUMBER(SEARCH("제외",$I${r}&$K${r}&$L${r}&$M${r})),ISNUMBER(SEARCH("정시",$I${r}&$K${r}&$L${r}&$M${r}))),"제외","반영"))` };
    }
  }

  function addReadWorksSheet(wb, items, baseDate) {
    const rows = [['이름', '날짜', '점포', '웍스스케줄 원값', '순환여부', '적용기준', '적용 조', '매장계획 원값', '매장계획 조', '읽은 조', '원본셀', '원본행', '원본열', '헤더셀', '헤더값', '휴무/제외값 여부']];
    for (const x of items.filter(v => v.date <= baseDate)) rows.push([
      x.name, x.date, x.store, x.value,
      hasRotation(x.value) ? '순환' : '',
      x.appliedBasis || (hasEducation(x.value) ? '교육 제외' : hasRotation(x.value) ? '순환→매장계획 기준' : isOffLike(x.value) ? '휴무/제외' : '웍스 기준'),
      x.appliedShift || '', x.planValue || '', x.planShift || '', x.shift,
      x.sourceCell, x.sourceRow, x.sourceCol, x.headerCell, x.headerValue,
      isOffLike(x.value) || hasEducation(x.value) ? '제외값' : '근무값'
    ]);
    addSheet(wb, '읽기 확인_웍스', rows, { widths: [12, 12, 16, 18, 10, 18, 10, 18, 10, 10, 10, 8, 8, 10, 12, 16] });
  }

  function addReadPlanSheet(wb, items, baseDate) {
    const rows = [['이름', '날짜', '점포', '매장근무계획 원값', '읽은 조', '원본셀', '원본행', '원본열', '휴무/제외값 여부']];
    for (const x of items.filter(v => v.date <= baseDate)) rows.push([x.name, x.date, x.store, x.value, x.shift, x.sourceCell, x.sourceRow, x.sourceCol, isOffLike(x.value) ? '제외값' : '근무값']);
    addSheet(wb, '읽기 확인_매장계획', rows, { widths: [12, 12, 16, 18, 10, 10, 8, 8, 16] });
  }

  function addRuleSheet(wb) {
    const rows = [
      ['구분', '기준', '감점/처리'],
      ['지각 판정', '기준시간과 같은 시간에 출근 찍어도 지각', '예: 10:00 기준 / 10:00 출근 = 지각 0분'],
      ['10분 이내 지각', '전체 지각 횟수가 3회 이상일 때만 감점', '1회당 -1점'],
      ['11~59분 지각', '횟수 상관없이 감점', '1회당 -1점'],
      ['60분 이상 지각', '횟수 상관없이 감점', '1회당 -2점'],
      ['근태 미입력', '근무 예정이나 출근 기록 없음', '1회당 -3점'],
      ['퇴근 미입력', '출근은 있으나 정상 퇴근시간 없음', '3회부터 -1점, 이후 1회 추가마다 -1점'],
      ['MX 기본 기준', '웍스스케줄 A/B/C조 기준', '영업시간DB의 점포별 조 시간 사용'],
      ['MX 순환 기준', '웍스스케줄 값 어디든 “순환” 글자 포함', '매장근무계획관리의 근무A/B/C 기준 사용'],
      ['제외 기준', '휴무/대체휴무/보상휴가/연차/공가/DIDA/예비군/교육', '근태 미입력 감점 제외'],
      ['CE 기준', '매장근무계획관리에 “근무” 포함, 단 휴무성 문구 제외', '출근 여부만 확인'],
      ['소명 처리', '소명처리 시트에서 처리결과 또는 처리사유에 입력해도 적용됨. 적용여부는 공란 가능', '정시인정/출근인정/퇴근인정/교육제외/기타제외 또는 인정/제외/정시 포함 문구 시 최종요약에서 제외. 적용여부=미적용이면 제외 안 함'],
    ];
    addSheet(wb, '판정기준', rows, { widths: [16, 55, 55] });
  }

  function colorNegative(ws, rowCount, colOneBased, startRow = 2) {
    for (let r = startRow; r <= rowCount; r++) {
      const addr = XLSX.utils.encode_cell({ r: r - 1, c: colOneBased - 1 });
      ws[addr] = ws[addr] || { t: 'n', v: 0 };
      ws[addr].s = ws[addr].s || {};
      ws[addr].s.font = { color: { rgb: 'C92A2A' }, bold: true, sz: 10 };
    }
  }
  function escapeHtml(s) { return String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch])); }
})();
