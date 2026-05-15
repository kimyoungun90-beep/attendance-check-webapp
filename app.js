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

  $('runBtn').addEventListener('click', runAnalysis);
  $('resetBtn').addEventListener('click', () => location.reload());

  function setStatus(message, type = '') {
    statusEl.className = `status ${type}`.trim();
    statusEl.innerHTML = message;
  }

  function clean(value) {
    if (value === null || value === undefined) return '';
    return String(value).replace(/\u00a0/g, '').trim();
  }

  function normalizeName(value) {
    return clean(value).replace(/\s+/g, '');
  }

  function normalizeStore(value) {
    let s = clean(value)
      .replace(/코스트코/g, '')
      .replace(/\s+/g, '')
      .replace(/점$/g, '점')
      .trim();
    if (s === '혁신점') s = '대구혁신점';
    if (s === '대구혁신') s = '대구혁신점';
    if (s && !s.endsWith('점') && !s.includes('사무소')) s += '점';
    return s;
  }

  function excelSerialToDate(serial) {
    const utcDays = Math.floor(Number(serial) - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date(utcValue * 1000);
    return new Date(dateInfo.getUTCFullYear(), dateInfo.getUTCMonth(), dateInfo.getUTCDate());
  }

  function dateKey(date) {
    if (!date) return '';
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
    let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    m = s.match(/^(\d{1,2})일$/);
    if (m && fallbackYear && fallbackMonth) return new Date(fallbackYear, fallbackMonth - 1, Number(m[1]));
    return null;
  }

  function parseTimeToMinutes(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value.getHours() * 60 + value.getMinutes();
    if (typeof value === 'number' && Number.isFinite(value)) {
      if (value >= 0 && value < 1) return Math.round(value * 24 * 60);
      return null; // 15 같은 숫자는 정상 퇴근시간으로 보지 않음
    }
    const s = clean(value);
    if (!s) return null;
    const m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
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

  function isWorkLike(value) {
    const s = clean(value);
    return s.includes('근무') || s.includes('A조') || s.includes('B조') || s.includes('C조') || s.includes('순환');
  }

  function isOffLike(value) {
    const s = clean(value);
    return !s || s.includes('휴무') || s.includes('휴일') || s.includes('연차') || s.includes('휴가') || s.includes('공가') || s.includes('대체') || s.includes('보상') || s.includes('DIDA') || s.includes('예비군');
  }

  function hasEducation(value) {
    return clean(value).includes('교육');
  }

  function hasRotation(value) {
    return clean(value).includes('순환');
  }

  function extractPlanShift(value) {
    const s = clean(value);
    if (s.includes('근무A')) return 'A조';
    if (s.includes('근무B')) return 'B조';
    if (s.includes('근무C')) return 'C조';
    if (s === 'A조' || s.includes('A조')) return 'A조';
    if (s === 'B조' || s.includes('B조')) return 'B조';
    if (s === 'C조' || s.includes('C조')) return 'C조';
    return '';
  }

  function extractWorksShift(value) {
    const s = clean(value);
    if (s.includes('A조')) return 'A조';
    if (s.includes('B조')) return 'B조';
    if (s.includes('C조')) return 'C조';
    return '';
  }

  async function readWorkbook(file) {
    const data = await file.arrayBuffer();
    return XLSX.read(data, { type: 'array', cellDates: true, raw: true });
  }

  function sheetRows(wb, preferredNames = []) {
    let sheetName = wb.SheetNames[0];
    for (const name of preferredNames) {
      if (wb.SheetNames.includes(name)) {
        sheetName = name;
        break;
      }
    }
    const ws = wb.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
  }

  async function runAnalysis() {
    try {
      const missingFiles = Object.entries(fileInputs).filter(([, el]) => !el.files || !el.files[0]).map(([key]) => key);
      if (missingFiles.length) {
        setStatus('엑셀 파일 4개를 모두 업로드해야 합니다.', 'error');
        return;
      }
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

      setStatus(`분석 완료: 지각 ${result.lateRows.length}건, 근태 미입력 ${result.noAttendanceRows.length}건, 퇴근 미입력 ${result.noCheckoutRows.length}건, 스케줄 불일치 ${result.mismatchRows.length}건을 확인했습니다.<br>결과 엑셀이 다운로드됩니다.`, 'ok');
    } catch (err) {
      console.error(err);
      setStatus(`오류가 발생했습니다.<br><b>${escapeHtml(err.message || err)}</b><br>파일 양식이나 시트명이 바뀌었는지 확인하세요.`, 'error');
    }
  }

  function inferYearMonth(attendanceRows, worksRows) {
    const manualMonth = clean($('monthInput').value);
    if (manualMonth) {
      const [y, m] = manualMonth.split('-').map(Number);
      return { year: y, month: m };
    }
    for (let r = 1; r < attendanceRows.length; r++) {
      const d = parseDate(attendanceRows[r][1]);
      if (d) return { year: d.getFullYear(), month: d.getMonth() + 1 };
    }
    const worksMonthCell = Number(worksRows?.[0]?.[0]);
    for (let c = 4; c < (worksRows[0] || []).length; c++) {
      const d = parseDate(worksRows[0][c]);
      if (d) return { year: d.getFullYear(), month: d.getMonth() + 1 };
    }
    const now = new Date();
    return { year: now.getFullYear(), month: worksMonthCell || now.getMonth() + 1 };
  }

  function analyze({ peopleRows, hourRows, planRows, worksRows, attendanceRows }) {
    const { year, month } = inferYearMonth(attendanceRows, worksRows);
    const manualBaseDate = clean($('baseDateInput').value);

    const peopleByName = new Map();
    const peopleByEmp = new Map();
    for (let r = 1; r < peopleRows.length; r++) {
      const row = peopleRows[r];
      const group = clean(row[0]);
      const store = normalizeStore(row[4]);
      const empNo = clean(row[5]);
      const name = normalizeName(row[6]);
      if (!name) continue;
      const person = { group, store, empNo, name, manager: clean(row[1]), region: clean(row[2]) };
      peopleByName.set(name, person);
      if (empNo) peopleByEmp.set(empNo, person);
    }

    const workHours = new Map();
    for (let r = 1; r < hourRows.length; r++) {
      const row = hourRows[r];
      const store = normalizeStore(row[0]);
      if (!store) continue;
      workHours.set(store, {
        'A조': parseTimeToMinutes(row[3]),
        'B조': parseTimeToMinutes(row[4]),
        'C조': parseTimeToMinutes(row[5]),
      });
    }

    const planByNameDate = new Map();
    const planByName = new Map();
    const planHeader = planRows[0] || [];
    const dayCols = [];
    for (let c = 0; c < planHeader.length; c++) {
      const header = clean(planHeader[c]);
      const m = header.match(/^(\d{1,2})일$/);
      if (m) dayCols.push({ c, day: Number(m[1]), date: new Date(year, month - 1, Number(m[1])) });
    }
    for (let r = 1; r < planRows.length; r++) {
      const row = planRows[r];
      const name = normalizeName(row[3]);
      if (!name) continue;
      const store = normalizeStore(row[0]);
      planByName.set(name, { row, store });
      for (const dc of dayCols) {
        const dKey = dateKey(dc.date);
        planByNameDate.set(`${name}|${dKey}`, {
          name, date: dKey, value: clean(row[dc.c]), shift: extractPlanShift(row[dc.c]), store,
        });
      }
    }

    const worksByNameDate = new Map();
    const worksByName = new Map();
    const worksDateCols = [];
    const worksHeader = worksRows[0] || [];
    for (let c = 4; c < worksHeader.length; c++) {
      let d = parseDate(worksHeader[c]);
      if (!d && c - 3 >= 1 && c - 3 <= 31) d = new Date(year, month - 1, c - 3);
      if (d) worksDateCols.push({ c, date: d, dKey: dateKey(d) });
    }
    for (let r = 2; r < worksRows.length; r++) {
      const row = worksRows[r];
      const name = normalizeName(row[3]);
      if (!name) continue;
      const store = normalizeStore(row[2] || row[0]);
      worksByName.set(name, { row, store });
      for (const dc of worksDateCols) {
        worksByNameDate.set(`${name}|${dc.dKey}`, {
          name, date: dc.dKey, value: clean(row[dc.c]), shift: extractWorksShift(row[dc.c]), store,
        });
      }
    }

    const attendanceByNameDate = new Map();
    let latestAttendanceDate = null;
    for (let r = 1; r < attendanceRows.length; r++) {
      const row = attendanceRows[r];
      const rawName = normalizeName(row[0]);
      const empNo = clean(row[3]);
      const person = peopleByEmp.get(empNo) || peopleByName.get(rawName);
      const name = person?.name || rawName;
      const d = parseDate(row[1]);
      if (!name || !d) continue;
      const dKey = dateKey(d);
      if (!latestAttendanceDate || d > latestAttendanceDate) latestAttendanceDate = d;
      const inMin = parseTimeToMinutes(row[2]);
      const outMin = parseTimeToMinutes(row[4]);
      const inStore = normalizeStore(row[7]);
      const outStore = normalizeStore(row[8]);
      const key = `${name}|${dKey}`;
      if (!attendanceByNameDate.has(key)) {
        attendanceByNameDate.set(key, { name, date: dKey, ins: [], outs: [], stores: [], outStores: [], rows: [] });
      }
      const rec = attendanceByNameDate.get(key);
      if (inMin !== null) rec.ins.push(inMin);
      if (outMin !== null) rec.outs.push(outMin);
      if (inStore) rec.stores.push({ store: inStore, inMin: inMin ?? 99999 });
      if (outStore) rec.outStores.push(outStore);
      rec.rows.push(row);
    }
    for (const rec of attendanceByNameDate.values()) {
      rec.firstIn = rec.ins.length ? Math.min(...rec.ins) : null;
      rec.hasOut = rec.outs.length > 0;
      rec.outTime = rec.outs.length ? minutesToHHMM(Math.max(...rec.outs)) : '';
      const sortedStores = rec.stores.slice().sort((a, b) => a.inMin - b.inMin);
      rec.store = sortedStores[0]?.store || '';
    }

    let baseDate = latestAttendanceDate;
    if (manualBaseDate) baseDate = parseDate(manualBaseDate);
    if (!baseDate) throw new Error('근태관리 파일에서 판정 기준일을 찾지 못했습니다.');
    const baseDateKey = dateKey(baseDate);

    const lateRows = [];
    const noAttendanceRows = [];
    const noCheckoutRows = [];
    const mismatchRows = [];
    const ceRows = [];
    const exceptionRows = [];

    const allPeople = Array.from(peopleByName.values()).sort((a, b) => a.store.localeCompare(b.store, 'ko') || a.name.localeCompare(b.name, 'ko'));

    // MX schedule mismatch check
    for (const person of allPeople) {
      if (person.group !== 'MX') continue;
      for (const dc of worksDateCols) {
        if (dc.dKey > baseDateKey) continue;
        const works = worksByNameDate.get(`${person.name}|${dc.dKey}`);
        const plan = planByNameDate.get(`${person.name}|${dc.dKey}`);
        const worksText = works?.value || '';
        const planText = plan?.value || '';
        if (hasEducation(worksText) || hasRotation(worksText)) continue;
        const ps = extractPlanShift(planText);
        const ws = extractWorksShift(worksText);
        if ((ps || ws) && ps !== ws) {
          mismatchRows.push({
            name: person.name, date: dc.dKey, store: person.store,
            plan: planText, works: worksText,
            note: `매장근무계획상 ${planText || '공란'} / 웍스스케줄상 ${worksText || '공란'}`,
          });
        }
      }
    }

    // MX lateness and no attendance check
    for (const person of allPeople) {
      if (person.group !== 'MX') continue;
      for (const dc of worksDateCols) {
        if (dc.dKey > baseDateKey) continue;
        const works = worksByNameDate.get(`${person.name}|${dc.dKey}`);
        const plan = planByNameDate.get(`${person.name}|${dc.dKey}`);
        const worksText = works?.value || '';
        const planText = plan?.value || '';
        if (!worksText || isOffLike(worksText)) continue;
        if (hasEducation(worksText)) continue;

        let basis = '웍스스케줄';
        let shift = extractWorksShift(worksText);
        let scheduleText = worksText;
        if (hasRotation(worksText)) {
          basis = '매장근무계획관리(순환)';
          shift = extractPlanShift(planText);
          scheduleText = planText || worksText;
        }
        if (!shift) continue;

        const rec = attendanceByNameDate.get(`${person.name}|${dc.dKey}`);
        if (!rec || rec.firstIn === null) {
          const row = {
            group: 'MX', name: person.name, date: dc.dKey, store: works?.store || person.store,
            basis, schedule: scheduleText, reason: '근무 예정이나 출근기록 없음', score: -3,
          };
          noAttendanceRows.push(row);
          exceptionRows.push(toExceptionRow(row, '근태미입력'));
          continue;
        }

        const attendStore = rec.store || works?.store || person.store;
        let standard = workHours.get(attendStore)?.[shift];
        let standardStore = attendStore;
        if (standard === null || standard === undefined) {
          standard = workHours.get(person.store)?.[shift];
          standardStore = person.store;
        }
        if (standard === null || standard === undefined) {
          lateRows.push({ group: 'MX', name: person.name, date: dc.dKey, store: attendStore, shift, basis, schedule: scheduleText, standardTime: '기준없음', actualTime: minutesToHHMM(rec.firstIn), lateMinutes: '', lateType: '기준시간없음', rawScore: 0 });
          continue;
        }
        if (rec.firstIn >= standard) {
          const lateMinutes = rec.firstIn - standard;
          const lateType = lateMinutes <= 10 ? '10분 이내' : lateMinutes < 60 ? '11~59분' : '60분 이상';
          const rawScore = lateType === '60분 이상' ? -2 : -1;
          const row = {
            group: 'MX', name: person.name, date: dc.dKey, store: attendStore, standardStore, shift, basis,
            schedule: scheduleText, standardTime: minutesToHHMM(standard), actualTime: minutesToHHMM(rec.firstIn),
            lateMinutes, lateType, rawScore,
          };
          lateRows.push(row);
          exceptionRows.push(toExceptionRow(row, '지각'));
        }
      }
    }

    // CE no attendance check
    for (const person of allPeople) {
      if (person.group !== 'CE') continue;
      for (const dc of dayCols) {
        const dKey = dateKey(dc.date);
        if (dKey > baseDateKey) continue;
        const plan = planByNameDate.get(`${person.name}|${dKey}`);
        const planText = plan?.value || '';
        if (!planText.includes('근무')) continue;
        const rec = attendanceByNameDate.get(`${person.name}|${dKey}`);
        const ceRow = { group: 'CE', name: person.name, date: dKey, store: person.store, schedule: planText, status: rec?.firstIn !== null && rec?.firstIn !== undefined ? '출근확인' : '근태미입력' };
        ceRows.push(ceRow);
        if (!rec || rec.firstIn === null) {
          const row = {
            group: 'CE', name: person.name, date: dKey, store: person.store,
            basis: '매장근무계획관리', schedule: planText, reason: '근무 예정이나 출근기록 없음', score: -3,
          };
          noAttendanceRows.push(row);
          exceptionRows.push(toExceptionRow(row, '근태미입력'));
        }
      }
    }

    // Checkout missing: apply to DB members with valid check-in and no valid checkout
    for (const rec of attendanceByNameDate.values()) {
      if (rec.date > baseDateKey) continue;
      const person = peopleByName.get(rec.name);
      if (!person) continue;
      if (rec.firstIn !== null && !rec.hasOut) {
        const row = {
          group: person.group, name: person.name, date: rec.date, store: rec.store || person.store,
          checkIn: minutesToHHMM(rec.firstIn), checkOut: '', reason: '출근은 있으나 정상 퇴근시간 없음', score: 0,
        };
        noCheckoutRows.push(row);
        exceptionRows.push(toExceptionRow(row, '퇴근미입력'));
      }
    }

    const autoSummary = buildSummary(allPeople, lateRows, noAttendanceRows, noCheckoutRows, false);

    return {
      year, month, baseDate: baseDateKey,
      people: allPeople,
      lateRows,
      noAttendanceRows,
      noCheckoutRows,
      mismatchRows,
      ceRows,
      exceptionRows,
      autoSummary,
    };
  }

  function toExceptionRow(row, kind) {
    let judgement = '';
    let lateType = '';
    let lateMinutes = '';
    let baseScore = '';
    if (kind === '지각') {
      judgement = `${row.actualTime} 출근 / ${row.lateMinutes}분 지각`;
      lateType = row.lateType;
      lateMinutes = row.lateMinutes;
      baseScore = row.rawScore;
    } else if (kind === '근태미입력') {
      judgement = row.reason;
      baseScore = -3;
    } else if (kind === '퇴근미입력') {
      judgement = row.reason;
      baseScore = 0;
    }
    return {
      apply: '', kind, name: row.name, date: row.date, autoJudgement: judgement,
      lateType, lateMinutes, baseScore,
      result: '', excludeFormula: '', reason: '', approver: '', memo: '',
    };
  }

  function buildSummary(people, lateRows, noAttendanceRows, noCheckoutRows) {
    const summary = new Map();
    const ensure = (person) => {
      if (!summary.has(person.name)) {
        summary.set(person.name, {
          group: person.group, name: person.name, store: person.store,
          late10: 0, late11: 0, late60: 0, lateTotal: 0, lateScore: 0,
          noAttend: 0, noAttendScore: 0, noCheckout: 0, noCheckoutScore: 0, totalScore: 0,
        });
      }
      return summary.get(person.name);
    };
    const byName = new Map(people.map(p => [p.name, p]));
    for (const row of lateRows) {
      if (row.lateType === '기준시간없음') continue;
      const s = ensure(byName.get(row.name) || { group: row.group, name: row.name, store: row.store });
      if (row.lateType === '10분 이내') s.late10 += 1;
      else if (row.lateType === '11~59분') s.late11 += 1;
      else if (row.lateType === '60분 이상') s.late60 += 1;
    }
    for (const row of noAttendanceRows) {
      const s = ensure(byName.get(row.name) || { group: row.group, name: row.name, store: row.store });
      s.noAttend += 1;
    }
    for (const row of noCheckoutRows) {
      const s = ensure(byName.get(row.name) || { group: row.group, name: row.name, store: row.store });
      s.noCheckout += 1;
    }
    for (const s of summary.values()) {
      s.lateTotal = s.late10 + s.late11 + s.late60;
      s.lateScore = (s.lateTotal >= 3 ? -s.late10 : 0) - s.late11 - (s.late60 * 2);
      s.noAttendScore = -3 * s.noAttend;
      s.noCheckoutScore = s.noCheckout >= 3 ? -(s.noCheckout - 2) : 0;
      s.totalScore = s.lateScore + s.noAttendScore + s.noCheckoutScore;
    }
    return Array.from(summary.values()).sort((a, b) => a.store.localeCompare(b.store, 'ko') || a.name.localeCompare(b.name, 'ko'));
  }

  function makeWorkbook(result) {
    const wb = XLSX.utils.book_new();
    wb.Workbook = { Views: [{ RTL: false }], CalcPr: { calcMode: 'auto' } };

    addSheet(wb, '사용방법', [
      ['근태 자동 분석 결과 사용방법'],
      ['1', '자동판정_지각 / 자동판정_근태미입력 / 자동판정_퇴근미입력 시트에서 원본 판정 내역을 확인합니다.'],
      ['2', '소명처리 시트에서 적용여부에 “적용”, 처리결과에 “정시인정/출근인정/퇴근인정/교육제외/기타제외” 중 하나를 입력합니다.'],
      ['3', '최종요약 시트는 소명처리의 최종제외 열을 기준으로 감점이 다시 계산됩니다.'],
      ['4', '자동판정 시트는 수정하지 말고, 소명처리 시트만 수정하는 것을 권장합니다.'],
      ['분석월', `${result.year}-${String(result.month).padStart(2, '0')}`],
      ['판정기준일', result.baseDate],
    ], { title: true, widths: [12, 110] });

    addSummarySheet(wb, '자동요약', result.autoSummary, false);
    addFinalSummarySheet(wb, result);
    addLateSheet(wb, result.lateRows);
    addNoAttendanceSheet(wb, result.noAttendanceRows);
    addNoCheckoutSheet(wb, result.noCheckoutRows);
    addMismatchSheet(wb, result.mismatchRows);
    addCeSheet(wb, result.ceRows);
    addExceptionSheet(wb, result.exceptionRows);
    addRuleSheet(wb);

    const fileName = `근태분석결과_${result.year}${String(result.month).padStart(2, '0')}_${result.baseDate.replace(/-/g, '')}.xlsx`;
    XLSX.writeFile(wb, fileName, { bookType: 'xlsx', cellStyles: true });
  }

  function headerStyle() {
    return {
      fill: { fgColor: { rgb: '1F4ED8' } },
      font: { color: { rgb: 'FFFFFF' }, bold: true },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: thinBorder(),
    };
  }

  function thinBorder() {
    return {
      top: { style: 'thin', color: { rgb: 'D9E0EA' } },
      bottom: { style: 'thin', color: { rgb: 'D9E0EA' } },
      left: { style: 'thin', color: { rgb: 'D9E0EA' } },
      right: { style: 'thin', color: { rgb: 'D9E0EA' } },
    };
  }

  function addSheet(wb, name, rows, options = {}) {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
    for (let R = range.s.r; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[addr]) continue;
        ws[addr].s = ws[addr].s || {};
        ws[addr].s.border = thinBorder();
        ws[addr].s.alignment = { vertical: 'center', wrapText: true };
        if (R === 0 || (options.title && R === 0)) ws[addr].s = headerStyle();
      }
    }
    ws['!cols'] = (options.widths || rows[0]?.map(() => 16) || [16]).map(wch => ({ wch }));
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };
    XLSX.utils.book_append_sheet(wb, ws, name);
    return ws;
  }

  function addSummarySheet(wb, name, summaryRows) {
    const rows = [[
      '구분', '이름', '점포', '10분 이내 지각', '11~59분 지각', '60분 이상 지각', '지각 총횟수', '지각감점', '근태미입력', '근태감점', '퇴근미입력', '퇴근감점', '총감점'
    ]];
    for (const s of summaryRows) {
      rows.push([s.group, s.name, s.store, s.late10, s.late11, s.late60, s.lateTotal, s.lateScore, s.noAttend, s.noAttendScore, s.noCheckout, s.noCheckoutScore, s.totalScore]);
    }
    const ws = addSheet(wb, name, rows, { widths: [8, 12, 16, 14, 14, 14, 12, 10, 12, 10, 12, 10, 10] });
    colorNegative(ws, rows.length, 8);
    colorNegative(ws, rows.length, 10);
    colorNegative(ws, rows.length, 12);
    colorNegative(ws, rows.length, 13);
  }

  function addFinalSummarySheet(wb, result) {
    const people = result.autoSummary.length ? result.autoSummary : result.people.map(p => ({ group: p.group, name: p.name, store: p.store }));
    const rows = [[
      '구분', '이름', '점포', '10분 이내 지각', '11~59분 지각', '60분 이상 지각', '지각 총횟수', '지각감점', '근태미입력', '근태감점', '퇴근미입력', '퇴근감점', '총감점'
    ]];
    for (const s of people) rows.push([s.group, s.name, s.store, '', '', '', '', '', '', '', '', '', '']);
    const ws = addSheet(wb, '최종요약', rows, { widths: [8, 12, 16, 14, 14, 14, 12, 10, 12, 10, 12, 10, 10] });

    for (let r = 2; r <= rows.length; r++) {
      ws[`D${r}`] = { t: 'n', f: `COUNTIFS('자동판정_지각'!$B:$B,$B${r},'자동판정_지각'!$J:$J,"10분 이내")-COUNTIFS('소명처리'!$B:$B,"지각",'소명처리'!$C:$C,$B${r},'소명처리'!$F:$F,"10분 이내",'소명처리'!$J:$J,"제외")` };
      ws[`E${r}`] = { t: 'n', f: `COUNTIFS('자동판정_지각'!$B:$B,$B${r},'자동판정_지각'!$J:$J,"11~59분")-COUNTIFS('소명처리'!$B:$B,"지각",'소명처리'!$C:$C,$B${r},'소명처리'!$F:$F,"11~59분",'소명처리'!$J:$J,"제외")` };
      ws[`F${r}`] = { t: 'n', f: `COUNTIFS('자동판정_지각'!$B:$B,$B${r},'자동판정_지각'!$J:$J,"60분 이상")-COUNTIFS('소명처리'!$B:$B,"지각",'소명처리'!$C:$C,$B${r},'소명처리'!$F:$F,"60분 이상",'소명처리'!$J:$J,"제외")` };
      ws[`G${r}`] = { t: 'n', f: `SUM(D${r}:F${r})` };
      ws[`H${r}`] = { t: 'n', f: `IF(G${r}>=3,-D${r},0)-E${r}-F${r}*2` };
      ws[`I${r}`] = { t: 'n', f: `COUNTIFS('자동판정_근태미입력'!$B:$B,$B${r})-COUNTIFS('소명처리'!$B:$B,"근태미입력",'소명처리'!$C:$C,$B${r},'소명처리'!$J:$J,"제외")` };
      ws[`J${r}`] = { t: 'n', f: `-I${r}*3` };
      ws[`K${r}`] = { t: 'n', f: `COUNTIFS('자동판정_퇴근미입력'!$B:$B,$B${r})-COUNTIFS('소명처리'!$B:$B,"퇴근미입력",'소명처리'!$C:$C,$B${r},'소명처리'!$J:$J,"제외")` };
      ws[`L${r}`] = { t: 'n', f: `IF(K${r}>=3,-(K${r}-2),0)` };
      ws[`M${r}`] = { t: 'n', f: `H${r}+J${r}+L${r}` };
    }
    colorNegative(ws, rows.length, 8);
    colorNegative(ws, rows.length, 10);
    colorNegative(ws, rows.length, 12);
    colorNegative(ws, rows.length, 13);
  }

  function addLateSheet(wb, items) {
    const rows = [['구분', '이름', '날짜', '출근지점', '기준점포', '조', '기준시트', '스케줄', '기준시간', '지각구분', '실제출근', '지각분', '기본감점']];
    for (const x of items) rows.push([x.group, x.name, x.date, x.store, x.standardStore || x.store, x.shift, x.basis, x.schedule, x.standardTime, x.lateType, x.actualTime, x.lateMinutes, x.rawScore]);
    addSheet(wb, '자동판정_지각', rows, { widths: [8, 12, 12, 16, 16, 8, 20, 16, 10, 12, 10, 8, 10] });
  }

  function addNoAttendanceSheet(wb, items) {
    const rows = [['구분', '이름', '날짜', '점포', '기준시트', '스케줄', '판정', '기본감점']];
    for (const x of items) rows.push([x.group, x.name, x.date, x.store, x.basis, x.schedule, x.reason, -3]);
    addSheet(wb, '자동판정_근태미입력', rows, { widths: [8, 12, 12, 16, 20, 16, 30, 10] });
  }

  function addNoCheckoutSheet(wb, items) {
    const rows = [['구분', '이름', '날짜', '출근지점', '출근시간', '퇴근시간', '판정']];
    for (const x of items) rows.push([x.group, x.name, x.date, x.store, x.checkIn, x.checkOut, x.reason]);
    addSheet(wb, '자동판정_퇴근미입력', rows, { widths: [8, 12, 12, 16, 10, 10, 32] });
  }

  function addMismatchSheet(wb, items) {
    const rows = [['이름', '날짜', '점포', '매장근무계획상', '웍스스케줄상', '내용']];
    for (const x of items) rows.push([x.name, x.date, x.store, x.plan, x.works, x.note]);
    addSheet(wb, '스케줄불일치', rows, { widths: [12, 12, 16, 18, 18, 50] });
  }

  function addCeSheet(wb, items) {
    const rows = [['구분', '이름', '날짜', '점포', '매장근무계획', '출근확인']];
    for (const x of items) rows.push([x.group, x.name, x.date, x.store, x.schedule, x.status]);
    addSheet(wb, 'CE출근확인', rows, { widths: [8, 12, 12, 16, 18, 12] });
  }

  function addExceptionSheet(wb, items) {
    const rows = [['적용여부', '구분', '이름', '날짜', '자동판정', '지각구분', '지각분', '기본감점', '처리결과', '최종제외', '처리사유', '승인자', '비고']];
    for (const x of items) rows.push([x.apply, x.kind, x.name, x.date, x.autoJudgement, x.lateType, x.lateMinutes, x.baseScore, x.result, '', x.reason, x.approver, x.memo]);
    const ws = addSheet(wb, '소명처리', rows, { widths: [10, 12, 12, 12, 35, 12, 8, 10, 14, 10, 35, 12, 25] });
    for (let r = 2; r <= rows.length; r++) {
      ws[`J${r}`] = {
        t: 's',
        f: `IF(AND($A${r}="적용",OR($I${r}="정시인정",$I${r}="출근인정",$I${r}="퇴근인정",$I${r}="교육제외",$I${r}="기타제외")),"제외","반영")`,
      };
    }
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
      ['MX 순환 기준', '웍스스케줄에 “순환” 포함', '매장근무계획관리의 근무A/B/C 기준 사용'],
      ['교육 기준', '웍스스케줄에 “교육” 포함', '지각/근태 미입력 판정 제외'],
      ['CE 기준', '매장근무계획관리에 “근무” 포함', '출근 여부만 확인'],
      ['소명 처리', '소명처리 시트에서 적용여부=적용, 처리결과 입력', '정시인정/출근인정/퇴근인정/교육제외/기타제외 시 최종요약에서 제외'],
    ];
    addSheet(wb, '판정기준', rows, { widths: [16, 48, 48] });
  }

  function colorNegative(ws, rowCount, colOneBased) {
    for (let r = 2; r <= rowCount; r++) {
      const addr = XLSX.utils.encode_cell({ r: r - 1, c: colOneBased - 1 });
      ws[addr] = ws[addr] || { t: 'n', v: 0 };
      ws[addr].s = ws[addr].s || {};
      ws[addr].s.font = { color: { rgb: 'C92A2A' }, bold: true };
    }
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
  }
})();
