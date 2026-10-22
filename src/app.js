/* 保護者会日程作成ツール：画面 */
(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);

  // ---------- 小さな道具 ----------
  function h(tag, attrs, ...kids) {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v == null || v === false) continue;
      if (k === 'class') el.className = v;
      else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
      else if (k === 'dataset') Object.assign(el.dataset, v);
      else el.setAttribute(k, v === true ? '' : v);
    }
    for (const k of kids.flat()) if (k != null && k !== false) el.append(k instanceof Node ? k : String(k));
    return el;
  }
  function toast(msg, ms = 3500) {
    const t = $('toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toast.timer);
    toast.timer = setTimeout(() => { t.hidden = true; }, ms);
  }
  function download(blob, name) {
    const a = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.append(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }
  function readSheets(buf) {
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    const out = {};
    for (const n of wb.SheetNames) out[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null });
    out._first = wb.SheetNames[0];
    return out;
  }
  function sheetsToBlob(sheets) {
    const wb = XLSX.utils.book_new();
    const widths = { 日程: [7, 5, 7, 7, 9, 18, 40, 20], 情報: [16, 22] };
    for (const [name, rows] of Object.entries(sheets)) {
      const ws = XLSX.utils.aoa_to_sheet(rows);
      ws['!cols'] = (widths[name] || rows[0].map((_, i) => (i < 3 ? 12 : i === rows[0].length - 1 ? 30 : 16))).map((w) => ({ wch: w }));
      XLSX.utils.book_append_sheet(wb, ws, name);
    }
    const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  function setupDrop(label, input, onFiles) {
    input.addEventListener('change', () => { if (input.files.length) onFiles([...input.files]); input.value = ''; });
    label.addEventListener('dragover', (e) => { e.preventDefault(); label.classList.add('over'); });
    label.addEventListener('dragleave', () => label.classList.remove('over'));
    label.addEventListener('drop', (e) => {
      e.preventDefault();
      label.classList.remove('over');
      onFiles([...e.dataTransfer.files], e.dataTransfer.items);
    });
  }
  const classTitle = (g, c) => `${g}年${c}組`;
  const fileNameOf = (g, c) => `日程案_${g}年${c}組.xlsx`;

  // ---------- 配布物の文面 ----------
  const LETTER_KEY = 'hogosyakai.letter';
  function storedLetter() {
    try {
      const v = JSON.parse(localStorage.getItem(LETTER_KEY) || 'null');
      if (v) return { ...Core.DEFAULT_LETTER, ...v };
    } catch (e) { /* 記憶できない環境でも動く */ }
    return { ...Core.DEFAULT_LETTER };
  }
  function storeLetter(letter) {
    try { localStorage.setItem(LETTER_KEY, JSON.stringify(letter)); } catch (e) { /* なし */ }
  }

  // 1枚分のお知らせ
  function letterPage(letter, p) {
    const date = letter.date.trim() || Core.reiwaDate(new Date());
    return h('div', { class: 'letter-page' },
      h('div', { class: 'r lp-date' }, date),
      h('div', { class: 'lp-to' }, letter.to),
      h('div', { class: 'r lp-school' }, letter.school),
      h('div', { class: 'r' }, letter.principal),
      h('div', { class: 'c lp-title' }, letter.title),
      h('div', { class: 'lp-body' }, letter.body),
      h('div', { class: 'c lp-ki' }, '記'),
      h('div', {}, letter.item1),
      h('div', { class: 'lp-box' },
        h('div', { class: 'lp-who' }, h('span', {}, `${p.grade}　年`), h('span', {}, `${p.cls}　組`), h('span', {}, `生徒氏名　　${p.name}`)),
        h('div', {}, '〔懇談予定日時〕'),
        h('div', { class: 'lp-time' }, p.time),
        h('div', { class: 'lp-note' }, letter.note)),
      h('div', { class: 'lp-items' }, letter.items));
  }

  /**
   * 文面の入力欄と見本を作る
   * @param box 入れる場所
   * @param get 今の文面を返す関数
   * @param onChange 文面が変わったとき（新しい文面）
   * @param sample 見本に使う生徒 {grade, cls, name, slot}
   */
  function letterEditor(box, get, onChange, sample) {
    box.replaceChildren();
    const letter = get();
    const pv = h('div', { class: 'preview-wrap' });
    const redraw = () => {
      const l = get();
      const sm = sample();
      pv.replaceChildren(h('div', { class: 'preview-cap' }, '見本（1人目）'),
        letterPage(l, { grade: sm.grade, cls: sm.cls, name: sm.name, time: sm.slot ? Core.letterTime(sm.slot, l.timeStyle) : Core.toFull('11月10日（火）　13:30～') }));
    };
    const els = {};
    const field = (key, label, opts = {}) => {
      const el = opts.rows
        ? h('textarea', { rows: opts.rows }, letter[key])
        : h('input', { type: 'text', value: letter[key], placeholder: opts.placeholder || '' });
      if (opts.rows) el.value = letter[key];
      els[key] = el;
      el.addEventListener('input', () => { onChange({ ...get(), [key]: el.value }); redraw(); syncSeason(); });
      return h('label', {}, label, el);
    };
    // 時候のあいさつ：選ぶと前文の先頭の「〇〇の候」が入れ替わる
    const sug = Core.suggestSeason(new Date());
    const seasonSel = h('select', {},
      h('option', { value: '' }, '（前文のまま）'),
      Core.SEASON_GREETINGS.map((g) => h('optgroup', { label: g.when + (g === sug ? '　← 今の時期' : '') }, g.items.map((x) => h('option', { value: x }, x)))));
    const syncSeason = () => {
      const cur = Core.getSeason(get().body);
      seasonSel.value = Core.SEASON_GREETINGS.some((g) => g.items.includes(cur)) ? cur : '';
    };
    seasonSel.addEventListener('change', () => {
      if (!seasonSel.value) return;
      const body = Core.setSeason(get().body, seasonSel.value);
      els.body.value = body;
      onChange({ ...get(), body });
      redraw();
    });
    const timeSel = h('select', {},
      h('option', { value: 'start', selected: letter.timeStyle !== 'range' }, '開始時刻だけ（例：１３：３０～）'),
      h('option', { value: 'range', selected: letter.timeStyle === 'range' }, '開始～終了（例：１３：３０～１３：４０）'));
    timeSel.addEventListener('change', () => { onChange({ ...get(), timeStyle: timeSel.value }); redraw(); });
    const form = h('div', { class: 'letter-form' },
      field('date', '発行日', { placeholder: '空欄なら印刷した日（例：令和８年１１月２日）' }),
      field('to', '宛名'),
      field('school', '学校名'),
      field('principal', '校長名'),
      field('title', '表題'),
      h('label', {}, '時候のあいさつ（選ぶと前文の最初が入れ替わります）', seasonSel),
      field('body', '前文', { rows: 7 }),
      field('item1', '１の見出し'),
      h('label', {}, '時刻の書き方', timeSel),
      field('note', '日時の下の注意書き'),
      field('items', '２以降の項目', { rows: 9 }),
      h('button', { class: 'btn small', onclick: () => { if (confirm('最初の文面に戻しますか？')) { onChange({ ...Core.DEFAULT_LETTER }); letterEditor(box, get, onChange, sample); } } }, '最初の文面に戻す'));
    box.append(h('div', { class: 'letter-edit' }, form, pv));
    syncSeason();
    redraw();
  }

  // ---------- 学級の日程表（A4・1枚） ----------
  /**
   * @param c {grade, cls, slots, unassigned}
   * @param who sid -> {no, name, sib}
   */
  function schedulePage(c, who) {
    const days = [...new Set(c.slots.map((x) => x.dateKey))].sort();
    const times = [...new Set(c.slots.map((x) => x.start))].sort((a, b) => a - b);
    const at = new Map(c.slots.map((x) => [`${x.dateKey} ${x.start}`, x]));
    const first = (dk) => c.slots.find((x) => x.dateKey === dk);
    let sibUsed = false;
    const cell = (sl) => {
      if (!sl) return h('td', { class: 'none' }, '－');
      if (sl.kind === 'break') return h('td', { class: 'brk' }, '休憩');
      if (sl.kind === 'blocked') return h('td', { class: 'brk' }, '担任不可');
      if (!sl.sid) return h('td', {});
      const w = who(sl.sid);
      if (w.sib) sibUsed = true;
      return h('td', {}, h('span', { class: 'sno' }, w.no), w.name, w.sib ? h('span', { class: 'smark' }, '◆') : null);
    };
    // 行が多いときは文字を小さくして1枚に収める
    const size = times.length <= 16 ? 'l' : times.length <= 24 ? 'm' : 's';
    const total = c.slots.filter((x) => x.sid).length;
    return h('div', { class: `sched-page size-${size}` },
      h('div', { class: 'sp-head' }, h('div', { class: 'sp-title' }, `${c.grade}年${c.cls}組　個人懇談会　日程表`),
        h('div', { class: 'sp-sub' }, `${total}人　${new Date().toLocaleDateString('ja-JP')} 現在`)),
      h('table', { class: 'sp-table' },
        h('thead', {}, h('tr', {}, h('th', { class: 'tcol' }, '時刻'),
          days.map((dk) => { const f = first(dk); return h('th', {}, `${f.month}/${f.day}（${f.dow}）`, h('br'), h('span', { class: 'cnt' }, `${c.slots.filter((x) => x.dateKey === dk && x.sid).length}人`)); }))),
        h('tbody', {}, times.map((t) => h('tr', {}, h('th', { class: 'tcol' }, Core.toHM(t)), days.map((dk) => cell(at.get(`${dk} ${t}`))))))),
      h('div', { class: 'sp-foot' },
        c.unassigned.length ? h('div', {}, '未割当：', c.unassigned.map((sid) => { const w = who(sid); return `${w.no}番 ${w.name}`; }).join('、')) : null,
        sibUsed ? h('div', {}, '◆ 兄弟姉妹あり（相手の学級と続けて面談）') : null));
  }

  function printPages(pages) {
    $('printArea').replaceChildren(...pages);
    document.body.classList.add('print-letters');
    window.print();
  }

  // ---------- タブ ----------
  function showTab(name) {
    document.querySelectorAll('.tab').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === name)));
    document.querySelectorAll('.panel').forEach((p) => { p.hidden = p.id !== name; });
    try { localStorage.setItem('hogosyakai.tab', name); } catch (e) { /* 保存できなくても動く */ }
  }
  document.querySelectorAll('.tab').forEach((b) => b.addEventListener('click', () => showTab(b.dataset.tab)));
  let firstTab = 'admin';
  try { firstTab = localStorage.getItem('hogosyakai.tab') || 'admin'; } catch (e) { /* なし */ }
  showTab(firstTab);

  // =====================================================================
  // ① 担当者
  // =====================================================================
  const A = {
    files: [], // {name, rows, grade}
    parsed: null, // {students, bandsByGrade, issues}
    blocked: new Map(), // classKey -> Set(bandKey)
    manual: new Map(), // 照合できなかった問題のキー -> 相手のid（'' = つながない）
    removed: new Set(), // 解除したリンク "a|b"
    result: null,
    letter: storedLetter(),
  };

  setupDrop($('adminDrop'), $('adminFiles'), async (files) => {
    for (const f of files) {
      if (!/\.xlsx$/i.test(f.name)) { toast(`${f.name} はExcel（.xlsx）ではありません`); continue; }
      try {
        const sheets = readSheets(await f.arrayBuffer());
        const rows = sheets[sheets._first];
        const grade = Core.guessGrade(f.name);
        const i = A.files.findIndex((x) => x.name === f.name);
        const entry = { name: f.name, rows, grade };
        if (i >= 0) A.files[i] = entry; else A.files.push(entry);
      } catch (e) {
        toast(`${f.name} を読み取れませんでした：${e.message}`);
      }
    }
    analyze();
  });

  function analyze() {
    A.result = null;
    $('result').replaceChildren();
    const students = [];
    const bandsByGrade = new Map();
    const issues = [];
    const errors = [];
    const fileRows = [];
    for (const f of A.files) {
      let info = '';
      if (!f.grade) {
        info = '← 学年を選んでください';
      } else {
        try {
          const r = Core.parseGradeRows(f.rows, f.grade, f.name);
          const dupGrade = students.some((s) => s.grade === f.grade);
          if (dupGrade) errors.push(`${f.grade}年のファイルが2つあります（${f.name}）`);
          students.push(...r.students);
          bandsByGrade.set(f.grade, r.bands);
          issues.push(...r.issues);
          const classes = [...new Set(r.students.map((s) => s.cls))].sort();
          info = `${r.students.length}人（${classes.map((c) => c + '組').join('・')}）` + (r.hasSibCol ? '' : '　※兄弟姉妹の列がありません');
        } catch (e) {
          errors.push(e.message);
          info = '読み取れません';
        }
      }
      fileRows.push({ f, info });
    }
    renderFileList(fileRows, errors);
    if (!students.length) {
      ['stepSettings', 'stepCheck', 'stepLetter', 'stepRun'].forEach((id) => { $(id).hidden = true; });
      return;
    }
    A.parsed = { students, bandsByGrade, issues };
    A.sib = Core.matchSiblings(students);
    ['stepSettings', 'stepCheck', 'stepLetter', 'stepRun'].forEach((id) => { $(id).hidden = false; });
    renderBlockedGrid();
    renderCheck();
    if (!$('adminLetter').firstChild) {
      letterEditor($('adminLetter'), () => A.letter, (l) => { A.letter = l; storeLetter(l); }, () => {
        const s = A.parsed.students[0];
        return { grade: s.grade, cls: s.cls, name: s.name, slot: null };
      });
    }
  }

  function renderFileList(fileRows, errors) {
    const box = $('fileList');
    box.replaceChildren();
    if (!fileRows.length) return;
    const tbl = h('table', { class: 'tbl' }, h('tr', {}, h('th', {}, 'ファイル'), h('th', {}, '学年'), h('th', {}, '内容'), h('th', {}, '')));
    for (const { f, info } of fileRows) {
      const sel = h('select', { 'aria-label': '学年', onchange: (e) => { f.grade = +e.target.value || null; analyze(); } },
        h('option', { value: '' }, '選ぶ'),
        [1, 2, 3, 4, 5, 6].map((g) => h('option', { value: g, selected: f.grade === g }, `${g}年`)));
      tbl.append(h('tr', {}, h('td', {}, f.name), h('td', {}, sel), h('td', {}, info),
        h('td', {}, h('button', { class: 'btn small', onclick: () => { A.files = A.files.filter((x) => x !== f); analyze(); } }, '外す'))));
    }
    box.append(tbl);
    if (errors.length) box.append(h('div', { class: 'box bad' }, h('ul', {}, errors.map((e) => h('li', {}, e)))));
  }

  function classKeysOf(grade) {
    return [...new Set(A.parsed.students.filter((s) => s.grade === grade).map((s) => s.cls))].sort();
  }

  function renderBlockedGrid() {
    const box = $('blockedGrid');
    box.replaceChildren();
    for (const [grade, bands] of [...A.parsed.bandsByGrade].sort((a, b) => a[0] - b[0])) {
      const sorted = [...bands].sort((a, b) => a.key.localeCompare(b.key));
      const head = h('tr', {}, h('th', {}, `${grade}年`), sorted.map((b) => h('th', {}, `${b.month}/${b.day}(${b.dow})`, h('br'), `${b.start}～`)));
      const rows = classKeysOf(grade).map((cls) => {
        const ck = `${grade}-${cls}`;
        const set = A.blocked.get(ck) || new Set();
        return h('tr', {}, h('th', {}, `${cls}組`), sorted.map((b) => {
          const on = set.has(b.key);
          return h('td', {
            class: 'cell' + (on ? ' on' : ''), title: `${cls}組 ${b.label}`,
            onclick: (e) => {
              if (!A.blocked.has(ck)) A.blocked.set(ck, new Set());
              const s = A.blocked.get(ck);
              if (s.has(b.key)) s.delete(b.key); else s.add(b.key);
              e.target.classList.toggle('on');
              e.target.textContent = s.has(b.key) ? '不可' : '';
              A.result = null;
            },
          }, on ? '不可' : '');
        }));
      });
      box.append(h('div', { class: 'grid-wrap' }, h('table', { class: 'bgrid' }, head, rows)));
    }
  }

  const linkKey = (a, b) => [a, b].sort().join('|');
  const probKey = (p) => `${p.id}|${p.entry.raw}`;

  function currentLinks() {
    const links = A.sib.links.filter(([a, b]) => !A.removed.has(linkKey(a, b)));
    for (const p of A.sib.problems) {
      const t = A.manual.get(probKey(p));
      if (t) links.push([p.id, t]);
    }
    return links;
  }

  function renderCheck() {
    const byId = new Map(A.parsed.students.map((s) => [s.id, s]));
    // 照合できなかった兄弟姉妹
    const pbox = $('sibProblems');
    pbox.replaceChildren();
    if (A.sib.problems.length) {
      const wrap = h('div', { class: 'box warn' }, h('h4', {}, `兄弟姉妹の欄を照合できなかった家庭（${A.sib.problems.length}件）`),
        h('p', { class: 'hint' }, '相手が分かる場合は選んでください。他校のきょうだいなどは「つなげない」のままで大丈夫です。'));
      for (const p of A.sib.problems) {
        const s = byId.get(p.id);
        const pool = A.parsed.students.filter((x) => x.id !== s.id && (!p.entry.grade || x.grade === p.entry.grade));
        const sel = h('select', { onchange: (e) => { A.manual.set(probKey(p), e.target.value); A.result = null; } },
          h('option', { value: '' }, 'つなげない'),
          pool.map((x) => h('option', { value: x.id, selected: A.manual.get(probKey(p)) === x.id }, Core.label(x))));
        wrap.append(h('div', { class: 'sibrow' }, h('span', { class: 'who' }, Core.label(s)), h('span', { class: 'raw' }, `記入：「${p.raw}」（${p.reason}）`), sel));
      }
      pbox.append(wrap);
    } else if (A.parsed.students.some((s) => s.sibText)) {
      pbox.append(h('div', { class: 'box ok' }, '兄弟姉妹の欄はすべて照合できました。'));
    }
    // 見つかった兄弟姉妹
    const lbox = $('sibLinks');
    lbox.replaceChildren();
    $('sibLinksSummary').textContent = `見つかった兄弟姉妹（${A.sib.links.length}組）`;
    for (const [a, b] of A.sib.links) {
      const k = linkKey(a, b);
      const cb = h('input', { type: 'checkbox', checked: !A.removed.has(k), onchange: (e) => { if (e.target.checked) A.removed.delete(k); else A.removed.add(k); A.result = null; } });
      lbox.append(h('label', { class: 'sibrow' }, cb, h('span', {}, `${Core.label(byId.get(a))} ⇔ ${Core.label(byId.get(b))}`)));
    }
    if (!A.sib.links.length) lbox.append(h('p', { class: 'hint' }, 'ありません'));
    // データの注意点
    $('issuesSummary').textContent = `データの注意点（${A.parsed.issues.length}件）`;
    $('issues').replaceChildren(...A.parsed.issues.map((i) => h('li', {}, i.text)));
    $('issuesBox').open = A.parsed.issues.some((i) => i.type === 'duplicate');
  }

  function options() {
    const minutes = +$('optMinutes').value;
    return {
      minutes,
      breakEvery: +$('optBreak').value,
      sibGap: $('optGap').value === 'slot' ? minutes : 0,
      spread: $('optSpread').value === '1',
      blockedByClass: A.blocked,
      links: currentLinks(),
    };
  }
  ['optMinutes', 'optBreak', 'optGap', 'optSpread'].forEach((id) => $(id).addEventListener('change', () => { A.result = null; $('result').replaceChildren(); }));

  $('runBtn').addEventListener('click', () => {
    const opt = options();
    const res = Core.schedule(A.parsed.students, A.parsed.bandsByGrade, opt);
    A.result = { res, opt, createdAt: new Date().toLocaleString('ja-JP') };
    renderResult();
  });

  // 兄弟姉妹欄（日程案に書く文字）
  function siblingInfo(res) {
    const info = new Map();
    const slotOf = (s) => res.classes.get(`${s.grade}-${s.cls}`).slots.find((x) => x.sid === s.id);
    for (const fam of res.families) {
      for (const s of fam) {
        const others = fam.filter((x) => x !== s).map((x) => {
          const sl = slotOf(x);
          return `${Core.label(x)}（${sl ? `${sl.month}/${sl.day} ${Core.toHM(sl.start)}～${Core.toHM(sl.end)}` : '未割当'}）`;
        });
        info.set(s.id, others.join('／'));
      }
    }
    for (const p of A.sib.problems) {
      if (A.manual.get(probKey(p))) continue;
      const prev = info.get(p.id);
      info.set(p.id, (prev ? prev + '／' : '') + `要確認「${p.raw}」`);
    }
    return info;
  }

  function classSheets(ck) {
    const { res, opt, createdAt } = A.result;
    const cls = res.classes.get(ck);
    return Core.classToSheets(cls, res.byId, A.parsed.bandsByGrade.get(cls.grade), siblingInfo(res), { minutes: opt.minutes, createdAt, letter: A.letter });
  }

  function renderResult() {
    const { res } = A.result;
    const box = $('result');
    box.replaceChildren();
    const keys = [...res.classes.keys()].sort((a, b) => { const [ga, ca] = a.split('-'); const [gb, cb] = b.split('-'); return ga - gb || ca.localeCompare(cb); });
    const tbl = h('table', { class: 'tbl' }, h('tr', {}, h('th', {}, '学級'), h('th', {}, '割り当て'), h('th', {}, '未割当'), h('th', {}, '兄弟姉妹'), h('th', {}, '')));
    for (const ck of keys) {
      const c = res.classes.get(ck);
      const n = c.slots.filter((x) => x.sid).length;
      const sib = res.families.flat().filter((s) => `${s.grade}-${s.cls}` === ck).length;
      tbl.append(h('tr', {}, h('td', {}, classTitle(c.grade, c.cls)), h('td', {}, `${n}人`),
        h('td', { style: c.unassigned.length ? 'color:var(--bad);font-weight:bold' : '' }, `${c.unassigned.length}人`),
        h('td', {}, sib ? `${sib}人` : ''),
        h('td', {},
          h('button', { class: 'btn small', onclick: () => download(sheetsToBlob(classSheets(ck)), fileNameOf(c.grade, c.cls)) }, '保存'), ' ',
          h('button', { class: 'btn small', onclick: () => printSchedules([ck]) }, '日程表を印刷'), ' ',
          h('button', { class: 'btn small', onclick: () => { T.load(classSheets(ck), fileNameOf(c.grade, c.cls), null); showTab('teacher'); } }, '開いて確認'))));
    }
    const notes = res.notes;
    box.append(
      h('div', { class: 'box ok' }, `${keys.length}学級の日程案ができました。下のボタンで学級ごとのExcelを保存し、共有フォルダに置いてください。`),
      h('div', { class: 'actions' },
        h('button', { class: 'btn primary', onclick: () => saveAllDownloads(keys) }, '「ダウンロード」フォルダに全学級を保存'),
        'showDirectoryPicker' in window ? h('button', { class: 'btn', onclick: () => saveAllToFolder(keys) }, '保存先フォルダを選んで保存') : null,
        h('button', { class: 'btn', onclick: () => saveAllZip(keys) }, 'ZIPにまとめてダウンロード'),
        h('button', { class: 'btn', onclick: () => printSchedules(keys) }, '全学級の日程表を印刷（1学級1枚）')),
      h('p', { class: 'hint' }, '「ダウンロード」に保存したときに「複数ファイルのダウンロード」の確認が出たら「許可」を押してください。'
        + '「保存先フォルダを選んで保存」では、「ダウンロード」「デスクトップ」「ドキュメント」そのものは選べません。共有フォルダか、その中に作ったフォルダ（例：保護者会）を選んでください。'),
      tbl,
      notes.length ? h('div', { class: 'box warn' }, h('h4', {}, `担任に伝えること（${notes.length}件）`), h('ul', {}, notes.map((n) => h('li', {}, n.text)))) : null,
    );
  }

  async function saveAllToFolder(keys) {
    let dir;
    try { dir = await window.showDirectoryPicker({ mode: 'readwrite', startIn: 'documents' }); } catch (e) { return; }
    try {
      for (const ck of keys) {
        const c = A.result.res.classes.get(ck);
        const fh = await dir.getFileHandle(fileNameOf(c.grade, c.cls), { create: true });
        const w = await fh.createWritable();
        await w.write(sheetsToBlob(classSheets(ck)));
        await w.close();
      }
      toast(`${keys.length}個のファイルを保存しました`);
    } catch (e) {
      toast(`保存できませんでした：${e.message}。ZIPでダウンロードしてください`, 6000);
    }
  }

  function printSchedules(keys) {
    const { res } = A.result;
    const sibs = new Set(res.families.flat().map((s) => s.id));
    const who = (sid) => { const s = res.byId.get(sid); return { no: s.no, name: s.name, sib: sibs.has(sid) }; };
    printPages(keys.map((ck) => schedulePage(res.classes.get(ck), who)));
  }

  // 1学級ずつ普通にダウンロードする（いちばん確実）
  async function saveAllDownloads(keys) {
    for (const ck of keys) {
      const c = A.result.res.classes.get(ck);
      download(sheetsToBlob(classSheets(ck)), fileNameOf(c.grade, c.cls));
      await new Promise((r) => setTimeout(r, 400));
    }
    toast(`${keys.length}個のファイルを「ダウンロード」フォルダに保存しました。共有フォルダに移してください`, 6000);
  }

  async function saveAllZip(keys) {
    const zip = new JSZip();
    for (const ck of keys) {
      const c = A.result.res.classes.get(ck);
      zip.file(fileNameOf(c.grade, c.cls), sheetsToBlob(classSheets(ck)));
    }
    download(await zip.generateAsync({ type: 'blob' }), '日程案_全学級.zip');
  }

  // =====================================================================
  // ② 担任
  // =====================================================================
  const T = {
    data: null, // Core.sheetsToClass の結果
    fileName: '',
    handle: null, // 上書き保存用（Edge/Chrome）
    selected: null, // 選択中の生徒id
    original: new Map(), // sid -> 開いたときの枠のキー
    template: null,
    dirty: false,
    letter: null,
  };
  const slotKey = (sl) => `${sl.dateKey} ${sl.start}`;

  T.load = function (sheets, fileName, handle) {
    let data;
    try { data = Core.sheetsToClass(sheets); } catch (e) { toast(e.message, 6000); return; }
    T.data = data;
    T.fileName = fileName;
    T.handle = handle;
    T.selected = null;
    T.dirty = false;
    T.original = new Map(data.slots.filter((s) => s.sid).map((s) => [s.sid, slotKey(s)]));
    $('boardStep').hidden = false;
    $('docStep').hidden = false;
    $('boardTitle').textContent = `${classTitle(data.grade, data.cls)}の日程（${fileName}）`;
    $('saveMsg').textContent = '';
    $('letterMsg').textContent = '';
    T.letter = data.letter || storedLetter();
    letterEditor($('teacherLetter'), () => T.letter, (l) => { T.letter = l; T.dirty = true; }, () => {
      const first = placedList()[0];
      return first ? { grade: first.s.grade, cls: first.s.cls, name: first.s.name, slot: first.sl } : { grade: data.grade, cls: data.cls, name: '（氏名）', slot: null };
    });
    renderBoard();
  };

  function placedList() {
    return T.data.slots.filter((x) => x.sid).map((sl) => ({ sl, s: studentOf(sl.sid) })).sort((a, b) => a.s.no - b.s.no);
  }

  async function openTeacherFile(files, items) {
    const f = files[0];
    if (!f) return;
    if (!/\.xlsx$/i.test(f.name)) { toast('日程案のExcel（.xlsx）を選んでください'); return; }
    let handle = null;
    try {
      if (items && items[0] && items[0].getAsFileSystemHandle) handle = await items[0].getAsFileSystemHandle();
    } catch (e) { handle = null; }
    try { T.load(readSheets(await f.arrayBuffer()), f.name, handle); } catch (e) { toast(`読み取れませんでした：${e.message}`, 6000); }
  }
  setupDrop($('teacherDrop'), $('teacherFile'), openTeacherFile);
  // Edge/Chrome ではファイルを開いた場所にそのまま上書き保存できるようにする
  if ('showOpenFilePicker' in window) {
    $('teacherDrop').addEventListener('click', async (e) => {
      e.preventDefault();
      try {
        const [hd] = await window.showOpenFilePicker({ types: [{ description: 'Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }] });
        const f = await hd.getFile();
        T.load(readSheets(await f.arrayBuffer()), f.name, hd);
      } catch (err) {
        if (err.name !== 'AbortError') $('teacherFile').click();
      }
    });
  }
  window.addEventListener('beforeunload', (e) => { if (T.dirty) { e.preventDefault(); e.returnValue = ''; } });

  function studentOf(sid) { return T.data.students.get(sid); }
  function isNg(sid, slot) {
    const ng = studentOf(sid).ng;
    return Core.bandsOfSlot(T.data.bands, slot).keys.some((k) => ng.includes(k));
  }
  function isOutside(slot) { return Core.bandsOfSlot(T.data.bands, slot).outside; }
  function timeChanged(sid, slot) { const o = T.original.get(sid); return !!o && o !== slotKey(slot); }
  const slotLen = (sl) => +T.data.minutes || sl.end - sl.start;

  // 枠・休憩の追加と削除（後ろの枠は時刻がずれる）
  function addSlot(ref, kind, dateKey) {
    const d = T.data;
    const len = slotLen(ref || d.slots.find((x) => x.dateKey === dateKey));
    Core.insertSlot(d.slots, d.bands, ref, kind, len, dateKey);
    T.dirty = true;
    renderBoard();
  }
  function deleteSlot(sl) {
    const d = T.data;
    if (sl.sid) {
      if (!confirm(`${Core.label(studentOf(sl.sid))} の枠を消します。この生徒は「未割当」に移ります。よろしいですか？`)) return;
      d.unassigned.push(sl.sid);
    }
    Core.removeSlot(d.slots, d.bands, sl);
    T.dirty = true;
    renderBoard();
  }
  function rowTools(sl, ...extra) {
    const btn = (label, title, fn) => h('button', { class: 'btn small', title, onclick: (e) => { e.stopPropagation(); fn(); } }, label);
    const len = slotLen(sl);
    return h('span', { class: 'mini' },
      ...extra.map(([label, fn]) => btn(label, '', fn)),
      btn('上に枠', `この上に面談枠を1つ入れます（下の枠は${len}分ずつ後ろにずれます）`, () => addSlot(sl, 'free')),
      btn('上に休憩', `この上に休憩を1つ入れます（下の枠は${len}分ずつ後ろにずれます）`, () => addSlot(sl, 'break')),
      btn('削除', `この枠を消します（下の枠は${sl.end - sl.start}分ずつ前に詰まります）`, () => deleteSlot(sl)));
  }

  function card(sid, slot) {
    const s = studentOf(sid);
    const ng = slot && isNg(sid, slot);
    const out = slot && !ng && isOutside(slot);
    const moved = slot && timeChanged(sid, slot);
    const sib = T.data.sibInfo.get(sid);
    const c = h('div', {
      class: 'card' + (T.selected === sid ? ' sel' : '') + (ng ? ' ng' : ''),
      draggable: 'true',
      title: [Core.label(s), sib ? `兄弟姉妹：${sib}` : '', ng ? 'この時間は×（都合が悪い）です' : ''].filter(Boolean).join('\n'),
      onclick: (e) => { e.stopPropagation(); if (T.selected && T.selected !== sid && slot) { moveTo(slot); return; } T.selected = T.selected === sid ? null : sid; renderBoard(); },
      ondragstart: (e) => { T.selected = sid; e.dataTransfer.setData('text/plain', sid); e.dataTransfer.effectAllowed = 'move'; },
    }, h('span', { class: 'no' }, `${s.no}`), h('span', { class: 'nm' }, s.name),
    sib ? h('span', { class: 'badge sib' }, '兄弟') : null,
    ng ? h('span', { class: 'badge ng' }, '×') : null,
    out ? h('span', { class: 'badge out' }, '時間外') : null,
    moved ? h('span', { class: 'badge moved', title: `元の時刻：${T.original.get(sid).split(' ')[0]} ${Core.toHM(+T.original.get(sid).split(' ')[1])}` }, '時刻変更') : null);
    return c;
  }

  function dropTarget(el, onDrop) {
    el.addEventListener('dragover', (e) => { e.preventDefault(); el.classList.add('over'); });
    el.addEventListener('dragleave', () => el.classList.remove('over'));
    el.addEventListener('drop', (e) => { e.preventDefault(); el.classList.remove('over'); T.selected = e.dataTransfer.getData('text/plain') || T.selected; onDrop(); });
    el.addEventListener('click', () => { if (T.selected) onDrop(); });
  }

  // 選択中の生徒を slot へ（入っていれば入れ替え）
  function moveTo(slot) {
    const sid = T.selected;
    if (!sid) return;
    const d = T.data;
    const from = d.slots.find((x) => x.sid === sid);
    if (from === slot) { T.selected = null; renderBoard(); return; }
    const other = slot.sid;
    if (from) from.sid = other; // 入れ替え（空なら空になる）
    else {
      d.unassigned = d.unassigned.filter((x) => x !== sid);
      if (other) d.unassigned.push(other);
    }
    slot.sid = sid;
    T.selected = null;
    T.dirty = true;
    renderBoard();
  }
  function moveToTray() {
    const sid = T.selected;
    if (!sid) return;
    const from = T.data.slots.find((x) => x.sid === sid);
    if (from) { from.sid = null; T.data.unassigned.push(sid); T.dirty = true; }
    T.selected = null;
    renderBoard();
  }

  function renderBoard() {
    const d = T.data;
    // 未割当
    const tray = $('tray');
    tray.classList.toggle('target', !!T.selected);
    $('trayBody').replaceChildren(...d.unassigned.slice().sort((a, b) => studentOf(a).no - studentOf(b).no).map((sid) => card(sid, null)));
    if (!d.unassigned.length) $('trayBody').append(h('span', { class: 'hint' }, 'なし'));
    if (!tray.dataset.ready) { dropTarget(tray, moveToTray); tray.dataset.ready = '1'; }
    // 日ごとの列
    const board = $('board');
    board.replaceChildren();
    const byDate = new Map();
    for (const sl of d.slots) { if (!byDate.has(sl.dateKey)) byDate.set(sl.dateKey, []); byDate.get(sl.dateKey).push(sl); }
    for (const [, list] of [...byDate].sort()) {
      const first = list[0];
      const last = list[list.length - 1];
      const col = h('div', { class: 'day' }, h('h3', {}, Core.slotDateText(first), h('span', { class: 'hint' }, ` ${list.filter((x) => x.sid).length}人・${Core.toHM(last.end)}まで`)));
      let prevEnd = null;
      for (const sl of list) {
        if (prevEnd != null && sl.start > prevEnd) col.append(h('div', { class: 'slot break' }, h('span', { class: 't' }, ''), `（${sl.start - prevEnd}分 あき）`));
        prevEnd = sl.end;
        const t = h('span', { class: 't' }, Core.toHM(sl.start));
        let row;
        if (sl.kind !== 'free') {
          row = h('div', { class: 'slot ' + sl.kind }, t, sl.kind === 'break' ? '休憩' : '担任不可',
            rowTools(sl, ['面談枠にする', () => { sl.kind = 'free'; T.dirty = true; renderBoard(); }]));
        } else if (sl.sid) {
          row = h('div', { class: 'slot' + (T.selected && T.selected !== sl.sid ? ' target' : '') }, t, card(sl.sid, sl), rowTools(sl));
          dropTarget(row, () => moveTo(sl));
        } else {
          row = h('div', { class: 'slot' + (T.selected ? ' target' : '') }, t, h('span', { class: 'empty' }, '空き'),
            rowTools(sl, ['休憩にする', () => { sl.kind = 'break'; T.dirty = true; renderBoard(); }]));
          dropTarget(row, () => moveTo(sl));
        }
        col.append(row);
      }
      col.append(h('div', { class: 'day-add' },
        h('button', { class: 'btn small', onclick: () => addSlot(null, 'free', first.dateKey) }, '＋ 最後に枠'),
        h('button', { class: 'btn small', onclick: () => addSlot(null, 'break', first.dateKey) }, '＋ 最後に休憩')));
      board.append(col);
    }
    renderWarnings();
  }

  function renderWarnings() {
    const d = T.data;
    const bad = [];
    const warn = [];
    for (const sl of d.slots) {
      if (!sl.sid) continue;
      const who = Core.label(studentOf(sl.sid));
      const when = `${Core.slotDateText(sl)}${Core.slotTimeText(sl)}`;
      if (isNg(sl.sid, sl)) bad.push(`${who}：${when} は都合が悪い（×）時間です${timeChanged(sl.sid, sl) ? '（枠を足したり消したりして時刻がずれました）' : ''}`);
      else if (isOutside(sl)) warn.push(`${who}：${when} は希望調査の時間帯の外です。保護者の都合を確認してください`);
    }
    const changed = d.slots.filter((sl) => sl.sid && timeChanged(sl.sid, sl));
    if (changed.length) {
      warn.push(`開いたときから時刻が変わった家庭が${changed.length}件あります（すでに配布物を配っていれば連絡が必要です）：`
        + changed.map((sl) => { const [dk, st] = T.original.get(sl.sid).split(' '); return `${studentOf(sl.sid).no}番 ${dk.replace(/^0/, '').replace('/0', '/')} ${Core.toHM(+st)}→${sl.month}/${sl.day} ${Core.toHM(sl.start)}`; }).join('、'));
    }
    for (const p of d.problems) bad.push(p);
    for (const sid of d.unassigned) bad.push(`${Core.label(studentOf(sid))}：まだ日時が決まっていません（未割当）`);
    for (const [sid, info] of d.sibInfo) {
      const now = d.slots.find((x) => x.sid === sid);
      const moved = (now ? slotKey(now) : null) !== (T.original.get(sid) || null);
      if (moved && !info.startsWith('要確認')) warn.push(`${Core.label(studentOf(sid))} を動かしました。兄弟姉妹（${info}）の時間とずれていないか、相手の担任に連絡してください`);
      if (info.includes('要確認')) warn.push(`${Core.label(studentOf(sid))}：兄弟姉妹の${info.slice(info.indexOf('要確認'))}（担当者が照合できませんでした）`);
    }
    const box = $('warnings');
    box.replaceChildren();
    if (bad.length) box.append(h('div', { class: 'box bad' }, h('h4', {}, '直してください'), h('ul', {}, bad.map((x) => h('li', {}, x)))));
    if (warn.length) box.append(h('div', { class: 'box warn' }, h('h4', {}, '確認してください'), h('ul', {}, warn.map((x) => h('li', {}, x)))));
    if (!bad.length && !warn.length) box.append(h('div', { class: 'box ok' }, '問題はありません。'));
  }

  function teacherSheets() {
    const d = T.data;
    return Core.classToSheets({ grade: d.grade, cls: d.cls, slots: d.slots, unassigned: d.unassigned }, d.students, d.bands, d.sibInfo, { minutes: d.minutes, createdAt: new Date().toLocaleString('ja-JP'), letter: T.letter });
  }

  $('saveBtn').addEventListener('click', async () => {
    const blob = sheetsToBlob(teacherSheets());
    if (T.handle && T.handle.createWritable) {
      try {
        const w = await T.handle.createWritable();
        await w.write(blob);
        await w.close();
        T.dirty = false;
        $('saveMsg').textContent = `${T.fileName} に上書き保存しました（${new Date().toLocaleTimeString('ja-JP')}）`;
        return;
      } catch (e) { /* 失敗したらダウンロードにする */ }
    }
    download(blob, T.fileName);
    T.dirty = false;
    $('saveMsg').textContent = `「ダウンロード」フォルダに ${T.fileName} を保存しました。共有フォルダのファイルと置き換えてください。`;
  });
  $('printBtn').addEventListener('click', () => {
    const d = T.data;
    const who = (sid) => { const s = studentOf(sid); const info = d.sibInfo.get(sid) || ''; return { no: s.no, name: s.name, sib: !!info && !info.startsWith('要確認') }; };
    printPages([schedulePage(d, who)]);
  });

  // ---------- 配布物の印刷 ----------
  $('printLettersBtn').addEventListener('click', () => {
    const placed = placedList();
    if (!placed.length) { toast('日時が決まった生徒がいません'); return; }
    const un = T.data.unassigned.length;
    if (un && !confirm(`未割当の生徒が${un}人います。この人たちの分は印刷されません。続けますか？`)) return;
    $('letterMsg').textContent = `${placed.length}人分（${placed.length}枚）を印刷します。`;
    printPages(placed.map(({ sl, s }) => letterPage(T.letter, { grade: s.grade, cls: s.cls, name: s.name, time: Core.letterTime(sl, T.letter.timeStyle) })));
  });
  window.addEventListener('afterprint', () => {
    document.body.classList.remove('print-letters');
    $('printArea').replaceChildren();
  });

  // ---------- 配布物 ----------
  setupDrop($('tplDrop'), $('tplFile'), async (files) => {
    const f = files[0];
    if (!f || !/\.docx$/i.test(f.name)) { toast('Wordファイル（.docx）を選んでください。古い形式（.doc）の場合はWordで「名前を付けて保存」→「Word文書(.docx)」にしてください', 7000); return; }
    T.template = { name: f.name, buf: await f.arrayBuffer() };
    $('tplName').textContent = `選んだWord：${f.name}`;
    $('docBtn').disabled = false;
  });

  $('docBtn').addEventListener('click', async () => {
    const d = T.data;
    const unassigned = d.unassigned.length;
    const placed = placedList();
    if (!placed.length) { toast('日時が決まった生徒がいません'); return; }
    if (unassigned && !confirm(`未割当の生徒が${unassigned}人います。この人たちの分は作られません。続けますか？`)) return;
    $('docMsg').textContent = '作成中…';
    try {
      const { blob, unknown, found } = await Docx.mergeDocx(T.template.buf, placed.map(({ sl, s }) => Core.mergeFields(s, sl)));
      download(blob, `配布物_${d.grade}年${d.cls}組.docx`);
      const msgs = [`${placed.length}人分の配布物を作りました（配布物_${d.grade}年${d.cls}組.docx）。Wordで開いて確認してから印刷してください。`];
      if (!found.length) msgs.push('※Wordの中に【氏名】などの目印が見つかりませんでした。同じ文書が人数分できています。');
      if (unknown.length) msgs.push(`※知らない目印があり、そのまま残しました：${unknown.map((u) => `【${u}】`).join('、')}`);
      $('docMsg').textContent = msgs.join(' ');
    } catch (e) {
      $('docMsg').textContent = `作れませんでした：${e.message}`;
    }
  });

  // テスト用
  window.__app = { A, T, analyze, showTab };
})();
