/*
 * 保護者会日程作成ツール：計算ロジック（画面に依存しない部分）
 * ブラウザでは window.Core、Node（テスト）では require で使う。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Core = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const VERSION = '1.0';
  const BREAK_LABEL = '（休憩）';
  const BLOCKED_LABEL = '（担任不可）';

  // 全角→半角、～→~ などをそろえる
  function norm(s) {
    return String(s == null ? '' : s).normalize('NFKC').trim();
  }
  // 名前照合用：空白・記号・敬称を落とす
  function normName(s) {
    return norm(s)
      .replace(/[（(][^）)]*[）)]/g, '')
      .replace(/(くん|君|さん|ちゃん|様)$/g, '')
      .replace(/[\s・･,、。.\-ー－_]/g, '');
  }
  const pad2 = (n) => String(n).padStart(2, '0');
  const toMin = (hm) => { const [h, m] = hm.split(':').map(Number); return h * 60 + m; };
  const toHM = (min) => `${Math.floor(min / 60)}:${pad2(min % 60)}`;

  const KANJI_NUM = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9 };

  // ---------- 時間帯（列見出し）の読み取り ----------
  // 例：「11月10日(火)13:30～14:30」
  const BAND_RE = /(\d{1,2})月(\d{1,2})日\s*[(（]?([月火水木金土日])?[)）]?\s*(\d{1,2}):(\d{2})\s*[~〜～\-－ー]\s*(\d{1,2}):(\d{2})/;

  function parseBand(header) {
    const m = norm(header).match(BAND_RE);
    if (!m) return null;
    const month = +m[1], day = +m[2];
    const start = `${+m[4]}:${m[5]}`, end = `${+m[6]}:${m[7]}`;
    return {
      key: `${pad2(month)}/${pad2(day)} ${pad2(+m[4])}:${m[5]}`,
      dateKey: `${pad2(month)}/${pad2(day)}`,
      month, day, dow: m[3] || '',
      start, end,
      label: `${month}/${day}(${m[3] || ''}) ${start}～${end}`,
    };
  }

  // ---------- Formsの書き出しExcel（行の配列）を読む ----------
  // rows: 1行目が見出しの2次元配列
  function findCols(header) {
    const h = header.map(norm);
    const find = (pred) => h.findIndex(pred);
    const cols = {
      cls: find((x) => /学級|組を/.test(x) && !/兄弟|姉妹/.test(x)),
      no: find((x) => /出席番号|番号/.test(x) && !/兄弟|姉妹/.test(x)),
      name: find((x) => /お子.*名前|生徒.*名前|児童.*名前|氏名/.test(x) && !/兄弟|姉妹/.test(x)),
      grade: find((x) => /学年/.test(x) && !/兄弟|姉妹/.test(x)),
      mode: find((x) => /どちらか|希望日時について/.test(x)),
      sib: find((x) => /兄弟|姉妹|きょうだい/.test(x)),
      done: find((x) => /完了時刻/.test(x)),
      bands: [],
    };
    h.forEach((x, i) => { const b = parseBand(x); if (b) cols.bands.push({ index: i, band: b }); });
    return cols;
  }

  // 学年をファイル名から推測（「3年」「第３学年」など）
  function guessGrade(text) {
    const m = norm(text).match(/([1-6一二三四五六])\s*(?:年|学年)/);
    if (!m) return null;
    return KANJI_NUM[m[1]] || +m[1];
  }

  /**
   * 1学年分のExcelを読み、生徒の配列と注意事項を返す
   * @returns {{students, bands, issues, cols}}
   */
  function parseGradeRows(rows, grade, fileName) {
    const issues = [];
    if (!rows.length) throw new Error(`${fileName}：データがありません`);
    const cols = findCols(rows[0]);
    if (cols.cls < 0 || cols.no < 0) throw new Error(`${fileName}：「学級」「出席番号」の列が見つかりません`);
    if (!cols.bands.length) throw new Error(`${fileName}：「11月10日(火)13:30～14:30」のような日時の列が見つかりません`);

    const byId = new Map();
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.every((v) => v == null || v === '')) continue;
      const cls = norm(row[cols.cls]).replace(/組$/, '').toUpperCase();
      const no = parseInt(norm(row[cols.no]), 10);
      if (!cls || !no) continue;
      const g = cols.grade >= 0 && parseInt(norm(row[cols.grade]), 10) ? parseInt(norm(row[cols.grade]), 10) : grade;
      const ng = [];
      cols.bands.forEach(({ index, band }) => { if (norm(row[index]) !== '') ng.push(band.key); });
      const modeText = cols.mode >= 0 ? norm(row[cols.mode]) : '';
      const anytime = /いつでも/.test(modeText);
      const done = cols.done >= 0 ? row[cols.done] : null;
      const st = {
        id: `${g}-${cls}-${no}`,
        grade: g, cls, no,
        name: cols.name >= 0 ? norm(row[cols.name]) : '',
        anytime,
        ng,
        sibText: cols.sib >= 0 ? norm(row[cols.sib]) : '',
        doneAt: done instanceof Date ? done.getTime() : (done ? Date.parse(done) || r : r),
        row: r + 1,
      };
      if (anytime && ng.length) {
        issues.push({ type: 'anytimeWithNg', id: st.id, text: `${label(st)}：「いつでもよい」なのに×が${ng.length}個あります → ×を優先して避けます` });
      }
      const prev = byId.get(st.id);
      if (prev) {
        const keep = st.doneAt >= prev.doneAt ? st : prev;
        issues.push({ type: 'duplicate', id: st.id, text: `${label(st)}：回答が2件あります（${prev.row}行目と${st.row}行目）→ 後から送信された${keep.row}行目を使います` });
        byId.set(st.id, keep);
      } else byId.set(st.id, st);
    }
    const students = [...byId.values()].sort(cmpStudent);

    // 欠番（未提出の可能性）
    const byCls = groupBy(students, (s) => s.cls);
    for (const [cls, list] of byCls) {
      const nos = new Set(list.map((s) => s.no));
      const max = Math.max(...nos);
      const miss = [];
      for (let i = 1; i < max; i++) if (!nos.has(i)) miss.push(i);
      if (miss.length) issues.push({ type: 'missing', id: `${grade}-${cls}`, text: `${grade}年${cls}組：出席番号 ${miss.join('、')} の回答がありません（未提出または欠番。最後の番号より後ろは判別できません）` });
    }
    return { students, bands: cols.bands.map((b) => b.band), issues, hasSibCol: cols.sib >= 0 };
  }

  function cmpStudent(a, b) {
    return a.grade - b.grade || a.cls.localeCompare(b.cls) || a.no - b.no;
  }
  function groupBy(arr, fn) {
    const m = new Map();
    for (const x of arr) { const k = fn(x); if (!m.has(k)) m.set(k, []); m.get(k).push(x); }
    return m;
  }
  function label(s) {
    return `${s.grade}年${s.cls}組${s.no}番 ${s.name}`;
  }

  // ---------- 兄弟姉妹欄の読み取り ----------
  // 「2年3組 山田花子」「２－Ｃ 山田」「山田花子(2年C組)」「姉:2年C組・山田花子、弟:1年A組・山田一郎」など
  const NONE_RE = /^(なし|無し|ない|いない|いません|おりません|ありません|特になし|該当なし|-|ー|―|×|0|なし。)$/;
  const GC_RE = /([1-6一二三四五六])\s*(?:年生?|ねん|[-‐－ー―の])\s*([A-Za-z]|[1-9一二三四五六七八九])\s*(?:組|くみ)?/g;

  const OTHER_SCHOOL_RE = /小学|小\d|高校|高\d|保育|幼稚|大学|未就学|社会人/;

  function parseSiblingText(text) {
    const t = norm(text);
    if (!t || NONE_RE.test(t.replace(/[。.、\s]/g, ''))) return [];
    const matches = [...t.matchAll(GC_RE)];
    if (!matches.length) return [{ raw: t, grade: null, cls: null, name: normName(t.replace(/^(兄|姉|弟|妹)\s*[:：]?/, '')), otherSchool: OTHER_SCHOOL_RE.test(t) }];
    const out = [];
    matches.forEach((m, i) => {
      const grade = KANJI_NUM[m[1]] || +m[1];
      let c = m[2];
      c = KANJI_NUM[c] ? String(KANJI_NUM[c]) : c.toUpperCase();
      // 名前は「この学年組の後ろ〜次の学年組の前」。後ろが空なら前を見る（例：山田花子(2年C組)）
      const after = t.slice(m.index + m[0].length, i + 1 < matches.length ? matches[i + 1].index : t.length);
      const before = t.slice(i === 0 ? 0 : matches[i - 1].index + matches[i - 1][0].length, m.index);
      const clean = (s) => normName(s.replace(/[（(）)\[\]【】「」]/g, ' ').replace(/(小学校|中学校|高校|中学|小学)/g, '').replace(/^(兄|姉|弟|妹|長男|長女|次男|次女)\s*[:：]?/, '').replace(/(兄|姉|弟|妹|長男|長女|次男|次女)\s*[:：]?\s*$/, ''));
      let name = clean(after.split(/[、,;\n]/).find((x) => x.trim()) || '');
      if (!name) name = clean(before);
      const school = OTHER_SCHOOL_RE.test(t.slice(Math.max(0, m.index - 6), m.index + m[0].length + 2));
      out.push({ raw: m[0] + (name ? ' ' + name : ''), grade, cls: c, name, otherSchool: school });
    });
    return out;
  }

  // 1組→A のような対応（学級が英字の学校向け）
  function clsVariants(c) {
    if (!c) return [];
    if (/^\d$/.test(c)) return [c, String.fromCharCode(64 + +c)];
    if (/^[A-Z]$/.test(c)) return [c, String(c.charCodeAt(0) - 64)];
    return [c];
  }

  /**
   * 兄弟姉妹欄を全学年の生徒と照合する
   * @returns {{links: Array<[idA,idB]>, problems: Array}}
   */
  function matchSiblings(students) {
    const byGrade = groupBy(students, (s) => s.grade);
    const links = [];
    const problems = [];
    for (const s of students) {
      if (!s.sibText) continue;
      const entries = parseSiblingText(s.sibText);
      for (const e of entries) {
        if (e.otherSchool) continue;
        const res = resolveEntry(e, s, byGrade);
        if (res.target) links.push([s.id, res.target.id]);
        else problems.push({ id: s.id, raw: s.sibText, entry: e, reason: res.reason });
      }
    }
    // 反対側の子の欄で照合できている場合は問題にしない
    const byId = new Map(students.map((x) => [x.id, x]));
    const partners = (id) => links.filter((l) => l.includes(id)).map((l) => byId.get(l[0] === id ? l[1] : l[0]));
    const rest = problems.filter((p) => !partners(p.id).some((q) => q.grade === p.entry.grade && (!p.entry.cls || clsVariants(p.entry.cls).includes(q.cls))));
    return { links: dedupeLinks(links), problems: rest };
  }

  function resolveEntry(e, self, byGrade) {
    const pool = e.grade ? (byGrade.get(e.grade) || []) : [...byGrade.values()].flat();
    if (e.grade && !byGrade.has(e.grade)) return { reason: `${e.grade}年のデータが読み込まれていません` };
    const others = pool.filter((p) => p.id !== self.id);
    const variants = clsVariants(e.cls);
    const inCls = e.cls ? others.filter((p) => variants.includes(p.cls)) : others;
    const nm = e.name;
    const eq = (p) => normName(p.name) === nm;
    const part = (p) => nm.length >= 2 && (normName(p.name).includes(nm) || nm.includes(normName(p.name)));
    const pick = (list) => (list.length === 1 ? list[0] : null);
    if (nm) {
      const hit = pick(inCls.filter(eq)) || pick(others.filter(eq)) || pick(inCls.filter(part)) || pick(others.filter(part));
      if (hit) return { target: hit };
      return { reason: '名前が見つかりません' };
    }
    return { reason: '名前が書かれていません' };
  }

  function dedupeLinks(links) {
    const seen = new Set();
    return links.filter(([a, b]) => {
      const k = [a, b].sort().join('|');
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    });
  }

  // 兄弟リンクから家族グループを作る（union-find）
  function buildFamilies(students, links) {
    const parent = new Map(students.map((s) => [s.id, s.id]));
    const find = (x) => (parent.get(x) === x ? x : (parent.set(x, find(parent.get(x))), parent.get(x)));
    for (const [a, b] of links) if (parent.has(a) && parent.has(b)) parent.set(find(a), find(b));
    const groups = groupBy(students, (s) => find(s.id));
    return [...groups.values()].filter((g) => g.length > 1).map((g) => g.sort(cmpStudent));
  }

  // ---------- 面談枠をつくる ----------
  /**
   * @param bands そのクラスで使う時間帯
   * @param opt {minutes, breakEvery, blocked:Set(bandKey)}
   */
  function buildSlots(bands, opt) {
    const minutes = Math.max(1, +opt.minutes || 10);
    const slots = [];
    const sorted = [...bands].sort((a, b) => a.key.localeCompare(b.key));
    for (const b of sorted) {
      for (let t = toMin(b.start); t + minutes <= toMin(b.end); t += minutes) {
        slots.push({ bandKey: b.key, dateKey: b.dateKey, month: b.month, day: b.day, dow: b.dow, start: t, end: t + minutes, kind: opt.blocked && opt.blocked.has(b.key) ? 'blocked' : 'free', sid: null });
      }
    }
    // N人ごとに休憩
    const every = +opt.breakEvery || 0;
    if (every > 0) {
      let run = 0, lastDate = null, lastEnd = null;
      for (const s of slots) {
        if (s.dateKey !== lastDate || s.start !== lastEnd) run = 0;
        lastDate = s.dateKey; lastEnd = s.end;
        if (s.kind !== 'free') { run = 0; continue; }
        if (run === every) { s.kind = 'break'; run = 0; } else run++;
      }
    }
    return slots;
  }

  // ---------- 割り当て ----------
  /**
   * @param students 全学年の生徒
   * @param bandsByGrade Map(grade -> bands[])
   * @param opt {minutes, breakEvery, sibGap(分), blockedByClass: Map(`${grade}-${cls}` -> Set(bandKey)), links}
   * @returns {classes: Map(classKey -> {grade, cls, slots, unassigned:[]}), families, notes}
   */
  function schedule(students, bandsByGrade, opt) {
    const byId = new Map(students.map((s) => [s.id, s]));
    const classes = new Map();
    for (const s of students) {
      const ck = `${s.grade}-${s.cls}`;
      if (!classes.has(ck)) {
        classes.set(ck, {
          grade: s.grade, cls: s.cls,
          slots: buildSlots(bandsByGrade.get(s.grade) || [], { minutes: opt.minutes, breakEvery: opt.breakEvery, blocked: (opt.blockedByClass && opt.blockedByClass.get(ck)) || new Set() }),
          unassigned: [],
        });
      }
    }
    const ngSet = new Map(students.map((s) => [s.id, new Set(s.ng)]));
    const slotsOf = (s) => classes.get(`${s.grade}-${s.cls}`).slots;
    const usable = (s, slot) => slot.kind === 'free' && !ngSet.get(s.id).has(slot.bandKey);
    const where = new Map(); // sid -> slot
    const put = (s, slot) => { slot.sid = s.id; where.set(s.id, slot); };
    const notes = [];

    const families = buildFamilies(students, opt.links || []);
    const locked = new Set();
    const gap = +opt.sibGap || 0;
    const tol = Math.max(+opt.minutes || 10, 10);

    const freeCount = (s) => slotsOf(s).filter((x) => !x.sid && usable(s, x)).length;

    // 1) 兄弟姉妹：同じ日に、前の子の終了から gap 分後〜(gap+tol)分後に始まる枠で連続させる
    const famOrder = families.slice().sort((a, b) => Math.min(...a.map(freeCount)) - Math.min(...b.map(freeCount)));
    for (const fam of famOrder) {
      const placed = placeFamily(fam);
      if (placed) {
        placed.forEach(([s, slot]) => { put(s, slot); locked.add(s.id); });
      } else {
        notes.push({ type: 'familySplit', ids: fam.map((s) => s.id), text: `兄弟（${fam.map(label).join('／')}）を同じ日の続きの時間に組めませんでした → 別々に配置しました` });
      }
    }
    function placeFamily(fam) {
      const dates = [...new Set(fam.flatMap((s) => slotsOf(s).map((x) => x.dateKey)))].sort();
      for (const perm of permutations(fam)) {
        for (const d of dates) {
          for (const first of slotsOf(perm[0])) {
            if (first.dateKey !== d || first.sid || !usable(perm[0], first)) continue;
            const chain = [[perm[0], first]];
            let prev = first, ok = true;
            for (let i = 1; i < perm.length; i++) {
              const s = perm[i];
              const used = new Set(chain.map((c) => c[1]));
              const next = slotsOf(s).find((x) => x.dateKey === d && !x.sid && !used.has(x) && usable(s, x) && x.start >= prev.end + gap && x.start <= prev.end + gap + tol);
              if (!next) { ok = false; break; }
              chain.push([s, next]);
              prev = next;
            }
            if (ok) return chain;
          }
        }
      }
      return null;
    }

    // 2) 残り：選べる枠が少ない家庭から先に、早い枠へ
    const rest = students.filter((s) => !where.has(s.id)).sort((a, b) => freeCount(a) - freeCount(b) || cmpStudent(a, b));
    for (const s of rest) {
      const cand = slotsOf(s).filter((x) => !x.sid && usable(s, x));
      if (!cand.length) continue;
      let slot = cand[0];
      if (opt.spread !== false) {
        // 各日の人数がそろうように、いちばん空いている日の早い枠へ
        const perDay = new Map();
        for (const x of slotsOf(s)) if (x.sid) perDay.set(x.dateKey, (perDay.get(x.dateKey) || 0) + 1);
        slot = cand.reduce((best, x) => ((perDay.get(x.dateKey) || 0) < (perDay.get(best.dateKey) || 0) ? x : best), cand[0]);
      }
      put(s, slot);
    }

    // 3) 入らなかった家庭は、他の家庭を空いている枠へずらして入れられないか試す（増加路）
    for (const s of students) {
      if (where.has(s.id)) continue;
      const seen = new Set();
      const tryPlace = (u) => {
        for (const slot of slotsOf(u)) {
          if (!usable(u, slot) || seen.has(slot)) continue;
          seen.add(slot);
          if (!slot.sid) { put(u, slot); return true; }
          const v = byId.get(slot.sid);
          if (locked.has(v.id)) continue;
          where.delete(v.id);
          slot.sid = null;
          if (tryPlace(v)) { put(u, slot); return true; }
          put(v, slot);
        }
        return false;
      };
      tryPlace(s);
    }

    // 未割当
    for (const s of students) {
      if (where.has(s.id)) continue;
      const c = classes.get(`${s.grade}-${s.cls}`);
      const any = c.slots.some((x) => usable(s, x));
      c.unassigned.push(s.id);
      notes.push({ type: 'unassigned', id: s.id, text: `${label(s)}：${any ? '空いている枠がありません（満員）' : 'すべての時間帯に×があります'} → 担任が手動で入れてください` });
    }
    return { classes, families, notes, byId };
  }

  function* permutations(arr) {
    if (arr.length <= 1) { yield arr.slice(); return; }
    for (let i = 0; i < arr.length; i++) {
      const rest = arr.slice(0, i).concat(arr.slice(i + 1));
      for (const p of permutations(rest)) yield [arr[i], ...p];
    }
  }

  // ---------- 表示用 ----------
  function slotDateText(s) {
    return `${s.month}月${s.day}日（${s.dow}）`;
  }
  function slotTimeText(s) {
    return `${toHM(s.start)}～${toHM(s.end)}`;
  }

  // ---------- 学級ごとのExcel（日程案）の中身 ----------
  const SCHED_HEAD = ['日付', '曜日', '開始', '終了', '出席番号', '生徒名', '兄弟姉妹', 'メモ'];

  /**
   * 学級の日程案を3シート分の2次元配列にする
   * sibInfo: Map(sid -> 表示文字列)
   */
  function classToSheets(cls, byId, bands, sibInfo, meta) {
    const sched = [SCHED_HEAD];
    for (const sl of cls.slots) {
      const s = sl.sid ? byId.get(sl.sid) : null;
      const nameCell = s ? s.name : sl.kind === 'break' ? BREAK_LABEL : sl.kind === 'blocked' ? BLOCKED_LABEL : '';
      sched.push([`${sl.month}/${sl.day}`, sl.dow, toHM(sl.start), toHM(sl.end), s ? s.no : '', nameCell, s ? sibInfo.get(s.id) || '' : '', sl.memo || '']);
    }
    for (const sid of cls.unassigned) {
      const s = byId.get(sid);
      sched.push(['未割当', '', '', '', s.no, s.name, sibInfo.get(s.id) || '', '']);
    }
    const studentsOfCls = [...byId.values()].filter((s) => s.grade === cls.grade && s.cls === cls.cls).sort(cmpStudent);
    const sortedBands = [...bands].sort((a, b) => a.key.localeCompare(b.key));
    const pref = [['出席番号', '生徒名', '回答', ...sortedBands.map((b) => b.label), '兄弟姉妹（記入）']];
    for (const s of studentsOfCls) {
      const ng = new Set(s.ng);
      pref.push([s.no, s.name, s.anytime ? 'いつでもよい' : '都合の悪い日時がある', ...sortedBands.map((b) => (ng.has(b.key) ? '×' : '')), s.sibText || '']);
    }
    const sheets = { 日程: sched, 希望: pref };
    if (meta.letter) sheets['文面'] = letterToRows(meta.letter);
    const info = [
      ['項目', '値'],
      ['学年', cls.grade],
      ['組', cls.cls],
      ['1人あたり（分）', meta.minutes],
      ['作成日時', meta.createdAt],
      ['ツール版', VERSION],
    ];
    sheets['情報'] = info;
    return sheets;
  }

  /**
   * 学級の日程案Excel（3シート）を読み戻す
   * sheets: {日程: rows, 希望: rows, 情報: rows}
   */
  function sheetsToClass(sheets) {
    const info = new Map((sheets['情報'] || []).slice(1).map((r) => [norm(r[0]), r[1]]));
    const infoGet = (k) => info.get(norm(k));
    const grade = parseInt(norm(infoGet('学年')), 10);
    const cls = norm(infoGet('組'));
    if (!grade || !cls) throw new Error('このファイルは「日程案」ではないようです（「情報」シートがありません）');
    const pref = sheets['希望'] || [];
    const head = (pref[0] || []).map(norm);
    const bandCols = [];
    head.forEach((h, i) => {
      const m = h.match(/^(\d{1,2})\/(\d{1,2})\(([^)]*)\)\s*(\d{1,2}:\d{2})\s*[~～〜]\s*(\d{1,2}:\d{2})$/);
      if (m) {
        const [hh, mm] = m[4].split(':');
        bandCols.push({ index: i, band: { key: `${pad2(+m[1])}/${pad2(+m[2])} ${pad2(+hh)}:${mm}`, dateKey: `${pad2(+m[1])}/${pad2(+m[2])}`, month: +m[1], day: +m[2], dow: m[3], start: m[4], end: m[5], label: `${+m[1]}/${+m[2]}(${m[3]}) ${m[4]}～${m[5]}` } });
      }
    });
    const sibCol = head.indexOf(norm('兄弟姉妹（記入）'));
    const students = new Map();
    for (const r of pref.slice(1)) {
      const no = parseInt(norm(r[0]), 10);
      if (!no) continue;
      const s = { id: `${grade}-${cls}-${no}`, grade, cls, no, name: norm(r[1]), anytime: /いつでも/.test(norm(r[2])), ng: bandCols.filter((b) => norm(r[b.index]) !== '').map((b) => b.band.key), sibText: sibCol >= 0 ? text(r[sibCol]) : '' };
      students.set(s.id, s);
    }
    const bands = bandCols.map((b) => b.band);
    const bandOf = (month, day, startMin) => bands.find((b) => b.month === month && b.day === day && toMin(b.start) <= startMin && startMin < toMin(b.end));
    const slots = [];
    const unassigned = [];
    const sibInfo = new Map();
    const problems = [];
    const sched = sheets['日程'] || [];
    for (const r of sched.slice(1)) {
      const d = dateCell(r[0]);
      const no = parseInt(norm(r[4]), 10);
      const name = norm(r[5]);
      const sid = no ? `${grade}-${cls}-${no}` : null;
      if (sid && !students.has(sid)) {
        students.set(sid, { id: sid, grade, cls, no, name, anytime: true, ng: [], sibText: '' });
        problems.push(`${no}番 ${name}：希望シートにいない生徒です`);
      }
      if (sid && text(r[6])) sibInfo.set(sid, text(r[6]));
      if (d === '未割当') { if (sid) unassigned.push(sid); continue; }
      const dm = d.match(/^(\d{1,2})\/(\d{1,2})$/);
      if (!dm) continue;
      const start = toMin(timeCell(r[2])), end = toMin(timeCell(r[3]));
      const band = bandOf(+dm[1], +dm[2], start);
      const kind = name === norm(BREAK_LABEL) || name === '休憩' ? 'break' : name === norm(BLOCKED_LABEL) ? 'blocked' : 'free';
      slots.push({ bandKey: band ? band.key : `${pad2(+dm[1])}/${pad2(+dm[2])} ?`, dateKey: `${pad2(+dm[1])}/${pad2(+dm[2])}`, month: +dm[1], day: +dm[2], dow: norm(r[1]), start, end, kind, sid: kind === 'free' ? sid : null, memo: text(r[7]) });
    }
    // 同じ生徒が2回出てくる（Excelで直接編集したとき）
    const seen = new Map();
    for (const sl of slots) {
      if (!sl.sid) continue;
      if (seen.has(sl.sid)) { problems.push(`${students.get(sl.sid).no}番 ${students.get(sl.sid).name}：2つの枠に入っています（${slotDateText(seen.get(sl.sid))} と ${slotDateText(sl)}）`); }
      else seen.set(sl.sid, sl);
    }
    // 希望シートにいるのに日程にも未割当にもいない生徒
    for (const s of students.values()) if (!seen.has(s.id) && !unassigned.includes(s.id)) unassigned.push(s.id);
    return { grade, cls, slots, unassigned, students, bands, sibInfo, minutes: infoGet('1人あたり（分）'), letter: rowsToLetter(sheets['文面']), problems };
  }

  // 表示用の文字（全角・半角はそのまま）
  function text(v) { return String(v == null ? '' : v).trim(); }

  // Excelの日付セル（"11/10"、Excelが日付に変えたもの、シリアル値）を "11/10" に
  function dateCell(v) {
    if (v instanceof Date) return `${v.getMonth() + 1}/${v.getDate()}`;
    if (typeof v === 'number' && v > 20000) { const d = new Date(Math.round((v - 25569) * 864e5)); return `${d.getUTCMonth() + 1}/${d.getUTCDate()}`; }
    const t = norm(v);
    const m = t.match(/^(?:\d{4}[\/-])?(\d{1,2})[\/-](\d{1,2})$/) || t.match(/^(\d{1,2})月(\d{1,2})日/);
    return m ? `${+m[1]}/${+m[2]}` : t;
  }

  // Excelの時刻セル（"13:30" または 0.5625 のような小数）を "13:30" に
  function timeCell(v) {
    if (typeof v === 'number') { const m = Math.round(v * 24 * 60); return `${Math.floor(m / 60)}:${pad2(m % 60)}`; }
    if (v instanceof Date) return `${v.getHours()}:${pad2(v.getMinutes())}`;
    const m = norm(v).match(/(\d{1,2}):(\d{2})/);
    return m ? `${+m[1]}:${m[2]}` : '0:00';
  }

  // ---------- 配布物の文面（保護印刷シートと同じ構成） ----------
  const LETTER_FIELDS = [
    ['date', '発行日'],
    ['to', '宛名'],
    ['school', '学校名'],
    ['principal', '校長名'],
    ['title', '表題'],
    ['body', '前文'],
    ['item1', '１の見出し'],
    ['note', '日時の下の注意書き'],
    ['items', '２以降の項目'],
    ['timeStyle', '時刻の書き方'],
  ];
  const DEFAULT_LETTER = {
    date: '',
    to: '保　護　者　様',
    school: '名古屋市立汐路中学校',
    principal: '鈴　木　暁　子',
    title: '保護者会（個別懇談会）のお知らせ',
    body: '　梅雨の候、皆様にはますますご健勝のこととお喜び申し上げます。平素より本校の教育活動にご理解とご協力を賜り、誠にありがとうございます。\n'
      + '　さて、下記のとおり保護者会（個人懇談会）の日時が決定いたしましたので、ご案内申し上げます。\n'
      + '　ご多用のところ恐縮ではございますが、ぜひご出席くださいますようお願いいたします。なお、やむを得ずご都合がつかない場合は、学級担任までお知らせください。',
    item1: '１　懇談日時',
    note: '※若干お時間が前後することがございます。ご了承ください。',
    items: '２　場　所　　　各教室\n\n'
      + '３　内　容　　　学習（進路）や生活等について\n\n'
      + '４　その他　　　※　スリッパをご持参ください。\n'
      + '　　　　　　　　※　正門からお入りください。\n'
      + '　　　　　　　　※　お車での来校はご遠慮ください。\n'
      + '　　　　　　　　※　自転車は正門の内側に駐輪してください。\n'
      + '　　　　　　　　※　ご来校の際は、つりさげ名札を必ずお付けください。',
    timeStyle: 'start',
  };

  // 令和の日付（全角）
  function reiwaDate(d) {
    return toFull(`令和${d.getFullYear() - 2018}年${d.getMonth() + 1}月${d.getDate()}日`);
  }
  function toFull(s) {
    return String(s).replace(/[0-9:]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 0xfee0));
  }
  function letterToRows(letter) {
    return [['項目', '内容'], ...LETTER_FIELDS.map(([k, label]) => [label, letter[k] == null ? '' : String(letter[k])])];
  }
  function rowsToLetter(rows) {
    if (!rows || rows.length < 2) return null;
    const byLabel = new Map(rows.slice(1).map((r) => [text(r[0]), r[1] == null ? '' : String(r[1]).replace(/\r\n?/g, '\n')]));
    const out = { ...DEFAULT_LETTER };
    for (const [k, label] of LETTER_FIELDS) if (byLabel.has(label)) out[k] = byLabel.get(label);
    return out;
  }
  // 配布物の日時（例：１１月１０日（火）　１３：３０～）
  function letterTime(slot, style) {
    const t = style === 'range' ? `${toHM(slot.start)}～${toHM(slot.end)}` : `${toHM(slot.start)}～`;
    return toFull(`${slot.month}月${slot.day}日（${slot.dow}）　${t}`);
  }

  // 差し込み用の値
  function mergeFields(s, slot) {
    return {
      学年: String(s.grade),
      組: s.cls,
      番号: String(s.no),
      出席番号: String(s.no),
      氏名: s.name,
      生徒名: s.name,
      月日: `${slot.month}月${slot.day}日`,
      曜日: slot.dow,
      時刻: toHM(slot.start),
      開始: toHM(slot.start),
      終了時刻: toHM(slot.end),
      終了: toHM(slot.end),
      日時: `${slot.month}月${slot.day}日（${slot.dow}）${toHM(slot.start)}～${toHM(slot.end)}`,
    };
  }

  return {
    VERSION, BREAK_LABEL, BLOCKED_LABEL,
    norm, normName, toHM, toMin,
    parseBand, findCols, guessGrade, parseGradeRows,
    parseSiblingText, matchSiblings, buildFamilies,
    buildSlots, schedule,
    classToSheets, sheetsToClass, timeCell,
    slotDateText, slotTimeText, mergeFields, label, cmpStudent,
    LETTER_FIELDS, DEFAULT_LETTER, reiwaDate, toFull, letterToRows, rowsToLetter, letterTime,
  };
});
