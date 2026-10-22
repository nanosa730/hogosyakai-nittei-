const test = require('node:test');
const assert = require('node:assert');
const path = require('path');
const XLSX = require('xlsx');
const Core = require('../src/core.js');

const DEMO = path.join(__dirname, '..', 'demo');

function loadGrade(file, grade) {
  const wb = XLSX.readFile(path.join(DEMO, file), { cellDates: true });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
  return Core.parseGradeRows(rows, grade, file);
}

function loadAll() {
  const g2 = loadGrade('demo_懇談会希望調査_2年.xlsx', 2);
  const g3 = loadGrade('demo_懇談会希望調査_3年.xlsx', 3);
  const students = [...g2.students, ...g3.students];
  const bandsByGrade = new Map([[2, g2.bands], [3, g3.bands]]);
  return { students, bandsByGrade, issues: [...g2.issues, ...g3.issues] };
}

test('時間帯の見出しを読む', () => {
  const b = Core.parseBand('11月10日(火)13:30～14:30');
  assert.equal(b.key, '11/10 13:30');
  assert.equal(b.dow, '火');
  assert.equal(b.end, '14:30');
  assert.equal(Core.parseBand('懇談の希望日時について'), null);
});

test('ファイル名から学年を推測する', () => {
  assert.equal(Core.guessGrade('懇談会希望調査_3年.xlsx'), 3);
  assert.equal(Core.guessGrade('第２学年.xlsx'), 2);
  assert.equal(Core.guessGrade('一年生.xlsx'), 1);
  assert.equal(Core.guessGrade('調査.xlsx'), null);
});

test('Formsの書き出しを読み、重複回答は後のものを使う', () => {
  const { students, issues } = loadAll();
  const d1 = students.find((s) => s.id === '3-D-1');
  assert.deepEqual(d1.ng.length, 4);
  assert.ok(issues.some((i) => i.type === 'duplicate' && i.id === '3-D-1'));
  assert.ok(issues.some((i) => i.type === 'missing'));
  assert.equal(new Set(students.map((s) => s.id)).size, students.length);
});

test('兄弟姉妹欄のいろいろな書き方', () => {
  const p = (t) => Core.parseSiblingText(t);
  assert.deepEqual(p('なし'), []);
  assert.deepEqual(p(''), []);
  assert.deepEqual(p('2年B組 山田花子').map((e) => [e.grade, e.cls, e.name]), [[2, 'B', '山田花子']]);
  assert.deepEqual(p('２－Ｃ　山田 花子').map((e) => [e.grade, e.cls, e.name]), [[2, 'C', '山田花子']]);
  assert.deepEqual(p('山田花子（2年D組）').map((e) => [e.grade, e.cls, e.name]), [[2, 'D', '山田花子']]);
  assert.deepEqual(p('姉：3年C組・山田花子、弟：1年A組・山田一郎').map((e) => [e.grade, e.cls, e.name]), [[3, 'C', '山田花子'], [1, 'A', '山田一郎']]);
  assert.deepEqual(p('2年3組 山田花子くん').map((e) => [e.grade, e.cls, e.name]), [[2, '3', '山田花子']]);
  assert.equal(p('弟 小学校4年 山田一郎')[0].otherSchool, true);
});

test('兄弟姉妹を照合する（組の数字↔英字、組の書き忘れ、反対側からの記入）', () => {
  const { students } = loadAll();
  const { links, problems } = Core.matchSiblings(students);
  const has = (a, b) => links.some((l) => l.includes(a) && l.includes(b));
  assert.ok(has('3-A-3', '2-B-3'));
  assert.ok(has('3-D-10', '2-C-10'), '2年3組 → 2年C組');
  assert.ok(has('2-A-15', '3-B-15'), '組の書き忘れでも名前で見つける');
  assert.ok(has('3-C-8', '2-D-8'));
  assert.ok(has('2-E-2', '2-E-3'), '同じ学年のきょうだい');
  // 照合できないもの：名前の誤り、1年が未読込。他校は問題にしない
  assert.deepEqual(problems.map((p) => p.id).sort(), ['3-A-20', '3-E-12']);
  assert.ok(!problems.some((p) => p.id === '3-C-8'), '相手側で照合できていれば問題にしない');
});

test('割り当て：×の時間・二重予約がなく、兄弟姉妹は同じ日の続きの時間', () => {
  const { students, bandsByGrade } = loadAll();
  const { links } = Core.matchSiblings(students);
  const blocked = new Map([['3-A', new Set(['11/10 13:30'])]]);
  const res = Core.schedule(students, bandsByGrade, { minutes: 10, breakEvery: 5, sibGap: 0, links, blockedByClass: blocked });
  const slotOf = new Map();
  for (const c of res.classes.values()) {
    for (const sl of c.slots) {
      if (!sl.sid) continue;
      const s = res.byId.get(sl.sid);
      assert.equal(sl.kind, 'free');
      assert.ok(!s.ng.includes(sl.bandKey), `${s.id} が×の時間に入っている`);
      assert.ok(!slotOf.has(sl.sid), `${s.id} が2回入っている`);
      slotOf.set(sl.sid, sl);
    }
  }
  // 全員どこかに入る（または未割当に入る）
  for (const s of students) assert.ok(slotOf.has(s.id) || res.classes.get(`${s.grade}-${s.cls}`).unassigned.includes(s.id));
  // 担任不可の時間には入らない
  assert.ok(res.classes.get('3-A').slots.filter((x) => x.bandKey === '11/10 13:30').every((x) => x.kind === 'blocked' && !x.sid));
  // 兄弟姉妹
  assert.ok(res.families.length >= 8);
  const split = new Set(res.notes.filter((n) => n.type === 'familySplit').flatMap((n) => n.ids));
  for (const fam of res.families) {
    if (fam.some((s) => split.has(s.id))) continue;
    const sl = fam.map((s) => slotOf.get(s.id)).sort((a, b) => a.start - b.start);
    assert.ok(sl.every((x) => x.dateKey === sl[0].dateKey), '同じ日');
    for (let i = 1; i < sl.length; i++) assert.ok(sl[i].start >= sl[i - 1].end && sl[i].start - sl[i - 1].end <= 10, '続きの時間');
  }
});

test('すべて×の家庭と満員は未割当になる', () => {
  const bands = [Core.parseBand('11月10日(火)13:30～13:50')];
  const mk = (no, ng) => ({ id: `1-A-${no}`, grade: 1, cls: 'A', no, name: `生徒${no}`, ng, anytime: false, sibText: '' });
  const students = [mk(1, ['11/10 13:30']), mk(2, []), mk(3, []), mk(4, [])];
  const res = Core.schedule(students, new Map([[1, bands]]), { minutes: 10 });
  const c = res.classes.get('1-A');
  assert.equal(c.slots.filter((x) => x.sid).length, 2);
  assert.equal(c.unassigned.length, 2);
  assert.ok(c.unassigned.includes('1-A-1'));
  assert.ok(res.notes.some((n) => n.id === '1-A-1' && /すべての時間帯に×/.test(n.text)));
});

test('選べる枠が少ない家庭を、ずらして入れる', () => {
  const bands = [Core.parseBand('11月10日(火)13:30～13:50')];
  const mk = (no, ng) => ({ id: `1-A-${no}`, grade: 1, cls: 'A', no, name: `生徒${no}`, ng, anytime: false, sibText: '' });
  // 2つの枠（13:30, 13:40）は同じ時間帯なので、×は帯ごと。2人とも入れる
  const res = Core.schedule([mk(1, []), mk(2, [])], new Map([[1, bands]]), { minutes: 10 });
  assert.equal(res.classes.get('1-A').unassigned.length, 0);
});

test('休憩は N人ごとに1枠', () => {
  const slots = Core.buildSlots([Core.parseBand('11月10日(火)13:30～14:30')], { minutes: 10, breakEvery: 2 });
  assert.deepEqual(Core.buildSlots([Core.parseBand('11月10日(火)13:30～14:30')], { minutes: 10, breakEvery: 3 }).map((s) => s.kind), ['free', 'free', 'free', 'break', 'free', 'free']);
  assert.deepEqual(slots.map((s) => s.kind), ['free', 'free', 'break', 'free', 'free', 'break']);
});

test('学級のExcelに書き出して読み戻しても同じ（×・兄弟・休憩・未割当）', () => {
  const { students, bandsByGrade } = loadAll();
  const { links } = Core.matchSiblings(students);
  const res = Core.schedule(students, bandsByGrade, { minutes: 10, breakEvery: 4, links });
  const cls = res.classes.get('3-A');
  cls.unassigned.push(cls.slots.find((x) => x.sid).sid);
  cls.slots.find((x) => x.sid && x.sid === cls.unassigned[0]).sid = null;
  const sib = new Map([['3-A-3', '2年B組3番 生徒2203（11/13 13:30～13:40）']]);
  const sheets = Core.classToSheets(cls, res.byId, bandsByGrade.get(3), sib, { minutes: 10, createdAt: 'x' });
  // 実際にxlsxにして読む
  const wb = XLSX.utils.book_new();
  for (const [n, rows] of Object.entries(sheets)) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), n);
  const wb2 = XLSX.read(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }), { cellDates: true });
  const back = {};
  for (const n of wb2.SheetNames) back[n] = XLSX.utils.sheet_to_json(wb2.Sheets[n], { header: 1, raw: true, defval: null });
  const d = Core.sheetsToClass(back);
  assert.equal(d.grade, 3);
  assert.equal(d.cls, 'A');
  assert.equal(d.minutes, 10);
  assert.equal(d.bands.length, 10);
  assert.equal(d.slots.length, cls.slots.length);
  assert.deepEqual(d.slots.map((s) => [s.kind, s.sid, s.bandKey]), cls.slots.map((s) => [s.kind, s.sid, s.bandKey]));
  assert.deepEqual(d.unassigned, cls.unassigned);
  assert.equal(d.sibInfo.get('3-A-3'), '2年B組3番 生徒2203（11/13 13:30～13:40）');
  for (const s of d.students.values()) assert.deepEqual(s.ng.slice().sort(), res.byId.get(s.id).ng.slice().sort());
  assert.deepEqual(d.problems, []);
  // 2回目の書き出しで見出しが変わらない
  const again = Core.classToSheets({ grade: d.grade, cls: d.cls, slots: d.slots, unassigned: d.unassigned }, d.students, d.bands, d.sibInfo, { minutes: d.minutes, createdAt: 'x' });
  assert.deepEqual(again['希望'][0], sheets['希望'][0]);
});

test('Excelで日付・時刻が変換されても読める', () => {
  assert.equal(Core.timeCell(0.5625), '13:30');
  assert.equal(Core.timeCell('13:30'), '13:30');
  const sheets = {
    情報: [['項目', '値'], ['学年', 1], ['組', 'A']],
    希望: [['出席番号', '生徒名', '回答', '11/10(火) 13:30～14:30', '兄弟姉妹（記入）'], [1, '生徒1', '都合の悪い日時がある', '×', '']],
    日程: [['日付'], [new Date(2026, 10, 10), '火', 0.5625, 0.5694444, 1, '生徒1', '', '']],
  };
  const d = Core.sheetsToClass(sheets);
  assert.equal(d.slots[0].bandKey, '11/10 13:30');
  assert.equal(d.slots[0].sid, '1-A-1');
});

test('差し込みの値', () => {
  const f = Core.mergeFields({ grade: 3, cls: 'A', no: 5, name: '生徒3105' }, { month: 11, day: 10, dow: '火', start: 810, end: 820 });
  assert.equal(f['日時'], '11月10日（火）13:30～13:40');
  assert.equal(f['氏名'], '生徒3105');
});

test('文面シートを書き出して読み戻す（改行も残る）', () => {
  const l = { ...Core.DEFAULT_LETTER, body: '　晩秋の候、\n　さて、', date: '令和８年１１月２日' };
  const back = Core.rowsToLetter(Core.letterToRows(l));
  assert.deepEqual(back, l);
  assert.equal(Core.letterTime({ month: 11, day: 10, dow: '火', start: 810, end: 820 }, 'start'), '１１月１０日（火）　１３：３０～');
  assert.equal(Core.letterTime({ month: 11, day: 10, dow: '火', start: 810, end: 820 }, 'range'), '１１月１０日（火）　１３：３０～１３：４０');
});

test('枠・休憩を足すと後ろがずれ、×の時間帯や時間外に入ったことが分かる', () => {
  const bands = [Core.parseBand('11月10日(火)13:30～14:00'), Core.parseBand('11月10日(火)14:00～14:30')];
  const slots = Core.buildSlots(bands, { minutes: 10 });
  slots.forEach((s, i) => { s.sid = `1-A-${i + 1}`; });
  // 13:40 の前に休憩 → 13:40 以降が10分ずれる
  const brk = Core.insertSlot(slots, bands, slots[1], 'break', 10);
  assert.equal(brk.start, Core.toMin('13:40'));
  assert.deepEqual(slots.map((s) => Core.toHM(s.start)), ['13:30', '13:40', '13:50', '14:00', '14:10', '14:20', '14:30']);
  const s3 = slots.find((x) => x.sid === '1-A-3'); // 元 13:50 → 14:00（次の時間帯）
  assert.equal(s3.bandKey, '11/10 14:00');
  const last = slots[slots.length - 1];
  assert.equal(Core.bandsOfSlot(bands, last).outside, true, '14:30以降は時間外');
  assert.equal(Core.bandsOfSlot(bands, s3).outside, false);
  // 最後に枠を追加
  const add = Core.insertSlot(slots, bands, null, 'free', 10, '11/10');
  assert.equal(Core.toHM(add.start), '14:40');
  // 休憩を消すと元に戻る
  Core.removeSlot(slots, bands, brk);
  assert.equal(Core.toHM(slots.find((x) => x.sid === '1-A-3').start), '13:50');
  assert.equal(slots.find((x) => x.sid === '1-A-3').bandKey, '11/10 13:30');
  // 時間帯をまたぐ枠は両方の時間帯を見る
  const cross = { month: 11, day: 10, start: Core.toMin('13:55'), end: Core.toMin('14:05') };
  assert.deepEqual(Core.bandsOfSlot(bands, cross).keys, ['11/10 13:30', '11/10 14:00']);
});
