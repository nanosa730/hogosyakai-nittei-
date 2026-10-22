/*
 * Wordの差し込み：テンプレートの【氏名】などを置き換え、1人1ページで全員分を1つの文書にする。
 * Wordは文字を細かい部品（run）に分けて保存するので、段落の中の文字をつないでから目印を探す。
 */
(function (root) {
  'use strict';
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const MARK_RE = /【([^【】]{1,10})】/g;

  function ownTexts(p) {
    // この段落に直接属する w:t（テキストボックス内の段落は別に処理される）
    return [...p.getElementsByTagNameNS(W, 't')].filter((t) => closestP(t) === p);
  }
  function closestP(n) {
    for (let x = n.parentNode; x; x = x.parentNode) if (x.namespaceURI === W && x.localName === 'p') return x;
    return null;
  }

  // 段落内の目印を置き換える。知らない目印はそのまま残し、名前を unknown に入れる
  function fillParagraph(p, fields, unknown) {
    const ts = ownTexts(p);
    if (!ts.length) return;
    const texts = ts.map((t) => t.textContent);
    const all = texts.join('');
    if (!all.includes('【')) return;
    const hits = [...all.matchAll(MARK_RE)].filter((m) => {
      if (Object.prototype.hasOwnProperty.call(fields, m[1])) return true;
      unknown.add(m[1]);
      return false;
    });
    // 後ろから置き換える（前の位置がずれないように）
    for (const m of hits.reverse()) {
      const from = m.index, to = m.index + m[0].length;
      let pos = 0, si = -1, so = 0, ei = -1, eo = 0;
      for (let i = 0; i < texts.length; i++) {
        const len = texts[i].length;
        if (si < 0 && from < pos + len) { si = i; so = from - pos; }
        if (ei < 0 && to <= pos + len) { ei = i; eo = to - pos; break; }
        pos += len;
      }
      const value = String(fields[m[1]]);
      if (si === ei) {
        texts[si] = texts[si].slice(0, so) + value + texts[si].slice(eo);
      } else {
        texts[si] = texts[si].slice(0, so) + value;
        for (let i = si + 1; i < ei; i++) texts[i] = '';
        texts[ei] = texts[ei].slice(eo);
      }
    }
    ts.forEach((t, i) => {
      if (t.textContent !== texts[i]) {
        t.textContent = texts[i];
        t.setAttribute('xml:space', 'preserve');
      }
    });
  }

  function pageBreakParagraph(doc) {
    const p = doc.createElementNS(W, 'w:p');
    const r = doc.createElementNS(W, 'w:r');
    const br = doc.createElementNS(W, 'w:br');
    br.setAttributeNS(W, 'w:type', 'page');
    r.appendChild(br);
    p.appendChild(r);
    return p;
  }

  function setPageBreakBefore(doc, p) {
    let pPr = [...p.childNodes].find((n) => n.namespaceURI === W && n.localName === 'pPr');
    if (!pPr) { pPr = doc.createElementNS(W, 'w:pPr'); p.insertBefore(pPr, p.firstChild); }
    if ([...pPr.childNodes].some((n) => n.localName === 'pageBreakBefore')) return;
    const el = doc.createElementNS(W, 'w:pageBreakBefore');
    // pPr の子の順番は決まっているので、pStyle/keepNext/keepLines の後ろに入れる
    const after = [...pPr.childNodes].filter((n) => ['pStyle', 'keepNext', 'keepLines'].includes(n.localName)).pop();
    pPr.insertBefore(el, after ? after.nextSibling : pPr.firstChild);
  }

  /**
   * @param {ArrayBuffer} template テンプレートのdocx
   * @param {Array<Object>} records 1人ずつの差し込み値
   * @returns {Promise<{blob: Blob, unknown: string[], found: string[]}>}
   */
  async function mergeDocx(template, records) {
    const zip = await JSZip.loadAsync(template);
    const file = zip.file('word/document.xml');
    if (!file) throw new Error('Wordファイル（.docx）ではないようです');
    const xml = await file.async('string');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('Wordファイルを読み取れませんでした');
    const body = doc.getElementsByTagNameNS(W, 'body')[0];
    const children = [...body.childNodes].filter((n) => n.nodeType === 1);
    const sectPr = children.length && children[children.length - 1].localName === 'sectPr' ? children.pop() : null;

    const found = new Set();
    for (const p of doc.getElementsByTagNameNS(W, 'p')) for (const m of ownTexts(p).map((t) => t.textContent).join('').matchAll(MARK_RE)) found.add(m[1]);

    children.forEach((c) => body.removeChild(c));
    const unknown = new Set();
    records.forEach((fields, idx) => {
      children.forEach((c, ci) => {
        const copy = c.cloneNode(true);
        if (idx > 0 && ci === 0) {
          // 2人目からは改ページ。先頭が段落なら「段落前で改ページ」にして空白ページを防ぐ
          if (copy.localName === 'p') setPageBreakBefore(doc, copy);
          else body.insertBefore(pageBreakParagraph(doc), sectPr);
        }
        const ps = copy.localName === 'p' ? [copy] : [];
        ps.push(...copy.getElementsByTagNameNS(W, 'p'));
        ps.forEach((p) => fillParagraph(p, fields, unknown));
        body.insertBefore(copy, sectPr);
      });
    });
    // 図形などのIDの重複はWordが自動で直すので、ここでは触らない
    zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
    const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    return { blob, unknown: [...unknown], found: [...found] };
  }

  root.Docx = { mergeDocx };
})(typeof self !== 'undefined' ? self : this);
