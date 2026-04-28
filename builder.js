/**
 * builder.js — The Red Box Newsletter Builder
 *
 * Generates .docx files matching the reference document (All_in_One_Special___24_10.docx)
 * with forensic accuracy on every measurement, font, border, spacing and colour.
 *
 * Reference measurements (all in DXA unless noted):
 *  Page:        US Letter 12240×15840, margins 1440 all sides
 *  Font:        Verdana throughout (doc default); Arial for summary Read More links
 *  Section hdg: sz=27 (13.5pt), bold, color 0B5394, align=both, no spacing override
 *  Article box: tblStyle "a", width 9350, borders single sz=4 #000000 all sides
 *  Global box:  width 9330, cell borders single sz=6 #000000, margins top/bottom=0 left/right=100
 *  Body text:   sz=20 (10pt), align=both, no explicit spacing (inherits doc default = 0/0)
 *  Empty para:  plain empty paragraph between body paragraphs (no spacing override)
 *  Hyperlinks:  rStyle "Hyperlink" (color 0000FF, underline from style) — no explicit font
 *  Summary tbl: TableGrid style, width auto, gridCols 3403/2601/3310 = 9314 total
 *    Cat header: gridSpan=3, borders sz=18 auto, centred, bold, run sz=28, NO shading
 *    2-col row:  headline gridSpan=2 w=6004 borders(outer sz=18, bottom sz=4, right nil)
 *                vAlign=center; read more w=3310 Arial sz=21 rStyle=Hyperlink left=nil
 *    3-col row:  headline w=3403 (outer sz=18, bottom sz=2, right nil) vAlign=center
 *                company w=2601 centred (outer sz=18, bottom sz=2, left/right nil) vAlign=center
 *                read more w=3310 (outer sz=18, bottom sz=2, left nil) Arial sz=21 rStyle=Hyperlink
 *  Supp heading: sz=27 bold color 0B5394 (same as section headings)
 *  Cat subhead:  sz=27 bold color 0B5394 (inside detail section)
 *  Article hdg:  bold, sz=20, align=both (no size change from body)
 *
 * Security:
 *  - Input sanitised via sanitiseText (no innerHTML, control chars stripped)
 *  - HTTPS-only URLs, private-IP/SSRF guard on all link fields
 *  - Strict length caps on all inputs
 *  - 100% client-side generation — no data leaves the browser
 *  - No eval(), no dynamic script injection
 */

'use strict';

/* ── CONSTANTS ── */
const MAX_BODY_WORDS = 220;
const MAX_TITLE_LEN  = 300;
const MAX_SOURCE_LEN = 100;
const MAX_URL_LEN    = 2048;
const MAX_CAT_LEN    = 80;

// RFC-reserved IP ranges — SSRF guard
const PRIVATE_IP_RE = /^(0\.|10\.|100\.(6[4-9]|[7-9]\d|1[01]\d|12[0-7])\.|127\.|169\.254\.|172\.(1[6-9]|2\d|3[01])\.|192\.168\.|::1$|[fF][cCdD][0-9a-fA-F]{2}:)/;

/* ── ID COUNTER ── */
let _uid = 0;
function uid() { return `e${++_uid}`; }

/* ── URL VALIDATION ── */
function validateUrl(raw) {
  if (!raw || raw.trim() === '') return { ok: true, url: '' };
  const s = raw.trim();
  if (s.length > MAX_URL_LEN) return { ok: false, reason: 'URL too long' };
  let parsed;
  try { parsed = new URL(s); } catch { return { ok: false, reason: 'Invalid URL format' }; }
  if (parsed.protocol !== 'https:') return { ok: false, reason: 'Only HTTPS URLs are allowed' };
  if (PRIVATE_IP_RE.test(parsed.hostname)) return { ok: false, reason: 'Private/reserved IP not allowed' };
  return { ok: true, url: s };
}

/* ── INPUT SANITISATION ── */
function sanitiseText(str, maxLen) {
  if (typeof str !== 'string') return '';
  let s = str.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
  if (maxLen) s = s.slice(0, maxLen);
  return s.trim();
}

/* ── WORD COUNT / TRUNCATION ── */
function wordCount(str) {
  return str.trim() === '' ? 0 : str.trim().split(/\s+/).length;
}

function truncateWords(str, limit) {
  const normalised = str.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const paragraphs = normalised.split(/\n{2,}/);
  const result = [];
  let remaining = limit;
  for (const para of paragraphs) {
    if (remaining <= 0) break;
    const words = para.trim().split(/\s+/).filter(Boolean);
    if (!words.length) continue;
    if (words.length <= remaining) {
      result.push(words.join(' '));
      remaining -= words.length;
    } else {
      result.push(words.slice(0, remaining).join(' ') + '….');
      remaining = 0;
    }
  }
  const totalWords = paragraphs.flatMap(p => p.trim().split(/\s+/).filter(Boolean)).length;
  const joined = result.join('\n\n');
  return totalWords > limit ? joined : normalised.trim();
}

/* ── TOAST ── */
let _toastTimer = null;
function toast(msg, isError = false) {
  let el = document.getElementById('toast-el');
  if (!el) {
    el = document.createElement('div');
    el.id = 'toast-el';
    el.className = 'toast';
    el.setAttribute('role', 'status');
    el.setAttribute('aria-live', 'polite');
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.className = 'toast' + (isError ? ' error' : '');
  void el.offsetWidth;
  el.classList.add('show');
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => el.classList.remove('show'), 3400);
}

/* ─────────────────────────────────────────────────────────────────────────
   FORM — ARTICLE ENTRY FACTORY
   ───────────────────────────────────────────────────────────────────────── */
function createArticleEntry(isSummary = false) {
  const tpl = document.getElementById(isSummary ? 'tpl-summary-article' : 'tpl-article');
  const clone = tpl.content.cloneNode(true);
  const entry = clone.querySelector('.article-entry');
  entry.dataset.id = uid();

  if (!isSummary) {
    const textarea = entry.querySelector('.field-body');
    const wcEl     = entry.querySelector('.wc-num');
    const wcWrap   = entry.querySelector('.word-count');
    textarea.addEventListener('input', () => {
      const wc = wordCount(textarea.value);
      wcEl.textContent = wc;
      wcWrap.classList.toggle('over', wc > MAX_BODY_WORDS);
    });
    const hasRelated  = entry.querySelector('.field-has-related');
    const relatedWrap = entry.querySelector('.field-related-wrap');
    hasRelated.addEventListener('change', () => { relatedWrap.hidden = !hasRelated.checked; });
  }

  entry.querySelector('.btn-remove').addEventListener('click', () => {
    const list = entry.parentElement;
    entry.remove();
    renumberEntries(list);
  });
  return entry;
}

function renumberEntries(list) {
  list.querySelectorAll('.article-entry').forEach((e, i) => {
    const num = e.querySelector('.article-num');
    if (num) num.textContent = `Article ${i + 1}`;
  });
}

function appendArticle(listEl, isSummary = false) {
  const entry = createArticleEntry(isSummary);
  listEl.appendChild(entry);
  renumberEntries(listEl);
  entry.querySelector('.field-headline').focus();
}

/* ─────────────────────────────────────────────────────────────────────────
   FORM — CATEGORY BLOCK FACTORY
   ───────────────────────────────────────────────────────────────────────── */
function createCategoryBlock(isSummary = false) {
  const tpl = document.getElementById('tpl-category');
  const clone = tpl.content.cloneNode(true);
  const block = clone.querySelector('.cat-block');
  block.dataset.catId = uid();
  const articleList = block.querySelector('.cat-articles');
  block.querySelector('.btn-add-in-cat').addEventListener('click', () => appendArticle(articleList, isSummary));
  block.querySelector('.btn-remove-cat').addEventListener('click', () => block.remove());
  return block;
}

function appendCategory(containerEl, isSummary = false) {
  const block = createCategoryBlock(isSummary);
  containerEl.appendChild(block);
  block.querySelector('.cat-name-input').focus();
}

/* ─────────────────────────────────────────────────────────────────────────
   FORM — COLLECT DATA
   ───────────────────────────────────────────────────────────────────────── */
function collectArticles(listEl, isSummary = false) {
  const articles = [];
  listEl.querySelectorAll('.article-entry').forEach(entry => {
    if (isSummary) {
      articles.push({
        headline : sanitiseText(entry.querySelector('.field-headline').value, MAX_TITLE_LEN),
        company  : sanitiseText(entry.querySelector('.field-company').value, MAX_SOURCE_LEN),
        source   : sanitiseText(entry.querySelector('.field-source').value, MAX_SOURCE_LEN),
        link     : entry.querySelector('.field-link').value.trim(),
      });
    } else {
      const bodyRaw   = entry.querySelector('.field-body').value;
      const bodyClean = sanitiseText(bodyRaw, 5000);
      const bodyTrunc = truncateWords(bodyClean, MAX_BODY_WORDS);
      const hasRelated = entry.querySelector('.field-has-related')?.checked;
      articles.push({
        headline        : sanitiseText(entry.querySelector('.field-headline').value, MAX_TITLE_LEN),
        body            : bodyTrunc,
        source          : sanitiseText(entry.querySelector('.field-source').value, MAX_SOURCE_LEN),
        link            : entry.querySelector('.field-link').value.trim(),
        hasRelated      : !!hasRelated,
        relatedHeadline : hasRelated ? sanitiseText(entry.querySelector('.field-related-headline')?.value, MAX_TITLE_LEN) : '',
        relatedSource   : hasRelated ? sanitiseText(entry.querySelector('.field-related-source')?.value, MAX_SOURCE_LEN) : '',
        relatedLink     : hasRelated ? (entry.querySelector('.field-related-link')?.value.trim() || '') : '',
      });
    }
  });
  return articles.filter(a => a.headline);
}

function collectCategories(containerEl, isSummary = false) {
  const cats = [];
  containerEl.querySelectorAll('.cat-block').forEach(block => {
    const name = sanitiseText(block.querySelector('.cat-name-input').value, MAX_CAT_LEN);
    if (!name) return;
    cats.push({ name, articles: collectArticles(block.querySelector('.cat-articles'), isSummary) });
  });
  return cats;
}

function collectFormData() {
  return {
    date        : document.getElementById('nl-date').value,
    topStory    : collectArticles(document.getElementById('articles-top-story')),
    global      : collectArticles(document.getElementById('articles-global')),
    suppSummary : collectCategories(document.getElementById('supp-summary-categories'), true),
    suppDetail  : collectCategories(document.getElementById('supp-detail-categories'), false),
  };
}

/* ─────────────────────────────────────────────────────────────────────────
   VALIDATION
   ───────────────────────────────────────────────────────────────────────── */
function validateData(data) {
  const errors = [];
  function checkUrls(articles) {
    articles.forEach((a, i) => {
      const r = validateUrl(a.link);
      if (!r.ok) errors.push(`Article ${i + 1} link: ${r.reason}`);
      if (a.hasRelated) {
        const r2 = validateUrl(a.relatedLink);
        if (!r2.ok) errors.push(`Article ${i + 1} related link: ${r2.reason}`);
      }
    });
  }
  checkUrls(data.topStory);
  checkUrls(data.global);
  data.suppSummary.forEach(cat => cat.articles.forEach((a, i) => {
    const r = validateUrl(a.link);
    if (!r.ok) errors.push(`Summary "${cat.name}" article ${i + 1}: ${r.reason}`);
  }));
  data.suppDetail.forEach(cat => checkUrls(cat.articles));
  return errors;
}

/* ─────────────────────────────────────────────────────────────────────────
   PREVIEW
   ───────────────────────────────────────────────────────────────────────── */
function buildPreviewHtml(data) {
  const frag = document.createDocumentFragment();
  const wrap = document.createElement('div');
  wrap.className = 'preview-doc';

  function el(tag, cls, text) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined) e.textContent = text;
    return e;
  }

  function sectionLabel(text) {
    wrap.appendChild(el('div', 'preview-section-label', text));
  }

  function renderFullArticles(articles) {
    articles.forEach(a => {
      const box = el('div', 'preview-article-box', '');
      box.appendChild(el('div', 'preview-article-title', a.headline));
      if (a.body) {
        a.body.replace(/\r\n/g, '\n').split(/\n{2,}/).filter(p => p.trim()).forEach(p => {
          box.appendChild(el('p', 'preview-article-body', p.replace(/\n/g, ' ').trim()));
        });
      }
      if (a.source || a.link) box.appendChild(el('div', 'preview-read-more', `Read More: ${a.source}`));
      if (a.hasRelated && a.relatedHeadline) {
        box.appendChild(el('div', 'preview-related', `Related: ${a.relatedHeadline}`));
        if (a.relatedSource) box.appendChild(el('div', 'preview-read-more', `Read More: ${a.relatedSource}`));
      }
      wrap.appendChild(box);
    });
  }

  if (data.date) {
    const d = new Date(data.date + 'T00:00:00');
    wrap.appendChild(el('div', 'preview-date',
      d.toLocaleDateString('en-GB', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })));
  }
  if (data.topStory.length) { sectionLabel('01 thing you need to know to start your day'); renderFullArticles(data.topStory); }
  if (data.global.length)   { sectionLabel('01 global updates to keep an eye on');         renderFullArticles(data.global); }

  if (data.suppSummary.length) {
    sectionLabel('Supplementary News – In Summary');
    const table = document.createElement('table');
    table.className = 'preview-table';
    data.suppSummary.forEach(cat => {
      const hr = table.insertRow(); hr.className = 'preview-cat-header';
      const td = hr.insertCell(); td.colSpan = 3; td.textContent = cat.name;
      cat.articles.forEach(a => {
        const row = table.insertRow();
        const hasTag = a.company && a.company.trim();
        if (hasTag) {
          row.insertCell().textContent = a.headline;
          row.insertCell().textContent = a.company;
        } else {
          const hCell = row.insertCell(); hCell.colSpan = 2; hCell.textContent = a.headline;
        }
        row.insertCell().textContent = `Read More: ${a.source}`;
      });
    });
    wrap.appendChild(table);
  }

  if (data.suppDetail.length) {
    sectionLabel('Supplementary News – In Detail');
    data.suppDetail.forEach(cat => {
      wrap.appendChild(el('div', 'preview-cat-name', cat.name));
      renderFullArticles(cat.articles);
    });
  }

  frag.appendChild(wrap);
  return frag;
}

/* ─────────────────────────────────────────────────────────────────────────
   DOCX GENERATION
   Forensically matched to All_in_One_Special___24_10.docx
   ───────────────────────────────────────────────────────────────────────── */
async function generateDocx(data) {
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, BorderStyle, WidthType, VerticalAlign,
    ExternalHyperlink, UnderlineType
  } = docx;

  /* ── Reference-exact constants ── */
  const PAGE_W  = 12240;   // US Letter width
  const PAGE_H  = 15840;   // US Letter height
  const MARGIN  = 1440;    // 1 inch all sides

  // Fonts
  const F_MAIN  = 'Verdana';   // body, headings, article titles
  const F_LINK  = 'Arial';     // Read More links in summary table (reference uses Arial)

  // Sizes (half-points)
  const SZ_BODY  = 20;   // 10pt  — body text, article titles
  const SZ_HEAD  = 27;   // 13.5pt — section headings, supplementary headings, cat subheadings
  const SZ_CAT   = 28;   // 14pt  — category header text in summary table
  const SZ_LINK  = 21;   // 10.5pt — Read More links in summary table (Arial)

  // Colours
  const C_BLUE  = '0B5394';  // section/supplementary/cat headings
  const C_BLACK = '000000';

  // Table widths (DXA)
  const TBL1_W  = 9350;   // article box (top story)
  const TBL2_W  = 9330;   // article box (global)
  const SUM_W   = 9314;   // summary table total
  const COL1    = 3403;   // summary col 1 (headline, 3-col)
  const COL2    = 2601;   // summary col 2 (company/sector)
  const COL3    = 3310;   // summary col 3 (read more)
  const COL2C1  = 6004;   // summary headline, 2-col (=COL1+COL2)
  const COL2C2  = COL3;   // summary read more, 2-col

  /* ── Border helpers ── */
  // Article box borders: single sz=4 black
  const bk4  = (sz=4)  => ({ style: BorderStyle.SINGLE, size: sz, color: C_BLACK, space: 0 });
  const bk6  = ()      => ({ style: BorderStyle.SINGLE, size: 6,  color: C_BLACK, space: 0 });
  const nil_ = ()      => ({ style: BorderStyle.NIL });
  // Summary table borders use 'auto' colour (Word default) sz=18 outer, sz=4/2 inner
  const au18 = ()      => ({ style: BorderStyle.SINGLE, size: 18, color: 'auto', space: 0 });
  const au4  = ()      => ({ style: BorderStyle.SINGLE, size: 4,  color: 'auto', space: 0 });
  const au2  = ()      => ({ style: BorderStyle.SINGLE, size: 2,  color: 'auto', space: 0 });

  const BOX1_BORDERS = { top: bk4(), bottom: bk4(), left: bk4(), right: bk4(), insideH: bk4(), insideV: bk4() };
  const BOX2_BORDERS = { top: bk6(), bottom: bk6(), left: bk6(), right: bk6(), insideH: { style: BorderStyle.NIL }, insideV: { style: BorderStyle.NIL } };

  /* ── Paragraph / run helpers ── */
  // Standard Verdana run — only sets what reference sets explicitly
  function vRun(text, { bold = false, sz = SZ_BODY, color = null } = {}) {
    return new TextRun({
      text,
      font: F_MAIN,
      size: sz,
      bold,
      ...(color ? { color } : {}),
    });
  }

  // Paragraph with alignment=both and NO explicit spacing (matches reference)
  function vPara(children, { before = 0, after = 0, align = AlignmentType.BOTH } = {}) {
    const ch = Array.isArray(children) ? children : [children];
    return new Paragraph({
      alignment: align,
      spacing: { before, after },
      children: ch,
    });
  }

  // Empty paragraph (paragraph separator — no spacing, matches reference)
  function emptyPara() {
    return new Paragraph({ spacing: { before: 0, after: 0 }, children: [] });
  }

  // Horizontal rule paragraph — used as visual separator between sections
  // Rendered as a paragraph with a bottom border (single, sz=6, black)
  function hrulePara() {
    return new Paragraph({
      spacing: { before: 0, after: 0 },
      border: {
        bottom: { style: BorderStyle.SINGLE, size: 6, color: C_BLACK, space: 1 },
      },
      children: [],
    });
  }

  // Hyperlink using rStyle "Hyperlink" — exactly as reference (no explicit font/color on run)
  function makeHyperlink(text, url) {
    const v = validateUrl(url);
    const safeUrl = v.ok && v.url ? v.url : null;
    // Explicitly set font and size to match body text (Verdana 10pt = sz 20)
    // so Read More links are always the same size/font as the article body.
    const linkRun = new TextRun({
      text,
      font: F_MAIN,
      size: SZ_BODY,
      style: 'Hyperlink',  // provides color 0000FF + underline from character style
    });
    if (safeUrl) {
      return new ExternalHyperlink({ link: safeUrl, children: [linkRun] });
    }
    return linkRun;
  }

  // Read More link in summary table — Arial sz=21, rStyle=Hyperlink (matches reference)
  function summaryReadMore(text, url) {
    const v = validateUrl(url);
    const safeUrl = v.ok && v.url ? v.url : null;
    const linkRun = new TextRun({
      text,
      font: F_LINK,
      size: SZ_LINK,
      style: 'Hyperlink',
    });
    if (safeUrl) {
      return new ExternalHyperlink({ link: safeUrl, children: [linkRun] });
    }
    return linkRun;
  }

  /* ── Section heading paragraph ──
     Bold, sz=27, color 0B5394, align=both, no spacing override */
  function sectionHeading(text) {
    return vPara(vRun(text, { bold: true, sz: SZ_HEAD, color: C_BLUE }));
  }

  /* ── Supplementary section heading (same style as section heading) ── */
  function suppHeading(text) {
    return vPara(vRun(text, { bold: true, sz: SZ_HEAD, color: C_BLUE }));
  }

  /* ── Article box internal paragraphs ──
     Returns array of Paragraph objects for one article */
  function articleParas(article) {
    const paras = [];

    // Headline: bold, sz=20, align=both
    paras.push(vPara(vRun(article.headline, { bold: true })));
    paras.push(emptyPara());

    // Body: split on double newlines to preserve paragraph structure
    if (article.body) {
      const normalised = article.body.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      const bodyParas = normalised.split(/\n{2,}/).map(p => p.replace(/\n/g, ' ').trim()).filter(Boolean);
      bodyParas.forEach((p, i) => {
        paras.push(vPara(vRun(p)));
        if (i < bodyParas.length - 1) paras.push(emptyPara());
      });
      paras.push(emptyPara());
    }

    // Read More hyperlink
    if (article.source || article.link) {
      paras.push(new Paragraph({
        alignment: AlignmentType.BOTH,
        spacing: { before: 0, after: 0 },
        children: [makeHyperlink(`Read More: ${article.source || article.link}`, article.link)],
      }));
    }

    // Related article
    if (article.hasRelated && article.relatedHeadline) {
      paras.push(emptyPara());
      paras.push(vPara(vRun(`Related: ${article.relatedHeadline}`, { bold: true })));
      if (article.relatedSource || article.relatedLink) {
        paras.push(new Paragraph({
          alignment: AlignmentType.BOTH,
          spacing: { before: 0, after: 0 },
          children: [makeHyperlink(`Read More: ${article.relatedSource || article.relatedLink}`, article.relatedLink)],
        }));
      }
    }

    return paras;
  }

  /* ── Article box table (top story) ──
     tblStyle "a", width 9350, borders single sz=4 #000000 */
  function articleBox1(article) {
    return new Table({
      style: 'a',
      width: { size: TBL1_W, type: WidthType.DXA },
      columnWidths: [TBL1_W],
      borders: BOX1_BORDERS,
      rows: [new TableRow({
        children: [new TableCell({
          width: { size: TBL1_W, type: WidthType.DXA },
          borders: BOX1_BORDERS,
          children: articleParas(article),
        })],
      })],
    });
  }

  /* ── Article box table (global) ──
     No tblStyle, width 9330, cell borders single sz=6 #000000,
     cell margins top/bottom=0 left/right=100 */
  function articleBox2(article) {
    const cellBorders = { top: bk6(), bottom: bk6(), left: bk6(), right: bk6() };
    return new Table({
      width: { size: TBL2_W, type: WidthType.DXA },
      columnWidths: [TBL2_W],
      borders: BOX2_BORDERS,
      rows: [new TableRow({
        children: [new TableCell({
          width: { size: TBL2_W, type: WidthType.DXA },
          borders: cellBorders,
          margins: { top: 0, bottom: 0, left: 100, right: 100 },
          children: articleParas(article),
        })],
      })],
    });
  }

  /* ── Summary table ──
     TableGrid style, columnWidths [3403, 2601, 3310]
     Category header: gridSpan=3, borders sz=18 auto, centred, bold sz=28, NO shading
     2-col article: headline gridSpan=2 w=6004, read more w=3310 Arial
     3-col article: headline w=3403, company centred w=2601, read more w=3310 Arial */
  function summaryTable(categories) {
    const rows = [];

    categories.forEach(cat => {
      // ── Category header row ──
      // Borders: sz=18 auto all sides, no shading, text centred bold sz=28
      rows.push(new TableRow({
        children: [new TableCell({
          width:      { size: SUM_W, type: WidthType.DXA },
          columnSpan: 3,
          borders:    { top: au18(), bottom: au18(), left: au18(), right: au18() },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing:   { before: 0, after: 0 },
            children:  [vRun(cat.name, { bold: true, sz: SZ_CAT })],
          })],
        })],
      }));

      // ── Article rows ──
      cat.articles.forEach(a => {
        const hasTag = a.company && a.company.trim();
        const rmText = `Read More: ${a.source || ''}`;

        if (hasTag) {
          /* 3-column row
             col1: outer sz=18, bottom sz=2, right=nil, vAlign=center
             col2: outer sz=18, bottom sz=2, left/right=nil, centred, vAlign=center
             col3: outer sz=18, bottom sz=2, left=nil, vAlign=center */
          rows.push(new TableRow({
            children: [
              new TableCell({
                width:   { size: COL1, type: WidthType.DXA },
                borders: { top: au18(), bottom: au2(), left: au18(), right: nil_() },
                verticalAlign: VerticalAlign.CENTER,
                children: [vPara(vRun(a.headline))],
              }),
              new TableCell({
                width:   { size: COL2, type: WidthType.DXA },
                borders: { top: au18(), bottom: au2(), left: nil_(), right: nil_() },
                verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing:   { before: 0, after: 0 },
                  children:  [vRun(a.company.trim())],
                })],
              }),
              new TableCell({
                width:   { size: COL3, type: WidthType.DXA },
                borders: { top: au18(), bottom: au2(), left: nil_(), right: au18() },
                verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({
                  alignment: AlignmentType.BOTH,
                  spacing:   { before: 0, after: 0 },
                  children:  [summaryReadMore(rmText, a.link)],
                })],
              }),
            ],
          }));
        } else {
          /* 2-column row
             col1+2: gridSpan=2, outer sz=18, bottom sz=4, right=nil, vAlign=center
             col3:   outer sz=18, bottom sz=4, left=nil, vAlign=center */
          rows.push(new TableRow({
            children: [
              new TableCell({
                width:      { size: COL2C1, type: WidthType.DXA },
                columnSpan: 2,
                borders:    { top: au18(), bottom: au4(), left: au18(), right: nil_() },
                verticalAlign: VerticalAlign.CENTER,
                children:   [vPara(vRun(a.headline))],
              }),
              new TableCell({
                width:   { size: COL2C2, type: WidthType.DXA },
                borders: { top: au18(), bottom: au4(), left: nil_(), right: au18() },
                verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({
                  alignment: AlignmentType.BOTH,
                  spacing:   { before: 0, after: 0 },
                  children:  [summaryReadMore(rmText, a.link)],
                })],
              }),
            ],
          }));
        }
      });
    });

    return new Table({
      style:        'TableGrid',
      width:        { size: 0, type: WidthType.AUTO },
      columnWidths: [COL1, COL2, COL3],
      rows,
    });
  }

  /* ── Detail section — plain paragraphs ──
     Category name: bold sz=27 color 0B5394 (same as section heading)
     Article headline: bold sz=20 align=both
     Body: sz=20 align=both, paragraphs separated by empty paragraphs
     Read More: Hyperlink style */
  function detailSection(categories) {
    const children = [];
    categories.forEach(cat => {
      children.push(vPara(vRun(cat.name, { bold: true, sz: SZ_HEAD, color: C_BLUE })));
      children.push(emptyPara());

      cat.articles.forEach(a => {
        // Article headline
        children.push(vPara(vRun(a.headline, { bold: true })));
        children.push(emptyPara());

        // Body paragraphs
        if (a.body) {
          const normalised = a.body.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
          const bodyParas  = normalised.split(/\n{2,}/).map(p => p.replace(/\n/g, ' ').trim()).filter(Boolean);
          bodyParas.forEach((p, i) => {
            children.push(vPara(vRun(p)));
            if (i < bodyParas.length - 1) children.push(emptyPara());
          });
          children.push(emptyPara());
        }

        // Read More
        if (article_has_readmore(a)) {
          children.push(new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing:   { before: 0, after: 0 },
            children:  [makeHyperlink(`Read More: ${a.source || a.link}`, a.link)],
          }));
          children.push(emptyPara());
        }

        // Related
        if (a.hasRelated && a.relatedHeadline) {
          children.push(vPara(vRun(`Related: ${a.relatedHeadline}`, { bold: true })));
          if (a.relatedSource || a.relatedLink) {
            children.push(new Paragraph({
              alignment: AlignmentType.BOTH,
              spacing:   { before: 0, after: 0 },
              children:  [makeHyperlink(`Read More: ${a.relatedSource || a.relatedLink}`, a.relatedLink)],
            }));
          }
          children.push(emptyPara());
        }
      });
    });
    return children;
  }

  function article_has_readmore(a) {
    return !!(a.source || a.link);
  }

  /* ── Assemble document ── */
  const docChildren = [];

  // Date line — bold Verdana, no size override (uses default 10pt)
  if (data.date) {
    const d = new Date(data.date + 'T00:00:00');
    const fmt = d.toLocaleDateString('en-GB', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
    docChildren.push(vPara(vRun(fmt, { bold: true })));
    docChildren.push(emptyPara());
  }

  // ── 01 thing you need to know to start your day ──
  if (data.topStory.length) {
    docChildren.push(sectionHeading('01 thing you need to know to start your day'));
    docChildren.push(emptyPara());
    data.topStory.forEach((a, i) => {
      docChildren.push(articleBox1(a));
      if (i < data.topStory.length - 1) docChildren.push(emptyPara());
    });
    docChildren.push(emptyPara());
    docChildren.push(hrulePara());
    docChildren.push(emptyPara());
  }

  // ── 01 global updates to keep an eye on ──
  if (data.global.length) {
    docChildren.push(sectionHeading('01 global updates to keep an eye on'));
    docChildren.push(emptyPara());
    data.global.forEach((a, i) => {
      docChildren.push(articleBox2(a));
      if (i < data.global.length - 1) docChildren.push(emptyPara());
    });
    docChildren.push(emptyPara());
    docChildren.push(hrulePara());
    docChildren.push(emptyPara());
  }

  // ── Supplementary News – In Summary (order: Summary before Detail) ──
  if (data.suppSummary.length) {
    docChildren.push(suppHeading('Supplementary News – In Summary'));
    docChildren.push(emptyPara());
    docChildren.push(summaryTable(data.suppSummary));
    docChildren.push(emptyPara());
    docChildren.push(hrulePara());
    docChildren.push(emptyPara());
  }

  // ── Supplementary News – In Detail ──
  if (data.suppDetail.length) {
    docChildren.push(suppHeading('Supplementary News – In Detail'));
    docChildren.push(emptyPara());
    detailSection(data.suppDetail).forEach(c => docChildren.push(c));
    docChildren.push(emptyPara());
  }

  /* ── Document definition ── */
  const doc = new Document({
    creator:     'The Red Box Newsletter Builder',
    description: 'Sri Lanka English-language news digest',

    // Character styles
    styles: {
      characterStyles: [{
        id:   'Hyperlink',
        name: 'Hyperlink',
        run: {
          color:     '0000FF',
          underline: { type: UnderlineType.SINGLE },
        },
      }],
    },

    sections: [{
      properties: {
        page: {
          size:   { width: PAGE_W, height: PAGE_H },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
        },
      },
      children: docChildren,
    }],
  });

  return Packer.toBlob(doc);
}

/* ─────────────────────────────────────────────────────────────────────────
   DOWNLOAD
   ───────────────────────────────────────────────────────────────────────── */
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; a.rel = 'noopener noreferrer';
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
}

function safeFilename(date) {
  const d = date ? date.replace(/-/g, '') : 'newsletter';
  return `red-box-newsletter-${d}.docx`;
}

/* ─────────────────────────────────────────────────────────────────────────
   GENERATE FLOW
   ───────────────────────────────────────────────────────────────────────── */
async function runGenerate() {
  const data   = collectFormData();
  const errors = validateData(data);
  if (errors.length) { toast('Fix these issues:\n• ' + errors.slice(0, 3).join('\n• '), true); return; }

  const btn = document.getElementById('btn-generate');
  btn.disabled = true; btn.textContent = 'Generating…';
  try {
    const blob = await generateDocx(data);
    downloadBlob(blob, safeFilename(data.date));
    toast('Document downloaded!');
  } catch (err) {
    console.error('docx generation failed:', err);
    toast('Generation failed — check console for details.', true);
  } finally {
    btn.disabled = false; btn.textContent = 'Generate .docx';
  }
}

/* ─────────────────────────────────────────────────────────────────────────
   WIRE UP
   ───────────────────────────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  // Default date to today
  const today = new Date();
  document.getElementById('nl-date').value =
    `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;

  // Add-article buttons
  document.querySelectorAll('.btn-add[data-target]').forEach(btn => {
    btn.addEventListener('click', () =>
      appendArticle(document.getElementById(`articles-${btn.dataset.target}`), false));
  });

  // Add-category buttons
  document.getElementById('btn-add-supp-cat').addEventListener('click', () =>
    appendCategory(document.getElementById('supp-summary-categories'), true));
  document.getElementById('btn-add-detail-cat').addEventListener('click', () =>
    appendCategory(document.getElementById('supp-detail-categories'), false));

  // Generate
  document.getElementById('btn-generate').addEventListener('click', runGenerate);

  // Preview
  document.getElementById('btn-preview').addEventListener('click', () => {
    const body = document.getElementById('modal-preview-body');
    body.innerHTML = '';
    body.appendChild(buildPreviewHtml(collectFormData()));
    const backdrop = document.getElementById('modal-backdrop');
    backdrop.hidden = false;
    backdrop.removeAttribute('aria-hidden');
  });

  // Modal close
  function closeModal() {
    const b = document.getElementById('modal-backdrop');
    b.hidden = true; b.setAttribute('aria-hidden', 'true');
  }
  document.getElementById('modal-close').addEventListener('click', closeModal);
  document.getElementById('modal-cancel').addEventListener('click', closeModal);
  document.getElementById('modal-backdrop').addEventListener('click', e => { if (e.target === e.currentTarget) closeModal(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });
  document.getElementById('modal-generate').addEventListener('click', () => { closeModal(); runGenerate(); });
});
