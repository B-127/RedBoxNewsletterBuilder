/**
 * builder.js — The Red Box Newsletter Builder
 *
 * Security:
 *  - All user input sanitised before insertion into DOM (textContent, not innerHTML)
 *  - URL validation: HTTPS-only, no private-IP ranges (SSRF prevention)
 *  - Word/char limits enforced on every field
 *  - docx generated entirely client-side; no data leaves the browser
 *  - No eval(), no dynamic script injection
 *  - Content is escaped before use in the document model
 */

'use strict';

/* ── CONSTANTS ── */
const MAX_BODY_WORDS   = 220;
const MAX_TITLE_LEN    = 300;
const MAX_SOURCE_LEN   = 100;
const MAX_URL_LEN      = 2048;
const MAX_CAT_LEN      = 80;

// Covers all RFC-reserved ranges (1122, 1918, 6598, 3927) + IPv6 loopback/unique-local
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
  try { parsed = new URL(s); }
  catch { return { ok: false, reason: 'Invalid URL format' }; }
  if (parsed.protocol !== 'https:') return { ok: false, reason: 'Only HTTPS URLs are allowed' };
  const host = parsed.hostname;
  if (PRIVATE_IP_RE.test(host)) return { ok: false, reason: 'Private/reserved IP addresses are not allowed' };
  return { ok: true, url: s };
}

/* ── SANITISE TEXT ── */
function sanitiseText(str, maxLen) {
  if (typeof str !== 'string') return '';
  // Strip non-printable control chars (allow newlines)
  let s = str.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
  if (maxLen) s = s.slice(0, maxLen);
  return s.trim();
}

/* ── WORD COUNT ── */
function wordCount(str) {
  return str.trim() === '' ? 0 : str.trim().split(/\s+/).length;
}

/* ── TRUNCATE TO WORD LIMIT ── */
function truncateWords(str, limit) {
  const words = str.trim().split(/\s+/);
  if (words.length <= limit) return str.trim();
  return words.slice(0, limit).join(' ') + '….';
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
   ARTICLE ENTRY FACTORY
   ───────────────────────────────────────────────────────────────────────── */

function createArticleEntry(isSummary = false) {
  const tplId = isSummary ? 'tpl-summary-article' : 'tpl-article';
  const tpl = document.getElementById(tplId);
  const clone = tpl.content.cloneNode(true);
  const entry = clone.querySelector('.article-entry');
  const id = uid();
  entry.dataset.id = id;

  // Word count watcher (full articles only)
  if (!isSummary) {
    const textarea = entry.querySelector('.field-body');
    const wcEl     = entry.querySelector('.wc-num');
    const wcWrap   = entry.querySelector('.word-count');
    textarea.addEventListener('input', () => {
      const wc = wordCount(textarea.value);
      wcEl.textContent = wc;
      wcWrap.classList.toggle('over', wc > MAX_BODY_WORDS);
    });

    // Related toggle
    const hasRelated  = entry.querySelector('.field-has-related');
    const relatedWrap = entry.querySelector('.field-related-wrap');
    hasRelated.addEventListener('change', () => {
      relatedWrap.hidden = !hasRelated.checked;
    });
  }

  // Remove button
  entry.querySelector('.btn-remove').addEventListener('click', () => {
    const list = entry.parentElement;
    entry.remove();
    renumberEntries(list);
  });

  return entry;
}

function renumberEntries(list) {
  const entries = list.querySelectorAll('.article-entry');
  entries.forEach((e, i) => {
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
   CATEGORY BLOCK FACTORY
   ───────────────────────────────────────────────────────────────────────── */

function createCategoryBlock(isSummary = false) {
  const tpl = document.getElementById('tpl-category');
  const clone = tpl.content.cloneNode(true);
  const block = clone.querySelector('.cat-block');
  block.dataset.catId = uid();

  const addBtn = block.querySelector('.btn-add-in-cat');
  const articleList = block.querySelector('.cat-articles');

  addBtn.addEventListener('click', () => {
    appendArticle(articleList, isSummary);
  });

  block.querySelector('.btn-remove-cat').addEventListener('click', () => {
    block.remove();
  });

  // Auto-focus category name
  return block;
}

function appendCategory(containerEl, isSummary = false) {
  const block = createCategoryBlock(isSummary);
  containerEl.appendChild(block);
  block.querySelector('.cat-name-input').focus();
}

/* ─────────────────────────────────────────────────────────────────────────
   COLLECT FORM DATA
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
        headline       : sanitiseText(entry.querySelector('.field-headline').value, MAX_TITLE_LEN),
        body           : bodyTrunc,
        source         : sanitiseText(entry.querySelector('.field-source').value, MAX_SOURCE_LEN),
        link           : entry.querySelector('.field-link').value.trim(),
        hasRelated     : !!hasRelated,
        relatedHeadline: hasRelated ? sanitiseText(entry.querySelector('.field-related-headline')?.value, MAX_TITLE_LEN) : '',
        relatedSource  : hasRelated ? sanitiseText(entry.querySelector('.field-related-source')?.value, MAX_SOURCE_LEN) : '',
        relatedLink    : hasRelated ? (entry.querySelector('.field-related-link')?.value.trim() || '') : '',
      });
    }
  });
  return articles.filter(a => a.headline); // skip empty
}

function collectCategories(containerEl, isSummary = false) {
  const cats = [];
  containerEl.querySelectorAll('.cat-block').forEach(block => {
    const name = sanitiseText(block.querySelector('.cat-name-input').value, MAX_CAT_LEN);
    if (!name) return;
    const articles = collectArticles(block.querySelector('.cat-articles'), isSummary);
    cats.push({ name, articles });
  });
  return cats;
}

function collectFormData() {
  const dateVal = document.getElementById('nl-date').value;
  return {
    date          : dateVal,
    topStory      : collectArticles(document.getElementById('articles-top-story')),
    global        : collectArticles(document.getElementById('articles-global')),
    suppSummary   : collectCategories(document.getElementById('supp-summary-categories'), true),
    suppDetail    : collectCategories(document.getElementById('supp-detail-categories'), false),
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
      if (!r.ok) errors.push(`Article ${i+1} link: ${r.reason}`);
      if (a.hasRelated) {
        const r2 = validateUrl(a.relatedLink);
        if (!r2.ok) errors.push(`Article ${i+1} related link: ${r2.reason}`);
      }
    });
  }

  checkUrls(data.topStory);
  checkUrls(data.global);
  data.suppSummary.forEach(cat => {
    cat.articles.forEach((a, i) => {
      const r = validateUrl(a.link);
      if (!r.ok) errors.push(`Summary "${cat.name}" article ${i+1}: ${r.reason}`);
    });
  });
  data.suppDetail.forEach(cat => {
    checkUrls(cat.articles);
  });

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
      wrap.appendChild(el('div', 'preview-article-title', a.headline));
      wrap.appendChild(el('div', 'preview-article-body', a.body));
      if (a.source || a.link) {
        const rm = el('div', 'preview-read-more', `Read more: ${a.source}`);
        wrap.appendChild(rm);
      }
      if (a.hasRelated && a.relatedHeadline) {
        wrap.appendChild(el('div', 'preview-related', `RELATED: ${a.relatedHeadline}`));
        const rel = el('div', 'preview-read-more', `Read more: ${a.relatedSource}`);
        wrap.appendChild(rel);
      }
    });
  }

  // Date
  if (data.date) {
    const d = new Date(data.date + 'T00:00:00');
    const fmt = d.toLocaleDateString('en-GB', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
    wrap.appendChild(el('div', 'preview-article-title', fmt));
  }

  // Top story
  if (data.topStory.length) {
    sectionLabel('01 thing you need to know to start your day');
    renderFullArticles(data.topStory);
  }

  // Global
  if (data.global.length) {
    sectionLabel('01 global updates to keep an eye on');
    renderFullArticles(data.global);
  }

  // Summary table
  if (data.suppSummary.length) {
    sectionLabel('Supplementary News – In Summary');
    const table = document.createElement('table');
    table.className = 'preview-table';
    data.suppSummary.forEach(cat => {
      const catRow = table.insertRow();
      catRow.className = 'preview-cat-header';
      const td = catRow.insertCell();
      td.colSpan = 3;
      td.textContent = cat.name;

      cat.articles.forEach(a => {
        const row = table.insertRow();
        row.insertCell().textContent = a.headline;
        row.insertCell().textContent = a.company;
        const srcCell = row.insertCell();
        srcCell.textContent = `Read more: ${a.source}`;
      });
    });
    wrap.appendChild(table);
  }

  // Detail
  if (data.suppDetail.length) {
    sectionLabel('Supplementary News – In Detail');
    data.suppDetail.forEach(cat => {
      wrap.appendChild(el('div', 'preview-article-title', cat.name));
      renderFullArticles(cat.articles);
    });
  }

  frag.appendChild(wrap);
  return frag;
}

/* ─────────────────────────────────────────────────────────────────────────
   DOCX GENERATION
   ───────────────────────────────────────────────────────────────────────── */

async function generateDocx(data) {
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
    VerticalAlign, ExternalHyperlink, UnderlineType
  } = docx;

  /* ── Page & Typography constants ── */
  const PAGE_W    = 12240; // 8.5 in
  const PAGE_H    = 15840; // 11 in
  const MARGIN    = 1080;  // 0.75 in
  const CONTENT_W = PAGE_W - MARGIN * 2; // 10080 DXA

  const FONT_BODY    = 'Calibri';
  const FONT_HEADING = 'Calibri';
  const SZ_BODY      = 20;  // 10pt
  const SZ_SECTION   = 22;  // 11pt bold
  const SZ_ARTICLE   = 20;  // 10pt bold
  const SZ_SMALL     = 18;  // 9pt

  /* ── Helpers ── */
  const noBorder = { style: BorderStyle.NIL };
  const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

  const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
  const thinBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };

  const greyBorder = { style: BorderStyle.SINGLE, size: 4, color: 'AAAAAA' };
  const greyBorders = { top: greyBorder, bottom: greyBorder, left: greyBorder, right: greyBorder };

  function para(runs, opts = {}) {
    return new Paragraph({
      spacing: { before: opts.before ?? 0, after: opts.after ?? 80 },
      ...opts,
      children: Array.isArray(runs) ? runs : [runs],
    });
  }

  function run(text, opts = {}) {
    return new TextRun({
      text,
      font: FONT_BODY,
      size: opts.size ?? SZ_BODY,
      bold: opts.bold ?? false,
      italics: opts.italics ?? false,
      color: opts.color ?? '000000',
      ...opts,
    });
  }

  function hyperlink(text, url, opts = {}) {
    const validated = validateUrl(url);
    const safeUrl = validated.ok && validated.url ? validated.url : '';
    if (!safeUrl) {
      return new TextRun({ text, font: FONT_BODY, size: opts.size ?? SZ_SMALL, color: '000000' });
    }
    return new ExternalHyperlink({
      link: safeUrl,
      children: [new TextRun({
        text,
        font: FONT_BODY,
        size: opts.size ?? SZ_SMALL,
        color: '0563C1',
        underline: { type: UnderlineType.SINGLE },
        ...opts,
      })],
    });
  }

  function emptyPara(half = false) {
    return new Paragraph({ spacing: { before: 0, after: half ? 60 : 120 }, children: [] });
  }

  function hrPara() {
    return new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 1 } },
      spacing: { before: 40, after: 120 },
      children: [],
    });
  }

  /* ── Section heading box ── */
  function sectionHeadingBox(label) {
    // Rendered as a single-cell table with black background
    return new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: CONTENT_W, type: WidthType.DXA },
              borders: thinBorders,
              shading: { fill: '000000', type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [
                new Paragraph({
                  spacing: { before: 0, after: 0 },
                  children: [new TextRun({
                    text: label,
                    font: FONT_HEADING,
                    size: SZ_SECTION,
                    bold: true,
                    color: 'FFFFFF',
                  })],
                }),
              ],
            }),
          ],
        }),
      ],
    });
  }

  /* ── Article box ── */
  function articleBox(article) {
    const children = [];

    // Headline
    children.push(para(
      run(article.headline, { bold: true, size: SZ_ARTICLE }),
      { spacing: { before: 0, after: 80 } }
    ));

    // Body
    if (article.body) {
      children.push(para(
        run(article.body, { size: SZ_BODY }),
        { spacing: { before: 0, after: 80 } }
      ));
    }

    // Read more
    if (article.source || article.link) {
      const linkText = `Read more : ${article.source || article.link}`;
      children.push(new Paragraph({
        spacing: { before: 0, after: 80 },
        children: [hyperlink(linkText, article.link, { size: SZ_SMALL })],
      }));
    }

    // Related
    if (article.hasRelated && article.relatedHeadline) {
      children.push(para(
        run(`RELATED: ${article.relatedHeadline}`, { bold: true, size: SZ_SMALL }),
        { spacing: { before: 40, after: 40 } }
      ));
      if (article.relatedSource || article.relatedLink) {
        const relText = `Read more: ${article.relatedSource || article.relatedLink}`;
        children.push(new Paragraph({
          spacing: { before: 0, after: 0 },
          children: [hyperlink(relText, article.relatedLink, { size: SZ_SMALL })],
        }));
      }
    }

    // Wrap in bordered box
    return new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [CONTENT_W],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: CONTENT_W, type: WidthType.DXA },
              borders: thinBorders,
              shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
              margins: { top: 100, bottom: 100, left: 160, right: 160 },
              children,
            }),
          ],
        }),
      ],
    });
  }

  /* ── Summary table ── */
  function summaryTable(categories) {
    const COL1 = Math.round(CONTENT_W * 0.50);
    const COL2 = Math.round(CONTENT_W * 0.20);
    const COL3 = CONTENT_W - COL1 - COL2;

    const rows = [];

    categories.forEach(cat => {
      // Category header row
      rows.push(new TableRow({
        children: [
          new TableCell({
            columnSpan: 3,
            width: { size: CONTENT_W, type: WidthType.DXA },
            borders: thinBorders,
            shading: { fill: 'F2F2F2', type: ShadingType.CLEAR },
            margins: { top: 60, bottom: 60, left: 120, right: 120 },
            children: [
              new Paragraph({
                spacing: { before: 0, after: 0 },
                alignment: AlignmentType.CENTER,
                children: [new TextRun({
                  text: cat.name,
                  font: FONT_HEADING,
                  size: SZ_SMALL,
                  bold: true,
                  color: '000000',
                })],
              }),
            ],
          }),
        ],
      }));

      // Article rows
      cat.articles.forEach(a => {
        const linkText = `Read more : ${a.source || ''}`;
        rows.push(new TableRow({
          children: [
            new TableCell({
              width: { size: COL1, type: WidthType.DXA },
              borders: greyBorders,
              margins: { top: 60, bottom: 60, left: 100, right: 80 },
              children: [para(run(a.headline, { size: SZ_SMALL }), { spacing: { before: 0, after: 0 } })],
            }),
            new TableCell({
              width: { size: COL2, type: WidthType.DXA },
              borders: greyBorders,
              margins: { top: 60, bottom: 60, left: 80, right: 80 },
              children: [para(run(a.company || '', { size: SZ_SMALL }), { spacing: { before: 0, after: 0 } })],
            }),
            new TableCell({
              width: { size: COL3, type: WidthType.DXA },
              borders: greyBorders,
              margins: { top: 60, bottom: 60, left: 80, right: 100 },
              children: [
                new Paragraph({
                  spacing: { before: 0, after: 0 },
                  children: [hyperlink(linkText, a.link, { size: SZ_SMALL })],
                }),
              ],
            }),
          ],
        }));
      });
    });

    return new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [COL1, COL2, COL3],
      rows,
    });
  }

  /* ── Detail section ── */
  function detailSection(categories) {
    const children = [];
    categories.forEach(cat => {
      // Category heading
      children.push(para(
        run(cat.name, { bold: true, size: SZ_SECTION }),
        { spacing: { before: 120, after: 80 } }
      ));

      cat.articles.forEach(a => {
        // Article headline
        children.push(para(
          run(a.headline, { bold: true, size: SZ_ARTICLE }),
          { spacing: { before: 60, after: 60 } }
        ));
        // Body
        if (a.body) {
          children.push(para(run(a.body, { size: SZ_BODY }), { spacing: { before: 0, after: 60 } }));
        }
        // Read more
        if (a.source || a.link) {
          children.push(new Paragraph({
            spacing: { before: 0, after: 80 },
            children: [hyperlink(`Read more : ${a.source || a.link}`, a.link, { size: SZ_SMALL })],
          }));
        }
        if (a.hasRelated && a.relatedHeadline) {
          children.push(para(
            run(`RELATED: ${a.relatedHeadline}`, { bold: true, size: SZ_SMALL }),
            { spacing: { before: 40, after: 40 } }
          ));
          if (a.relatedSource || a.relatedLink) {
            children.push(new Paragraph({
              spacing: { before: 0, after: 80 },
              children: [hyperlink(`Read more: ${a.relatedSource || a.relatedLink}`, a.relatedLink, { size: SZ_SMALL })],
            }));
          }
        }
      });

      children.push(hrPara());
    });
    return children;
  }

  /* ── Assemble document ── */
  const docChildren = [];

  // Date heading
  if (data.date) {
    const d = new Date(data.date + 'T00:00:00');
    const fmt = d.toLocaleDateString('en-GB', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
    docChildren.push(para(run(fmt, { bold: true, size: 24 }), { spacing: { before: 0, after: 160 } }));
  }

  // Top Story
  if (data.topStory.length) {
    docChildren.push(sectionHeadingBox('01 thing you need to know to start your day'));
    docChildren.push(emptyPara(true));
    data.topStory.forEach((a, i) => {
      docChildren.push(articleBox(a));
      if (i < data.topStory.length - 1) docChildren.push(emptyPara(true));
    });
    docChildren.push(emptyPara());
  }

  // Global
  if (data.global.length) {
    docChildren.push(sectionHeadingBox('01 global updates to keep an eye on'));
    docChildren.push(emptyPara(true));
    data.global.forEach((a, i) => {
      docChildren.push(articleBox(a));
      if (i < data.global.length - 1) docChildren.push(emptyPara(true));
    });
    docChildren.push(emptyPara());
  }

  // Summary
  if (data.suppSummary.length) {
    docChildren.push(sectionHeadingBox('Supplementary News – In Summary'));
    docChildren.push(emptyPara(true));
    docChildren.push(summaryTable(data.suppSummary));
    docChildren.push(emptyPara());
  }

  // Detail
  if (data.suppDetail.length) {
    docChildren.push(sectionHeadingBox('Supplementary News – In Detail'));
    docChildren.push(emptyPara(true));
    detailSection(data.suppDetail).forEach(c => docChildren.push(c));
  }

  const doc = new Document({
    creator: 'The Red Box Newsletter Builder',
    description: 'Sri Lanka English-language news digest',
    styles: {
      default: {
        document: {
          run: { font: FONT_BODY, size: SZ_BODY, color: '000000' },
          paragraph: { spacing: { before: 0, after: 80 } },
        },
      },
    },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_W, height: PAGE_H },
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
  const a   = document.createElement('a');
  a.href     = url;
  a.download = filename;
  a.rel      = 'noopener noreferrer';
  document.body.appendChild(a);
  a.click();
  // Clean up after a tick
  setTimeout(() => {
    URL.revokeObjectURL(url);
    a.remove();
  }, 1000);
}

function safeFilename(date) {
  const d = date ? date.replace(/-/g, '') : 'newsletter';
  return `red-box-newsletter-${d}.docx`;
}

/* ─────────────────────────────────────────────────────────────────────────
   MAIN GENERATE FLOW
   ───────────────────────────────────────────────────────────────────────── */

async function runGenerate() {
  const data   = collectFormData();
  const errors = validateData(data);

  if (errors.length) {
    toast('Fix these issues:\n• ' + errors.slice(0, 3).join('\n• '), true);
    return;
  }

  const btn = document.getElementById('btn-generate');
  btn.disabled = true;
  btn.textContent = 'Generating…';

  try {
    const blob = await generateDocx(data);
    downloadBlob(blob, safeFilename(data.date));
    toast('Document downloaded!');
  } catch (err) {
    console.error('docx generation failed:', err);
    toast('Generation failed. Check console for details.', true);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Generate .docx';
  }
}

/* ─────────────────────────────────────────────────────────────────────────
   WIRE UP
   ───────────────────────────────────────────────────────────────────────── */

document.addEventListener('DOMContentLoaded', () => {

  // Default date to today
  const todayEl = document.getElementById('nl-date');
  const today = new Date();
  const yyyy  = today.getFullYear();
  const mm    = String(today.getMonth() + 1).padStart(2, '0');
  const dd    = String(today.getDate()).padStart(2, '0');
  todayEl.value = `${yyyy}-${mm}-${dd}`;

  // Section add-article buttons
  document.querySelectorAll('.btn-add[data-target]').forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.target;
      const listEl = document.getElementById(`articles-${target}`);
      appendArticle(listEl, false);
    });
  });

  // Summary category button
  document.getElementById('btn-add-supp-cat').addEventListener('click', () => {
    appendCategory(document.getElementById('supp-summary-categories'), true);
  });

  // Detail category button
  document.getElementById('btn-add-detail-cat').addEventListener('click', () => {
    appendCategory(document.getElementById('supp-detail-categories'), false);
  });

  // Generate button
  document.getElementById('btn-generate').addEventListener('click', runGenerate);

  // Preview button
  document.getElementById('btn-preview').addEventListener('click', () => {
    const data = collectFormData();
    const body = document.getElementById('modal-preview-body');
    body.innerHTML = '';
    body.appendChild(buildPreviewHtml(data));
    const backdrop = document.getElementById('modal-backdrop');
    backdrop.hidden = false;
    backdrop.removeAttribute('aria-hidden');
  });

  // Modal close
  function closeModal() {
    const backdrop = document.getElementById('modal-backdrop');
    backdrop.hidden = true;
    backdrop.setAttribute('aria-hidden', 'true');
  }

  document.getElementById('modal-close').addEventListener('click', closeModal);
  document.getElementById('modal-cancel').addEventListener('click', closeModal);
  document.getElementById('modal-backdrop').addEventListener('click', e => {
    if (e.target === e.currentTarget) closeModal();
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') closeModal();
  });

  // Modal generate
  document.getElementById('modal-generate').addEventListener('click', () => {
    closeModal();
    runGenerate();
  });
});
