/**
 * builder.js — The Red Box Newsletter Builder
 *
 * Generates .docx files matching the exact styling of the reference document:
 *   - Font: Verdana throughout
 *   - Section headings: Bold, color #0B5394, 13.5pt (sz=27)
 *   - Article boxes: Single-cell table, 9350 DXA wide, black borders sz=4
 *   - Body text: Verdana 10pt (sz=20), justified, yellow highlight on key paras
 *   - Hyperlinks: Blue #0000FF, underlined
 *   - Summary table: 3-col (3403/2601/3310 DXA), category headers span full width
 *   - Page: US Letter, 1" margins all sides (1440 DXA)
 *
 * Security:
 *   - All user input sanitised (textContent, not innerHTML)
 *   - HTTPS-only URL validation with private-IP/SSRF guard
 *   - Strict length caps on all fields
 *   - docx generated 100% client-side; no data leaves the browser
 *   - No eval(), no dynamic script injection
 */

'use strict';

/* ── CONSTANTS ── */
const MAX_BODY_WORDS  = 220;
const MAX_TITLE_LEN   = 300;
const MAX_SOURCE_LEN  = 100;
const MAX_URL_LEN     = 2048;
const MAX_CAT_LEN     = 80;

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
  if (PRIVATE_IP_RE.test(parsed.hostname)) return { ok: false, reason: 'Private/reserved IP addresses not allowed' };
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
  // Preserve paragraph breaks (\n\n) while counting total words across all paras
  const normalised = str.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  // Split into paragraphs first
  const paragraphs = normalised.split(/\n{2,}/);
  const result = [];
  let remaining = limit;
  for (const para of paragraphs) {
    if (remaining <= 0) break;
    const words = para.trim().split(/\s+/).filter(Boolean);
    if (words.length === 0) continue;
    if (words.length <= remaining) {
      result.push(words.join(' '));
      remaining -= words.length;
    } else {
      result.push(words.slice(0, remaining).join(' ') + '….');
      remaining = 0;
    }
  }
  const joined = result.join('\n\n');
  // Only add ellipsis suffix if we actually cut something
  const totalWords = paragraphs.flatMap(p => p.trim().split(/\s+/).filter(Boolean)).length;
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
   ARTICLE ENTRY FACTORY
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
   CATEGORY BLOCK FACTORY
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
      const bodyRaw  = entry.querySelector('.field-body').value;
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
      if (!r.ok) errors.push(`Article ${i+1} link: ${r.reason}`);
      if (a.hasRelated) {
        const r2 = validateUrl(a.relatedLink);
        if (!r2.ok) errors.push(`Article ${i+1} related link: ${r2.reason}`);
      }
    });
  }
  checkUrls(data.topStory);
  checkUrls(data.global);
  data.suppSummary.forEach(cat => cat.articles.forEach((a, i) => {
    const r = validateUrl(a.link);
    if (!r.ok) errors.push(`Summary "${cat.name}" article ${i+1}: ${r.reason}`);
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

  function sectionLabel(text) { wrap.appendChild(el('div', 'preview-section-label', text)); }

  function renderFullArticles(articles) {
    articles.forEach(a => {
      wrap.appendChild(el('div', 'preview-article-box', ''));
      const box = wrap.lastChild;
      box.appendChild(el('div', 'preview-article-title', a.headline));
      if (a.body) box.appendChild(el('div', 'preview-article-body', a.body));
      if (a.source || a.link) box.appendChild(el('div', 'preview-read-more', `Read More: ${a.source}`));
      if (a.hasRelated && a.relatedHeadline) {
        box.appendChild(el('div', 'preview-related', `Related: ${a.relatedHeadline}`));
        if (a.relatedSource) box.appendChild(el('div', 'preview-read-more', `Read More: ${a.relatedSource}`));
      }
    });
  }

  if (data.date) {
    const d = new Date(data.date + 'T00:00:00');
    wrap.appendChild(el('div', 'preview-date', d.toLocaleDateString('en-GB', { weekday:'long', year:'numeric', month:'long', day:'numeric' })));
  }
  if (data.topStory.length) { sectionLabel('01 thing you need to know to start your day'); renderFullArticles(data.topStory); }
  if (data.global.length)   { sectionLabel('01 global updates to keep an eye on'); renderFullArticles(data.global); }

  if (data.suppSummary.length) {
    sectionLabel('Supplementary News – In Summary');
    const table = document.createElement('table');
    table.className = 'preview-table';
    data.suppSummary.forEach(cat => {
      const hr = table.insertRow(); hr.className = 'preview-cat-header';
      const td = hr.insertCell(); td.colSpan = 3; td.textContent = cat.name;
      cat.articles.forEach(a => {
        const row = table.insertRow();
        row.insertCell().textContent = a.headline;
        row.insertCell().textContent = a.company;
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
   DOCX GENERATION — matches reference document exactly
   ───────────────────────────────────────────────────────────────────────── */
async function generateDocx(data) {
  const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, BorderStyle, WidthType, ShadingType,
    ExternalHyperlink, UnderlineType
  } = docx;

  /* ── Reference-matched constants ── */
  const PAGE_W    = 12240;  // US Letter
  const PAGE_H    = 15840;
  const MARGIN    = 1440;   // 1 inch
  const TBL_W     = 9350;   // article box table width (from reference)
  const TBL_W_SUM = 9314;   // summary table total width (from reference)
  const COL1      = 3403;   // summary col 1 (headline)
  const COL2      = 2601;   // summary col 2 (company/tag)
  const COL3      = 3310;   // summary col 3 (read more)
  const COL_2C_1  = 6004;   // 2-col summary: headline
  const COL_2C_2  = 3310;   // 2-col summary: read more

  const FONT       = 'Verdana';
  const SZ_BODY    = 20;    // 10pt
  const SZ_LABEL   = 27;    // 13.5pt — section headings ("01 thing...")
  const SZ_SMALL   = 19;    // 9.5pt — read more links
  const COLOR_HEAD = '0B5394'; // dark blue for section headings
  const COLOR_LINK = '0000FF'; // hyperlink blue

  /* ── Border presets ── */
  const blackBorder  = { style: BorderStyle.SINGLE, size: 4, color: '000000', space: 0 };
  const blackBorders = { top: blackBorder, bottom: blackBorder, left: blackBorder, right: blackBorder, insideH: blackBorder, insideV: blackBorder };
  const nilBorder    = { style: BorderStyle.NIL };
  const nilBorders   = { top: nilBorder, bottom: nilBorder, left: nilBorder, right: nilBorder };

  /* ── Helpers ── */
  function vRun(text, opts = {}) {
    return new TextRun({
      text,
      font: FONT,
      size: opts.size ?? SZ_BODY,
      bold: opts.bold ?? false,
      italics: opts.italics ?? false,
      color: opts.color ?? '000000',
      highlight: opts.highlight ?? undefined,
      underline: opts.underline ?? undefined,
    });
  }

  function vPara(children, opts = {}) {
    return new Paragraph({
      alignment: opts.alignment ?? AlignmentType.BOTH,
      spacing: { before: opts.before ?? 0, after: opts.after ?? 80 },
      children: Array.isArray(children) ? children : [children],
    });
  }

  function emptyPara(after = 120) {
    return new Paragraph({ spacing: { before: 0, after }, children: [] });
  }

  function hyperlink(text, url) {
    const v = validateUrl(url);
    if (!v.ok || !v.url) {
      return vRun(text, { size: SZ_SMALL, color: COLOR_LINK, underline: { type: UnderlineType.SINGLE } });
    }
    return new ExternalHyperlink({
      link: v.url,
      children: [new TextRun({
        text,
        font: FONT,
        size: SZ_SMALL,
        color: COLOR_LINK,
        underline: { type: UnderlineType.SINGLE },
      })],
    });
  }

  /* ── Section heading paragraph ("01 thing...", "01 global...") ──
     Blue bold 13.5pt, NOT in a table — matches reference exactly */
  function sectionHeadingPara(text) {
    return vPara(
      vRun(text, { bold: true, size: SZ_LABEL, color: COLOR_HEAD }),
      { before: 0, after: 80 }
    );
  }

  /* ── Article box — single-cell table with black border ── */
  function articleBox(article) {
    const cellChildren = [];

    // Headline — bold
    cellChildren.push(vPara(
      vRun(article.headline, { bold: true }),
      { after: 80 }
    ));

    // Empty line after headline
    cellChildren.push(emptyPara(0));

    // Body paragraphs: split on blank lines to preserve paragraph structure.
    // Pasted text from Word/browser uses \n\n between paragraphs; a single \n
    // within a paragraph is a soft wrap and becomes a space.
    if (article.body) {
      const normalised = article.body.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      const paras = normalised.split(/\n{2,}/).map(p => p.replace(/\n/g, ' ').trim()).filter(p => p);
      paras.forEach((p, i) => {
        const highlight = i < 2 ? 'yellow' : undefined;
        cellChildren.push(vPara(
          vRun(p, { highlight }),
          { after: 80 }
        ));
        if (i < paras.length - 1) cellChildren.push(emptyPara(0));
      });
    }

    // Read More link
    if (article.source || article.link) {
      cellChildren.push(emptyPara(0));
      cellChildren.push(new Paragraph({
        alignment: AlignmentType.BOTH,
        spacing: { before: 0, after: 80 },
        children: [hyperlink(`Read More: ${article.source || article.link}`, article.link)],
      }));
    }

    // RELATED articles
    if (article.hasRelated && article.relatedHeadline) {
      cellChildren.push(vPara(
        vRun(`Related: ${article.relatedHeadline}`, { bold: true }),
        { before: 80, after: 40 }
      ));
      if (article.relatedSource || article.relatedLink) {
        cellChildren.push(new Paragraph({
          alignment: AlignmentType.BOTH,
          spacing: { before: 0, after: 0 },
          children: [hyperlink(`Read More: ${article.relatedSource || article.relatedLink}`, article.relatedLink)],
        }));
      }
    }

    return new Table({
      width: { size: TBL_W, type: WidthType.DXA },
      columnWidths: [TBL_W],
      borders: blackBorders,
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: TBL_W, type: WidthType.DXA },
              borders: blackBorders,
              margins: { top: 0, bottom: 0, left: 100, right: 100 },
              children: cellChildren,
            }),
          ],
        }),
      ],
    });
  }

  /* ── Summary table ──
     Matches reference exactly:
     - Category header: full-width (9314), bold, gridSpan=3
     - 2-col rows: headline (6004) + read more (3310) — for cats without company tag
     - 3-col rows: headline (3403) + company (2601) + read more (3310) — for cats with tags
  */
  function summaryTable(categories) {
    const rows = [];

    categories.forEach(cat => {
      // Determine if this category uses 3-col (has any article with a company tag)
      const hasCompany = cat.articles.some(a => a.company && a.company.trim());

      // Category header row — full width, bold
      rows.push(new TableRow({
        children: [
          new TableCell({
            width: { size: TBL_W_SUM, type: WidthType.DXA },
            columnSpan: 3,
            borders: blackBorders,
            margins: { top: 0, bottom: 0, left: 100, right: 100 },
            children: [
              new Paragraph({
                alignment: AlignmentType.BOTH,
                spacing: { before: 0, after: 0 },
                children: [vRun(cat.name, { bold: true })],
              }),
            ],
          }),
        ],
      }));

      // Article rows
      cat.articles.forEach(a => {
        if (hasCompany) {
          // 3-column row
          rows.push(new TableRow({
            children: [
              new TableCell({
                width: { size: COL1, type: WidthType.DXA },
                borders: blackBorders,
                margins: { top: 0, bottom: 0, left: 100, right: 100 },
                children: [vPara(vRun(a.headline), { after: 0 })],
              }),
              new TableCell({
                width: { size: COL2, type: WidthType.DXA },
                borders: blackBorders,
                margins: { top: 0, bottom: 0, left: 100, right: 100 },
                children: [vPara(vRun(a.company || ''), { after: 0 })],
              }),
              new TableCell({
                width: { size: COL3, type: WidthType.DXA },
                borders: blackBorders,
                margins: { top: 0, bottom: 0, left: 100, right: 100 },
                children: [new Paragraph({
                  alignment: AlignmentType.BOTH,
                  spacing: { before: 0, after: 0 },
                  children: [hyperlink(`Read More: ${a.source || ''}`, a.link)],
                })],
              }),
            ],
          }));
        } else {
          // 2-column row (no company tag)
          rows.push(new TableRow({
            children: [
              new TableCell({
                width: { size: COL_2C_1, type: WidthType.DXA },
                columnSpan: 2,
                borders: blackBorders,
                margins: { top: 0, bottom: 0, left: 100, right: 100 },
                children: [vPara(vRun(a.headline), { after: 0 })],
              }),
              new TableCell({
                width: { size: COL_2C_2, type: WidthType.DXA },
                borders: blackBorders,
                margins: { top: 0, bottom: 0, left: 100, right: 100 },
                children: [new Paragraph({
                  alignment: AlignmentType.BOTH,
                  spacing: { before: 0, after: 0 },
                  children: [hyperlink(`Read More: ${a.source || ''}`, a.link)],
                })],
              }),
            ],
          }));
        }
      });
    });

    return new Table({
      width: { size: TBL_W_SUM, type: WidthType.DXA },
      columnWidths: [COL1, COL2, COL3],
      borders: blackBorders,
      rows,
    });
  }

  /* ── Detail section — plain paragraphs, bold blue category names ── */
  function detailSection(categories) {
    const children = [];
    categories.forEach(cat => {
      // Category name — bold blue, same style as top-level section headings
      children.push(vPara(
        vRun(cat.name, { bold: true, size: SZ_LABEL, color: COLOR_HEAD }),
        { before: 160, after: 80 }
      ));

      cat.articles.forEach(a => {
        // Article headline — bold
        children.push(vPara(vRun(a.headline, { bold: true }), { before: 80, after: 80 }));

        // Body: split on blank lines to preserve paragraph spacing
        if (a.body) {
          const normalised = a.body.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
          const paras = normalised.split(/\n{2,}/).map(p => p.replace(/\n/g, ' ').trim()).filter(p => p);
          paras.forEach((p, i) => {
            children.push(vPara(vRun(p), { after: 80 }));
            if (i < paras.length - 1) children.push(emptyPara(0));
          });
          children.push(emptyPara(0));
        }

        // Read More
        if (a.source || a.link) {
          children.push(new Paragraph({
            alignment: AlignmentType.BOTH,
            spacing: { before: 0, after: 80 },
            children: [hyperlink(`Read More: ${a.source || a.link}`, a.link)],
          }));
        }

        // Related
        if (a.hasRelated && a.relatedHeadline) {
          children.push(vPara(vRun(`Related: ${a.relatedHeadline}`, { bold: true }), { before: 40, after: 40 }));
          if (a.relatedSource || a.relatedLink) {
            children.push(new Paragraph({
              alignment: AlignmentType.BOTH,
              spacing: { before: 0, after: 80 },
              children: [hyperlink(`Read More: ${a.relatedSource || a.relatedLink}`, a.relatedLink)],
            }));
          }
        }

        children.push(emptyPara(0));
      });
    });
    return children;
  }

  /* ── Assemble document ── */
  const docChildren = [];

  // Date
  if (data.date) {
    const d = new Date(data.date + 'T00:00:00');
    const fmt = d.toLocaleDateString('en-GB', { weekday:'long', year:'numeric', month:'long', day:'numeric' });
    docChildren.push(vPara(vRun(fmt, { bold: true }), { after: 160 }));
  }

  // Top story section
  if (data.topStory.length) {
    docChildren.push(sectionHeadingPara('01 thing you need to know to start your day'));
    docChildren.push(emptyPara(0));
    data.topStory.forEach((a, i) => {
      docChildren.push(articleBox(a));
      if (i < data.topStory.length - 1) docChildren.push(emptyPara(80));
    });
    docChildren.push(emptyPara(160));
  }

  // Global updates section
  if (data.global.length) {
    docChildren.push(sectionHeadingPara('01 global updates to keep an eye on'));
    docChildren.push(emptyPara(0));
    data.global.forEach((a, i) => {
      docChildren.push(articleBox(a));
      if (i < data.global.length - 1) docChildren.push(emptyPara(80));
    });
    docChildren.push(emptyPara(160));
  }

  // Supplementary – In Detail (matches reference order: Detail comes before Summary)
  if (data.suppDetail.length) {
    docChildren.push(vPara(
      [vRun('Supplementary News', { bold: true }), vRun(' – In Detail', { bold: true })],
      { after: 80 }
    ));
    detailSection(data.suppDetail).forEach(c => docChildren.push(c));
    docChildren.push(emptyPara(160));
  }

  // Supplementary – In Summary
  if (data.suppSummary.length) {
    docChildren.push(vPara(
      [vRun('Supplementary News', { bold: true }), vRun(' – In Summary', { bold: true })],
      { after: 80 }
    ));
    docChildren.push(summaryTable(data.suppSummary));
    docChildren.push(emptyPara(0));
  }

  const doc = new Document({
    creator: 'The Red Box Newsletter Builder',
    description: 'Sri Lanka English-language news digest',
    styles: {
      default: {
        document: {
          run: { font: FONT, size: SZ_BODY, color: '000000' },
          paragraph: { spacing: { before: 0, after: 80 }, alignment: AlignmentType.BOTH },
        },
      },
      characterStyles: [{
        id: 'Hyperlink',
        name: 'Hyperlink',
        run: { color: COLOR_LINK, underline: { type: UnderlineType.SINGLE } },
      }],
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
  const data = collectFormData();
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
    `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;

  // Add-article buttons
  document.querySelectorAll('.btn-add[data-target]').forEach(btn => {
    btn.addEventListener('click', () => {
      appendArticle(document.getElementById(`articles-${btn.dataset.target}`), false);
    });
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
