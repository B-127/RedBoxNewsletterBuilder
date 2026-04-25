# The Red Box: Newsletter Builder 📰

A standalone, client-side tool for composing **The Red Box** weekly email digest and exporting it as a formatted `.docx` file.

No backend. No database. No data leaves the browser. The Word document is generated entirely in-page and downloaded directly.

---

## What it does

Fill in the four sections of the newsletter:

| Section | Format |
|---|---|
| 01 thing you need to know to start your day | Full article — headline, body (auto-truncated to ~200 words), Read More link, optional RELATED link |
| 01 global updates to keep an eye on | Same as above |
| Supplementary News – In Summary | Table view — headline, company tag, source link per article, grouped by category |
| Supplementary News – In Detail | Full article per item, grouped by category |

Each section supports a variable number of articles via the **+** button. Categories in the summary and detail sections can also be added dynamically.

Click **Preview** to review before exporting. Click **Generate .docx** to download the formatted Word document.

---

## Project structure

```
red-box-newsletter/
├── index.html              # Single-page form app
├── builder.css             # All styles
├── builder.js              # All logic — form, validation, docx generation
└── .github/workflows/
    └── deploy.yml          # GitHub Pages deployment
```

---

## Setup

### 1. Fork / push to GitHub

```bash
git init
git add .
git commit -m "init: Red Box Newsletter Builder"
git remote add origin https://github.com/YOUR_USERNAME/red-box-newsletter.git
git push -u origin main
```

### 2. Enable GitHub Pages

Go to **Settings → Pages → Source → GitHub Actions**.

### 3. Trigger first deploy

Go to **Actions → Deploy to GitHub Pages → Run workflow**.

The builder will be live at `https://YOUR_USERNAME.github.io/red-box-newsletter/`.

### 4. Local development

No build step required. Just open `index.html` in a browser, or serve it locally:

```bash
python3 -m http.server 8080
# then open http://localhost:8080
```

---

## Security

All security measures are enforced entirely client-side:

- **No data leaves the browser** — `.docx` generated locally via the `docx` library
- **HTTPS-only URLs** — all Read More links validated; `http://` rejected
- **SSRF guard** — private/reserved IP ranges blocked in URL fields (RFC 1918, 6598, loopback, link-local, IPv6 unique-local)
- **Input sanitisation** — non-printable control characters stripped from all fields
- **Length caps** — headlines (300 chars), source names (100 chars), URLs (2048 chars), body text (5000 chars / ~200 words)
- **DOM safety** — all user content inserted via `textContent`, never `innerHTML`
- **Strict CSP** — `Content-Security-Policy` header blocks inline scripts, mixed content, framing, and form submissions
- **Supply-chain safety** — all GitHub Actions pinned to full commit SHAs

---
