# WikiZotRef — Wikipedia references to Zotero citations

A single self-contained HTML page that converts a Wikipedia article's wikitext into a Word document whose `<ref>` citations are live Zotero fields, ready for Zotero's "Refresh" to format. No server, no build step, no install: everything runs in the browser.

## For users

### What it does

You write or edit a Wikipedia (or Wikiversity, Wikibooks, etc.) article and want to turn it into a Word document with proper, editable Zotero citations instead of Wikipedia-style footnotes. WikiZotRef:

1. Reads the article's wikitext and finds every `<ref>…</ref>`.
2. Matches each one to an item in your Zotero library, using whichever identifier it can find: a Wikidata QID (via `{{Cite Q}}` or an item's `Extra` field), a DOI, or a URL.
3. Lets you confirm or correct each match in a table before anything is written.
4. Converts the rest of the wikitext — headings, bold/italic, lists, tables, links — into Word formatting.
5. Downloads a `.docx` file where each matched reference is a real Zotero citation field (the same kind Zotero's own "Add Citation" creates), not plain text.

Open the file in Word, go to the **Zotero** tab, click **Refresh**, and pick a citation style — Zotero formats every citation and can build the bibliography for you.

### Step by step

1. **Add the article's references to your Zotero library first.** WikiZotRef only *matches* references to items that already exist in your library — it never creates or imports new ones. If you haven't done this yet, see [Wikipedia: Citing sources with Zotero](https://en.wikipedia.org/wiki/Wikipedia:Citing_sources_with_Zotero#Adding_the_references_of_a_Wikipedia_article_to_your_Zotero_library) and [Zotero: Zotero and Wikipedia](https://www.zotero.org/support/kb/zotero_and_wikipedia).
2. **Export your library.** In Zotero: right-click your library or collection → Export → format "CSL JSON". Load that file into WikiZotRef.
3. **Paste the wikitext.** Use the article's "Edit source" view so the `<ref>` tags are visible (not the rendered page). Optionally set the base URL used to turn internal `[[links]]` into real hyperlinks — it defaults to English Wikipedia.
4. **Check the matches.** Each reference shows its detected QID, DOI or URL and the Zotero item WikiZotRef thinks it matches. Fix any wrong guesses with the dropdown, or leave unmatched ones as skipped.
5. **Download and refresh.** Open the `.docx` in Word, click Refresh in the Zotero tab, choose a style.

### What's supported

- **Matching:** Wikidata QID (from `{{Cite Q|Q…}}` or a `Q\d+` anywhere in the ref, matched against a library item's `Extra` field written as `QID: Q…`), DOI, and — as a fallback when neither is found — a URL matched against the item's `URL` field.
- **Reused references:** `<ref name="x">…</ref>` and later `<ref name="x" />` resolve to the same citation.
- **Grouped citations:** references cited back-to-back (separated only by whitespace, commas or semicolons, e.g. `<ref>…</ref><ref>…</ref>` or with a comma between them) are merged into a single Zotero field with several sources, instead of one field per reference.
- **Page numbers:** `|page=` or `|pages=` on a `{{Cite Q}}` becomes the citation's locator.
- **Unmatched references:** kept as visible red text in the Word file (`[unmatched reference: …]`) so you can spot and fix them by hand with Zotero's own "Add/Edit Citation".
- **Formatting:** headings, `'''bold'''`/`''italic''`, bulleted and numbered lists, simple tables.
- **Links:** `[[Internal links]]` (if you set a base URL), `[https://example.com External links]`, bare URLs, and `{{Wikidata entity link|Q…}}` templates — the last is turned into a link to wikidata.org, labelled with the item's name (fetched live from Wikidata; falls back to showing the bare QID if that lookup fails).
- **Reference lists:** `{{Reflist}}` and `<references />` are removed, since Zotero builds the bibliography itself.

### Limitations

- Only the **first** QID, DOI or URL in a `<ref>` is used — multiple sources packed into one `<ref>` tag aren't split out.
- Grouping only merges references that are *directly* adjacent in the source; a short aside or an unmatched reference between two citations breaks the chain.
- Wikipedia-specific templates other than `{{Cite Q}}` and `{{Wikidata entity link}}` are stripped, not expanded — infoboxes, `{{convert}}`, `{{lang}}`, etc. will disappear and need manual cleanup.
- Sentence-end detection (moving a citation before the final `.`/`!`/`?`) is a simple pattern match; it doesn't know about abbreviations or closing quotation marks.
- Needs an internet connection the first time it looks up a Wikidata label; everything else works fully offline.

## For developers

### Why no dependencies

An earlier version used [Pandoc compiled to WebAssembly](https://github.com/pandoc/pandoc-wasm) for the wikitext→docx conversion, loaded from a CDN. In practice this turned out not to be "open and it just works": the package needs bundler-level `.wasm` asset handling that a CDN's `+esm` endpoint can't provide, and even once bundled locally, ES module imports are blocked from a `file://` page, which is how most people would actually open the tool. Rather than asking users to run a build step and a local web server just to use a single HTML file, the tool now does the wikitext→docx conversion itself, by hand, trading full MediaWiki fidelity for something that works by double-clicking the file.

### File layout

Everything lives in one file, `wikitext-to-zotero-docx.html`:

- Inline `<style>` — no external stylesheet, light/dark mode via `prefers-color-scheme`.
- A short HTML body with four numbered `<section>`s mirroring the user-facing steps, plus a closing "After downloading" note.
- One `<script type="module">` containing all the logic, in this order: matching helpers → UI event handlers for loading the library and scanning refs → the citation field builder → the docx builder → the "download" handler that ties it all together.

### Data flow

```
CSL JSON file ──► lib (array of CSL-JSON items)
wikitext       ──► scan() on 'Find references' click
                     │
                     ▼
                 refs[] { start, end, body, qid, doi, url, page, cands, sel, dup? }
                     │  (user edits r.sel via the review table's <select>)
                     ▼
            'Download Word file' click
                     │
                     ▼
     group adjacent matched refs ──► groups[] { tok, items: [{it, page}] }
     unmatched refs              ──► missText{ tok: displayString }
     {{Wikidata entity link}}    ──► fetch labels ──► wdLabels{ QID: label }
                     │
                     ▼
     out = wikitext with <ref> spans replaced by ZOTREF####X / ZOTMISS####X tokens,
           citation fields repositioned before sentence-final punctuation,
           a leading space inserted before every ZOTREF token
                     │
                     ▼
              buildDocx(out, fields, missText, wikiBase, wdLabels)
                     │
                     ▼
                article.docx (downloaded)
```

### Matching (`scan()`)

A single regex finds both self-closing and paired `<ref>` tags:

```js
/<ref\b([^>]*?)\/>|<ref\b([^>]*)>([\s\S]*?)<\/ref>/gi
```

For each match it extracts, in order of preference: a Wikidata QID (`\bQ\d+\b`), a DOI (`10\.\d{4,9}/…`), and a URL (`|url=` parameter or a bare `https?://` link). It looks up library items whose `Extra` field (CSL JSON's `note`) contains a line matching `/^\s*qid\s*:\s*(Q\d+)\s*$/im`, or whose `DOI`/`URL` field matches after normalisation (`normDoi`, `normUrl` strip scheme/case/trailing slashes so `https://doi.org/X` and `X` compare equal). Named refs (`<ref name="x" />` reused later) are linked back to the first occurrence of that name via the `dup` field rather than re-matched.

The result, one object per `<ref>` found, is stored in the module-level `refs` array and rendered as an editable table — `r.sel` (index into `r.cands`, or `-1` for "skip") is the only field the UI mutates after the initial scan.

### Grouping and token substitution (`$('go')` handler)

References are walked in source order. A ref is merged into the *immediately preceding* group (not just "the last group created") only if the text between them matches `/^[\s,;]*$/` — this distinction matters: an unmatched reference sitting between two matched ones must break the chain even though it doesn't push anything onto the `groups` array itself, which is why `lastGroup` is tracked and explicitly reset to `null` in the unmatched branch rather than re-derived from `groups[groups.length - 1]`.

Each group becomes one placeholder token (`ZOTREF0000X`, `ZOTREF0001X`, …); each unmatched ref becomes a `ZOTMISS0000X` token mapped to its display string in `missText`. These tokens are plain alphanumeric strings so they survive the docx builder's wikitext-stripping regexes untouched, the same trick used for Pandoc-style "protect this span" filters.

After substitution, two regex passes handle the two cosmetic requests that needed whole-string context rather than per-ref logic:

```js
out = out.replace(/([.!?])(ZOTREF\d{4}X)/g, '$2$1');  // field before sentence-final punctuation
out = out.replace(/(ZOTREF\d{4}X)/g, ' $1');            // space before every field
```

### Citation fields (`fieldXml`, `shortCite`)

A Zotero/Word citation is an OOXML complex field: `begin` → `instrText` (the actual field code) → `separate` → the visible result → `end`. The field code Zotero recognises is `ADDIN ZOTERO_ITEM CSL_CITATION {json}`, where the JSON's `citationItems` array holds one entry per source (each with the item's CSL-JSON under `itemData`, plus an optional `locator`/`label` for a page number) and `properties.formattedCitation` is what displays until the user clicks Refresh in Word. `shortCite()` builds that placeholder as `Author, Year` (`Author & Author2, Year` for two authors, `Author et al., Year` for more), joined with `; ` across a group's items — approximating what most author-date styles will eventually render, though the real text comes from Zotero itself.

### The docx builder (`buildDocx`)

No dependency, built from three pieces:

- **`zip()`** — a minimal, store-only (uncompressed) ZIP writer: local file headers, central directory, end record, with its own CRC-32 table. A `.docx` is just a ZIP of XML parts, and Word doesn't require `deflate` compression, so skipping it keeps the code short.
- **`clean()` / `runs()`** — a line-oriented wikitext→OOXML converter. `clean()` strips templates and converts `[[links]]`, `[url label]`, bare URLs and `{{Wikidata entity link|Q…}}` into placeholder tokens (`LNKREF####X`) *before* the generic `{{…}}`-stripping pass runs, so those templates survive long enough to become hyperlinks; `runs()` then turns a paragraph's text into a sequence of OOXML `<w:r>` runs, recognising bold/italic markers (`'''`, `''`) and all three token families (`ZOTREF`/`ZOTMISS`/`LNKREF`).
- **The line loop** — walks the wikitext line by line, recognising `{|…|}` tables, `== headings ==`, `*`/`#`/`:`/`;` list markers, and blank lines as paragraph breaks; everything else accumulates into the current paragraph. This is intentionally a small subset of MediaWiki syntax, not a full parser — nested templates, `<nowiki>`, and most extension tags are not specifically handled.

Hyperlinks need a package-level relationship, not just inline XML: `relFor()` assigns each link target a fresh `rId`, these are collected into `word/_rels/document.xml.rels`, and the document root declares the `xmlns:r` namespace the `r:id` attribute depends on. A `Hyperlink` character style is added to `styles.xml` purely for appearance (blue, underlined) — Word doesn't require it to follow the link.

### Wikidata label lookup

`{{Wikidata entity link|Q…}}` resolution is split across two places because `buildDocx` is synchronous but fetching a label is not: the `$('go')` handler scans `out` for every `Q\d+` inside such a template, batches them (50 per request — the MediaWiki API's practical limit for `wbgetentities`), and calls `https://www.wikidata.org/w/api.php?action=wbgetentities&…&origin=*` (the `origin=*` parameter is what makes this work cross-origin from a page that isn't on `wikidata.org`, without needing an API key). The resulting `{QID: label}` map is passed into `buildDocx`, where `clean()` consumes it synchronously. The language tried first is guessed from the `wikiBase` input's subdomain (`fr.wikipedia.org` → `fr`), falling back to English, then to the bare QID if even that's missing — a network failure here degrades gracefully rather than aborting the conversion.

### Known rough edges / where to look first if something breaks

- **CSL JSON shape assumptions:** the code assumes `it.note` holds the Extra field's text and `it.id` holds a usable identifier (either a full `http://zotero.org/...` item URI, used directly as the citation's `uris`, or an opaque key used as a fallback `id`). If Zotero changes its export shape, `qidOf`/`doiOf`/`urlOf`/`fieldXml` are the places to adjust.
- **The wikitext parser is regex-based, not a real tokenizer.** Nested or malformed markup (an unclosed `{{`, a table inside a template) can produce visibly wrong output; there's no error recovery beyond "something didn't get stripped".
- **Grouping separator** is defined once as `const SEP = /^[\s,;]*$/` in the `go` handler — change it there if a different separator convention needs to count as "cited together".
