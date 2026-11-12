# docx-track

**Edit Word documents with tracked changes — no Word, no GUI, no broken formatting.**

[![PyPI](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
![No Word required](https://img.shields.io/badge/Word-not%20required-informational.svg)

`docxtrack` is a small command-line tool for the job that office automation
usually gets wrong: change a `.docx` and leave the change **reviewable**. Every
edit becomes a proper Word revision (`w:ins` / `w:del`), every untouched word
keeps its own formatting, and accepting or rejecting the revisions gives you
back exactly the new or the old document — byte-for-byte identical apart from the
edits you asked for.

It runs headless. No Microsoft Word, no LibreOffice, no COM automation, no
Office 365 login. Pure Python and `lxml`.

```bash
docxtrack list contract.docx --grep indemnity
docxtrack replace contract.docx "indemnity clause" --find "10 days" --with "5 business days" -o contract_v2.docx
docxtrack verify contract_v2.docx      # paragraphs: 214, ins: 2, del: 2, zip_ok: True
```

Open `contract_v2.docx` in Word and the change is sitting there as a revision,
attributed to the author and date you chose.

## Why another docx tool

Most "edit a Word file" scripts do a string replace inside `word/document.xml`
and save. That works until it doesn't: mixed bold inside a sentence, text split
across runs, a table cell, a hyperlink, a bookmark — and the formatting quietly
collapses into one style. Adding revision marks correctly is harder still, and
getting *reject* to restore the original document is where most implementations
give up.

`docxtrack` is built around three guarantees:

| Guarantee | What it means in practice |
|---|---|
| **Fidelity** | `rPr`, tables, bookmarks, comment anchors, `sectPr` and every other package part are preserved. Paragraph-level edits only rewrite the runs they touch. |
| **Reversibility** | `accept` produces exactly the revised document. `reject` produces exactly the original. Both are verified paragraph by paragraph by the test suite. |
| **Atomicity** | A batch of edits either applies completely or writes nothing. Every anchor is resolved against the original document before the first mutation. |

And one convenience that matters for review work: **word-level diffs**. Setting
a paragraph does not strike the old paragraph and mark the new one in full — it
marks only the words that actually changed, each keeping its original bold,
italic, font and size.

## Install

```bash
pip install git+https://github.com/socamalo/docx-track.git
```

With [uv](https://github.com/astral-sh/uv), the recommended way for a CLI:

```bash
git clone https://github.com/socamalo/docx-track.git
cd docx-track
uv tool install --editable .
```

You get a `docxtrack` executable on your PATH.

## Quick start

```bash
# 1. Look at the document. Paragraphs are numbered from 0.
docxtrack list answer.docx
docxtrack list answer.docx --grep fence

# 2. Describe the edits, once, in JSON.
cat > ops.json <<'JSON'
{
  "author": "Reviewer",
  "date": "2026-01-01T00:00:00Z",
  "ops": [
    {"op": "replace", "para": "mentions a fence", "find": "fence", "with": "port boundary"},
    {"op": "set",     "para": 12, "text": "Shortened answer that fits the word limit."},
    {"op": "insert_after", "para": 12, "text": "Evidence link: [URL pending.]", "style": "BodyText"},
    {"op": "comment", "para": 0, "text": "Check the title against the entry form."}
  ]
}
JSON

# 3. Validate addresses first, then write.
docxtrack apply answer.docx ops.json --dry-run
docxtrack apply answer.docx ops.json -o answer_v2.docx
docxtrack verify answer_v2.docx

# 4. Ship the tracked file, or settle it yourself.
docxtrack accept answer_v2.docx -o answer_v2_final.docx --drop-comments
docxtrack reject answer_v2.docx -o answer_v2_original.docx
```

Two versions to compare? You do not need to describe the edits at all:

```bash
docxtrack diff v1.docx v2.docx -o review.docx
```

`review.docx` opens in Word with every difference already marked as a revision.

## Commands

```text
# read
docxtrack list FILE [--json] [--grep TEXT] [--from N] [--to N] [--all]
docxtrack show FILE PARA [--all] [--runs]
docxtrack find FILE TEXT [--regex] [--all] [--json]

# edit — one change per call
docxtrack set FILE PARA TEXT [-o OUT] [--text-file F] [--no-track]
docxtrack insert FILE PARA TEXT [-o OUT] [--style S] [--no-track]
docxtrack delete FILE PARA [-o OUT] [--hard] [--to N]
docxtrack replace FILE PARA --find A --with B [-o OUT] [--every] [--no-track]
docxtrack replace-all FILE --find A --with B [-o OUT] [--every]
docxtrack merge FILE PARA [--count N] [--separator S] [-o OUT]
docxtrack split FILE PARA --at MARKER [-o OUT]
docxtrack comment FILE PARA TEXT [-o OUT]

# edit — the whole set at once
docxtrack apply FILE ops.json -o OUT [--dry-run] [--json]

# compare and settle
docxtrack diff OLD.docx NEW.docx -o OUT [--author A] [--date D]
docxtrack compare A.docx B.docx [--json] [--all]      # exit 0 equal, 1 different
docxtrack accept FILE [-o OUT] [--drop-comments] [--dry-run]
docxtrack reject FILE [-o OUT] [--dry-run]
docxtrack verify FILE [--json]
```

`PARA` is either a **0-based paragraph index** in document order, or a
**substring that matches exactly one paragraph**. A substring that matches two
paragraphs is a hard error, and the message lists the candidates — no silent
"first match wins" surprises in an automated pipeline.

`list` and `show` display the document *with revisions accepted*. Add `--all` to
see struck-out text as `[-deleted-]` / `[+inserted+]`, and `--runs` to inspect
per-run formatting.

## Batch `ops.json`

```json
{
  "author": "Reviewer",
  "date": "2026-01-01T00:00:00Z",
  "ops": [
    {"op": "replace", "para": "mentions a fence", "find": "fence", "with": "port boundary"},
    {"op": "set", "para": 12, "text": "Full new paragraph text."},
    {"op": "insert_after", "para": 12, "text": "New paragraph.", "style": "BodyText"},
    {"op": "delete", "para": "obsolete sentence", "hard": false},
    {"op": "comment", "para": 0, "text": "Check the title."}
  ]
}
```

| op | fields | description |
|---|---|---|
| `set` | `para`, `text`, `track?`, `comment?` | replace the paragraph with a word-level diff |
| `replace` | `para`, `find`, `with`, `track?`, `every?`, `comment?` | replace inside one paragraph |
| `insert_after` | `para`, `text`, `style?`, `track?`, `comment?` | insert a paragraph after |
| `insert_before` | `para`, `text`, `style?`, `track?`, `comment?` | insert a paragraph before |
| `replace_all` | `find`, `with`, `every?`, `track?` | every paragraph; no `para` needed |
| `split` | `para`, `at`, `track?` | split at `at` into two paragraphs |
| `delete` | `para`, `hard?`, `comment?` | strikethrough by default, `hard: true` removes it |
| `comment` | `para`, `text` | comment only, body untouched |

Addresses are resolved against the original document before anything is
mutated, so inserting a paragraph earlier in the batch never shifts the index of
a later op. If one anchor fails, the batch stops and nothing is written.

The same paragraph can be edited more than once in a batch. Each edit rebuilds
its diff against the pre-revision original, so `accept` still lands on the final
text and `reject` still returns to the untouched paragraph; the earlier
revisions on that paragraph simply merge into one diff.

## Python API

The CLI is a thin layer over a library you can call directly.

```python
from docxtrack import Docx, apply_ops, build_tracked, accept_all, reject_all, compare_docs, verify

d = Docx("answer.docx")
para = d.resolve("fence")                  # unique substring or index
d.set_paragraph(para, "port boundary")     # tracked by default
d.save("answer_v2.docx")

build_tracked("v1.docx", "v2.docx", "v2_review.docx")   # -> tracked document
d = Docx("v2_review.docx")
accept_all(d)
d.save("v2_final.docx")
```

## How it stays faithful

The engine never rebuilds a document from scratch. It loads the package, edits
the XML tree in place, and writes every other part out unchanged:

- Paragraph edits touch only the runs inside that paragraph. Unchanged words are
  re-emitted with their original `rPr`, and inserted text inherits the formatting
  of the word it was inserted next to.
- `diff` aligns documents at paragraph level, then reuses the same word-level
  insertion/deletion machinery for each pair of changed paragraphs.
- Table cells and text boxes are addressable like any other paragraph; headers,
  footers, footnotes and endnotes are outside the addressing range and are
  preserved untouched.
- Comments get a proper `comments.xml` part, content-type override and
  relationship, created on demand and dropped cleanly by `--drop-comments`.

Known boundaries, stated plainly: `diff` compares text, so a formatting-only
change produces no revision; `delete --hard` edits the XML directly and cannot
be undone by `reject`; a document open in Word cannot be written to until it is
closed.

## Tests

```bash
python -m unittest discover -s tests
```

26 tests cover tracked `set` (mixed formatting, repeated edits), `replace`,
`replace-all`, `merge`, `split`, inserts, tracked/hard/range deletes,
two-version `diff` (tables, paragraph prepended before the first one), batch
`apply` (anchors, ambiguity, out-of-range, atomic failure), `compare`, comments,
accept/reject round-trips, package integrity and CLI smoke tests.

The `.docx` fixtures are checked in together with the `.md` files they were
generated from, so the suite runs without Word and the expected content is
readable in review.

## License

MIT — see [LICENSE](LICENSE).
