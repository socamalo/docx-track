---
name: docx-track
description: Read and revise .docx files with Word tracked changes from the command line. List paragraphs by index or unique substring, replace/insert/delete text as w:ins/w:del, diff two Word files into a tracked document, accept/reject revisions, add comments, and verify the package. Use whenever an agent must edit a Word file for review, leave revision marks, or produce a clean copy from a revised one.
---

# docx-track — Word tracked changes from the CLI

One executable, `docxtrack`, that turns "read a Word file, change it, leave a
revision mark" into scriptable commands. No Word installation: it rewrites
`word/document.xml` (and `word/comments.xml` when comments are used) and copies
every other package part unchanged.

## Before you start

Confirm the command is available:

```bash
docxtrack --version        # docxtrack 0.1.0
```

If it is not on PATH, install it once (`pip install git+https://github.com/socamalo/docx-track.git`
or `uv tool install --editable <repo>`), then call the installed executable by
its full path in scripts.

## When to use this skill

- Applying review feedback to a `.docx` where the change must be visible as a
  revision (`w:ins` / `w:del`).
- Comparing two versions of a Word file and handing over the differences as a
  tracked document.
- Accepting or rejecting all revisions to get a clean copy or restore the
  original.
- Reading a paragraph, locating it by keyword, or attaching a comment.

## Core concepts

**Paragraph address (`PARA`)** — two forms, interchangeable in every command:

1. **Index**: 0-based, in document order, including paragraphs inside tables.
   `0` is the first paragraph.
2. **Unique substring**: a fragment of the paragraph text. It must match exactly
   one paragraph; ambiguity is an error and the message lists the candidates.
   More stable than an index and better for scripting.

**Tracked by default** — `set`, `replace`, `insert`, `delete` write revisions.
`--no-track` edits directly, and `delete --hard` removes the paragraph from the
document entirely.

**Round-trip** — `accept` gives the revised document, `reject` gives the original.
Both are verified byte-for-byte at paragraph level by the test suite.

## Commands

```bash
# read
docxtrack list FILE                      # #000 [Heading1] text...
docxtrack list FILE --json               # [{"index":0,"style":"...","text":"..."}]
docxtrack list FILE --grep TEXT          # only matching paragraphs
docxtrack list FILE --from N --to M      # limit the range
docxtrack list FILE --all --json         # include tracked_text with [-del-]/[+ins+]
docxtrack show FILE 12                   # one paragraph (index or substring)
docxtrack show FILE 12 --all             # include struck-out text
docxtrack show FILE 12 --runs            # per-run bold/italic/underline/strike
docxtrack find FILE "fence"              # paragraphs containing the string
docxtrack find FILE "fence" --all        # also search deleted text
docxtrack find FILE "fence" --json       # machine readable

# edit — one change per call
docxtrack set FILE 12 "new text" -o out.docx
docxtrack set FILE 12 --text-file body.txt -o out.docx
docxtrack insert FILE 12 "new paragraph" --style BodyText -o out.docx
docxtrack delete FILE 12 -o out.docx              # strikethrough
docxtrack delete FILE 12 --hard -o out.docx       # paragraph removed outright
docxtrack delete FILE 12 --to 18 -o out.docx      # delete 12..18 inclusive
docxtrack replace FILE 12 --find "fence" --with "boundary" -o out.docx
docxtrack replace FILE 12 --find "fence" --with "boundary" --every -o out.docx
docxtrack replace-all FILE --find "fence" --with "boundary" -o out.docx
docxtrack merge FILE 12 --count 2 -o out.docx     # fold the next 2 paragraphs in
docxtrack split FILE 12 --at "3, How does" -o out.docx
docxtrack comment FILE 12 "why" -o out.docx

# edit — batch (preferred for more than one change)
docxtrack apply FILE ops.json -o out.docx
docxtrack apply FILE ops.json --dry-run           # validate addresses, write nothing

# compare two versions
docxtrack diff old.docx new.docx -o tracked.docx

# settle and inspect
docxtrack accept tracked.docx -o clean.docx [--drop-comments] [--dry-run]
docxtrack reject tracked.docx -o original.docx [--dry-run]
docxtrack verify FILE [--json]     # paragraphs / ins / del / delText / comments / zip_ok
docxtrack compare A.docx B.docx [--json] [--all]  # exit 0 equal, 1 different
```

`apply`, `diff`, `accept`, `reject`, `verify`, `compare`, `list`, `find`, and the
range form of `delete` accept `--json` for machine-readable output.

**Do not write Python or temporary scripts for a revision task.** Reading,
editing, comparing, accepting/rejecting, commenting and verifying are all covered
by the commands above. Only two kinds of file are meant to be written: `ops.json`
(what to change) and `--text-file` (the body text to insert).

### Batch `ops.json`

```json
{
  "author": "Reviewer",
  "date": "2026-01-01T00:00:00Z",
  "ops": [
    {"op": "replace", "para": "mentions a fence", "find": "fence", "with": "port boundary"},
    {"op": "set",     "para": 12, "text": "Full new paragraph text."},
    {"op": "insert_after", "para": 12, "text": "New paragraph.", "style": "BodyText"},
    {"op": "delete",  "para": "obsolete sentence", "hard": false},
    {"op": "comment", "para": 0, "text": "Check the title."}
  ]
}
```

| op | fields | description |
|---|---|---|
| `set` | `para`, `text`, `track?`, `comment?` | replace the paragraph; word-level diff, unchanged words stay clean |
| `replace` | `para`, `find`, `with`, `track?`, `every?`, `comment?` | replace inside one paragraph |
| `insert_after` | `para`, `text`, `style?`, `track?`, `comment?` | insert a paragraph after |
| `insert_before` | `para`, `text`, `style?`, `track?`, `comment?` | insert a paragraph before |
| `replace_all` | `find`, `with`, `every?`, `track?` | every paragraph, no `para` needed |
| `split` | `para`, `at`, `track?` | split at `at` into two paragraphs |
| `delete` | `para`, `hard?`, `comment?` | strikethrough by default, `hard: true` removes it |
| `comment` | `para`, `text` | comment only, body untouched |

All `para` values resolve against the original document at load time, so an
insert or delete earlier in the batch never shifts a later address. Any failed
address aborts the whole batch without writing.

One paragraph can be edited several times in a batch (e.g. `replace_all` then
`split`): each edit rebuilds the diff against the pre-revision original, so
`accept` lands on the final text and `reject` returns to the untouched original.
The cost is that earlier revisions on that paragraph merge into one diff.

## Recommended agent workflow

1. `docxtrack list FILE --json` (or `--grep`) to find the target paragraphs and
   note their index or a unique fragment.
2. Write the changes into `ops.json` (UTF-8).
3. `apply --dry-run` to validate every address without touching the file.
4. `apply` to write; `verify` to confirm `ins` / `del` are non-zero.
5. Hand over the tracked file as-is. Produce a clean copy with `accept -o clean.docx`
   when the recipient does not want revision marks.
6. To confirm nothing else moved, `reject` a copy and `compare` it with the
   original.

## Pitfalls and conventions

- **Keep non-ASCII text out of the command line on Windows bash** — a command
  containing non-ASCII characters is rejected wholesale (exit 127). Put body text
  in `ops.json` or use `--text-file`.
- **Anchors must be unique** — a substring that matches more than one paragraph
  raises an error listing the candidates; use an index or a longer fragment.
- **A file open in Word cannot be written** (PermissionError). Close the document
  first.
- **`--hard` is for internal working material** (version notes, review checklists,
  shortlists). It edits the XML directly, so `reject` cannot restore it. If the
  deletion should be visible and reversible, use the default tracked delete.
  `delete --to` is the range form of the same command.
- **Table paragraphs are addressable** in document order; `insert_after` inserts
  into the same cell.
- **Formatting is preserved** — `set` diffs word by word and keeps each word's
  original bold/font; inserted text inherits the formatting at the insertion
  point. A paragraph with mixed formatting stays sliced per run.
- **`diff` boundaries** — both inputs should be clean documents (no pending
  revisions); the output is the tracked document. Alignment is per paragraph and
  compares text only, so a pure formatting change yields no revision. Good for
  prose documents: answers, press releases, reports.
- **Scope** — body paragraphs of `word/document.xml`, including table cells and
  text boxes. Headers, footers, footnotes and endnotes are outside the addressing
  range: preserved as-is, never modified.
- **`list` / `show` show text with revisions accepted**; add `--all` for the
  struck-out original. Substring addressing also only sees visible text, so
  already-deleted sentences cannot be targeted by accident.

## Typical session

```bash
# Review note: "explain 'fence' as an electronic boundary, and trim the long answer"
docxtrack list answer.docx --grep fence
# write ops.json: replace fence -> port boundary; set the long answer -> short answer
docxtrack apply answer.docx ops.json --dry-run
docxtrack apply answer.docx ops.json -o answer_v2.docx
docxtrack verify answer_v2.docx

# Show a client what changed between two versions
docxtrack diff answer_v1.docx answer_v2.docx -o answer_v2_review.docx

# Final: accept everything, drop internal comments
docxtrack accept answer_v2_review.docx -o answer_v2_final.docx --drop-comments
```
