# -*- coding: utf-8 -*-
"""docxtrack - command line interface.

Commands
--------
  docxtrack list FILE [--json] [--grep TEXT] [--from N] [--to N]
  docxtrack show FILE PARA
  docxtrack find FILE TEXT [--regex]
  docxtrack apply FILE OPS.json -o OUT [--dry-run]
  docxtrack set FILE PARA TEXT [-o OUT] [--no-track]
  docxtrack insert FILE PARA TEXT [-o OUT] [--style S] [--no-track]
  docxtrack delete FILE PARA [-o OUT] [--hard]
  docxtrack replace FILE PARA --find TEXT --with TEXT [-o OUT] [--no-track]
  docxtrack comment FILE PARA TEXT [-o OUT]
  docxtrack diff OLD.docx NEW.docx -o OUT [--author A] [--date D]
  docxtrack accept FILE [-o OUT] [--drop-comments]
  docxtrack reject FILE [-o OUT]
  docxtrack verify FILE [--json]

PARA is either a 0-based paragraph index or a unique substring of its text.
"""
from __future__ import annotations

import argparse
import json
import re
import sys

from . import core


def _addr(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return v


def _read_text(args):
    if getattr(args, "text_file", None):
        with open(args.text_file, encoding="utf-8") as f:
            return f.read().rstrip("\n")
    if args.text is None:
        raise core.DocxError("missing text (pass TEXT or --text-file)")
    return args.text


def _dump_paras(doc, args):
    rows = []
    lo = getattr(args, "from_", None)
    hi = getattr(args, "to", None)
    for i, p in enumerate(doc.paras):
        if lo is not None and i < lo:
            continue
        if hi is not None and i > hi:
            continue
        text = core.visible_text(p)
        annotated = core.annotated_text(p)
        if getattr(args, "grep", None) and args.grep not in core._text_of(p):
            continue
        row = {"index": i, "style": core._style_of(p), "text": text}
        if annotated != text:
            row["tracked_text"] = annotated
        rows.append(row)
    if getattr(args, "json", False):
        print(json.dumps(rows, ensure_ascii=False, indent=2))
    else:
        for r in rows:
            style = ("[%s] " % r["style"]) if r["style"] else ""
            print("#%03d %s%s" % (r["index"], style, r["text"]))


def cmd_list(args):
    doc = core.Docx(args.file)
    _dump_paras(doc, args)


def cmd_show(args):
    doc = core.Docx(args.file)
    p = doc.resolve(_addr(args.para))
    print("index: %d" % doc.index_of(p))
    print("style: %s" % (core._style_of(p) or "(none)"))
    print("text:")
    print(core.visible_text(p))
    if args.all:
        annotated = core.annotated_text(p)
        if annotated != core.visible_text(p):
            print("with revisions ([-deleted-] / [+inserted+]):")
            print(annotated)
    if args.runs:
        print("runs:")
        for k, r in enumerate(p.findall(core.q("r"))):
            rpr = r.find(core.q("rPr"))
            flags = []
            if rpr is not None:
                for tag in ("b", "i", "u", "strike"):
                    if rpr.find(core.q(tag)) is not None:
                        flags.append(tag)
            txt = "".join(t.text or "" for t in r.iter(core.q("t")))
            print("  [%d] %s %r" % (k, ",".join(flags) or "-", txt))


def cmd_find(args):
    doc = core.Docx(args.file)
    if args.all:
        hits = []
        for i, p in enumerate(doc.paras):
            t = core._text_of(p)
            ok = re.search(args.text, t) if args.regex else (args.text in t)
            if ok:
                hits.append((i, core._style_of(p), core.annotated_text(p)))
    else:
        hits = doc.find(args.text, regex=args.regex)
    if args.json:
        print(json.dumps(
            [{"index": i, "style": s, "text": t} for i, s, t in hits],
            ensure_ascii=False, indent=2))
    else:
        for i, style, text in hits:
            print("#%03d %s" % (i, text))
    if not hits:
        if not args.json:
            print("no match")
        return 2
    return 0


def _report(args, ok, **fields):
    if getattr(args, "json", False):
        payload = {"ok": ok}
        payload.update(fields)
        print(json.dumps(payload, ensure_ascii=False))
    else:
        msg = "OK" if ok else "FAILED"
        if fields:
            msg += ": " + ", ".join("%s=%s" % (k, v) for k, v in fields.items())
        print(msg)


def cmd_apply(args):
    with open(args.ops, encoding="utf-8") as f:
        spec = json.load(f)
    doc = core.Docx(args.file)
    n = core.apply_ops(doc, spec, dry_run=args.dry_run)
    if args.dry_run:
        _report(args, True, ops=n, dry_run=True)
        return 0
    out = args.out or args.file
    doc.save(out)
    _report(args, True, ops=n, out=out)


def _simple(args, fn):
    doc = core.Docx(args.file)
    p = doc.resolve(_addr(args.para))
    fn(doc, p, args)
    out = args.out or args.file
    doc.save(out)
    print("OK -> %s" % out)


def cmd_set(args):
    text = _read_text(args)
    _simple(args, lambda d, p, a: d.set_paragraph(p, text, not a.no_track))


def cmd_insert(args):
    text = _read_text(args)
    _simple(args, lambda d, p, a: d.insert_after(p, text, a.style, not a.no_track))


def cmd_delete(args):
    if getattr(args, "to", None) is not None:
        start = _addr(args.para)
        if not isinstance(start, int):
            raise core.DocxError("--to requires an integer start index")
        if args.to < start:
            raise core.DocxError("--to must be >= start")
        doc = core.Docx(args.file)
        for i in range(start, args.to + 1):
            doc.delete_paragraph(doc.resolve(i), args.hard)
        out = args.out or args.file
        doc.save(out)
        _report(args, True, paragraphs=args.to - start + 1, out=out)
        return 0
    _simple(args, lambda d, p, a: d.delete_paragraph(p, a.hard))


def cmd_replace(args):
    _simple(args, lambda d, p, a: d.replace_text(
        p, a.find, a.with_text, not a.no_track, every=a.every))


def cmd_comment(args):
    text = _read_text(args)
    _simple(args, lambda d, p, a: d.add_comment(p, text))


def cmd_split(args):
    doc = core.Docx(args.file)
    p = doc.resolve(_addr(args.para))
    before, after = doc.split_paragraph(p, args.at, not args.no_track)
    out = args.out or args.file
    doc.save(out)
    _report(args, True, before=len(before), after=len(after), out=out)


def cmd_merge(args):
    doc = core.Docx(args.file)
    p = doc.resolve(_addr(args.para))
    n = doc.merge_next(p, args.count, args.separator, not args.no_track)
    out = args.out or args.file
    doc.save(out)
    _report(args, True, paragraphs=n, out=out)


def cmd_replace_all(args):
    doc = core.Docx(args.file)
    n = doc.replace_all(args.find, args.with_text, not args.no_track,
                        every=args.every)
    if n == 0:
        sys.stderr.write("ERROR: text not found in any paragraph: %r\n" % args.find)
        return 2
    out = args.out or args.file
    doc.save(out)
    _report(args, True, paragraphs=n, out=out)


def cmd_compare(args):
    info = core.compare_docs(args.a, args.b, include_deleted=args.all)
    if args.json:
        print(json.dumps(info, ensure_ascii=False, indent=2))
    else:
        print("EQUAL" if info["equal"] else "DIFFERENT")
        print("paragraphs: %d vs %d" % (info["paragraphs_a"], info["paragraphs_b"]))
        for d in info["diffs"][:10]:
            print("#%03d A=%r B=%r" % (d["index"], d["a"], d["b"]))
    return 0 if info["equal"] else 1


def cmd_diff(args):
    core.build_tracked(args.old, args.new, args.out, args.author, args.date)
    _report(args, True, out=args.out)


def _drop_comments(doc):
    if doc.comments is None:
        return 0
    ids = set(c.get(core.q("id")) for c in doc.comments.findall(core.q("comment")))
    for c in list(doc.comments):
        doc.comments.remove(c)
    for el in list(doc.root.iter()):
        if el.tag in (core.q("commentRangeStart"), core.q("commentRangeEnd")):
            if el.get(core.q("id")) in ids:
                el.getparent().remove(el)
        elif el.tag == core.q("commentReference"):
            r = el.getparent()
            if r is not None and r.getparent() is not None:
                r.getparent().remove(r)
    return len(ids)


def cmd_accept(args):
    doc = core.Docx(args.file)
    dropped = _drop_comments(doc) if args.drop_comments else 0
    core.accept_all(doc)
    if args.dry_run:
        _report(args, True, dropped_comments=dropped, dry_run=True)
        return 0
    out = args.out or args.file
    doc.save(out)
    _report(args, True, dropped_comments=dropped, out=out)


def cmd_reject(args):
    doc = core.Docx(args.file)
    core.reject_all(doc)
    if args.dry_run:
        _report(args, True, dry_run=True)
        return 0
    out = args.out or args.file
    doc.save(out)
    _report(args, True, out=out)


def cmd_verify(args):
    info = core.verify(args.file)
    if args.json:
        print(json.dumps(info, ensure_ascii=False, indent=2))
    else:
        for k, v in info.items():
            print("%s: %s" % (k, v))
    return 0 if info.get("zip_ok") else 1


def build_parser():
    ap = argparse.ArgumentParser(
        prog="docxtrack",
        description="Inspect and revise .docx files with Word tracked changes.",
    )
    ap.add_argument("--version", action="version", version="docxtrack 0.1.0")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("list", help="list paragraphs with indices")
    p.add_argument("file")
    p.add_argument("--json", action="store_true")
    p.add_argument("--grep", default=None)
    p.add_argument("--from", dest="from_", type=int, default=None)
    p.add_argument("--to", type=int, default=None)
    p.add_argument("--all", action="store_true", help="include deleted text in --json")
    p.set_defaults(func=cmd_list)

    p = sub.add_parser("show", help="show one paragraph")
    p.add_argument("file"); p.add_argument("para")
    p.add_argument("--all", action="store_true", help="also show deleted text")
    p.add_argument("--runs", action="store_true", help="show run-level formatting")
    p.set_defaults(func=cmd_show)

    p = sub.add_parser("find", help="find paragraphs containing text")
    p.add_argument("file"); p.add_argument("text")
    p.add_argument("--regex", action="store_true")
    p.add_argument("--all", action="store_true", help="include deleted text")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_find)

    p = sub.add_parser("apply", help="apply a JSON ops file")
    p.add_argument("file"); p.add_argument("ops")
    p.add_argument("-o", "--out")
    p.add_argument("--dry-run", action="store_true")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_apply)

    for name, fn in (("set", cmd_set), ("insert", cmd_insert)):
        p = sub.add_parser(name, help="%s a paragraph" % name)
        p.add_argument("file"); p.add_argument("para"); p.add_argument("text", nargs="?")
        p.add_argument("--text-file")
        p.add_argument("-o", "--out")
        p.add_argument("--no-track", action="store_true")
        if name == "insert":
            p.add_argument("--style", default=None)
        p.set_defaults(func=fn)

    p = sub.add_parser("delete", help="delete a paragraph (or a range with --to)")
    p.add_argument("file"); p.add_argument("para")
    p.add_argument("--to", type=int, default=None, help="delete [para..to] inclusive")
    p.add_argument("-o", "--out"); p.add_argument("--hard", action="store_true")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_delete)

    p = sub.add_parser("replace", help="replace text inside one paragraph")
    p.add_argument("file"); p.add_argument("para")
    p.add_argument("--find", required=True)
    p.add_argument("--with", dest="with_text", required=True)
    p.add_argument("--every", action="store_true", help="replace every occurrence")
    p.add_argument("-o", "--out"); p.add_argument("--no-track", action="store_true")
    p.set_defaults(func=cmd_replace)

    p = sub.add_parser("split", help="split a paragraph at a marker")
    p.add_argument("file"); p.add_argument("para")
    p.add_argument("--at", required=True, help="marker text where the split starts")
    p.add_argument("-o", "--out"); p.add_argument("--no-track", action="store_true")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_split)

    p = sub.add_parser("merge", help="merge following paragraphs into this one")
    p.add_argument("file"); p.add_argument("para")
    p.add_argument("--count", type=int, default=1)
    p.add_argument("--separator", default=" ")
    p.add_argument("-o", "--out"); p.add_argument("--no-track", action="store_true")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_merge)

    p = sub.add_parser("replace-all", help="replace text in every paragraph")
    p.add_argument("file")
    p.add_argument("--find", required=True)
    p.add_argument("--with", dest="with_text", required=True)
    p.add_argument("--every", action="store_true", help="replace every occurrence")
    p.add_argument("-o", "--out"); p.add_argument("--no-track", action="store_true")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_replace_all)

    p = sub.add_parser("compare", help="compare paragraph text of two documents")
    p.add_argument("a"); p.add_argument("b")
    p.add_argument("--all", action="store_true", help="include deleted text")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_compare)

    p = sub.add_parser("comment", help="attach a comment to a paragraph")
    p.add_argument("file"); p.add_argument("para"); p.add_argument("text", nargs="?")
    p.add_argument("--text-file"); p.add_argument("-o", "--out")
    p.set_defaults(func=cmd_comment)

    p = sub.add_parser("diff", help="build tracked changes from two documents")
    p.add_argument("old"); p.add_argument("new"); p.add_argument("-o", "--out", required=True)
    p.add_argument("--author", default=core.DEFAULT_AUTHOR)
    p.add_argument("--date", default=core.DEFAULT_DATE)
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_diff)

    p = sub.add_parser("accept", help="accept all revisions")
    p.add_argument("file"); p.add_argument("-o", "--out")
    p.add_argument("--drop-comments", action="store_true")
    p.add_argument("--dry-run", action="store_true")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_accept)

    p = sub.add_parser("reject", help="reject all revisions")
    p.add_argument("file"); p.add_argument("-o", "--out")
    p.add_argument("--dry-run", action="store_true")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_reject)

    p = sub.add_parser("verify", help="count revisions and check the package")
    p.add_argument("file"); p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_verify)
    return ap


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    ap = build_parser()
    args = ap.parse_args(argv)
    try:
        rc = args.func(args)
    except core.DocxError as e:
        sys.stderr.write("ERROR: %s\n" % e)
        return 1
    except (OSError, ValueError, KeyError) as e:
        sys.stderr.write("ERROR: %s\n" % e)
        return 1
    return rc or 0


if __name__ == "__main__":
    sys.exit(main())
