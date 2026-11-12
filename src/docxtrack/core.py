# -*- coding: utf-8 -*-
"""docxtrack.core - tracked-changes engine for .docx files.

Design goals
------------
* Paragraph-addressed: every operation names a paragraph by its 0-based index
  in document order, or by a unique substring of its text.
* Tracked by default: text edits become w:ins / w:del; paragraph deletes also
  delete the paragraph mark, so "accept all" in Word gives the clean result.
* Word-level diffs: setting a paragraph keeps unchanged words untouched
  instead of striking the whole paragraph.
* Round-trip safe: accept(reject) of a tracked file yields exactly the revised
  (original) document.
* Everything else in the package is copied byte-for-byte.

Only word/document.xml (and word/comments.xml when comments are used) is
rewritten.  No Word installation is required.
"""
from __future__ import annotations

import copy
import difflib
import os
import re
import zipfile

from lxml import etree

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_COMMENTS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
CT_COMMENTS = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

DEFAULT_AUTHOR = "Reviewer"
DEFAULT_DATE = "2026-01-01T00:00:00Z"


def q(tag: str) -> str:
    return "{%s}%s" % (NS_W, tag)


# --------------------------------------------------------------------------
# small XML helpers
# --------------------------------------------------------------------------

def _text_of(p, include_deleted: bool = True) -> str:
    out = []
    for node in p.iter():
        if node.tag == q("t") or (include_deleted and node.tag == q("delText")):
            out.append(node.text or "")
    return "".join(out)


def visible_text(p) -> str:
    """Text a reader sees after accepting revisions (deleted text excluded)."""
    return _text_of(p, include_deleted=False)


def original_text(p) -> str:
    """Text before any revisions: deleted text restored, inserted text removed."""
    out = []
    for node in p.iter():
        if node.tag == q("t"):
            parent = node.getparent()
            inside_ins = parent is not None and parent.getparent() is not None \
                and parent.getparent().tag == q("ins")
            if not inside_ins:
                out.append(node.text or "")
        elif node.tag == q("delText"):
            out.append(node.text or "")
    return "".join(out)


def annotated_text(p) -> str:
    """Text with [-deleted-] and [+inserted+] markers, in run order."""
    out = []
    for node in p.iter():
        if node.tag == q("t"):
            parent = node.getparent()
            inside_ins = parent is not None and parent.getparent() is not None \
                and parent.getparent().tag == q("ins")
            out.append(("[+%s+]" % node.text) if inside_ins else (node.text or ""))
        elif node.tag == q("delText"):
            out.append("[-%s-]" % node.text)
    return "".join(out)


def _style_of(p) -> str:
    st = p.find(q("pPr") + "/" + q("pStyle"))
    if st is None:
        return ""
    return st.get(q("val")) or ""


def _direct_runs(p):
    return [r for r in p if r.tag == q("r")]


def _first_rpr(p):
    for r in p.iter(q("r")):
        rp = r.find(q("rPr"))
        if rp is not None:
            return rp
    return None


def _uniform_rpr(p):
    """True when all direct runs share one formatting; returns (rPr, uniform)."""
    rprs = []
    for r in _direct_runs(p):
        rp = r.find(q("rPr"))
        rprs.append(etree.tostring(rp) if rp is not None else b"")
    rpr = _first_rpr(p)
    return rpr, len(set(rprs)) <= 1


def _make_run(text, rpr=None):
    r = etree.Element(q("r"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, q("t"))
    t.set(XML_SPACE, "preserve")
    t.text = text
    return r


def _make_mark(tag, text, wid, author, date, rpr=None):
    el = etree.Element(q(tag))
    el.set(q("id"), str(wid))
    el.set(q("author"), author)
    el.set(q("date"), date)
    r = etree.SubElement(el, q("r"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, q("delText") if tag == "del" else q("t"))
    t.set(XML_SPACE, "preserve")
    t.text = text
    return el


def _unwrap(el):
    parent = el.getparent()
    idx = list(parent).index(el)
    children = list(el)
    parent.remove(el)
    for i, child in enumerate(children):
        parent.insert(idx + i, child)


def _word_diff_spans(old, new):
    """Yield (kind, text, old_start, old_end) with kind in keep/del/ins.

    Offsets are character positions in ``old`` (useful for slicing runs).
    """
    a = re.findall(r"\s+|\S+", old)
    b = re.findall(r"\s+|\S+", new)
    oa = [0]
    for tok in a:
        oa.append(oa[-1] + len(tok))
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        os_, oe = oa[i1], oa[i2]
        if tag == "equal":
            yield "keep", "".join(a[i1:i2]), os_, oe
        else:
            if i1 < i2:
                yield "del", "".join(a[i1:i2]), os_, oe
            if j1 < j2:
                yield "ins", "".join(b[j1:j2]), os_, oe


def _word_diff(old, new):
    for kind, frag, _s, _e in _word_diff_spans(old, new):
        yield kind, frag


def _format_map(p):
    """[(start, end, rPr, text)] over the *original* text (before revisions).

    Deleted runs count (their delText), inserted runs are skipped. This lets a
    second edit rebuild diff(original -> new) and keep reject exact.
    """
    spans = []
    pos = 0
    for r in p.iter(q("r")):
        node = r.getparent()
        inserted = False
        while node is not None and node is not p:
            if node.tag == q("ins"):
                inserted = True
                break
            node = node.getparent()
        if inserted:
            continue
        txt = "".join((t.text or "") for t in r.iter()
                      if t.tag in (q("t"), q("delText")))
        rpr = r.find(q("rPr"))
        spans.append((pos, pos + len(txt), rpr, txt))
        pos += len(txt)
    return spans


def _slice_spans(spans, start, end):
    """Yield (text, rPr) covering [start, end) using the format map."""
    for s, e, rpr, txt in spans:
        a, b = max(s, start), min(e, end)
        if a < b:
            yield txt[a - s:b - s], rpr


def _rpr_at(spans, pos):
    if not spans:
        return None
    for s, e, rpr, _txt in spans:
        if s <= pos < e:
            return rpr
    return spans[-1][2] if spans else None


class DocxError(Exception):
    pass


# --------------------------------------------------------------------------
# document model
# --------------------------------------------------------------------------

class Docx:
    def __init__(self, path):
        self.path = path
        if not os.path.exists(path):
            raise DocxError("file not found: %s" % path)
        try:
            z = zipfile.ZipFile(path)
        except zipfile.BadZipFile:
            raise DocxError("not a valid .docx (zip): %s" % path)
        with z:
            self.names = list(z.namelist())
            self.parts = {n: z.read(n) for n in self.names}
        if "word/document.xml" not in self.parts:
            raise DocxError("missing word/document.xml: %s" % path)
        self.root = etree.fromstring(self.parts["word/document.xml"])
        self.body = self.root.find(q("body"))
        self.paras = [p for p in self.body.iter(q("p"))]
        self.comments = None
        if "word/comments.xml" in self.parts:
            self.comments = etree.fromstring(self.parts["word/comments.xml"])
        self._next_id = self._max_id() + 1000
        self._next_comment = self._max_comment() + 1

    # ---- ids ---------------------------------------------------------
    def _max_id(self):
        mx = 0
        for el in self.root.iter():
            v = el.get(q("id"))
            if v and v.isdigit():
                mx = max(mx, int(v))
        return mx

    def _max_comment(self):
        if self.comments is None:
            return 0
        mx = 0
        for c in self.comments.findall(q("comment")):
            v = c.get(q("id"))
            if v and v.isdigit():
                mx = max(mx, int(v))
        return mx

    def nid(self):
        self._next_id += 1
        return self._next_id

    # ---- addressing --------------------------------------------------
    def resolve(self, addr):
        """int index or unique substring -> paragraph element."""
        if isinstance(addr, bool):
            raise DocxError("invalid paragraph address: %r" % addr)
        if isinstance(addr, int):
            if addr < 0 or addr >= len(self.paras):
                raise DocxError(
                    "paragraph index out of range: %d (0..%d)"
                    % (addr, len(self.paras) - 1)
                )
            return self.paras[addr]
        if not isinstance(addr, str):
            raise DocxError("invalid paragraph address: %r" % addr)
        needle = addr.strip()
        if not needle:
            raise DocxError("empty paragraph anchor")
        hits = [p for p in self.paras if needle in visible_text(p)]
        if not hits:
            raise DocxError("no paragraph contains: %r" % needle)
        if len(hits) > 1:
            where = "; ".join(
                "#%d %r" % (self.paras.index(p), _text_of(p)[:40]) for p in hits[:6]
            )
            raise DocxError(
                "anchor is ambiguous (%d matches): %r -> %s"
                % (len(hits), needle, where)
            )
        return hits[0]

    def index_of(self, p):
        return self.paras.index(p)

    def find(self, needle, regex=False):
        out = []
        for i, p in enumerate(self.paras):
            t = visible_text(p)
            ok = re.search(needle, t) if regex else (needle in t)
            if ok:
                out.append((i, _style_of(p), t))
        return out

    # ---- paragraph mutation -----------------------------------------
    def _clear_runs(self, p):
        # Remove old content, including revision wrappers left by an earlier op
        # in the same batch. Otherwise a second edit would append runs beside
        # the previous w:ins/w:del and corrupt reject.
        for r in list(p):
            if r.tag in (q("r"), q("hyperlink"), q("smartTag"), q("ins"), q("del")):
                p.remove(r)

    def set_paragraph(self, p, text, track=True, author=DEFAULT_AUTHOR, date=DEFAULT_DATE):
        # Always diff against the pre-revision original, so repeated edits in
        # one batch still reject back to the untouched document.
        old = original_text(p)
        spans = _format_map(p)
        rpr = spans[0][2] if spans else _first_rpr(p)
        if not track:
            self._clear_runs(p)
            p.append(_make_run(text, rpr))
            return
        self._clear_runs(p)
        for kind, frag, os_, oe in _word_diff_spans(old, text):
            if not frag:
                continue
            if kind == "ins":
                p.append(_make_mark("ins", frag, self.nid(), author, date,
                                    _rpr_at(spans, os_ if os_ < len(old) else len(old) - 1)))
                continue
            segs = list(_slice_spans(spans, os_, oe)) or [(frag, rpr)]
            for seg, seg_rpr in segs:
                if not seg:
                    continue
                if kind == "keep":
                    p.append(_make_run(seg, seg_rpr))
                else:
                    p.append(_make_mark("del", seg, self.nid(), author, date, seg_rpr))

    def _new_paragraph(self, text, style, track, author, date):
        newp = etree.Element(q("p"))
        ppr = etree.SubElement(newp, q("pPr"))
        if style:
            st = etree.SubElement(ppr, q("pStyle"))
            st.set(q("val"), style)
        if track:
            newp.append(_make_mark("ins", text, self.nid(), author, date, None))
            self._mark_para_mark(newp, "ins", author, date)
        else:
            newp.append(_make_run(text, None))
        return newp

    def insert_after(self, p, text, style=None, track=True,
                     author=DEFAULT_AUTHOR, date=DEFAULT_DATE):
        newp = self._new_paragraph(text, style, track, author, date)
        p.addnext(newp)
        return newp

    def insert_before(self, p, text, style=None, track=True,
                      author=DEFAULT_AUTHOR, date=DEFAULT_DATE):
        newp = self._new_paragraph(text, style, track, author, date)
        p.addprevious(newp)
        return newp

    def delete_paragraph(self, p, hard=False, author=DEFAULT_AUTHOR, date=DEFAULT_DATE):
        if hard:
            p.getparent().remove(p)
            return
        for r in _direct_runs(p):
            rpr = r.find(q("rPr"))
            text = "".join(t.text or "" for t in r.iter(q("t")))
            p.replace(r, _make_mark("del", text, self.nid(), author, date, rpr))
        self._mark_para_mark(p, "del", author, date)

    def replace_text(self, p, find, with_text, track=True,
                     author=DEFAULT_AUTHOR, date=DEFAULT_DATE, every=False):
        old = _text_of(p)
        if find not in old:
            raise DocxError("text not found in paragraph: %r" % find)
        new = old.replace(find, with_text) if every else old.replace(find, with_text, 1)
        self.set_paragraph(p, new, track, author, date)

    def split_paragraph(self, p, marker, track=True,
                        author=DEFAULT_AUTHOR, date=DEFAULT_DATE):
        """Split a paragraph at ``marker`` into two paragraphs (tracked by default)."""
        text = visible_text(p)
        idx = text.find(marker)
        if idx <= 0:
            raise DocxError("split marker not found (or at start): %r" % marker)
        before = text[:idx].rstrip()
        after = text[idx:]
        self.set_paragraph(p, before, track, author, date)
        self.insert_after(p, after, _style_of(p) or None, track, author, date)
        return before, after

    def merge_next(self, p, count=1, separator=" ", track=True,
                   author=DEFAULT_AUTHOR, date=DEFAULT_DATE):
        """Merge the next ``count`` paragraphs into this one (tracked by default)."""
        parts = [visible_text(p)]
        nxts = []
        cur = p
        for _ in range(max(0, int(count))):
            nxt = cur.getnext()
            while nxt is not None and nxt.tag != q("p"):
                nxt = nxt.getnext()
            if nxt is None:
                break
            nxts.append(nxt)
            parts.append(visible_text(nxt))
            cur = nxt
        merged = separator.join(x for x in parts if x)
        self.set_paragraph(p, merged, track, author, date)
        for n in nxts:
            self.delete_paragraph(n, hard=not track, author=author, date=date)
        return len(nxts) + 1

    def replace_all(self, find, with_text, track=True,
                    author=DEFAULT_AUTHOR, date=DEFAULT_DATE, every=False):
        """Replace text in every paragraph that contains it (first hit per paragraph)."""
        n = 0
        for p in list(self.paras):
            old = _text_of(p)
            if find in old:
                new = old.replace(find, with_text) if every \
                    else old.replace(find, with_text, 1)
                self.set_paragraph(p, new, track, author, date)
                n += 1
        return n

    def _mark_para_mark(self, p, tag, author, date):
        ppr = p.find(q("pPr"))
        if ppr is None:
            ppr = etree.Element(q("pPr"))
            p.insert(0, ppr)
        rpr = ppr.find(q("rPr"))
        if rpr is None:
            rpr = etree.SubElement(ppr, q("rPr"))
        el = etree.SubElement(rpr, q(tag))
        el.set(q("id"), str(self.nid()))
        el.set(q("author"), author)
        el.set(q("date"), date)

    # ---- comments ----------------------------------------------------
    def _ensure_comments(self, author, date):
        if self.comments is not None:
            return
        self.comments = etree.Element(q("comments"), nsmap={"w": NS_W})
        self.parts["word/comments.xml"] = b""
        if "word/comments.xml" not in self.names:
            self.names.append("word/comments.xml")
        # content types
        ct = etree.fromstring(self.parts["[Content_Types].xml"])
        if not any(o.get("PartName") == "/word/comments.xml"
                   for o in ct.findall("{%s}Override" % NS_CT)):
            ov = etree.SubElement(ct, "{%s}Override" % NS_CT)
            ov.set("PartName", "/word/comments.xml")
            ov.set("ContentType", CT_COMMENTS)
            self.parts["[Content_Types].xml"] = etree.tostring(
                ct, xml_declaration=True, encoding="UTF-8", standalone=True)
        # relationship
        rels_name = "word/_rels/document.xml.rels"
        rels = etree.fromstring(self.parts[rels_name])
        if not any(r.get("Type") == REL_COMMENTS
                   for r in rels.findall("{%s}Relationship" % NS_REL)):
            rid = "rIdDocxTrackComments"
            r = etree.SubElement(rels, "{%s}Relationship" % NS_REL)
            r.set("Id", rid)
            r.set("Type", REL_COMMENTS)
            r.set("Target", "comments.xml")
            self.parts[rels_name] = etree.tostring(
                rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    def add_comment(self, p, text, author=DEFAULT_AUTHOR, date=DEFAULT_DATE):
        self._ensure_comments(author, date)
        cid = str(self._next_comment)
        self._next_comment += 1
        c = etree.SubElement(self.comments, q("comment"))
        c.set(q("id"), cid)
        c.set(q("author"), author)
        c.set(q("date"), date)
        cp = etree.SubElement(c, q("p"))
        cr = etree.SubElement(cp, q("r"))
        ct = etree.SubElement(cr, q("t"))
        ct.set(XML_SPACE, "preserve")
        ct.text = text
        # point anchor right after the paragraph mark
        anchor = p.find(q("pPr"))
        pos = 0 if anchor is None else list(p).index(anchor) + 1
        start = etree.Element(q("commentRangeStart")); start.set(q("id"), cid)
        end = etree.Element(q("commentRangeEnd")); end.set(q("id"), cid)
        ref = etree.Element(q("r"))
        cref = etree.SubElement(ref, q("commentReference")); cref.set(q("id"), cid)
        p.insert(pos, start); p.insert(pos + 1, end); p.insert(pos + 2, ref)
        return cid

    # ---- save --------------------------------------------------------
    def save(self, out):
        self.parts["word/document.xml"] = etree.tostring(
            self.root, xml_declaration=True, encoding="UTF-8", standalone=True)
        if self.comments is not None and "word/comments.xml" in self.parts:
            self.parts["word/comments.xml"] = etree.tostring(
                self.comments, xml_declaration=True, encoding="UTF-8", standalone=True)
        tmp = out + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for n in self.names:
                data = self.parts.get(n)
                if data is None:
                    continue
                z.writestr(n, data)
        os.replace(tmp, out)
        return out


# --------------------------------------------------------------------------
# batch ops
# --------------------------------------------------------------------------

def _check_keys(op, allowed, i):
    extra = set(op) - set(allowed)
    if extra:
        raise DocxError("op #%d: unknown field(s): %s" % (i, ", ".join(sorted(extra))))


def apply_ops(doc: Docx, ops, dry_run=False):
    """Apply a list of ops. Addresses are resolved before any mutation."""
    author = ops.get("author", DEFAULT_AUTHOR)
    date = ops.get("date", DEFAULT_DATE)
    items = ops.get("ops", [])
    if not isinstance(items, list):
        raise DocxError("'ops' must be a list")

    # Resolve every address first so a bad index never leaves a half-edited doc.
    targets = []
    for i, op in enumerate(items):
        if not isinstance(op, dict) or "op" not in op:
            raise DocxError("op #%d: missing 'op'" % i)
        kind = op["op"]
        if kind == "replace_all":
            targets.append(None)
            continue
        if "para" not in op:
            raise DocxError("op #%d (%s): missing 'para'" % (i, kind))
        try:
            p = doc.resolve(op["para"])
        except DocxError as e:
            raise DocxError("op #%d (%s): %s" % (i, kind, e))
        targets.append(p)

    for i, (op, p) in enumerate(zip(items, targets)):
        kind = op["op"]
        if kind == "set":
            _check_keys(op, {"op", "para", "text", "track", "comment"}, i)
            if "text" not in op:
                raise DocxError("op #%d (set): missing 'text'" % i)
            doc.set_paragraph(p, op["text"], op.get("track", True), author, date)
            if op.get("comment"):
                doc.add_comment(p, op["comment"], author, date)
        elif kind == "replace":
            _check_keys(op, {"op", "para", "find", "with", "track", "comment",
                             "every"}, i)
            for k in ("find", "with"):
                if k not in op:
                    raise DocxError("op #%d (replace): missing %r" % (i, k))
            doc.replace_text(p, op["find"], op["with"], op.get("track", True),
                             author, date, op.get("every", False))
            if op.get("comment"):
                doc.add_comment(p, op["comment"], author, date)
        elif kind in ("insert_after", "insert_before"):
            _check_keys(op, {"op", "para", "text", "style", "track", "comment"}, i)
            if "text" not in op:
                raise DocxError("op #%d (%s): missing 'text'" % (i, kind))
            fn = doc.insert_after if kind == "insert_after" else doc.insert_before
            newp = fn(p, op["text"], op.get("style"),
                      op.get("track", True), author, date)
            if op.get("comment"):
                doc.add_comment(newp, op["comment"], author, date)
        elif kind == "delete":
            _check_keys(op, {"op", "para", "hard", "comment"}, i)
            doc.delete_paragraph(p, op.get("hard", False), author, date)
            if op.get("comment"):
                doc.add_comment(p, op["comment"], author, date)
        elif kind == "comment":
            _check_keys(op, {"op", "para", "text"}, i)
            if "text" not in op:
                raise DocxError("op #%d (comment): missing 'text'" % i)
            doc.add_comment(p, op["text"], author, date)
        elif kind == "replace_all":
            _check_keys(op, {"op", "find", "with", "every", "track"}, i)
            for k in ("find", "with"):
                if k not in op:
                    raise DocxError("op #%d (replace_all): missing %r" % (i, k))
            doc.replace_all(op["find"], op["with"], op.get("track", True),
                            author, date, op.get("every", False))
        elif kind == "split":
            _check_keys(op, {"op", "para", "at", "track"}, i)
            if "at" not in op:
                raise DocxError("op #%d (split): missing 'at'" % i)
            doc.split_paragraph(p, op["at"], op.get("track", True), author, date)
        else:
            raise DocxError("op #%d: unknown op %r" % (i, kind))

    if not dry_run:
        doc.paras = [p for p in doc.body.iter(q("p"))]
    return len(items)


# --------------------------------------------------------------------------
# accept / reject
# --------------------------------------------------------------------------

def _para_marks(doc):
    out = {"del": [], "ins": []}
    for p in doc.body.iter(q("p")):
        rpr = p.find(q("pPr") + "/" + q("rPr"))
        if rpr is None:
            continue
        for tag in ("del", "ins"):
            if rpr.find(q(tag)) is not None:
                out[tag].append(p)
    return out


def accept_all(doc: Docx):
    # 1. run-level deletions out, insertions unwrapped
    for el in list(doc.root.iter(q("del"))):
        if el.getparent() is not None and el.getparent().tag == q("rPr"):
            continue
        el.getparent().remove(el)
    for el in list(doc.root.iter(q("ins"))):
        if el.getparent() is not None and el.getparent().tag == q("rPr"):
            continue
        _unwrap(el)
    # 2. deleted paragraph marks -> drop the paragraph (merge if it still has text)
    for p in _para_marks(doc)["del"]:
        if p.getparent() is None:
            continue
        if _text_of(p).strip():
            nxt = p.getnext()
            while nxt is not None and nxt.tag != q("p"):
                nxt = nxt.getnext()
            if nxt is not None:
                for r in list(p):
                    if r.tag == q("r"):
                        nxt.insert(0, r)
                p.getparent().remove(p)
                continue
        p.getparent().remove(p)
    # 3. strip leftover paragraph-mark markers
    for rpr in list(doc.root.iter(q("rPr"))):
        for tag in ("del", "ins"):
            el = rpr.find(q(tag))
            if el is not None:
                rpr.remove(el)
    doc.paras = [p for p in doc.body.iter(q("p"))]


def reject_all(doc: Docx):
    # 1. inserted paragraphs disappear
    for p in _para_marks(doc)["ins"]:
        if p.getparent() is not None:
            p.getparent().remove(p)
    # 2. inserted runs disappear
    for el in list(doc.root.iter(q("ins"))):
        if el.getparent() is not None and el.getparent().tag == q("rPr"):
            continue
        el.getparent().remove(el)
    # 3. deletions come back
    for el in list(doc.root.iter(q("del"))):
        if el.getparent() is not None and el.getparent().tag == q("rPr"):
            continue
        for dt in el.iter(q("delText")):
            dt.tag = q("t")
        _unwrap(el)
    for rpr in list(doc.root.iter(q("rPr"))):
        for tag in ("del", "ins"):
            el = rpr.find(q(tag))
            if el is not None:
                rpr.remove(el)
    doc.paras = [p for p in doc.body.iter(q("p"))]


# --------------------------------------------------------------------------
# diff two documents
# --------------------------------------------------------------------------

def build_tracked(old_path, new_path, out_path,
                  author=DEFAULT_AUTHOR, date=DEFAULT_DATE):
    old = Docx(old_path)
    new = Docx(new_path)
    ot = [_text_of(p) for p in old.paras]
    nt = [_text_of(p) for p in new.paras]
    sm = difflib.SequenceMatcher(None, ot, nt, autojunk=False)

    def deleted_copy(src):
        d = copy.deepcopy(src)
        for r in _direct_runs(d):
            d.replace(r, _make_mark(
                "del", "".join(t.text or "" for t in r.iter(q("t"))),
                old.nid(), author, date, r.find(q("rPr"))))
        _mark_mark(d, "del", old, author, date)
        return d

    def inserted_copy(src):
        n = copy.deepcopy(src)
        for r in _direct_runs(n):
            n.replace(r, _make_mark(
                "ins", "".join(t.text or "" for t in r.iter(q("t"))),
                old.nid(), author, date, r.find(q("rPr"))))
        _mark_mark(n, "ins", old, author, date)
        return n

    # map each diff block to output paragraphs
    out_map = {}          # old paragraph element -> list of new paragraph elements
    inserts = {}          # anchor old paragraph (or None for start) -> list
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                out_map[old.paras[i1 + k]] = [copy.deepcopy(old.paras[i1 + k])]
            continue
        if tag == "replace" and (i2 - i1) == (j2 - j1):
            for k in range(i2 - i1):
                op, np_ = old.paras[i1 + k], new.paras[j1 + k]
                tmp = copy.deepcopy(op)
                old.set_paragraph(tmp, _text_of(np_), True, author, date)
                out_map[op] = [tmp]
            continue
        if tag == "replace":
            # unequal block (e.g. one paragraph split into two): delete the old
            # paragraphs in place, then insert the new ones after the last one.
            for op in old.paras[i1:i2]:
                out_map[op] = [deleted_copy(op)]
            anchor = old.paras[i2 - 1] if i2 > i1 else (None if i1 == 0 else old.paras[i1 - 1])
            inserts.setdefault(anchor, []).extend(
                inserted_copy(p) for p in new.paras[j1:j2])
            continue
        if tag == "delete":
            for op in old.paras[i1:i2]:
                out_map[op] = [deleted_copy(op)]
            continue
        if tag == "insert":
            prev = None if i1 == 0 else old.paras[i1 - 1]
            inserts.setdefault(prev, []).extend(
                inserted_copy(p) for p in new.paras[j1:j2])
            continue

    # Insertions before the very first paragraph (do this before replacements
    # detach that paragraph from its parent).
    if inserts.get(None) and old.paras:
        first = old.paras[0]
        parent = first.getparent()
        if parent is not None:
            idx = list(parent).index(first)
            for k, el in enumerate(inserts[None]):
                parent.insert(idx + k, el)

    # Replace paragraphs in place so table cells and nested containers survive.
    for p in list(old.paras):
        outs = out_map.get(p)
        if outs is None:
            continue
        parent = p.getparent()
        if parent is None:
            continue
        idx = list(parent).index(p)
        parent.remove(p)
        for k, el in enumerate(outs):
            parent.insert(idx + k, el)
        extra = inserts.get(p) or []
        for k, el in enumerate(extra):
            parent.insert(idx + len(outs) + k, el)
    old.save(out_path)
    return out_path


def _mark_mark(p, tag, doc, author, date):
    ppr = p.find(q("pPr"))
    if ppr is None:
        ppr = etree.Element(q("pPr"))
        p.insert(0, ppr)
    rpr = ppr.find(q("rPr"))
    if rpr is None:
        rpr = etree.SubElement(ppr, q("rPr"))
    el = etree.SubElement(rpr, q(tag))
    el.set(q("id"), str(doc.nid()))
    el.set(q("author"), author)
    el.set(q("date"), date)


# --------------------------------------------------------------------------
# compare
# --------------------------------------------------------------------------

def compare_docs(a_path, b_path, include_deleted=False):
    """Compare paragraph text between two documents.

    Returns {"equal": bool, "paragraphs_a": n, "paragraphs_b": n,
             "diffs": [{"index", "a", "b"}]} (at most 50 diffs).
    """
    a = Docx(a_path)
    b = Docx(b_path)
    ta = [_text_of(p, include_deleted) for p in a.paras]
    tb = [_text_of(p, include_deleted) for p in b.paras]
    diffs = []
    for i in range(max(len(ta), len(tb))):
        x = ta[i] if i < len(ta) else None
        y = tb[i] if i < len(tb) else None
        if x != y:
            diffs.append({"index": i, "a": x, "b": y})
    return {"equal": not diffs, "paragraphs_a": len(ta),
            "paragraphs_b": len(tb), "diffs": diffs[:50]}


# --------------------------------------------------------------------------
# verification
# --------------------------------------------------------------------------

def verify(path):
    d = Docx(path)
    xml = etree.tostring(d.root)
    counts = {
        "paragraphs": len(d.paras),
        "ins": xml.count(b"<w:ins "),
        "del": xml.count(b"<w:del "),
        "delText": xml.count(b"<w:delText"),
        "comments": 0 if d.comments is None else len(d.comments.findall(q("comment"))),
    }
    with zipfile.ZipFile(path) as z:
        counts["zip_ok"] = z.testzip() is None
    return counts
