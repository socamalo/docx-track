# -*- coding: utf-8 -*-
"""Unit tests for docxtrack core (run: python -m unittest discover tests)."""
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "src"))

from docxtrack import core  # noqa: E402

FIX = os.path.join(HERE, "fixtures")


def texts(path):
    d = core.Docx(path)
    return [core._text_of(p) for p in d.paras]


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.base = os.path.join(self.tmp, "base.docx")
        self.rev = os.path.join(self.tmp, "revised.docx")
        shutil.copyfile(os.path.join(FIX, "base.docx"), self.base)
        shutil.copyfile(os.path.join(FIX, "revised.docx"), self.rev)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _roundtrip(self, tracked):
        acc = os.path.join(self.tmp, "acc.docx")
        rej = os.path.join(self.tmp, "rej.docx")
        d = core.Docx(tracked); core.accept_all(d); d.save(acc)
        d = core.Docx(tracked); core.reject_all(d); d.save(rej)
        return texts(acc), texts(rej)


class TestSet(Base):
    def test_set_tracked_roundtrip(self):
        new = "The second paragraph mentions the port boundary, clearly."
        d = core.Docx(self.base)
        d.set_paragraph(d.resolve(2), new)
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        acc, rej = self._roundtrip(t)
        self.assertEqual(acc[2], new)
        self.assertEqual(rej[2], texts(self.base)[2])
        self.assertGreater(core.verify(t)["del"], 0)
        self.assertGreater(core.verify(t)["ins"], 0)

    def test_set_no_track(self):
        new = "Direct replacement."
        d = core.Docx(self.base)
        d.set_paragraph(d.resolve(2), new, track=False)
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        self.assertEqual(texts(t)[2], new)
        self.assertEqual(core.verify(t)["del"], 0)

    def test_repeated_edits_compose(self):
        d = core.Docx(self.base)
        p = d.resolve(2)
        d.replace_text(p, "fence", "boundary")
        d.set_paragraph(p, "Final wording for the paragraph.")
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        acc, rej = self._roundtrip(t)
        self.assertEqual(acc[2], "Final wording for the paragraph.")
        self.assertEqual(rej, texts(self.base))

    def test_mixed_formatting_kept(self):
        d = core.Docx(self.base)
        d.set_paragraph(d.resolve(1), "Bold lead. This is the first body paragraph, revised.")
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        acc = os.path.join(self.tmp, "acc.docx")
        d = core.Docx(t); core.accept_all(d); d.save(acc)
        root = core.Docx(acc).root
        p = list(core.Docx(acc).paras)[1]
        runs = p.findall(core.q("r"))
        self.assertTrue(runs[0].find(core.q("rPr") + "/" + core.q("b")) is not None)
        self.assertIsNone(runs[-1].find(core.q("rPr") + "/" + core.q("b")))


class TestReplaceInsertDelete(Base):
    def test_replace(self):
        d = core.Docx(self.base)
        d.replace_text(d.resolve("a fence"), "fence", "boundary")
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        acc, rej = self._roundtrip(t)
        self.assertIn("boundary", acc[2])
        self.assertEqual(rej[2], texts(self.base)[2])

    def test_insert_tracked_roundtrip(self):
        d = core.Docx(self.base)
        d.insert_after(d.resolve(3), "Inserted paragraph.")
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        acc, rej = self._roundtrip(t)
        self.assertEqual(len(acc), len(texts(self.base)) + 1)
        self.assertIn("Inserted paragraph.", acc)
        self.assertEqual(rej, texts(self.base))

    def test_delete_tracked_roundtrip(self):
        d = core.Docx(self.base)
        d.delete_paragraph(d.resolve(4))
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        acc, rej = self._roundtrip(t)
        self.assertEqual(len(acc), len(texts(self.base)) - 1)
        self.assertNotIn("bullet one", acc)
        self.assertEqual(rej, texts(self.base))

    def test_delete_hard(self):
        d = core.Docx(self.base)
        d.delete_paragraph(d.resolve(4), hard=True)
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        self.assertEqual(len(texts(t)), len(texts(self.base)) - 1)


class TestDiff(Base):
    def test_diff_roundtrip(self):
        out = os.path.join(self.tmp, "d.docx")
        core.build_tracked(self.base, self.rev, out)
        acc, rej = self._roundtrip(out)
        self.assertEqual(acc, texts(self.rev))
        self.assertEqual(rej, texts(self.base))

    def test_diff_split_paragraph(self):
        d = core.Docx(self.base)
        d.split_paragraph(d.resolve(2), "which is")
        new = os.path.join(self.tmp, "new.docx"); d.save(new)
        out = os.path.join(self.tmp, "d.docx")
        core.build_tracked(self.base, new, out)
        acc, rej = self._roundtrip(out)
        self.assertEqual(acc, texts(new))
        self.assertEqual(rej, texts(self.base))

    def test_diff_prepend_paragraph(self):
        new = os.path.join(FIX, "prepend.docx")
        out = os.path.join(self.tmp, "d.docx")
        core.build_tracked(self.base, new, out)
        acc, rej = self._roundtrip(out)
        self.assertEqual(acc, texts(new))
        self.assertEqual(rej, texts(self.base))

    def test_diff_has_revisions(self):
        out = os.path.join(self.tmp, "d.docx")
        core.build_tracked(self.base, self.rev, out)
        info = core.verify(out)
        self.assertGreater(info["ins"], 0)
        self.assertGreater(info["del"], 0)
        self.assertTrue(info["zip_ok"])


class TestApply(Base):
    def _spec(self):
        return {"author": "Reviewer", "date": "2026-11-11T00:00:00Z", "ops": [
            {"op": "replace", "para": "mentions a fence", "find": "fence",
             "with": "port boundary"},
            {"op": "set", "para": "Bold lead.", "text": "Bold lead. Revised."},
            {"op": "insert_after", "para": 3, "text": "Batch insert.", "style": "BodyText"},
            {"op": "delete", "para": "bullet two"},
            {"op": "comment", "para": 0, "text": "Check title."},
        ]}

    def test_apply_batch(self):
        d = core.Docx(self.base)
        core.apply_ops(d, self._spec())
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        acc, rej = self._roundtrip(t)
        self.assertIn("port boundary", acc[2])
        self.assertIn("Batch insert.", acc)
        self.assertNotIn("bullet two", acc)
        self.assertEqual(rej, texts(self.base))
        self.assertEqual(core.verify(t)["comments"], 1)

    def test_ambiguous_anchor_rejected(self):
        d = core.Docx(self.base)
        spec = {"ops": [{"op": "set", "para": "paragraph", "text": "x"}]}
        with self.assertRaises(core.DocxError):
            core.apply_ops(d, spec)

    def test_bad_index_rejected_without_write(self):
        d = core.Docx(self.base)
        spec = {"ops": [{"op": "set", "para": 99, "text": "x"}]}
        with self.assertRaises(core.DocxError):
            core.apply_ops(d, spec)


class TestComments(Base):
    def test_comment_part_created(self):
        d = core.Docx(self.base)
        d.add_comment(d.resolve(0), "A note.")
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        info = core.verify(t)
        self.assertEqual(info["comments"], 1)
        self.assertTrue(info["zip_ok"])
        import zipfile
        with zipfile.ZipFile(t) as z:
            self.assertIn("word/comments.xml", z.namelist())
            ct = z.read("[Content_Types].xml").decode("utf8")
            self.assertIn("comments+xml", ct)


class TestReplaceAll(Base):
    def test_replace_all_counts_and_roundtrip(self):
        d = core.Docx(self.base)
        n = d.replace_all("paragraph", "PARA")
        self.assertEqual(n, 3)  # three paragraphs contain the word
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        acc, rej = self._roundtrip(t)
        self.assertEqual(sum(1 for x in acc if "PARA" in x), 3)
        self.assertEqual(rej, texts(self.base))


class TestInsertBefore(Base):
    def test_insert_before_roundtrip(self):
        d = core.Docx(self.base)
        d.insert_before(d.resolve(3), "Before third.")
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        acc, rej = self._roundtrip(t)
        self.assertEqual(len(acc), len(texts(self.base)) + 1)
        self.assertEqual(acc.index("Before third."), 3)
        self.assertEqual(rej, texts(self.base))


class TestReplaceEvery(Base):
    def test_replace_every_occurrence(self):
        d = core.Docx(self.base)
        p = d.resolve(2)
        d.set_paragraph(p, "fence fence fence", track=False)
        d.replace_text(p, "fence", "wall", every=True)
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        acc, _ = self._roundtrip(t)
        self.assertEqual(acc[2], "wall wall wall")


class TestSplit(Base):
    def test_split_roundtrip(self):
        d = core.Docx(self.base)
        before, after = d.split_paragraph(d.resolve(2), "which is")
        self.assertTrue(before.endswith("yard,"))
        self.assertTrue(after.startswith("which is"))
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        acc, rej = self._roundtrip(t)
        self.assertEqual(len(acc), len(texts(self.base)) + 1)
        self.assertEqual(rej, texts(self.base))


class TestMerge(Base):
    def test_merge_next_roundtrip(self):
        d = core.Docx(self.base)
        n = d.merge_next(d.resolve(1), count=1)
        self.assertEqual(n, 2)
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        acc, rej = self._roundtrip(t)
        self.assertEqual(len(acc), len(texts(self.base)) - 1)
        self.assertIn("kept short. The second paragraph", acc[1])
        self.assertEqual(rej, texts(self.base))


class TestCompare(Base):
    def test_compare_equal_and_diff(self):
        self.assertTrue(core.compare_docs(self.base, self.base)["equal"])
        info = core.compare_docs(self.base, self.rev)
        self.assertFalse(info["equal"])
        self.assertGreater(len(info["diffs"]), 0)


class TestTables(Base):
    def setUp(self):
        super().setUp()
        self.table = os.path.join(self.tmp, "table.docx")
        shutil.copyfile(os.path.join(FIX, "table.docx"), self.table)

    def test_table_cell_edit_roundtrip(self):
        d = core.Docx(self.table)
        d.set_paragraph(d.resolve(4), "20")
        t = os.path.join(self.tmp, "t.docx"); d.save(t)
        acc, rej = self._roundtrip(t)
        self.assertEqual(acc[4], "20")
        self.assertEqual(rej[4], "19")

    def test_diff_table_roundtrip(self):
        d = core.Docx(self.table)
        d.set_paragraph(d.resolve(4), "20")
        new = os.path.join(self.tmp, "new.docx"); d.save(new)
        out = os.path.join(self.tmp, "d.docx")
        core.build_tracked(self.table, new, out)
        acc, rej = self._roundtrip(out)
        self.assertEqual(acc, texts(new))
        self.assertEqual(rej, texts(self.table))


class TestCliSmoke(unittest.TestCase):
    def setUp(self):
        self.exe = shutil.which("docxtrack") or os.path.expanduser(
            r"~\.local\bin\docxtrack.exe")
        if not os.path.exists(self.exe):
            self.skipTest("docxtrack not installed")
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def run_cli(self, *args):
        import subprocess
        return subprocess.run([self.exe] + list(args), capture_output=True,
                              text=True, encoding="utf-8")

    def test_cli_list_and_verify(self):
        out = self.run_cli("list", os.path.join(FIX, "base.docx"))
        self.assertEqual(out.returncode, 0)
        self.assertIn("#000", out.stdout)
        out = self.run_cli("verify", os.path.join(FIX, "base.docx"), "--json")
        self.assertEqual(out.returncode, 0)
        self.assertIn('"zip_ok": true', out.stdout)

    def test_cli_compare_replace_all_delete_range(self):
        a = os.path.join(self.tmp, "a.docx")
        b = os.path.join(self.tmp, "b.docx")
        shutil.copyfile(os.path.join(FIX, "base.docx"), a)
        self.assertEqual(self.run_cli("compare", a, a).returncode, 0)
        out = self.run_cli("replace-all", a, "--find", "paragraph",
                           "--with", "PARA", "-o", b, "--json")
        self.assertEqual(out.returncode, 0, out.stderr)
        self.assertIn('"paragraphs": 3', out.stdout)
        c = os.path.join(self.tmp, "c.docx")
        out = self.run_cli("delete", a, "4", "--to", "5", "-o", c, "--json")
        self.assertEqual(out.returncode, 0, out.stderr)
        d = self.run_cli("accept", c, "-o", os.path.join(self.tmp, "ca.docx"))
        self.assertEqual(d.returncode, 0)
        rows = core.Docx(os.path.join(self.tmp, "ca.docx")).paras
        self.assertEqual(len(rows), len(core.Docx(a).paras) - 2)


if __name__ == "__main__":
    unittest.main()
