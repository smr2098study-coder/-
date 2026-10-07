"""Structural regression tests for native Mode B citations."""
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile, ZIP_DEFLATED

from lxml import etree as E

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / 'plugins/shimen-format/skills/shimen-word-format/scripts/citation_audit.py'
spec = importlib.util.spec_from_file_location('citation_audit', SCRIPT)
audit = importlib.util.module_from_spec(spec)
spec.loader.exec_module(audit)


def w(name):
    return '{' + audit.W + '}' + name


def run(parent, text=None, superscript=False, baseline=False):
    node = E.SubElement(parent, w('r'))
    if superscript or baseline:
        rpr = E.SubElement(node, w('rPr'))
        align = E.SubElement(rpr, w('vertAlign'))
        align.set(w('val'), 'superscript' if superscript else 'baseline')
    if text is not None:
        t = E.SubElement(node, w('t'))
        t.text = text
    return node


def valid_fixture(path):
    document = E.Element(w('document'), nsmap={'w': audit.W})
    body = E.SubElement(document, w('body'))
    p = E.SubElement(body, w('p'))
    run(p, '[', superscript=True)
    start = E.SubElement(p, w('bookmarkStart'))
    start.set(w('id'), '7')
    start.set(w('name'), '_RefCite_A7F32C')
    ref_run = run(p, superscript=True)
    ref = E.SubElement(ref_run, w('endnoteReference'))
    ref.set(w('id'), '2')
    end = E.SubElement(p, w('bookmarkEnd'))
    end.set(w('id'), '7')
    run(p, ']', superscript=True)
    p2 = E.SubElement(body, w('p'))
    run(p2, '[', superscript=True)
    field_runs = [
        ('fldChar', {'fldCharType': 'begin'}, None),
        ('instrText', {}, ' NOTEREF _RefCite_A7F32C \\h \\* MERGEFORMAT '),
        ('fldChar', {'fldCharType': 'separate'}, None),
        ('t', {}, '1'),
        ('fldChar', {'fldCharType': 'end'}, None),
    ]
    for kind, attrs, text in field_runs:
        r = run(p2, superscript=True)
        child = E.SubElement(r, w(kind))
        for key, value in attrs.items():
            child.set(w(key), value)
        child.text = text
    run(p2, ']', superscript=True)

    endnotes = E.Element(w('endnotes'), nsmap={'w': audit.W})
    for note_id, note_type in [('-1', 'separator'), ('0', 'continuationSeparator')]:
        note = E.SubElement(endnotes, w('endnote'))
        note.set(w('id'), note_id)
        note.set(w('type'), note_type)
        E.SubElement(note, w('p'))
    note = E.SubElement(endnotes, w('endnote'))
    note.set(w('id'), '2')
    np = E.SubElement(note, w('p'))
    ppr = E.SubElement(np, w('pPr'))
    ind = E.SubElement(ppr, w('ind'))
    ind.set(w('left'), '420')
    ind.set(w('hanging'), '420')
    run(np, '[', baseline=True)
    rr = run(np, baseline=True)
    E.SubElement(rr, w('endnoteRef'))
    run(np, ']', baseline=True)
    tab_run = run(np)
    E.SubElement(tab_run, w('tab'))
    run(np, 'Reference A')

    settings = E.Element(w('settings'), nsmap={'w': audit.W})
    ep = E.SubElement(settings, w('endnotePr'))
    for name, value in [('pos', 'sectEnd'), ('numFmt', 'decimal'), ('numStart', '1'), ('numRestart', 'continuous')]:
        node = E.SubElement(ep, w(name))
        node.set(w('val'), value)

    types = E.Element('{%s}Types' % audit.CT, nsmap={None: audit.CT})
    E.SubElement(types, '{%s}Override' % audit.CT, PartName='/word/endnotes.xml',
                 ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml')
    rels = E.Element('{%s}Relationships' % audit.PR, nsmap={None: audit.PR})
    E.SubElement(rels, '{%s}Relationship' % audit.PR, Id='rId1', Target='endnotes.xml',
                 Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes')

    with ZipFile(path, 'w', ZIP_DEFLATED) as z:
        z.writestr('word/document.xml', E.tostring(document))
        z.writestr('word/endnotes.xml', E.tostring(endnotes))
        z.writestr('word/settings.xml', E.tostring(settings))
        z.writestr('[Content_Types].xml', E.tostring(types))
        z.writestr('word/_rels/document.xml.rels', E.tostring(rels))


class CitationAuditTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_native_mode_b_requires_word_validation_for_completion(self):
        path = self.path / 'valid.docx'
        valid_fixture(path)
        static = audit.audit(path, 'B')
        self.assertEqual(static['status'], 'pending_word_validation', static)
        self.assertTrue(static['navigation_ok'])
        self.assertTrue(static['dynamic_field_structure_ok'])
        self.assertIsNone(static['dynamic_field_ok'])
        self.assertEqual(static['counts']['noteref_field_count'], 1)
        self.assertEqual(static['counts']['broken_noteref_count'], 0)
        verified = audit.audit(path, 'B', {
            'word_field_update_validation': 'passed',
            'dynamic_insert_test_passed': True,
            'dynamic_delete_test_passed': True,
            'save_close_reopen_passed': True,
            'broken_noteref_count': 0,
        })
        self.assertEqual(verified['status'], 'passed', verified)
        self.assertTrue(verified['dynamic_field_ok'])

    def test_static_hyperlink_and_missing_endnotes_fail(self):
        path = self.path / 'fake.docx'
        document = E.Element(w('document'), nsmap={'w': audit.W})
        body = E.SubElement(document, w('body'))
        p = E.SubElement(body, w('p'))
        link = E.SubElement(p, w('hyperlink'))
        link.set(w('anchor'), 'ref_1')
        run(link, '[1]', superscript=True)
        with ZipFile(path, 'w', ZIP_DEFLATED) as z:
            z.writestr('word/document.xml', E.tostring(document))
            z.writestr('[Content_Types].xml', '<Types xmlns="%s"/>' % audit.CT)
            z.writestr('word/_rels/document.xml.rels', '<Relationships xmlns="%s"/>' % audit.PR)
        result = audit.audit(path, 'B')
        self.assertEqual(result['status'], 'failed')
        self.assertEqual(result['counts']['citation_hyperlink_count'], 1)
        self.assertFalse(result['dynamic_field_structure_ok'])

    def test_broken_and_display_number_bookmark_are_rejected(self):
        path = self.path / 'broken.docx'
        valid_fixture(path)
        with ZipFile(path) as z:
            parts = {name: z.read(name) for name in z.namelist()}
        root = E.fromstring(parts['word/document.xml'])
        root.find('.//w:bookmarkStart', audit.NS).set(w('name'), 'ref_1')
        parts['word/document.xml'] = E.tostring(root)
        with ZipFile(path, 'w', ZIP_DEFLATED) as z:
            for name, data in parts.items():
                z.writestr(name, data)
        result = audit.audit(path, 'B')
        self.assertEqual(result['status'], 'failed')
        self.assertGreater(result['counts']['broken_noteref_count'], 0)


if __name__ == '__main__':
    unittest.main()
