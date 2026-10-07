from io import BytesIO
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

from lxml import etree
from pypdf import PdfWriter
from pypdf.generic import DictionaryObject, NameObject, DecodedStreamObject

from converter.docx import prepare_docx, W
from converter.errors import ConversionError
from converter.fonts import family_key, inventory
from converter.pdf import inspect_pdf, embed_supported_fonts
from tests.helpers import docx_bytes


class DocxTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.target = Path(self.directory.name) / "input.docx"

    def tearDown(self):
        self.directory.cleanup()

    def test_preserve_original_bytes_and_style_inheritance(self):
        original = docx_bytes()
        self.assertEqual(prepare_docx(original, self.target), {"Bell MT"})
        self.assertEqual(self.target.read_bytes(), original)

    def test_direct_run_overrides_style(self):
        run = '<w:rPr><w:rFonts w:ascii="LM Roman 10" w:hAnsi="LM Roman 10"/></w:rPr>'
        self.assertEqual(prepare_docx(docx_bytes(run=run), self.target), {"LM Roman 10"})

    def test_theme_overrides_explicit_font_from_parent(self):
        run = '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/></w:rPr>'
        theme = '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:minorFont><a:latin typeface="LM Roman 10"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>'
        self.assertEqual(prepare_docx(docx_bytes(run=run, extra={"word/theme/theme1.xml": theme}), self.target), {"LM Roman 10"})

    def test_font_override_preserves_bold_and_sizes(self):
        run = '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr>'
        self.assertEqual(prepare_docx(docx_bytes(run=run), self.target, "LM Roman 10"), {"LM Roman 10"})
        with ZipFile(self.target) as archive:
            root = etree.fromstring(archive.read("word/document.xml"))
            self.assertEqual(root.find('.//' + W + 'rFonts').get(W + 'ascii'), 'LM Roman 10')
            self.assertIsNotNone(root.find('.//' + W + 'b'))
            self.assertEqual(root.find('.//' + W + 'sz').get(W + 'val'), '32')

    def test_alias_is_canonicalized(self):
        self.assertEqual(prepare_docx(docx_bytes(font="Latin Modern Roman 10"), self.target), {"LM Roman 10"})
        with ZipFile(self.target) as archive:
            self.assertIn(b'LM Roman 10', archive.read('word/styles.xml'))

    def test_namespace_prefixes_used_in_ignorable_survive_rewrite(self):
        data = docx_bytes()
        with ZipFile(BytesIO(data)) as archive:
            doc = archive.read('word/document.xml').decode().replace('<w:document ', '<w:document xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14" ')
        prepare_docx(docx_bytes(extra={'word/document.xml': doc}), self.target, 'LM Roman 10')
        with ZipFile(self.target) as archive:
            root = etree.fromstring(archive.read('word/document.xml'))
            self.assertIn('w14', root.nsmap)

    def test_headers_are_checked_but_unused_comments_are_ignored(self):
        xml = f'<w:root xmlns:w="{W[1:-1]}"><w:p><w:r><w:rPr><w:rFonts w:ascii="LM Roman 10"/></w:rPr><w:t>Header</w:t></w:r></w:p></w:root>'
        self.assertEqual(prepare_docx(docx_bytes(font='Arial', extra={'word/header1.xml': xml}), self.target), {'LM Roman 10'})
        self.assertEqual(prepare_docx(docx_bytes(font='Arial', extra={'word/comments.xml': xml}), self.target), set())

    def test_bad_packages_are_rejected(self):
        for content in (b'fake docx', b'', docx_bytes(extra={'word/vbaProject.bin': b'x'}), docx_bytes(extra={'word/embeddings/data.bin': b'x'}), docx_bytes(extra={'word/bad.xml': '<!DOCTYPE x><x/>'})):
            with self.subTest(content=content[:20]), self.assertRaises(ConversionError):
                prepare_docx(content, self.target)

    def test_external_templates_are_rejected(self):
        rel = '<Relationships><Relationship TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" Target="https://example.com/template.dotm"/></Relationships>'
        with self.assertRaises(ConversionError):
            prepare_docx(docx_bytes(extra={'word/_rels/settings.xml.rels': rel}), self.target)

    def test_unknown_font_option_is_rejected(self):
        with self.assertRaises(ConversionError):
            prepare_docx(docx_bytes(), self.target, 'Wrong font')


class FontTests(unittest.TestCase):
    def test_postscript_and_family_aliases(self):
        self.assertEqual(family_key('BellMT-Bold'), 'Bell MT')
        self.assertEqual(family_key('LMRoman10-BoldItalic'), 'LM Roman 10')
        self.assertIsNone(family_key('LMRoman12-Regular'))

    def test_bundled_lm_has_all_four_embeddable_variants(self):
        family = inventory()['LM Roman 10']
        self.assertTrue(family['available'])
        self.assertEqual(set(family['variants']), {'regular', 'bold', 'italic', 'bolditalic'})
        self.assertTrue(all(v['embeddable'] for v in family['variants'].values()))
        self.assertEqual({v['version'] for v in family['variants'].values()}, {'1.106'})


class PdfTests(unittest.TestCase):
    def create_pdf(self, path, name='BellMT', embedded=True, subtype='/TrueType', subset=True):
        writer = PdfWriter()
        page = writer.add_blank_page(595, 842)
        descriptor = DictionaryObject({NameObject('/Type'): NameObject('/FontDescriptor')})
        if embedded:
            stream = DecodedStreamObject()
            stream.set_data(b'test font stream')
            descriptor[NameObject('/FontFile2')] = writer._add_object(stream)
        font = DictionaryObject({NameObject('/Type'): NameObject('/Font'), NameObject('/Subtype'): NameObject(subtype), NameObject('/BaseFont'): NameObject('/' + ('ABCDEF+' if subset else '') + name), NameObject('/FontDescriptor'): writer._add_object(descriptor)})
        page[NameObject('/Resources')] = DictionaryObject({NameObject('/Font'): DictionaryObject({NameObject('/F1'): writer._add_object(font)})})
        writer.write(path)

    def test_missing_substituted_and_nonembedded_fonts_are_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'test.pdf'
            for name, embedded in [('Arial', True), ('BellMT', False)]:
                self.create_pdf(path, name, embedded)
                with self.assertRaises(ConversionError):
                    inspect_pdf(path, {'Bell MT'})
            self.create_pdf(path)
            report = inspect_pdf(path, {'Bell MT'})
            self.assertEqual(report['verified_families'], ['Bell MT'])
            self.assertTrue(report['fonts'][0]['embedded'])

    def test_original_lm_cff_is_embedded_with_version_1106(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'test.pdf'
            self.create_pdf(path, 'LMRoman10-Regular', False, '/Type1', False)
            embed_supported_fonts(path)
            report = inspect_pdf(path, {'LM Roman 10'})
            self.assertTrue(report['fonts'][0]['embedded'])
            self.assertEqual(report['fonts'][0]['version'], '1.106')

    def test_unknown_or_subsetted_fonts_are_never_patched(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'test.pdf'
            for name, subset in [('LMRoman10-Regular', True), ('UnrelatedFont', False)]:
                self.create_pdf(path, name, False, '/Type1', subset)
                original = path.read_bytes()
                embed_supported_fonts(path)
                self.assertEqual(path.read_bytes(), original)
