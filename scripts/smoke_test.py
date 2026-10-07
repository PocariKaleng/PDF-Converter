"""Real Word conversion check; requires Windows, Word, and Bell MT."""
import json
from io import BytesIO
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from converter.engine import convert
from converter.fonts import inventory
from pypdf import PdfReader
from tests.helpers import styled_docx_bytes


if __name__ == '__main__':
    for font in ('Bell MT', 'LM Roman 10'):
        variants = ('regular', 'bold', 'italic', 'bolditalic') if font == 'LM Roman 10' else ('regular', 'bold', 'italic')
        for mode in ('original', font):
            source_font = font if mode == 'original' else 'Arial'
            pdf, report = convert(styled_docx_bytes(source_font, variants), mode)
            assert pdf.startswith(b'%PDF-')
            assert report['verified_families'] == [font]
            used_fonts = {}

            def record_font(text, cm, tm, font_dict, size):
                if font_dict:
                    name = str(font_dict.get('/BaseFont', '')).lstrip('/').split('+')[-1]
                    for line in text.strip().splitlines():
                        if line.startswith('Contoh '):
                            used_fonts[line.strip().removeprefix('Contoh ')] = name

            for page in PdfReader(BytesIO(pdf)).pages:
                page.extract_text(visitor_text=record_font)
            for variant in variants:
                expected = inventory()[font]['variants'][variant]['postscript']
                assert used_fonts.get(variant) == expected, (mode, variant, used_fonts)
                entry = next(item for item in report['fonts'] if item['name'] == expected)
                assert entry['embedded'], entry
                if font == 'LM Roman 10':
                    assert entry['version'] == '1.106', entry
            print(json.dumps({'mode': mode, 'font': font, **report}))
    print('All four real Word conversions passed, including bold and italic font mappings.')
