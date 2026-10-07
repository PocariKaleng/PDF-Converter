"""Real Word conversion check; requires Windows, Word, and Bell MT."""
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from converter.engine import convert
from tests.helpers import docx_bytes


if __name__ == '__main__':
    for font in ('Bell MT', 'LM Roman 10'):
        for mode in ('original', font):
            pdf, report = convert(docx_bytes(font=font), mode)
            assert pdf.startswith(b'%PDF-')
            assert report['verified_families'] == [font]
            print(json.dumps({'mode': mode, 'font': font, **report}))
    print('All four real Word conversions passed.')
