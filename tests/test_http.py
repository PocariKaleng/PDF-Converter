import http.client
import json
import threading
import unittest
from unittest.mock import patch
from urllib.parse import unquote

from app import make_server


class HttpTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = make_server(0)
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join()

    def request(self, method, path, body=None, headers=None):
        connection = http.client.HTTPConnection('127.0.0.1', self.server.server_port, timeout=10)
        connection.request(method, path, body, headers or {})
        response = connection.getresponse()
        result = (response.status, response.read(), dict(response.getheaders()))
        connection.close()
        return result

    def test_static_path_traversal_is_rejected(self):
        self.assertEqual(self.request('GET', '/../app.py')[0], 404)

    def test_host_rebinding_is_rejected(self):
        self.assertEqual(self.request('GET', '/', headers={'Host': 'attacker.example'})[0], 403)

    def test_conversion_requires_token_and_same_origin(self):
        self.assertEqual(self.request('POST', '/api/convert', b'x')[0], 403)
        self.assertEqual(self.request('POST', '/api/convert', b'x', {'X-Converter-Token': self.server.token, 'Origin': 'https://example.com'})[0], 403)

    def test_wrong_extension_is_rejected(self):
        self.assertEqual(self.request('POST', '/api/convert', b'x', {'X-Converter-Token': self.server.token, 'X-Filename': 'document.txt'})[0], 400)

    def test_binary_pdf_and_verification_report_are_returned(self):
        report = {'pages': 2, 'fonts': [], 'verified_families': ['Bell MT']}
        with patch('app.convert', return_value=(b'%PDF-test', report)) as conversion:
            status, body, headers = self.request('POST', '/api/convert', b'input', {'X-Converter-Token': self.server.token, 'X-Filename': 'laporan.docx'})
            self.assertEqual(status, 200)
            self.assertEqual(body, b'%PDF-test')
            self.assertEqual(json.loads(unquote(headers['X-Conversion-Report'])), report)
            self.assertIn('laporan.pdf', headers['Content-Disposition'])
            download_status, download_body, download_headers = self.request('GET', headers['X-Download-Url'])
            self.assertEqual(download_status, 200)
            self.assertEqual(download_body, body)
            self.assertIn('attachment', download_headers['Content-Disposition'])
            self.assertEqual(self.request('GET', '/api/download/invalid')[0], 404)
            conversion.assert_called_once_with(b'input', 'original')
