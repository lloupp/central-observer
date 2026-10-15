"""Testes com dados sinteticos, sem bind de LAN nem TLS real."""
import base64
import importlib
import json
import os
import sys
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app
import server_lan as lan


class LANTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp=tempfile.TemporaryDirectory()
        cls.saved_db=app.DB_PATH
        app.DB_PATH=Path(cls.tmp.name)/'synthetic.db'
        app.init_db()
        lan.USERNAME='central'
        lan.PASSWORD='segredo-de-teste-com-32-caracteres'
        lan.TLS=False
        lan.SCHEME='http'
        lan.PUBLIC_ORIGIN=''
        cls.server=lan.LanServer(('127.0.0.1',0),lan.LanHandler)
        cls.port=cls.server.server_port
        cls.host='http://127.0.0.1:'+str(cls.port)
        cls.thread=threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        app.DB_PATH=cls.saved_db
        cls.tmp.cleanup()

    def request(self, path, method='GET', headers=None, data=None):
        req=urllib.request.Request(self.host+path, method=method, data=data, headers=headers or {})
        try:
            with urllib.request.urlopen(req, timeout=5) as res:
                return res.status, res.read()
        except urllib.error.HTTPError as ex:
            return ex.code, ex.read()

    def auth(self, password=None):
        credentials='central:'+(password if password is not None else lan.PASSWORD)
        return {'Authorization':'Basic '+base64.b64encode(credentials.encode()).decode()}

    def test_access_blocked_without_password(self):
        for path in ('/', '/static/app.js', '/api/estagios','/api/exportar'):
            self.assertEqual(self.request(path)[0], 401)
            self.assertEqual(self.request(path,headers=self.auth('senha-incorreta'))[0],401)

    def test_authorized_get(self):
        code,body=self.request('/api/estagios',headers=self.auth())
        self.assertEqual(code,200)
        self.assertEqual(json.loads(body)['registros'],[])
        self.assertEqual(self.request('/', headers=self.auth())[0],200)

    def test_write_requires_same_origin(self):
        body=json.dumps({'nome':'Pessoa Ficticia'}).encode()
        hdr={**self.auth(), 'Content-Type':'application/json'}
        self.assertEqual(self.request('/api/estagios','POST',hdr,body)[0],403)
        self.assertEqual(self.request('/api/estagios','POST',{**hdr, 'Origin':'https://evil.example'},body)[0],403)
        self.assertEqual(self.request('/api/estagios','POST',{**hdr, 'Origin':self.host,'Sec-Fetch-Site':'cross-site'},body)[0],403)
        code,_=self.request('/api/estagios','POST',{**hdr,'Origin':self.host},body)
        self.assertEqual(code,201)

    def test_reject_insecure_defaults_and_unc_database(self):
        with patch.object(lan, 'HOST', '0.0.0.0'), patch.object(lan,'TLS',False), patch.dict(os.environ,{},clear=True):
            with self.assertRaisesRegex(ValueError,'HTTPS'):
                lan.validate_configuration()
        with patch.object(lan,'HOST','127.0.0.1'), patch.object(lan,'TLS',False), patch.object(app,'DB_PATH',Path('//server/compartilhamento/banco.db')):
            with self.assertRaisesRegex(ValueError,'UNC'):
                lan.validate_configuration()


if __name__=='__main__': unittest.main()
