import io
import os
import sys
import tempfile
import unittest
import zipfile
from datetime import date
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import app
import certificate_service as cert


def tiny_docx(missing=None):
    tags = cert.REQUIRED_TAGS - ({missing} if missing else set())
    body = ''.join(f'<w:p><w:r><w:t>{{{{{t}}}}}</w:t></w:r></w:p>' for t in tags)
    xml = ('<?xml version="1.0" encoding="UTF-8"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           f'<w:body>{body}</w:body></w:document>')
    b=io.BytesIO()
    with zipfile.ZipFile(b,'w') as z:z.writestr('word/document.xml',xml)
    return b.getvalue()


class CertificateTests(unittest.TestCase):
    def test_01_upload_model_validation(self):
        with self.assertRaisesRegex(ValueError,'inválido'):
            cert.check_template(b'not a docx')
        with self.assertRaisesRegex(ValueError,'NOME'):
            cert.check_template(tiny_docx('NOME'))
        self.assertEqual(cert.check_template(tiny_docx()),cert.REQUIRED_TAGS)

    def test_02_issuance_guardrails(self):
        record={'nome':'Pessoa Fictícia','email':'pessoa@example.org','especialidade':'Pediatria',
                'inicio':'2026-10-12','fim':'2026-10-16','modalidade':'Observership',
                'carga_horaria':'40','situacao_aprovacao':'aprovado','status':'ativo',
                'status_formulario':'recebido','status_cracha':'devolvido','status_certificado':'pendente'}
        with tempfile.TemporaryDirectory() as d:
            path=Path(d)/'template.docx'
            path.write_bytes(tiny_docx())
            with patch.object(cert,'MODEL_PATH',path),patch.object(cert,'template_status',return_value={'modelo_configurado':True,'dependencias_instaladas':True}):
                p=cert.validate_record(record)
                self.assertEqual(p['ch'],40)
                for changes,message in [({'status_formulario':'pendente'},'formulário'),
                                         ({'status_cracha':'pendente'},'crachá'),
                                         ({'carga_horaria':''},'carga horária'),
                                         ({'carga_horaria':'41'},'40 horas'),
                                         ({'fim':'2026-10-01'},'anterior'),
                                         ({'status_certificado':'enviado'},'enviado')]:
                    with self.subTest(changes=changes):
                        with self.assertRaisesRegex(ValueError,message):
                            cert.validate_record({**record,**changes})
                resident=cert.validate_record({**record,'modalidade':'Residente','carga_horaria':'80'})
                self.assertEqual(resident['ch'],80)

    def test_03_old_database_migration(self):
        original=app.DB_PATH
        with tempfile.TemporaryDirectory() as d:
            db=Path(d)/'old.db'
            import sqlite3
            with sqlite3.connect(db) as con:
                con.execute("CREATE TABLE estagios (id INTEGER PRIMARY KEY,nome TEXT, fingerprint TEXT, criado_em TEXT, atualizado_em TEXT)")
                con.execute("INSERT INTO estagios(nome,criado_em,atualizado_em) VALUES('Exemplo','2026','2026')")
            try:
                app.DB_PATH=db
                app.init_db()
                with app.connect() as con:
                    row=con.execute('SELECT * FROM estagios WHERE id=1').fetchone()
                    self.assertEqual(row['nome'],'Exemplo')
                    self.assertIn('carga_horaria',row.keys())
                    self.assertIn('certificado_arquivo',row.keys())
            finally:
                app.DB_PATH=original

    def test_04_local_storage_and_path_guard(self):
        with tempfile.TemporaryDirectory() as d:
            with patch.object(cert,'OUTPUT_DIR',Path(d)):
                name=cert.store_certificate(14,b'bytes-de-teste')
                self.assertEqual(cert.stored_certificate(name),b'bytes-de-teste')
                with self.assertRaisesRegex(ValueError,'inválido'):
                    cert.stored_certificate('../../privado.docx')

if __name__ == '__main__':unittest.main()
