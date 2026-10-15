import email, importlib, io, json, os, tempfile, threading, unittest, urllib.error, urllib.request
from pathlib import Path
from http.server import ThreadingHTTPServer
from tempfile import TemporaryDirectory
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))

import app
from excel_io import rows_from_excel, export_xlsx
from mail_templates import render, TEMPLATES

class CentralTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp=TemporaryDirectory()
        app.DB_PATH=Path(cls.tmp.name)/'database.db'
        cls.form_original=app.FORM
        app.FORM=Path(cls.tmp.name)/'formulario.doc'
        app.FORM.write_bytes(bytes.fromhex('D0CF11E0A1B11AE1')+b'arquivo sintetico ' * 150)
        app.init_db()
        cls.server=ThreadingHTTPServer(('127.0.0.1',0),app.Handler)
        app.PORT=cls.server.server_port
        cls.thread=threading.Thread(target=cls.server.serve_forever,daemon=True)
        cls.thread.start()
        cls.base=f'http://127.0.0.1:{app.PORT}'
    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown();cls.server.server_close();app.FORM=cls.form_original;cls.tmp.cleanup()
    def call(self,route,payload=None,method='GET',ctype='application/json'):
        data=(json.dumps(payload).encode() if ctype=='application/json' and payload is not None else payload)
        request=urllib.request.Request(self.base+route,data=data,method=method,headers={'Content-Type':ctype})
        try:
            with urllib.request.urlopen(request,timeout=5) as res:
                content=res.read()
                return res.status,(json.loads(content) if res.headers.get_content_type()=='application/json' else content)
        except urllib.error.HTTPError as e:
            return e.code,json.loads(e.read())
    def test_01_empty_dashboard_and_templates(self):
        code,data=self.call('/api/estagios')
        self.assertEqual(code,200);self.assertEqual(len(data['registros']),0)
        code,data=self.call('/api/modelos')
        self.assertEqual(len(data['modelos']),6)
    def test_02_manual_flow_and_guardrails(self):
        r={'nome':'Estudante Exemplo','email':'teste@example.com','especialidade':'Clínica Médica',
           'inicio':'2026-10-19','fim':'2026-10-23','credenciamento_hora':'8h'}
        code,data=self.call('/api/estagios',r,'POST')
        self.assertEqual(code,201)
        id=data['registro']['id'];self.__class__.rid=id
        self.assertEqual(data['registro']['etapa'],'Aguardando aprovação')
        code,data=self.call(f'/api/email?id={id}&modelo=orientacoes')
        self.assertEqual(code,400);self.assertIn('aprovação',data['erro'])
        code,data=self.call('/api/estagios/'+str(id),{'situacao_aprovacao':'aprovado','situacao_documentos':'recebidos'},'PATCH')
        self.assertEqual(code,200);self.assertEqual(data['registro']['etapa'],'Pronto para orientar')
        code,data=self.call(f'/api/email?id={id}&modelo=orientacoes')
        self.assertEqual(code,200)
        self.assertIn('19/10/2026',data['corpo'])
        self.assertIn('Sala Espaço Educação',data['corpo'])
        self.assertTrue(data['anexo'])
        code,binary=self.call(f'/api/email/rascunho?id={id}&modelo=orientacoes')
        self.assertEqual(code,200)
        msg=email.message_from_bytes(binary)
        self.assertEqual(msg['X-Unsent'],'1')
        attachments=[p for p in msg.walk() if p.get_content_disposition()=='attachment']
        self.assertEqual(len(attachments),1)
        self.assertGreater(len(attachments[0].get_payload(decode=True)),1000)
        code,data=self.call('/api/registrar-envio',{'id':id,'modelo':'orientacoes'},'POST')
        self.assertEqual(code,200)
        code,data=self.call('/api/estagios')
        self.assertTrue(data['registros'][0]['orientacoes_enviadas_em'])
        code,data=self.call('/api/estagios/'+str(id),{'status_certificado':'enviado'},'PATCH')
        self.assertEqual(code,400)
        code,data=self.call('/api/estagios/'+str(id),{'status_cracha':'devolvido','status_formulario':'recebido','status_certificado':'enviado'},'PATCH')
        self.assertEqual(code,200)
        self.assertEqual(data['registro']['etapa'],'Certificado enviado')
    def test_03_excel_preview_repeat_import_preserves_changes(self):
        col=['Inicio','CONFIRMADO','Nome ','Especialidade ','Crachá','Término ','E-mail','Documentos','Etapa Certificado ']
        rows=[['45683','CONFIRMADO','Pessoa de Teste','Pediatria','','45687','student@example.com','Enviado',''],
              ['45683','CONFIRMADO','Outro Teste','Cardiologia','','46052','outro@example.com','Pendente','']]
        payload=export_xlsx(col,rows)
        parsed,sheets=rows_from_excel(payload)
        self.assertEqual(len(parsed),2)
        self.assertTrue(parsed[1]['alertas_importacao'])
        url='/api/importar/preview?nome=controle.xlsx'
        code,preview=self.call(url,payload,'POST','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.assertEqual(code,200);self.assertEqual(preview['total'],2);self.assertEqual(preview['total_alertas'],1)
        code,imported=self.call('/api/importar/confirmar?nome=controle.xlsx',payload,'POST','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.assertEqual(imported['criados'],2)
        code,again=self.call('/api/importar/confirmar?nome=controle.xlsx',payload,'POST','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.assertEqual(again['criados'],0);self.assertEqual(again['duplicados_ignorados'],2)
        code,xlsx=self.call('/api/exportar')
        self.assertEqual(code,200)
        self.assertIn(b'PK',xlsx[:4])
        exported,_=rows_from_excel(xlsx)
        self.assertEqual(len(exported),3)
    def test_04_template_edit_and_invalid_data(self):
        code,res=self.call('/api/modelos/aprovacao',{'assunto':'Aprovado! {{NOME}}','corpo':'Olá, {{PRIMEIRO_NOME}}!'},'POST')
        self.assertEqual(code,200)
        code,mail=self.call(f'/api/email?id={self.rid}&modelo=aprovacao')
        self.assertEqual(code,200);self.assertIn('Estudante Exemplo',mail['assunto'])
        self.assertIn('Olá, Estudante!',mail['corpo'])
        code,bad=self.call('/api/estagios/'+str(self.rid),{'situacao_aprovacao':'inventado'},'PATCH')
        self.assertEqual(code,400)
        code,bad=self.call('/api/estagios',{ 'nome':'Sem validação','inicio':'2026-44-00'},'POST')
        self.assertEqual(code,400)
    def test_05_registros_iguais_em_linhas_distintas_sao_preservados(self):
        r={'nome':'Pessoa Igual','email':'duplicado@example.com','especialidade':'Cardiologia',
           'inicio':'2026-10-19','arquivo_aba':'2026'}
        fp1=app.fingerprint({**r,'linha_origem':20})
        fp2=app.fingerprint({**r,'linha_origem':21})
        self.assertNotEqual(fp1,fp2)
        self.assertEqual(fp1,app.fingerprint({**r,'linha_origem':20}))
    def test_06_origin_protection(self):
        request=urllib.request.Request(self.base+'/api/estagios',method='GET',headers={'Origin':'https://site-malicioso.example'})
        with self.assertRaises(urllib.error.HTTPError) as context:urllib.request.urlopen(request,timeout=3)
        self.assertEqual(context.exception.code,403)

if __name__=='__main__':unittest.main()
