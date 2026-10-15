#!/usr/bin/env python3
"""Central de Estágios Observacionais — aplicativo local sem serviços externos."""
from __future__ import annotations
import csv, hashlib, io, json, os, re, sqlite3, sys, traceback, urllib.parse, zipfile
from datetime import date, datetime, timezone
from contextlib import contextmanager
from email.message import EmailMessage
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import RLock
from excel_io import rows_from_excel, export_xlsx
from mail_templates import TEMPLATES, render
import certificate_service as certificates

BASE=Path(__file__).resolve().parent
DB_PATH=Path(os.environ.get('OBS_DB',str(BASE/'data'/'observership.db')))
HOST='127.0.0.1'
PORT=int(os.environ.get('OBS_PORT','8765'))
FORM=BASE/'data'/'formulario_conclusao.doc'
if not FORM.exists() and (BASE/'resources'/'formulario_conclusao.doc').exists():
    FORM=BASE/'resources'/'formulario_conclusao.doc'
LOCK=RLock()

FIELDS=['nome','email','telefone','especialidade','hospital','universidade','semestre','inicio','fim',
 'credenciamento_data','credenciamento_hora','valor','situacao_aprovacao','situacao_documentos',
 'status_cracha','cracha_numero','status_formulario','status_certificado','status',
 'orientacoes_enviadas_em','observacoes','documentos_origem','certificado_origem',
 'aprovacao_origem','presenca','arquivo_aba','linha_origem','arquivo_origem','alertas_importacao','modalidade','carga_horaria','certificado_arquivo','certificado_gerado_em']
READONLY={'documentos_origem','certificado_origem','aprovacao_origem','arquivo_aba','linha_origem','arquivo_origem','certificado_arquivo','certificado_gerado_em'}
ENUMS={
 'situacao_aprovacao':{'pendente','aprovado','recusado'},
 'situacao_documentos':{'pendente','recebidos'},
 'status_cracha':{'pendente','retirado','devolvido','extraviado'},
 'status_formulario':{'pendente','recebido'},
 'status_certificado':{'pendente','emissao','enviado'},
 'status':{'ativo','cancelado'},
}

def now():return datetime.now(timezone.utc).replace(microsecond=0).isoformat()

@contextmanager
def connect():
    DB_PATH.parent.mkdir(parents=True,exist_ok=True)
    con=sqlite3.connect(DB_PATH,timeout=15)
    con.row_factory=sqlite3.Row
    try:
        con.execute('PRAGMA foreign_keys=ON')
        con.execute('PRAGMA busy_timeout=15000')
        yield con
        con.commit()
    except BaseException:
        con.rollback()
        raise
    finally:
        con.close()

def init_db():
    with connect() as c:
        c.execute('PRAGMA journal_mode=WAL')
        fields='\n'.join(f'{field} TEXT DEFAULT \'\',' for field in FIELDS)
        c.execute(f'''CREATE TABLE IF NOT EXISTS estagios (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          {fields}
          fingerprint TEXT UNIQUE,
          criado_em TEXT NOT NULL,
          atualizado_em TEXT NOT NULL
        )''')
        c.execute('''CREATE TABLE IF NOT EXISTS modelos (
           chave TEXT PRIMARY KEY, assunto TEXT NOT NULL, corpo TEXT NOT NULL,
           atualizado_em TEXT NOT NULL)''')
        c.execute('''CREATE TABLE IF NOT EXISTS historico (
           id INTEGER PRIMARY KEY AUTOINCREMENT, estagio_id INTEGER NOT NULL,
           data TEXT NOT NULL, acao TEXT NOT NULL, detalhes TEXT NOT NULL,
           FOREIGN KEY (estagio_id) REFERENCES estagios(id))''')
        # Migração não destrutiva para bancos criados na versão anterior.
        existing={row[1] for row in c.execute('PRAGMA table_info(estagios)')}
        for field in FIELDS:
            if field not in existing:
                c.execute(f"ALTER TABLE estagios ADD COLUMN {field} TEXT DEFAULT ''")
        c.execute('CREATE INDEX IF NOT EXISTS idx_estagios_inicio ON estagios(inicio)')
        c.execute('CREATE INDEX IF NOT EXISTS idx_estagios_nome ON estagios(nome)')
        for k,v in TEMPLATES.items():
            c.execute('INSERT OR IGNORE INTO modelos VALUES (?,?,?,?)',(k,v['assunto'],v['corpo'],now()))

def fingerprint(record):
    parts=[record.get(k,'').strip().casefold() for k in ('email','nome','especialidade','inicio')]
    # A mesma pessoa pode ter duas solicitações idênticas; cada linha da origem
    # é tratada como registro distinto, mas reimportações da linha são ignoradas.
    parts.append(str(record.get('arquivo_aba','')).strip().casefold())
    parts.append(str(record.get('linha_origem','')).strip())
    return hashlib.sha256('|'.join(parts).encode('utf-8')).hexdigest()

def rowdict(row):return dict(row) if row else None

def stage(row):
    if row['status']=='cancelado' or row['situacao_aprovacao']=='recusado':return 'Cancelado'
    if row['status_certificado']=='enviado':return 'Certificado enviado'
    if row['status_certificado']=='emissao':return 'Em emissão/assinatura'
    if row['status_formulario']=='recebido' and row['status_cracha']=='devolvido':return 'Aguardando certificado'
    if row['fim'] and row['fim']<date.today().isoformat() and row['orientacoes_enviadas_em']:
        return 'Pós-estágio'
    if row['orientacoes_enviadas_em']:return 'Orientações enviadas'
    if row['situacao_aprovacao']=='aprovado' and row['situacao_documentos']=='recebidos':
        return 'Pronto para orientar'
    if row['situacao_aprovacao']=='aprovado':return 'Aguardando documentos'
    return 'Aguardando aprovação'

def enriched(row):
    d=rowdict(row)
    if d:d['etapa']=stage(d)
    return d

def validate(fields,update=False):
    if not isinstance(fields,dict):raise ValueError('Dados inválidos.')
    clean={}
    for k,v in fields.items():
        if k not in FIELDS or (update and k in READONLY):continue
        if v is None:v=''
        if not isinstance(v,(str,int,float,bool)):raise ValueError(f'Campo inválido: {k}')
        v=str(v).strip()
        if len(v)>5000:raise ValueError(f'Campo {k} muito grande.')
        if k=='modalidade' and v and v not in certificates.MODALIDADES:raise ValueError('Modalidade inválida.')
        if k in ENUMS and v and v not in ENUMS[k]:raise ValueError(f'Valor inválido para {k}.')
        if k in ['inicio','fim','credenciamento_data','orientacoes_enviadas_em'] and v:
            try:
                if k=='orientacoes_enviadas_em':datetime.fromisoformat(v.replace('Z','+00:00'))
                else:date.fromisoformat(v)
            except ValueError:raise ValueError(f'Data inválida em {k}.')
        clean[k]=v
    if not update and not clean.get('nome'):raise ValueError('Informe o nome do estudante.')
    return clean

def insert_record(con, fields, fp):
    v=validate(fields)
    if not v.get('modalidade'):v['modalidade']='Observership'
    v['fingerprint']=fp
    v.update(criado_em=now(),atualizado_em=now())
    cols=list(v)
    cur=con.execute('INSERT INTO estagios ('+','.join(cols)+') VALUES ('+','.join('?' for _ in cols)+')',list(v.values()))
    con.execute('INSERT INTO historico (estagio_id,data,acao,detalhes) VALUES (?,?,?,?)',(cur.lastrowid,now(),'Cadastro criado','Novo cadastro'))
    return cur.lastrowid

def import_records(records,file_name):
    created=0;duplicates=0;warnings=0
    with LOCK,connect() as con:
        for r in records:
            r=dict(r)
            r['arquivo_origem']=Path(file_name).name[:200]
            if r.get('alertas_importacao'):warnings+=1
            fp=fingerprint(r)
            cur=con.execute('SELECT id FROM estagios WHERE fingerprint=?',(fp,))
            exists=cur.fetchone()
            if exists:
                duplicates+=1
                # Evita sobrescrever atualizações manuais após uma segunda importação.
                continue
            insert_record(con,r,fp);created+=1
    return {'criados':created,'duplicados_ignorados':duplicates,'registros_com_alertas':warnings}

def require_email_ready(record,key):
    if not record['email'] or not re.fullmatch(r'[^\s@]+@[^\s@]+\.[^\s@]+',record['email'].strip()):
        raise ValueError('Cadastre um e-mail válido para gerar o rascunho.')
    if key=='orientacoes':
        if not FORM.is_file():raise ValueError('Carregue o formulário original .doc antes de preparar as orientações.')
        if record['situacao_aprovacao']!='aprovado' or record['situacao_documentos']!='recebidos':
            raise ValueError('Para enviar orientações, marque aprovação e documentos recebidos.')
        if not record['inicio'] or not record['fim'] or not record['credenciamento_hora']:
            raise ValueError('Preencha as datas do estágio e o horário de retirada do crachá.')
        if record.get('alertas_importacao') and ('Término anterior' in record['alertas_importacao'] or 'Data de' in record['alertas_importacao']):
            raise ValueError('Confira as datas inconsistentes antes de preparar as orientações.')

def get_model(con,key):
    row=con.execute('SELECT * FROM modelos WHERE chave=?',(key,)).fetchone()
    if not row:raise ValueError('Modelo não encontrado.')
    return dict(row)

def make_mail(con, id, key):
    rec=con.execute('SELECT * FROM estagios WHERE id=?',(id,)).fetchone()
    if not rec:raise ValueError('Cadastro não encontrado.')
    rec=dict(rec);require_email_ready(rec,key)
    t=get_model(con,key)
    return rec,render(t['assunto'],rec),render(t['corpo'],rec)

def make_eml(rec,subject,body,key):
    msg=EmailMessage()
    msg['To']=rec['email']
    msg['Subject']=subject
    msg['X-Unsent']='1'
    msg.set_content(body)
    if key=='orientacoes' and FORM.is_file():
        msg.add_attachment(FORM.read_bytes(),maintype='application',subtype='msword',filename='Formulario_de_Conclusao_Observership.doc')
    return msg.as_bytes()

class Handler(BaseHTTPRequestHandler):
    server_version='ObservershipLocal/1.0'
    def log_message(self,fmt,*args):
        if '/api/' in self.path or self.path=='/':super().log_message(fmt,*args)
    def reply(self,obj,code=200):
        data=json.dumps(obj,ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type','application/json; charset=utf-8')
        self.send_header('Content-Length',str(len(data)))
        self.send_header('Cache-Control','no-store')
        self.send_header('X-Content-Type-Options','nosniff')
        self.end_headers();self.wfile.write(data)
    def binary(self,data,typ,name):
        self.send_response(200)
        self.send_header('Content-Type',typ)
        self.send_header('Content-Disposition',f'attachment; filename="{name}"')
        self.send_header('Content-Length',str(len(data)))
        self.send_header('X-Content-Type-Options','nosniff')
        self.end_headers();self.wfile.write(data)
    def get_body(self,max_bytes=12*1024*1024):
        n=int(self.headers.get('Content-Length','0'))
        if n<1 or n>max_bytes:raise ValueError('Arquivo/requisição vazia ou grande demais.')
        return self.rfile.read(n)
    def json_body(self):return json.loads(self.get_body(1_000_000).decode('utf-8'))
    def safe_origin(self):
        origin=self.headers.get('Origin','')
        return not origin or origin in {f'http://127.0.0.1:{PORT}',f'http://localhost:{PORT}'}
    def run_action(self):
        if not self.safe_origin():self.reply({'erro':'Origem não permitida'},403);return
        try:
            parsed=urllib.parse.urlsplit(self.path)
            path=parsed.path
            q=urllib.parse.parse_qs(parsed.query)
            if self.command=='GET':return self.get_handler(path,q)
            if self.command=='POST':return self.post_handler(path,q)
            if self.command=='PATCH':return self.patch_handler(path,q)
            self.reply({'erro':'Método não permitido'},405)
        except (ValueError,KeyError,sqlite3.IntegrityError,zipfile.BadZipFile, json.JSONDecodeError) as ex:
            self.reply({'erro':str(ex)},400)
        except Exception:
            traceback.print_exc()
            self.reply({'erro':'Erro interno ao processar solicitação.'},500)
    def do_GET(self):self.run_action()
    def do_POST(self):self.run_action()
    def do_PATCH(self):self.run_action()
    def get_handler(self,path,q):
        if path=='/':return self.static_file('/index.html')
        if path.startswith('/static/'):
            return self.static_file(path.replace('/static/','/',1))
        if path=='/api/estagios':
            with connect() as con:
                records=[enriched(r) for r in con.execute('SELECT * FROM estagios ORDER BY inicio DESC,id DESC')]
            return self.reply({'registros':records})
        if path=='/api/modelos':
            with connect() as con:
                model=[dict(r) for r in con.execute('SELECT * FROM modelos ORDER BY chave')]
            return self.reply({'modelos':[{**r,'nome':TEMPLATES[r['chave']]['nome']} for r in model]})
        if path=='/api/historico':
            id=int(q.get('id',['0'])[0]);
            with connect() as con:
                data=[dict(r) for r in con.execute('SELECT * FROM historico WHERE estagio_id=? ORDER BY id DESC LIMIT 50',(id,))]
            return self.reply({'historico':data})
        if path=='/api/email':
            id=int(q.get('id',['0'])[0]);key=q.get('modelo',['orientacoes'])[0]
            with connect() as con:rec,subject,body=make_mail(con,id,key)
            return self.reply({'destinatario':rec['email'],'assunto':subject,'corpo':body,
                'anexo':key=='orientacoes' and FORM.is_file(),'modelo':key})
        if path=='/api/email/rascunho':
            id=int(q.get('id',['0'])[0]);key=q.get('modelo',['orientacoes'])[0]
            with connect() as con:rec,subject,body=make_mail(con,id,key)
            return self.binary(make_eml(rec,subject,body,key),'message/rfc822',f'rascunho_observership_{id}_{key}.eml')
        if path=='/api/exportar':
            with connect() as con:rows=[dict(x) for x in con.execute('SELECT * FROM estagios ORDER BY inicio,id')]
            columns=['id']+FIELDS+['etapa','criado_em','atualizado_em']
            matrix=[[stage(x) if k=='etapa' else x.get(k,'') for k in columns] for x in rows]
            return self.binary(export_xlsx(columns,matrix),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','observership_exportacao.xlsx')
        if path=='/api/certificados/configuracao':
            return self.reply(certificates.template_status())
        if path=='/api/certificados/baixar':
            id=int(q.get('id',['0'])[0])
            with connect() as con:
                rec=con.execute('SELECT certificado_arquivo FROM estagios WHERE id=?',(id,)).fetchone()
            if not rec or not rec['certificado_arquivo']:
                raise ValueError('Nenhum certificado gerado para este cadastro.')
            return self.binary(certificates.stored_certificate(rec['certificado_arquivo']),
                               'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                               f'certificado_observership_{id}.docx')
        if path=='/api/formulario/configuracao':return self.reply({'formulario_configurado':FORM.is_file()})
        if path=='/api/formulario':
            if not FORM.is_file():raise ValueError('Formulário não encontrado.')
            return self.binary(FORM.read_bytes(),'application/msword','Formulario_de_Conclusao_Observership.doc')
        return self.reply({'erro':'Não encontrado'},404)
    def static_file(self,path):
        if path not in ('/index.html','/app.js','/style.css'):
            return self.reply({'erro':'Não encontrado'},404)
        p=BASE/'static'/path.lstrip('/')
        if not p.is_file():return self.reply({'erro':'Recurso não encontrado'},404)
        typ='text/html' if path.endswith('.html') else ('text/javascript' if path.endswith('.js') else 'text/css')
        b=p.read_bytes();self.send_response(200)
        self.send_header('Content-Type',typ+'; charset=utf-8')
        self.send_header('Content-Length',str(len(b)))
        self.send_header('Cache-Control','no-store')
        self.send_header('X-Content-Type-Options','nosniff')
        self.send_header('Content-Security-Policy',"default-src 'self'; style-src 'self'; script-src 'self'; img-src 'self' data:; connect-src 'self'; base-uri 'none'; object-src 'none'; form-action 'self'")
        self.end_headers();self.wfile.write(b)
    def post_handler(self,path,q):
        if path=='/api/formulario/modelo':
            if self.headers.get('Content-Type','').split(';')[0] not in ('application/msword','application/octet-stream'):
                raise ValueError('Envie o formulário original no formato .doc.')
            raw=self.get_body(2_000_000)
            if not raw.startswith(bytes.fromhex('D0CF11E0A1B11AE1')):
                raise ValueError('Arquivo não é um documento Word .doc válido.')
            target=BASE/'data'/'formulario_conclusao.doc'
            target.parent.mkdir(parents=True,exist_ok=True)
            target.write_bytes(raw)
            global FORM
            FORM=target
            return self.reply({'ok':True,'formulario_configurado':True})
        if path=='/api/certificados/modelo':
            if self.headers.get('Content-Type','').split(';')[0] != 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                raise ValueError('Envie um modelo Word .docx.')
            certificates.save_template(self.get_body(5_000_000))
            return self.reply({'ok':True,**certificates.template_status()})
        if path=='/api/certificados/gerar':
            data=self.json_body();id=int(data['id'])
            with LOCK,connect() as con:
                row=con.execute('SELECT * FROM estagios WHERE id=?',(id,)).fetchone()
                if not row:raise ValueError('Cadastro não encontrado.')
                # Só marcar emissão após o arquivo Word ter sido realmente gerado.
                filename,docx=certificates.generate_record(dict(row))
                stored=certificates.store_certificate(id,docx)
                try:
                    con.execute('UPDATE estagios SET status_certificado=?,certificado_arquivo=?,certificado_gerado_em=?,atualizado_em=? WHERE id=?',
                                ('emissao',stored,now(),now(),id))
                    con.execute('INSERT INTO historico(estagio_id,data,acao,detalhes) VALUES (?,?,?,?)',
                                (id,now(),'Certificado gerado','Arquivo DOCX criado para conferência e assinatura; não enviado'))
                except BaseException:
                    (certificates.OUTPUT_DIR/stored).unlink(missing_ok=True)
                    raise
            return self.reply({'ok':True,'arquivo':filename,'download':f'/api/certificados/baixar?id={id}'})
        if path=='/api/estagios':
            values=validate(self.json_body())
            if not values.get('situacao_aprovacao'):values['situacao_aprovacao']='pendente'
            if not values.get('situacao_documentos'):values['situacao_documentos']='pendente'
            with LOCK,connect() as con:
                id=insert_record(con,values,None)
                result=enriched(con.execute('SELECT * FROM estagios WHERE id=?',(id,)).fetchone())
            return self.reply({'registro':result},201)
        if path in ('/api/importar/preview','/api/importar/confirmar'):
            fn=urllib.parse.unquote(q.get('nome',['controle.xlsx'])[0])
            if not fn.lower().endswith('.xlsx'):raise ValueError('Envie um arquivo .xlsx (Excel).')
            if self.headers.get('Content-Type','')!='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                raise ValueError('Tipo de arquivo inválido.')
            data=self.get_body();rows,sheets=rows_from_excel(data)
            issues=[{'linha':r['linha_origem'],'aba':r['arquivo_aba'],'problema':r['alertas_importacao']} for r in rows if r['alertas_importacao']]
            if path.endswith('preview'):
                return self.reply({'total':len(rows),'abas':sheets,'amostra':[{'nome':r['nome'],'especialidade':r['especialidade'],'inicio':r['inicio']} for r in rows[:6]],
                  'alertas':issues[:100],'total_alertas':len(issues)})
            return self.reply(import_records(rows,fn))
        if path=='/api/registrar-envio':
            data=self.json_body();id=int(data['id']);key=data.get('modelo')
            if key not in TEMPLATES:raise ValueError('Modelo inválido.')
            with LOCK,connect() as con:
                rec=con.execute('SELECT * FROM estagios WHERE id=?',(id,)).fetchone()
                if not rec:raise ValueError('Cadastro não encontrado')
                require_email_ready(dict(rec),key)
                if key=='orientacoes':
                    con.execute('UPDATE estagios SET orientacoes_enviadas_em=?,atualizado_em=? WHERE id=?',(now(),now(),id))
                con.execute('INSERT INTO historico(estagio_id,data,acao,detalhes) VALUES(?,?,?,?)',(id,now(),'Envio confirmado manualmente',TEMPLATES[key]['nome']))
            return self.reply({'ok':True})
        if path.startswith('/api/modelos/'):
            key=path.rsplit('/',1)[-1]
            if key not in TEMPLATES:raise ValueError('Modelo não encontrado.')
            body=self.json_body()
            subject=str(body.get('assunto','')).strip();content=str(body.get('corpo','')).strip()
            if not subject or not content or len(subject)>300 or len(content)>15000:raise ValueError('Assunto/corpo inválido.')
            with LOCK,connect() as con:
                con.execute('UPDATE modelos SET assunto=?,corpo=?,atualizado_em=? WHERE chave=?',(subject,content,now(),key))
            return self.reply({'ok':True})
        return self.reply({'erro':'Não encontrado'},404)
    def patch_handler(self,path,q):
        if not path.startswith('/api/estagios/'):
            return self.reply({'erro':'Não encontrado'},404)
        id=int(path.rsplit('/',1)[-1]);values=validate(self.json_body(),update=True)
        if not values:raise ValueError('Nenhuma atualização informada.')
        with LOCK,connect() as con:
            existing=con.execute('SELECT * FROM estagios WHERE id=?',(id,)).fetchone()
            if not existing:raise ValueError('Cadastro não encontrado.')
            merged={**dict(existing),**values}
            if values.get('status_certificado')=='enviado' and existing['status_certificado']!='enviado' and not (merged['status_formulario']=='recebido' and merged['status_cracha']=='devolvido'):
                raise ValueError('Confirme formulário recebido e crachá devolvido antes de marcar certificado enviado.')
            values['atualizado_em']=now()
            con.execute('UPDATE estagios SET '+','.join(f'{k}=?' for k in values)+' WHERE id=?',list(values.values())+[id])
            notes=', '.join(f'{k}: {str(existing[k])[:30]} → {str(v)[:30]}' for k,v in values.items() if k!='atualizado_em' and str(existing[k] or '')!=v)
            if notes:con.execute('INSERT INTO historico(estagio_id,data,acao,detalhes) VALUES(?,?,?,?)',(id,now(),'Cadastro atualizado',notes[:1000]))
            record=enriched(con.execute('SELECT * FROM estagios WHERE id=?',(id,)).fetchone())
        return self.reply({'registro':record})

def main():
    init_db()
    print(f'\nCentral de Estágios Observacionais\nAbra: http://{HOST}:{PORT}\nBanco: {DB_PATH}\nPressione Ctrl+C para encerrar.\n')
    server=ThreadingHTTPServer((HOST,PORT),Handler)
    try:
        import webbrowser
        if os.environ.get('OBS_NO_BROWSER')!='1':webbrowser.open(f'http://{HOST}:{PORT}')
    except Exception:pass
    try:server.serve_forever()
    except KeyboardInterrupt:print('\nEncerrado.')

if __name__=='__main__':main()
