"""Leitura de XLSX OOXML sem dependências; os dados nunca saem do computador."""
from __future__ import annotations
from datetime import date, datetime, timedelta
from io import BytesIO
from collections import Counter
import json, re, unicodedata, zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

NS = {'x':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'p':'http://schemas.openxmlformats.org/package/2006/relationships'}
MAX_INPUT = 12 * 1024 * 1024
MAX_ROWS = 10000

def norm(v):
    return re.sub(r'\s+',' ', ''.join(c for c in unicodedata.normalize('NFKD', str(v or '').strip().lower()) if not unicodedata.combining(c))).strip()

def excel_date(v):
    if v is None or str(v).strip() == '': return ''
    if isinstance(v, (int,float)) or re.fullmatch(r'\d+(?:\.\d+)?',str(v).strip()):
        n=float(v)
        if n < 10000 or n > 100000: return ''
        try:return (date(1899,12,30)+timedelta(days=int(n))).isoformat()
        except (OverflowError,ValueError):return ''
    s=str(v).strip()
    for fmt in ('%d/%m/%Y','%Y-%m-%d','%d/%m/%y'):
        try:return datetime.strptime(s[:10],fmt).date().isoformat()
        except ValueError:pass
    return ''

def colnum(cell):
    col=re.match(r'[A-Z]+',cell)
    if not col:return -1
    n=0
    for c in col.group():n=n*26+ord(c)-64
    return n-1

def parse_xlsx(content:bytes):
    if len(content)>MAX_INPUT:raise ValueError('O Excel excede 12 MB.')
    with zipfile.ZipFile(BytesIO(content)) as z:
        if sum(i.file_size for i in z.infolist())>70*1024*1024:raise ValueError('Arquivo XLSX expandido excessivamente grande.')
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            root=ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('x:si',NS):
                shared.append(''.join(t.text or '' for t in si.findall('.//x:t',NS)))
        book=ET.fromstring(z.read('xl/workbook.xml'))
        rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        paths={i.get('Id'):i.get('Target') for i in rels}
        sheets=[]
        for s in book.findall('.//x:sheet',NS):
            target=paths.get(s.get('{'+NS['r']+'}id'))
            if not target:continue
            path=target.lstrip('/') if target.startswith('/') else 'xl/'+target
            if path not in z.namelist():continue
            xml=ET.fromstring(z.read(path))
            rows=[]
            for r in xml.findall('.//x:sheetData/x:row',NS):
                if len(rows)>=MAX_ROWS:raise ValueError('Limite de 10 mil linhas por aba.')
                cells={}
                for c in r.findall('x:c',NS):
                    k=colnum(c.get('r',''))
                    if k<0 or k>=80:continue
                    t=c.get('t');v=c.find('x:v',NS)
                    if t=='inlineStr':value=''.join(x.text or '' for x in c.findall('.//x:t',NS))
                    elif v is None:value=''
                    elif t=='s':
                        try:value=shared[int(v.text)]
                        except (ValueError,IndexError,TypeError):value=''
                    else:value=v.text or ''
                    cells[k]=value
                if cells:rows.append((int(r.get('r',len(rows)+1)),cells))
            sheets.append((s.get('name',''),rows))
        return sheets

ALIASES={
 'inicio':['inicio','data inicio','data de inicio'],
 'mes':['mes'],
 'confirmado':['confirmado','situacao','aprovacao'],
 'nome':['nome','nome completo','aluno','estagiario'],
 'especialidade':['especialidade','servico'],
 'cracha':['cracha'],
 'valor':['valor'],
 'fim':['termino','fim','data termino'],
 'email':['e-mail','email','correio eletronico'],
 'telefone':['telefone','celular','whatsapp'],
 'documentos':['documentos','documentacao'],
 'observacoes':['observacao','observacoes'],
 'certificado':['etapa certificado','certificado'],
 'presenca':['presenca'],
}

def rows_from_excel(data:bytes):
    sheets=parse_xlsx(data)
    found=[];sheet_info=[]
    for sheet_name,rows in sheets:
        header_i=None;headers={}
        for i,(rn,cells) in enumerate(rows[:25]):
            candidate={}
            for k,v in cells.items():
                nv=norm(v)
                for key,aliases in ALIASES.items():
                    if nv in aliases:candidate[key]=k
            if 'nome' in candidate and ('especialidade' in candidate or 'email' in candidate):
                header_i=i;headers=candidate;break
        if header_i is None:continue
        n=0
        for row_number,cells in rows[header_i+1:]:
            def v(key):return str(cells.get(headers.get(key,-1),'') or '').strip()
            nome=v('nome')
            if not nome or norm(nome) in ('total','nome'):continue
            first=excel_date(v('inicio'));last=excel_date(v('fim'))
            confirmed=v('confirmado');docs=v('documentos');badge=v('cracha');cert=v('certificado')
            if 'cancel' in norm(confirmed):approval='recusado';status='cancelado'
            elif 'confirmad' in norm(confirmed) or 'aprovad' in norm(confirmed):approval='aprovado';status='ativo'
            else:approval='pendente';status='ativo'
            dn=norm(docs)
            docstatus='pendente' if not dn or 'pendent' in dn or 'falta' in dn else ('recebidos' if 'envi' in dn else 'pendente')
            bn=norm(badge)
            badge_status='devolvido' if 'devolvid' in bn else ('retirado' if bn and ('possui' in bn or any(c.isdigit() for c in bn)) else 'pendente')
            certn=norm(cert)
            cert_status='enviado' if ('enviado' in certn and 'formulario' not in certn and 'form.' not in certn) else ('emissao' if 'emissao' in certn else 'pendente')
            duration=None
            if first and last:
                duration=(date.fromisoformat(last)-date.fromisoformat(first)).days
            issues=[]
            if not first and v('inicio'):issues.append('Data de início inválida')
            if not last and v('fim'):issues.append('Data de término inválida')
            if duration is not None:
                if duration<0:issues.append('Término anterior ao início')
                elif duration>7:issues.append(f'Período de {duration+1} dias — verificar')
            item={'nome':nome,'email':v('email').replace('\u00a0','').strip(),'telefone':v('telefone'),
             'especialidade':v('especialidade'),'inicio':first,'fim':last,
             'valor':v('valor'),'situacao_aprovacao':approval,'situacao_documentos':docstatus,
             'status_cracha':badge_status,'cracha_numero':badge if badge_status=='retirado' and 'possui' not in bn else '',
             'status_certificado':cert_status,'status':status,'observacoes':v('observacoes'),
             'documentos_origem':docs,'certificado_origem':cert,'aprovacao_origem':confirmed,
             'presenca':v('presenca'),'arquivo_aba':sheet_name,'linha_origem':row_number,
             'alertas_importacao':'; '.join(issues)}
            found.append(item);n+=1
        sheet_info.append({'aba':sheet_name,'registros':n})
    if not found:raise ValueError('Não encontrei colunas Nome + Especialidade/E-mail no Excel.')
    return found,sheet_info

# Exportação mínima, interoperável, em Excel .xlsx, sem dependências externas.
def export_xlsx(headers, rows):
    def cell(ref, value, style=0):
        if value is None:value=''
        safe=escape(str(value))
        return f'<c r="{ref}" t="inlineStr" s="{style}"><is><t xml:space="preserve">{safe}</t></is></c>'
    def letter(n):
        out='';n+=1
        while n:n,rem=divmod(n-1,26);out=chr(65+rem)+out
        return out
    xmlrows=[]
    for ri,vals in enumerate([headers]+rows,1):
        xmlrows.append(f'<row r="{ri}">'+''.join(cell(f'{letter(ci)}{ri}',v,1 if ri==1 else 0) for ci,v in enumerate(vals))+'</row>')
    cols=''.join(f'<col min="{i+1}" max="{i+1}" width="{32 if i in (1,3,4,18) else 17}" customWidth="1"/>' for i in range(len(headers)))
    ref=f'A1:{letter(len(headers)-1)}{len(xmlrows)}'
    sheet=f'''<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{NS['x']}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>{cols}</cols><sheetData>{''.join(xmlrows)}</sheetData><autoFilter ref="{ref}"/></worksheet>'''
    out=BytesIO()
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml','''<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>''')
        z.writestr('_rels/.rels','''<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>''')
        z.writestr('xl/workbook.xml','''<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Estágios" sheetId="1" r:id="rId1"/></sheets></workbook>''')
        z.writestr('xl/_rels/workbook.xml.rels','''<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>''')
        z.writestr('xl/styles.xml','''<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF133B52"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>''')
        z.writestr('xl/worksheets/sheet1.xml',sheet)
    return out.getvalue()
