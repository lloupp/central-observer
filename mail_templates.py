"""Modelos revisáveis; nenhum e-mail é enviado automaticamente."""
import re
from datetime import date

TEMPLATES={
'orientacoes':{
  'nome':'Orientações para início do estágio',
  'assunto':'Estágio Observacional — Orientações | {{NOME}} | {{INICIO}}',
  'corpo':'''Prezado(a) {{PRIMEIRO_NOME}},

É com prazer que encaminhamos as orientações para o seu período de estágio observacional na Santa Casa de Porto Alegre.

Pedimos que leia atentamente as informações abaixo para que tenha uma ótima experiência durante sua vivência em nossa instituição.

DADOS DO ESTÁGIO
Especialidade: {{ESPECIALIDADE}}
Hospital/serviço: {{HOSPITAL}}
Período: {{INICIO}} a {{FIM}}

RETIRADA DO CRACHÁ E ACOLHIMENTO
Local: Sala Espaço Educação — Praça Central (Hospital Santa Clara).
Endereço: Rua Professor Annes Dias, 295 — Centro Histórico, Porto Alegre/RS.
Mapa: https://www.google.com/maps/search/?api=1&query=Rua+Professor+Annes+Dias+295+Porto+Alegre
Data: {{DATA_CRACHA}}
Horário: {{HORA_CRACHA}}

• Você fará a retirada do crachá de acesso e será orientado(a) até o local de serviço.
• Tolerância de 10 minutos.
• Nos demais dias, compareça diretamente ao local do serviço.
• O estágio é realizado de segunda a sexta-feira e é totalmente observacional.

O QUE LEVAR / USAR
• Jaleco branco (obrigatório).
• Máscara individual (opcional).
• Não utilizar adornos, como brincos, anéis, relógio ou pulseiras.

FORMULÁRIO DE CONCLUSÃO — ANEXO
• Enviamos em anexo o formulário de conclusão, no qual será registrada a carga horária cumprida.
• O documento deverá ser preenchido somente no último dia do estágio, com todas as informações solicitadas, assinatura e carimbo do preceptor/responsável do serviço. A chefia médica ou residente poderá realizar a assinatura conforme as orientações do serviço.
• Após o término do estágio, envie exclusivamente a versão digital para relacionamento.ep@santacasa.org.br.
• Não recebemos a folha física.

DEVOLUÇÃO DO CRACHÁ
• No último dia, devolva o crachá e assine a folha de devolução na recepção do Centro Administrativo — 2º andar do Edifício Garagem.
• Não devolver ou extraviar o crachá poderá gerar pendência na emissão do certificado e cobrança de taxa, conforme orientações do formulário.
• Estágios de mais de uma semana podem utilizar o mesmo crachá.
• Atendimento para devolução: segunda a quinta, das 8h às 18h; sexta, das 8h às 17h.

CERTIFICADO
O certificado será emitido conforme a carga horária informada no formulário de frequência e ficará pronto em até 30 dias úteis após o recebimento do formulário e a devolução do crachá.

OBSERVAÇÕES IMPORTANTES
• Não é disponibilizado acesso ao refeitório institucional para estagiários. Há outras opções para alimentação: Bistrô do Centro Histórico Cultural (CHC), Bistrô do Hospital Santo Antônio (térreo, próximo à Panvel) e Bistrô do Hospital Santa Rita (2º andar).
• Não é possível acompanhar mais de uma especialidade por semana/dia.
• Não há plano mensal de estacionamento para estagiários; o pagamento é diário. Consulte valores diretamente na administração do estacionamento.
• Leve somente o necessário para acompanhar o serviço.
• O cronograma das atividades será repassado pela equipe do serviço.
• A vivência acompanha a rotina hospitalar do serviço escolhido, limitada a 40 horas semanais, distribuídas a critério do serviço.

Em caso de dúvidas, esclarecimentos ou feedback, entre em contato com nosso WhatsApp comercial:
https://w.app/equipecomercialensino

Agradecemos a preferência pela Santa Casa de Porto Alegre!

Equipe de Ensino e Pesquisa
Santa Casa de Porto Alegre'''
 },
'aprovacao':{
  'nome':'Aprovação — solicitação de documentos',
  'assunto':'Estágio Observacional — Aprovação e documentos | {{NOME}}',
  'corpo':'''Prezado(a) {{PRIMEIRO_NOME}},

Informamos que sua solicitação para o estágio observacional na especialidade de {{ESPECIALIDADE}} foi aprovada.

Para dar continuidade ao processo, pedimos que encaminhe a documentação exigida, conforme as orientações recebidas no processo de inscrição.

Período previsto: {{INICIO}} a {{FIM}}.

Após a conferência da documentação, encaminharemos o e-mail com as orientações para o primeiro dia, incluindo o local e horário de retirada do crachá e o formulário de conclusão.

Em caso de dúvidas, fale conosco pelo WhatsApp: https://w.app/equipecomercialensino

Atenciosamente,
Equipe de Ensino e Pesquisa — Santa Casa de Porto Alegre'''
 },
'pendencia_documentos':{
  'nome':'Lembrete de documentos pendentes',
  'assunto':'Estágio Observacional — Documentação pendente | {{NOME}}',
  'corpo':'''Prezado(a) {{PRIMEIRO_NOME}},

Estamos acompanhando sua solicitação de estágio observacional na especialidade de {{ESPECIALIDADE}}, prevista para {{INICIO}}.

Até o momento, precisamos concluir a conferência da documentação necessária para o início da vivência. Solicitamos o envio ou regularização dos documentos pendentes, conforme orientações já repassadas.

Assim que a documentação estiver completa, poderemos encaminhar as informações do primeiro dia do estágio.

Dúvidas: https://w.app/equipecomercialensino

Atenciosamente,
Equipe de Ensino e Pesquisa — Santa Casa de Porto Alegre'''
 },
'formulario_pendente':{
  'nome':'Após o estágio — formulário de conclusão',
  'assunto':'Estágio Observacional — Formulário de conclusão | {{NOME}}',
  'corpo':'''Prezado(a) {{PRIMEIRO_NOME}},

Esperamos que sua experiência de estágio observacional na especialidade de {{ESPECIALIDADE}} tenha sido proveitosa.

Para darmos continuidade à emissão do certificado, solicitamos o encaminhamento do formulário de conclusão preenchido integralmente, com a carga horária efetivamente cumprida, assinatura e carimbo do responsável pelo serviço.

O formulário deve ser enviado exclusivamente por e-mail para relacionamento.ep@santacasa.org.br. Não recebemos a versão física.

Lembramos também que a devolução do crachá é obrigatória no Centro Administrativo — 2º andar do Edifício Garagem. A não devolução pode gerar pendência no certificado e cobrança de taxa.

Após o recebimento do formulário e a devolução do crachá, o prazo de emissão é de até 30 dias úteis.

Dúvidas: https://w.app/equipecomercialensino

Atenciosamente,
Equipe de Ensino e Pesquisa — Santa Casa de Porto Alegre'''
 },
'cracha_pendente':{
  'nome':'Após o estágio — devolução do crachá',
  'assunto':'Estágio Observacional — Devolução de crachá | {{NOME}}',
  'corpo':'''Prezado(a) {{PRIMEIRO_NOME}},

Identificamos que a devolução do crachá referente ao seu estágio observacional em {{ESPECIALIDADE}} ainda precisa ser confirmada.

Solicitamos que devolva o crachá e assine a folha de devolução no Centro Administrativo — 2º andar do Edifício Garagem.

Horários: segunda a quinta, das 8h às 18h; sexta, das 8h às 17h.

A não devolução ou o extravio pode gerar pendência na emissão do certificado e cobrança de taxa. O prazo de até 30 dias úteis para emissão do certificado começa após o recebimento do formulário de conclusão e a devolução do crachá.

Dúvidas: https://w.app/equipecomercialensino

Atenciosamente,
Equipe de Ensino e Pesquisa — Santa Casa de Porto Alegre'''
 },
'certificado':{
  'nome':'Certificado disponível',
  'assunto':'Estágio Observacional — Certificado | {{NOME}}',
  'corpo':'''Prezado(a) {{PRIMEIRO_NOME}},

Informamos que o certificado referente ao seu estágio observacional na especialidade de {{ESPECIALIDADE}}, realizado de {{INICIO}} a {{FIM}}, encontra-se disponível.

Pedimos que confira seus dados no documento encaminhado.

Agradecemos por ter realizado sua vivência em nossa instituição.

Atenciosamente,
Equipe de Ensino e Pesquisa — Santa Casa de Porto Alegre'''
 }
}

def br_date(v):
    if not v:return '[não informado]'
    try:return date.fromisoformat(v).strftime('%d/%m/%Y')
    except ValueError:return str(v)

def render(template, record):
    name=record.get('nome','').strip()
    values={
      'NOME':name or '[nome não informado]',
      'PRIMEIRO_NOME':name.split()[0] if name else '[nome]',
      'EMAIL':record.get('email') or '[não informado]',
      'ESPECIALIDADE':record.get('especialidade') or '[não informada]',
      'HOSPITAL':record.get('hospital') or 'a confirmar com a equipe',
      'INICIO':br_date(record.get('inicio')),
      'FIM':br_date(record.get('fim')),
      'DATA_CRACHA':br_date(record.get('credenciamento_data') or record.get('inicio')),
      'HORA_CRACHA':record.get('credenciamento_hora') or '[horário a confirmar]',
      'UNIVERSIDADE':record.get('universidade') or '',
      'SEMESTRE':record.get('semestre') or '',
    }
    return re.sub(r'\{\{([A-Z_]+)\}\}',lambda m: str(values.get(m.group(1), m.group(0))),template)
