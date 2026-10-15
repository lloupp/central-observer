# Central de Estágios Observacionais — Santa Casa

Aplicação local de gestão de estagiários observacionais, criada com base no fluxo descrito e na estrutura de controle da planilha fornecida. Inspirada na lógica operacional da Central de Notas, **sem compartilhar o banco de dados com aquele projeto**.

## Como iniciar no Windows

1. Extraia o arquivo ZIP em uma pasta de sua escolha.
2. Execute `INICIAR.bat` com duplo clique (requer **Python 3.10 ou superior** instalado).
3. Abra `http://127.0.0.1:8765` no navegador, caso não abra automaticamente.
4. Em **Importar Excel**, selecione a sua planilha `.xlsx`, confira os alertas e clique em **Confirmar importação**.

No Linux/macOS: `python3 app.py` ou `./iniciar.sh`.

**Para gestão, importação e e-mails não é necessário instalar bibliotecas adicionais.** Para gerar certificados Word, instale os pacotes opcionais descritos adiante.

## Principais funcionalidades

- Painel com totais, pendências operacionais, próximos inícios e cadastros recentes.
- Lista de estagiários com busca por nome, e-mail ou especialidade e filtro por etapa.
- Ficha com aprovação, documentação, datas, hospital, crachá, formulário, certificado, telefone, observações, universidade e semestre.
- Importação da planilha Excel atual, inclusive status existentes; **mostra prévia e alertas de datas antes de importar**.
- Importação repetida ignora registros já encontrados, sem sobrescrever edições efetuadas no aplicativo.
- Seis modelos de e-mail editáveis, preenchidos automaticamente com os dados de cada pessoa.
- Download de **rascunho `.eml` para revisão manual**, com o formulário `.doc` original anexado automaticamente ao modelo de orientações.
- Link `mailto:` como alternativa, **sem anexos** e limitado a mensagens curtas.
- Registro manual de envio após você efetivamente enviar o e-mail.
- Histórico de alterações por cadastro, exportação `.xlsx` de backup e controles para emissão de certificados.

## Fluxo recomendado

**Cadastro ou importação → aprovação → documentos recebidos → dados do primeiro dia → gerar rascunho de orientações → revisar e enviar no Outlook → marcar envio → registrar término → confirmar formulário e crachá devolvido → emissão e envio do certificado.**

Para gerar **Orientações de início**, é necessário registrar **Aprovado + Documentos recebidos + e-mail + datas + horário de retirada do crachá**. O aplicativo não envia nada sozinho. Para registrar **Certificado enviado**, confirme o recebimento do formulário de conclusão e a devolução do crachá.

## Segurança e dados

- O sistema atende apenas em `127.0.0.1` (computador local), não está publicado na internet.
- Os dados ficam em `data/observership.db`. **Faça cópia de segurança periódica deste arquivo com o aplicativo fechado**; exporte também um `.xlsx` pelo botão **Exportar Excel**.
- Não compartilhe a pasta `data/`, os arquivos Excel com dados pessoais nem seus rascunhos reais no GitHub.
- Este MVP foi feito para **uso em um único computador**. Para acesso de várias pessoas ou por rede, será preciso adicionar autenticação individual, permissões, armazenamento de anexos, backup automatizado e adequações LGPD.
- Não foi disponibilizado servidor de envio automático nem conexão de e-mail institucional: os rascunhos precisam ser abertos e enviados no cliente de e-mail.
- O formulário .doc deve ser carregado em Modelos de e-mail e permanece local. A equipe deve confirmar periodicamente o valor de taxas e as normas institucionais contidas no documento.

## Sobre a planilha de origem

O importador reconhece a aba `2026` e as 14 colunas da planilha recebida. Mantém como texto os campos de origem (ex.: “ENVIADO (ECOMMERCE)”, “FORMULÁRIO ENVIADO”, “S/REGISTRO”) e utiliza estados internos padronizados. **Não interpreta automaticamente “FORMULÁRIO ENVIADO” como comprovação de formulário de conclusão recebido.** A conferência humana é necessária.

Alertas são gerados para datas inválidas, datas finais anteriores às iniciais ou períodos de mais de 7 dias. Não há alteração automática nessas datas.

## Tecnologias

Python 3 (biblioteca padrão), servidor HTTP local, SQLite, JavaScript, HTML/CSS, leitor/escritor OOXML `.xlsx` sem dependências adicionais.

## Testes

```bash
python -m unittest discover -s tests -v
```

O comando usa apenas dados sintéticos de teste; não grava dados pessoais no repositório.

## Emissão de certificados — integração com gerador existente

O código do gerador em `certificados/gerar_certificados_ensino.py` veio do gerador de **Observership, Residentes, Doutorandos e Mini Fellowship**, reutilizado na aplicação. A criação de certificados individuais mantém os textos por modalidade, tratamento de nomes/especialidades, datas e carga horária previstos naquele código. A Central **não lê, reescreve ou apaga** a planilha `CERTIFICADOS ENSINO.xlsx` do gerador anterior.

1. No Windows, execute uma única vez `INSTALAR_CERTIFICADOS.bat` para instalar `docxtpl` e `openpyxl` no Python utilizado (é necessária conexão com a internet). No Linux, execute `python3 -m pip install -r requirements-certificados.txt`.
2. Abra a **Central**, clique em **Certificados** e carregue o **`Certificado modelo.docx` original** disponível no seu pacote anterior, na pasta `MODELOS`. É necessário carregar o modelo apenas uma vez por computador.
3. Na ficha do estudante, informe **modalidade**, **carga horária comprovada no formulário**, as datas e confira o status de formulário recebido e crachá devolvido.
4. Na tela **Certificados**, use **Gerar**. O Word ficará salvo na pasta privada `data/certificados/` e disponível para download. A situação muda para **Em emissão/assinatura**, sem marcar o documento como enviado.
5. Confira a formatação, providencie as assinaturas pela rotina institucional e **marque como enviado somente após o envio real**.

**Proteções:** exige aprovação, formulário e crachá; impede emissão com período invertido, campos essenciais inválidos e mais de 40 horas em Observership; guarda histórico; impede que geração seja confundida com envio. O modelo Word é conferido antes de ser salvo. A interface é de um único operador local; ainda não existe assinatura digital ou envio automático do certificado.

**Repositório público:** não inclui o modelo institucional, nenhum arquivo de certificados emitidos, nenhuma planilha com dados reais e nenhum banco de dados. Os arquivos de trabalho permanecem em `data/`, ignorados pelo Git. Para compartilhar o aplicativo com colegas, avalie autenticação e normas de acesso antes de disponibilizá-lo em rede.

### Validação

```bash
python -m unittest discover -s tests -v
```

O teste de geração completa com o modelo institucional requer `docxtpl` e o `.docx` disponibilizado localmente e, por isso, não executa esse caminho em CI pública com dados reais.

**Formulário de conclusão:** abra **Modelos de e-mail** e carregue o arquivo Word `.doc` fornecido anteriormente. Assim, o modelo será anexado aos rascunhos de orientações sem ser enviado para o GitHub.
