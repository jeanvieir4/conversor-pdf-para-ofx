# Conversor de extrato bancario PDF -> Excel - Jean Vieira

## O que e o projeto

Ferramenta offline (sem IA, sem API, sem custo) que le extratos bancarios em
PDF de bancos e cooperativas usados pelos clientes de Jean, contador
(igrejas evangelicas e associacoes do terceiro setor) e gera, pra cada PDF:
um Excel com 4 colunas fixas (Data dd/mm/aaaa, Historico, Valor sempre
positivo com virgula decimal, Tipo C ou D) e um `.ofx` por banco pronto pra
importar no sistema de conciliacao bancaria.

Motivo: o cliente (Jean, contador) atualmente faz isso colando o PDF num
prompt de IA (Gemini) e copiando o resultado a mao pra planilha, depois
colando os valores no OFX Generator pra gerar o OFX. Objetivo e eliminar
essas etapas manuais com um programa que ele roda com dois cliques.

## Arquitetura

- `converter.py`: ponto de entrada. Le o(s) PDF(s) passado(s) como argumento,
  identifica o banco de cada pagina (funcao `identificar_banco`, lista
  `DETECTORES`), agrupa paginas consecutivas do mesmo banco
  (`extrair_blocos_por_banco`), chama o parser correspondente de `bancos.py`,
  e gera um `.xlsx` com **uma aba por banco** (nao mistura bancos na mesma
  aba, e nao coloca o nome do banco dentro do texto do Historico) mais uma
  aba "Pendencias" com avisos gerais.
- `bancos.py`: um `parse_<banco>(texto)` por banco, cada um retornando
  `(lista_de_transacoes, aviso_ou_None)`. Cada transacao e um dict com
  `data` (objeto `date`), `historico` (str), `valor` (Decimal, sempre
  positivo), `tipo` ('C' ou 'D'), `obs` (string, geralmente vazia).
- `ofx_export.py`: `gerar_ofx(transacoes, banco, caminho_saida)` gera um
  `.ofx` (formato OFX 1.02 SGML) a partir da mesma lista de transacoes usada
  no Excel. Formato validado batendo **byte a byte** contra um `.ofx` real
  exportado do OFX Generator (ferramenta que Jean usava manualmente antes
  disso) pro mesmo extrato Sicoob - por isso o cabecalho, a ausencia de
  bloco LEDGERBAL, o `CURDEF>USD` (nao BRL - e assim no arquivo de
  referencia) e o formato de `FITID`/`CHECKNUM` (`aammdd` + sequencial de
  2 digitos que reinicia a cada dia) nao devem ser alterados sem novo
  arquivo de referencia pra comparar. `BANKID` sai preenchido sozinho via
  `COMPE_BANCO` (codigo Febraban de 4 digitos, com zero a esquerda; 0756
  do Sicoob confirmado contra o arquivo de referencia, os demais sao a
  tabela publica Febraban). `ACCTID` sai fixo em `00000` (numero de conta
  nao da pra extrair do PDF de forma confiavel; usuario edita manualmente
  se o sistema de conciliacao exigir o numero real). Um `.ofx` por banco
  encontrado no PDF, nomeado `<nome_do_pdf>_<Banco>.ofx`.
- `Converter_Extratos.bat`: roda `python converter.py` em todos os PDFs da
  pasta onde o .bat esta, pra nao precisar usar linha de comando.
- `Para_Equipe/Conversor_Extratos.exe`: executavel standalone (PyInstaller,
  `--onefile`) gerado a partir de `converter.py`, pra distribuir pra equipe
  sem precisar instalar Python. Aceita duplo clique (processa todo PDF da
  pasta onde o .exe esta) ou arrastar PDFs em cima do icone (processa so os
  arrastados). Ver `Para_Equipe/LEIA-ME.txt` pras instrucoes que vao pra
  equipe.
- `Para_Equipe/LICENSE.txt`: licenca MIT (uso livre, "como esta", sem
  garantia), pra distribuicao mais ampla (nao so o time direto do Jean).

## Como gerar/atualizar o executavel

Sempre que `converter.py` ou `bancos.py` mudar e for hora de distribuir uma
nova versao pra equipe:

```
python -m pip install pyinstaller   # so na primeira vez
python -m PyInstaller --onefile --name Conversor_Extratos --console converter.py
```

Isso gera `dist/Conversor_Extratos.exe`. Copiar esse arquivo por cima do
`Para_Equipe/Conversor_Extratos.exe` pra atualizar a versao da equipe. O
`Conversor_Extratos.spec` guarda a configuracao do build (nao precisa mexer
nele pra um build padrao).

Depois de copiar o novo `.exe`, regerar o `Para_Equipe/Conversor_Extratos.zip`
(exe + `LEIA-ME.txt` + `LICENSE.txt`) e subir pra release do GitHub por cima
do asset existente (mesmo link de sempre, tag `1.0` - nunca criar uma tag
nova so pra isso, senao os links ja compartilhados quebram). Ver historico
de commits recentes pra copiar o comando de upload via API (delete do asset
antigo + upload do novo, usando o token do `git credential fill`).

### Verificacao automatica de atualizacao

O programa checa sozinho (numa thread separada, sem travar nada, ignora
qualquer erro silenciosamente) se tem versao mais nova disponivel,
comparando `VERSAO_ATUAL` (constante em `converter.py`) contra o conteudo
do arquivo `VERSION` no repositorio (lido via
`raw.githubusercontent.com` - nao manda nenhum dado do usuario, so le um
numero de volta). **Sempre que gerar uma nova versao pra distribuir:**

1. Atualizar `VERSAO_ATUAL` em `converter.py` (ex: `'1.1'`).
2. Atualizar o arquivo `VERSION` na raiz do repo com o MESMO numero.
3. Recompilar, testar, e subir o novo `.exe`/zip pro mesmo link de sempre
   (ver instrucoes acima - a tag da release NAO muda).
4. Fazer commit e push de tudo (codigo + VERSION) ANTES ou junto de
   publicar o novo `.exe`, senao quem ja tem a versao antiga vai ver o
   aviso de atualizacao mas o link vai entregar a mesma versao antiga.
5. **Arquivar essa versao numa release separada** (pedido do Jean: quem
   precisar de uma versao antiga tem que conseguir baixar). Criar uma
   release nova com tag `vX.Y` (ex: `v1.1`) e subir o MESMO zip la, com
   nome tipo `Conversor_Extratos_v1.1.zip`. Essa release e so pra arquivo/
   historico - o link estavel principal (tag `1.0`) continua sendo o unico
   divulgado no README/LEIA-ME/pagina do projeto e o unico usado pelo
   aviso de atualizacao.

Cuidado: o `raw.githubusercontent.com` tem cache de CDN (na pratica
alguns minutos) - depois de dar push no `VERSION`, pode levar um tempinho
pra quem ja esta rodando o programa ver o aviso. Isso e esperado, nao e
bug.

## Regra de ouro de cada parser

Nunca advinhar layout de banco sem ver uma amostra real. Sempre que um
parser for escrito ou ajustado, validar o resultado batendo a soma dos
creditos e a soma dos debitos contra os totais que o proprio extrato declara
(normalmente aparecem como "Total", "Saldo do periodo" ou similar). Isso ja
pegou varios bugs sutis de extracao de texto de PDF (linhas quebradas,
colunas que se misturam etc).

Armadilha mais comum: a extracao de texto de PDF (pdfplumber) as vezes
quebra uma linha visual em 2 ou 3 linhas de texto (descricao antes ou depois
do valor), ou junta coisas de forma inconsistente dentro do MESMO documento
(ja aconteceu do Bradesco ter um padrao numa pagina e outro padrao 2 paginas
depois). Sempre inspecionar o texto bruto (`page.extract_text()`) antes de
escrever o regex, e testar contra o extrato inteiro, nao so a primeira
pagina.

## Status por banco (atualizado nesta sessao)

### Validados com confianca total (bate exatamente com os totais do extrato)
- Sicoob - CUIDADO com valores de milhar (>= 1.000,00): a coluna do valor
  fica mais larga no PDF e o pdfplumber perde o alinhamento vertical com a
  linha "DD/MM descricao" de tres jeitos diferentes vistos no MESMO
  extrato real (bug encontrado e corrigido na v1.6, reportado pelo Jean -
  valores grandes sumindo do .ofx sem nenhum aviso):
  1. Só o C/D "estoura" pra linha seguinte, sozinho:
     `'05/08 PIX EMIT.OUTRA IF 4.554,00'` / `'D'`.
  2. O VALOR inteiro fica deslocado pra ANTES da linha data+descricao
     (bloco de 3 linhas): `'1.200,00'` / `'03/08 PIX RECEB.OUTRA IF'` /
     `'C'`. Testado com `layout=True` no pdfplumber tambem - o problema
     persiste, entao NAO e ordem de leitura errada, e o proprio PDF que
     tem esse campo numa posicao vertical ligeiramente diferente quando o
     numero fica mais largo.
  3. `parse_sicoob` reconstroi os dois casos ANTES de aplicar o regex
     principal, juntando as linhas fragmentadas numa so.
  - "DEP.CHEQUE BLOQ.XD" (deposito de cheque ainda em compensacao) nao usa
    C/D no final da linha, usa um asterisco (`'6.254,00*'`). ARMADILHA:
    parece um credito (e um deposito), mas NAO e - o extrato so soma esse
    valor no saldo quando aparece depois um lancamento separado
    "LIBER.DEPOSITO BLOQ" (esse sim com C normal) com o MESMO valor.
    Contar os dois como credito duplica o dinheiro - foi exatamente isso
    que quebrou a reconciliacao de saldo na v1.6 antes da correcao final
    (excesso de credito batendo exatamente com a soma dos "BLOQ.XD"
    daquele extrato). Corrigido ignorando toda linha com asterisco (igual
    as linhas de SALDO) - o credito real chega via "LIBER.DEPOSITO BLOQ"
    normalmente. "SALDO BLOQ.ANTERIOR" tambem usa asterisco e ja era
    ignorada por outro motivo (esta na lista `excluir`).
  Validado com extrato real de 251 lancamentos: saldo anterior (35.715,72)
  + creditos - debitos bateu exato com o saldo final declarado
  (88.070,80).
  DOIS layouts diferentes, `parse_sicoob` detecta qual e pela presenca da
  linha "PERIODO:" no cabecalho:
  1. `_parse_sicoob_curto` (o de cima) - linha "DD/MM descricao valorCD"
     (sem ano na data, sem coluna Documento), precisa da linha
     "PERIODO: DD/MM/AAAA - DD/MM/AAAA" pra saber o ano de cada DD/MM.
  2. `_parse_sicoob_detalhado` (v1.8) - visto noutra cooperativa do
     sistema (SICOOB VALE DOS PINHAIS). Nao tem linha "PERIODO:" (usa
     "DD/MM/AAAA EXTRATO CONTA CORRENTE HH:MM:SS" como data de emissao).
     Linha de lancamento tem o ANO completo na data e uma coluna
     "Documento" a mais (numero ou a palavra "Pix") entre data e
     historico: "DD/MM/AAAA [documento] Historico ValorCD". Linhas de
     detalhe complementar depois (REM.:, "Recebimento Pix", nome do
     favorecido, CPF mascarado, "Dizimo") sao ignoradas, mesma decisao
     usada no Conversor OCR - nao vale o risco de juntar errado.
     Lancamentos de "RDC AUTOMATICO" (aplicacao automatica) SAO contados
     normalmente (debito ao aplicar, credito ao resgatar) - diferente do
     "Aplic Aut Mais" do Itau (ver abaixo), aqui e a propria conta
     corrente que movimenta esse dinheiro. Validado com extrato real: 89
     lancamentos, saldo anterior (20.456,14) + creditos - debitos bateu
     exato com o "SALDO EM CONTA" do resumo final (5.494,91) - NAO com o
     "SALDO DISPONIVEL" (24.766,37), que soma tambem o RDC automatico
     (produto separado, fora do escopo de conta corrente).
- Caixa Economica Federal (Gerenciador Caixa)
- Itau (excluir linhas de "Aplic Aut Mais", que sao varredura automatica
  pra CDB e nao contam como movimento de conta corrente conforme o proprio
  extrato)
- UniCred
- Cresol - TRES formatos completamente diferentes, `parse_cresol` detecta
  qual e e roteia pro parser certo:
  1. Extrato tradicional (descricao as vezes quebra em 2 linhas com a linha
     de data+valor no meio).
  2. Extrato do sistema web "Colmeia" (`_parse_cresol_colmeia`), identificado
     por `EXTRATO CONSOLIDADO DE CONTA CORRENTE` ou pela URL
     `sistema.confesol/colmeia` no texto. Esse formato tem um problema
     serio de extracao: o `page.extract_text()` padrao do pdfplumber
     EMBARALHA a pagina inteira (parece ser peculiaridade de como esse
     sistema gera o PDF por "impressao" de pagina web). A solucao foi um
     fallback em `converter.py` (`_texto_robusto_por_caracteres`) que
     reconstroi as linhas agrupando os caracteres pela posicao vertical
     (`char['top']`) na ordem em que aparecem no fluxo do PDF, em vez de
     confiar no algoritmo de layout do pdfplumber - os caracteres em si
     estao em ordem de leitura correta no PDF, so o agrupamento em linhas
     que falha. O fallback so ativa quando detecta a URL
     `sistema.confesol/colmeia` no texto (ver `_FINGERPRINTS_TEXTO_EMBARALHADO`).
     Cada pagina tem tambem uma marca d'agua diagonal (nome de quem gerou o
     extrato, repetida varias vezes) que vira uma sopa de fragmentos curtos
     nesse agrupamento - `_parse_cresol_colmeia` filtra isso (linhas curtas
     demais, linhas com a URL do sistema, e usa o fato de que toda
     categoria real de transacao vem em CAIXA ALTA pra nao deixar um resto
     de marca d'agua sobrescrever a categoria certa quando uma transacao
     cai bem na quebra de pagina). Ver `EXTRATO CONSOLIDADO...pdf` usado
     pra validar - saldo inicial + soma dos lancamentos bateu exatamente
     com o saldo final declarado no extrato (13.474,42).
  3. `_parse_cresol_consolidado_diario` (v1.5) - MESMO titulo de pagina
     que o formato 2 ("EXTRATO CONSOLIDADO DE CONTA CORRENTE"), mas
     estrutura de linha bem mais simples - cada lancamento inteiro numa
     unica linha, sem quebrar categoria/detalhe/valor:
     `03/08/2026 PIX CREDITO DE: FERNANDA DOMINGUES S - 01/08 220,00 C`.
     "SALDO ANTERIOR" aparece uma vez por DIA (nao so no inicio do
     periodo) mostrando o saldo de abertura daquele dia - ignorado, nao e
     lancamento. Termina com um bloco `(=)SALDO: ...` seguido de
     `LANCAMENTOS FUTUROS/PENDENTES` (parcelas ainda nao lancadas, sem
     C/D no final da linha) - tudo isso e cortado. Distinguido do formato
     2 pela AUSENCIA de `SALDO ANT.:` (abreviado, com ponto e dois
     pontos) - o formato 2 sempre tem essa string exata (e o proprio
     regex dele depende disso), o formato 3 usa "SALDO ANTERIOR" por
     extenso. Validado com extrato real: 113 lancamentos, saldo do
     periodo inicial (49.452,12) + creditos - debitos bateu exato com o
     saldo final declarado (44.849,73).
  IMPORTANTE (bug real encontrado e corrigido na v1.5): o formato 3 expos
  uma variante nova do problema de falso positivo do Bradesco. O nome do
  banco ("CRESOL") so aparece no cabecalho da 1a pagina do PDF - as
  paginas seguintes repetem so o titulo "EXTRATO CONSOLIDADO DE CONTA
  CORRENTE", sem a palavra "CRESOL". Como `identificar_banco` roda
  PAGINA POR PAGINA (ver `extrair_blocos_por_banco`), uma pagina sem
  nenhum sinal forte de cresol podia cair no fallback frouxo do Bradesco
  (`'bradesco' in texto.lower()`) se tivesse, por exemplo, um pagamento
  de titulo pra "BRADESCO SEGUROS" (nome de terceiro) - isso quebrava o
  extrato em DOIS blocos no meio de um unico banco (paginas 2+ processadas
  como se fossem Bradesco, perdendo a maioria dos lancamentos: 45 de 113
  na amostra real). Corrigido adicionando `'EXTRATO CONSOLIDADO DE CONTA
  CORRENTE' in t.upper()` como mais um sinal do detector `cresol` em
  `DETECTORES` (converter.py) - essa frase aparece em TODAS as paginas
  desse tipo de extrato Cresol (formatos 2 e 3), entao agora toda pagina
  e identificada como cresol diretamente, sem depender de heranca de
  pagina anterior nem competir com o fallback do Bradesco. Se aparecer
  um bug parecido com outro banco no futuro, suspeitar do mesmo padrao:
  fallback frouxo tipo `'nome_banco' in texto.lower()` pegando mencao a
  esse banco como texto livre dentro da descricao de uma transacao de
  outro banco - E que o detector do banco correto so bate em ALGUMAS
  paginas do extrato (nao em todas), pela combinacao dos dois.
- Ailos / ViaCredi (mesmo sistema, cooperativa aparece no cabecalho)
- Civia (v1.7) - cooperativa do Sistema Ailos (COMPE 085, mesmo do Ailos/
  ViaCredi - ver `ofx_export.py`), mas com layout de relatorio proprio
  (gerado pelo sistema "Schema", `parse_civia` separado, nao reusa
  `parse_ailos`). A palavra "Civia" NUNCA aparece no texto do extrato -
  o fingerprint em `DETECTORES` (converter.py) usa a combinacao
  `'Conta/dv:' in t and 'Finalidade da Conta' in t`, especifica desse
  layout. Cada lancamento vem numa unica linha "DD/MM/AAAA Historico
  Documento D/C Valor [Saldo]" - Saldo so aparece esporadicamente (nao
  serve pra validar linha a linha, so usado pra reconciliar o total).
  "SALDO ANTERIOR" tem formato diferente (sem D/C) e e ignorado.
  Validado com extrato real: 94 lancamentos, saldo anterior (776,47) +
  creditos - debitos bateu exato com o ultimo saldo do extrato (2.088,18).
  IMPORTANTE: o PDF de amostra tinha a estrutura interna malformada
  (faltava o trailer/xref - erro "No /Root object" no pdfplumber), embora
  abrisse normal em navegador. Isso motivou adicionar um FALLBACK em
  `converter.py` (`_paginas_texto`): se `pdfplumber.open` falhar ao abrir
  o arquivo inteiro (excecao, nao so pagina vazia), tenta de novo com
  PyMuPDF (`fitz`), que reconstroi a xref na hora e e bem mais tolerante
  a PDF malformado - beneficia qualquer banco, nao so o Civia. Limitacao:
  no modo fallback nao tem como aplicar o fix de texto embaralhado do
  Cresol (que depende de `page.chars`, exclusivo do pdfplumber) - nunca
  precisamos dos dois ao mesmo tempo ate agora.
- Banco do Brasil - TRES formatos, `parse_bb` detecta qual e:
  1. Formato atual (com colunas "Ag. origem"/"Lote" no cabecalho) - excluir
     linhas "BB Rende Facil", mesma logica do Itau.
  2. Formato de 2016 (`_parse_bb_2016`), sem "Ag. origem"/"Lote". Saldo so
     aparece na ULTIMA linha de um grupo de lancamentos do mesmo dia
     (todas as linhas anteriores do grupo ficam sem saldo), e "Documento"
     se distingue de "Valor" por nao ter virgula decimal. Descricao pode
     ter uma 2a linha de continuacao (ex: "SEFAZ RECURSOS ORDINARIOS")
     sem data na frente - cuidado pra nao confundir com a linha de saldo
     final "S A L D O" (tambem sem "documento" antes do valor, mas
     comeca com data, entao nao deve ser tratada como continuacao).
  3. Formato "Dia Lote Documento" (`_parse_bb_dia_lote`), identificado
     pelo cabecalho "Dia Lote Documento Historico Valor". Aqui nao tem
     coluna C/D separada - o sinal vem no final da linha, tipo
     "1.234,56 (+)" ou "98,00 (-)". Estrutura de cada lancamento e uma
     linha de "categoria" (ex: "Pix - Enviado"), seguida de uma linha
     com data+lote+historico+valor+sinal, as vezes seguida de mais uma
     linha de continuacao da descricao (quando a linha seguinte nao
     bate com o padrao de data). LIMITACAO CONHECIDA: se dois
     lancamentos vierem colados um no outro sem nada entre eles, o
     parser pode perder um pedaco da descricao de um deles - mas o
     valor e a data sempre ficam corretos (o que importa pra bater o
     saldo). Linhas "SALDO ANTERIOR" / "SALDO DO DIA" sao ignoradas.
     Validado com extrato real: 126 lancamentos, 0 avisos, saldo bateu
     exato (3.673,73 + creditos - debitos = 3.387,63).

     VARIANTE do formato 3 (`_parse_bb_dia_lote_data_separada`): o mesmo
     cabecalho "Dia Lote Documento", mas o pdfplumber as vezes extrai a
     DATA sozinha numa linha propria (ou "DD/MM/AAAA categoria" junto),
     SEM repetir a data na linha de lote/documento/valor - diferente do
     formato 3 padrao, onde data e valor vem na mesma linha. Rodava e
     nao achava NENHUMA transacao (nem aviso, nem erro - so silencio) no
     formato 3 padrao. Tratado como fallback dentro de `parse_bb`: tenta
     `_parse_bb_dia_lote` primeiro (intocado), e SO chama a variante se
     vier vazio - nunca roda no lugar do parser original ja validado.
     LIMITACAO CONHECIDA: linhas de texto livre entre uma transacao e a
     proxima (continuacao da anterior + categoria da proxima, sem
     separador confiavel) sao concatenadas no historico da transacao
     seguinte - descricao pode sair meio misturada, mas data/valor/tipo
     sempre corretos. Validado com extrato real: 11 lancamentos, saldo
     bateu exato (1.812,75 + creditos - debitos = 386,36).
  ARMADILHA: os formatos 1 e 2 tem "Dt. movimento" E "Dt. balancete" no
  cabecalho (so a ordem dos dois muda) - NAO da pra distinguir por isso
  sozinho (peguei uma regressao no formato atual testando so com isso).
  O sinal confiavel e a AUSENCIA de "Ag. origem" pra identificar o
  formato de 2016. O formato 3 e identificado antes dos outros dois,
  pela presenca de "Dia Lote Documento" (string exclusiva dele).
- Bradesco (cuidado: o layout dos primeiros lancamentos de credito difere
  do layout dos lancamentos de debito dentro do MESMO extrato - ver
  `parse_bradesco` pra entender a logica de linha "tipo" + linha de
  valor + continuacao)
- Sicredi - DOIS layouts, `parse_sicredi` detecta qual e:
  1. `_parse_sicredi_valor_unico` (o de cima) - uma unica coluna de Valor
     com sinal (positivo=credito, negativo=debito), seguida da coluna
     Saldo, tudo na mesma linha - da pra extrair so com regex de texto.
  2. `_parse_sicredi_colunas` (v1.9) - visto noutra cooperativa (CCPI DA
     REGIAO ALTOS DA SERRA). Colunas SEPARADAS de Debito/Credito/Saldo -
     o texto corrido sozinho e AMBIGUO aqui: um lancamento pode ter so
     UM numero no final da linha, sem dar pra saber se e Debito, Credito
     ou Saldo so pelo texto (todos no formato "1.234,56"). Resolvido com
     uma reconstrucao por POSICAO em `converter.py`
     (`_texto_sicredi_colunas_por_pagina`): usa a posicao x de cada
     palavra (via `page.extract_words()`) pra decidir em que coluna um
     valor cai, e marca o resultado num texto intermediario com um
     sufixo sem ambiguidade (`@@D@@1.234,56` ou `@@C@@1.234,56`) que o
     parser em bancos.py so precisa ler com um regex simples.
     ARMADILHA no cabecalho: as palavras "DEBITO"/"CREDITO"/"SALDO" podem
     aparecer de novo dentro do HISTORICO de alguma transacao (ex:
     historico real "DEBITO T.E.D.") - pegar a primeira ocorrencia de
     cada palavra isolada da errado; e preciso achar a UNICA linha onde
     as tres aparecem JUNTAS (mesma posicao vertical) pra ter certeza que
     e o cabecalho de verdade, nao uma coincidencia dentro de um
     lancamento.
     Lancamentos "CAPTACAO APLIC.FINANC.AVISO PREVIO" / "CAPTACAO
     RESG.APLIC.FIN.AVISO PREV" sao a varredura automatica pra uma
     aplicacao financeira (mesmo padrao do "Aplic Aut Mais" do Itau) -
     excluidos, nao contam como movimento real da conta corrente.
     Validado com extrato real (4 paginas, incluindo as linhas de
     CAPTACAO): reconstruindo o saldo linha a linha a partir de "SALDO
     ANTERIOR" e comparando com o "Saldo" que o proprio extrato declara
     em cada linha, 0 divergencias em todo o extrato.
  ARMADILHA de deteccao (bug real corrigido na v1.9): o detector em
  `DETECTORES` (converter.py) so procurava a palavra 'Sicredi' com essa
  capitalizacao exata (ou 'Associado:') - um extrato real veio com o
  rodape todo em CAIXA ALTA ("SICREDI, A VIDA E MELHOR QUANDO E
  COOPERATIVA!") e nao batia. Pior: esse rodape SO aparece na ULTIMA
  pagina do PDF - as paginas com as transacoes de verdade (as primeiras)
  nao tinham NENHUM sinal do nome do banco nelas, entao mesmo corrigindo
  a capitalizacao, `extrair_blocos_por_banco` ia quebrar o PDF em dois
  blocos errados (paginas 1-3 como "banco nao identificado", so a pagina
  4 como sicredi). Corrigido usando os proprios marcadores @@D@@/@@C@@
  como sinal adicional do detector - eles so existem quando
  `_texto_sicredi_colunas_por_pagina` ja confirmou que e esse layout
  especifico, entao servem como fingerprint seguro mesmo sem a palavra
  "sicredi" aparecer na pagina.

### Funciona mas precisa de revisao humana ocasional
- Santander: o texto extraido do PDF NAO tem coluna ou sinal confiavel de
  credito x debito (o layout usa duas colunas visuais que se perdem na
  extracao de texto simples). A classificacao e feita por palavra-chave da
  descricao (RECEBIDO/CREDITO = credito, PAGAMENTO/DEBITO/TARIFA = debito).
  Linhas incertas vem marcadas `[A VERIFICAR]` na coluna Observacao.

### Sem nenhuma amostra ainda (dos 14 bancos da carteira)
- Banrisul
- Sulcredi
- Votorantim

### Extrato digitalizado (PDF escaneado, sem texto) - ver secao "Conversor OCR"
O Conversor_Extratos PRINCIPAL continua sem suportar isso (decisao
mantida - ver por que na secao "Conversor OCR" mais abaixo). Quando um
PDF assim e enviado, o programa detecta (nenhuma pagina do PDF tem
texto extraivel - `page.extract_text()` vazio em todas) e gera um aviso
especifico na aba Pendencias: "Esse PDF parece ser uma imagem escaneada...".
Antes disso (ate a v1.3) esse caso caia no aviso generico de "banco nao
identificado", que confundia (parecia layout novo, nao arquivo sem texto).
A partir de 2026-09-14 existe uma FERRAMENTA SEPARADA (`Conversor_OCR`)
pra esse caso - ver secao propria mais abaixo.

## Proximo passo

Quando o usuario mandar uma amostra completa (nao cortada) de algum desses
4 bancos, seguir o mesmo processo: extrair texto bruto com pdfplumber,
identificar o padrao de linha pra data/descricao/valor/tipo, escrever
`parse_<banco>` em `bancos.py`, adicionar o detector em `DETECTORES` no
`converter.py`, testar e validar contra os totais do extrato.

## Conversor OCR (ferramenta separada) - extrato digitalizado

Historico: em 2026-09-11 o Jean pediu suporte a extrato digitalizado
(foto/print do app do banco, sem texto no PDF) e a decisao foi NAO
implementar - risco alto demais pra ferramenta contabil (ver secao
"Extrato digitalizado" em "Status por banco"). Em 2026-09-14 ele voltou
com um caso real (pastor mandou um extrato assim, gerado direto do app
da Caixa) e pediu pra resolver "nem que seja uma versao separada do
sistema". Decisao final: SIM, mas como ferramenta separada
(`Conversor_OCR`), nunca misturada com o `Conversor_Extratos` principal
- assim o risco de OCR fica isolado, e o conversor principal (confiavel,
le texto nativo) continua intocado.

### Por que e uma ferramenta a parte (nao um modo dentro do principal)
- OCR tem uma taxa de erro real e residual que NAO da pra zerar por mais
  que se ajuste parametro (testado: DPI, pre-processamento de contraste,
  upscale de imagem - a melhor combinacao achada ainda deixa ~5% das
  linhas com o sinal credito/debito indeterminado, num extrato real de
  teste). Misturar isso no conversor principal arriscaria contaminar a
  confianca de quem usa o fluxo normal (texto nativo, bem mais preciso).
- O motor de OCR (Tesseract) e pesado (~70MB de binarios+dados so pro
  essencial) - decisao do Jean foi embutir tudo num .exe so mesmo assim
  (aceitando o .exe ficar grande, ~120MB), em vez de pedir instalacao
  separada - ver [[feedback-console-vs-gui]] sobre o risco de antivirus
  com pacotes grandes/muitos arquivos; aqui o risco existe mas foi aceito
  conscientemente.

### Arquitetura (pasta `OCR/`, isolada do resto do projeto)
- `OCR/conversor_ocr.py`: script principal, mesmo padrao de UX do
  console principal (mensagem de cabecalho, aceita PDFs arrastados ou
  processa tudo da pasta, "Pressione Enter pra sair"). Reusa
  `gerar_excel`/`NOMES_BANCO` de `converter.py` e `gerar_ofx` de
  `ofx_export.py` via import (sys.path aponta pra pasta pai) - NUNCA
  duplica nem altera essas funcoes.
- `OCR/bancos_ocr.py`: parsers de OCR, um por layout de app bancario
  (so tem `parse_caixa_app_ocr` por enquanto). Modulo SEPARADO do
  `bancos.py` principal (nunca importa nem altera ele) porque trabalha
  com texto reconhecido por OCR (sujeito a erro), nao com texto extraido
  direto do PDF.
- `OCR/vendor/tesseract/`: copia local dos binarios do Tesseract
  (tesseract.exe + DLLs + tessdata/por.traineddata) usada pra empacotar
  o .exe. NAO fica no git (`.gitignore`: `OCR/vendor/`) - pra reproduzir:
  1. Baixar e instalar o Tesseract-OCR pra Windows (build UB-Mannheim,
     https://github.com/UB-Mannheim/tesseract/wiki).
  2. Copiar `tesseract.exe` + todos os `*.dll` de
     `C:\Program Files\Tesseract-OCR\` pra `OCR/vendor/tesseract/`.
  3. Copiar `tessdata\por.traineddata` pra
     `OCR/vendor/tesseract/tessdata/` (so o portugues - nao precisa dos
     outros ~100 idiomas nem do `osd.traineddata`, que sozinho tem 10MB).

### Como o OCR e feito (tecnica - importante pra novos parsers)
1. `PyMuPDF` (`fitz`) rasteriza cada pagina do PDF em imagem (300 DPI).
2. A imagem e ampliada 2x (`Image.LANCZOS`) ANTES do OCR - melhora
   bastante o reconhecimento de caracteres pequenos tipo o "C"/"D" colado
   no valor (testado: reduziu linhas incertas de ~11% pra ~5%, mas
   sozinho nao resolveu tudo - trade-off real, nao solucao magica).
3. **NUNCA usar `pytesseract.image_to_string()` (texto corrido) pra
   tabela.** O Tesseract as vezes le a tabela COLUNA POR COLUNA (todas
   as datas juntas, depois todos os valores juntos) em vez de linha por
   linha - embaralha completamente a correspondencia entre campos, e
   pior, isso pode mudar ENTRE PAGINAS DO MESMO PDF (visto no extrato
   real: paginas 1-3 vieram em blocos de coluna, pagina 4 veio linha por
   linha corrida). A solucao e usar `pytesseract.image_to_data()` (retorna
   a posicao x/y de cada palavra) e remontar as linhas pela posicao
   (agrupar por proximidade vertical, ordenar por x dentro do grupo) -
   ver `linhas_por_posicao()` em `bancos_ocr.py`. Mesmo principio do
   `_texto_robusto_por_caracteres()` no `converter.py` (bug de texto
   embaralhado do Cresol), so que a fonte aqui e OCR, nao pdfplumber.
4. Cada parser de banco (ex: `parse_caixa_app_ocr`) trabalha em cima
   dessas linhas ja remontadas, com um regex que casa
   `data - hora doc historico [ruido] valor sinal saldo sinal` - o
   "ruido" no meio (favorecido, CPF/CNPJ) e DELIBERADAMENTE IGNORADO
   (decisao do Jean: nao vale o risco de desalinhar e colar o nome
   errado numa transacao errada). Historico usa uma lista FECHADA de
   categorias conhecidas (vocabulario observado no extrato real testado)
   - historico fora da lista e descartado, nunca adivinhado.
5. **Inferencia de sinal por delta de saldo**: quando o OCR nao consegue
   ler o caractere C/D com confianca, o parser compara o saldo da linha
   com o saldo da linha anterior na sequencia - se o delta bater com o
   valor da transacao (tolerancia 0,02), usa o sinal do delta. Quando
   NEM o OCR nem o delta confirmam, a transacao fica com tipo '?' - nunca
   inventa um sinal.

### Como o resultado sai marcado (nunca finge ser confiavel)
- TODA transacao sai com a coluna Observacao preenchida "Gerado por OCR -
  confira contra o extrato original antes de usar."
- Transacoes com sinal indeterminado (tipo='?') saem com
  "[A VERIFICAR - OCR]" na Observacao E com a LINHA INTEIRA destacada em
  vermelho claro (`_destacar_linhas_incertas()` em `conversor_ocr.py` -
  pos-processamento com openpyxl por cima do `gerar_excel` generico, sem
  alterar ele).
- Transacoes incertas NAO entram no `.ofx` - `gerar_ofx` trata qualquer
  coisa != 'C' como debito (ver `ofx_export.py` linha ~88), entao passar
  tipo='?' pra ele esconderia a incerteza dentro do arquivo que alimenta
  o sistema contabil. O `conversor_ocr.py` filtra
  (`transacoes_confiaveis = [t for t in transacoes if not t['incerta']]`)
  antes de chamar `gerar_ofx`.

### Validacao feita (extrato real, app da Caixa, 4 paginas)
64 transacoes reconhecidas na 1a tentativa (DPI 300, sem upscale): 7
incertas (~11%), saldo NAO bateu (calculado R$ 7.274,88 vs real
R$ 8.938,58). Com upscale 2x: 58 transacoes (perdeu 6 - trade-off real,
upscale melhora uns casos e piora outros), 3 incertas (~5%), saldo ainda
nao bateu exato mas ficou bem mais perto. Conclusao: aceito como
"rascunho pra conferir", NUNCA como resultado pronto pra usar - e assim
que a ferramenta se apresenta (LEIA-ME e mensagens do console deixam
isso explicito).

### Escopo de bancos suportados
So o extrato do APP da Caixa por enquanto (unico com amostra real
testada). Regra de ouro reforcada aqui: NUNCA escrever parser de OCR
pra um banco sem amostra real em maos - seria adivinhar o layout do app
daquele banco, e o proprio caso da Caixa mostrou que OCR erra mesmo com
amostra real testada; sem teste, o risco e maior ainda. Quando aparecer
extrato digitalizado de outro banco, tratar como novo layout: escrever
um novo `parse_<banco>_app_ocr` em `bancos_ocr.py`, adicionar o
fingerprint em `DETECTORES_OCR`, validar batendo saldo, documentar aqui.

### Distribuicao
Mesmo fluxo do `Conversor_Extratos` (link estavel + releases arquivadas
por versao - ver "Verificacao automatica de atualizacao"), mas como
asset SEPARADO no mesmo repositorio (`Conversor_OCR.zip`), com seu
proprio numero de versao. Pasta `Para_Equipe_OCR/` (analoga a
`Para_Equipe/`) com `Conversor_OCR.exe` + `LEIA-ME.txt` proprio
(bem mais extenso em avisos que o do conversor principal) + `LICENSE.txt`.

## Interface grafica (GUI) - em teste

Motivo: a tela estilo DOS do `.exe` console estava assustando usuarios
menos tecnicos (relatado pelo Jean, sessao 2026-09-08). Escolhida a opcao
`pywebview` (janela nativa reaproveitando o mesmo visual "ledger" usado em
`docs/index.html`) em vez de `tkinter`, por ficar bem mais bonita sem
precisar reescrever o design.

- `gui.py`: janela pywebview. Classe `Api` expoe `escolher_pdfs()`
  (dialogo nativo de arquivo) e `converter(caminhos)` (chama o mesmo
  `processar_pdf` do `converter.py` - a logica de conversao nao mudou,
  so a "casca") pro JS via `window.pywebview.api.*`.
- `gui.html`: interface (dropzone, lista de arquivos, botao Converter,
  lista de resultados com pill verde/vermelho). Mesma paleta de cores do
  `docs/index.html`, mas layout de ferramenta (nao de pagina de
  apresentacao) - sem tema escuro de proposito, e um app so, nao artifact.

**ARMADILHA IMPORTANTE - pywebview 6.x trava a janela**: a versao 6.2.1
(a mais recente no momento) tem um bug real (nao so log poluido) na
integracao com WebView2 nesta maquina - ao abrir a janela, tenta
enumerar `window.native.AccessibilityObject...` por reflexao e entra em
recursao infinita, deixando a janela com "Nao esta respondendo". A
correcao foi fixar a versao: `pip install pywebview==4.4.1` (sem esse
recurso experimental `window.native`, mais estavel). **Nunca fazer
`pip install --upgrade pywebview` sem testar antes que a janela abre e
continua respondendo (`Get-Process ... | Select Responding` no
PowerShell) por pelo menos uns 10 segundos** - o travamento nao aparece
imediato no log, só quando o processo trava mesmo.

Drag-and-drop de arquivo pra DENTRO da janela ja aberta (`file.pywebviewFullPath`
no JS) testado e confirmado que NAO funciona nesta maquina/versao do
WebView2 Runtime - cai no aviso "Nao consegui ler o caminho do arquivo
arrastado" (comportamento esperado, ver `dropzone.addEventListener('drop', ...)`
em `gui.html`). Por isso tem dois caminhos que SAO garantidos:
1. Botao "clique pra escolher" (dialogo nativo, `Api.escolher_pdfs`).
2. Arrastar o(s) PDF(s) em cima do ICONE do `.exe` no Explorer (nao pra
   dentro da janela) - o Windows chama o programa com os caminhos em
   `sys.argv`, igual o `.exe` de console ja fazia. `Api.arquivos_iniciais()`
   devolve esses caminhos pro JS assim que a janela carrega (evento
   `pywebviewready`), pré-preenchendo a lista. Os dois caminhos foram
   clicados e confirmados funcionando (nao so testado por logica).

**ARMADILHA (RESOLVIDA) - primeira execucao do .exe empacotado demorava
MUITO (10 a 30+ segundos) com a tela em branco**: onefile extrai tudo de
novo pra uma pasta temporaria TODA vez que abre (+ antivirus escaneando
os arquivos recem-extraidos de novo a cada vez), diferente de rodar
`python gui.py` direto (~7s, sem essa extracao). Corrigido com duas
mudancas, ambas testadas com controle de tela real (nao so por logica):
1. Build trocado de `--onefile` pra `--onedir` (pasta com o .exe e os
   DLLs ja extraidos, sem extracao repetida a cada abertura) - eliminou a
   demora na pratica, abriu rapido nas duas execucoes testadas.
2. `background_color='#EEF2E9'` no `webview.create_window()` (mesma cor
   de fundo do app) - pywebview usa branco puro por padrao, o que fazia
   qualquer atraso remanescente parecer tela travada. Sem isso o
   parametro default e `#FFFFFF`.

Build agora: `python -m PyInstaller --onedir --console --add-data
"gui.html;." gui.py` (gera uma PASTA em vez de um arquivo so - precisa
zipar junto com LEIA-ME.txt e LICENSE.txt igual o `.exe` de console,
nao da mais pra mandar um arquivo unico). Manter `--console` ate decidir
junto com o Jean se troca pra `--windowed`.

Status: testado e validado pelo Claude com controle de tela real (nao so
por logica) - janela renderiza, os dois fluxos de selecao de arquivo
funcionam, conversao bate com os totais corretos (testado com Sicoob 310,
Caixa 597, Itau 104 lancamentos, todos conferem). Falta o Jean decidir se
quer substituir o .exe de console pela GUI na distribuicao oficial (`Para_Equipe`,
release do GitHub) ou oferecer as duas opcoes.

## Preferencias do usuario (aplicam a este projeto)

- Portugues do Brasil, padrao numerico brasileiro (milhar com ponto,
  decimal com virgula).
- Nao usar travessao como pontuacao.
- Nao reconstruir de memoria um layout sem amostra real - pedir o PDF.
- Sempre confirmar contagem de lancamentos e totais batendo com o extrato
  antes de considerar um parser fechado.
