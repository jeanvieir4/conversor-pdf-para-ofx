# -*- coding: utf-8 -*-
"""
Conversor PDF para OFX - Jean Vieira
Uso: python converter.py caminho_do_extrato.pdf [outro.pdf ...]
Gera um .ofx por banco encontrado no PDF (ou um .txt explicando o problema, se nao gerar):
"""
import sys
import os
import re
import glob
import threading
import urllib.request
import atualizador
import relatorio
import pdfplumber
import fitz
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bancos import BANK_PARSERS
from ofx_export import gerar_ofx

# Sempre que gerar uma nova versao pra distribuir: atualizar esse numero
# E o conteudo do arquivo VERSION no repositorio (mesmo numero nos dois).
# So assim quem ja tem uma versao antiga instalada fica sabendo que saiu
# uma nova - ver "_verificar_atualizacao" mais abaixo e o CONTEXTO_PROJETO.md.
VERSAO_ATUAL = '3.1'
_REPO_GITHUB = 'jeanvieir4/conversor-pdf-para-ofx'
_URL_VERSION = f'https://raw.githubusercontent.com/{_REPO_GITHUB}/main/VERSION'
_URL_DOWNLOAD = f'https://github.com/{_REPO_GITHUB}/releases/download/1.0/Conversor_Extratos.zip'
_URL_HASH = f'https://raw.githubusercontent.com/{_REPO_GITHUB}/main/SHA256_EXE'
_NOME_EXE = 'Conversor_Extratos.exe'
# Preenchido pela verificacao de versao (ver __main__); usado nos relatorios .txt.
_INFO_UPDATE = {}

# Observacao: nem todo banco escreve o proprio nome como texto selecionavel no PDF
# (varios usam o nome so na logo, que e imagem). Por isso alguns detectores usam
# combinacoes de campos do cabecalho/rodape em vez do nome do banco.
DETECTORES = [
    ('sicoob',    lambda t: 'SICOOB' in t and 'SISTEMA DE COOPERATIVAS' in t),
    ('caixa',     lambda t: ('GERENCIADOR' in t.upper() and 'CAIXA' in t.upper()) or 'Extrato no per\u00edodo de' in t),
    # cresol vem antes do bradesco: o fallback frouxo do bradesco ('bradesco' in
    # texto) da falso positivo quando um extrato de OUTRO banco tem um boleto
    # pago pra "Bradesco Seguros" ou similar (nome de terceiro, nao do banco).
    # Isso e especialmente critico pro Cresol porque o nome do banco so
    # aparece no cabecalho da 1a pagina - as paginas seguintes (que repetem
    # so "EXTRATO CONSOLIDADO DE CONTA CORRENTE") ficavam sem nenhum sinal
    # forte de cresol e podiam cair de vez no fallback frouxo do bradesco
    # (ex.: um pagamento de titulo pra "BRADESCO SEGUROS" na pagina 2),
    # quebrando o extrato em dois blocos no meio de um unico banco. Por
    # isso o titulo da pagina tambem conta como sinal de cresol aqui.
    ('cresol',    lambda t: 'Consulta Posi\u00e7\u00e3o consolidada' in t or 'CRESOL' in t.upper()
                              or 'sistema.confesol' in t.lower()
                              or 'EXTRATO CONSOLIDADO DE CONTA CORRENTE' in t.upper()),
    ('bradesco',  lambda t: ('Lan\u00e7amento' in t and 'Dcto.' in t and 'D\u00e9bito' in t)
                              or 'Nome do usu\u00e1rio:' in t or ('bradesco' in t.lower())),
    ('itau',      lambda t: 'extrato mensal' in t.lower() and ('ita\u00fa' in t.lower() or 'B001A' in t)),
    ('unicred',   lambda t: 'CENTRAL DE RELACIONAMENTO' in t or ('Coop:' in t and 'AG:' in t and 'Conta:' in t)),
    ('ailos',     lambda t: 'AILOS' in t.upper() or 'VIACREDIALTOVALE' in t.upper() or 'VIACREDI' in t.upper()),
    ('bb',        lambda t: 'BB Rende F' in t or ('Ag. origem' in t and 'Lote' in t) or 'Dt. balancete' in t
                              or 'Dia Lote Documento' in t),
    # sicredi vem antes do santander: os dois tem fallback frouxo
    # ('nome_banco' in texto.lower()) que da falso positivo quando um
    # extrato de OUTRO banco tem um boleto ou transferencia mencionando
    # o nome do outro banco como texto livre (mesmo padrao ja visto entre
    # cresol/bradesco). Bug real: um extrato do Sicredi tinha uma linha
    # "LIQUIDACAO BOLETO ... SANTANDER SANTA..." (nome do beneficiario, nao
    # o banco do extrato) e a pagina inteira - com as transacoes de
    # verdade - ia pro detector errado (santander), sobrando so a ultima
    # pagina (rodape) como sicredi.
    # 'SICREDI' em maiusculas (nao so 'Sicredi') porque um extrato real veio
    # com o rodape todo em caixa alta ("SICREDI, A VIDA E MELHOR QUANDO E
    # COOPERATIVA!"), e esse rodape SO aparece na ULTIMA pagina do PDF -
    # as paginas com as transacoes de verdade (as primeiras) nao tem
    # nenhum sinal do nome do banco nelas. Por isso tambem conta como
    # sicredi quando a pagina foi reconstruida pelo layout de colunas
    # Debito/Credito/Saldo (`_texto_sicredi_colunas_por_pagina` ja so
    # ativa pra esse layout especifico, entao os marcadores @@D@@/@@C@@
    # sao um sinal seguro por si so, mesmo sem a palavra "sicredi" na
    # pagina). Bug real reportado pelo Jean - extrato inteiro nao batia
    # com nenhum detector.
    ('sicredi',   lambda t: 'sicredi' in t.lower() or 'Associado:' in t
                              or '@@D@@' in t or '@@C@@' in t),
    ('santander', lambda t: 'santander' in t.lower() or 'Extrato_PJ_A4' in t or 'BALP_UY' in t),
    # Civia nao escreve o proprio nome em lugar nenhum do texto (extrato
    # gerado pelo sistema "Schema", usado por varias cooperativas/bancos
    # pequenos - visto no metadado 'author: schemaprd' do PDF, mas isso
    # nao aparece no texto das paginas). Fingerprint por combinacao de
    # campos do cabecalho, que sao bem especificos desse layout.
    ('civia',     lambda t: 'Conta/dv:' in t and 'Finalidade da Conta' in t),
    # Itau - relatorio "Entradas e Saidas" (Itau Empresas, v2.1): a palavra
    # "Itau" so aparece no rodape da ULTIMA pagina, e o cabecalho so na
    # primeira - detecta pela combinacao de campos do cabecalho (pagina 1);
    # as demais paginas herdam o banco da anterior (ver extrair_blocos_por_banco).
    ('itau',      lambda t: 'Lançamentos do período' in t and 'Razão Social' in t
                              and 'SALDO TOTAL DISPON' in t),
    # Civia - 'EXTRATO ESPECIAL' (v2.2): outro cabecalho do Civia, ver parse_civia.
    ('civia',     lambda t: 'CIVIA - EXTRATO ESPECIAL' in t and 'CONTA/DV:' in t),
]


def identificar_banco(texto):
    for nome, teste in DETECTORES:
        try:
            if teste(texto):
                return nome
        except Exception:
            continue
    return None


def _texto_robusto_por_caracteres(page):
    """Fallback de extracao: reconstroi as linhas agrupando os caracteres
    pela posicao vertical, na ordem em que aparecem no fluxo do PDF (em vez
    de deixar o algoritmo de layout do pdfplumber decidir a ordem). Usado
    quando o extract_text() padrao embaralha a pagina - visto em extratos
    gerados por "impressao" de sistema web (ex: Cresol / sistema Colmeia),
    onde os caracteres tem posicoes que confundem o agrupador padrao."""
    linhas = []
    atual = []
    top_atual = None
    for c in page.chars:
        t = round(c['top'], 1)
        if top_atual is None or abs(t - top_atual) < 1.5:
            atual.append(c['text'])
            if top_atual is None:
                top_atual = t
        else:
            linhas.append(''.join(atual))
            atual = [c['text']]
            top_atual = t
    if atual:
        linhas.append(''.join(atual))
    return '\n'.join(linhas)


# Fingerprint de paginas com texto embaralhado pelo extract_text() padrao.
# A URL do sistema sobrevive ao embaralhamento (fica intacta), por isso da
# pra usar ela pra decidir quando trocar pelo fallback por caracteres.
_FINGERPRINTS_TEXTO_EMBARALHADO = ('sistema.confesol/colmeia',)

_MARCADOR_DEBITO = '@@D@@'
_MARCADOR_CREDITO = '@@C@@'


def _texto_sicredi_colunas_por_pagina(page):
    """Reconstroi as linhas de um extrato Sicredi com colunas SEPARADAS de
    Debito/Credito/Saldo (em vez de uma coluna so de Valor com sinal ou
    letra C/D) usando a posicao (x) de cada palavra pra decidir em que
    coluna um valor cai. O texto corrido sozinho e ambiguo aqui: um
    lancamento pode ter so UM numero no fim da linha, e sem saber a
    posicao nao da pra saber se aquele numero e Debito, Credito ou
    Saldo (todos com o mesmo formato "1.234,56"). Marca o valor
    encontrado com um sufixo (@@D@@valor ou @@C@@valor) que
    `_parse_sicredi_colunas` em bancos.py le sem ambiguidade nenhuma. So
    ativa quando a pagina tem esse layout especifico (cabecalho com
    "DEBITO", "CREDITO" e "SALDO" como palavras separadas - o outro
    layout do Sicredi ja suportado, com Valor+sinal numa coluna so, nao
    tem esses 3 cabecalhos)."""
    palavras = page.extract_words()
    # "DEBITO"/"CREDITO"/"SALDO" podem aparecer de novo dentro do HISTORICO
    # de alguma transacao (ex: historico "DEBITO T.E.D.") - por isso nao da
    # pra so pegar a primeira ocorrencia de cada palavra isolada. O
    # cabecalho de verdade e a UNICA linha onde as tres aparecem juntas.
    por_linha = {}
    for p in palavras:
        if p['text'] in ('DEBITO', 'CREDITO', 'SALDO'):
            por_linha.setdefault(round(p['top'], 1), {})[p['text']] = p
    linha_cabecalho = next((v for v in por_linha.values() if len(v) == 3), None)
    if linha_cabecalho is None:
        return None

    x_debito = linha_cabecalho['DEBITO']['x0']
    x_credito = linha_cabecalho['CREDITO']['x0']
    x_saldo = linha_cabecalho['SALDO']['x0']
    limite_debito_credito = (x_debito + x_credito) / 2
    limite_credito_saldo = (x_credito + x_saldo) / 2
    padrao_valor = re.compile(r'^-?[\d.]+,\d{2}$')

    linhas_por_top = {}
    for p in palavras:
        linhas_por_top.setdefault(round(p['top'], 1), []).append(p)

    linhas_texto = []
    for top in sorted(linhas_por_top.keys()):
        ps = sorted(linhas_por_top[top], key=lambda p: p['x0'])
        partes_historico = []
        valor_debito = None
        valor_credito = None
        for p in ps:
            # so trata como valor de coluna se estiver bem depois de onde
            # a coluna Debito comeca - assim um numero de documento que por
            # coincidencia tenha formato "123,45" (raro) dentro do texto
            # normal nao e confundido com um lancamento
            if padrao_valor.match(p['text']) and p['x0'] >= x_debito - 20:
                if p['x0'] < limite_debito_credito:
                    valor_debito = p['text']
                elif p['x0'] < limite_credito_saldo:
                    valor_credito = p['text']
                # senao e Saldo - nao usado pro lancamento, so ignora
            else:
                partes_historico.append(p['text'])
        linha = ' '.join(partes_historico)
        if valor_debito:
            linha += f' {_MARCADOR_DEBITO}{valor_debito}'
        if valor_credito:
            linha += f' {_MARCADOR_CREDITO}{valor_credito}'
        linhas_texto.append(linha)
    return '\n'.join(linhas_texto)


def _paginas_texto(caminho_pdf):
    """Extrai o texto de cada pagina do PDF. Tenta pdfplumber primeiro (e
    o unico que suporta o fallback de texto embaralhado acima, que depende
    de page.chars). Se o arquivo tiver a estrutura interna malformada (ex:
    trailer/xref incompletos - visto num extrato real que abria normal em
    navegador mas o pdfplumber recusava com "No /Root object") e
    pdfplumber nem conseguir ABRIR o arquivo, cai pro PyMuPDF (fitz), que
    e bem mais tolerante a esse tipo de problema (os leitores de PDF em
    geral reconstroem a xref table na hora quando ela falta ou esta
    quebrada - fitz faz isso, pdfminer/pdfplumber nao). Nesse caso o
    fallback de texto embaralhado fica indisponivel (nunca precisamos dos
    dois ao mesmo tempo)."""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            paginas = []
            for page in pdf.pages:
                texto = page.extract_text() or ''
                if any(fp in texto.lower() for fp in _FINGERPRINTS_TEXTO_EMBARALHADO):
                    texto = _texto_robusto_por_caracteres(page)
                elif 'DEBITO' in texto and 'CREDITO' in texto and 'SALDO' in texto:
                    texto_colunas = _texto_sicredi_colunas_por_pagina(page)
                    if texto_colunas is not None:
                        texto = texto_colunas
                paginas.append(texto)
            return paginas
    except Exception:
        doc = fitz.open(caminho_pdf)
        try:
            return [page.get_text() for page in doc]
        finally:
            doc.close()


def extrair_blocos_por_banco(caminho_pdf):
    """Agrupa paginas consecutivas do mesmo banco. Paginas sem cabecalho
    reconhecivel herdam o banco da pagina anterior (paginas de continuacao)."""
    blocos = []  # lista de (nome_banco, texto_concatenado)
    banco_atual = None
    texto_atual = []
    for texto in _paginas_texto(caminho_pdf):
        banco_pagina = identificar_banco(texto)
        if banco_pagina and banco_pagina != banco_atual and texto_atual:
            blocos.append((banco_atual, '\n'.join(texto_atual)))
            texto_atual = []
        if banco_pagina:
            banco_atual = banco_pagina
        texto_atual.append(texto)
    if texto_atual:
        blocos.append((banco_atual, '\n'.join(texto_atual)))
    return blocos


def _formatar_valor_brl(valor):
    """Decimal(1234.56) -> '1.234,56' (padrao numerico brasileiro, com milhar)."""
    inteiro, _, centavos = f"{valor:.2f}".partition('.')
    negativo = inteiro.startswith('-')
    if negativo:
        inteiro = inteiro[1:]
    grupos = []
    while len(inteiro) > 3:
        grupos.insert(0, inteiro[-3:])
        inteiro = inteiro[:-3]
    grupos.insert(0, inteiro)
    return ('-' if negativo else '') + '.'.join(grupos) + ',' + centavos


NOMES_BANCO = {
    'sicoob': 'Sicoob', 'caixa': 'Caixa', 'itau': 'Itau', 'unicred': 'UniCred',
    'cresol': 'Cresol', 'ailos': 'Ailos_ViaCredi', 'bb': 'BancoDoBrasil',
    'santander': 'Santander', 'sicredi': 'Sicredi', 'bradesco': 'Bradesco',
    'civia': 'Civia',
}


# Nome do banco como aparece na tela, e resumo do ultimo PDF processado
# (banco, data inicial, data final, quantidade) - usado so pra mostrar na tela.
_NOME_TELA = {'bb': 'Banco do Brasil', 'ailos': 'Ailos/ViaCredi'}
_RESUMO_ULTIMO = []


def gerar_excel(transacoes_por_banco, avisos, caminho_saida):
    wb = Workbook()
    wb.remove(wb.active)  # a aba ativa padrao e substituida por uma aba por banco

    for banco, transacoes in transacoes_por_banco.items():
        if not transacoes:
            continue
        titulo_aba = NOMES_BANCO.get(banco, banco)[:31]  # limite do Excel p/ nome de aba
        ws = wb.create_sheet(titulo_aba)
        ws.append(['Data', 'Historico', 'Valor', 'Tipo', 'Observacao'])
        for c in ws[1]:
            c.font = Font(bold=True, color='FFFFFF')
            c.fill = PatternFill(start_color='1B4E8C', end_color='1B4E8C', fill_type='solid')

        transacoes_ordenadas = sorted(transacoes, key=lambda t: (t['data'] is None, t['data']))
        for t in transacoes_ordenadas:
            valor_str = _formatar_valor_brl(t['valor']) if t['valor'] is not None else '[A VERIFICAR]'
            ws.append([
                t['data'].strftime('%d/%m/%Y') if t['data'] else '[A VERIFICAR]',
                t['historico'],
                valor_str,
                t['tipo'],
                t.get('obs', ''),
            ])

        larguras = [12, 55, 14, 8, 60]
        for i, w in enumerate(larguras, start=1):
            ws.column_dimensions[chr(64 + i)].width = w

    if avisos:
        ws2 = wb.create_sheet('Pendencias')
        ws2.append(['Origem', 'Aviso'])
        for c in ws2[1]:
            c.font = Font(bold=True, color='FFFFFF')
            c.fill = PatternFill(start_color='1B4E8C', end_color='1B4E8C', fill_type='solid')
        for origem, aviso in avisos:
            ws2.append([NOMES_BANCO.get(origem, origem), aviso])
        ws2.column_dimensions['A'].width = 15
        ws2.column_dimensions['B'].width = 100

    if not wb.sheetnames:
        wb.create_sheet('Extrato')  # evita salvar arquivo sem nenhuma aba

    wb.save(caminho_saida)


def processar_pdf(caminho_pdf, pasta_saida):
    nome_base = os.path.splitext(os.path.basename(caminho_pdf))[0]
    _RESUMO_ULTIMO.clear()
    blocos = extrair_blocos_por_banco(caminho_pdf)
    transacoes_por_banco = {}
    avisos = []
    total = 0
    escaneado = False
    for banco, texto in blocos:
        if banco is None:
            if not texto.strip():
                escaneado = True
            else:
                avisos.append(('(banco nao identificado)',
                                'Um trecho do PDF nao bateu com nenhum layout de banco conhecido e foi ignorado.'))
            continue
        parser = BANK_PARSERS.get(banco)
        if not parser:
            avisos.append((banco, 'Banco identificado mas sem parser implementado ainda.'))
            continue
        transacoes, obs = parser(texto)
        transacoes_por_banco.setdefault(banco, []).extend(transacoes)
        total += len(transacoes)
        if obs:
            avisos.append((banco, obs))

    if escaneado:
        if total == 0:
            total += _ler_por_ocr(caminho_pdf, transacoes_por_banco, avisos)
        else:
            avisos.append(('(sem texto no PDF)',
                            'Parte do PDF e imagem escaneada e foi ignorada (so as paginas com texto foram lidas).'))

    arquivos_ofx = []
    for banco, transacoes in transacoes_por_banco.items():
        titulo_banco = NOMES_BANCO.get(banco, banco)
        caminho_ofx = os.path.join(pasta_saida, f'{nome_base}_{titulo_banco}.ofx')
        if gerar_ofx(transacoes, banco, caminho_ofx):
            arquivos_ofx.append(caminho_ofx)
            datas = [t['data'] for t in transacoes if t.get('data')]
            _RESUMO_ULTIMO.append((_NOME_TELA.get(banco, titulo_banco),
                                   min(datas) if datas else None, max(datas) if datas else None,
                                   len(transacoes)))

    if arquivos_ofx:
        saida = arquivos_ofx[0]
        if avisos:
            # O OFX saiu, mas parte do PDF ficou de fora ou tem ressalva: deixa registrado.
            relatorio.escrever_relatorio(
                caminho_pdf, pasta_saida, '_AVISOS',
                'Avisos sobre a conversão (o arquivo .ofx FOI gerado)',
                'O .ofx foi gerado, mas o programa encontrou o seguinte:\n\n' + _texto_avisos(avisos),
                'Confira se está faltando algum lançamento comparando o .ofx com o extrato original '
                'antes de importar no sistema.',
                _INFO_UPDATE, VERSAO_ATUAL, _URL_DOWNLOAD)
        return saida, total, avisos, arquivos_ofx

    saida = relatorio.escrever_relatorio(
        caminho_pdf, pasta_saida, '_ERRO',
        'NÃO FOI POSSÍVEL GERAR O ARQUIVO .OFX',
        _explicar_falha(avisos, total), _orientar_falha(avisos),
        _INFO_UPDATE, VERSAO_ATUAL, _URL_DOWNLOAD)
    return saida, total, avisos, arquivos_ofx


def _ler_por_ocr(caminho_pdf, transacoes_por_banco, avisos):
    """PDF que e imagem: le por OCR (ocr_motor). Linhas com sinal/valor incerto
    NUNCA entram no OFX - ficam listadas no aviso (-> _AVISOS.txt). Retorna
    quantos lancamentos confiaveis foram adicionados."""
    try:
        import ocr_motor
    except Exception:
        avisos.append(('(sem texto no PDF)',
                        'Este PDF e uma imagem escaneada e o recurso de OCR nao esta disponivel nesta copia do programa.'))
        return 0
    print('      (PDF e imagem: lendo por OCR, pode demorar um pouco...)', flush=True)
    try:
        banco, transacoes, aviso = ocr_motor.ler_pdf_ocr(caminho_pdf)
    except Exception as e:
        avisos.append(('(sem texto no PDF)', f'Este PDF e uma imagem escaneada e a leitura por OCR falhou ({e}).'))
        return 0
    if banco is None:
        avisos.append(('(banco nao identificado - OCR)', 'layout nao reconhecido'))
        return 0
    confiaveis = [t for t in transacoes if not t.get('incerta')]
    incertas = [t for t in transacoes if t.get('incerta')]
    texto = 'LIDO POR OCR (imagem). ' + (aviso or '')
    if incertas:
        texto += ('\nLINHAS QUE FICARAM DE FORA DO .OFX (nao deu pra confirmar valor/sinal pela imagem) '
                  '- lance a mao depois de conferir no extrato:')
        for t in incertas:
            texto += (f"\n    {t['data'].strftime('%d/%m/%Y')}  {t['tipo']}  "
                      f"{str(t['valor']).replace('.', ',')}  {t['historico']}")
    avisos.append((banco, texto))
    transacoes_por_banco.setdefault(banco, []).extend(confiaveis)
    return len(confiaveis)


def _texto_avisos(avisos):
    return '\n'.join(f'- {NOMES_BANCO.get(o, o)}: {a}' for o, a in avisos)


def _explicar_falha(avisos, total):
    origens = [o for o, _ in avisos]
    if '(banco nao identificado - OCR)' in origens:
        return ('Este PDF é uma IMAGEM (extrato escaneado, foto ou print). O programa leu a imagem por OCR, '
                'mas o layout não foi reconhecido. Hoje o OCR entende extratos da Caixa (aplicativo) e do '
                'Banco do Brasil (internet banking e extrato de conta corrente escaneado). Pode ser outro '
                'banco ou um modelo novo.')
    if '(sem texto no PDF)' in origens:
        return ('Este PDF é uma IMAGEM (extrato escaneado, foto ou print) e o programa não conseguiu '
                'lê-lo por OCR.\n\n' + _texto_avisos(avisos))
    if '(banco nao identificado)' in origens and total == 0 and len(origens) == len(
            [x for x in origens if x == '(banco nao identificado)']):
        return ('O programa não reconheceu o banco/layout deste extrato. Pode ser um banco que ainda '
                'não é suportado ou um modelo de extrato novo desse banco. Nenhum lançamento foi lido.')
    if avisos:
        return ('O programa reconheceu o extrato, mas não conseguiu tirar nenhum lançamento dele.\n\n'
                + _texto_avisos(avisos))
    if total > 0:
        return ('Foram lidos lançamentos, mas nenhum tinha data e valor válidos para gravar no .ofx.')
    return 'O programa leu o PDF, mas não encontrou nenhum lançamento nele (o extrato pode estar vazio).'


def _orientar_falha(avisos):
    origens = [o for o, _ in avisos]
    if '(banco nao identificado - OCR)' in origens:
        return ('1) Confira se a imagem está nítida, reta e completa (sem cortes).\n'
                '2) Veja abaixo se existe versão nova do programa e atualize.\n'
                '3) Se persistir, envie este PDF para o Jean (contato no fim deste arquivo) para incluir '
                'suporte a este layout. Só dá para incluir com uma amostra real.')
    if '(sem texto no PDF)' in origens:
        return ('Veja abaixo se existe versão nova do programa e atualize. Se persistir, envie este PDF '
                'para o Jean (contato no fim deste arquivo).')
    return ('1) Confira se o PDF é mesmo um extrato bancário (e não um comprovante ou relatório) e se está '
            'completo.\n2) Veja abaixo se existe versão nova do programa e atualize.\n'
            '3) Se persistir, envie este PDF para o Jean (contato no fim deste arquivo) para incluir '
            'suporte a este layout.')


def relatorio_excecao(caminho_pdf, pasta_saida, erro):
    """Chamado quando o processamento do PDF quebrou de vez (arquivo corrompido,
    protegido por senha, em uso por outro programa, etc.)."""
    return relatorio.escrever_relatorio(
        caminho_pdf, pasta_saida, '_ERRO',
        'NÃO FOI POSSÍVEL GERAR O ARQUIVO .OFX',
        'O programa não conseguiu abrir ou ler este PDF.\n'
        'Causas mais comuns: arquivo corrompido ou incompleto (download interrompido), PDF protegido por '
        'senha, ou arquivo aberto/bloqueado por outro programa.\n\n'
        f'Detalhe técnico (se o Jean pedir): {erro}',
        'Abra o PDF para ver se ele abre normalmente. Se estiver protegido por senha, salve uma cópia sem '
        'senha. Se estiver corrompido, baixe o extrato de novo no banco. Depois tente outra vez.',
        _INFO_UPDATE, VERSAO_ATUAL, _URL_DOWNLOAD)


def _versao_maior(remota, atual):
    """Compara '1.10' > '1.9' corretamente (numero a numero, nao texto)."""
    def normalizar(v):
        partes = []
        for p in v.strip().split('.'):
            if not p.isdigit():
                return None
            partes.append(int(p))
        return tuple(partes) if partes else None
    a, b = normalizar(remota), normalizar(atual)
    if a is None or b is None:
        return False
    return a > b


def _verificar_atualizacao(resultado):
    """Roda em thread separada (ver __main__) pra nao atrasar o programa.
    So consulta um arquivo de texto simples no GitHub (VERSION) - nao
    manda nenhum dado do usuario, so le um numero de volta. Qualquer falha
    (sem internet, GitHub fora do ar, etc.) e ignorada silenciosamente:
    verificar atualizacao nunca pode travar nem atrapalhar a conversao."""
    try:
        with urllib.request.urlopen(_URL_VERSION, timeout=2.5) as resp:
            versao_remota = resp.read().decode('utf-8').strip()
        if _versao_maior(versao_remota, VERSAO_ATUAL):
            resultado['nova_versao'] = versao_remota
    except Exception:
        pass


def _pasta_do_programa():
    """Pasta onde esta o .exe (ou o .py, se rodando via Python direto)."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _mostrar_aviso_atualizacao(resultado_update, thread_update):
    """Espera a verificacao terminar (se ainda nao tiver terminado) so um
    pouco mais - na pratica ja rodou em paralelo com a conversao, entao
    quase sempre essa espera e de 0 segundos."""
    thread_update.join(timeout=2.5)
    if resultado_update.get('nova_versao'):
        print(f"Ha uma versao nova (v{resultado_update['nova_versao']}). Abra de novo e responda S pra atualizar,")
        print(f'ou baixe em: {_URL_DOWNLOAD}')


def _pausar():
    """Espera Enter no fim (duplo clique); nao quebra se nao houver teclado/console."""
    try:
        input('\nPressione Enter para sair...')
    except (EOFError, OSError):
        pass


if __name__ == '__main__':
    # O console do Windows costuma usar uma codificacao antiga (cp1252 ou
    # similar) que nao sabe imprimir todo caractere Unicode - um nome de
    # arquivo com acento (ex: "Bancário", com o acento como caractere
    # combinante U+0301, comum quando o PDF/pasta veio de outro sistema)
    # jogava uma UnicodeEncodeError no meio do print() e derrubava o
    # programa inteiro (bug real reportado pelo Jean). Reconfigura pra
    # UTF-8 com fallback de substituicao - nunca mais quebra por causa de
    # um caractere que o console nao consegue desenhar.
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

    rodando_como_exe = getattr(sys, 'frozen', False)
    pasta_programa = _pasta_do_programa()

    resultado_update = _INFO_UPDATE
    thread_update = threading.Thread(target=atualizador.verificar_versao,
                                     args=(_URL_VERSION, VERSAO_ATUAL, resultado_update), daemon=True)
    thread_update.start()

    print(f'Conversor PDF para OFX v{VERSAO_ATUAL} - Jean Vieira')
    print('-' * 44)

    # Atualizacao automatica (com confirmacao): ver atualizador.py.
    atualizador.limpar_restos(_NOME_EXE)
    thread_update.join(timeout=2.5)
    if resultado_update.get('nova_versao'):
        atualizador.oferecer_atualizacao(_NOME_EXE, _URL_DOWNLOAD, _URL_HASH, VERSAO_ATUAL,
                                         resultado_update['nova_versao'], _URL_DOWNLOAD)

    if len(sys.argv) > 1:
        # PDFs arrastados e soltos em cima do .exe (ou passados na linha de comando)
        arquivos = sys.argv[1:]
    else:
        # Duplo clique sem arrastar nada: processa todo PDF que estiver na mesma pasta
        arquivos = sorted(glob.glob(os.path.join(pasta_programa, '*.pdf')))

    if not arquivos:
        print('Nenhum PDF encontrado. Coloque os extratos na mesma pasta do programa')
        print('(ou arraste os PDFs em cima do .exe) e abra de novo.')
        _mostrar_aviso_atualizacao(resultado_update, thread_update)
        if rodando_como_exe:
            _pausar()
        sys.exit(1)

    pasta_saida = os.environ.get('PASTA_SAIDA') or pasta_programa
    erros = 0
    for caminho in arquivos:
        nome = os.path.basename(caminho)
        try:
            saida, n, avisos, arquivos_ofx = processar_pdf(caminho, pasta_saida)
            if arquivos_ofx:
                for banco_tela, d1, d2, qtd in _RESUMO_ULTIMO:
                    periodo = f'{d1:%d/%m/%Y} a {d2:%d/%m/%Y}' if d1 and d2 else 'periodo nao identificado'
                    print(f'OK    {banco_tela}, {periodo}: {qtd} lancamentos, .ofx gerado')
                if avisos:
                    print(f'      ATENCAO: leia o arquivo {os.path.splitext(nome)[0]}_AVISOS.txt')
            else:
                erros += 1
                print(f'ERRO  {nome}: o .ofx NAO foi gerado. Abra: {os.path.basename(saida)}')
        except Exception as e:
            erros += 1
            try:
                txt = relatorio_excecao(caminho, pasta_saida, e)
                print(f'ERRO  {nome}: o .ofx NAO foi gerado. Abra: {os.path.basename(txt)}')
            except Exception:
                print(f'ERRO  {nome}: {e}')

    print('-' * 44)
    ok = len(arquivos) - erros
    print(f'Resultado: {ok} com .ofx gerado, {erros} com problema.')
    if erros:
        print('Nos com problema, abra o arquivo _ERRO.txt (fica na mesma pasta do PDF).')
    _mostrar_aviso_atualizacao(resultado_update, thread_update)
    if rodando_como_exe:
        _pausar()
    sys.exit(1 if erros else 0)
