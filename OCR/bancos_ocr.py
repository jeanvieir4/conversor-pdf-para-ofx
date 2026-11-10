# -*- coding: utf-8 -*-
"""
Parsers de extrato lido por OCR (imagem/foto/print de app, sem texto
nativo no PDF). Modulo separado do bancos.py principal - NUNCA reusa nem
altera a logica de la, porque essa aqui trabalha com texto reconhecido
por OCR (sujeito a erro de leitura), nao com texto extraido diretamente
do PDF (confiavel). Ver secao "Extrato digitalizado (OCR)" no
CONTEXTO_PROJETO.md pra entender a decisao de manter isso separado.
"""
import re
import unicodedata
from decimal import Decimal
from datetime import date

import pytesseract


def _dec(s):
    return Decimal(s.replace('.', '').replace(',', '.'))


def linhas_por_posicao(img, escala=1):
    """Reconstroi linhas de texto a partir da posicao (x, y) de cada
    palavra reconhecida pelo Tesseract, em vez de usar a ordem 'natural'
    que ele devolve numa string corrida - testado no extrato real: o
    Tesseract as vezes le a tabela coluna-por-coluna (todas as datas
    juntas, depois todos os valores juntos) em vez de linha por linha,
    o que embaralha completamente a correspondencia entre campos.
    Reconstruindo pela posicao evita isso (mesmo principio do
    `_texto_robusto_por_caracteres` no converter.py, pro bug de texto
    embaralhado do Cresol - so que aqui a fonte e OCR, nao pdfplumber).
    `escala` divide as coordenadas de volta pra escala original quando a
    imagem foi ampliada antes do OCR (upscale melhora a leitura de
    caracteres pequenos como o sinal C/D colado no valor)."""
    dados = pytesseract.image_to_data(img, lang='por', output_type=pytesseract.Output.DICT)
    palavras = []
    for i in range(len(dados['text'])):
        texto = dados['text'][i].strip()
        if not texto:
            continue
        palavras.append({'texto': texto, 'left': dados['left'][i] / escala, 'top': dados['top'][i] / escala})
    palavras.sort(key=lambda p: (p['top'], p['left']))

    linhas = []
    linha_atual = []
    top_atual = None
    for p in palavras:
        if top_atual is None or abs(p['top'] - top_atual) <= 10:
            linha_atual.append(p)
            top_atual = p['top'] if top_atual is None else top_atual
        else:
            linhas.append(sorted(linha_atual, key=lambda x: x['left']))
            linha_atual = [p]
            top_atual = p['top']
    if linha_atual:
        linhas.append(sorted(linha_atual, key=lambda x: x['left']))
    return [' '.join(p['texto'] for p in linha) for linha in linhas]


_HISTORICOS_CAIXA_APP = [
    'SALDO ANTERIOR', 'SALDO DIA', 'CRED PIX CHAVE', 'CRED PIX QR COD EST',
    'DEB PIX CHAVE', 'PIX RECEBIDO DADOS CONTA', 'PAGAMENTO DE BOLETO',
    'DEPOSITO DINH LOTERICO', 'PAGAMENTO GPS INSS IBC', 'PAGAMENTO AGUA',
    'MENSALIDADE CESTA SERVICO',
]

_RE_DATA_INICIO_CAIXA = re.compile(r'^(\d{2}/\d{2}/\d{4})\s*-\s*\d{2}:\d{2}:\d{2}\s+(\d{6})\s+(.*)$')
_RE_VALOR = re.compile(r'([\d.]+,\d{2})\s*(\S)')


def identificar_caixa_app(linhas):
    """Fingerprint do extrato gerado pelo APP da Caixa (diferente do PDF
    'gerenciador' que ja tem parser proprio em bancos.py - texto nativo,
    nao precisa de OCR)."""
    cabecalho = ' '.join(linhas[:6]).upper()
    return 'CAIXA' in cabecalho and 'EXTRATO' in cabecalho


def parse_caixa_app_ocr(linhas):
    """Reconstroi lancamentos do extrato do app da Caixa a partir das
    linhas ja remontadas por posicao (`linhas_por_posicao`). Cada linha
    de lancamento tem o formato:
    'DD/MM/AAAA - HH:MM:SS NRDOC HISTORICO [favorecido] [cpf] VALOR SINAL SALDO SINAL'

    So extrai Data/Historico/Valor/Tipo - NUNCA tenta reconstruir o nome
    do favorecido nem o CPF/CNPJ. Decisao deliberada: esses dois campos
    vem em quantidade desalinhada em relacao as datas (linhas "SALDO DIA"
    nao tem favorecido, por exemplo), entao juntar tudo por posicao
    arrisca colar o nome errado numa transacao errada. Historico usa uma
    lista fechada de categorias conhecidas (vocabulario observado no
    extrato real testado) - historico fora dessa lista e ignorado (linha
    descartada, nao vira lancamento) em vez de arriscar interpretar
    errado.

    LIMITACAO CONHECIDA (testada num extrato real de 4 paginas): uma
    fatia real das linhas (~5 a 11%% dependendo do preprocessamento de
    imagem) sai com o sinal C/D indeterminado - as vezes porque o
    Tesseract nao reconhece o caractere C/D (fonte pequena, colado no
    numero), as vezes porque o saldo mostrado na linha repete o saldo da
    transacao anterior (peculiaridade do proprio extrato do app, nao bug
    do parser - confirmado inspecionando a imagem original). Pra tentar
    resolver esses casos, o sinal tambem e inferido comparando o saldo da
    linha com o saldo da linha anterior (delta): se o delta bater com o
    valor da transacao, usa o sinal do delta como confirmacao adicional.
    Quando NEM o OCR nem o delta conseguem confirmar o sinal, a transacao
    fica marcada '[A VERIFICAR]' na planilha e e EXCLUIDA do arquivo OFX
    (nunca inventa um sinal - ver `conversor_ocr.py`)."""
    saldo_anterior_declarado = None
    sequencia = []
    for linha in linhas:
        if 'SALDO ANTERIOR' in linha.upper() and saldo_anterior_declarado is None:
            m = re.search(r'([\d.]+,\d{2})', linha)
            if m:
                saldo_anterior_declarado = _dec(m.group(1))

        m = _RE_DATA_INICIO_CAIXA.match(linha.strip())
        if not m:
            continue
        data_str, _doc, resto = m.groups()
        valores = _RE_VALOR.findall(resto)
        if len(valores) < 2:
            continue
        valor_str, sinal_valor = valores[-2]
        saldo_str, _sinal_saldo = valores[-1]

        historico = None
        resto_upper = resto.upper()
        for h in _HISTORICOS_CAIXA_APP:
            if resto_upper.startswith(h):
                historico = h
                break
        if historico is None:
            continue

        dd, mm, yyyy = data_str.split('/')
        try:
            dt = date(int(yyyy), int(mm), int(dd))
        except ValueError:
            continue

        sinal_ocr = 'C' if sinal_valor.upper() == 'C' else ('D' if sinal_valor.upper() == 'D' else None)
        sequencia.append({
            'data': dt, 'historico': historico, 'valor': _dec(valor_str),
            'sinal_ocr': sinal_ocr, 'saldo_linha': _dec(saldo_str),
        })

    saldo_ref = saldo_anterior_declarado
    transacoes = []
    n_incertas = 0
    for t in sequencia:
        if t['historico'] in ('SALDO DIA', 'SALDO ANTERIOR'):
            saldo_ref = t['saldo_linha']
            continue

        delta = t['saldo_linha'] - saldo_ref if saldo_ref is not None else None
        sinal_por_delta = None
        if delta is not None and abs(abs(delta) - t['valor']) <= Decimal('0.02'):
            sinal_por_delta = 'C' if delta > 0 else 'D'
        sinal_final = t['sinal_ocr'] or sinal_por_delta

        if sinal_final is None:
            n_incertas += 1
            obs = ('[A VERIFICAR - OCR] Credito/debito nao pode ser confirmado. '
                   'Nao entra no .ofx - confira no extrato original e lance manualmente se for o caso.')
        else:
            # obs fica vazio nas confiaveis de proposito: esse campo tambem vira o
            # MEMO do .ofx (ver ofx_export.py) - um aviso generico repetido em toda
            # linha so polui o arquivo financeiro. O aviso de "gerado por OCR" ja
            # fica uma vez so, na aba Pendencias (ver `aviso` no fim da funcao).
            obs = ''

        transacoes.append({
            'data': t['data'], 'historico': t['historico'], 'valor': t['valor'],
            'tipo': sinal_final or '?', 'obs': obs, 'incerta': sinal_final is None,
        })
        saldo_ref = t['saldo_linha']

    aviso = (f'Extrato lido por OCR (imagem, nao PDF nativo). {len(transacoes)} lancamento(s) '
             f'encontrados, {n_incertas} com credito/debito incerto (marcados [A VERIFICAR] e '
             f'ausentes do .ofx). CONFIRA TODOS OS VALORES contra o extrato original antes de usar '
             f'- leitura por OCR pode errar mesmo quando parece certa.')
    return transacoes, aviso


def identificar_bb_ocr(linhas):
    """Fingerprint do extrato do Banco do Brasil lido por OCR - visto
    como um PRINT/PDF salvo do internet banking (autoatendimento.bb.com.br),
    nao um PDF gerado nativamente pelo banco (esse ja tem parser proprio
    em bancos.py, texto nativo, nao precisa de OCR). Mesmo layout de
    colunas do formato "atual" nativo do BB (Ag. origem/Lote/Documento/
    Historico/Valor/Saldo), so que aqui vem como imagem."""
    texto = ' '.join(linhas).upper()
    return 'BANCO DO BRASIL' in texto and ('AUTOATENDIMENTO.BB.COM.BR' in texto or 'AG. ORIGEM' in texto)


_RE_BB_TRANSACAO = re.compile(
    r'^(\d{2}/\d{2}/\d{4})\s+(\d{4})\s+(\d{5})\s+(\d{3})\s*(.+?)\s+([\d.]+,\d{2})\s*([CD€])'
    r'(?:\s+[\d.]+,\d{2}\s*[CD€])?$'
)
_BB_EXCLUIR_HIST = ('SALDO ANTERIOR', 'BB RENDE F', 'RENDE FACIL', 'SALDO')


def parse_bb_ocr(linhas):
    """Reconstroi lancamentos do extrato do Banco do Brasil (print do
    internet banking) a partir das linhas ja remontadas por posicao.
    Cada lancamento vem numa linha so:
    'DD/MM/AAAA AGORIGEM LOTE DOCUMENTO HISTORICO VALOR TIPO [SALDO TIPO]'
    - o "historico" fica com tudo entre o documento e o valor (inclusive
    um numero de identificacao de transacao que aparece no meio, tipo
    "890.471.573.141.144") sem tentar separar mais - mesma decisao ja
    usada nos outros parsers de OCR (nao vale o risco de cortar errado).

    PECULIARIDADE DE OCR: o Tesseract as vezes le o simbolo circular "C"
    (credito) como "€" (euro) - visualmente parecidos em certas fontes.
    Tratado como equivalente a "C" aqui.

    "BB Rende Facil"/"Rende Facil" e a varredura automatica pra uma
    aplicacao financeira que soma e resgata o dinheiro no mesmo dia
    (mesmo padrao do "Aplic Aut Mais" do Itau e do "CAPTACAO" do Sicredi,
    ja excluidos nos parsers deles) - excluido aqui tambem, nao conta
    como movimento real de conta corrente. "Saldo Anterior" e a linha
    "SALDO" final (resumo) tambem sao ignoradas, nao sao lancamento.

    LIMITACAO CONHECIDA: quando o OCR erra um DIGITO dentro do valor (ex:
    le "93,75" como "93,/5", ou acrescenta um digito a mais tipo
    "98,095"), a linha inteira nao bate no regex e o lancamento e
    perdido silenciosamente - nao da pra adivinhar o digito certo. O
    aviso final conta quantas linhas com data nao foram reconhecidas,
    pra deixar isso visivel.

    Validado com extrato real (4 paginas): 45 lancamentos reconhecidos,
    3 linhas perdidas por erro de digito no OCR. Reconstruindo o saldo
    linha a linha (incluindo as linhas de Rende Facil, que tem valor E
    saldo na mesma linha) contra o que o proprio extrato declara: 15 de
    19 batem exato, as 4 divergencias restantes sao explicadas
    exatamente pelas 3 linhas perdidas (a diferenca do saldo bate com o
    valor da linha perdida correspondente) - confirma que a extracao dos
    valores reconhecidos esta correta, o erro e so nas poucas linhas que
    o OCR realmente nao conseguiu ler direito.
    """
    if not identificar_bb_ocr(linhas) and identificar_bb_ocr_dep_origem(linhas):
        return _parse_bb_ocr_dep_origem(linhas)
    linhas_com_data = 0
    linhas_nao_reconhecidas = 0
    transacoes = []
    for linha in linhas:
        l = linha.strip()
        if not re.match(r'^\d{2}/\d{2}/\d{4}\s+\d{4}', l):
            continue
        linhas_com_data += 1
        m = _RE_BB_TRANSACAO.match(l)
        if not m:
            linhas_nao_reconhecidas += 1
            continue
        data_str, _ag, _lote, _doc, historico, valor_str, tipo = m.groups()
        historico = historico.strip()
        hist_upper = historico.upper()
        if any(hist_upper.startswith(k) or hist_upper == k for k in _BB_EXCLUIR_HIST):
            continue
        dd, mm, yyyy = data_str.split('/')
        try:
            dt = date(int(yyyy), int(mm), int(dd))
        except ValueError:
            continue
        tipo_norm = 'C' if tipo in ('C', '€') else 'D'
        transacoes.append({
            'data': dt, 'historico': historico, 'valor': _dec(valor_str),
            'tipo': tipo_norm, 'obs': '', 'incerta': False,
        })

    aviso = (f'Extrato lido por OCR (imagem, nao PDF nativo). {len(transacoes)} lancamento(s) '
             f'encontrados. {linhas_nao_reconhecidas} linha(s) com data nao puderam ser lidas '
             f'corretamente (numero com digito ilegivel) e foram IGNORADAS - podem faltar '
             f'lancamentos. Linhas "BB Rende Facil" (varredura automatica pra aplicacao) tambem '
             f'foram excluidas, nao contam como movimento de conta corrente. CONFIRA TODOS OS '
             f'VALORES contra o extrato original antes de usar.')
    return transacoes, aviso


def identificar_bb_ocr_dep_origem(linhas):
    """Segundo layout do BB lido por OCR: extrato de conta corrente
    escaneado (papel, "Mod. 0.50.817-2", SISBB) com colunas Data / Dep.
    origem / Historico / Documento / Valor / Saldo e uma linha de detalhe
    abaixo de cada lancamento. So vale se o layout 1 (autoatendimento)
    NAO bateu - o layout 1 segue tratado exatamente como antes."""
    texto = ' '.join(linhas).upper()
    return ('SISBB' in texto and 'EXTRATO DE CONTA CORRENTE' in texto
            and 'DEP.' in texto and 'SALDO ANTERIOR' in texto)


_RE_BB2_VALOR_SALDO = re.compile(r'([\d.]+,\d{2})\s*([CD€])\s+([\d.]+,\d{2})\s*([CD€])\s*$')
_RE_BB2_SALDO_ANT = re.compile(r'SALDO ANTERIOR\s+([\d.]+,\d{2})\s*([CD€])', re.IGNORECASE)
_RE_BB2_DATA = re.compile(r'^(\d{2}/\d{2}/\d{4})\b')


def _parse_bb_ocr_dep_origem(linhas):
    """Cada lancamento tem 'DATA [AG] HISTORICO [DOC] VALOR C|D SALDO C|D' e,
    na linha de baixo, um detalhe ('03/05 19:46 ... NOME'). O OCR as vezes
    separa a data do resto da linha (papel levemente torto), entao uma data
    sozinha fica "pendente" pra proxima linha com valor. O saldo corrente
    declarado em cada linha serve de conferencia: se saldo novo - saldo
    anterior nao for exatamente o valor lido, o lancamento vira incerto
    (vermelho no Excel, fora do OFX); se so o sinal C/D foi lido errado
    mas o valor bate, o sinal e corrigido pelo saldo."""
    def sinal(v, t):
        return v if t in ('C', '€') else -v

    saldo_prev = None
    for l in linhas:
        m = _RE_BB2_SALDO_ANT.search(l)
        if m:
            saldo_prev = sinal(_dec(m.group(1)), m.group(2))
            break

    transacoes = []
    pendente = None
    ult = None
    n_incertas = 0
    for linha in linhas:
        l = linha.strip()
        if '***' in l or re.search(r'\bSALDO\s*$', l, re.IGNORECASE):
            ult = None
            continue
        m = _RE_BB2_VALOR_SALDO.search(l)
        if m and 'SALDO ANTERIOR' not in l.upper():
            prefixo = l[:m.start()].strip()
            md = _RE_BB2_DATA.match(prefixo)
            if md:
                data_str = md.group(1)
                prefixo = prefixo[md.end():].strip()
            else:
                data_str = pendente
            pendente = None
            prefixo = re.sub(r'^\d{4}-?\d\b\s*', '', prefixo)
            prefixo = re.sub(r'\s+[\d.]*\d\.\d{3}$', '', prefixo).strip()
            prefixo = prefixo.replace('COMPRACICARTAO', 'COMPRA C/CARTAO')
            if not data_str:
                ult = None
                continue
            dd, mm, yyyy = data_str.split('/')
            try:
                dt = date(int(yyyy), int(mm), int(dd))
            except ValueError:
                ult = None
                continue
            valor = _dec(m.group(1))
            saldo = sinal(_dec(m.group(3)), m.group(4))
            tipo = 'C' if m.group(2) in ('C', '€') else 'D'
            incerta = False
            if saldo_prev is not None:
                delta = saldo - saldo_prev
                if abs(delta) == valor:
                    tipo = 'C' if delta > 0 else 'D'
                else:
                    incerta = True
            saldo_prev = saldo
            if incerta:
                n_incertas += 1
            ult = {'data': dt, 'historico': prefixo, 'valor': valor, 'tipo': tipo,
                   'obs': '[A VERIFICAR - OCR] valor/saldo nao conferem' if incerta else '',
                   'incerta': incerta, '_detalhes': 0}
            transacoes.append(ult)
            continue
        md = _RE_BB2_DATA.match(l)
        if md:
            pendente = md.group(1)
            ult = None
            continue
        if ult is not None and ult['_detalhes'] < 2:
            limpo = re.sub(r'^[^0-9A-Za-z]+', '', l)
            if len(re.sub(r'[^0-9A-Za-z]', '', limpo)) >= 4 and not re.fullmatch(r'[\d.\s=\-]+', limpo):
                ult['historico'] = (ult['historico'] + ' ' + limpo).strip()
                ult['_detalhes'] += 1
    for t in transacoes:
        t.pop('_detalhes', None)

    aviso = (f'Extrato lido por OCR (imagem escaneada, nao PDF nativo). {len(transacoes)} lancamento(s) '
             f'encontrados; cada um foi conferido contra o saldo corrente que o proprio extrato declara. '
             f'{n_incertas} nao conferiram e ficaram marcados "[A VERIFICAR - OCR]" (fora do OFX). '
             f'CONFIRA TODOS OS VALORES contra o extrato original antes de usar.')
    return transacoes, aviso


def _sem_acento(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')


def identificar_sicoob_ocr(linhas):
    """Extrato Sicoob (SISBR) escaneado: conta corrente ou extrato de
    aplicacoes. So e usado se Caixa e BB nao bateram antes."""
    t = _sem_acento(' '.join(linhas).upper()).replace(' ', '')   # alguns OCR perdem os espacos
    return 'SICOOB' in t and ('SISBR' in t or 'SISTEMADECOOPERATIVAS' in t) and 'MOVIMENTA' in t


_NUM = r'[\d.]+,\d{2}'
# "DD/MM[/AAAA] descricao valor<C|D|ruido>"; o ultimo caractere e o sinal (o OCR
# le o C como "€" ou "(" as vezes; por isso o sinal nunca e confiado sozinho).
_RE_SC_LINHA = re.compile(r'^(\d{2}/\d{2})(?:/(\d{4}))?\s+(.+?)\s+(' + _NUM + r')\s*([^\d\s]?)$')
# extrato de aplicacoes: "DD/MM/AAAA HISTORICO valor [C|D] saldo"
_RE_SC_APLIC = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(' + _NUM + r')(?:\s*([^\d\s]))?(?:\s+(' + _NUM + r'))?$')
_RE_SC_PERIODO = re.compile(r'PER.{1,2}DO:?\s*(\d{2}/\d{2}/\d{4})\D{0,8}(\d{2}/\d{2}/\d{4})', re.IGNORECASE)


def _sc_parece_lancamento(linha):
    """Linha que comeca com data e tem valor, mas que o leitor nao conseguiu
    interpretar (OCR bagunçou). Nao pode sumir em silencio. Linhas que so
    quebraram o C/D ou o valor pra linha vizinha (tratadas em
    _sc_juntar_quebradas) e cabecalhos nao contam."""
    l = linha.strip()
    return bool(re.match(r'^\d{2}/\d{2}(/\d{4})?\s+\S', l) and re.search(_NUM, l)
                and not re.match(r'^\d{2}/\d{2}/\d{4}\s+EXTRATO', l, re.IGNORECASE))


def _sc_aviso_nao_lidas(n):
    return (f'. ATENCAO: {n} linha(s) com data e valor NAO puderam ser lidas e estao FORA do .ofx - '
            f'compare com o extrato' if n else '')


def _sc_sinal(letra):
    return -1 if letra == 'D' else 1


def _sc_juntar_quebradas(linhas):
    """Valor >= 1.000,00 faz o OCR (e o pdfplumber) as vezes quebrar a linha:
    o C/D sozinho na linha de baixo. Junta de volta."""
    sozinho_cd = re.compile(r'^[CD€(]$')
    data_desc = re.compile(r'^\d{2}/\d{2}(/\d{4})?\s+\S.*$')
    data_desc_valor = re.compile(r'^\d{2}/\d{2}(/\d{4})?\s+.+?\s' + _NUM + r'\s*[^\d\s]?$')
    out = []
    i = 0
    while i < len(linhas):
        l = linhas[i].strip()
        # (b) valor sozinho, depois "DD/MM descricao" sem valor, depois C/D sozinho
        if (re.match(r'^' + _NUM + r'$', l) and i + 2 < len(linhas)
                and data_desc.match(linhas[i + 1].strip()) and not data_desc_valor.match(linhas[i + 1].strip())
                and sozinho_cd.match(linhas[i + 2].strip())):
            out.append(f'{linhas[i + 1].strip()} {l}{linhas[i + 2].strip()}')
            i += 3
            continue
        # (a) "DD/MM descricao valor" sem sinal, depois C/D sozinho
        if (re.match(r'^\d{2}/\d{2}(/\d{4})?\s+.+\s' + _NUM + r'$', l) and i + 1 < len(linhas)
                and sozinho_cd.match(linhas[i + 1].strip())):
            out.append(l + linhas[i + 1].strip())
            i += 2
            continue
        out.append(l)
        i += 1
    return out


def _parse_sicoob_ocr_aplicacoes(linhas):
    """Extrato de aplicacoes (RDC): cada linha traz o saldo corrente
    ('valor D saldo'), entao cada lancamento e conferido (saldo novo - saldo
    anterior == valor) e o sinal vem do saldo, nao da letra lida. A 1a linha
    (aplicacao inicial) tem so o valor, que e tambem o saldo."""
    transacoes = []
    saldo_prev = None
    n_incertas = 0
    n_nao_lidas = 0
    for l in linhas:
        m = _RE_SC_APLIC.match(l.strip())
        if not m:
            if _sc_parece_lancamento(l):
                n_nao_lidas += 1
            continue
        data_str, desc, v, letra, s = m.groups()
        dd, mm, yyyy = data_str.split('/')
        try:
            dt = date(int(yyyy), int(mm), int(dd))
        except ValueError:
            continue
        valor = _dec(v)
        incerta = False
        if s is None:                      # primeira linha: aplicacao inicial
            tipo = 'C'
            saldo = valor
        else:
            saldo = _dec(s)
            tipo = 'D' if letra == 'D' else 'C'
            if saldo_prev is not None:
                delta = saldo - saldo_prev
                if abs(delta) == valor:
                    tipo = 'C' if delta > 0 else 'D'
                else:
                    incerta = True
        saldo_prev = saldo
        if incerta:
            n_incertas += 1
        transacoes.append({'data': dt, 'historico': desc.strip(), 'valor': valor, 'tipo': tipo,
                           'obs': '[A VERIFICAR - OCR] valor/saldo nao conferem' if incerta else '',
                           'incerta': incerta})
    aviso = (f'Extrato Sicoob lido por OCR (imagem escaneada). {len(transacoes)} lancamento(s); cada um foi '
             f'conferido contra o saldo corrente do proprio extrato. {n_incertas} nao conferiram e ficaram '
             f'"[A VERIFICAR - OCR]" (fora do OFX)'
             + _sc_aviso_nao_lidas(n_nao_lidas)
             + '. CONFIRA TODOS OS VALORES contra o extrato original.')
    return transacoes, aviso


def parse_sicoob_ocr(linhas):
    """Sicoob escaneado. Conta corrente: 'DD/MM descricao valorC|D' (ou com
    ano e documento). As linhas 'SALDO DO DIA' sao pontos de conferencia: a
    soma dos lancamentos desde o ultimo saldo tem que dar o saldo declarado.
    Se nao der, tenta achar UM lancamento com sinal C/D lido errado; se nao
    achar, todos os lancamentos do trecho ficam incertos (fora do OFX)."""
    texto = _sem_acento('\n'.join(linhas).upper()).replace(' ', '')
    if 'EXTRATODEAPLICA' in texto:
        return _parse_sicoob_ocr_aplicacoes(linhas)

    mp = _RE_SC_PERIODO.search('\n'.join(linhas))
    ini = fim = None
    if mp:
        ini = date(*map(int, reversed(mp.group(1).split('/'))))
        fim = date(*map(int, reversed(mp.group(2).split('/'))))

    base = None           # ultimo saldo CONFERIDO (saldo anterior, ou SALDO DO DIA que bateu)
    ult_decl = None       # ultimo SALDO DO DIA lido (mesmo que nao tenha batido)
    pend = []             # lancamentos de dias anteriores que ainda nao bateram
    dia = []              # lancamentos desde o ultimo SALDO DO DIA
    transacoes = []
    n_corrigidas = 0
    n_sem_base = 0
    n_lacunas = 0
    sem_ano = False

    def soma(lista):
        return sum(t['valor'] * (1 if t['tipo'] == 'C' else -1) for t in lista)

    def tenta(ref, lista, alvo):
        """True se ref + soma(lista) == alvo; senao tenta achar UM lancamento
        com C/D lido errado que faca bater (so aceita se for unico)."""
        total = ref + soma(lista)
        if total == alvo:
            return True
        cands = [t for t in lista if total + (-2 if t['tipo'] == 'C' else 2) * t['valor'] == alvo]
        if len(cands) == 1:
            cands[0]['tipo'] = 'D' if cands[0]['tipo'] == 'C' else 'C'
            return True
        return False

    def marcar(lista, motivo):
        for t in lista:
            t['incerta'] = True
            t['obs'] = '[A VERIFICAR - OCR] ' + motivo

    n_nao_lidas = 0
    for l in _sc_juntar_quebradas(linhas):
        m = _RE_SC_LINHA.match(l)
        if not m:
            if _sc_parece_lancamento(l):
                n_nao_lidas += 1
            continue
        dm, ano, desc, v, letra = m.groups()
        desc = desc.strip()
        d_up = _sem_acento(desc.upper()).replace(' ', '')
        if letra == '*':
            continue                      # SALDO BLOQ., cheque bloqueado: so aviso, nao e credito
        valor = _dec(v)
        if d_up.startswith('SALDOANTERIOR'):
            base = ult_decl = valor * _sc_sinal(letra)
            pend, dia = [], []
            continue
        if d_up.startswith('SALDOBLOQ'):
            continue
        if d_up.startswith('SALDODODIA'):
            declarado = valor * _sc_sinal(letra)
            if base is None:
                # saldo anterior nao foi lido: o primeiro trecho nao tem como ser
                # conferido (fica incerto); a partir daqui o saldo declarado vira a base.
                marcar(dia, 'saldo anterior nao foi lido')
                n_sem_base += len(dia)
                base = ult_decl = declarado
                dia = []
                continue
            # Um SALDO DO DIA que nao bate pode ser (1) lancamento com C/D lido
            # errado, (2) o proprio saldo lido errado ou (3) lancamento perdido/
            # com valor lido errado. (2): o proximo saldo que bater confirma tudo
            # que veio no meio. (3): o proximo saldo bate considerando so o dia
            # novo a partir do saldo anterior declarado - o trecho ruim fica incerto.
            if tenta(base, pend + dia, declarado):
                base = declarado
                pend, dia = [], []
            elif tenta(ult_decl, dia, declarado):
                marcar(pend, 'saldo do dia nao conferiu')
                base = declarado
                pend, dia = [], []
            else:
                if not (pend + dia):
                    n_lacunas += 1      # saldo mudou e nao ha lancamento lido que explique
                pend = pend + dia
                dia = []
            ult_decl = declarado
            continue
        dd, mm = dm.split('/')
        if ano:
            yyyy = int(ano)
        elif ini:
            yyyy = ini.year if int(mm) >= ini.month else fim.year
        else:
            sem_ano = True
            continue
        data_ruim = False
        try:
            dt = date(yyyy, int(mm), int(dd))
        except ValueError:
            # o OCR leu a data errado (ex.: "60/60"): o lancamento NAO pode sumir em
            # silencio. Entra como incerto, com a data do lancamento anterior.
            data_ruim = True
            dt = transacoes[-1]['data'] if transacoes else (ini or date.today())
        t = {'data': dt, 'historico': desc, 'valor': valor, 'tipo': 'D' if letra == 'D' else 'C',
             'obs': '', 'incerta': False}
        if data_ruim:
            t['incerta'] = True
            t['obs'] = '[A VERIFICAR - OCR] data ilegivel (' + dm + ')'
            t['historico'] = desc + ' [data lida: ' + dm + ']'
        transacoes.append(t)
        dia.append(t)

    # o que sobrou sem conferir: se ha saldo de referencia, nenhum SALDO DO DIA
    # bateu depois deles -> incerto; sem nenhum saldo no extrato, nao da pra conferir nada.
    resto = pend + dia
    n_sem_conferencia = 0
    if base is not None:
        marcar(resto, 'saldo do dia nao conferiu')
    else:
        n_sem_conferencia = len(resto)
    n_incertas = sum(1 for t in transacoes if t['incerta'])

    if sem_ano and not transacoes:
        return [], 'Sicoob (OCR): nao achei a linha PERIODO: no extrato; ano das datas sem referencia segura'
    aviso = (f'Extrato Sicoob lido por OCR (imagem escaneada). {len(transacoes)} lancamento(s); cada trecho foi '
             f'conferido contra a linha "SALDO DO DIA" do proprio extrato (sinal C/D lido errado e corrigido '
             f'pelo saldo); {n_incertas} lancamento(s) em trecho que nao conferiu ficaram '
             f'"[A VERIFICAR - OCR]" (fora do OFX)'
             + (f'; {n_sem_conferencia} sem saldo pra conferir' if n_sem_conferencia else '')
             + (f'. ATENCAO: em {n_lacunas} ponto(s) o saldo do extrato mudou sem lancamento lido que explique - '
                f'PODE TER FALTADO LANCAMENTO, compare com o extrato' if n_lacunas else '')
             + _sc_aviso_nao_lidas(n_nao_lidas)
             + '. CONFIRA TODOS OS VALORES contra o extrato original.')
    return transacoes, aviso


# ===========================================================================
# v3.2 - PagBank, Banco do Brasil (Dt. balancete/Ag. origem/Lote) e Caixa
# (Gerenciador Caixa) lidos por OCR. Lidos preferencialmente pelo RapidOCR
# (ocr_rapid.py), que perde os espacos entre palavras - por isso todos os
# detectores/parsers abaixo ignoram espacos. Cada lancamento e conferido
# pelos saldos que o proprio extrato declara; o que nao confere fica
# incerto (fora do OFX) e o que nao deu pra ler vai pro aviso.
# ===========================================================================
def _padrao_hist(texto):
    partes = []
    for ch in texto:
        if ch == ' ':
            continue
        base = _sem_acento(ch)
        partes.append('[' + re.escape(ch) + re.escape(base) + ']' if base != ch else re.escape(ch))
    return re.compile(r'\s*'.join(partes), re.IGNORECASE)


def _restaurador(lista):
    """Devolve f(texto) que recoloca os espacos de historicos padrao ('PIXRECEBIDO' ->
    'PIX RECEBIDO'). So deixa o .ofx legivel; nao afeta valor/data/sinal."""
    pads = sorted(((_padrao_hist(h), h) for h in lista), key=lambda x: -len(x[1]))

    def restaurar(txt):
        for pad, canon in pads:
            novo, n = pad.subn(canon, txt, count=1)
            if n:
                return novo
        return txt
    return restaurar


class _Conferidor:
    """Confere lancamentos contra saldos declarados no extrato (mesma logica do
    Sicoob). Uso: tx(t) para cada lancamento e saldo(valor_com_sinal) a cada saldo
    declarado. Lancamentos com 'oculta' True entram na soma mas nao no resultado
    (ex.: varredura automatica pro BB Rende Facil)."""

    def __init__(self, base=None):
        self.base = base
        self.ult = base
        self.pend = []
        self.dia = []
        self.n_lacunas = 0
        self.n_sem_base = 0
        self.n_sem_conferencia = 0

    @staticmethod
    def soma(lista):
        return sum(t['valor'] * (1 if t['tipo'] == 'C' else -1) for t in lista)

    @staticmethod
    def _tenta(ref, lista, alvo):
        total = ref + _Conferidor.soma(lista)
        if total == alvo:
            return True
        cands = [t for t in lista if total + (-2 if t['tipo'] == 'C' else 2) * t['valor'] == alvo]
        if len(cands) == 1:
            cands[0]['tipo'] = 'D' if cands[0]['tipo'] == 'C' else 'C'
            return True
        return False

    @staticmethod
    def _marcar(lista, motivo):
        for t in lista:
            t['incerta'] = True
            t['obs'] = '[A VERIFICAR - OCR] ' + motivo

    def tx(self, t):
        self.dia.append(t)

    def saldo(self, declarado):
        if self.base is None:
            self._marcar(self.dia, 'saldo anterior nao foi lido')
            self.n_sem_base += len(self.dia)
            self.base = self.ult = declarado
            self.dia = []
            return
        if self._tenta(self.base, self.pend + self.dia, declarado):
            self.base = declarado
            self.pend, self.dia = [], []
        elif self._tenta(self.ult, self.dia, declarado):
            self._marcar(self.pend, 'saldo nao conferiu')
            self.base = declarado
            self.pend, self.dia = [], []
        else:
            if not (self.pend + self.dia):
                self.n_lacunas += 1
            self.pend = self.pend + self.dia
            self.dia = []
        self.ult = declarado

    def fim(self):
        resto = self.pend + self.dia
        if self.base is not None:
            self._marcar(resto, 'saldo nao conferiu')
        else:
            self.n_sem_conferencia = len(resto)


def _montar_aviso(banco, transacoes, conf, n_nao_lidas, extra=''):
    n_inc = sum(1 for t in transacoes if t.get('incerta'))
    return (f'Extrato {banco} lido por OCR (imagem escaneada). {len(transacoes)} lancamento(s); cada trecho '
            f'foi conferido contra os saldos do proprio extrato (sinal C/D lido errado e corrigido pelo saldo); '
            f'{n_inc} lancamento(s) em trecho que nao conferiu ficaram "[A VERIFICAR - OCR]" (fora do OFX)'
            + (f'; {conf.n_sem_conferencia} sem saldo pra conferir' if conf.n_sem_conferencia else '')
            + (f'. ATENCAO: em {conf.n_lacunas} ponto(s) o saldo do extrato mudou sem lancamento lido que '
               f'explique - PODE TER FALTADO LANCAMENTO, compare com o extrato' if conf.n_lacunas else '')
            + _sc_aviso_nao_lidas(n_nao_lidas) + extra
            + '. CONFIRA TODOS OS VALORES contra o extrato original.')


def _sem_espaco(texto):
    return _sem_acento(texto.upper()).replace(' ', '')


def _sinal_letra(letra):
    return -1 if letra == 'D' else 1


# ---------------------------------------------------------------- PagBank
def identificar_pagbank_ocr(linhas):
    t = _sem_espaco(' '.join(linhas))
    return 'PAGSEGUROINTERNET' in t and 'EXTRATODACONTA' in t


_RE_PAGBANK = re.compile(r'^(\d{2}/\d{2}/\d{4})\s*(.+?)\s*(-?)\s*R\$\s*(' + _NUM + r')$')
_HIST_PAGBANK = _restaurador(['Pix recebido', 'Pix enviado', 'Vendas - Disponivel', 'Transferencia recebida',
                              'Transferencia enviada'])


def _limpa_hist_pagbank(desc):
    desc = _HIST_PAGBANK(desc)
    desc = re.sub(r'^(Pix (?:recebido|enviado))\s*-\s*', r'\1 - ', desc)
    desc = re.sub(r'^(Vendas - Disponivel)\s*', r'\1 ', desc)
    return desc.strip()


def parse_pagbank_ocr(linhas):
    """PagBank: 'DD/MM/AAAA Descricao [-]R$ valor'; 'Saldo do dia' confere (o
    saldo inicial nao vem no extrato: assume 0,00 como o leitor de texto)."""
    transacoes = []
    conf = _Conferidor(base=Decimal('0'))
    n_nao_lidas = 0
    for l in linhas:
        m = _RE_PAGBANK.match(l.strip())
        if not m:
            if _sc_parece_lancamento(l):
                n_nao_lidas += 1
            continue
        data_str, desc, neg, v = m.groups()
        dd, mm, yyyy = data_str.split('/')
        valor = _dec(v)
        if _sem_espaco(desc) == 'SALDODODIA':
            conf.saldo(valor * (-1 if neg else 1))
            continue
        try:
            dt = date(int(yyyy), int(mm), int(dd))
            data_ruim = False
        except ValueError:
            dt = transacoes[-1]['data'] if transacoes else date.today()
            data_ruim = True
        t = {'data': dt, 'historico': _limpa_hist_pagbank(desc), 'valor': valor, 'tipo': 'D' if neg else 'C',
             'obs': '', 'incerta': False}
        if data_ruim:
            t['incerta'] = True
            t['obs'] = '[A VERIFICAR - OCR] data ilegivel (' + data_str[:5] + ')'
        transacoes.append(t)
        conf.tx(t)
    conf.fim()
    return transacoes, _montar_aviso('PagBank', transacoes, conf, n_nao_lidas)


# ------------------------------------------------- Banco do Brasil (balancete)
def identificar_bb_balancete_ocr(linhas):
    """BB 'Consultas - Extrato de conta corrente' com colunas Dt. balancete /
    Dt. movimento / Ag. origem / Lote / Historico / Documento / Valor / Saldo."""
    t = _sem_espaco(' '.join(linhas))
    return 'EXTRATODECONTACORRENTE' in t and 'DT.BALANCETE' in t and 'AG.ORIGEM' in t


_RE_BB3 = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(\d{4})\s*(.*?)\s*(' + _NUM + r')\s*([CD€])\s*(?:(' + _NUM +
                     r')\s*([CD€]))?$')
_HIST_BB3 = _restaurador([
    'Pix - Recebido QR Code', 'Pix - Recebido', 'Pix - Enviado', 'Dep dinheiro ATM', 'Dep dinheiro inter ag',
    'Dep disponivel agencia', 'Dep CORBAN dinheiro', 'Tarifa Pix Enviado', 'Tarifa Pacote de Servicos',
    'Pagamentos Diversos', 'Cheque Compensado', 'Aplicacao BB CDB DI', 'BB Rende Facil', 'SALDO ANTERIOR',
    'Transferencia enviada', 'Transferencia recebida', 'Deposito', 'Saque'])


def _bb3_assinado(valor, letra):
    return valor * (1 if letra in ('C', '€') else -1)


def parse_bb_balancete_ocr(linhas):
    """Cada lancamento: 'DD/MM/AAAA Ag [lote cod]Historico Documento valor C|D [saldo C|D]'.
    Todo dia termina com uma varredura pro 'BB Rende Facil' (some do OFX, como no
    leitor de texto) e saldo 0,00 - esse saldo confere o dia."""
    transacoes = []
    conf = _Conferidor()
    n_nao_lidas = 0
    for l in linhas:
        l = l.strip()
        m = _RE_BB3.match(l)
        if not m:
            if _sc_parece_lancamento(l) and 'SALDO' not in _sem_espaco(l):
                n_nao_lidas += 1
            continue
        data_str, _ag, meio, v, letra, s, sletra = m.groups()
        valor = _dec(v)
        meio = meio.strip()
        mc = re.match(r'^([\dOo]{4,5}\s?[\dOo]{3})\s*(.*)$', meio)
        codigo, resto = (mc.group(1), mc.group(2)) if mc else ('', meio)
        resto = re.sub(r'\s+[\d.]*\d\.\d{3}$', '', resto).strip()   # coluna Documento
        resto = re.sub(r'\s+\d+$', '', resto).strip() if re.search(r'\s\d{1,3}$', resto) else resto
        d_ne = _sem_espaco(resto)
        if d_ne.startswith('SALDOANTERIOR'):
            conf.base = conf.ult = _bb3_assinado(valor, letra)
            continue
        if d_ne in ('SALDO',):
            conf.saldo(_bb3_assinado(valor, letra))
            continue
        dd, mm, yyyy = data_str.split('/')
        data_ruim = False
        try:
            dt = date(int(yyyy), int(mm), int(dd))
        except ValueError:
            dt = transacoes[-1]['data'] if transacoes else date.today()
            data_ruim = True
        rende = d_ne.startswith('BBRENDEF')
        cod = re.sub(r'\s', '', codigo).replace('o', '0').replace('O', '0')
        t = {'data': dt, 'historico': (cod + ' ' + _HIST_BB3(resto)).strip(),
             'valor': valor, 'tipo': 'C' if letra in ('C', '€') else 'D', 'obs': '', 'incerta': False}
        if rende:
            t['oculta'] = True
        if data_ruim:
            t['incerta'] = True
            t['obs'] = '[A VERIFICAR - OCR] data ilegivel (' + data_str[:5] + ')'
        if not rende:
            transacoes.append(t)
        conf.tx(t)
        if s is not None:
            conf.saldo(_bb3_assinado(_dec(s), sletra))
    conf.fim()
    return transacoes, _montar_aviso('Banco do Brasil', transacoes, conf, n_nao_lidas)


# ------------------------------------------------ Caixa (Gerenciador Caixa)
def identificar_caixa_gerenciador_ocr(linhas):
    t = _sem_espaco(' '.join(linhas))
    # (o Tesseract embaralha o logo 'GERENCIADOR CAIXA' e parte 'Extrato no periodo': nao dependem dele)
    return 'DATAEFETIVA' in t and 'PERIODOSOLICITADO' in t and 'SALDOANTERIOR' in t and 'R$' in ' '.join(linhas)


_RE_CAIXA3 = re.compile(r'^(\d{2}/\d{2}/\d{4})\s*(.*?)\s*(-?)\s*R\$\s*(' + _NUM + r')\s*R\$\s*(' + _NUM +
                        r')\s*([CD€])?$')
_RE_CAIXA3_SALDO = re.compile(r'^(\d{2}/\d{2}/\d{4})\s*SALDO\s*DIA\s*R\$\s*(' + _NUM + r')\s*([CD€])?$', re.IGNORECASE)
_RE_CAIXA3_ANT = re.compile(r'SALDO\s*ANTERIOR.*?R\$\s*(' + _NUM + r')\s*([CD€])?', re.IGNORECASE)
_HIST_CAIXA3 = _restaurador(['CRED PIX QR COD EST', 'PIX RECEBIDO DADOS CONTA', 'CRED PIX CHAVE', 'PIX RECEBIDO',
                             'DEB PIX CHAVE', 'ENVIO DE TED', 'APLIC AUTOMATICA FUNDO', 'DEB PIX QR COD',
                             'DEBITO AUTOMATICO', 'TARIFA'])
_CAIXA3_CABECALHO = ('DATADOCUMENTO', 'DATAEFETIVA', 'DOCUMENTOHISTORICO')


def parse_caixa_gerenciador_ocr(linhas):
    """Caixa 'Gerenciador Caixa': cada lancamento tem o historico numa linha, e na
    seguinte 'DD/MM/AAAA doc detalhe [-]R$ valor R$ saldo C|D' (a letra C/D e o sinal
    do SALDO; o debito vem com '-' antes do valor). Cada lancamento e conferido contra
    o saldo corrente da propria linha."""
    transacoes = []
    base = None
    for l in linhas:
        m = _RE_CAIXA3_ANT.search(l)
        if m and 'EXTRATONO' not in _sem_espaco(l):
            base = _dec(m.group(1)) * _sinal_letra(m.group(2) or 'C')
            break
    conf = _Conferidor(base=base)
    n_nao_lidas = 0
    hist_prev = ''
    for l in linhas:
        l = l.strip()
        ms = _RE_CAIXA3_SALDO.match(l)
        if ms:
            conf.saldo(_dec(ms.group(2)) * _sinal_letra(ms.group(3) or 'C'))
            hist_prev = ''
            continue
        m = _RE_CAIXA3.match(l)
        if m:
            data_str, det, neg, v, s, sletra = m.groups()
            dd, mm, yyyy = data_str.split('/')
            data_ruim = False
            try:
                dt = date(int(yyyy), int(mm), int(dd))
            except ValueError:
                dt = transacoes[-1]['data'] if transacoes else date.today()
                data_ruim = True
            det = det.strip()
            hist = (_HIST_CAIXA3(hist_prev) + ' - ' + det).strip(' -') if hist_prev else det
            t = {'data': dt, 'historico': hist, 'valor': _dec(v), 'tipo': 'D' if neg else 'C',
                 'obs': '', 'incerta': False}
            if data_ruim:
                t['incerta'] = True
                t['obs'] = '[A VERIFICAR - OCR] data ilegivel (' + data_str[:5] + ')'
            transacoes.append(t)
            conf.tx(t)
            conf.saldo(_dec(s) * _sinal_letra(sletra or 'C'))
            hist_prev = ''
            continue
        if _sc_parece_lancamento(l) and re.search(r'R\$', l):
            n_nao_lidas += 1
            hist_prev = ''
            continue
        ne = _sem_espaco(l)
        if (not re.match(r'^\d{2}/\d{2}', l) and not re.match(r'^E\d{10,}', ne) and ne
                and not any(ne.startswith(c) for c in _CAIXA3_CABECALHO) and 'R$' not in l
                and len(ne) <= 60):
            hist_prev = l
    conf.fim()
    return transacoes, _montar_aviso('Caixa', transacoes, conf, n_nao_lidas)


# despacho por layout: o layout antigo de cada banco continua com o leitor de sempre
def _parse_bb_ocr_todos(linhas):
    if identificar_bb_balancete_ocr(linhas):
        return parse_bb_balancete_ocr(linhas)
    return parse_bb_ocr(linhas)


def _parse_caixa_ocr_todos(linhas):
    if identificar_caixa_gerenciador_ocr(linhas):
        return parse_caixa_gerenciador_ocr(linhas)
    return parse_caixa_app_ocr(linhas)


# layouts melhor lidos pelo RapidOCR (ocr_rapid.py): (banco, detector)
RAPID_DETECTORES = [
    ('sicoob', identificar_sicoob_ocr),
    ('pagbank', identificar_pagbank_ocr),
    ('bb', identificar_bb_balancete_ocr),
    ('caixa', identificar_caixa_gerenciador_ocr),
]


DETECTORES_OCR = [
    ('caixa', identificar_caixa_app),
    ('bb', identificar_bb_ocr),
    ('bb', identificar_bb_ocr_dep_origem),
    ('sicoob', identificar_sicoob_ocr),
    ('pagbank', identificar_pagbank_ocr),
    ('bb', identificar_bb_balancete_ocr),
    ('caixa', identificar_caixa_gerenciador_ocr),
]

PARSERS_OCR = {
    'caixa': _parse_caixa_ocr_todos,
    'bb': _parse_bb_ocr_todos,
    'sicoob': parse_sicoob_ocr,
    'pagbank': parse_pagbank_ocr,
}
