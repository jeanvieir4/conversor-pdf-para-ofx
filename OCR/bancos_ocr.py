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


DETECTORES_OCR = [
    ('caixa', identificar_caixa_app),
    ('bb', identificar_bb_ocr),
]

PARSERS_OCR = {
    'caixa': parse_caixa_app_ocr,
    'bb': parse_bb_ocr,
}
