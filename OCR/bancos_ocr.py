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
            obs = 'Gerado por OCR - confira contra o extrato original antes de usar.'

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


DETECTORES_OCR = [
    ('caixa', identificar_caixa_app),
]

PARSERS_OCR = {
    'caixa': parse_caixa_app_ocr,
}
