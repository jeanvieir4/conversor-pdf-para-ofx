# -*- coding: utf-8 -*-
"""
Segundo motor de OCR (RapidOCR / ONNX, offline) - usado so para os layouts
em que ele foi validado (hoje: Sicoob). Le muito melhor numeros e tabelas do
que o Tesseract (teste com 3 extratos Sicoob escaneados: 188/188 lancamentos
conferidos contra a leitura direta, contra 126/188 do Tesseract), mas e mais
lento. Importado so quando precisa (import preguicoso); se faltar ou falhar,
quem chama volta pro Tesseract.

Limitacao conhecida do RapidOCR: perde os espacos entre palavras
('PIXRECEB.OUTRAIF'). Os parsers de OCR do Sicoob ja toleram isso e o
historico padrao do Sicoob e restaurado por `_restaurar_historico`.
"""
import io
import re
import unicodedata

import fitz
import numpy as np
from PIL import Image

DPI_RENDER = 150

# Historicos padrao do Sicoob (texto fixo do sistema). Restaurar os espacos so
# deixa o .ofx legivel; nao afeta valor/data/sinal.
HISTORICOS_SICOOB = [
    'APLICAÇÃO FINANCEIRA', 'CAPITALIZAÇÃO DE CORREÇÃO MONETÁRIA', 'COMP MASTER MAESTRO',
    'DEP DIN AG', 'DÉB.CNV.EN.ELET.GAS', 'DÉB.CONV.SANEAMENTO', 'DÉB.EMPRÉSTIMO',
    'DÉB.PGTO.BOLETO INT', 'DÉB.TIT.COMPE.EFETI', 'EST.DÉB.EMPRÉSTIMO', 'PIX EMIT.OUTRA IF',
    'PIX REC.OUTRA IF MT', 'PIX RECEB.OUTRA IF', 'PIX.EMIT.OUT IF-MSM',
    'RESGATE DE APLICAÇÃO FINANCEIRA', 'RETENÇÃO DE IRRF', 'TRANSF. PIX SICOOB',
    'TRANSF.RECEB-PIX SI', 'SALDO ANTERIOR', 'SALDO BLOQ.ANTERIOR', 'SALDO DO DIA',
]

_motor = None


def _sem_acento(c):
    return ''.join(x for x in unicodedata.normalize('NFD', c) if unicodedata.category(x) != 'Mn')


def _padrao_historico(texto):
    partes = []
    for ch in texto:
        if ch == ' ':
            continue
        base = _sem_acento(ch)
        partes.append('[' + re.escape(ch) + re.escape(base) + ']' if base != ch else re.escape(ch))
    return re.compile(r'\s*'.join(partes), re.IGNORECASE)


# mais longos primeiro (ex.: 'RESGATE DE APLICACAO FINANCEIRA' antes de 'APLICACAO FINANCEIRA')
_PADROES = sorted(((_padrao_historico(h), h) for h in HISTORICOS_SICOOB), key=lambda x: -len(x[1]))
_RE_INICIO = re.compile(r'^(\d{2}/\d{2}(?:/\d{4})?\s*)(.*)$')


def _restaurar_historico(linha):
    m = _RE_INICIO.match(linha)
    if not m:
        return linha
    resto = m.group(2)
    for pad, canon in _PADROES:
        novo, n = pad.subn(canon, resto, count=1)
        if n:
            resto = novo
            break
    return m.group(1).rstrip() + ' ' + resto


def _limpar_numeros(linha):
    """Cola numeros que o OCR partiu com espaco: '55.114, 04' -> '55.114,04'."""
    linha = re.sub(r'(?<=\d)\s+(?=,\d{2}(?!\d))', '', linha)
    linha = re.sub(r'(?<=\d,)\s+(?=\d{2}(?!\d))', '', linha)
    linha = re.sub(r'(?<=\d\.)\s+(?=\d{3}(?!\d))', '', linha)
    return linha


def _obter_motor():
    global _motor
    if _motor is None:
        from rapidocr_onnxruntime import RapidOCR
        _motor = RapidOCR()
    return _motor


def linhas_pdf(caminho_pdf, restaurar_sicoob=True):
    """Rasteriza o PDF, roda o RapidOCR e junta as caixas de texto em linhas
    (mesma altura = mesma linha). Retorna lista de strings."""
    motor = _obter_motor()
    linhas = []
    doc = fitz.open(caminho_pdf)
    try:
        for pg in doc:
            pix = pg.get_pixmap(dpi=DPI_RENDER)
            img = np.array(Image.open(io.BytesIO(pix.tobytes('png'))).convert('RGB'))
            res, _ = motor(img)
            itens = []
            for box, txt, _sc in (res or []):
                ys = [p[1] for p in box]
                xs = [p[0] for p in box]
                itens.append((sum(ys) / 4, min(xs), max(ys) - min(ys), txt.strip()))
            itens.sort()
            atual, yref, alt = [], None, None
            for cy, x, h, t in itens:
                if yref is None or abs(cy - yref) <= max(8, 0.6 * alt):
                    atual.append((x, t))
                    if yref is None:
                        yref, alt = cy, h
                else:
                    linhas.append(' '.join(s for _, s in sorted(atual)))
                    atual, yref, alt = [(x, t)], cy, h
            if atual:
                linhas.append(' '.join(s for _, s in sorted(atual)))
    finally:
        doc.close()
    linhas = [_limpar_numeros(l) for l in linhas]
    return [_restaurar_historico(l) for l in linhas] if restaurar_sicoob else linhas
