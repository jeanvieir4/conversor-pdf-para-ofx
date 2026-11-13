# -*- coding: utf-8 -*-
"""
Motor de OCR usado pelo Conversor_Extratos quando o PDF e uma IMAGEM
(extrato escaneado, foto ou print, sem texto selecionavel). Importado so
quando necessario (import preguicoso em converter.py), pra o programa
continuar funcionando mesmo se o OCR nao estiver disponivel.

Os parsers por banco ficam em OCR/bancos_ocr.py e NUNCA reusam a logica de
bancos.py (texto nativo e confiavel; OCR erra) - ver CONTEXTO_PROJETO.md.
"""
import io
import os
import sys

import fitz
import pytesseract
from PIL import Image

_AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(_AQUI, 'OCR'))
from bancos_ocr import (linhas_por_posicao, DETECTORES_OCR, PARSERS_OCR,  # noqa: E402
                        RAPID_DETECTORES, PARSERS_RAPID)

ESCALA_UPSCALE = 2  # ampliar a imagem antes do OCR melhora a leitura de C/D colado no valor
DPI_RENDER = 300


def configurar_tesseract():
    """Empacotado (.exe): o Tesseract vai dentro do bundle do PyInstaller.
    Rodando via Python: usa a copia em OCR/vendor, se existir."""
    if getattr(sys, 'frozen', False):
        base = os.path.join(sys._MEIPASS, 'tesseract')
    else:
        base = os.path.join(_AQUI, 'OCR', 'vendor', 'tesseract')
    exe = os.path.join(base, 'tesseract.exe')
    if os.path.exists(exe):
        pytesseract.pytesseract.tesseract_cmd = exe
        os.environ['TESSDATA_PREFIX'] = os.path.join(base, 'tessdata')


def _extrair_linhas_pdf(caminho_pdf, max_paginas=None):
    """Rasteriza cada pagina do PDF e reconstroi as linhas de texto via
    OCR + posicao (ver bancos_ocr.linhas_por_posicao). `max_paginas` limita
    a leitura as primeiras paginas (usado so pra identificar o banco)."""
    doc = fitz.open(caminho_pdf)
    todas_linhas = []
    for page in doc:
        if max_paginas is not None and page.number >= max_paginas:
            break
        pix = page.get_pixmap(dpi=DPI_RENDER)
        img = Image.open(io.BytesIO(pix.tobytes('png')))
        if ESCALA_UPSCALE != 1:
            img = img.resize((img.width * ESCALA_UPSCALE, img.height * ESCALA_UPSCALE), Image.LANCZOS)
        todas_linhas.extend(linhas_por_posicao(img, escala=ESCALA_UPSCALE))
    doc.close()
    return todas_linhas


def ler_pdf_ocr(caminho_pdf):
    """Retorna (banco, transacoes, aviso). banco e None se o layout nao foi
    reconhecido. Transacoes com 'incerta' True nao devem ir pro OFX."""
    configurar_tesseract()
    # Layouts validados com o RapidOCR (Sicoob, PagBank, BB com 'Dt. balancete' e
    # Caixa 'Gerenciador'): a 1a pagina (so ela, rapido) identifica o layout e o
    # resto e lido pelo RapidOCR, bem mais preciso que o Tesseract. Se o RapidOCR
    # faltar ou falhar, segue o fluxo normal (Tesseract) abaixo.
    erro_rapid = None
    banco_rapid = None
    try:
        linhas_p1 = _extrair_linhas_pdf(caminho_pdf, max_paginas=1)
        banco_rapid = next((nome for nome, det in RAPID_DETECTORES if det(linhas_p1)), None)
        if banco_rapid:
            import ocr_rapid
            linhas_rapid = ocr_rapid.linhas_pdf(caminho_pdf, restaurar_sicoob=(banco_rapid == 'sicoob'))
            if any(nome == banco_rapid and det(linhas_rapid) for nome, det in RAPID_DETECTORES):
                parser = PARSERS_RAPID.get(banco_rapid, PARSERS_OCR[banco_rapid])
                transacoes, aviso = parser(linhas_rapid)
                return banco_rapid, transacoes, aviso
            erro_rapid = 'RapidOCR nao reconheceu o layout'
    except Exception as e:
        erro_rapid = f'{type(e).__name__}: {e}'
    linhas = _extrair_linhas_pdf(caminho_pdf)
    banco = None
    for nome, detector in DETECTORES_OCR:
        if detector(linhas):
            banco = nome
            break
    if banco is None:
        return None, [], None
    transacoes, aviso = PARSERS_OCR[banco](linhas)
    if erro_rapid and banco == banco_rapid:
        aviso = (aviso or '') + (f' [O leitor de OCR preciso nao funcionou ({erro_rapid}); foi usado o leitor '
                                 f'basico, que erra mais - confira com cuidado extra.]')
    return banco, transacoes, aviso
