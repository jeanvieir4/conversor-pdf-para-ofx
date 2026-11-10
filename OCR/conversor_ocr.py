# -*- coding: utf-8 -*-
"""
Conversor de extrato digitalizado (OCR) para Excel + OFX - Jean Vieira
Ferramenta SEPARADA do Conversor_Extratos principal. So usa essa aqui
quando o PDF for uma imagem escaneada/foto/print sem texto (o conversor
principal ja avisa quando e esse o caso). Ve o aviso de responsabilidade
no console e na planilha antes de confiar em qualquer resultado - OCR
erra, mesmo quando o numero parece certo. Ver "Extrato digitalizado
(OCR)" no CONTEXTO_PROJETO.md pra entender as limitacoes conhecidas.

Uso: python conversor_ocr.py caminho_do_extrato.pdf [outro.pdf ...]
"""
import sys
import os
import glob
import io
import threading

import fitz
import pytesseract
from PIL import Image
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from bancos_ocr import linhas_por_posicao, DETECTORES_OCR, PARSERS_OCR
from converter import gerar_excel, NOMES_BANCO
from ofx_export import gerar_ofx
import atualizador
import relatorio

# Sempre que gerar uma nova versao do OCR: atualizar esse numero E o arquivo
# VERSION_OCR no repositorio (mesmo numero), e o hash em SHA256_OCR.
VERSAO_ATUAL = '1.4'
_REPO_GITHUB = 'jeanvieir4/conversor-pdf-para-ofx'
_URL_VERSION = f'https://raw.githubusercontent.com/{_REPO_GITHUB}/main/VERSION_OCR'
_URL_HASH = f'https://raw.githubusercontent.com/{_REPO_GITHUB}/main/SHA256_OCR'
_URL_DOWNLOAD = f'https://github.com/{_REPO_GITHUB}/releases/download/ocr-1.0/Conversor_OCR.zip'
_NOME_EXE = 'Conversor_OCR.exe'
_INFO_UPDATE = {}  # preenchido pela verificacao de versao; usado nos .txt

_PREENCHIMENTO_INCERTA = PatternFill(start_color='F4CCCC', end_color='F4CCCC', fill_type='solid')


def _destacar_linhas_incertas(caminho_xlsx):
    """Colore de vermelho claro as linhas marcadas [A VERIFICAR] - o
    `gerar_excel` do converter.py e generico (reusado tambem pelo
    conversor principal) e nao faz esse tipo de destaque condicional,
    entao aplica por cima aqui, so nesse arquivo, sem tocar em
    converter.py."""
    wb = load_workbook(caminho_xlsx)
    for nome_aba in wb.sheetnames:
        if nome_aba == 'Pendencias':
            continue
        ws = wb[nome_aba]
        for row in ws.iter_rows(min_row=2):
            obs = row[4].value if len(row) > 4 else None
            if obs and 'A VERIFICAR' in str(obs):
                for cel in row:
                    cel.fill = _PREENCHIMENTO_INCERTA
    wb.save(caminho_xlsx)

ESCALA_UPSCALE = 2  # ampliar a imagem antes do OCR melhora a leitura de C/D colado no valor
DPI_RENDER = 300


def _pasta_do_programa():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _configurar_tesseract():
    """Quando empacotado (.exe), o Tesseract vai embutido dentro do
    bundle do PyInstaller - aponta pytesseract pra ele em vez de
    depender de uma instalacao separada no PC de quem usa."""
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
        tesseract_exe = os.path.join(base, 'tesseract', 'tesseract.exe')
        if os.path.exists(tesseract_exe):
            pytesseract.pytesseract.tesseract_cmd = tesseract_exe
            os.environ['TESSDATA_PREFIX'] = os.path.join(base, 'tesseract', 'tessdata')


def _extrair_linhas_pdf(caminho_pdf):
    """Rasteriza cada pagina do PDF e reconstroi as linhas de texto via
    OCR + posicao (ver bancos_ocr.linhas_por_posicao)."""
    doc = fitz.open(caminho_pdf)
    todas_linhas = []
    for page in doc:
        pix = page.get_pixmap(dpi=DPI_RENDER)
        img = Image.open(io.BytesIO(pix.tobytes('png')))
        if ESCALA_UPSCALE != 1:
            img = img.resize((img.width * ESCALA_UPSCALE, img.height * ESCALA_UPSCALE), Image.LANCZOS)
        todas_linhas.extend(linhas_por_posicao(img, escala=ESCALA_UPSCALE))
    doc.close()
    return todas_linhas


def identificar_banco_ocr(linhas):
    for nome_banco, detector in DETECTORES_OCR:
        if detector(linhas):
            return nome_banco
    return None


def _relatorio(caminho_pdf, pasta_saida, sufixo, titulo, o_que_houve, o_que_fazer):
    return relatorio.escrever_relatorio(caminho_pdf, pasta_saida, sufixo, titulo, o_que_houve,
                                        o_que_fazer, _INFO_UPDATE, VERSAO_ATUAL, _URL_DOWNLOAD)


def processar_pdf_ocr(caminho_pdf, pasta_saida):
    nome_base = os.path.splitext(os.path.basename(caminho_pdf))[0]
    linhas = _extrair_linhas_pdf(caminho_pdf)
    banco = identificar_banco_ocr(linhas)

    if banco is None:
        txt = _relatorio(
            caminho_pdf, pasta_saida, '_ERRO', 'NÃO FOI POSSÍVEL GERAR O ARQUIVO .OFX',
            'O programa leu a imagem do extrato, mas o layout não foi reconhecido. Hoje o OCR só entende '
            'extratos da Caixa (aplicativo) e do Banco do Brasil (internet banking e extrato de conta '
            'corrente escaneado). Pode ser outro banco ou um modelo novo.',
            '1) Confira se a imagem está nítida, reta e completa (sem cortes).\n'
            '2) Veja abaixo se existe versão nova do programa e atualize.\n'
            '3) Se persistir, envie este PDF para o Jean (contato no fim deste arquivo) para incluir '
            'suporte a este layout.')
        return txt, 0, [('(banco nao identificado - OCR)', 'layout nao reconhecido')], []

    parser = PARSERS_OCR[banco]
    transacoes, aviso = parser(linhas)
    avisos = [(banco, aviso)] if aviso else []

    # transacoes com sinal incerto NUNCA entram no ofx - nao da pra inventar
    # credito ou debito num arquivo que alimenta o sistema contabil direto.
    transacoes_confiaveis = [t for t in transacoes if not t.get('incerta')]
    incertas = [t for t in transacoes if t.get('incerta')]
    arquivos_ofx = []
    if transacoes_confiaveis:
        titulo_banco = NOMES_BANCO.get(banco, banco)
        caminho_ofx = os.path.join(pasta_saida, f'{nome_base}_{titulo_banco}.ofx')
        if gerar_ofx(transacoes_confiaveis, banco, caminho_ofx):
            arquivos_ofx.append(caminho_ofx)

    if arquivos_ofx:
        if incertas or aviso:
            detalhe = (aviso or '') + '\n'
            if incertas:
                detalhe += ('\nLINHAS QUE FICARAM DE FORA DO .OFX (não deu para confirmar o valor/sinal '
                            'pela leitura da imagem) - lance à mão depois de conferir no extrato:\n')
                for t in incertas:
                    detalhe += (f"  {t['data'].strftime('%d/%m/%Y')}  {t['tipo']}  "
                                f"{str(t['valor']).replace('.', ',')}  {t['historico']}\n")
            _relatorio(caminho_pdf, pasta_saida, '_AVISOS',
                       'Avisos sobre a conversão (o arquivo .ofx FOI gerado)', detalhe,
                       'Confira TODOS os lançamentos do .ofx contra o extrato original antes de importar: '
                       'a leitura por OCR pode errar, mesmo em números que parecem certos.')
        return arquivos_ofx[0], len(transacoes), avisos, arquivos_ofx

    txt = _relatorio(
        caminho_pdf, pasta_saida, '_ERRO', 'NÃO FOI POSSÍVEL GERAR O ARQUIVO .OFX',
        'O layout do extrato foi reconhecido, mas nenhum lançamento pôde ser lido com segurança.\n\n'
        + (aviso or ''),
        '1) Confira se a imagem está nítida, reta e completa.\n'
        '2) Veja abaixo se existe versão nova do programa e atualize.\n'
        '3) Se persistir, envie este PDF para o Jean (contato no fim deste arquivo).')
    return txt, len(transacoes), avisos, arquivos_ofx


def _pausar():
    """Espera Enter no fim (duplo clique); nao quebra se nao houver teclado/console."""
    try:
        input('\nPressione Enter para sair...')
    except (EOFError, OSError):
        pass


if __name__ == '__main__':
    _configurar_tesseract()
    pasta_programa = _pasta_do_programa()

    resultado_update = _INFO_UPDATE
    thread_update = threading.Thread(target=atualizador.verificar_versao,
                                     args=(_URL_VERSION, VERSAO_ATUAL, resultado_update), daemon=True)
    thread_update.start()

    print('============================================')
    print('  Conversor OCR (extrato digitalizado) - Jean Vieira')
    print(f'  Versao {VERSAO_ATUAL}')
    print('============================================')
    print()
    print('ATENCAO: essa ferramenta le extrato por IMAGEM (OCR), nao por')
    print('texto nativo do PDF. Erros de leitura SAO ESPERADOS, mesmo em')
    print('numeros que parecem corretos. Toda linha gerada vem marcada')
    print('"Gerado por OCR" e precisa ser conferida contra o extrato')
    print('original antes de usar em qualquer lancamento contabil.')
    print('============================================')
    print()

    # Atualizacao automatica (com confirmacao): ver atualizador.py.
    atualizador.limpar_restos(_NOME_EXE)
    thread_update.join(timeout=2.5)
    if resultado_update.get('nova_versao'):
        atualizador.oferecer_atualizacao(_NOME_EXE, _URL_DOWNLOAD, _URL_HASH, VERSAO_ATUAL,
                                         resultado_update['nova_versao'], _URL_DOWNLOAD)

    if len(sys.argv) > 1:
        arquivos = sys.argv[1:]
    else:
        arquivos = sorted(glob.glob(os.path.join(pasta_programa, '*.pdf')))

    if not arquivos:
        print('Nenhum arquivo PDF encontrado.')
        print(f'Coloque os extratos digitalizados nesta pasta ({pasta_programa})')
        print('e execute o programa novamente, ou arraste os PDFs em cima do .exe.')
        _pausar()
        sys.exit(0)

    total_arquivos = 0
    total_erros = 0
    for caminho in arquivos:
        pasta_saida = os.path.dirname(os.path.abspath(caminho))
        try:
            saida, total, avisos, ofx = processar_pdf_ocr(caminho, pasta_saida)
            if ofx:
                print(f'OK: {caminho} -> {total} lancamento(s) lido(s), {len(ofx)} arquivo(s) .ofx gerado(s)'
                      + (' (veja tambem o arquivo _AVISOS.txt)' if avisos else ''))
            else:
                total_erros += 1
                print(f'NAO GEROU OFX: {caminho}')
                print(f'   Motivo e o que fazer: abra o arquivo {saida}')
                continue
        except Exception as e:
            total_erros += 1
            try:
                txt = _relatorio(
                    caminho, pasta_saida, '_ERRO', 'NÃO FOI POSSÍVEL GERAR O ARQUIVO .OFX',
                    'O programa não conseguiu abrir ou ler este PDF.\n'
                    'Causas mais comuns: arquivo corrompido ou incompleto, PDF protegido por senha, ou '
                    f'arquivo aberto/bloqueado por outro programa.\n\nDetalhe técnico: {e}',
                    'Abra o PDF para ver se ele abre normalmente e tente de novo. Veja abaixo se existe '
                    'versão nova do programa.')
                print(f'NAO GEROU OFX: {caminho}')
                print(f'   Motivo e o que fazer: abra o arquivo {txt}')
            except Exception:
                print(f'ERRO ao processar {caminho}: {e}')
            continue
        total_arquivos += 1

    print()
    print('============================================')
    print(f'Concluido. {total_arquivos} arquivo(s) processado(s), {total_erros} erro(s).')
    print('Os arquivos .ofx (e, se houve problema, o .txt explicando) estao nesta mesma pasta.')
    print('LEMBRETE: confira os valores contra o extrato original antes de usar.')
    print('============================================')
    if resultado_update.get('nova_versao'):
        print(f"Nova versao disponivel: v{resultado_update['nova_versao']} (voce esta usando a v{VERSAO_ATUAL})")
        print(f'Baixe em: {_URL_DOWNLOAD}')
    _pausar()
