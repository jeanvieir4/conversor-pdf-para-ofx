# -*- coding: utf-8 -*-
"""
Conversor PDF para OFX - Jean Vieira
Uso: python converter.py caminho_do_extrato.pdf [outro.pdf ...]
Gera um .xlsx por PDF de entrada, na mesma pasta, com 4 colunas:
Data (dd/mm/aaaa), Historico, Valor (positivo, virgula decimal), Tipo (C/D).
"""
import sys
import os
import re
import glob
import threading
import urllib.request
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
VERSAO_ATUAL = '2.0'
_REPO_GITHUB = 'jeanvieir4/conversor-pdf-para-ofx'
_URL_VERSION = f'https://raw.githubusercontent.com/{_REPO_GITHUB}/main/VERSION'
_URL_DOWNLOAD = f'https://github.com/{_REPO_GITHUB}/releases/download/1.0/Conversor_Extratos.zip'

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
    blocos = extrair_blocos_por_banco(caminho_pdf)
    transacoes_por_banco = {}
    avisos = []
    total = 0
    for banco, texto in blocos:
        if banco is None:
            if not texto.strip():
                avisos.append(('(sem texto no PDF)',
                                'Esse PDF parece ser uma imagem escaneada (sem nenhum texto '
                                'selecionavel). O programa nao le extrato digitalizado, so PDF '
                                'gerado digitalmente pelo banco.'))
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

    caminho_saida = os.path.join(pasta_saida, f'{nome_base}.xlsx')
    gerar_excel(transacoes_por_banco, avisos, caminho_saida)

    arquivos_ofx = []
    for banco, transacoes in transacoes_por_banco.items():
        titulo_banco = NOMES_BANCO.get(banco, banco)
        caminho_ofx = os.path.join(pasta_saida, f'{nome_base}_{titulo_banco}.ofx')
        if gerar_ofx(transacoes, banco, caminho_ofx):
            arquivos_ofx.append(caminho_ofx)

    return caminho_saida, total, avisos, arquivos_ofx


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
        print()
        print('============================================')
        print(f"Nova versao disponivel: v{resultado_update['nova_versao']} (voce esta usando a v{VERSAO_ATUAL})")
        print(f'Baixe em: {_URL_DOWNLOAD}')
        print('============================================')


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

    resultado_update = {}
    thread_update = threading.Thread(target=_verificar_atualizacao, args=(resultado_update,), daemon=True)
    thread_update.start()

    print('============================================')
    print('  Conversor PDF para OFX - Jean Vieira')
    print(f'  Versao {VERSAO_ATUAL}')
    print('============================================')

    if len(sys.argv) > 1:
        # PDFs arrastados e soltos em cima do .exe (ou passados na linha de comando)
        arquivos = sys.argv[1:]
    else:
        # Duplo clique sem arrastar nada: processa todo PDF que estiver na mesma pasta
        arquivos = sorted(glob.glob(os.path.join(pasta_programa, '*.pdf')))

    if not arquivos:
        print()
        print('Nenhum arquivo PDF encontrado.')
        print(f'Coloque os extratos em PDF nesta pasta ({pasta_programa})')
        print('e execute o programa novamente, ou arraste os PDFs em cima do .exe.')
        _mostrar_aviso_atualizacao(resultado_update, thread_update)
        if rodando_como_exe:
            input('\nPressione Enter para sair...')
        sys.exit(1)

    pasta_saida = os.environ.get('PASTA_SAIDA') or pasta_programa
    erros = 0
    for caminho in arquivos:
        try:
            saida, n, avisos, arquivos_ofx = processar_pdf(caminho, pasta_saida)
            print(f'{caminho} -> {saida} ({n} lancamentos, {len(avisos)} avisos, {len(arquivos_ofx)} ofx gerado(s))')
        except Exception as e:
            erros += 1
            print(f'ERRO ao processar "{caminho}": {e}')

    print()
    print('============================================')
    print(f'Concluido. {len(arquivos)} arquivo(s) processado(s), {erros} erro(s).')
    print('Os arquivos .xlsx e .ofx foram gerados nesta mesma pasta.')
    print('============================================')
    _mostrar_aviso_atualizacao(resultado_update, thread_update)
    if rodando_como_exe:
        input('\nPressione Enter para sair...')
    sys.exit(1 if erros else 0)
