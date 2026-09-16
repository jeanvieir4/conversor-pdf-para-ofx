# -*- coding: utf-8 -*-
"""
Parsers por banco para o conversor de extrato PDF -> Excel de Jean Vieira.
Cada parser recebe o texto bruto (pdfplumber) de UM extrato de UM banco e
devolve uma lista de dicts: {data: date, historico: str, valor: Decimal, tipo: 'C'|'D', obs: str}
'obs' e usado para sinalizar [A VERIFICAR] quando a classificacao C/D nao e 100% confiavel.
"""
import re
from decimal import Decimal, InvalidOperation
from datetime import date

MESES_PT = {'jan':1,'fev':2,'mar':3,'abr':4,'mai':5,'jun':6,'jul':7,'ago':8,'set':9,'out':10,'nov':11,'dez':12}

def _dec(valor_str):
    """Converte '1.234,56' -> Decimal('1234.56'). Nunca usa float."""
    v = valor_str.strip().replace('.', '').replace(',', '.')
    try:
        return Decimal(v)
    except InvalidOperation:
        return None

def _linhas_uteis(texto):
    return [l for l in texto.split('\n') if l.strip()]


# ---------------------------------------------------------------------------
# SICOOB
# ---------------------------------------------------------------------------
def parse_sicoob(texto):
    m_periodo = re.search(r'PER[ÍI]ODO:\s*(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})', texto)
    if not m_periodo:
        return [], 'Sicoob: nao encontrei a linha PERIODO: no extrato, ano das datas fica sem referencia segura'
    ini = date(*map(int, reversed(m_periodo.group(1).split('/'))))
    fim = date(*map(int, reversed(m_periodo.group(2).split('/'))))

    padrao = re.compile(r'^(\d{2})/(\d{2})\s+(.+?)\s+([\d\.]+,\d{2})([CD])$')
    # Quando o valor tem separador de milhar (>= 1.000,00), a coluna do valor
    # fica mais larga e o pdfplumber as vezes perde o alinhamento vertical
    # com a linha "DD/MM descricao" - visto de duas formas diferentes no
    # mesmo extrato:
    #   a) so o C/D "estoura" pra linha seguinte, sozinho:
    #        '05/08 PIX EMIT.OUTRA IF 4.554,00'
    #        'D'
    #   b) o VALOR inteiro fica deslocado pra ANTES da linha data+descricao,
    #      com o C/D ainda depois (bloco de 3 linhas):
    #        '1.200,00'
    #        '03/08 PIX RECEB.OUTRA IF'
    #        'C'
    # Sem reconstruir isso a linha nao bate no `padrao` e o lancamento
    # inteiro era descartado em silencio - bug real reportado pelo Jean
    # (valores >1.000 sumindo do .ofx).
    padrao_sem_cd = re.compile(r'^\d{2}/\d{2}\s+.+?\s+[\d\.]+,\d{2}$')
    padrao_data_desc_sem_valor = re.compile(r'^\d{2}/\d{2}\s+\S.*$')
    padrao_valor_sozinho = re.compile(r'^[\d\.]+,\d{2}$')
    linhas_brutas = _linhas_uteis(texto)
    linhas = []
    i = 0
    while i < len(linhas_brutas):
        l = linhas_brutas[i].strip()
        # caso (b): valor sozinho, seguido de "DD/MM descricao" sem valor, seguido de C/D sozinho
        if (padrao_valor_sozinho.match(l) and i + 2 < len(linhas_brutas)
                and padrao_data_desc_sem_valor.match(linhas_brutas[i + 1].strip())
                and not padrao_sem_cd.match(linhas_brutas[i + 1].strip())
                and linhas_brutas[i + 2].strip() in ('C', 'D')):
            linhas.append(f'{linhas_brutas[i + 1].strip()} {l}{linhas_brutas[i + 2].strip()}')
            i += 3
            continue
        # caso (a): "DD/MM ... valor" (sem C/D) seguido de C/D sozinho
        if padrao_sem_cd.match(l) and i + 1 < len(linhas_brutas) and linhas_brutas[i + 1].strip() in ('C', 'D'):
            linhas.append(l + linhas_brutas[i + 1].strip())
            i += 2
            continue
        linhas.append(l)
        i += 1

    excluir = ('SALDO ANTERIOR', 'SALDO BLOQ.ANTERIOR', 'SALDO DO DIA')
    # "DEP.CHEQUE BLOQ.XD" (cheque em compensacao) nao usa C/D no final, usa
    # um asterisco ('6.254,00*') - e so um AVISO de que o deposito chegou
    # mas ainda esta bloqueado, NAO e credito ainda (nao entra no saldo do
    # dia declarado no extrato). O credito real so acontece quando aparece
    # "LIBER.DEPOSITO BLOQ" (lancamento normal, com C), tipicamente uns
    # dias depois, com o MESMO valor - contar os dois como credito duplica
    # o dinheiro (confirmado no extrato real: DEP.CHEQUE BLOQ.1D de
    # 6.254,00 e 231,00 tem LIBER.DEPOSITO BLOQ correspondente com o mesmo
    # valor depois - contar ambos estourava o saldo em exatamente 6.485,00,
    # a soma dos dois). Por isso as linhas com asterisco sao ignoradas
    # aqui, igual as linhas de SALDO.
    padrao_asterisco = re.compile(r'^\d{2}/\d{2}\s+.+?\s+[\d\.]+,\d{2}\*$')
    out = []
    for l in linhas:
        if padrao_asterisco.match(l):
            continue
        mm_ = padrao.match(l)
        if not mm_:
            continue
        dd, mm, desc, valor_str, tipo = mm_.groups()
        if desc.strip().upper() in excluir:
            continue
        mes = int(mm)
        ano = ini.year if mes >= ini.month else fim.year
        try:
            dt = date(ano, mes, int(dd))
        except ValueError:
            continue
        valor = _dec(valor_str)
        out.append({'data': dt, 'historico': desc.strip(), 'valor': valor, 'tipo': tipo, 'obs': ''})
    return out, None


# ---------------------------------------------------------------------------
# CAIXA (Gerenciador Caixa)
# ---------------------------------------------------------------------------
def parse_caixa(texto):
    padrao = re.compile(
        r'(?P<desc>[^\n]+)\n(?P<data>\d{2}/\d{2}/\d{4})\n(?P<doc>\S+)\s+(?P<det>.+?)\s+R\$\s*(?P<valor>[\d\.]+,\d{2})\s+R\$\s*[\d\.]+,\d{2}\s+(?P<tipo>[CD])'
    )
    excluir = ('SALDO ANTERIOR', 'SALDO DO DIA', 'SALDO ANTERIOR AO PERÍODO SOLICITADO')
    out = []
    for m in padrao.finditer(texto):
        desc = m.group('desc').strip()
        if desc.upper() in excluir or 'SALDO' in desc.upper() and 'DIA' in desc.upper():
            continue
        dd, mm, yyyy = m.group('data').split('/')
        try:
            dt = date(int(yyyy), int(mm), int(dd))
        except ValueError:
            continue
        valor = _dec(m.group('valor'))
        out.append({'data': dt, 'historico': f"{desc} - {m.group('det').strip()}", 'valor': valor,
                     'tipo': m.group('tipo'), 'obs': ''})
    return out, None


# ---------------------------------------------------------------------------
# ITAU
# ---------------------------------------------------------------------------
def parse_itau(texto):
    m_periodo = re.search(r'\b(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\s+(20\d{2})\b', texto, re.IGNORECASE)
    if not m_periodo:
        return [], 'Itau: nao encontrei o mes/ano do extrato (ex: "abr 2026") no cabecalho'
    ano_ref = int(m_periodo.group(2))

    # Linhas de "Aplic Aut Mais" (varredura automatica diaria de saldo para um
    # CDB atrelado a conta) aparecem repetidas vezes ao longo de TODO o extrato,
    # uma vez por dia, e nao so no final. Sao excluidas por nao dar pra confirmar
    # pelo texto extraido se contam como credito ou debito real na conta corrente.
    excluir_kw = ('SALDO ANTERIOR', 'APLIC AUT MAIS', 'SALDO EM C/C', 'SALDO FINAL', 'TOTAL')
    padrao = re.compile(r'^(?P<desc>.+?)\s+(?P<valor>[\d\.]+,\d{2})(?P<deb>-)?(?:\s+[\d\.]+,\d{2})?$')
    out = []
    obs_aplic = False
    dia_atual = None
    for l in _linhas_uteis(texto):
        l = l.strip()
        # Tabelas-resumo que reaparecem 1x no final do extrato (Compras a debito,
        # Cheques compensados, Debitos automaticos) usam data no formato DD/MM/AA
        # (com ano de 2 digitos) em vez do DD/MM usado na movimentacao diaria, e
        # os lancamentos ali ja foram contados na movimentacao principal - pular.
        if re.match(r'^\d{2}/\d{2}/\d{2}\b', l):
            continue
        mdata = re.match(r'^(\d{2})/(\d{2})\s+(.+)$', l)
        desc_completa = l
        if mdata:
            dia_atual = (mdata.group(1), mdata.group(2))
            desc_completa = mdata.group(3)
        # A tabela "movimentacao - aplicacoes/resgates" tambem usa DD/MM no
        # inicio da linha, mas o resto e so numero (sem nenhuma letra) -
        # descricao de transacao real sempre tem letras (PIX, Sispag, Rshop...).
        if not re.search(r'[A-Za-zÀ-ÖØ-öø-ÿ]', desc_completa):
            continue
        m = padrao.match(desc_completa)
        if not m:
            continue
        desc = m.group('desc').strip()
        # Linhas de tabelas-resumo (resumo mensal do CDB, totalizador de
        # aplicacoes) tem varios numeros com vírgula decimal dentro da propria
        # "descricao" (ex: "principal 5.784,11 15.901,17 0,00 16.980,46").
        # Uma descricao de lancamento real nunca tem valor embutido assim.
        if re.search(r'\d,\d{2}', desc):
            continue
        if any(k in desc.upper() for k in excluir_kw):
            if 'APLIC AUT MAIS' in desc.upper():
                obs_aplic = True
            continue
        valor = _dec(m.group('valor'))
        tipo = 'D' if m.group('deb') else 'C'
        if dia_atual is None:
            continue  # ainda nao vimos nenhuma data nesta pagina: nao da pra atribuir com seguranca
        dd, mm = dia_atual
        try:
            dt = date(ano_ref, int(mm), int(dd))
        except ValueError:
            continue
        out.append({'data': dt, 'historico': desc, 'valor': valor, 'tipo': tipo, 'obs': ''})
    obs = ('Itau: linhas de "Aplic Aut Mais" (aplicacao automatica em CDB atrelada a conta) '
           'foram excluidas por nao ser possivel confirmar pelo texto extraido se contam como '
           'credito ou debito na conta corrente - confira manualmente se precisar delas.') if obs_aplic else None
    return out, obs


# ---------------------------------------------------------------------------
# UNICRED
# ---------------------------------------------------------------------------
def parse_unicred(texto):
    linhas = _linhas_uteis(texto)
    out = []
    i = 0
    padrao_data = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(-\s*)?R\$\s*([\d\.]+,\d{2})\s+-?R\$\s*[\d\.]+,\d{2}\s*(.*)$')
    while i < len(linhas):
        l = linhas[i].strip()
        m = padrao_data.match(l)
        if m:
            data_str, sinal, valor_str, resto = m.groups()
            dd, mm, yyyy = data_str.split('/')
            try:
                dt = date(int(yyyy), int(mm), int(dd))
            except ValueError:
                i += 1
                continue
            desc_antes = linhas[i-1].strip() if i > 0 else ''
            desc_depois = resto.strip()
            if not desc_depois and i + 1 < len(linhas):
                prox = linhas[i+1].strip()
                if not padrao_data.match(prox) and 'Saldo do Dia' not in prox and not prox.startswith('Consulta') and not prox.startswith('Periodo') and not prox.startswith('P\u00e1gina'):
                    desc_depois = prox
            desc = ' '.join(p for p in (desc_antes, desc_depois) if p and 'Saldo do Dia' not in p and 'Lan\u00e7amentos' != p).strip(' -(')
            valor = _dec(valor_str)
            tipo = 'D' if sinal else 'C'
            out.append({'data': dt, 'historico': desc, 'valor': valor, 'tipo': tipo, 'obs': ''})
        i += 1
    # remove "lancamentos futuros" (nao compensados ainda) - ficam apos a linha "Saldo no final do periodo"
    if 'Saldo no final do per' in texto:
        corte = texto.split('Saldo no final do per')[1]
        # aproximação: nada a fazer aqui pois já não capturamos texto após o corte na leitura por página
        pass
    return out, None


# ---------------------------------------------------------------------------
# CRESOL
# ---------------------------------------------------------------------------
def _parse_cresol_colmeia(texto):
    """Segundo layout de extrato do Cresol, gerado pelo sistema web 'Colmeia'
    (impressao de pagina, nao o extrato tradicional). Formato bem diferente
    do parse_cresol original:
      '03/08/2026 SALDO ANT.:               15.384,73 C'   <- cabecalho do dia
      'PIX CREDITO                                     '    <- categoria
      '(DE: THIAGO CHAGAS - 01/08)              78,00 C'    <- detalhe + valor
    Descricao longa quebra em 3 linhas, com o valor sozinho na 3a:
      'PAGAMENTO DE TITULOS'
      '(BEVILAQUA CONSTRUTORA E MATERIAL)'
      '                                      1.400,00 D'
    Validado batendo saldo inicial + soma dos lancamentos == saldo final
    declarado no proprio extrato (ver CONTEXTO_PROJETO.md).
    """
    # O rodape de cada pagina (URL do sistema + cabecalho repetido) e a marca
    # d'agua diagonal com o nome de quem tirou o extrato (repetida varias
    # vezes, rotacionada) viram fragmentos curtos quando o pdfplumber padrao
    # nao da conta do layout e caimos no fallback por posicao de caractere
    # (ver _texto_robusto_por_caracteres em converter.py). Sem filtrar isso,
    # uma transacao que cai bem na quebra de pagina herda esse lixo na
    # descricao. Fragmento generico (nao depende do nome de quem gerou o
    # extrato): linha de 4 caracteres uteis ou menos, ou que contenha a URL.
    linhas = [l for l in _linhas_uteis(texto)
              if 'sistema.confesol' not in l.lower()
              and 'sistema de gest' not in l.lower()
              and len(l.strip()) > 4]
    for i, l in enumerate(linhas):
        if 'LANCAMENTOS FUTUROS' in l.upper() or l.strip().startswith('(=)SALDO'):
            linhas = linhas[:i]
            break

    padrao_dia = re.compile(r'^(\d{2})/(\d{2})/(\d{4})\s+SALDO ANT\.:')
    padrao_valor_fim = re.compile(r'([\d.]+,\d{2})\s*([CD])\s*$')

    data_atual = None
    categoria = None
    out = []
    i = 0
    while i < len(linhas):
        l = linhas[i].strip()
        m_dia = padrao_dia.match(l)
        if m_dia:
            dd, mm, yyyy = m_dia.groups()
            try:
                data_atual = date(int(yyyy), int(mm), int(dd))
            except ValueError:
                data_atual = None
            categoria = None
            i += 1
            continue
        if not l:
            i += 1
            continue
        if l.startswith('('):
            m_valor = padrao_valor_fim.search(l)
            if m_valor:
                valor_str, tipo = m_valor.groups()
                desc_extra = l[:m_valor.start()].strip()
                historico = categoria if not desc_extra else f'{categoria} {desc_extra}'.strip()
                out.append({'data': data_atual, 'historico': historico or '(sem descricao)',
                             'valor': _dec(valor_str), 'tipo': tipo, 'obs': ''})
                i += 1
                continue
            # descricao longa: o valor vem sozinho na proxima linha
            if i + 1 < len(linhas):
                m_valor2 = padrao_valor_fim.search(linhas[i + 1].strip())
                if m_valor2:
                    valor_str, tipo = m_valor2.groups()
                    historico = f'{categoria} {l}'.strip() if categoria else l
                    out.append({'data': data_atual, 'historico': historico,
                                 'valor': _dec(valor_str), 'tipo': tipo, 'obs': ''})
                    i += 2
                    continue
            i += 1
            continue
        # linha de categoria (PIX CREDITO, PAGAMENTO DE TITULOS etc) - nao e
        # transacao, so contexto pra descricao da proxima linha "(...)". Toda
        # categoria real vem em CAIXA ALTA; usa isso pra nao deixar um resto
        # de marca d'agua (texto normal, tipo "Extrato emitido por...") que
        # escapou do filtro acima sobrescrever a categoria certa quando uma
        # transacao cai bem na quebra de pagina.
        if l.isupper():
            categoria = l
        i += 1
    return out, None


def _parse_cresol_consolidado_diario(texto):
    """Terceiro layout do Cresol - tambem tem o cabecalho 'EXTRATO
    CONSOLIDADO DE CONTA CORRENTE' (mesmo titulo do layout Colmeia), mas a
    estrutura da linha e bem mais simples: cada lancamento vem inteiro
    numa unica linha, sem quebrar categoria/detalhe/valor em linhas
    separadas:
      '03/08/2026 PIX CREDITO DE: FERNANDA DOMINGUES S - 01/08 220,00 C'
    E 'SALDO ANTERIOR' aparece uma vez por DIA (nao so no inicio do
    periodo), mostrando o saldo de abertura daquele dia - nao e
    lancamento, so o Colmeia usa 'SALDO ANT.:' (abreviado, com ponto e
    dois pontos); esse aqui usa 'SALDO ANTERIOR' por extenso - e esse o
    sinal usado pra rotear entre os dois em `parse_cresol` (ver mais
    abaixo), nunca os dois aparecem juntos no mesmo extrato.
    Termina com um bloco de resumo '(=)SALDO: ...' e depois
    'LANCAMENTOS FUTUROS/PENDENTES' (parcelas/faturas ainda nao
    lancadas, sem 'C'/'D' no final da linha) - tudo isso e ignorado.
    Validado com extrato real: 113 lancamentos, saldo do periodo inicial
    (49.452,12) + creditos - debitos bateu exato com o saldo final
    declarado '(=)SALDO: 44.849,73'.
    """
    padrao = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d.]+,\d{2})\s+([CD])$')
    out = []
    parar = False
    for l in _linhas_uteis(texto):
        l = l.strip()
        if '(=)SALDO' in l or 'LANCAMENTOS FUTUROS' in l.upper():
            parar = True
        if parar:
            continue
        m = padrao.match(l)
        if not m:
            continue
        data_str, desc, valor_str, tipo = m.groups()
        desc = desc.strip()
        if desc.upper() == 'SALDO ANTERIOR':
            continue
        dd, mm, yyyy = data_str.split('/')
        try:
            dt = date(int(yyyy), int(mm), int(dd))
        except ValueError:
            continue
        out.append({'data': dt, 'historico': desc, 'valor': _dec(valor_str), 'tipo': tipo, 'obs': ''})
    return out, None


def parse_cresol(texto):
    if 'SALDO ANT.:' not in texto and 'EXTRATO CONSOLIDADO DE CONTA CORRENTE' in texto.upper():
        return _parse_cresol_consolidado_diario(texto)
    if 'EXTRATO CONSOLIDADO DE CONTA CORRENTE' in texto.upper() or 'sistema.confesol' in texto.lower():
        return _parse_cresol_colmeia(texto)
    linhas = _linhas_uteis(texto)
    # O layout real quebra descricoes longas em 2 linhas, com a linha de
    # data+valor ficando ENTRE elas (por causa do alinhamento visual no PDF):
    #   "PAGAMENTO DE TITULOS - IB WESTHOUSE"
    #   "22/07/2026 - R$ 208,74"
    #   "CLEAN PRODUTOS DE HIGI"
    # Descricoes curtas (uma linha so) vem completas na mesma linha da data:
    #   "22/07/2026 PIX DEBITO PARA: ELIEZER SOUZA - R$ 1.000,00"
    padrao = re.compile(r'^(\d{2}/\d{2}/\d{4})\s*(.*?)\s*([+-])\s*R\$\s*([\d\.]+,\d{2})$')
    ignorar_prefixo = ('Consulta Posi', 'Periodo de', 'P\u00e1gina', 'Lan\u00e7amentos', 'Saldo em Conta',
                        'Limite de Cr', 'Saldo Dispon', 'Ag\u00eancia', 'Saldo Anterior', 'IGREJA', 'Saldo do Dia')
    out = []
    for i, l in enumerate(linhas):
        l = l.strip()
        m = padrao.match(l)
        if not m:
            continue
        if 'Saldo do Dia' in l or 'Saldo Anterior' in l:
            continue
        data_str, desc_inline, sinal, valor_str = m.groups()
        desc_inline = desc_inline.strip()
        if desc_inline:
            desc = desc_inline
        else:
            partes = []
            if i > 0:
                prev = linhas[i - 1].strip()
                if prev and not padrao.match(prev) and not any(prev.startswith(p) for p in ignorar_prefixo):
                    partes.append(prev)
            if i + 1 < len(linhas):
                prox = linhas[i + 1].strip()
                if prox and not padrao.match(prox) and not any(prox.startswith(p) for p in ignorar_prefixo):
                    partes.append(prox)
            desc = ' '.join(partes) if partes else '(descricao nao identificada pelo texto extraido)'
        dd, mm, yyyy = data_str.split('/')
        try:
            dt = date(int(yyyy), int(mm), int(dd))
        except ValueError:
            continue
        valor = _dec(valor_str)
        tipo = 'C' if sinal == '+' else 'D'
        out.append({'data': dt, 'historico': desc, 'valor': valor, 'tipo': tipo, 'obs': ''})
    return out, None


# ---------------------------------------------------------------------------
# AILOS / VIACREDI (mesmo sistema)
# ---------------------------------------------------------------------------
def parse_ailos(texto):
    padrao = re.compile(
        r'^(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\w\.\-]+)\s+(-?[\d\.]+,\d{2})\s+(-?[\d\.]+,\d{2})$'
    )
    excluir = ('SALDO ANTERIOR', 'TOTAL')
    out = []
    for l in _linhas_uteis(texto):
        m = padrao.match(l.strip())
        if not m:
            continue
        data_str, desc, doc, valor_str, saldo_str = m.groups()
        if desc.strip().upper() in excluir:
            continue
        dd, mm, yyyy = data_str.split('/')
        try:
            dt = date(int(yyyy), int(mm), int(dd))
        except ValueError:
            continue
        valor = _dec(valor_str.lstrip('-'))
        tipo = 'D' if valor_str.strip().startswith('-') else 'C'
        out.append({'data': dt, 'historico': desc.strip(), 'valor': valor, 'tipo': tipo, 'obs': ''})
    return out, None


# ---------------------------------------------------------------------------
# BANCO DO BRASIL
# ---------------------------------------------------------------------------
def _parse_bb_2016(texto):
    """Layout mais antigo do extrato do BB (visto em extrato de 2016),
    bem diferente do layout atual com 'Ag. origem'/'Lote':
      'Dt. movimento Dt. balancete Historico Documento Valor R$ Saldo'
      '25/01/2016 Ordem Bancaria 201.601.250.006.341 34.493.267,52 C 34.493.267,52 C'
      '26/01/2016 +Ordem Bancaria 201.601.250.006.350 41.909,68 C'
      'SEFAZ RECURSOS ORDINARIOS'                          <- linha de continuacao
    O saldo corrente (ultimos 2 tokens) so aparece na ULTIMA linha de um
    grupo de lancamentos do mesmo dia - por isso e sempre opcional no
    regex. 'Documento' distingue de 'Valor' por nao ter virgula decimal.
    Validado batendo soma dos creditos - soma dos debitos == saldo final
    'S A L D O' declarado no proprio extrato (ver CONTEXTO_PROJETO.md).
    """
    padrao = re.compile(
        r'^(\d{2}/\d{2}/\d{4})\s+(\+?.+?)\s+([\d.]+)\s+([\d.]+,\d{2})\s+([CD])(?:\s+[\d.]+,\d{2}\s+[CD])?$'
    )
    linhas = _linhas_uteis(texto)
    out = []
    i = 0
    while i < len(linhas):
        l = linhas[i].strip()
        m = padrao.match(l)
        if not m:
            i += 1
            continue
        data_str, historico, documento, valor_str, tipo = m.groups()
        chave = historico.replace(' ', '').upper()
        if chave in ('SALDOANTERIOR', 'SALDO'):
            i += 1
            continue
        historico = historico.lstrip('+').strip()
        if i + 1 < len(linhas):
            prox = linhas[i + 1].strip()
            comeca_com_data = re.match(r'^\d{2}/\d{2}/\d{4}\s', prox)
            if not padrao.match(prox) and not comeca_com_data and prox and not prox.upper().startswith('OBSERVA'):
                historico = f'{historico} {prox}'.strip()
                i += 1
        dd, mm, yyyy = data_str.split('/')
        try:
            dt = date(int(yyyy), int(mm), int(dd))
        except ValueError:
            i += 1
            continue
        out.append({'data': dt, 'historico': historico, 'valor': _dec(valor_str), 'tipo': tipo, 'obs': ''})
        i += 1
    return out, None


def _parse_bb_dia_lote(texto):
    """Terceiro layout do BB, cabecalho 'Dia Lote Documento Historico
    Valor' (extrato tipo internet banking/app, sem 'Dt. movimento'/
    'Dt. balancete' nenhum). Sinal vem como '(+)'/'(-)' no final da
    linha, nao 'C'/'D'. Cada lancamento tem uma linha de "categoria"
    ANTES (ex: 'Pix - Recebido', 'Compra com Cartao') e frequentemente
    uma linha de continuacao (data/hora originais + documento + nome)
    DEPOIS - mas o pdfplumber as vezes funde a continuacao dentro da
    propria linha do lancamento (antes do valor), entao o meio da linha
    (entre o numero de lote e o valor) pode ser so o numero de documento
    OU documento+descricao fundida. Validado batendo saldo inicial + soma
    dos lancamentos == saldo final 'S A L D O' declarado no proprio
    extrato (ver CONTEXTO_PROJETO.md).
    LIMITACAO CONHECIDA: se um lancamento nao tiver NENHUMA linha de
    continuacao nem categoria fundida (2 lancamentos colados sem nada
    entre eles), a categoria do lancamento seguinte pode ser engolida
    como se fosse continuacao deste - o valor/tipo/data continuam
    corretos (bate no total), so a descricao daquele lancamento seguinte
    fica vazia. Nao visto na amostra testada, mas pode acontecer.
    """
    padrao = re.compile(
        r'^(\d{2}/\d{2}/\d{4})\s+(?:(\d+)\s+)?(.*?)\s+([\d.]+,\d{2})\s*\(([+-])\)$'
    )
    linhas = _linhas_uteis(texto)
    out = []
    categoria = None
    i = 0
    while i < len(linhas):
        l = linhas[i].strip()
        m = padrao.match(l)
        if not m:
            categoria = l
            i += 1
            continue

        data_str, _lote, meio, valor_str, sinal = m.groups()
        meio = (meio or '').strip()
        chave = re.sub(r'\s+', '', meio).upper()
        if 'SALDOANTERIOR' in chave or 'SALDODODIA' in chave or chave == 'SALDO':
            categoria = None
            i += 1
            continue

        partes_meio = meio.split(None, 1)
        if partes_meio and partes_meio[0].isdigit():
            resto_meio = partes_meio[1] if len(partes_meio) > 1 else ''
        else:
            resto_meio = meio

        historico = categoria or ''
        if resto_meio:
            historico = f'{historico} {resto_meio}'.strip()

        if i + 1 < len(linhas):
            prox = linhas[i + 1].strip()
            if not padrao.match(prox):
                historico = f'{historico} {prox}'.strip()
                i += 1

        dd, mm, yyyy = data_str.split('/')
        try:
            dt = date(int(yyyy), int(mm), int(dd))
        except ValueError:
            categoria = None
            i += 1
            continue
        out.append({'data': dt, 'historico': historico or '(sem descricao)',
                     'valor': _dec(valor_str), 'tipo': ('C' if sinal == '+' else 'D'), 'obs': ''})
        categoria = None
        i += 1
    return out, None


def _parse_bb_dia_lote_data_separada(texto):
    """Variante do terceiro layout do BB ('Dia Lote Documento') onde a
    extracao de texto do pdfplumber separa a DATA numa linha propria
    (sozinha, ou com uma categoria colada: 'DD/MM/AAAA Transferencia
    recebida'), sem repetir a data na linha de lote/documento/valor -
    diferente do padrao tratado por `_parse_bb_dia_lote` (onde data e
    valor vem na MESMA linha). So e chamada como fallback quando
    `_parse_bb_dia_lote` roda e nao encontra nenhuma transacao nesse
    texto (ver `parse_bb` mais abaixo) - nunca roda no lugar dela.
    Validado batendo saldo anterior + soma dos lancamentos == saldo
    final 'S A L D O' declarado no proprio extrato.
    LIMITACAO CONHECIDA: as linhas de texto livre entre uma transacao e
    a proxima (continuacao da transacao anterior + categoria da
    proxima, sem separador confiavel entre as duas) sao todas
    concatenadas no historico da transacao seguinte - a descricao pode
    sair um pouco misturada, mas data/valor/tipo sempre ficam corretos.
    """
    padrao_data = re.compile(r'^(\d{2}/\d{2}/\d{4})\s*(.*)$')
    padrao_transacao = re.compile(
        r'^(?:(\d+)\s+)?(.*?)\s+([\d.]+,\d{2})\s*\(([+-])\)$'
    )
    linhas = _linhas_uteis(texto)
    out = []
    data_atual = None
    texto_pendente = []

    for linha in linhas:
        l = linha.strip()
        if l == '00/00/0000':
            continue

        m_data = padrao_data.match(l)
        if m_data:
            data_str, sufixo = m_data.groups()
            dd, mm, yyyy = data_str.split('/')
            try:
                data_atual = date(int(yyyy), int(mm), int(dd))
            except ValueError:
                data_atual = None
            if sufixo.strip():
                texto_pendente.append(sufixo.strip())
            continue

        m_trans = padrao_transacao.match(l)
        if m_trans:
            _lote, meio, valor_str, sinal = m_trans.groups()
            meio = (meio or '').strip()
            chave = re.sub(r'\s+', '', meio).upper()
            if 'SALDOANTERIOR' in chave or 'SALDODODIA' in chave or chave == 'SALDO':
                texto_pendente = []
                data_atual = None
                continue

            partes_meio = meio.split(None, 1)
            if partes_meio and partes_meio[0].isdigit():
                resto_meio = partes_meio[1] if len(partes_meio) > 1 else ''
            else:
                resto_meio = meio

            if data_atual is not None:
                historico = ' '.join(texto_pendente + ([resto_meio] if resto_meio else []))
                out.append({'data': data_atual, 'historico': historico.strip() or '(sem descricao)',
                             'valor': _dec(valor_str), 'tipo': ('C' if sinal == '+' else 'D'), 'obs': ''})
            texto_pendente = []
            data_atual = None
            continue

        texto_pendente.append(l)

    return out, None


def parse_bb(texto):
    if 'Dia Lote Documento' in texto:
        transacoes, aviso = _parse_bb_dia_lote(texto)
        if not transacoes:
            transacoes, aviso = _parse_bb_dia_lote_data_separada(texto)
        return transacoes, aviso

    # Os dois layouts abaixo tem 'Dt. movimento' e 'Dt. balancete' no
    # cabecalho (so a ordem muda) - nao da pra distinguir por isso. O
    # layout atual (com 'Ag. origem'/'Lote') tem essas 2 colunas extras
    # que o layout de 2016 nao tem; usa a AUSENCIA delas como sinal.
    if 'Dt. balancete' in texto and 'Ag. origem' not in texto:
        return _parse_bb_2016(texto)

    padrao = re.compile(
        r'^(\d{2}/\d{2}/\d{4})\s+\d+\s+(.+?)\s+([\d\.]+)\s+([\d\.]+,\d{2})\s+([CD])(?:\s+[\d\.]+,\d{2}\s+[CD])?$'
    )
    excluir_kw = ('SALDO ANTERIOR', 'BB RENDE F', 'RENDE FACIL')
    out = []
    obs_rende = False
    for l in _linhas_uteis(texto):
        m = padrao.match(l.strip())
        if not m:
            continue
        data_str, desc, doc, valor_str, tipo = m.groups()
        if any(k in desc.upper() for k in excluir_kw):
            obs_rende = True
            continue
        dd, mm, yyyy = data_str.split('/')
        try:
            dt = date(int(yyyy), int(mm), int(dd))
        except ValueError:
            continue
        valor = _dec(valor_str)
        out.append({'data': dt, 'historico': desc.strip(), 'valor': valor, 'tipo': tipo, 'obs': ''})
    obs = ('BB: linhas "BB Rende Facil" (varredura automatica para aplicacao) foram excluidas pelo '
           'mesmo motivo do Itau - confirme se precisa incluir.') if obs_rende else None
    return out, obs


# ---------------------------------------------------------------------------
# SANTANDER  (sem sinal/coluna confiavel no texto extraido -> classificacao por palavra-chave)
# ---------------------------------------------------------------------------
_KW_DEBITO_SANTANDER = ('PAGAMENTO', 'ENVIADO', 'LIQUIDACAO', 'TARIFA', 'DEBITO', 'BOLETO',
                         'JUROS', 'IOF', 'SAQUE', 'COMPRA')
_KW_CREDITO_SANTANDER = ('RECEBIDO', 'CREDITO', 'DEPOSITO', 'DEP ', 'REND')

def parse_santander(texto):
    m_periodo = re.search(r'Resumo\s*-\s*(\w+)/(\d{4})', texto, re.IGNORECASE)
    if not m_periodo:
        return [], 'Santander: nao encontrei "Resumo - mes/ano" para saber o ano das datas dd/mm'
    mes_ref = MESES_PT.get(m_periodo.group(1)[:3].lower())
    ano_ref = int(m_periodo.group(2))
    if not mes_ref:
        return [], 'Santander: nao consegui interpretar o mes no cabecalho "Resumo"'

    padrao = re.compile(r'^(\d{2}/\d{2})\s+(.+?)\s+-?([\d\.]+,\d{2})-?$')
    excluir = ('SALDO EM', 'SALDO DISPON', 'SALDO ANTERIOR', 'SALDO BLOQUEADO', 'SALDO FINAL')
    out = []
    data_atual = None
    for l in _linhas_uteis(texto):
        l = l.strip()
        m = padrao.match(l)
        if m:
            data_atual = m.group(1)
            desc = m.group(2)
            valor_str = m.group(3)
        else:
            # linha de continuacao (sem data, pertence ao lancamento anterior) - ignorar valor aqui
            continue
        if any(k in desc.upper() for k in excluir):
            continue
        dd, mm = data_atual.split('/')
        try:
            dt = date(ano_ref, int(mm), int(dd))
        except ValueError:
            continue
        valor = _dec(valor_str)
        desc_up = desc.upper()
        if any(k in desc_up for k in _KW_CREDITO_SANTANDER):
            tipo, obs = 'C', ''
        elif any(k in desc_up for k in _KW_DEBITO_SANTANDER) or l.rstrip().endswith('-'):
            tipo, obs = 'D', ''
        else:
            tipo, obs = 'C', '[A VERIFICAR] tipo C/D inferido por palavra-chave, sem confirmacao no PDF'
        out.append({'data': dt, 'historico': desc.strip(), 'valor': valor, 'tipo': tipo, 'obs': obs})
    return out, ('Santander: o texto extraido do PDF nao tem uma coluna/sinal confiavel de credito x '
                 'debito (o layout usa duas colunas visuais que se perdem na extracao de texto). '
                 'A classificacao foi feita por palavra-chave da descricao e pode falhar em casos '
                 'atipicos - reviser as linhas marcadas [A VERIFICAR].')


# ---------------------------------------------------------------------------
# SICREDI
# ---------------------------------------------------------------------------
def parse_sicredi(texto):
    m_periodo = re.search(r'Per[íi]odo\s*(?:de|\()\s*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
    padrao = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(.+?)\s+(\S+)\s+(-?[\d\.]+,\d{2})\s+[\d\.]+,\d{2}$')
    excluir = ('SALDO ANTERIOR',)
    out = []
    for l in _linhas_uteis(texto):
        l = l.strip()
        if l.upper().startswith('SALDO ANTERIOR'):
            continue
        m = padrao.match(l)
        if not m:
            continue
        data_str, desc, doc, valor_str = m.groups()
        dd, mm, yyyy = data_str.split('/')
        try:
            dt = date(int(yyyy), int(mm), int(dd))
        except ValueError:
            continue
        valor = _dec(valor_str.lstrip('-'))
        tipo = 'D' if valor_str.strip().startswith('-') else 'C'
        out.append({'data': dt, 'historico': desc.strip(), 'valor': valor, 'tipo': tipo, 'obs': ''})
    return out, None


# ---------------------------------------------------------------------------
# BRADESCO  (somente credito confirmado na amostra recebida)
# ---------------------------------------------------------------------------
def parse_bradesco(texto):
    linhas = _linhas_uteis(texto)
    # Padrao geral de uma linha de valor: [texto opcional antes] DOC VALOR SALDO
    # (texto opcional cobre "REM: fulano dd/mm" no credito, ou nada no debito)
    trailing = re.compile(r'^(.*?)\s*(\S+)\s+(-?[\d\.]+,\d{2})\s+([\d\.]+,\d{2})$')
    date_re = re.compile(r'^(\d{2}/\d{2}/\d{4})\s*(.*)$')
    ignorar_prefixo = ('SALDO ANTERIOR', 'TOTAL', 'SALDOS INVEST', 'SALDO INVEST', 'OS DADOS ACIMA')
    # Prefixos que sempre marcam o INICIO de um novo lancamento (linha "tipo").
    # Usados para saber se uma linha depois do valor e continuacao da descricao
    # anterior (ex: "CIADESCP" apos "PAGTO ELETRON COBRANCA") ou o comeco do
    # proximo lancamento.
    tipos_conhecidos = ('PIX RECEBIDO', 'PIX ENVIADO', 'PIX QR CODE', 'PAGTO ELETRON',
                         'PAGTO ELETRONICO', 'TARIFA BANCARIA', 'RENTAB.INVEST', 'TED ',
                         'DOC ', 'DEB ', 'CRED ')
    out = []
    data_atual = None
    desc_buffer = []
    i = 0
    n = len(linhas)
    while i < n:
        l = linhas[i].strip()
        if any(l.upper().startswith(k) for k in ignorar_prefixo):
            desc_buffer = []
            i += 1
            continue
        dm = date_re.match(l)
        if dm:
            data_str, resto = dm.groups()
            dd, mm, yyyy = data_str.split('/')
            try:
                data_atual = date(int(yyyy), int(mm), int(dd))
            except ValueError:
                pass
            l = resto.strip()
            if not l:
                i += 1
                continue
        m = trailing.match(l)
        if not m:
            desc_buffer.append(l)
            i += 1
            continue
        prefixo, doc, valor_str, saldo_str = m.groups()
        prefixo = prefixo.strip()
        partes = list(desc_buffer)
        if prefixo:
            partes.append(prefixo)
        desc_buffer = []
        i += 1
        # Olha a proxima linha: se nao for data, nao for outro valor, e nao
        # comecar com um prefixo de tipo conhecido, e continuacao da descricao
        # deste mesmo lancamento (ex: "CIADESCP" depois de "PAGTO ELETRON
        # COBRANCA ... -13.229,77").
        if i < n:
            prox = linhas[i].strip()
            if (prox and not date_re.match(prox) and not trailing.match(prox)
                    and not any(prox.upper().startswith(t) for t in tipos_conhecidos)
                    and not prox.upper().startswith(('REM:', 'DES:'))):
                partes.append(prox)
                i += 1
        desc = ' - '.join(p for p in partes if p) if partes else '(descricao nao identificada pelo texto extraido)'
        if data_atual is None:
            continue
        valor = _dec(valor_str.lstrip('-'))
        tipo = 'D' if valor_str.strip().startswith('-') else 'C'
        out.append({'data': data_atual, 'historico': desc, 'valor': valor, 'tipo': tipo, 'obs': ''})
    return out, None


BANK_PARSERS = {
    'sicoob': parse_sicoob,
    'caixa': parse_caixa,
    'itau': parse_itau,
    'unicred': parse_unicred,
    'cresol': parse_cresol,
    'ailos': parse_ailos,
    'bb': parse_bb,
    'santander': parse_santander,
    'sicredi': parse_sicredi,
    'bradesco': parse_bradesco,
}
