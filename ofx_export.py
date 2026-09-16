# -*- coding: utf-8 -*-
"""
Geracao de arquivo OFX (Open Financial Exchange) a partir das transacoes ja
extraidas de um extrato. Formato (cabecalho, tags, FITID por dia, ausencia
de LEDGERBAL) validado campo a campo contra um .ofx real gerado no OFX
Generator para o mesmo extrato Sicoob - nao foi "adivinhado" a partir da
especificacao OFX generica.
"""
from datetime import date

# COMPE (Banco Central) de cada banco, sempre com 4 digitos.
# 0756 confirmado contra o .ofx de referencia do Sicoob; os demais sao o
# codigo COMPE padrao Febraban de cada banco/cooperativa central.
COMPE_BANCO = {
    'sicoob': '0756',
    'caixa': '0104',
    'itau': '0341',
    'unicred': '0136',
    'cresol': '0133',
    'ailos': '0085',
    'bb': '0001',
    'santander': '0033',
    'sicredi': '0748',
    'bradesco': '0237',
    'civia': '0085',  # Civia e uma cooperativa do Sistema Ailos - mesmo COMPE do Ailos
}

ACCTID_PADRAO = '00000'  # numero de conta nao e extraido do PDF; mesmo valor usado no arquivo de referencia


def _montar_linhas(transacoes, banco):
    validas = sorted(
        (t for t in transacoes if t.get('data') and t.get('valor') is not None),
        key=lambda t: t['data'],
    )
    if not validas:
        return None

    data_ini = validas[0]['data']
    data_fim = validas[-1]['data']
    bank_id = COMPE_BANCO.get(banco, '0000')

    linhas = [
        'OFXHEADER:100',
        'DATA:OFXSGML',
        'VERSION:102',
        'SECURITY:NONE',
        'ENCODING:USASCII',
        'CHARSET:1252',
        'COMPRESSION:NONE',
        'OLDFILEUID:NONE',
        'NEWFILEUID:NONE',
        '',
        '<OFX>',
        '<SIGNONMSGSRSV1>',
        '<SONRS>',
        '<STATUS>',
        '<CODE>0',
        '<SEVERITY>INFO',
        '</STATUS>',
        '<DTSERVER>' + data_fim.strftime('%Y%m%d') + '235959',
        '<LANGUAGE>POR',
        '</SONRS>',
        '</SIGNONMSGSRSV1>',
        '<BANKMSGSRSV1>',
        '<STMTTRNRS>',
        '<TRNUID>1001',
        '<STATUS>',
        '<CODE>0',
        '<SEVERITY>INFO',
        '</STATUS>',
        '<STMTRS>',
        '<CURDEF>USD',
        '<BANKACCTFROM>',
        '<BANKID>' + bank_id,
        '<ACCTID>' + ACCTID_PADRAO,
        '<ACCTTYPE>CHECKING',
        '</BANKACCTFROM>',
        '<BANKTRANLIST>',
        '<DTSTART>' + data_ini.strftime('%Y%m%d'),
        '<DTEND>' + data_fim.strftime('%Y%m%d'),
    ]

    contador_do_dia = {}
    for t in validas:
        dia = t['data'].strftime('%y%m%d')
        contador_do_dia[dia] = contador_do_dia.get(dia, 0) + 1
        fitid = f'{dia}{contador_do_dia[dia]:02d}'
        tipo_ofx = 'CREDIT' if t['tipo'] == 'C' else 'DEBIT'
        valor = t['valor'] if t['tipo'] == 'C' else -t['valor']
        memo = t['historico']
        if t.get('obs'):
            memo = f"{memo} {t['obs']}"
        linhas += [
            '<STMTTRN>',
            '<TRNTYPE>' + tipo_ofx,
            '<DTPOSTED>' + t['data'].strftime('%Y%m%d'),
            f'<TRNAMT>{valor:.2f}',
            '<FITID>' + fitid,
            '<CHECKNUM>' + fitid,
            '<MEMO>' + memo,
            '</STMTTRN>',
        ]

    linhas += [
        '</BANKTRANLIST>',
        '</STMTRS>',
        '</STMTTRNRS>',
        '</BANKMSGSRSV1>',
        '</OFX>',
        '',
    ]
    return linhas


def gerar_ofx(transacoes, banco, caminho_saida):
    """Gera um .ofx com as transacoes de UM banco.
    Retorna True se o arquivo foi gerado, False se nao havia nenhuma
    transacao com data e valor validos (nada a exportar)."""
    linhas = _montar_linhas(transacoes, banco)
    if linhas is None:
        return False
    conteudo = '\r\n'.join(linhas)
    with open(caminho_saida, 'w', encoding='cp1252', errors='replace', newline='') as f:
        f.write(conteudo)
    return True
