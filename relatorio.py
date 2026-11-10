# -*- coding: utf-8 -*-
"""
Relatorio em .txt gerado quando o .ofx NAO sai (ou sai com ressalvas) -
usado pelo Conversor_Extratos e (a partir da v3.0, com OCR embutido) pelo mesmo exe. Escrito pra quem nao
e da area tecnica: diz o que houve, o que fazer e se existe versao nova do
programa (muitos erros ja estao corrigidos numa versao mais nova).
"""
import os

CONTATO = 'Jean Vieira - (49) 99907-9884 - jean.vieira@hotmail.com'


def texto_versao(info_update, versao_atual, url_download):
    """Bloco sobre versao nova. info_update tem 'nova_versao' (se achou) e
    'verificado' (se conseguiu consultar a internet)."""
    nova = info_update.get('nova_versao')
    if nova:
        return (
            f'*** EXISTE UMA VERSÃO NOVA DO PROGRAMA: v{nova} (você está usando a v{versao_atual}) ***\n'
            'Esse problema pode já ter sido corrigido na versão nova.\n'
            'ATUALIZE ANTES DE TENTAR DE NOVO:\n'
            '  1) Abra o programa outra vez e, quando ele perguntar "Atualizar agora? (S/N)", digite S e Enter.\n'
            f'  2) Ou baixe a versão nova em: {url_download}\n'
        )
    if info_update.get('verificado'):
        return (f'Versão do programa: v{versao_atual}. Você já está usando a versão mais recente, '
                'então atualizar não vai resolver este caso.\n')
    return (f'Versão do programa: v{versao_atual}. Não consegui verificar se existe versão nova '
            '(talvez o computador esteja sem internet). Se tiver internet, abra o programa de novo '
            'e veja se ele oferece a atualização, pois o problema pode já estar corrigido.\n')


def escrever_relatorio(caminho_pdf, pasta_saida, sufixo, titulo, o_que_houve, o_que_fazer,
                       info_update, versao_atual, url_download):
    """Grava '<nome do pdf><sufixo>.txt' na pasta de saida e devolve o caminho."""
    nome_base = os.path.splitext(os.path.basename(caminho_pdf))[0]
    caminho = os.path.join(pasta_saida, f'{nome_base}{sufixo}.txt')
    linhas = [
        titulo,
        '=' * len(titulo),
        '',
        f'Arquivo: {os.path.basename(caminho_pdf)}',
        '',
        'O QUE ACONTECEU',
        '---------------',
        o_que_houve.strip(),
        '',
        'O QUE FAZER',
        '-----------',
        o_que_fazer.strip(),
        '',
        'VERSÃO DO PROGRAMA',
        '------------------',
        texto_versao(info_update, versao_atual, url_download).strip(),
        '',
        f'Se continuar com problema, envie o PDF para: {CONTATO}',
        '',
    ]
    with open(caminho, 'w', encoding='utf-8-sig', newline='\r\n') as f:
        f.write('\n'.join(linhas))
    return caminho
