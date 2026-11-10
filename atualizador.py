# -*- coding: utf-8 -*-
"""
Atualizacao automatica (com confirmacao) dos executaveis - usado pelo
Conversor_Extratos e pelo Conversor_OCR.

Fluxo: ao abrir, o programa le o arquivo de versao no GitHub; se houver
versao nova e o usuario aceitar, baixa o .zip da release estavel, confere
o SHA-256 do .exe contra o hash publicado no repositorio e troca o proprio
.exe NA MESMA PASTA (o exe em uso e renomeado pra .old - o Windows permite
renomear um exe em execucao - e o novo entra com o nome de sempre). Depois
reabre o programa ja atualizado, com os mesmos argumentos.

Qualquer falha (sem internet, pasta sem permissao de escrita, hash que nao
bate, etc.) deixa o exe atual intacto e o programa continua funcionando
normalmente na versao antiga.
"""
import sys
import os
import io
import hashlib
import subprocess
import tempfile
import urllib.request
import zipfile

_ENV_JA_ATUALIZADO = 'CONVERSOR_JA_ATUALIZADO'
# Variavel so pra testes locais: so e aceita se apontar pra file:// (nunca
# redireciona a atualizacao pra um endereco da internet).
_ENV_BASE_TESTE = 'CONVERSOR_URL_BASE_TESTE'


def _url(url):
    base = os.environ.get(_ENV_BASE_TESTE, '')
    if base.startswith('file:'):
        return base.rstrip('/') + '/' + url.rsplit('/', 1)[-1]
    return url


def _ler_texto(url, timeout):
    with urllib.request.urlopen(_url(url), timeout=timeout) as resp:
        return resp.read().decode('utf-8').strip()


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


def verificar_versao(url_version, versao_atual, resultado):
    """Roda em thread (ver chamadores) pra nao atrasar o programa. So le um
    numero de volta, nao manda nenhum dado do usuario. Qualquer falha e
    ignorada em silencio."""
    try:
        remota = _ler_texto(url_version, 2.5)
        resultado['verificado'] = True
        if _versao_maior(remota, versao_atual):
            resultado['nova_versao'] = remota
    except Exception:
        pass


def limpar_restos(nome_exe):
    """Apaga o .old/.novo deixado por uma atualizacao anterior. Se o .old
    ainda estiver em uso, tenta de novo na proxima abertura."""
    if not getattr(sys, 'frozen', False):
        return
    pasta = os.path.dirname(sys.executable)
    for sufixo in ('.old', '.novo'):
        try:
            os.remove(os.path.join(pasta, nome_exe + sufixo))
        except OSError:
            pass


def ja_foi_atualizado():
    return bool(os.environ.get(_ENV_JA_ATUALIZADO))


def _pasta_gravavel(pasta):
    try:
        with tempfile.NamedTemporaryFile(dir=pasta, prefix='.teste_', delete=True):
            pass
        return True
    except OSError:
        return False


def _baixar(url, timeout=30):
    with urllib.request.urlopen(_url(url), timeout=timeout) as resp:
        total = int(resp.headers.get('Content-Length') or 0)
        buf = io.BytesIO()
        lido = 0
        ultimo = -1
        while True:
            bloco = resp.read(1024 * 256)
            if not bloco:
                break
            buf.write(bloco)
            lido += len(bloco)
            if total:
                pct = lido * 100 // total
                if pct != ultimo and pct % 10 == 0:
                    print(f'  baixando... {pct}%', flush=True)
                    ultimo = pct
        return buf.getvalue()


def atualizar_exe(nome_exe, url_zip, url_hash, versao_nova):
    """Baixa, confere e troca o exe. Retorna (ok, mensagem). Em caso de
    sucesso o chamador deve chamar reabrir_e_sair()."""
    if not getattr(sys, 'frozen', False):
        return False, 'so atualiza o .exe (rodando via Python direto, nada a fazer)'
    exe_atual = sys.executable
    pasta = os.path.dirname(exe_atual)
    if not _pasta_gravavel(pasta):
        return False, f'sem permissao pra gravar na pasta {pasta}'

    novo = os.path.join(pasta, nome_exe + '.novo')
    old = os.path.join(pasta, nome_exe + '.old')
    try:
        hash_esperado = _ler_texto(url_hash, 10).split()[0].lower()
        if len(hash_esperado) != 64:
            return False, 'arquivo de verificacao (hash) invalido'
        print(f'Baixando a versao {versao_nova}...', flush=True)
        dados_zip = _baixar(url_zip)
        with zipfile.ZipFile(io.BytesIO(dados_zip)) as z:
            conteudo = z.read(nome_exe)
        if hashlib.sha256(conteudo).hexdigest() != hash_esperado:
            return False, 'o arquivo baixado nao confere com o hash publicado (nao foi instalado)'
        with open(novo, 'wb') as f:
            f.write(conteudo)
    except Exception as e:
        try:
            os.remove(novo)
        except OSError:
            pass
        return False, f'falha no download ({e})'

    try:
        try:
            os.remove(old)
        except OSError:
            pass
        os.replace(exe_atual, old)
        try:
            os.replace(novo, exe_atual)
        except Exception:
            os.replace(old, exe_atual)
            raise
    except Exception as e:
        try:
            os.remove(novo)
        except OSError:
            pass
        return False, f'nao consegui trocar o arquivo ({e})'
    return True, 'ok'


def reabrir_e_sair():
    """Reabre o exe (agora o novo) com os mesmos argumentos, espera ele
    terminar e sai com o mesmo codigo. Espera (em vez de largar o processo)
    pra o console continuar sendo do programa novo ate o fim."""
    env = dict(os.environ)
    env[_ENV_JA_ATUALIZADO] = '1'
    # Um exe --onefile aberto por dentro de outro herdaria a pasta temporaria
    # do pai; isso forca o filho a extrair a propria.
    env['PYINSTALLER_RESET_ENVIRONMENT'] = '1'
    env.pop('_MEIPASS2', None)
    try:
        rc = subprocess.call([sys.executable] + sys.argv[1:], env=env,
                             cwd=os.path.dirname(sys.executable))
    except Exception as e:
        print(f'Atualizado, mas nao consegui reabrir sozinho ({e}). Abra o programa de novo.')
        rc = 0
    sys.exit(rc)


def oferecer_atualizacao(nome_exe, url_zip, url_hash, versao_atual, versao_nova, url_manual):
    """Pergunta ao usuario e, se aceitar, atualiza e reabre. Se nao der pra
    perguntar (sem console interativo), nao faz nada alem do aviso passivo
    que o chamador ja mostra no fim. Retorna sem fazer nada se recusar ou
    falhar; so sai do processo quando a atualizacao deu certo."""
    if not getattr(sys, 'frozen', False) or ja_foi_atualizado():
        return
    try:
        interativo = sys.stdin is not None and sys.stdin.isatty()
    except Exception:
        interativo = False
    if os.environ.get(_ENV_BASE_TESTE, '').startswith('file:'):
        interativo = True  # modo de teste local: aceita resposta por stdin redirecionado
    if not interativo:
        return
    print()
    print('============================================')
    print(f'Ha uma versao nova: v{versao_nova} (voce esta usando a v{versao_atual}).')
    print('Atualizar agora? O programa baixa, troca o arquivo nesta mesma pasta e')
    print('reabre sozinho (seus PDFs nao sao mexidos).')
    print('============================================')
    try:
        resp = input('Atualizar agora? (S/N): ').strip().lower()
    except (EOFError, OSError):
        return
    if resp not in ('s', 'sim', 'y', 'yes'):
        return
    ok, msg = atualizar_exe(nome_exe, url_zip, url_hash, versao_nova)
    if ok:
        print('Atualizado! Reabrindo...')
        reabrir_e_sair()
    print(f'Nao foi possivel atualizar: {msg}.')
    print(f'Continuando com a versao atual. Baixe manualmente em: {url_manual}')
    print()
