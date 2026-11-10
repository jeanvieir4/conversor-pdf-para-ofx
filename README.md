# Conversor PDF para OFX

Converte extrato bancario em PDF (Sicoob, Itau, Bradesco, Caixa, Banco do
Brasil, Sicredi, Cresol, UniCred, Ailos/ViaCredi, Santander e mais) para
OFX pronto para conciliacao. 100% offline, sem IA, sem custo, sem
copiar e colar.

Desenvolvido por [Jean Vieira](mailto:jean.vieira@hotmail.com), contador,
pra eliminar a etapa manual de colar extrato em IA e digitar o resultado
numa planilha.

## Baixar (Windows, pronto pra usar)

Nao precisa instalar Python nem nada. Baixe o zip (contem o `.exe`, o
`LEIA-ME.txt` com instrucoes e a `LICENSE.txt`) e extraia:

**[Conversor_Extratos.zip (GitHub Releases)](https://github.com/jeanvieir4/conversor-pdf-para-ofx/releases/download/1.0/Conversor_Extratos.zip)**

O LEIA-ME.txt tambem esta neste repositorio em
[`Para_Equipe/LEIA-ME.txt`](Para_Equipe/LEIA-ME.txt).

Na primeira execucao o Windows SmartScreen avisa "Editor desconhecido" -
isso e normal pra um executavel sem certificado de assinatura paga, nao e
virus. Clique em "Mais informacoes" -> "Executar assim mesmo".

## O que gera

Pra cada PDF, um arquivo por banco encontrado nele:

- **`.ofx`** (formato OFX 1.02 SGML) pronto pra importar na conciliacao
  bancaria do seu sistema contabil, com o codigo do banco (COMPE)
  preenchido automaticamente.

## Bancos suportados

Sicoob, Caixa Economica Federal, Itau, UniCred, Cresol, Ailos/ViaCredi,
Banco do Brasil, Bradesco, Sicredi e Santander.

PDF de um banco ainda nao suportado nao trava a execucao: ele aparece na
um arquivo `_ERRO.txt` explicando o que houve (e se ha versao nova), e o resto dos extratos e processado normal.
Santander tem uma ressalva: o layout do PDF nao tem marcacao confiavel de
credito/debito no texto extraido, entao alguns lancamentos vem marcados
`[A VERIFICAR]` pra conferencia manual.

## Rodar a partir do codigo-fonte

Requer Python 3.10+.

```bash
pip install pdfplumber openpyxl
python converter.py extrato1.pdf extrato2.pdf
```

Gera os `.ofx` (ou o `_ERRO.txt`) na pasta atual (ou na pasta definida na variavel de
ambiente `PASTA_SAIDA`).

Pra gerar o `.exe` standalone (PyInstaller):

```bash
pip install pyinstaller
python -m PyInstaller --onefile --name Conversor_Extratos --console converter.py
```

## Estrutura

- `converter.py` - ponto de entrada: identifica o banco de cada pagina do
  PDF, agrupa por banco, chama o parser correspondente e gera o `.ofx`.
- `bancos.py` - um `parse_<banco>(texto)` por banco.
- `ofx_export.py` - gera o `.ofx` a partir das transacoes lidas, no
  mesmo (formato validado contra um `.ofx` real de referencia).
- `CONTEXTO_PROJETO.md` - notas de desenvolvimento: status de validacao de
  cada parser, decisoes de formato do OFX, proximos passos.

## Aviso de responsabilidade

Ferramenta oferecida como esta, gratuita, sem garantia. Agiliza a
digitacao, nao substitui a conferencia profissional. Antes de usar
qualquer planilha ou `.ofx` gerado no fechamento contabil de um cliente,
confira a soma dos creditos e a soma dos debitos contra os totais que o
proprio extrato do banco declara.

## Licenca

[MIT](LICENSE) - livre pra usar, copiar e repassar, inclusive em trabalho
comercial.

## Apoie o projeto

Se este programa te ajudou, voce pode apoiar o desenvolvimento via Pix:

**Chave Pix:** jean.vieira@hotmail.com

![QR Code Pix](./assets/qrcode-pix.png)

## Contato

Duvidas, bug ou PDF de banco ainda nao suportado: Jean Vieira -
(49) 99907-9884 - jean.vieira@hotmail.com
