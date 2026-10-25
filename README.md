# ClaraConvert

Aplicativo desktop em Python para conversões de PDF, Word (DOCX) e imagens.
Tem interface em português com tema verde-escuro, seletor de destino do
Windows e uma lista dos arquivos convertidos durante a sessão, que podem ser
abertos com duplo clique ou pelo botão "Abrir arquivo".

É possível selecionar vários arquivos de uma vez ou arrastá-los para a área
de seleção. A lista de entrada permite remover um arquivo selecionado ou
limpar toda a seleção. O ClaraConvert oferece formatos de saída que funcionam
para todos os arquivos selecionados; resultados com nomes repetidos recebem
numeração para não sobrescrever arquivos existentes.

## Conversões disponíveis

- PDF para DOCX, PNG ou JPG.
- DOCX para PDF, PNG ou JPG.
- Imagens PNG, JPG/JPEG, WEBP, BMP, TIFF e GIF para PDF.
- Conversão de imagens entre PNG, JPG e WEBP.
- PDFs com várias páginas e documentos Word convertidos em imagens geram um arquivo por página, com numeração no nome.
- PDFs digitalizados como imagem precisam de OCR para virar texto editável; esta versão não inclui OCR.

## Executar durante o desenvolvimento

1. Instale Python 3.10 ou superior no Windows.
2. Abra um terminal nesta pasta.
3. Execute:

```
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
py conversor.py
```

## Criar o executável para Windows

Execute `build.bat` com Python instalado. O arquivo será criado em
`dist\ClaraConvert.exe`. Copie esse executável para outro computador Windows
para usar o aplicativo sem instalar Python.

Baixe a versão mais recente na
[página de versões do ClaraConvert](https://github.com/Diogo020/ConversoSimples/releases/latest).

O executável é compilado para Windows e não é um binário multiplataforma.
Para converter DOCX em PDF ou imagens, o computador de destino também precisa
ter o LibreOffice instalado. As outras conversões usam as bibliotecas
empacotadas no executável.

## Privacidade

As conversões são feitas localmente no computador. Os arquivos não são
enviados para um serviço online.
