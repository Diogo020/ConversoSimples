# ClaraConvert

Aplicativo desktop em Python para conversoes de PDF, Word (DOCX) e imagens.
Tem interface em portugues, tema escuro com destaque verde-claro e nao precisa
de terminal para ser usado depois de iniciado.

E possivel selecionar varios arquivos de uma vez ou arrasta-los para a area
de selecao. O ClaraConvert oferece formatos de saida que funcionam para todos
os arquivos selecionados; resultados com nomes repetidos recebem numeracao
para nao sobrescrever arquivos existentes.

## Conversoes disponiveis

- PDF para DOCX, PNG ou JPG.
- DOCX para PDF, PNG ou JPG.
- Imagens PNG, JPG/JPEG, WEBP, BMP, TIFF e GIF para PDF.
- Conversao de imagens entre PNG, JPG e WEBP.
- PDFs com varias paginas e documentos Word convertidos em imagens geram um
  arquivo por pagina, com numeracao no nome.
- PDFs digitalizados como imagem precisam de OCR para virar texto editavel;
  esta versao nao inclui OCR.

## Executar durante o desenvolvimento

1. Instale Python 3.10 ou superior no Windows.
2. Abra um terminal nesta pasta.
3. Execute:

   ```powershell
   py -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   py conversor.py
   ```

## Criar o executavel para Windows

Execute `build.bat` com Python instalado. O arquivo sera criado em
`dist\ClaraConvert.exe`. Copie esse executavel para outro computador Windows
para usar o aplicativo sem instalar Python.

O executavel e compilado para Windows e nao e um binario multiplataforma.
Para converter DOCX em PDF ou imagens, o computador de destino tambem precisa
ter o LibreOffice instalado. As outras conversoes usam as bibliotecas
empacotadas no executavel.

## Privacidade

As conversoes sao feitas localmente no computador. Os arquivos nao sao enviados
para um servico online.
