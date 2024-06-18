# Bot Da Vinci
Bot escrito em Python que tem como objetivo a recriação de imagens numa folha de Excel (Suportado a partir do Excel 2016).

Bot written in Python that aims to recreate images in an Excel sheet (Supported from Excel 2016).

# V.0.1 (15/06/2024)

Primeiro commit do projeto no GitHub. O bot de momento consegue pegar em qualquer imagem em formato JPEG e PNG (em PNG a imagem fica preta na área "vazia").
Ainda é instável, pelo que é necessário colocar a imagem num diretório do próprio computador e não na pasta do projeto em si. ("Exemplo: C:\Users\User\Downloads\bot_davinci\bot pintor\menina.png" em vez de "menina.png").

First commit of the project on GitHub. The bot is currently able to take any image in JPEG and PNG format (in PNG the image turns black in the "empty" area).
It's still unstable, so you need to put the image in a directory on your computer and not in the project folder itself ("Example: C:\Users\User\Downloads\bot_davinci\bot pintor\menina.png" instead of "menina.png").

# Instruções de Instalação/Install Instructions

## Passo 1: Instalar Python:

Para fazer a instalação de Python será necessário ir ao website oficial (python.org) e instalar a versão mais recente em "Downloads". Para realizar o download do instalador de Python (3.12.4) mais rapidamente, clique neste link: https://www.python.org/ftp/python/3.12.4/python-3.12.4-amd64.exe (O link irá instalar o Python no diretório escolhido).

Depois de executar o instalador, por favor selecione ambas as caixas e prossiga com a instalação: 

![image](https://github.com/dabrelity1/Bot-DaVinci/assets/147398154/47f250a2-e902-4f6b-ab98-1f4672986754)

## Passo 2: Instalar as Livrarias usadas no código:

Para executar o código sem erros será necessário a instalação de 3 livrarias, o **OpenCV**, o **XLWings**, o **XLSXWriter** e o **NumPy**.

### OpenCV:

Para instalar o OpenCV (*e o NumPy*) diriga-se à linha de comandos do seu computador (Windows):

![image](https://github.com/dabrelity1/Bot-DaVinci/assets/147398154/c94c1d73-8e20-4300-b4b7-e90339208af7)

Após executar a linha de comandos, escreva esta síntaxe:

*py -mpip install opencv-python*.

**MANTENHA-SE NA LINHA DE COMANDOS DO SEU COMPUTADOR**

### XLWings:

Para instalar o XLWings escreva:

*py -mpip install xlwings*

**MANTENHA-SE NA LINHA DE COMANDOS DO SEU COMPUTADOR**

### XLSXWriter:

Para instalar o XLSXWriter escreva:

*py -mpip install xlsxwriter*

## Passo 3: Instalar o Visual Studio Code:

Caso já possuir o Visual Studio Code instalado no seu computador **deve ignorar este passo**, caso não tiver, deve **seguir este passo**.

Para a instalação do Visual Studio Code diriga-se a https://code.visualstudio.com/ (O site oficial).
Após chegar ao site, clique no botão "Download for Windows":

![image](https://github.com/dabrelity1/Bot-DaVinci/assets/147398154/70136bf9-7bce-47c8-b9ec-8d38da48029e)


# Notes (for English users)

If you pretend to analyze the code itself, be aware that all the comments on the code explaining it's functions are in Portuguese.
