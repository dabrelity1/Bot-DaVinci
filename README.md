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

Depois de executar o instalador, selecione ambas as caixas e prossiga com a instalação: 

![image](https://github.com/dabrelity1/Bot-DaVinci/assets/147398154/47f250a2-e902-4f6b-ab98-1f4672986754)

## Passo 2: Instalar as Livrarias usadas no código:

Para executar o código sem erros será necessário a instalação de 3 livrarias, o **OpenCV**, o **XLWings**, o **XLSXWriter** e o **NumPy**.

### OpenCV:

Para instalar o OpenCV (*e o NumPy*) diriga-se à linha de comandos do seu computador (Windows):

![image](https://github.com/dabrelity1/Bot-DaVinci/assets/147398154/c94c1d73-8e20-4300-b4b7-e90339208af7)

Após executar a linha de comandos, escreva esta síntaxe:

*py -mpip install opencv-python*.

**MANTENHA-SE NA LINHA DE COMANDOS DO SEU COMPUTADOR!!**

### XLWings:

Para instalar o XLWings escreva:

*py -mpip install xlwings*

### XLSXWriter:

Para instalar o XLSXWriter escreva:

*py -mpip install xlsxwriter*

## Passo 3: Instalar o Visual Studio Code:

Caso já possuir o Visual Studio Code instalado no seu computador **deve ignorar este passo**, caso não tiver, deve **seguir este passo**.

Para a instalação do Visual Studio Code diriga-se a https://code.visualstudio.com/ (O site oficial).
Após chegar ao site, clique no botão "Download for Windows":

![image](https://github.com/dabrelity1/Bot-DaVinci/assets/147398154/70136bf9-7bce-47c8-b9ec-8d38da48029e)

Após executar o instalador, pode prosseguir sempre (Clicar sempre em Seguinte) e siga para o passo seguinte.

## Passo 4: Realizar o download do bot e executá-lo a partir do Visual Studio Code:

Agora, faça download do código fonte do bot neste repositório e abra a pasta *Bot-DaVinci-main*.
Após abrir a pasta, execute o ficheiro *davinci_bot.py* através do Visual Studio Code:

![image](https://github.com/dabrelity1/Bot-DaVinci/assets/147398154/95d656d4-1feb-41cc-bd4f-1f0db0c5f1bd)

Após abrir o Visual Studio Code através do ficheiro, permita que o Visual Studio Code abra ficheiros externos:

![image](https://github.com/dabrelity1/Bot-DaVinci/assets/147398154/54218051-5f23-4ee8-851a-14c2eae38620)

## Passo 5: Fazer o bot detectar a imagem:

Deve ter reparado que na pasta *Bot-DaVinci-main* existe um ficheiro de imagem chamado *menina.png*. Se preferir, pode mudar por outro, mas neste exemplo vamos usar esse.

Em todos os computadores, os diretórios são diferentes, ou seja, encontram-se em localizações distintas. Para resolver isto basta:

1. Selecionar a imagem no Explorador de Ficheiros do Windows;

![image](https://github.com/dabrelity1/Bot-DaVinci/assets/147398154/de8c2c35-428a-475c-a4b8-0289da147cc5)

2. Clicar em "Base", no lado direito do botão azul "Ficheiro" no canto superior esquerdo.

![image](https://github.com/dabrelity1/Bot-DaVinci/assets/147398154/5780c508-b5ed-4ffa-b6f3-c5389b98e2ee)

3. Clicar em "Copiar caminho":

![image](https://github.com/dabrelity1/Bot-DaVinci/assets/147398154/2f6fd934-171b-4ae3-b9c4-bb70d7a659ae)

No meu caso, o diretório é "C:\Users\14475\Downloads\Bot-DaVinci-main\menina.png", mas é perfeitamente normal ser diferente em outros computadores.

Agora, volte ao Visual Studio Code e, na linha 53 (V.0.1) elimine o 'C:\Users\User\Downloads\picasso_bot-main\bot pintor\menina.png' e substitua com o seu diretório (No meu caso seria 'C:\Users\14475\Downloads\Bot-DaVinci-main\menina.png'). Tenha em atenção remover as "", e substituir com '' no início e fim do diretório:

![image](https://github.com/dabrelity1/Bot-DaVinci/assets/147398154/37de383e-e7e2-4e2d-987f-c6d4ee74e5e7)

## Passo 6: Executar o código:

Para executar o código, basta carregar com o **botão direito do rato** num espaço vazio (sem código) e clicar em "Run Python":

![image](https://github.com/dabrelity1/Bot-DaVinci/assets/147398154/85b51841-f739-49fc-8fae-b5c3db63d200)

Depois, para finalmente executar o código, basta clicar em "Run Python File in Terminal":

![image](https://github.com/dabrelity1/Bot-DaVinci/assets/147398154/eec11255-0af8-42a5-9a38-91d50879d50f)

E agora o Excel deverá abrir e dentro de alguns segundos começar a renderizar a imagem.

(**Para parar o código basta terminar o Excel**)

# Notes (for English users)

If you pretend to analyze the code itself, be aware that all the comments on the code explaining it's functions are in Portuguese.
