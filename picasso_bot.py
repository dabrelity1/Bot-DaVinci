# Livrarias usadas

import cv2 # Aqui também tinha previamente carregado o OpenCV, mas por razões "desconhecidas" tive que optar por usar o cv2 que é quase equivalente
import xlsxwriter # Livraria usada para dar acesso total ao excel
import numpy as np
import xlwings as xw # Livraria para implementação de Python em Excel (Lógica)
import time

def rgb2hex(r,g,b): # esta função converte o código da imagem RGB em Hexadecimal (Senão não era possível meter cores no Excel)
    r = max(0,min(r,255)) # Supostamente a livraria xlwings usa o formato BGR (Blue, Green, Red) mas para mim é mais fácil trabalhar com RGB por isso foi o que usei (pode resultar em erros de cor)
    g = max(0,min(g,255))
    b = max(0,min(b,255))
    return f'#{r:02x}{g:02x}{b:02x}'

def imgparaexcel_xlwings(img, direction=None, x=1, y=1):
    if x != 1 and y != 1:
        img_resize = cv2.resize(img, None, fx=x, fy=y)
    else:
        img_resize = img
    if len(img_resize.shape) == 3:
        rows, cols, _ = img_resize.shape
    else:
        rows, cols = img_resize.shape
    book = xw.Book()
    sheet = book.sheets[0]
    sheet.range((1, 1), (1, cols)).column_width = 2
    time.sleep(5)
    
    color_matrix = []
    
    # -- Retirar os if's

    if direction == 'horizontal':
        for r in range(rows):
            row_colors = []
            for c in range(cols):
                color = img_resize[r, c] if isinstance(img_resize[r, c], np.ndarray) else [img_resize[r, c]]*3
                row_colors.append((int(color[2]), int(color[1]), int(color[0])))
            color_matrix.append(row_colors)
    else:  # Default para vertical
        for r in range(rows):
            row_colors = []
            for c in range(cols):
                color = img_resize[r, c] if isinstance(img_resize[r, c], np.ndarray) else [img_resize[r, c]]*3
                row_colors.append((int(color[2]), int(color[1]), int(color[0])))
            color_matrix.append(row_colors)
    
    for r in range(rows):
        for c in range(cols):
            sheet.cells(r + 1, c + 1).color = color_matrix[r][c]

# Carregar a imagem
img2 = cv2.imread(r'C:\Users\User\Downloads\picasso_bot-main\bot pintor\menina.png', cv2.IMREAD_COLOR)
# Aqui converteu-se a imagem em vertical
imgparaexcel_xlwings(img2, direction='vertical', x=1, y=1)