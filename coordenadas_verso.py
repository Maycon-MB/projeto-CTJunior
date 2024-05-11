import cv2

# Carregue a imagem
imagem = cv2.imread('imgs/verso_carteirinha600.jpg')

# Função de callback para capturar as coordenadas ao clicar na imagem
def capturar_coordenadas(event, x, y, flags, param):
    if event == cv2.EVENT_LBUTTONDOWN:
        # Calcula as coordenadas em relação à imagem
        altura, largura, _ = imagem.shape
        proporcao_x = largura / largura_da_janela
        proporcao_y = altura / altura_da_janela
        coord_x = int(x * proporcao_x)
        coord_y = int(y * proporcao_y)
        print(f'Coordenadas: ({coord_x}, {coord_y})')

# Crie uma janela para a imagem com o modo WINDOW_AUTOSIZE
cv2.namedWindow('Imagem', cv2.WINDOW_AUTOSIZE)

# Associe a função de callback à janela
cv2.setMouseCallback('Imagem', capturar_coordenadas)

# Exiba a imagem
cv2.imshow('Imagem', imagem)

# Obtenha as dimensões da janela
altura_da_janela, largura_da_janela, _ = imagem.shape

# Espere até que a janela seja fechada
cv2.waitKey(0)
cv2.destroyAllWindows()
