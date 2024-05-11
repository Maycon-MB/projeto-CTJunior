from PIL import Image, ImageDraw, ImageFont

# Função para preencher a frente da carteirinha com informações da planilha
def preencher_frente(nome, cpf, contato):
    # Abrir a imagem da frente da carteirinha
    frente_carteirinha = Image.open("frente_carteirinha.jpg")
    
    # Carregar a fonte para o texto
    fonte = ImageFont.truetype("arial.ttf", 16)
    
    # Preencher a imagem com as informações
    draw = ImageDraw.Draw(frente_carteirinha)
    draw.text((211, 230), f"Nome: {nome}", fill=(0, 0, 0), font=fonte)
    draw.text((214, 280), f"CPF: {cpf}", fill=(0, 0, 0), font=fonte)
    draw.text((215, 327), f"Contato: {contato}", fill=(0, 0, 0), font=fonte)
    
    # Salvar a frente da carteirinha preenchida
    frente_carteirinha.save("frente_carteirinha_preenchida.jpg")

# Função para preencher o verso da carteirinha com informações da planilha
def preencher_verso(pagamentos):
    # Abrir a imagem do verso da carteirinha
    verso_carteirinha = Image.open("verso_carteirinha.jpg")
    
    # Carregar a fonte para o texto
    fonte = ImageFont.truetype("arial.ttf", 16)
    
    # Preencher a imagem com as informações dos pagamentos
    draw = ImageDraw.Draw(verso_carteirinha)
    y = 183
    for mes, status in pagamentos.items():
        draw.text((300, y), f"{mes}: {status}", fill=(0, 0, 0), font=fonte)
        y += 30
    
    # Salvar o verso da carteirinha preenchido
    verso_carteirinha.save("verso_carteirinha_preenchido.jpg")

# Dados de exemplo da planilha
nome = "João da Silva"
cpf = "123.456.789-00"
contato = "(00) 12345-6789"
pagamentos = {
    "Janeiro": "Pago",
    "Fevereiro": "Não pago",
    "Março": "Pago"
}

# Chamando as funções para preencher as carteirinhas com os dados da planilha
preencher_frente(nome, cpf, contato)
preencher_verso(pagamentos)
