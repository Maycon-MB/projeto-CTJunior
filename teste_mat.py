def gerar_matricula(cpf, data_nascimento, ano_inscricao):
    # Remove caracteres não numéricos do CPF e da data de nascimento
    cpf_numerico = ''.join(filter(str.isdigit, cpf))
    data_nascimento_numerica = ''.join(filter(str.isdigit, data_nascimento))
    
    # Pega os dois últimos dígitos do ano de inscrição
    ano_inscricao_ultimos_digitos = str(ano_inscricao)[-2:]
    
    # Concatena as informações e retorna como matrícula
    matricula = ano_inscricao_ultimos_digitos + cpf_numerico + data_nascimento_numerica
    
    # Se a matrícula tiver mais de 12 dígitos, remove os extras
    if len(matricula) > 12:
        matricula = matricula[:12]
    # Se a matrícula tiver menos de 12 dígitos, preenche com zeros à esquerda
    elif len(matricula) < 12:
        matricula = matricula.zfill(12)
    
    return int(matricula)

# Exemplo de uso
cpf_aluno = "123.456.789-00"  # Substitua pelo CPF do aluno
data_nascimento = "01/01/2000"  # Substitua pela data de nascimento do aluno (formato: dd/mm/aaaa)
ano_inscricao = 2024  # Substitua pelo ano da inscrição do aluno

matricula = gerar_matricula(cpf_aluno, data_nascimento, ano_inscricao)
print("Matrícula do aluno:", matricula)
