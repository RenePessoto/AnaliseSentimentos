import pandas as pd
from nltk.sentiment import SentimentIntensityAnalyzer
import nltk

nltk.download('vader_lexicon')

# Criação do objeto SentimentIntensityAnalyzer
sia = SentimentIntensityAnalyzer()

# Nome do arquivo Excel para atualização
nome_arquivo = "analise_sentimentos.xlsx"

# Carregar o arquivo Excel existente, se existir
try:
    df = pd.read_excel(nome_arquivo)
except FileNotFoundError:
    # Se o arquivo não existir, criar um DataFrame vazio
    df = pd.DataFrame()

# Criação das listas para armazenar os dados
instituicoes = []
tipos = []
comentarios = []
classificacoes = []

while True:
    # Recebendo os dados da instituição
    instituicao = input("Digite o nome da instituição (ou 'sair' para encerrar): ")
    if instituicao.lower() == "sair":
        break

    tipo = input("Digite se a instituição é privada ou pública: ")

    # Recebendo o comentário
    comentario = input("Digite o comentário do Reclame Aqui (pressione Enter duas vezes para encerrar): ")
    comentarios.append(comentario)

    # Realizando a análise de sentimento
    sentiment = sia.polarity_scores(comentario)
    if sentiment["compound"] >= 0.05:
        classificacao = "Positivo"
    elif sentiment["compound"] <= -0.05:
        classificacao = "Negativo"
    else:
        classificacao = "Neutro"

    # Armazenando os dados nas listas
    instituicoes.append(instituicao)
    tipos.append(tipo)
    classificacoes.append(classificacao)

# Adicionando os novos dados ao DataFrame existente
novos_dados = pd.DataFrame({
    "Nome da Instituição": instituicoes,
    "Tipo": tipos,
    "Comentário": comentarios,
    "Classificação": classificacoes
})

df = pd.concat([df, novos_dados], ignore_index=True)

# Salvando o DataFrame de volta no mesmo arquivo Excel
df.to_excel(nome_arquivo, index=False)

