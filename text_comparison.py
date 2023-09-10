import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import CountVectorizer

# NOME COMPLETO DA PLANILHA QUE SERÁ MANIPULADA
planilha = 'planilha_comparacao.xlsx'

# NOME DA NOVA PLANILHA PARA SALVAR OS RESULTADOS
nova_planilha = 'planilha_comparacao_resultados.xlsx'

# CRIA UMA LISTA QUE SERÁ POVOADA COM OS PERCENTUAIS
lista_percentual = []

# MODELO DO TIPO ANIGRAMA (CADA PALAVRA VIRA UM NÚMERO)
n = 1

# LÊ O DATASET
df = pd.read_excel(planilha)

# INDICA O TAMANHO DE LINHAS DO DATASET
qtd_linhas = len(df)
print(f'Total de linhas a comparar {qtd_linhas}')
print()

# PERCORRE DUAS COLUNAS DO DATASET
for i, texto_fonte in enumerate(df['TEXTO_FONTE']):
    texto_a_ser_comparado = df.loc[i, 'TEXTO_A_COMPARAR']

    counts = CountVectorizer(analyzer='word', ngram_range=(n, n))

    n_grams = counts.fit_transform([texto_a_ser_comparado, texto_fonte])

    vocab2int = counts.fit([texto_a_ser_comparado, texto_fonte]).vocabulary_

    n_grams_array = n_grams.toarray()

    intersection_list = np.amin(n_grams.toarray(), axis=0)

    intersection_count = np.sum(intersection_list)

    index_A = 0
    A_count = np.sum(n_grams.toarray()[index_A])
    percentual = (intersection_count / A_count)

    # ADICIONA O PERCENTUAL NA LISTA
    lista_percentual.append(percentual)
    print(f'Comparando linha {i + 1} do total de {qtd_linhas}...')

# CRIA UM DATAFRAME COM A LISTA DE PERCENTUAIS
df_percentual = pd.DataFrame(lista_percentual, columns=['PERCENTUAL_SEMELHANCA'])

# CRIA UM DATAFRAME COM OS DADOS ANTIGOS E OS RESULTADOS
df_resultados = pd.concat([df, df_percentual], axis=1)

# SALVA O DATAFRAME COM OS RESULTADOS EM UMA NOVA PLANILHA
df_resultados.to_excel(nova_planilha, sheet_name='BD', index=False)

print()
print('Processo finalizado com sucesso!')
