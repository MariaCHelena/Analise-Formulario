import pandas as pd
import openpyxl
import math

'''A todos os gráficos foram feitos manualmente através do excel, esse código servindo exclusivamente para contabilizar a quantidade de respostas no formulário e os pontos que essas respostas geraram em relação ao seu peso'''

url = 'https://drive.google.com/file/d/1tOh_dQSYaOgSLcBbMGK17GlSCiCfgj7o/view?usp=sharing'
path = 'https://drive.google.com/uc?export=download&id='+url.split('/')[-2]
df = pd.read_csv(path, sep=',')
df = df.values.tolist()

url = 'https://drive.google.com/file/d/1rHcavR0fRGS_t559uZ2wjfqepet_XxEh/view?usp=sharing'
path = 'https://drive.google.com/uc?export=download&id='+url.split('/')[-2]
dfPontos = pd.read_csv(path, sep=',')
dfPontos = dfPontos.values.tolist()

pontosTipos = [0, 0, 0, 0, 0, 0]


def perguntas(posicao):
    valor = []
    contador = []
    for i in range(len(dfPontos)):
        if int(dfPontos[i][0]) < posicao:
            continue
        if int(dfPontos[i][0]) > posicao:
            break
        if int(dfPontos[i][0]) == posicao:
            contador.append(0)
            valor.append(dfPontos[i][1])
    for i in df:
        if isinstance(i[posicao + 2], float) == False:
            contador[valor.index(i[posicao + 2])] += 1
    return contarPontos(posicao, contador)

def contarPontos(posicao, valores):
    contador = 0
    pontosPorGrafico = [0, 0, 0, 0, 0, 0]
    for i in range(len(dfPontos)):
        if int(dfPontos[i][0]) < posicao:
            continue
        if int(dfPontos[i][0]) > posicao:
            break
        if int(dfPontos[i][0]) == posicao:
            for j in range(2, 8):
                if math.isnan(dfPontos[i][j]) == False:
                    pontosTipos[j - 2] += 1*int(dfPontos[i][j])*valores[contador]
                    pontosPorGrafico[j - 2] += 1*int(dfPontos[i][j])*valores[contador]
            contador += 1
    return pontosPorGrafico

dataFrames = []
valores = []
colunas = ['Free Spirit', 'Philanthropist', 'Socializer', 'Achiever', 'Disruptor', 'Player']
for i in range(1, dfPontos[-1][0] + 1):
    valores.append(perguntas(i))

total = pd.DataFrame([pontosTipos],
                  columns= colunas)

path = 'Pontuação.xlsx'

for i in range(len(valores)):
    dataFrames.append(pd.DataFrame([valores[i]],
                  columns= colunas))

with pd.ExcelWriter(path) as writer:
    total.to_excel(writer, sheet_name='Total', index=False)
    for i in range(len(dataFrames)):
        dataFrames[i].to_excel(writer, sheet_name=f'Pergunta {i + 1}', index=False)
