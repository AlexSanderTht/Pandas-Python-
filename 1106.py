import pandas as pd
import os
import time
import hvplot.pandas
import panel as pn
caminho = r'C:\Users\e806128\Desktop\ipdos'
lista_arquivos = os.listdir(caminho)
lista_datas = []
nomes = []
ipdo1 = []
caminhos_arquivos = []


for arquivo in lista_arquivos:
    data = os.path.getmtime(F"{caminho}/{arquivo}")
    data =   time.ctime(data)
    lista_datas.append((data, arquivo))

if os.path.exists(caminho):
    arquivos_na_pasta = os.listdir(caminho)

    for arquivo in arquivos_na_pasta:
        caminho_completo = os.path.join(caminho, arquivo)
        if os.path.isfile(caminho_completo):
            caminhos_arquivos.append(caminho_completo)

def Primeiro(ipdo):
    colunas =  {6: 'Hidro Nac',
            7: 'Itaip',
            8: 'Termo Nuc',
            9: 'Termo Conv',
            10: 'Eólica',
            11: 'Solar',
            12: 'Total SIN ',
            13: 'Interc. Inter',
            14: 'Carga',
            15: 'Interc. Inter.'}
    for caminho in caminhos_arquivos:
        ipdo1 = [pd.read_excel(arquivo, sheet_name="IPDO") for arquivo in caminhos_arquivos]

        nomes = [os.path.basename(arquivo).split('/')[-1].replace(".xlsm", "") for arquivo in caminhos_arquivos]
        n = 0
        cont = 1
        comparacao = []
        comparacao2 = []
        no = 2
        todos_ipdos_programado = pd.DataFrame()
        todos_ipdos_verificado = pd.DataFrame()

    for n, s in enumerate(zip(ipdo1, nomes)):
        if cont < 10:
            n = ipdo1[n]
            s = s[1]
            no = no + 1000
            caminho_arquivo = r'C:\Users\e806128\Desktop\1\prograo_sin.xlsx'
            no = 2
            #print("O arquivo não existe!")
            
            programado = n.iloc[6:16, 10:13]
            programado = (programado.drop(programado.columns[1], axis=1)
                        .transpose()
                        .rename(index={'Unnamed: 12': data})
                        .rename(columns=colunas)
                        .iloc[1:2])
            todos_ipdos_programado = todos_ipdos_programado._append(programado)
            todos_ipdos_programado = todos_ipdos_programado.rename(index={'Mon May  6 09:23:58 2024': str(s)}) 
            todos_ipdos_programado.to_excel(r"C:\Users\e806128\Desktop\1\todos_dados_Programado" + '.xlsx' , index=False)
            programado.to_excel(
                r'C:\Users\e806128\Desktop\1\prograo_sinComparação' + '.xlsx')
            verificado = n.iloc[6:16, 10:15]
            colunass = [1, 2, 3]                                                                                                                                                                                                                                                                                                                                                             

            verificado= (verificado.drop(verificado.columns[colunass], axis=1)
                        .transpose()
                        .rename(columns=colunas)
                        .rename(index={'Unnamed: 14': data})
                        .iloc[1:2])
            todos_ipdos_verificado = todos_ipdos_verificado._append(verificado)
            todos_ipdos_verificado = todos_ipdos_verificado.rename(index={'Mon May  6 09:23:58 2024': str(s)}) 
            todos_ipdos_verificado.to_excel(r"C:\Users\e806128\Desktop\1\todos_dados_Verificado" + '.xlsx' , index=False)
            verificado.to_excel(
                r"C:\Users\e806128\Desktop\1\verifi_sin" + '.xlsx')

            cont += 1
    return todos_ipdos_programado, todos_ipdos_verificado


def plotar(todos_ipdos_programado, todos_ipdos_verificado):
    plot1 = todos_ipdos_programado.hvplot(columns=['Hidro Nac'], index=['Hidro Nac', 'Itaip', 'Termo Nuc', 'Termo Conv', 'Eólica', 'Solar', 'Total SIN', 'Interc. Inter', 'Carga', 'Interc. Inte'], kind='line', title='Programado Sin', ylabel='MW', xlabel='Colunas', legend='top_right', grid=True, height=500, width=1000).opts(yformatter='%.0f')

    plot2 = todos_ipdos_verificado.hvplot(columns=['Hidro Nac'], index=['Hidro Nac', 'Itaip', 'Termo Nuc', 'Termo Conv', 'Eólica', 'Solar', 'Total SIN', 'Interc. Inter', 'Carga', 'Interc. Inte'], kind='line', title='Verificado Sin', ylabel='MW', xlabel='Colunas', legend='top_right', grid=True, height=500, width=1000).opts(yformatter='%.0f')
    pn.panel(plot1 + plot2).show()

if __name__ == "__main__":
    ipdo = pd.read_excel(r'C:\Users\e806128\Downloads\IPDO-13-03-2024.xlsm',
                         sheet_name='IPDO')
plotar(todos_ipdos_programado , todos_ipdos_verificado)
