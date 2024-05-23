import pandas as pd
import os
import time
caminho = r'C:\Users\e806128\Desktop\ipdos'
lista_arquivos = os.listdir(caminho)
lista_datas = []
for arquivo in lista_arquivos:
    data = os.path.getmtime(F"{caminho}/{arquivo}")
    data = time.ctime(data)
    lista_datas.append((data, arquivo))
    df = pd.DataFrame(lista_datas)
df1 = df.iloc[0, 0]
data = pd.to_datetime(df1, dayfirst=True)


linha_inicio = 1
dado1 = data
"""Alterar a pasta de armazenamento dos xlsx, e alterar o caminho ali cima,
juntamente com o caminho do ipdo no main no final"""


def primeiro(ipdo : pd.DataFrame):
    colunasdpai = ({6: 'Hidro Nac',
                    7: 'Itaip',
                    8: 'Termo Nuc',
                    9: 'Termo Conv',
                    10: 'Eólica',
                    11: 'Solar',
                    12: 'Total SIN',
                    13: 'Interc.Inter',
                    14: 'Carga',
                    15: 'Interc.Inter'})
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\prograo_sin.xlsx'
    no = 2
    sin = ipdo.iloc[6:16, 10:13]
    sin1 = sin.drop(sin.columns[1], axis=1)
    sin1_d = sin1.transpose()

    sin2 = ipdo.iloc[6:16, 10:15]
    colunass = [1, 2, 3]
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(r'C:\Users\e806128\Desktop\1\prograo_sin.xlsx')
        a_in1 = pd.read_excel(r"C:\Users\e806128\Desktop\1\verifi_sin.xlsx")
        sin1_re1 = sin1_d.rename(index={'Unnamed: 12': data})
        sin1_f = sin1_re1.iloc[1:2]
        sin3 = sin2.drop(sin2.columns[colunass], axis=1)
        sin3_f = sin3.transpose()
        sin2_re1 = sin3_f.rename(index={'Unnamed: 14': data})
        sin2_v = sin2_re1.iloc[1:2]
        no = no + 1000
        add_inf.loc[no] = [data] + sin1_f.iloc[0, :11].tolist()
        add_inf.rename(columns={'Unnamed: 0': ' '}, inplace=True)
        add_inf.to_excel(r'C:\Users\e806128\Desktop\1\prograo_sin.xlsx', index=False)
        a_in1.loc[no] = [data] + sin2_v.iloc[0, :10].tolist()
        a_in1.rename(columns={'Unnamed: 0': ' '}, inplace=True)
        a_in1.to_excel(r'C:\Users\e806128\Desktop\1\verifi_sin.xlsx', index=False)
    else:
        print("O arquivo não existe.")
        sin1_re = sin1_d.rename(columns=colunasdpai)
        sin1_re1 = sin1_re.rename(index={'Unnamed: 12': data})
        sin1_f = sin1_re1.iloc[1:2]
        sin1_f.to_excel(r'C:\Users\e806128\Desktop\1\prograo_sin.xlsx')
        sin3 = sin2.drop(sin2.columns[colunass], axis=1)
        sin3_f = sin3.transpose()
        sin2_re = sin3_f.rename(columns=colunasdpai)
        sin2_re1 = sin2_re.rename(index={'Unnamed: 14': data})
        sin2_v = sin2_re1.iloc[1:2]
        sin2_v.to_excel(r"C:\Users\e806128\Desktop\1\verifi_sin.xlsx")


if __name__ == "__main__":
    ipdo = pd.read_excel(r'C:\Users\e806128\Downloads\IPDO-13-03-2024.xlsm',
                         sheet_name='IPDO')
primeiro(ipdo)
