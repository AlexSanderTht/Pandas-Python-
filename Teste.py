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


def dadosprimeiro(ipdo: pd.DataFrame):
    """
    data,colunass,transpose.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito valor verificado
    e o programado do nordeste em uma pasta pré-escolhida.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\prograo_sin.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(r'C:\Users\e806128\Desktop\1\prograo_sin.xlsx')
        a_in1 = pd.read_excel(r"C:\Users\e806128\Desktop\1\verifi_sin.xlsx")

        sin = ipdo.iloc[6:16, 10:13]
        sin1 = sin.drop(sin.columns[1], axis=1)
        sin1_d = sin1.transpose()

        sin1_re1 = sin1_d.rename(index={'Unnamed: 12': data})
        sin1_f = sin1_re1.iloc[1:2]
        sin2 = ipdo.iloc[6:16, 10:15]
        colunass = [1, 2, 3]
        sin3 = sin2.drop(sin2.columns[colunass], axis=1)
        sin3_f = sin3.transpose()
        sin2_re1 = sin3_f.rename(index={'Unnamed: 14': data})
        sin2_v = sin2_re1.iloc[1:2]
        no = no + 3
        dado1 = data
        dado2 = sin1_f.iloc[0, 0]
        dado3 = sin1_f.iloc[0, 1]
        dado4 = sin1_f.iloc[0, 2]
        dado5 = sin1_f.iloc[0, 3]
        dado6 = sin1_f.iloc[0, 4]
        dado7 = sin1_f.iloc[0, 5]
        dado8 = sin1_f.iloc[0, 6]
        dado9 = sin1_f.iloc[0, 7]
        dado10 = sin1_f.iloc[0, 8]
        dado11 = sin1_f.iloc[0, 9]
        dado21 = sin2_v.iloc[0, 0]
        dado22 = sin2_v.iloc[0, 1]
        dado23 = sin2_v.iloc[0, 2]
        dado24 = sin2_v.iloc[0, 3]
        dado25 = sin2_v.iloc[0, 4]
        dado26 = sin2_v.iloc[0, 5]
        dado27 = sin2_v.iloc[0, 6]
        dado28 = sin2_v.iloc[0, 7]
        dado29 = sin2_v.iloc[0, 8]
        dado30 = sin2_v.iloc[0, 9]

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           dado9,
                           dado10,
                           dado11)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(r'C:\Users\e806128\Desktop\1\prograo_sin.xlsx',
                         index=False)
        a_in1.loc[no] = (dado1,
                         dado21,
                         dado22,
                         dado23,
                         dado24,
                         dado25,
                         dado26,
                         dado27,
                         dado28,
                         dado29,
                         dado30)
        a_in1 = a_in1.rename(columns={'Unnamed: 0': ' '})
        a_in1.to_excel(r"C:\Users\e806128\Desktop\1\verifi_sin.xlsx",
                       index=False)

    else:
        print("O arquivo não existe.")

        sin = ipdo.iloc[6:16, 10:13]
        sin1 = sin.drop(sin.columns[1], axis=1)
        sin1_d = sin1.transpose()
        sin1_re = sin1_d.rename(columns={6: 'Hidro Nac',
                                         7: 'Itaip',
                                         8: 'Termo Nuc',
                                         9: 'Termo Conv',
                                         10: 'Eólica',
                                         11: 'Solar',
                                         12: 'Total SIN',
                                         13: 'Interc.Inter',
                                         14: 'Carga',
                                         15: 'Interc.Inter'})
        sin1_re1 = sin1_re.rename(index={'Unnamed: 12': data})
        sin1_f = sin1_re1.iloc[1:2]
        sin1_f.to_excel(r'C:\Users\e806128\Desktop\1\prograo_sin.xlsx')
        sin2 = ipdo.iloc[6:16, 10:15]
        colunass = [1, 2, 3]
        sin3 = sin2.drop(sin2.columns[colunass], axis=1)
        sin3_f = sin3.transpose()
        sin2_re = sin3_f.rename(columns={6: 'Hidro Nac',
                                         7: 'Itaip',
                                         8: 'Termo Nuc',
                                         9: 'Termo Conv',
                                         10: 'Eólica',
                                         11: 'Solar',
                                         12: 'Total SIN ',
                                         13: 'Interc. Inter',
                                         14: 'Carga',
                                         15: 'Interc. Inter '})
        sin2_re1 = sin2_re.rename(index={'Unnamed: 14': data})
        sin2_v = sin2_re1.iloc[1:2]
        sin2_v.to_excel(r"C:\Users\e806128\Desktop\1\verifi_sin.xlsx")


def intercambio(ipdo: pd.DataFrame):
    """
    data,colums,transpose.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito do intercâmbio
    de energia em uma pasta pré-escolhida.
    -------
    """
    caminho_arquivo = r"C:\Users\e806128\Desktop\1\Verifi_Intcâm.xlsx"
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(r"C:\Users\e806128\Desktop\1\prog_Intcâm.xlsx")
        a_in1 = pd.read_excel(r"C:\Users\e806128\Desktop\1\Verifi_Intcâm.xlsx")
        inte = ipdo.iloc[17:21, 17:20]
        inte1 = inte.drop(inte.columns[1], axis=1)
        inte1_d = inte1.transpose()
        inte1_re1 = inte1_d.rename(index={'Unnamed: 19': data})
        inte1_f = inte1_re1.iloc[1:2]

        inte2 = ipdo.iloc[17:21, 17:22]
        colunassi = [1, 2, 3]
        inte3 = inte2.drop(inte2.columns[colunassi], axis=1)
        inte3_f = inte3.transpose()
        inte2_re1 = inte3_f.rename(index={'Unnamed: 21': data})
        inte2_v = inte2_re1.iloc[1:2]

        no = no + 3
        dado1 = data
        dado2 = inte1_f.iloc[0, 0]
        dado3 = inte1_f.iloc[0, 1]
        dado4 = inte1_f.iloc[0, 2]
        dado5 = inte1_f.iloc[0, 3]
        dado21 = inte2_v.iloc[0, 0]
        dado22 = inte2_v.iloc[0, 1]
        dado23 = inte2_v.iloc[0, 2]
        dado24 = inte2_v.iloc[0, 3]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(r"C:\Users\e806128\Desktop\1\prog_Intcâm.xlsx",
                         index=False)
        a_in1.loc[no] = (dado1,
                         dado21,
                         dado22,
                         dado23,
                         dado24,)
        a_in1 = a_in1.rename(columns={'Unnamed: 0': ' '})
        a_in1.to_excel(r"C:\Users\e806128\Desktop\1\Verifi_Intcâm.xlsx",
                       index=False)
    else:
        print("O arquivo não existe.")
        inte = ipdo.iloc[17:21, 17:20]
        inte1 = inte.drop(inte.columns[1], axis=1)
        inte1_d = inte1.transpose()
        inte1_re = inte1_d.rename(columns={17: 'Interc N',
                                           18: 'Interc NE',
                                           19: 'Interc SE',
                                           20: 'Interc S'})
        inte1_re1 = inte1_re.rename(index={'Unnamed: 19': data})
        inte1_f = inte1_re1.iloc[1:2]
        inte1_f.to_excel(r"C:\Users\e806128\Desktop\1\prog_Intcâm.xlsx")

        inte2 = ipdo.iloc[17:21, 17:22]
        colunassi = [1, 2, 3]
        inte3 = inte2.drop(inte2.columns[colunassi], axis=1)
        inte3_f = inte3.transpose()
        inte2_re = inte3_f.rename(columns={17: 'Interc N',
                                           18: 'Interc NE',
                                           19: 'Interc SE',
                                           20: 'Interc S'})
        inte2_re1 = inte2_re.rename(index={'Unnamed: 21': data})
        inte2_v = inte2_re1.iloc[1:2]
        inte2_v.to_excel(r"C:\Users\e806128\Desktop\1\Verifi_Intcâm.xlsx")


def internacional(ipdo):
    """
    data,colunassr,transpose.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito valor verificado
    e o programado do nordeste em uma pasta pré-escolhida.
    -------
    """
    caminho_arquivo = r"C:\Users\e806128\Desktop\1\Resul_inter.xlsx"
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(r"C:\Users\e806128\Desktop\1\Resul_inter.xlsx")
        inter2 = ipdo.iloc[38:47, 16:24]
        colunassr = [1, 2, 3, 4, 5, 6]
        inter3 = inter2.drop(inter2.columns[colunassr], axis=1)
        intert = inter3.iloc[3:9]
        inter3_f = intert.transpose()
        inter2_re1 = inter3_f.rename(index={'Unnamed: 23': data})
        inter2 = inter2_re1.iloc[1:2]

        dado1 = data
        dado2 = inter2.iloc[0, 0]
        dado3 = inter2.iloc[0, 1]
        dado4 = inter2.iloc[0, 2]
        dado5 = inter2.iloc[0, 3]
        dado6 = inter2.iloc[0, 4]
        dado7 = inter2.iloc[0, 5]
        no = no + 3
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(r"C:\Users\e806128\Desktop\1\Resul_inter.xlsx",
                         index=False)
    else:
        print("O arquivo não existe.")
        inter2 = ipdo.iloc[38:47, 16:24]
        colunassr = [1, 2, 3, 4, 5, 6]
        inter3 = inter2.drop(inter2.columns[colunassr], axis=1)
        intert = inter3.iloc[3:9]
        inter3_f = intert.transpose()
        inter2_re = inter3_f.rename(columns={41: ' Acaray',
                                             42: 'Uruguaiana',
                                             43: 'Garabi I',
                                             44: 'Garabi II',
                                             45: 'Rivera',
                                             46: 'Melo'})
        inter2_re1 = inter2_re.rename(index={'Unnamed: 23': data})
        inter2 = inter2_re1.iloc[1:2]
        inter2.to_excel(r"C:\Users\e806128\Desktop\1\Resul_inter.xlsx")


def itaipu(ipdo: pd.DataFrame):
    """
    data,colunassn,transpose.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito valor verificado
    e o programado do nordeste em uma pasta pré-escolhida.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\Veri_Itaipu.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(r'C:\Users\e806128\Desktop\1\Veri_Itaipu.xlsx')
        a_in1 = pd.read_excel(r'C:\Users\e806128\Desktop\1\progr_Itaipu.xlsx')
        itau = ipdo.iloc[28:30, 17:20]
        itau1 = itau.drop(itau.columns[1], axis=1)
        itau1_d = itau1.transpose()

        itau1_re1 = itau1_d.rename(index={'Unnamed: 19': data})
        itau1_f = itau1_re1.iloc[1:2]
        itau1_f.to_excel(r'C:\Users\e806128\Desktop\1\Veri_Itaipu.xlsx',)
        itau2 = ipdo.iloc[28:30, 17:22]
        colunassitau = [1, 2]
        itau3 = itau2.drop(itau2.columns[colunassitau], axis=1)
        itau3_f = itau3.transpose()

        itau2_re1 = itau3_f.rename(index={'Unnamed: 21': data})
        itau2_v = itau2_re1.iloc[2:3]
        itau2_v.to_excel(r'C:\Users\e806128\Desktop\1\progr_Itaipu.xlsx', )
        dado1 = data
        dado2 = itau1_f.iloc[0, 0]
        dado3 = itau1_f.iloc[0, 1]
        dado4 = itau2_v.iloc[0, 0]
        dado5 = itau2_v.iloc[0, 1]
        no = no + 3

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(r'C:\Users\e806128\Desktop\1\Veri_Itaipu.xlsx',
                         index=False)
        a_in1.loc[no] = (dado1,
                         dado4,
                         dado5)
        a_in1 = a_in1.rename(columns={'Unnamed: 0': ' '})
        a_in1.to_excel(r'C:\Users\e806128\Desktop\1\progr_Itaipu.xlsx',
                       index=False)
    else:
        print("O arquivo não existe.")
        itau = ipdo.iloc[28:30, 17:20]
        itau1 = itau.drop(itau.columns[1], axis=1)
        itau1_d = itau1.transpose()
        itau2_re = itau1_d.rename(columns={28: 'Elo 50Hz',
                                           29: 'Itai. 60Hz'})
        itau1_re1 = itau2_re.rename(index={'Unnamed: 19': data})
        itau1_f = itau1_re1.iloc[1:2]
        itau1_f.to_excel(r'C:\Users\e806128\Desktop\1\Veri_Itaipu.xlsx',)
        itau2 = ipdo.iloc[28:30, 17:22]
        colunassitau = [1, 2]
        itau3 = itau2.drop(itau2.columns[colunassitau], axis=1)
        itau3_f = itau3.transpose()
        itau2_re = itau3_f.rename(columns={28: 'Elo 50Hz',
                                           29: 'Itai. 60Hz'})
        itau2_re1 = itau2_re.rename(index={'Unnamed: 21': data})
        itau2_v = itau2_re1.iloc[2:3]
        itau2_v.to_excel(r'C:\Users\e806128\Desktop\1\progr_Itaipu.xlsx', )


def nordeste(ipdo: pd.DataFrame):
    """
    data,colunassn,transpose.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito valor verificado
    e o programado do nordeste em uma pasta pré-escolhida.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\progr_Nordt.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(r'C:\Users\e806128\Desktop\1\progr_Nordt.xlsx')
        a_in1 = pd.read_excel(r'C:\Users\e806128\Desktop\1\Veri_Nordeste.xlsx')
        nordestd = ipdo.iloc[25:31, 10:13]
        nordest1 = nordestd.drop(nordestd.columns[1], axis=1)
        nordest1_d = nordest1.transpose()
        nordest1_re1 = nordest1_d.rename(index={'Unnamed: 12': data})
        norde_f = nordest1_re1.iloc[1:2]

        nordest2 = ipdo.iloc[25:31, 10:15]
        colunassn = [1, 2, 3]
        nordest3 = nordest2.drop(nordest2.columns[colunassn], axis=1)
        nordest3_f = nordest3.transpose()

        nordest2_re1 = nordest3_f.rename(index={'Unnamed: 14': data})
        nord_v = nordest2_re1.iloc[1:2]

        dado1 = data
        dado2 = norde_f.iloc[0, 0]
        dado3 = norde_f.iloc[0, 1]
        dado4 = norde_f.iloc[0, 2]
        dado5 = norde_f.iloc[0, 3]
        dado6 = norde_f.iloc[0, 4]
        dado7 = norde_f.iloc[0, 5]
        dado8 = nord_v.iloc[0, 0]
        dado9 = nord_v.iloc[0, 1]
        dado10 = nord_v.iloc[0, 2]
        dado11 = nord_v.iloc[0, 3]
        dado12 = nord_v.iloc[0, 4]
        dado13 = nord_v.iloc[0, 5]
        no = no + 3
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(r'C:\Users\e806128\Desktop\1\progr_Nordt.xlsx',
                         index=False)
        a_in1.loc[no] = (dado1,
                         dado8,
                         dado9,
                         dado10,
                         dado11,
                         dado12,
                         dado13)
        a_in1 = a_in1.rename(columns={'Unnamed: 0': ' '})
        a_in1.to_excel(r'C:\Users\e806128\Desktop\1\Veri_Nordeste.xlsx',
                       index=False)
    else:
        print("O arquivo não existe.")
        nordestd = ipdo.iloc[25:31, 10:13]
        nordest1 = nordestd.drop(nordestd.columns[1], axis=1)
        nordest1_d = nordest1.transpose()
        nordest1_re = nordest1_d.rename(columns={25: 'Hidro',
                                                 26: 'Termo',
                                                 27: 'Eólica',
                                                 28: 'Solar',
                                                 29: 'Total Ger',
                                                 30: 'Carga'})
        nordest1_re1 = nordest1_re.rename(index={'Unnamed: 12': data})
        norde_f = nordest1_re1.iloc[1:2]
        norde_f.to_excel(r'C:\Users\e806128\Desktop\1\progr_Nordt.xlsx',)
        nordest2 = ipdo.iloc[25:31, 10:15]
        colunassn = [1, 2, 3]
        nordest3 = nordest2.drop(nordest2.columns[colunassn], axis=1)
        nordest3_f = nordest3.transpose()
        nordest2_re = nordest3_f.rename(columns={25: 'Hidro',
                                                 26: 'Termo',
                                                 27: 'Eólica',
                                                 28: 'Solar',
                                                 29: 'Total Ger',
                                                 30: 'Carga'})
        nordest2_re1 = nordest2_re.rename(index={'Unnamed: 14': data})
        nord_v = nordest2_re1.iloc[1:2]
        nord_v.to_excel(r'C:\Users\e806128\Desktop\1\Veri_Nordeste.xlsx',)


def roraima(ipdo: pd.DataFrame):
    """
    data,colunassr,transpose.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito valor verificado
    e o programado do roraima em uma pasta pré-escolhida.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\prog_Roráima.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(r'C:\Users\e806128\Desktop\1\Ver_Roraíma.xlsx')
        a_in1 = pd.read_excel(r'C:\Users\e806128\Desktop\1\prog_Roráima.xlsx')
        ror = ipdo.iloc[22:26, 17:20]
        ror1 = ror.drop(ror.columns[1], axis=1)
        ror1_d = ror1.transpose()

        ror1_re1 = ror1_d.rename(index={'Unnamed: 19': data})
        ror1_f = ror1_re1.iloc[1:2]

        ror2 = ipdo.iloc[22:26, 17:22]
        colunassr = [1, 2, 3]
        ror3 = ror2.drop(ror2.columns[colunassr], axis=1)
        ror3_f = ror3.transpose()
        ror2_re1 = ror3_f.rename(index={'Unnamed: 21': data})
        ror2_v = ror2_re1.iloc[1:2]

        dado1 = data
        dado2 = ror1_f.iloc[0, 0]
        dado3 = ror1_f.iloc[0, 1]
        dado4 = ror1_f.iloc[0, 2]
        dado5 = ror1_f.iloc[0, 3]

        dado7 = ror2_v.iloc[0, 0]
        dado8 = ror2_v.iloc[0, 1]
        dado9 = ror2_v.iloc[0, 2]
        dado10 = ror2_v.iloc[0, 3]
        no = no + 3

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(r'C:\Users\e806128\Desktop\1\Ver_Roraíma.xlsx',
                         index=False)
        a_in1.loc[no] = (dado1,
                         dado7,
                         dado8,
                         dado9,
                         dado10)
        a_in1 = a_in1.rename(columns={'Unnamed: 0': ' '})
        a_in1.to_excel(r'C:\Users\e806128\Desktop\1\prog_Roráima.xlsx',
                       index=False)
    else:
        print("O arquivo não existe.")
        ror = ipdo.iloc[22:26, 17:20]
        ror1 = ror.drop(ror.columns[1], axis=1)
        ror1_d = ror1.transpose()
        ror1_re = ror1_d.rename(columns={22: 'Termo',
                                         23: 'Interc',
                                         24: 'Total',
                                         25: 'Carga'})
        ror1_re1 = ror1_re.rename(index={'Unnamed: 19': data})
        ror1_f = ror1_re1.iloc[1:2]
        ror1_f.to_excel(r'C:\Users\e806128\Desktop\1\Ver_Roraíma.xlsx')
        ror2 = ipdo.iloc[22:26, 17:22]
        colunassr = [1, 2, 3]
        ror3 = ror2.drop(ror2.columns[colunassr], axis=1)
        ror3_f = ror3.transpose()
        ror2_re = ror3_f.rename(columns={22: 'Termo',
                                         23: 'Interc',
                                         24: 'Total',
                                         25: 'Carga'})
        ror2_re1 = ror2_re.rename(index={'Unnamed: 21': data})
        ror2_v = ror2_re1.iloc[1:2]
        ror2_v.to_excel(r'C:\Users\e806128\Desktop\1\prog_Roráima.xlsx')


def norte(ipdo: pd.DataFrame):
    """
   data,colunass,transpose,to_excel, n_index.
   ----------
   ipdo : DataFrame da tabela.

   Retorna um arquivo .xlsx em formato de tabela a respeito valor verificado
   e o programado do norte em uma pasta pré-escolhida.
   -------
   """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\pro_Norte.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(r'C:\Users\e806128\Desktop\1\pro_Norte.xlsx')
        a_in1 = pd.read_excel(r'C:\Users\e806128\Desktop\1\Veri_Norte.xlsx')
        norte = ipdo.iloc[17:23, 10:13]
        norte1 = norte.drop(norte.columns[1], axis=1)
        norte1_d = norte1.transpose()
        norte1_re1 = norte1_d.rename(index={'Unnamed: 12': data})
        norte1_f = norte1_re1.iloc[1:2]
        norte2 = ipdo.iloc[17:23, 10:15]
        colunass = [1, 2, 3]
        norte3 = norte2.drop(norte2.columns[colunass], axis=1)
        norte3_f = norte3.transpose()
        norte2_re1 = norte3_f.rename(index={'Unnamed: 14': data})
        norte2_v = norte2_re1.iloc[1:2]
        dado1 = data
        dado2 = norte1_f.iloc[0, 0]
        dado3 = norte1_f.iloc[0, 1]
        dado4 = norte1_f.iloc[0, 2]
        dado5 = norte1_f.iloc[0, 3]
        dado6 = norte1_f.iloc[0, 4]
        dado7 = norte1_f.iloc[0, 5]

        dado71 = norte2_v.iloc[0, 0]
        dado8 = norte2_v.iloc[0, 1]
        dado9 = norte2_v.iloc[0, 2]
        dado10 = norte2_v.iloc[0, 3]
        dado11 = norte2_v.iloc[0, 4]
        dado12 = norte2_v.iloc[0, 5]
        no = no + 3
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(r'C:\Users\e806128\Desktop\1\pro_Norte.xlsx',
                         index=False)
        a_in1.loc[no] = (dado1,
                         dado71,
                         dado8,
                         dado9,
                         dado10,
                         dado11,
                         dado12)

        a_in1 = a_in1.rename(columns={'Unnamed: 0': ' '})
        a_in1.to_excel(r'C:\Users\e806128\Desktop\1\Veri_Norte.xlsx',
                       index=False)
    else:
        print("O arquivo não existe.")
        norte = ipdo.iloc[17:23, 10:13]
        norte1 = norte.drop(norte.columns[1], axis=1)
        norte1_d = norte1.transpose()
        norte1_re = norte1_d.rename(columns={17: 'Hidro',
                                             18: 'Termo',
                                             19: 'Eólica',
                                             20: 'Solar',
                                             21: 'Total Ger',
                                             22: 'Carga'})
        norte1_re1 = norte1_re.rename(index={'Unnamed: 12': data})
        norte1_f = norte1_re1.iloc[1:2]
        norte1_f.to_excel(r'C:\Users\e806128\Desktop\1\pro_Norte.xlsx')

        norte2 = ipdo.iloc[17:23, 10:15]
        colunass = [1, 2, 3]
        norte3 = norte2.drop(norte2.columns[colunass], axis=1)
        norte3_f = norte3.transpose()
        norte2_re = norte3_f.rename(columns={17: 'Hidro',
                                             18: 'Termo',
                                             19: 'Eólica',
                                             20: 'Solar',
                                             21: 'Total Ger',
                                             22: 'Carga'})
        norte2_re1 = norte2_re.rename(index={'Unnamed: 14': data})
        norte2_v = norte2_re1.iloc[1:2]
        norte2_v.to_excel(r'C:\Users\e806128\Desktop\1\Veri_Norte.xlsx')


def sc(ipdo: pd.DataFrame):
    """
   data,colunass,transpose,to_excel, n_index.
   ----------
   ipdo : DataFrame da tabela.

   Retorna um arquivo .xlsx em formato de tabela a respeito valor verificado
   e o programado do norte em uma pasta pré-escolhida.
   -------
   """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\progr_SE-CO.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(r'C:\Users\e806128\Desktop\1\progr_SE-CO.xlsx')
        a_in1 = pd.read_excel(r'C:\Users\e806128\Desktop\1\Veri_SE-CO.xlsx')
        sc = ipdo.iloc[33:39, 10:13]
        sc1 = sc.drop(sc.columns[1], axis=1)
        sc1_d = sc1.transpose()
        sc1_re1 = sc1_d.rename(index={'Unnamed: 12': data})
        sc1_f = sc1_re1.iloc[1:2]
        sc2 = ipdo.iloc[33:39, 10:15]
        colunassc = [1, 2, 3]
        sc3 = sc2.drop(sc2.columns[colunassc], axis=1)
        sc3_f = sc3.transpose()
        sc2_re1 = sc3_f.rename(index={'Unnamed: 14': data})
        sc2_v = sc2_re1.iloc[1:2]

        dado1 = data
        dado2 = sc1_f.iloc[0, 0]
        dado3 = sc1_f.iloc[0, 1]
        dado4 = sc1_f.iloc[0, 2]
        dado5 = sc1_f.iloc[0, 3]
        dado6 = sc1_f.iloc[0, 4]
        dado7 = sc1_f.iloc[0, 5]

        dado8 = sc2_v.iloc[0, 0]
        dado9 = sc2_v.iloc[0, 1]
        dado10 = sc2_v.iloc[0, 2]
        dado11 = sc2_v.iloc[0, 3]
        dado12 = sc2_v.iloc[0, 4]
        dado13 = sc2_v.iloc[0, 5]
        no = no + 3

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(r'C:\Users\e806128\Desktop\1\progr_SE-CO.xlsx',
                         index=False)
        a_in1.loc[no] = (dado1,
                         dado8,
                         dado9,
                         dado10,
                         dado11,
                         dado12,
                         dado13)
        a_in1 = a_in1.rename(columns={'Unnamed: 0': ' '})
        a_in1.to_excel(r'C:\Users\e806128\Desktop\1\Veri_SE-CO.xlsx',
                       index=False)
    else:
        print("O arquivo não existe.")
        sc = ipdo.iloc[33:39, 10:13]
        sc1 = sc.drop(sc.columns[1], axis=1)
        sc1_d = sc1.transpose()
        sc1_re = sc1_d.rename(columns={33: 'Hidro',
                                       34: 'Termo',
                                       35: 'Eólica',
                                       36: 'Solar',
                                       37: 'Total Ger',
                                       38: 'Carga'})
        sc1_re1 = sc1_re.rename(index={'Unnamed: 12': data})
        sc1_f = sc1_re1.iloc[1:2]
        sc1_f.to_excel(r'C:\Users\e806128\Desktop\1\progr_SE-CO.xlsx', )

        sc2 = ipdo.iloc[33:39, 10:15]
        colunassc = [1, 2, 3]
        sc3 = sc2.drop(sc2.columns[colunassc], axis=1)
        sc3_f = sc3.transpose()
        sc2_re = sc3_f.rename(columns={33: 'Hidro',
                                       34: 'Termo',
                                       35: 'Eólica',
                                       36: 'Solar',
                                       37: 'Total Ger',
                                       38: 'Carga'})
        sc2_re1 = sc2_re.rename(index={'Unnamed: 14': data})

        sc2_v = sc2_re1.iloc[1:2]
        sc2_v.to_excel(r'C:\Users\e806128\Desktop\1\Veri_SE-CO.xlsx', )


def termos(ipdo: pd.DataFrame):
    """
   data,colunasst,transpose,to_excel, n_index.
   ----------
   ipdo : DataFrame da tabela.

   Retorna um arquivo .xlsx em formato de tabela a respeito valor verificado
   e o programado do norte em uma pasta pré-escolhida.
   -------
   """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\prog_Termo.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(r'C:\Users\e806128\Desktop\1\prog_Termo.xlsx')
        a_in1 = pd.read_excel(r'C:\Users\e806128\Desktop\1\Verif_Termo.xlsx')
        termo = ipdo.iloc[34:36, 17:20]
        termo1 = termo.drop(termo.columns[1], axis=1)
        termo1_d = termo1.transpose()
        termo1_re1 = termo1_d.rename(index={'Unnamed: 19': data})
        termo1_f = termo1_re1.iloc[1:2]

        termo2 = ipdo.iloc[34:36, 17:22]
        colunasst = [1, 2, 3]
        termo3 = termo2.drop(termo2.columns[colunasst], axis=1)
        termo3_f = termo3.transpose()
        termo2_re1 = termo3_f.rename(index={'Unnamed: 21': data})
        termo2_v = termo2_re1.iloc[1:2]
        dado1 = data
        dado2 = termo1_f.iloc[0, 0]
        dado3 = termo1_f.iloc[0, 1]
        dado8 = termo2_v.iloc[0, 0]
        dado9 = termo2_v.iloc[0, 1]
        no = no + 3
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(r'C:\Users\e806128\Desktop\1\prog_Termo.xlsx',
                         index=False)
        a_in1.loc[no] = (dado1,
                         dado8,
                         dado9)

        a_in1 = a_in1.rename(columns={'Unnamed: 0': ' '})
        a_in1.to_excel(r'C:\Users\e806128\Desktop\1\Verif_Termo.xlsx',
                       index=False)
    else:
        print("O arquivo não existe.")
        termo = ipdo.iloc[34:36, 17:20]
        termo1 = termo.drop(termo.columns[1], axis=1)
        termo1_d = termo1.transpose()
        termo1_re = termo1_d.rename(columns={34: 'Termo Nuc',
                                             35: 'Termo Conv'})
        termo1_re1 = termo1_re.rename(index={'Unnamed: 19': data})
        termo1_f = termo1_re1.iloc[1:2]

        termo1_f.to_excel(r'C:\Users\e806128\Desktop\1\prog_Termo.xlsx', )
        termo2 = ipdo.iloc[34:36, 17:22]
        colunasst = [1, 2, 3]
        termo3 = termo2.drop(termo2.columns[colunasst], axis=1)
        termo3_f = termo3.transpose()
        termo2_re = termo3_f.rename(columns={34: 'Termo Nuc',
                                             35: 'Termo Conv'})
        termo2_re1 = termo2_re.rename(index={'Unnamed: 21': data})
        termo2_v = termo2_re1.iloc[1:2]

        termo2_v.to_excel(r'C:\Users\e806128\Desktop\1\Verif_Termo.xlsx')


def carga(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito valor verificado
    e o programado do Carga em uma pasta pré-escolhida.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\progr_carga.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(r'C:\Users\e806128\Desktop\1\progr_carga.xlsx')
        a_in1 = pd.read_excel(r'C:\Users\e806128\Desktop\1\Máxi_carga.xlsx')
        carga = ipdo.iloc[41:46, 1:4]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index1 = d_transp.rename(index={'Unnamed: 3': data})
        form_tab = n_index1.iloc[1:2]

        carga2 = ipdo.iloc[41:46, 1:9]
        colunassc = [1, 2, 4, 3, 6]
        carga3 = carga2.drop(carga2.columns[colunassc], axis=1)
        carga3_f = carga3.transpose()

        carga2_re1 = carga3_f.rename(index={'Unnamed: 6': data,
                                            'Unnamed: 8': 'Data Verificação'})
        carga2_v = carga2_re1.iloc[1:3]
        carga_3 = carga2_v.iloc[1:3]

        dado1 = data
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado8 = carga_3.iloc[0, 0]
        dado9 = carga_3.iloc[0, 1]
        dado10 = carga_3.iloc[0, 2]
        dado11 = carga_3.iloc[0, 3]
        dado12 = carga_3.iloc[0, 4]
        dado21 = carga2_v.iloc[0, 0]
        dado31 = carga2_v.iloc[0, 1]
        dado41 = carga2_v.iloc[0, 2]
        dado51 = carga2_v.iloc[0, 3]
        dado61 = carga2_v.iloc[0, 4]
        no = no + 3
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(r'C:\Users\e806128\Desktop\1\progr_carga.xlsx',
                         index=False)
        a_in1.loc[no] = (dado1,
                         dado21,
                         dado31,
                         dado41,
                         dado51,
                         dado61)
        no = len(add_inf) + 1
        dado1 = ('Data Verificação')
        a_in1.loc[no] = (dado1,
                         dado8,
                         dado9,
                         dado10,
                         dado11,
                         dado12)
        a_in1 = a_in1.rename(columns={'Unnamed: 0': ' '})
        a_in1.to_excel(r'C:\Users\e806128\Desktop\1\Máxi_carga.xlsx',
                       index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[41:46, 1:4]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={41: 'Sul',
                                           42: 'Sudeste CO',
                                           43: 'Norte',
                                           44: 'NORDESTE',
                                           45: 'SIN'})
        n_index1 = n_index.rename(index={'Unnamed: 3': data})
        form_tab = n_index1.iloc[1:2]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\progr_carga.xlsx')

        carga2 = ipdo.iloc[41:46, 1:9]
        colunassc = [1, 2, 4, 3, 6]
        carga3 = carga2.drop(carga2.columns[colunassc], axis=1)
        carga3_f = carga3.transpose()
        c2_v = carga3_f.rename(columns={41: 'Sul',
                                        42: 'Sudeste CO',
                                        43: 'Norte',
                                        44: 'NORDESTE',
                                        45: 'SIN'})
        c2_v = c2_v.rename(index={'Unnamed: 6': data,
                                  'Unnamed: 8': 'Data Verificação'})
        c_v = c2_v.iloc[1:3]
        c_v.to_excel(r'C:\Users\e806128\Desktop\1\Máxi_carga.xlsx')


def dados_hidrologicos(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito valor da EnaMwMed,
    EnaBruta, EnaArmaz, ValorEARD, ValorEARD, ValorEardDia, DesvioEar, Varia-
    ção, CapMax e o VarMensal.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\Dados_CapMax.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Dados_Ena Mwmed.xlsx',)
        a_in2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Dados_EnaBruta.xlsx')
        a_in3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Dados_EnaArmaz.xlsx')
        a_in4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Dados_ValorEARD.xlsx')
        a_in5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Dados_ValrEarDiaP100.xlsx')
        a_in6 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Dados_DesvioEAR.xlsx')
        a_in7 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Dados_VariacaoEmP100.xlsx')
        a_in8 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Dados_Variacao.xlsx')
        a_in9 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Dados_CapMax.xlsx')
        a_in10 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Dados_VarMensal.xlsx')
        hidro = ipdo.iloc[60:65, 10:24]
        hidro1 = hidro.drop(hidro.columns[1], axis=1)
        hidro1_d = hidro1.transpose()
        hidro1_re1 = hidro1_d.rename(index={'Unnamed: 12': data})
        hidro1_f = hidro1_re1.iloc[1:2]
        no = no + 3
        dado1 = data
        dado2 = hidro1_f.iloc[0, 0]
        dado3 = hidro1_f.iloc[0, 1]
        dado4 = hidro1_f.iloc[0, 2]
        dado5 = hidro1_f.iloc[0, 3]
        dado6 = hidro1_f.iloc[0, 4]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(r'C:\Users\e806128\Desktop\1\Dados_Ena Mwmed.xlsx',
                         index=False)

        hidro2 = ipdo.iloc[60:65, 10:24]
        colunassh = [1, 2, 3]
        hidro3 = hidro2.drop(hidro2.columns[colunassh], axis=1)
        hidro3_f = hidro3.transpose()
        hidro2_re = hidro3_f.rename(columns={60: 'Norte',
                                             61: 'Nordeste',
                                             62: 'Sul',
                                             63: 'Sudeste',
                                             64: 'SIN'})
        hidro2_re1 = hidro2_re.rename(index={'Unnamed: 14': data})
        hidro2_v = hidro2_re1.iloc[1:2]

        dado7 = hidro2_v.iloc[0, 0]
        dado8 = hidro2_v.iloc[0, 1]
        dado9 = hidro2_v.iloc[0, 2]
        dado10 = hidro2_v.iloc[0, 3]
        dado11 = hidro2_v.iloc[0, 4]
        a_in2.loc[no] = (dado1,
                         dado7,
                         dado8,
                         dado9,
                         dado10,
                         dado11,)
        a_in2 = a_in2.rename(columns={'Unnamed: 0': ' '})
        a_in2.to_excel(r'C:\Users\e806128\Desktop\1\Dados_EnaBruta.xlsx',
                       index=False)
        hidro31 = ipdo.iloc[60:65, 10:24]
        colunassh1 = [1, 2, 3, 4]
        hidro4 = hidro31.drop(hidro31.columns[colunassh1], axis=1)
        hidro4_f = hidro4.transpose()
        hidro3_re = hidro4_f.rename(columns={60: 'Norte',
                                             61: 'Nordeste',
                                             62: 'Sul',
                                             63: 'Sudeste',
                                             64: 'SIN'})
        hidro3_re1 = hidro3_re.rename(index={'Unnamed: 15': data})
        hidro3_v = hidro3_re1.iloc[1:2]
        dado12 = hidro3_v.iloc[0, 0]
        dado13 = hidro3_v.iloc[0, 1]
        dado14 = hidro3_v.iloc[0, 2]
        dado15 = hidro3_v.iloc[0, 3]
        dado16 = hidro3_v.iloc[0, 4]
        a_in3.loc[no] = (dado1,
                         dado12,
                         dado13,
                         dado14,
                         dado15,
                         dado16)
        a_in3 = a_in3.rename(columns={'Unnamed: 0': ' '})
        a_in3.to_excel(r'C:\Users\e806128\Desktop\1\Dados_EnaArmaz.xlsx',
                       index=False)
        hidro4 = ipdo.iloc[60:65, 10:24]
        colunassh2 = [1, 2, 3, 4]
        hidro5 = hidro4.drop(hidro4.columns[colunassh2], axis=1)
        hidro4_f = hidro5.transpose()
        hidro3_re = hidro4_f.rename(columns={60: 'Norte',
                                             61: 'Nordeste',
                                             62: 'Sul',
                                             63: 'Sudeste',
                                             64: 'SIN'})
        hidro3_re1 = hidro3_re.rename(index={'Unnamed: 16': data})
        hidroE = hidro3_re1.iloc[2:3]
        dado12 = hidroE.iloc[0, 0]
        dado13 = hidroE.iloc[0, 1]
        dado14 = hidroE.iloc[0, 2]
        dado15 = hidroE.iloc[0, 3]
        dado16 = hidroE.iloc[0, 4]
        a_in4.loc[no] = (dado1,
                         dado12,
                         dado13,
                         dado14,
                         dado15,
                         dado16)

        a_in4 = a_in4.rename(columns={'Unnamed: 0': ' '})
        a_in4.to_excel(r'C:\Users\e806128\Desktop\1\Dados_ValorEARD.xlsx',
                       index=False)
        hidro5 = ipdo.iloc[60:65, 10:24]
        colunassh3 = [1, 2, 3, 4]
        hidro6 = hidro5.drop(hidro5.columns[colunassh3], axis=1)
        hidro5_f = hidro6.transpose()
        hidro4_re = hidro5_f.rename(columns={60: 'Norte',
                                             61: 'Nordeste',
                                             62: 'Sul',
                                             63: 'Sudeste',
                                             64: 'SIN'})
        hidro4_re1 = hidro4_re.rename(index={'Unnamed: 17': data})
        hidro4_v = hidro4_re1.iloc[3:4]
        dado12 = hidro4_v.iloc[0, 0]
        dado13 = hidro4_v.iloc[0, 1]
        dado14 = hidro4_v.iloc[0, 2]
        dado15 = hidro4_v.iloc[0, 3]
        dado16 = hidro4_v.iloc[0, 4]
        a_in5.loc[no] = (dado1,
                         dado12,
                         dado13,
                         dado14,
                         dado15,
                         dado16)

        a_in5 = a_in5.rename(columns={'Unnamed: 0': ' '})
        a_in5.to_excel(r'C:\Users\e806128\Desktop\1\Dados_ValrEarDiaP100.xlsx',
                       index=False)
        hidro6 = ipdo.iloc[60:65, 10:24]
        colunassh4 = [1, 2, 3, 4]
        hidro7 = hidro6.drop(hidro6.columns[colunassh4], axis=1)
        hidro6_f = hidro7.transpose()
        hidro5_re = hidro6_f.rename(columns={60: 'Norte',
                                             61: 'Nordeste',
                                             62: 'Sul',
                                             63: 'Sudeste',
                                             64: 'SIN'})
        hidro5_re1 = hidro5_re.rename(index={'Unnamed: 19': data})
        hidro5_v = hidro5_re1.iloc[5:6]
        dado12 = hidro5_v.iloc[0, 0]
        dado13 = hidro5_v.iloc[0, 1]
        dado14 = hidro5_v.iloc[0, 2]
        dado15 = hidro5_v.iloc[0, 3]
        dado16 = hidro5_v.iloc[0, 4]
        a_in6.loc[no] = (dado1,
                         dado12,
                         dado13,
                         dado14,
                         dado15,
                         dado16)
        a_in6 = a_in6.rename(columns={'Unnamed: 0': ' '})
        a_in6.to_excel(r'C:\Users\e806128\Desktop\1\Dados_DesvioEAR.xlsx',
                       index=False)
        hidro7 = ipdo.iloc[60:65, 10:24]
        colunassh5 = [1, 2, 3, 4]
        hidro8 = hidro7.drop(hidro7.columns[colunassh5], axis=1)
        hidro7_f = hidro8.transpose()
        hidro6_re = hidro7_f.rename(columns={60: 'Norte',
                                             61: 'Nordeste',
                                             62: 'Sul',
                                             63: 'Sudeste',
                                             64: 'SIN'})
        hidro6_re1 = hidro6_re.rename(index={'Unnamed: 21': data})
        hidro6_v = hidro6_re1.iloc[7:8]

        dado12 = hidro6_v.iloc[0, 0]
        dado13 = hidro6_v.iloc[0, 1]
        dado14 = hidro6_v.iloc[0, 2]
        dado15 = hidro6_v.iloc[0, 3]
        dado16 = hidro6_v.iloc[0, 4]
        a_in7.loc[no] = (dado1,
                         dado12,
                         dado13,
                         dado14,
                         dado15,
                         dado16)
        a_in7 = a_in7.rename(columns={'Unnamed: 0': ' '})
        a_in7.to_excel(r'C:\Users\e806128\Desktop\1\Dados_VariacaoEmP100.xlsx',
                       index=False)
        hidro8 = ipdo.iloc[60:65, 10:24]
        colunassh6 = [1, 2, 3, 4]
        hidro9 = hidro8.drop(hidro8.columns[colunassh6], axis=1)
        hidro8_f = hidro9.transpose()
        hidro7_re = hidro8_f.rename(columns={60: 'Norte',
                                             61: 'Nordeste',
                                             62: 'Sul',
                                             63: 'Sudeste',
                                             64: 'SIN'})
        hidro7_re1 = hidro7_re.rename(index={'Unnamed: 23': data})
        hidro7_v = hidro7_re1.iloc[9:10]

        dado12 = hidro7_v.iloc[0, 0]
        dado13 = hidro7_v.iloc[0, 1]
        dado14 = hidro7_v.iloc[0, 2]
        dado15 = hidro7_v.iloc[0, 3]
        dado16 = hidro7_v.iloc[0, 4]
        a_in8.loc[no] = (dado1,
                         dado12,
                         dado13,
                         dado14,
                         dado15,
                         dado16)

        a_in8 = a_in8.rename(columns={'Unnamed: 0': ' '})
        a_in8.to_excel(r'C:\Users\e806128\Desktop\1\Dados_Variacao.xlsx',
                       index=False)
        hidro9 = ipdo.iloc[68:73, 10:24]
        hidro10 = hidro9.drop(hidro9.columns[1], axis=1)
        hidro10_d = hidro10.transpose()
        hidro10_re = hidro10_d.rename(columns={68: 'Norte',
                                               69: 'Nordeste',
                                               70: 'Sul',
                                               71: 'Sudeste',
                                               72: 'SIN'})
        hidro10_re1 = hidro10_re.rename(index={'Unnamed: 12': data})
        hidro10_f = hidro10_re1.iloc[1:2]
        dado12 = hidro10_f.iloc[0, 0]
        dado13 = hidro10_f.iloc[0, 1]
        dado14 = hidro10_f.iloc[0, 2]
        dado15 = hidro10_f.iloc[0, 3]
        dado16 = hidro10_f.iloc[0, 4]
        a_in9.loc[no] = (dado1,
                         dado12,
                         dado13,
                         dado14,
                         dado15,
                         dado16)

        a_in9 = a_in9.rename(columns={'Unnamed: 0': ' '})
        a_in9.to_excel(r'C:\Users\e806128\Desktop\1\Dados_CapMax.xlsx',
                       index=False)
        hidro10 = ipdo.iloc[68:73, 10:24]
        colunassv = [1, 2, 3]
        hidro11 = hidro10.drop(hidro10.columns[colunassv], axis=1)
        hidro11_f = hidro11.transpose()
        hidro11_re = hidro11_f.rename(columns={68: 'Norte',
                                               69: 'Nordeste',
                                               70: 'Sul',
                                               71: 'Sudeste',
                                               72: 'SIN'})
        hidro10_re1 = hidro11_re.rename(index={'Unnamed: 14': data})
        hidro10_v = hidro10_re1.iloc[1:2]
        hidro10_v.to_excel(r'C:\Users\e806128\Desktop\1\Dados_VarMensal.xlsx')

        dado12 = hidro10_v.iloc[0, 0]
        dado13 = hidro10_v.iloc[0, 1]
        dado14 = hidro10_v.iloc[0, 2]
        dado15 = hidro10_v.iloc[0, 3]
        dado16 = hidro10_v.iloc[0, 4]
        a_in10.loc[no] = (
                        dado1,
                        dado12,
                        dado13,
                        dado14,
                        dado15,
                        dado16)

        a_in10 = a_in10.rename(columns={'Unnamed: 0': ' '})
        a_in10.to_excel(r'C:\Users\e806128\Desktop\1\Dados_VarMensal.xlsx',
                        index=False)
    else:
        print("O arquivo não existe.")
        hidro = ipdo.iloc[60:65, 10:24]
        hidro1 = hidro.drop(hidro.columns[1], axis=1)
        hidro1_d = hidro1.transpose()
        hidro1_re = hidro1_d.rename(columns={60: 'Norte',
                                             61: 'Nordeste',
                                             62: 'Sul',
                                             63: 'Sudeste',
                                             64: 'SIN'})
        hidro1_re1 = hidro1_re.rename(index={'Unnamed: 12': data})
        hidro1_f = hidro1_re1.iloc[1:2]
        hidro1_f.to_excel(r'C:\Users\e806128\Desktop\1\Dados_Ena Mwmed.xlsx', )
        hidro2 = ipdo.iloc[60:65, 10:24]
        colunassh = [1, 2, 3]
        hidro3 = hidro2.drop(hidro2.columns[colunassh], axis=1)
        hidro3_f = hidro3.transpose()
        hidro2_re = hidro3_f.rename(columns={60: 'Norte',
                                             61: 'Nordeste',
                                             62: 'Sul',
                                             63: 'Sudeste',
                                             64: 'SIN'})
        hidro2_re1 = hidro2_re.rename(index={'Unnamed: 14': data})
        hidro2_v = hidro2_re1.iloc[1:2]

        hidro2_v.to_excel(r'C:\Users\e806128\Desktop\1\Dados_EnaBruta.xlsx', )

        hidro31 = ipdo.iloc[60:65, 10:24]
        colunassh1 = [1, 2, 3, 4]
        hidro4 = hidro31.drop(hidro31.columns[colunassh1], axis=1)
        hidro4_f = hidro4.transpose()
        hidro3_re = hidro4_f.rename(columns={60: 'Norte',
                                             61: 'Nordeste',
                                             62: 'Sul',
                                             63: 'Sudeste',
                                             64: 'SIN'})
        hidro3_re1 = hidro3_re.rename(index={'Unnamed: 15': data})
        hidro3_v = hidro3_re1.iloc[1:2]

        hidro3_v.to_excel(r'C:\Users\e806128\Desktop\1\Dados_EnaArmaz.xlsx')
        hidro4 = ipdo.iloc[60:65, 10:24]
        colunassh2 = [1, 2, 3, 4]
        hidro5 = hidro4.drop(hidro4.columns[colunassh2], axis=1)
        hidro4_f = hidro5.transpose()
        hidro3_re = hidro4_f.rename(columns={60: 'Norte',
                                             61: 'Nordeste',
                                             62: 'Sul',
                                             63: 'Sudeste',
                                             64: 'SIN'})
        hidro3_re1 = hidro3_re.rename(index={'Unnamed: 16': data})
        hidro3_v = hidro3_re1.iloc[2:3]

        hidro3_v.to_excel(r'C:\Users\e806128\Desktop\1\Dados_ValorEARD.xlsx')

        hidro5 = ipdo.iloc[60:65, 10:24]
        colunassh3 = [1, 2, 3, 4]
        hidro6 = hidro5.drop(hidro5.columns[colunassh3], axis=1)
        hidro5_f = hidro6.transpose()
        hidro4_re = hidro5_f.rename(columns={60: 'Norte',
                                             61: 'Nordeste',
                                             62: 'Sul',
                                             63: 'Sudeste',
                                             64: 'SIN'})
        hidro4_re1 = hidro4_re.rename(index={'Unnamed: 17': data})
        h4 = hidro4_re1.iloc[3:4]

        h4.to_excel(r'C:\Users\e806128\Desktop\1\Dados_ValrEarDiaP100.xlsx')
        hidro6 = ipdo.iloc[60:65, 10:24]
        colunassh4 = [1, 2, 3, 4]
        hidro7 = hidro6.drop(hidro6.columns[colunassh4], axis=1)
        hidro6_f = hidro7.transpose()
        hidro5_re = hidro6_f.rename(columns={60: 'Norte',
                                             61: 'Nordeste',
                                             62: 'Sul',
                                             63: 'Sudeste',
                                             64: 'SIN'})
        hidro5_re1 = hidro5_re.rename(index={'Unnamed: 19': data})
        hidro5_v = hidro5_re1.iloc[5:6]
        hidro5_v.to_excel(r'C:\Users\e806128\Desktop\1\Dados_DesvioEAR.xlsx')
        hidro7 = ipdo.iloc[60:65, 10:24]
        colunassh5 = [1, 2, 3, 4]
        hidro8 = hidro7.drop(hidro7.columns[colunassh5], axis=1)
        hidro7_f = hidro8.transpose()
        hidro6_re = hidro7_f.rename(columns={60: 'Norte',
                                             61: 'Nordeste',
                                             62: 'Sul',
                                             63: 'Sudeste',
                                             64: 'SIN'})
        hidro6_re1 = hidro6_re.rename(index={'Unnamed: 21': data})
        h6 = hidro6_re1.iloc[7:8]

        h6.to_excel(r'C:\Users\e806128\Desktop\1\Dados_VariacaoEmP100.xlsx')
        hidro8 = ipdo.iloc[60:65, 10:24]
        colunassh6 = [1, 2, 3, 4]
        hidro9 = hidro8.drop(hidro8.columns[colunassh6], axis=1)
        hidro8_f = hidro9.transpose()
        hidro7_re = hidro8_f.rename(columns={60: 'Norte',
                                             61: 'Nordeste',
                                             62: 'Sul',
                                             63: 'Sudeste',
                                             64: 'SIN'})
        hidro7_re1 = hidro7_re.rename(index={'Unnamed: 23': data})
        hidro7_v = hidro7_re1.iloc[9:10]
        hidro7_v.to_excel(r'C:\Users\e806128\Desktop\1\Dados_Variacao.xlsx', )

        hidro9 = ipdo.iloc[68:73, 10:24]
        hidro10 = hidro9.drop(hidro9.columns[1], axis=1)
        hidro10_d = hidro10.transpose()
        hidro10_re = hidro10_d.rename(columns={68: 'Norte',
                                               69: 'Nordeste',
                                               70: 'Sul',
                                               71: 'Sudeste',
                                               72: 'SIN'})
        hidro10_re1 = hidro10_re.rename(index={'Unnamed: 12': data})

        hidro10_f = hidro10_re1.iloc[1:2]
        hidro10_f.to_excel(r'C:\Users\e806128\Desktop\1\Dados_CapMax.xlsx')
        hidro10 = ipdo.iloc[68:73, 10:24]
        colunassv = [1, 2, 3]
        hidro11 = hidro10.drop(hidro10.columns[colunassv], axis=1)
        hidro11_f = hidro11.transpose()
        hidro11_re = hidro11_f.rename(columns={68: 'Norte',
                                               69: 'Nordeste',
                                               70: 'Sul',
                                               71: 'Sudeste',
                                               72: 'SIN'})
        hidro10_re1 = hidro11_re.rename(index={'Unnamed: 14': data})
        hidro10_v = hidro10_re1.iloc[1: 2]
        hidro10_v.to_excel(r'C:\Users\e806128\Desktop\1\Dados_VarMensal.xlsx')


def valoresMDUTT(ipdo: pd.DataFrame):
    """
   data,colunassh,transpose,to_excel, n_index.
   ----------
   ipdo : DataFrame da tabela.

   Retorna um arquivo .xlsx em formato de tabela a respeito dos valores da mé-
   dia diária das usinas termicas.
   -------
   """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\Valores_AngraVMDUT.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_AngraVMDUT.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_AngraIIVMDUT.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Val_N.FluminesVMDUT.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_MarlimAzul.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Val_BaiFluminense.xlsx')
        add_inf6 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Val_SantaCruzNova.xlsx')
        add_inf7 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Val_DoAtlântico.xlsx')
        add_inf8 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Val_LuizO.R.Melo.xlsx')
        add_inf9 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Val_GNA1.xlsx')
        add_inf10 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Termorio.xlsx')
        add_inf11 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Val_Cubatao.xlsx')
        add_inf12 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_ibirete.xlsx')
        add_inf13 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_TresLagoas.xlsx')
        add_inf14 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Karkey013.xlsx')
        add_inf15 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\'Valores_Karkey019.xlsx')
        add_inf16 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Val_NovaPiratininga.xlsx')
        add_inf17 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Seropédica.xlsx')
        add_inf18 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Termomacaé.xlsx')
        add_inf19 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_PorsudII.xlsx')
        add_inf20 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_PorsudI.xlsx')
        add_inf21 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Viana.xlsx')
        add_inf22 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Juiz_deFora.xlsx')
        add_inf23 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Povoacao1.xlsx')
        add_inf24 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Viana1.xlsx')
        add_inf25 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Val_PalmeirasdeGoias.xlsx')
        add_inf26 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Val_Goiania2.xlsx')
        add_inf27 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Val_Cuiabá.xlsx')
        add_inf28 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\'Valores_Daia.xlsx')
        add_inf29 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_W.Arjona.xlsx')
        add_inf30 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Valores_TotalSE.xlsx')

        no = no + 3
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0,  1]
        hidro1 = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro1_r = hidro1.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro1_re1 = hidro1_r.rename(index={259: data})
        vari1 = hidro1_re1.iloc[0: 1]

        dado1 = data
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           dado9)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(r'C:\Users\e806128\Desktop\1\Valores_AngraVMDUT.xlsx',
                         index=False)
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0,  1]
        hidro1_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_r = hidro1_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_r.rename(index={260: data})
        vari1 = hidro10_re1.iloc[1:2]

        dado1 = data
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_AngraIIVMDUT.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0,  1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={261: data})
        vari1 = hidro10_re1.iloc[2:3]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\Val_N.FluminesVMDUT.xlsx',
            index=False)
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0,  1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={262: data})
        vari1 = hidro10_re1.iloc[3:4]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_MarlimAzul.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0,  1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={263: data})
        vari1 = hidro10_re1.iloc[4:5]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\Val_BaiFluminense.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0,  1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={264: data})
        vari1 = hidro10_re1.iloc[5:6]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf6.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf6 = add_inf6.rename(columns={'Unnamed: 0': ' '})
        add_inf6.to_excel(
            r'C:\Users\e806128\Desktop\1\Val_SantaCruzNova.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0,  1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={265: data})
        vari1 = hidro10_re1.iloc[6:7]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf7.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf7 = add_inf7.rename(columns={'Unnamed: 0': ' '})
        add_inf7.to_excel(
            r'C:\Users\e806128\Desktop\1\Val_DoAtlântico.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0,  1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={266: data})
        vari1 = hidro10_re1.iloc[7:8]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf8.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf8 = add_inf8.rename(columns={'Unnamed: 0': ' '})
        add_inf8.to_excel(
            r'C:\Users\e806128\Desktop\1\Val_LuizO.R.Melo.xlsx',
            index=False)
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={267: data})
        vari1 = hidro10_re1.iloc[8:9]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf9.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf9 = add_inf9.rename(columns={'Unnamed: 0': ' '})
        add_inf9.to_excel(r'C:\Users\e806128\Desktop\1\Val_GNA1.xlsx',
                          index=False)
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={268: data})
        vari1 = hidro10_re1.iloc[9:10]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf10.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf10 = add_inf10.rename(columns={'Unnamed: 0': ' '})
        add_inf10.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Termorio.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={269: data})
        vari1 = hidro10_re1.iloc[10:11]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf11.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf11 = add_inf11.rename(columns={'Unnamed: 0': ' '})
        add_inf11.to_excel(
            r'C:\Users\e806128\Desktop\1\Val_Cubatao.xlsx', index=False)
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={270: data})
        vari1 = hidro10_re1.iloc[11:12]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf12.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf12 = add_inf12.rename(columns={'Unnamed: 0': ' '})
        add_inf12.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_ibirete.xlsx', index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={271: data})
        vari1 = hidro10_re1.iloc[12:13]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_TresLagoas.xlsx',
            index=False)

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf13.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf13 = add_inf13.rename(columns={'Unnamed: 0': ' '})
        add_inf13.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_TresLagoas.xlsx',
            index=False)
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={272: data})
        vari1 = hidro10_re1.iloc[13:14]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Karkey013.xlsx',
            index=False)
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf14.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf14 = add_inf14.rename(columns={'Unnamed: 0': ' '})
        add_inf14.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Karkey013.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={273: data})
        vari1 = hidro10_re1.iloc[14:15]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf15.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf15 = add_inf15.rename(columns={'Unnamed: 0': ' '})
        add_inf15.to_excel(
            r'C:\Users\e806128\Desktop\1\'Valores_Karkey019.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={274: data})
        vari1 = hidro10_re1.iloc[15:16]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf16.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf16 = add_inf16.rename(columns={'Unnamed: 0': ' '})
        add_inf16.to_excel(
            r'C:\Users\e806128\Desktop\1\Val_NovaPiratininga.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={275: data})
        vari1 = hidro10_re1.iloc[16:17]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf17.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf17 = add_inf17.rename(columns={'Unnamed: 0': ' '})
        add_inf17.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Seropédica.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={276: data})
        vari1 = hidro10_re1.iloc[17:18]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf18.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf18 = add_inf18.rename(columns={'Unnamed: 0': ' '})
        add_inf18.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Termomacaé.xlsx',
            index=False)
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={277: data})
        vari1 = hidro10_re1.iloc[18:19]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf19.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf19 = add_inf19.rename(columns={'Unnamed: 0': ' '})
        add_inf19.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_PorsudII.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={278: data})
        vari1 = hidro10_re1.iloc[19:20]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_PorsudI.xlsx',
            index=False)
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf20.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf20 = add_inf20.rename(columns={'Unnamed: 0': ' '})
        add_inf20.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_PorsudI.xlsx',
            index=False)
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={279: data})
        vari1 = hidro10_re1.iloc[20:21]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf21.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf21 = add_inf21.rename(columns={'Unnamed: 0': ' '})
        add_inf21.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Viana.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={280: data})
        vari1 = hidro10_re1.iloc[21:22]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf22.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf22 = add_inf22.rename(columns={'Unnamed: 0': ' '})
        add_inf22.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Juiz_deFora.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={281: data})
        vari1 = hidro10_re1.iloc[22:23]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf23.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf23 = add_inf23.rename(columns={'Unnamed: 0': ' '})
        add_inf23.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Povoacao1.xlsx',
            index=False)
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={282: data})
        vari1 = hidro10_re1.iloc[23:24]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf24.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf24 = add_inf24.rename(columns={'Unnamed: 0': ' '})
        add_inf24.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_Viana1.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={283: data})
        vari1 = hidro10_re1.iloc[24:25]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf25.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf25 = add_inf25.rename(columns={'Unnamed: 0': ' '})
        add_inf25.to_excel(
            r'C:\Users\e806128\Desktop\1\Val_PalmeirasdeGoias.xlsx',
            index=False)
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={284: data})
        vari1 = hidro10_re1.iloc[25:26]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf26.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf26 = add_inf26.rename(columns={'Unnamed: 0': ' '})
        add_inf26.to_excel(
            r'C:\Users\e806128\Desktop\1\Val_Goiania2.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={285: data})
        vari1 = hidro10_re1.iloc[26:27]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf27.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf27 = add_inf27.rename(columns={'Unnamed: 0': ' '})
        add_inf27.to_excel(
            r'C:\Users\e806128\Desktop\1\Val_Cuiabá.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={286: data})
        vari1 = hidro10_re1.iloc[27:28]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]

        add_inf28.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf28 = add_inf28.rename(columns={'Unnamed: 0': ' '})
        add_inf28.to_excel(
            r'C:\Users\e806128\Desktop\1\'Valores_Daia.xlsx',
            index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={287: data})
        vari1 = hidro10_re1.iloc[28:29]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf29.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf29 = add_inf29.rename(columns={'Unnamed: 0': ' '})
        add_inf29.to_excel(r'C:\Users\e806128\Desktop\1\Valores_W.Arjona.xlsx',
                           index=False)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={291: data})
        vari1 = hidro10_re1.iloc[32:33]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf30.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf30 = add_inf30.rename(columns={'Unnamed: 0': ' '})
        add_inf30.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_TotalSE.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0,  1]
        hidro1 = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro1_r = hidro1.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro1_re1 = hidro1_r.rename(index={259: data})
        hidro10_v = hidro1_re1.iloc[0: 1]
        hidro10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_AngraVMDUT.xlsx')

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0,  1]
        hidro1_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_r = hidro1_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_r.rename(index={260: data})
        vari1 = hidro10_re1.iloc[1:2]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_AngraIIVMDUT.xlsx')

        hidro9 = ipdo.iloc[259: 292, 0: 10]
        colunassh = [0,  1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={261: data})
        hidro10_v = hidro10_re1.iloc[2:3]

        hidro10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\Val_N.FluminesVMDUT.xlsx',)

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0,  1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={262: data})
        vari1 = hidro10_re1.iloc[3:4]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\Valores_MarlimAzul.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0,  1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={263: data})
        vari1 = hidro10_re1.iloc[4:5]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\Val_BaiFluminense.xlsx')

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0,  1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={264: data})
        vari1 = hidro10_re1.iloc[5:6]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\Val_SantaCruzNova.xlsx')

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0,  1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={265: data})
        vari1 = hidro10_re1.iloc[6:7]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\Val_DoAtlântico.xlsx')

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0,  1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={266: data})
        vari1 = hidro10_re1.iloc[7:8]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Val_LuizO.R.Melo.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={267: data})
        vari1 = hidro10_re1.iloc[8:9]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Val_GNA1.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={268: data})
        vari1 = hidro10_re1.iloc[9:10]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Valores_Termorio.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={269: data})
        vari1 = hidro10_re1.iloc[10:11]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Val_Cubatao.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={270: data})
        vari1 = hidro10_re1.iloc[11:12]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Valores_ibirete.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={271: data})
        vari1 = hidro10_re1.iloc[12:13]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Valores_TresLagoas.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={272: data})
        vari1 = hidro10_re1.iloc[13:14]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Valores_Karkey013.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={273: data})
        vari1 = hidro10_re1.iloc[14:15]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\'Valores_Karkey019.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={274: data})
        vari1 = hidro10_re1.iloc[15:16]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Val_NovaPiratininga.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={275: data})
        vari1 = hidro10_re1.iloc[16:17]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Valores_Seropédica.xlsx')

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={276: data})
        vari1 = hidro10_re1.iloc[17:18]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Valores_Termomacaé.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={277: data})
        vari1 = hidro10_re1.iloc[18:19]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Valores_PorsudII.xlsx')

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={278: data})
        vari1 = hidro10_re1.iloc[19:20]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Valores_PorsudI.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={279: data})
        vari1 = hidro10_re1.iloc[20:21]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Valores_Viana.xlsx')

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={280: data})
        vari1 = hidro10_re1.iloc[21:22]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Valores_Juiz_deFora.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={281: data})
        vari1 = hidro10_re1.iloc[22:23]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Valores_Povoacao1.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={282: data})
        vari1 = hidro10_re1.iloc[23:24]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Valores_Viana1.xlsx')

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={283: data})
        vari1 = hidro10_re1.iloc[24:25]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Val_PalmeirasdeGoias.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={284: data})
        vari1 = hidro10_re1.iloc[25:26]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Val_Goiania2.xlsx')

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={285: data})
        vari1 = hidro10_re1.iloc[26:27]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Val_Cuiabá.xlsx')

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={286: data})
        vari1 = hidro10_re1.iloc[27:28]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\'Valores_Daia.xlsx')
        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)
        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={287: data})
        vari1 = hidro10_re1.iloc[28:29]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Valores_W.Arjona.xlsx')

        hidro9 = ipdo.iloc[259:292, 0:10]
        colunassh = [0, 1]
        hidro10_d = hidro9.drop(hidro9.columns[colunassh], axis=1)

        hidro10_re = hidro10_d.rename(columns={
            'Unnamed: 2': 'Razão Despacho',
            'Unnamed: 3': 'Capacidade Instalada',
            'Unnamed: 4': 'Capacidade Disponível',
            'Unnamed: 5': 'Media Diária programada',
            'Unnamed: 6': 'Media Diária verificada',
            'Unnamed: 7': 'Média diária Difer',
            'Unnamed: 8': 'Variação %',
            'Unnamed: 9': 'OBS'})
        hidro10_re1 = hidro10_re.rename(index={291: data})
        vari1 = hidro10_re1.iloc[32:33]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\Valores_TotalSE.xlsx')


def ValMDUTTNorte(ipdo: pd.DataFrame):
    """
   data,colunassh,transpose,to_excel, n_index.
   ----------
   ipdo : DataFrame da tabela.

   Retorna um arquivo .xlsx em formato de tabela a respeito dos valores da mé-
   dia diária das usinas termicas do norte.
   -------
   """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\ValMedNorte_Ap.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Ap.xlsx')
        add_inf1 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Mauá3.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_MarIII.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_ParIV.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_MarIV.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_MarV.xlsx')
        add_inf6 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_PaV.xlsx')
        add_inf7 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_NVenecia.xlsx')
        add_inf8 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_PdoI.xlsx')
        add_inf9 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_GeI.xlsx')
        add_inf10 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_GeraII.xlsx')
        add_inf11 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_CrisRo1.xlsx')
        add_inf12 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Jaraqui.xlsx')
        add_inf13 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Manauara.xlsx')
        add_inf14 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_PontaNegra.xlsx')
        add_inf15 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Tambaqui.xlsx')
        add_inf16 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ ValMedNorte_Total.xlsx')

        no = no + 3
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={377: data})
        vari1 = vari1_re1.iloc[2:3]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           dado9)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Ap.xlsx',
            index=False)

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={378: data})
        vari1 = vari1_re1.iloc[3:4]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf1.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf1 = add_inf1.rename(columns={'Unnamed: 0': ' '})
        add_inf1.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Mauá3.xlsx',
            index=False)
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={379: data})
        vari1 = vari1_re1.iloc[4:5]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMedNorte_MarIII.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_MarIII.xlsx',
            index=False)
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={380: data})
        vari1 = vari1_re1.iloc[5:6]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMedNorte_ParIV.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_ParIV.xlsx',
            index=False)
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={381: data})
        vari1 = vari1_re1.iloc[6:7]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMedNorte_MarIV.xlsx', )

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_MarIV.xlsx',
            index=False)

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={382: data})
        vari1 = vari1_re1.iloc[7:8]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMedNorte_MarV.xlsx')

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_MarV.xlsx',
            index=False)

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={383: data})
        vari1 = vari1_re1.iloc[8:9]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMedNorte_PaV.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf6.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf6 = add_inf6.rename(columns={'Unnamed: 0': ' '})
        add_inf6.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_PaV.xlsx',
            index=False)

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={384: data})
        vari1 = vari1_re1.iloc[9:10]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMedNorte_NVenecia.xlsx')

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf7.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf7 = add_inf7.rename(columns={'Unnamed: 0': ' '})
        add_inf7.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_NVenecia.xlsx',
            index=False)

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={385: data})
        vari1 = vari1_re1.iloc[10:11]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMedNorte_PdoI.xlsx', )
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf8.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf8 = add_inf8.rename(columns={'Unnamed: 0': ' '})
        add_inf8.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_PdoI.xlsx',
            index=False)

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                  'Unnamed: 2': 'Razão Despacho',
                                  'Unnamed: 3': 'Capacidade Instalada',
                                  'Unnamed: 4': 'Capacidade Disponível',
                                  'Unnamed: 5': 'Media Diária Programada',
                                  'Unnamed: 6': 'Media Diária verificada ',
                                  'Unnamed: 7': 'Média diária Difer',
                                  'Unnamed: 8': 'Variação%',
                                  'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={386: data})
        vari1 = vari1_re1.iloc[11:12]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMedNorte_GeI.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf9.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf9 = add_inf9.rename(columns={'Unnamed: 0': ' '})
        add_inf9.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_GeI.xlsx',
            index=False)
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                  'Unnamed: 2': 'Razão Despacho',
                                  'Unnamed: 3': 'Capacidade Instalada',
                                  'Unnamed: 4': 'Capacidade Disponível',
                                  'Unnamed: 5': 'Media Diária Programada',
                                  'Unnamed: 6': 'Media Diária verificada ',
                                  'Unnamed: 7': 'Média diária Difer',
                                  'Unnamed: 8': 'Variação%',
                                  'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={387: data})
        vari1 = vari1_re1.iloc[12:13]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMedNorte_GeraII.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf10.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf10 = add_inf10.rename(columns={'Unnamed: 0': ' '})
        add_inf10.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_GeraII.xlsx',
            index=False)
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                  'Unnamed: 2': 'Razão Despacho',
                                  'Unnamed: 3': 'Capacidade Instalada',
                                  'Unnamed: 4': 'Capacidade Disponível',
                                  'Unnamed: 5': 'Media Diária Programada',
                                  'Unnamed: 6': 'Media Diária verificada ',
                                  'Unnamed: 7': 'Média diária Difer',
                                  'Unnamed: 8': 'Variação%',
                                  'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={388: data})
        vari1 = vari1_re1.iloc[13:14]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMedNorte_CrisRo1.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf11.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf11 = add_inf11.rename(columns={'Unnamed: 0': ' '})
        add_inf11.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_CrisRo1.xlsx',
            index=False)

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                  'Unnamed: 2': 'Razão Despacho',
                                  'Unnamed: 3': 'Capacidade Instalada',
                                  'Unnamed: 4': 'Capacidade Disponível',
                                  'Unnamed: 5': 'Media Diária Programada',
                                  'Unnamed: 6': 'Media Diária verificada ',
                                  'Unnamed: 7': 'Média diária Difer',
                                  'Unnamed: 8': 'Variação%',
                                  'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={389: data})
        vari1 = vari1_re1.iloc[14:15]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMedNorte_Jaraqui.xlsx')

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf12.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf12 = add_inf12.rename(columns={'Unnamed: 0': ' '})
        add_inf12.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Jaraqui.xlsx',
            index=False)

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                  'Unnamed: 2': 'Razão Despacho',
                                  'Unnamed: 3': 'Capacidade Instalada',
                                  'Unnamed: 4': 'Capacidade Disponível',
                                  'Unnamed: 5': 'Media Diária Programada',
                                  'Unnamed: 6': 'Media Diária verificada ',
                                  'Unnamed: 7': 'Média diária Difer',
                                  'Unnamed: 8': 'Variação%',
                                  'Unnamed: 9': 'OBS'})
        vari1 = vari1_re.rename(index={390: data})
        vari1 = vari1_re1.iloc[15:16]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Manauara.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf13.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf13 = add_inf13.rename(columns={'Unnamed: 0': ' '})
        add_inf13.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Manauara.xlsx',
            index=False)
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                  'Unnamed: 2': 'Razão Despacho',
                                  'Unnamed: 3': 'Capacidade Instalada',
                                  'Unnamed: 4': 'Capacidade Disponível',
                                  'Unnamed: 5': 'Media Diária Programada',
                                  'Unnamed: 6': 'Media Diária verificada ',
                                  'Unnamed: 7': 'Média diária Difer',
                                  'Unnamed: 8': 'Variação%',
                                  'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={391: data})
        vari1 = vari1_re1.iloc[16:17]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf14.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf14 = add_inf14.rename(columns={'Unnamed: 0': ' '})
        add_inf14.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_PontaNegra.xlsx',
            index=False)

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                  'Unnamed: 2': 'Razão Despacho',
                                  'Unnamed: 3': 'Capacidade Instalada',
                                  'Unnamed: 4': 'Capacidade Disponível',
                                  'Unnamed: 5': 'Media Diária Programada',
                                  'Unnamed: 6': 'Media Diária verificada ',
                                  'Unnamed: 7': 'Média diária Difer',
                                  'Unnamed: 8': 'Variação%',
                                  'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={392: data})
        vari1 = vari1_re1.iloc[17:18]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMedNorte_Tambaqui.xlsx')

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf15.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf15 = add_inf15.rename(columns={'Unnamed: 0': ' '})
        add_inf15.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Tambaqui.xlsx',
            index=False)

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                  'Unnamed: 2': 'Razão Despacho',
                                  'Unnamed: 3': 'Capacidade Instalada',
                                  'Unnamed: 4': 'Capacidade Disponível',
                                  'Unnamed: 5': 'Media Diária Programada',
                                  'Unnamed: 6': 'Media Diária verificada ',
                                  'Unnamed: 7': 'Média diária Difer',
                                  'Unnamed: 8': 'Variação%',
                                  'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={405: data})
        vari1 = vari1_re1.iloc[30:31]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ ValMedNorte_Total.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf16.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf16 = add_inf16.rename(columns={'Unnamed: 0': ' '})
        add_inf16.to_excel(
            r'C:\Users\e806128\Desktop\1\ ValMedNorte_Total.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={377: data})
        vari1 = vari1_re1.iloc[2:3]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Ap.xlsx')
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={378: data})
        vari1_v = vari1_re1.iloc[3:4]

        vari1_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Mauá3.xlsx', )

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={379: data})
        vari1_v = vari1_re1.iloc[4:5]

        vari1_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_MarIII.xlsx')
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={380: data})
        vari1_v = vari1_re1.iloc[5:6]

        vari1_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_ParIV.xlsx')
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={381: data})
        vari1_v = vari1_re1.iloc[6:7]

        vari1_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_MarIV.xlsx', )

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={382: data})
        vari1_v = vari1_re1.iloc[7:8]

        vari1_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_MarV.xlsx')
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={383: data})
        vari1_v = vari1_re1.iloc[8:9]

        vari1_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_PaV.xlsx')

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={384: data})
        vari1_v = vari1_re1.iloc[9:10]

        vari1_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_NVenecia.xlsx')

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={385: data})
        vari1_v = vari1_re1.iloc[10:11]

        vari1_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_PdoI.xlsx', )
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={386: data})
        vari1_v = vari1_re1.iloc[11:12]

        vari1_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_GeI.xlsx')
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={387: data})
        vari1_v = vari1_re1.iloc[12:13]

        vari1_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_GeraII.xlsx')
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)
        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={388: data})
        vari1_v = vari1_re1.iloc[13:14]

        vari1_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_CrisRo1.xlsx')

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={389: data})
        vari1_v = vari1_re1.iloc[14:15]

        vari1_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Jaraqui.xlsx')
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1 = vari1_re.rename(index={390: data})
        vari1 = vari1_re1.iloc[15:16]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Manauara.xlsx')
        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={391: data})
        vari1 = vari1_re1.iloc[16:17]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_PontaNegra.xlsx')

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={392: data})
        vari1_v = vari1_re1.iloc[17:18]

        vari1_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMedNorte_Tambaqui.xlsx')

        valor9 = ipdo.iloc[375:410, 0:10]
        colunassh = [0, 1]
        vari1_d = valor9.drop(valor9.columns[colunassh], axis=1)

        vari1_re = vari1_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação%',
                                'Unnamed: 9': 'OBS'})
        vari1_re1 = vari1_re.rename(index={405: data})
        vari1_v = vari1_re1.iloc[30:31]

        vari1_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ ValMedNorte_Total.xlsx')


def valoresMDUTTSUL(ipdo: pd.DataFrame):
    """
   data,colunassh,transpose,to_excel, n_index.
   ----------
   ipdo : DataFrame da tabela.

   Retorna um arquivo .xlsx em formato de tabela a respeito dos valores da mé-
   dia diária das usinas termicas do sul.
   -------
   """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\ValMeSul_Pampa.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
               r'C:\Users\e806128\Desktop\1\ValMeSul_Pampa.xlsx')
        add_inf2 = pd.read_excel(
               r'C:\Users\e806128\Desktop\1\ValMeSulCandiotaII_.xlsx')
        add_inf3 = pd.read_excel(
               r'C:\Users\e806128\Desktop\1\ValMeSul_J.LacerdaC.xlsx')
        add_inf4 = pd.read_excel(
              r'C:\Users\e806128\Desktop\1\ValMeSul_Figueira.xlsx')
        add_inf5 = pd.read_excel(
               r'C:\Users\e806128\Desktop\1\ValMeSul_JLacerB.xlsx')
        add_inf6 = pd.read_excel(
               r'C:\Users\e806128\Desktop\1\ValMeSul_JLacerA.xlsx')
        add_inf7 = pd.read_excel(
               r'C:\Users\e806128\Desktop\1\ValMeSul_Canoas.xlsx')
        add_inf8 = pd.read_excel(
               r'C:\Users\e806128\Desktop\1\ValMeSul_Araucária.xlsx')
        add_inf9 = pd.read_excel(
               r'C:\Users\e806128\Desktop\1\ValMeSul_Uruguaiana.xlsx')
        add_inf10 = pd.read_excel(
               r'C:\Users\e806128\Desktop\1\ValMeSul_Total.xlsx')

        no = no + 3
        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={315: data})
        vari1 = valor10_re1.iloc[0:1]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMeSul_Pampa.xlsx', )
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           dado9)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_Pampa.xlsx',
            index=False)
        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={316: data})
        vari1 = valor10_re1.iloc[1:2]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMeSulCandiotaII_.xlsx',)

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSulCandiotaII_.xlsx',
            index=False)

        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={317: data})
        vari1 = valor10_re1.iloc[2:3]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMeSul_J.LacerdaC.xlsx',)

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_J.LacerdaC.xlsx',
            index=False)
        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={318: data})
        vari1 = valor10_re1.iloc[3:4]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMeSul_Figueira.xlsx', )
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_Figueira.xlsx',
            index=False)

        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={319: data})
        vari1 = valor10_re1.iloc[4:5]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMeSul_JLacerB.xlsx', )
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_JLacerB.xlsx',
            index=False)
        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={320: data})
        vari1 = valor10_re1.iloc[5:6]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMeSul_JLacerA.xlsx', )
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf6.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf6 = add_inf6.rename(columns={'Unnamed: 0': ' '})
        add_inf6.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_JLacerA.xlsx',
            index=False)

        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={321: data})
        vari1 = valor10_re1.iloc[6:7]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMeSul_Canoas.xlsx', )
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf7.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf7 = add_inf7.rename(columns={'Unnamed: 0': ' '})
        add_inf7.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_Canoas.xlsx',
            index=False)

        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={322: data})
        vari1 = valor10_re1.iloc[7:8]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMeSul_Araucária.xlsx', )
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf8.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf8 = add_inf8.rename(columns={'Unnamed: 0': ' '})
        add_inf8.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_Araucária.xlsx',
            index=False)

        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={323: data})
        vari1 = valor10_re1.iloc[8:9]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMeSul_Uruguaiana.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf9.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf9 = add_inf9.rename(columns={'Unnamed: 0': ' '})
        add_inf9.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_Uruguaiana.xlsx',
            index=False)

        valor9 = ipdo.iloc[315:334, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={330: data})
        vari1 = valor10_re1.iloc[15:16]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMeSul_Total.xlsx', )
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf10.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf10 = add_inf10.rename(columns={'Unnamed: 0': ' '})
        add_inf10.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_Total.xlsx',
            index=False)
    else:
        print('O arquivo não existe.')
        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={315: data})
        valor10_v = valor10_re1.iloc[0:1]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_Pampa.xlsx', )

        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={316: data})
        valor10_v = valor10_re1.iloc[1:2]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSulCandiotaII_.xlsx',)

        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={317: data})
        valor10_v = valor10_re1.iloc[2:3]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_J.LacerdaC.xlsx',)

        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={318: data})
        valor10_v = valor10_re1.iloc[3:4]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_Figueira.xlsx', )

        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={319: data})
        valor10_v = valor10_re1.iloc[4:5]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_JLacerB.xlsx', )

        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={320: data})
        valor10_v = valor10_re1.iloc[5:6]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_JLacerA.xlsx', )

        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={321: data})
        valor10_v = valor10_re1.iloc[6:7]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_Canoas.xlsx', )

        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={322: data})
        valor10_v = valor10_re1.iloc[7:8]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_Araucária.xlsx', )

        valor9 = ipdo.iloc[315:324, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={323: data})
        valor10_v = valor10_re1.iloc[8:9]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_Uruguaiana.xlsx',)

        valor9 = ipdo.iloc[315:334, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={330: data})
        valor10_v = valor10_re1.iloc[15:16]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeSul_Total.xlsx', )


def valoresMDTTNordeste(ipdo: pd.DataFrame):
    """
     data,colunassh,transpose,to_excel, n_index.
   ----------
   ipdo : DataFrame da tabela.

   Retorna um arquivo .xlsx em formato de tabela a respeito dos valores da mé-
   dia diária das usinas termicas do nordeste.
   -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\ValMeNor_Term.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_Term.xlsx')
        add_inf2 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_PorPecémI.xlsx')
        add_inf3 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_PorPecémII.xlsx')
        add_inf4 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_PortoSe.xlsx')
        add_inf5 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_ValedoAcu.xlsx')
        add_inf6 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_Termoceará.xlsx')
        add_inf7 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_TermoBahia.xlsx')
        add_inf8 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_PerloIII.xlsx')
        add_inf9 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_MracanaúI.xlsx')
        add_inf10 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_Termocabo.xlsx')
        add_inf11 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_Termonorde.xlsx')
        add_inf12 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_Termona.xlsx')
        add_inf13 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_CampinaG.xlsx')
        add_inf14 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_SuapeII.xlsx')
        add_inf15 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_GlobalI.xlsx')
        add_inf16 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_GlobalII.xlsx')
        add_inf17 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_Fortelaza.xlsx')
        add_inf18 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_Apoena.xlsx')
        add_inf19 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_guarani.xlsx')
        add_inf20 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_pretolina.xlsx')
        add_inf21 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_portiguar.xlsx')
        add_inf22 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_paudeferI.xlsx')
        add_inf23 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_termom.xlsx')
        add_inf24 = pd.read_excel(
                  r'C:\Users\e806128\Desktop\1\ValMeNor_Total.xlsx')

        no = no + 3
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={334: data})
        vari1 = valor10_re1.iloc[0:1]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           dado9)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_Term.xlsx',
            index=False)

        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={335: data})
        vari1 = valor10_re1.iloc[1:2]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_PorPecémI.xlsx',
            index=False)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={336: data})
        vari1 = valor10_re1.iloc[2:3]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_PorPecémII.xlsx',
            index=False)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={337: data})
        vari1 = valor10_re1.iloc[3:4]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_PortoSe.xlsx',
            index=False)

        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={338: data})
        vari1 = valor10_re1.iloc[4:5]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_ValedoAcu.xlsx',
            index=False)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={339: data})
        vari1 = valor10_re1.iloc[5:6]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf6.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf6 = add_inf6.rename(columns={'Unnamed: 0': ' '})
        add_inf6.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_Termoceará.xlsx',
            index=False)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={340: data})
        vari1 = valor10_re1.iloc[6:7]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf7.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf7 = add_inf7.rename(columns={'Unnamed: 0': ' '})
        add_inf7.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_TermoBahia.xlsx',
            index=False)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={341: data})
        vari1 = valor10_re1.iloc[7:8]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf8.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf8 = add_inf8.rename(columns={'Unnamed: 0': ' '})
        add_inf8.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_PerloIII.xlsx',
            index=False)

        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={342: data})
        vari1 = valor10_re1.iloc[8:9]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf9.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf9 = add_inf9.rename(columns={'Unnamed: 0': ' '})
        add_inf9.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_MracanaúI.xlsx',
            index=False)

        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={343: data})
        vari1 = valor10_re1.iloc[9:10]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf10.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf10 = add_inf10.rename(columns={'Unnamed: 0': ' '})
        add_inf10.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_Termocabo.xlsx',
            index=False)

        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={344: data})
        vari1 = valor10_re1.iloc[10:11]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf11.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf11 = add_inf11.rename(columns={'Unnamed: 0': ' '})
        add_inf11.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_Termonorde.xlsx',
            index=False)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={345: data})
        vari1 = valor10_re1.iloc[11:12]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf12.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf12 = add_inf12.rename(columns={'Unnamed: 0': ' '})
        add_inf12.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_Termona.xlsx',
            index=False)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={346: data})
        vari1 = valor10_re1.iloc[12:13]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf13.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf13 = add_inf13.rename(columns={'Unnamed: 0': ' '})
        add_inf13.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_CampinaG.xlsx',
            index=False)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={347: data})
        vari1 = valor10_re1.iloc[13:14]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf14.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf14 = add_inf14.rename(columns={'Unnamed: 0': ' '})
        add_inf14.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_SuapeII.xlsx',
            index=False)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={348: data})
        vari1 = valor10_re1.iloc[14:15]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf15.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf15 = add_inf15.rename(columns={'Unnamed: 0': ' '})
        add_inf15.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_GlobalI.xlsx',
            index=False)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={349: data})
        vari1 = valor10_re1.iloc[15:16]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf16.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf16 = add_inf16.rename(columns={'Unnamed: 0': ' '})
        add_inf16.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_GlobalII.xlsx',
            index=False)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={350: data})
        vari1 = valor10_re1.iloc[16:17]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf17.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf17 = add_inf17.rename(columns={'Unnamed: 0': ' '})
        add_inf17.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_Fortelaza.xlsx',
            index=False)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={351: data})
        vari1 = valor10_re1.iloc[17:18]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf18.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf18 = add_inf18.rename(columns={'Unnamed: 0': ' '})
        add_inf18.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_Apoena.xlsx',
            index=False)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={352: data})
        vari1 = valor10_re1.iloc[18:19]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf19.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf19 = add_inf19.rename(columns={'Unnamed: 0': ' '})
        add_inf19.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_guarani.xlsx',
            index=False)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                   'Unnamed: 2': 'Razão Despacho',
                                   'Unnamed: 3': 'Capacidade Instalada',
                                   'Unnamed: 4': 'Capacidade Disponível',
                                   'Unnamed: 5': 'Media Diária programado',
                                   'Unnamed: 6': 'Media Diária verificada ',
                                   'Unnamed: 7': 'Média diária Difer',
                                   'Unnamed: 8': 'Variação %',
                                   'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={353: data})
        vari1 = valor10_re1.iloc[19:20]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf20.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf20 = add_inf20.rename(columns={'Unnamed: 0': ' '})
        add_inf20.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_pretolina.xlsx',
            index=False)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                   'Unnamed: 2': 'Razão Despacho',
                                   'Unnamed: 3': 'Capacidade Instalada',
                                   'Unnamed: 4': 'Capacidade Disponível',
                                   'Unnamed: 5': 'Media Diária programado',
                                   'Unnamed: 6': 'Media Diária verificada ',
                                   'Unnamed: 7': 'Média diária Difer',
                                   'Unnamed: 8': 'Variação %',
                                   'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={354: data})
        vari1 = valor10_re1.iloc[20:21]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMeNor_portiguar.xlsx', )
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf21.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf21 = add_inf21.rename(columns={'Unnamed: 0': ' '})
        add_inf21.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_portiguar.xlsx',
            index=False)

        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                   'Unnamed: 2': 'Razão Despacho',
                                   'Unnamed: 3': 'Capacidade Instalada',
                                   'Unnamed: 4': 'Capacidade Disponível',
                                   'Unnamed: 5': 'Media Diária programado',
                                   'Unnamed: 6': 'Media Diária verificada ',
                                   'Unnamed: 7': 'Média diária Difer',
                                   'Unnamed: 8': 'Variação %',
                                   'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={356: data})
        vari1 = valor10_re1.iloc[21:22]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMeNor_paudeferI.xlsx',)
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf22.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf22 = add_inf22.rename(columns={'Unnamed: 0': ' '})
        add_inf22.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_paudeferI.xlsx',
            index=False)

        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                  'Unnamed: 2': 'Razão Despacho',
                                  'Unnamed: 3': 'Capacidade Instalada',
                                  'Unnamed: 4': 'Capacidade Disponível',
                                  'Unnamed: 5': 'Media Diária programado',
                                  'Unnamed: 6': 'Media Diária verificada ',
                                  'Unnamed: 7': 'Média diária Difer',
                                  'Unnamed: 8': 'Variação %',
                                  'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={357: data})
        vari1 = valor10_re1.iloc[22:23]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMeNor_termom.xlsx',)
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf23.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf23 = add_inf23.rename(columns={'Unnamed: 0': ' '})
        add_inf23.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_termom.xlsx',
            index=False)

        valor9 = ipdo.iloc[334:370, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                  'Unnamed: 2': 'Razão Despacho',
                                  'Unnamed: 3': 'Capacidade Instalada',
                                  'Unnamed: 4': 'Capacidade Disponível',
                                  'Unnamed: 5': 'Media Diária programado',
                                  'Unnamed: 6': 'Media Diária verificada ',
                                  'Unnamed: 7': 'Média diária Difer',
                                  'Unnamed: 8': 'Variação %',
                                  'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={366: data})
        vari1 = valor10_re1.iloc[32:33]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ValMeNor_Total.xlsx',)
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf24.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf24 = add_inf24.rename(columns={'Unnamed: 0': ' '})
        add_inf24.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_Total.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={334: data})
        valor10_v = valor10_re1.iloc[0:1]
        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\ValMeNor_Term.xlsx', )

        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={335: data})
        valor10_v = valor10_re1.iloc[1:2]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_PorPecémI.xlsx',)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={336: data})
        valor10_v = valor10_re1.iloc[2:3]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_PorPecémII.xlsx',)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={337: data})
        valor10_v = valor10_re1.iloc[3:4]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_PortoSe.xlsx',)
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={338: data})
        valor10_v = valor10_re1.iloc[4:5]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_ValedoAcu.xlsx', )
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={339: data})
        valor10_v = valor10_re1.iloc[5:6]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_Termoceará.xlsx')
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={340: data})
        valor10_v = valor10_re1.iloc[6:7]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_TermoBahia.xlsx')
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={341: data})
        valor10_v = valor10_re1.iloc[7:8]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_PerloIII.xlsx')

        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={342: data})
        valor10_v = valor10_re1.iloc[8:9]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_MracanaúI.xlsx',)

        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={343: data})
        valor10_v = valor10_re1.iloc[9:10]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_Termocabo.xlsx')

        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={344: data})
        valor10_v = valor10_re1.iloc[10:11]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_Termonorde.xlsx')
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={345: data})
        valor10_v = valor10_re1.iloc[11:12]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_Termona.xlsx')
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={346: data})
        valor10_v = valor10_re1.iloc[12:13]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_CampinaG.xlsx')
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={347: data})
        valor10_v = valor10_re1.iloc[13:14]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_SuapeII.xlsx', )
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={348: data})
        valor10_v = valor10_re1.iloc[14:15]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_GlobalI.xlsx', )
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={349: data})
        valor10_v = valor10_re1.iloc[15:16]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_GlobalII.xlsx', )
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={350: data})
        valor10_v = valor10_re1.iloc[16:17]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_Fortelaza.xlsx', )
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={351: data})
        valor10_v = valor10_re1.iloc[17:18]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_Apoena.xlsx', )
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programado',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={352: data})
        valor10_v = valor10_re1.iloc[18:19]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_guarani.xlsx', )
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                   'Unnamed: 2': 'Razão Despacho',
                                   'Unnamed: 3': 'Capacidade Instalada',
                                   'Unnamed: 4': 'Capacidade Disponível',
                                   'Unnamed: 5': 'Media Diária programado',
                                   'Unnamed: 6': 'Media Diária verificada ',
                                   'Unnamed: 7': 'Média diária Difer',
                                   'Unnamed: 8': 'Variação %',
                                   'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={353: data})
        valor10_v = valor10_re1.iloc[19:20]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_pretolina.xlsx', )
        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                   'Unnamed: 2': 'Razão Despacho',
                                   'Unnamed: 3': 'Capacidade Instalada',
                                   'Unnamed: 4': 'Capacidade Disponível',
                                   'Unnamed: 5': 'Media Diária programado',
                                   'Unnamed: 6': 'Media Diária verificada ',
                                   'Unnamed: 7': 'Média diária Difer',
                                   'Unnamed: 8': 'Variação %',
                                   'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={354: data})
        valor10_v = valor10_re1.iloc[20:21]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_portiguar.xlsx', )

        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                   'Unnamed: 2': 'Razão Despacho',
                                   'Unnamed: 3': 'Capacidade Instalada',
                                   'Unnamed: 4': 'Capacidade Disponível',
                                   'Unnamed: 5': 'Media Diária programado',
                                   'Unnamed: 6': 'Media Diária verificada ',
                                   'Unnamed: 7': 'Média diária Difer',
                                   'Unnamed: 8': 'Variação %',
                                   'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={355: data})
        valor10_v = valor10_re1.iloc[21:22]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_paudeferI.xlsx',)

        valor9 = ipdo.iloc[334:367, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                  'Unnamed: 2': 'Razão Despacho',
                                  'Unnamed: 3': 'Capacidade Instalada',
                                  'Unnamed: 4': 'Capacidade Disponível',
                                  'Unnamed: 5': 'Media Diária programado',
                                  'Unnamed: 6': 'Media Diária verificada ',
                                  'Unnamed: 7': 'Média diária Difer',
                                  'Unnamed: 8': 'Variação %',
                                  'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={356: data})
        valor10_v = valor10_re1.iloc[22:23]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_termom.xlsx',)

        valor9 = ipdo.iloc[334:370, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                  'Unnamed: 2': 'Razão Despacho',
                                  'Unnamed: 3': 'Capacidade Instalada',
                                  'Unnamed: 4': 'Capacidade Disponível',
                                  'Unnamed: 5': 'Media Diária programado',
                                  'Unnamed: 6': 'Media Diária verificada ',
                                  'Unnamed: 7': 'Média diária Difer',
                                  'Unnamed: 8': 'Variação %',
                                  'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={366: data})
        valor10_v = valor10_re1.iloc[32:33]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeNor_Total.xlsx', )


def valormedioUsinaT1(ipdo: pd.DataFrame):
    """
   data,colunassh,transpose,to_excel, n_index.
   ----------
   ipdo : DataFrame da tabela.

   Retorna um arquivo .xlsx em formato de tabela a respeito dos valores da mé-
   dia diária das usinas T1.
   -------
   """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\ValMeT11Ro_JaqII.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
          r'C:\Users\e806128\Desktop\1\ValMeT11Ro_JaqII.xlsx')
        add_inf2 = pd.read_excel(
           r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_Bm.xlsx')
        add_inf3 = pd.read_excel(
           r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_cant.xlsx')
        add_inf4 = pd.read_excel(
           r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_pa.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1R_santaluz.xlsx')
        add_inf6 = pd.read_excel(
           r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_pal.xlsx')
        add_inf7 = pd.read_excel(
           r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_ba.xlsx')
        add_inf8 = pd.read_excel(
          r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_mont.xlsx')
        add_inf9 = pd.read_excel(
           r'C:\Users\e806128\Desktop\1\ValMeT!Ro_montstoII.xlsx')
        add_inf10 = pd.read_excel(
           r'C:\Users\e806128\Desktop\1\ValMeT1Romontsto.xlsx')
        add_inf11 = pd.read_excel(
           r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_f.xlsx')
        add_inf12 = pd.read_excel(
           r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_do.xlsx')
        add_inf13 = pd.read_excel(
           r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_totrr.xlsx')
        no = no + 3
        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={409: data})
        vari1 = valor10_re1.iloc[0:1]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           dado9)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT11Ro_JaqII.xlsx',
            index=False)

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={410: data})
        vari1 = valor10_re1.iloc[1:2]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_Bm.xlsx',
            index=False)
        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={411: data})
        vari1 = valor10_re1.iloc[2:3]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_cant.xlsx',
            index=False)
        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={412: data})
        vari1 = valor10_re1.iloc[3:4]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_pa.xlsx',
            index=False)

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={413: data})
        vari1 = valor10_re1.iloc[4:5]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1R_santaluz.xlsx',
            index=False)
        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={414: data})
        vari1 = valor10_re1.iloc[5:6]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf6.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf6 = add_inf6.rename(columns={'Unnamed: 0': ' '})
        add_inf6.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_pal.xlsx',
            index=False)

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={415: data})
        vari1 = valor10_re1.iloc[6:7]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf7.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf7 = add_inf7.rename(columns={'Unnamed: 0': ' '})
        add_inf7.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_ba.xlsx',
            index=False)

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={416: data})
        vari1 = valor10_re1.iloc[7:8]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf8.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf8 = add_inf8.rename(columns={'Unnamed: 0': ' '})
        add_inf8.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_mont.xlsx',
            index=False)

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={417: data})
        vari1 = valor10_re1.iloc[8:9]
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf9.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf9 = add_inf9.rename(columns={'Unnamed: 0': ' '})
        add_inf9.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT!Ro_montstoII.xlsx',
            index=False)

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={418: data})
        vari1 = valor10_re1.iloc[9:10]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf10.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf10 = add_inf10.rename(columns={'Unnamed: 0': ' '})
        add_inf10.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Romontsto.xlsx',
            index=False)

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={419: data})
        vari1 = valor10_re1.iloc[10:11]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf11.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf11 = add_inf11.rename(columns={'Unnamed: 0': ' '})
        add_inf11.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_f.xlsx',
            index=False)
        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={420: data})
        vari1 = valor10_re1.iloc[11:12]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf12.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf12 = add_inf12.rename(columns={'Unnamed: 0': ' '})
        add_inf12.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_do.xlsx',
            index=False)

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={423: data})
        vari1 = valor10_re1.iloc[14:15]

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf13.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf13 = add_inf13.rename(columns={'Unnamed: 0': ' '})
        add_inf13.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_totrr.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={409: data})
        valor10_v = valor10_re1.iloc[0:1]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT11Ro_JaqII.xlsx', )

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={410: data})
        valor10_v = valor10_re1.iloc[1:2]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_Bm.xlsx',)
        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={411: data})
        valor10_v = valor10_re1.iloc[2:3]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_cant.xlsx',)
        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={412: data})
        valor10_v = valor10_re1.iloc[3:4]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_pa.xlsx',)

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={413: data})
        valor10_v = valor10_re1.iloc[4:5]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1R_santaluz.xlsx',)
        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={414: data})
        valor10_v = valor10_re1.iloc[5:6]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_pal.xlsx',)

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={415: data})
        valor10_v = valor10_re1.iloc[6:7]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_ba.xlsx',)

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={416: data})
        valor10_v = valor10_re1.iloc[7:8]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_mont.xlsx',)

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={417: data})
        valor10_v = valor10_re1.iloc[8:9]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT!Ro_montstoII.xlsx',)

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={418: data})
        valor10_v = valor10_re1.iloc[9:10]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Romontsto.xlsx', )

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={419: data})
        valor10_v = valor10_re1.iloc[10:11]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_f.xlsx')
        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={420: data})
        valor10_v = valor10_re1.iloc[11:12]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_do.xlsx',)

        valor9 = ipdo.iloc[409:425, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária Programada',
                                'Unnamed: 6': 'Media Diária verificada',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={423: data})
        valor10_v = valor10_re1.iloc[14:15]

        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\ValMeT1Roraima_totrr.xlsx')


def valormediusinaT2(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores da mé-
    dia diária das usinas T2.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\valMedT2nord_totalne.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2SUl_saosape.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2sul_barrabonitaI.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2sul_total.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_oncapintada.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_santavitoria.xlsx')
        add_inf6 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_paulinaverde.xlsx')
        add_inf7 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_xavantes.xlsx')
        add_inf8 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_totalse.xlsx')
        add_inf9 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ERBCandeias.xlsx')
        add_inf10 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ProsperidadeI.xlsx')
        add_inf11 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ProsperidadeIII.xlsx')
        add_inf12 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ProsperidadeII.xlsx')
        add_inf13 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_Curumim.xlsx')
        add_inf14 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_totalne.xlsx')
        no = no + 3
        valor9 = ipdo.iloc[426:437, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={431: data})
        vari1 = valor10_re1.iloc[5:6]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2SUl_saosape.xlsx', )
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           dado9)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2SUl_saosape.xlsx',
            index=False)

        valor9 = ipdo.iloc[426:439, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={432: data})
        vari1 = valor10_re1.iloc[6:7]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2sul_barrabonitaI.xlsx', )
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2sul_barrabonitaI.xlsx',
            index=False)

        valor9 = ipdo.iloc[426:439, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={436: data})
        vari1 = valor10_re1.iloc[10:11]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2sul_total.xlsx', )

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2sul_total.xlsx',
            index=False)

        valor9 = ipdo.iloc[437:450, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={439: data})
        vari1 = valor10_re1.iloc[2:3]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_oncapintada.xlsx')

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_oncapintada.xlsx',
            index=False)

        valor9 = ipdo.iloc[437:450, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={440: data})
        vari1 = valor10_re1.iloc[3:4]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_santavitoria.xlsx')

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_santavitoria.xlsx',
            index=False)
        valor9 = ipdo.iloc[437:450, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={441: data})
        vari1 = valor10_re1.iloc[4:5]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_paulinaverde.xlsx', )
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf6.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf6 = add_inf6.rename(columns={'Unnamed: 0': ' '})
        add_inf6.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_paulinaverde.xlsx',
            index=False)

        valor9 = ipdo.iloc[437:450, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={442: data})
        vari1 = valor10_re1.iloc[5:6]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_xavantes.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf7.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf7 = add_inf7.rename(columns={'Unnamed: 0': ' '})
        add_inf7.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_xavantes.xlsx',
            index=False)

        valor9 = ipdo.iloc[437:451, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={443: data})
        vari1 = valor10_re1.iloc[13:14]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_totalse.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf8.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf8 = add_inf8.rename(columns={'Unnamed: 0': ' '})
        add_inf8.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_totalse.xlsx',
            index=False)

        valor9 = ipdo.iloc[451:461, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={453: data})
        vari1 = valor10_re1.iloc[2:3]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ERBCandeias.xlsx', )
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf9.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9)
        add_inf9 = add_inf9.rename(columns={'Unnamed: 0': ' '})
        add_inf9.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ERBCandeias.xlsx',
            index=False)

        valor9 = ipdo.iloc[451:461, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={454: data})
        vari1 = valor10_re1.iloc[3:4]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ProsperidadeI.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf10.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf10 = add_inf10.rename(columns={'Unnamed: 0': ' '})
        add_inf10.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ProsperidadeI.xlsx',
            index=False)

        valor9 = ipdo.iloc[451:461, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={455: data})
        vari1 = valor10_re1.iloc[4:5]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ProsperidadeIII.xlsx')

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf11.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf11 = add_inf11.rename(columns={'Unnamed: 0': ' '})
        add_inf11.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ProsperidadeIII.xlsx',
            index=False)

        valor9 = ipdo.iloc[451:461, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={456: data})
        vari1 = valor10_re1.iloc[5:6]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ProsperidadeII.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf12.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf12 = add_inf12.rename(columns={'Unnamed: 0': ' '})
        add_inf12.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ProsperidadeII.xlsx',
            index=False)

        valor9 = ipdo.iloc[451:461, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={457: data})
        vari1 = valor10_re1.iloc[6:7]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_Curumim.xlsx')

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf13.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf13 = add_inf13.rename(columns={'Unnamed: 0': ' '})
        add_inf13.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_Curumim.xlsx',
            index=False)

        valor9 = ipdo.iloc[451:461, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={457: data})
        vari1 = valor10_re1.iloc[9:10]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_totalne.xlsx')

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        add_inf14.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6,
                             dado7,
                             dado8,
                             dado9)
        add_inf14 = add_inf14.rename(columns={'Unnamed: 0': ' '})
        add_inf14.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_totalne.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        valor9 = ipdo.iloc[426:437, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={431: data})
        vari1 = valor10_re1.iloc[5:6]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2SUl_saosape.xlsx', )

        valor9 = ipdo.iloc[426:439, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={432: data})
        vari1 = valor10_re1.iloc[6:7]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2sul_barrabonitaI.xlsx', )

        valor9 = ipdo.iloc[426:439, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={436: data})
        vari1 = valor10_re1.iloc[10:11]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2sul_total.xlsx', )

        valor9 = ipdo.iloc[437:450, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={439: data})
        vari1 = valor10_re1.iloc[2:3]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_oncapintada.xlsx')

        valor9 = ipdo.iloc[437:450, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={440: data})
        vari1 = valor10_re1.iloc[3:4]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_santavitoria.xlsx')
        valor9 = ipdo.iloc[437:450, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={441: data})
        vari1 = valor10_re1.iloc[4:5]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_paulinaverde.xlsx', )

        valor9 = ipdo.iloc[437:450, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={442: data})
        vari1 = valor10_re1.iloc[5:6]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_xavantes.xlsx')

        valor9 = ipdo.iloc[437:451, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={450: data})
        vari1 = valor10_re1.iloc[13:14]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2Sud_totalse.xlsx')

        valor9 = ipdo.iloc[451:461, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={453: data})
        vari1 = valor10_re1.iloc[2:3]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ERBCandeias.xlsx', )

        valor9 = ipdo.iloc[451:461, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={454: data})
        vari1 = valor10_re1.iloc[3:4]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ProsperidadeI.xlsx')

        valor9 = ipdo.iloc[451:461, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={455: data})
        vari1 = valor10_re1.iloc[4:5]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ProsperidadeIII.xlsx')

        valor9 = ipdo.iloc[451:461, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={456: data})
        vari1 = valor10_re1.iloc[5:6]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_ProsperidadeII.xlsx')

        valor9 = ipdo.iloc[451:461, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={457: data})
        vari1 = valor10_re1.iloc[6:7]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_Curumim.xlsx')

        valor9 = ipdo.iloc[451:461, 0:10]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária Programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={460: data})
        vari1 = valor10_re1.iloc[9:10]

        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\valMedT2nord_totalne.xlsx')


def usinascommaisrazao(ipdo: pd.DataFrame):
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\UsinCMRaz_aparecida.xlsx'
    no = 2
    dado1 = data
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
          r'C:\Users\e806128\Desktop\1\UsinCMRaz_aparecida.xlsx')
        add_inf2 = pd.read_excel(
          r'C:\Users\e806128\Desktop\1\UsinCMRaz_maua3.xlsx')
        valor9 = ipdo.iloc[473:484, 0:10]
        colunassh = [0, 1, 3, 5, 7, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Media Verificada',
                                    'Unnamed: 4': 'Media programada',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Razaõ de despacho ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={477: data, 478: ' '})
        vari1 = valor10_re1.iloc[4:6]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\UsinCMRaz_aparecida.xlsx')
        no = len(add_inf) + 1
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        no = len(add_inf) + 1
        vari11 = vari1
        vari11 = vari11.iloc[1:2]
        dado23 = vari11.iloc[0, 0]
        dado33 = vari11.iloc[0, 1]
        dado43 = vari11.iloc[0, 2]
        dado1 = (' ')
        add_inf.loc[no] = (dado1,
                           dado23,
                           dado33,
                           dado43
                           )
        add_inf.to_excel(
           r'C:\Users\e806128\Desktop\1\UsinCMRaz_aparecida.xlsx',
           index=False)

        valor9 = ipdo.iloc[473:484, 0:10]
        colunassh = [0, 1, 3, 5, 7, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Media Verificada',
                                    'Unnamed: 4': 'Media programada',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Razaõ de despacho ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={480: data, 481: ' ', 482: ' '})
        valor10_v = valor10_re1.iloc[7:10]
        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\UsinCMRaz_maua3.xlsx')
        no = len(add_inf2) + 1
        dado1 = data
        valor10_v1 = valor10_v
        dado2 = valor10_v1.iloc[0, 0]
        dado3 = valor10_v1.iloc[0, 1]
        dado4 = valor10_v1.iloc[0, 2]

        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        no = len(add_inf2) + 1
        valor10_v2 = valor10_v
        valor10_v2 = valor10_v2.iloc[1:2]
        dado23 = valor10_v2.iloc[0, 0]
        dado33 = valor10_v2.iloc[0, 1]
        dado43 = valor10_v2.iloc[0, 2]
        dado1 = (' ')
        add_inf2.loc[no] = (dado1,
                            dado23,
                            dado33,
                            dado43)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        no = len(add_inf2) + 1
        valor10_v3 = valor10_v
        valor10_v3 = valor10_v.iloc[2:3]
        dado23 = valor10_v3.iloc[0, 0]
        dado33 = valor10_v3.iloc[0, 1]
        dado43 = valor10_v3.iloc[0, 2]
        dado1 = (' ')
        add_inf2.loc[no] = (dado1,
                            dado23,
                            dado33,
                            dado43)
        add_inf2.to_excel(
          r'C:\Users\e806128\Desktop\1\UsinCMRaz_maua3.xlsx',
          index=False)
    else:
        print("O arquivo não existe.")
        valor9 = ipdo.iloc[473:484, 0:10]
        colunassh = [0, 1, 3, 5, 7, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Media Verificada',
                                    'Unnamed: 4': 'Media programada',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Razaõ de despacho ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={477: data, 478: ' '})
        vari1 = valor10_re1.iloc[4:6]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\UsinCMRaz_aparecida.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]

        valor9 = ipdo.iloc[473:484, 0:10]
        colunassh = [0, 1, 3, 5, 7, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Media Verificada',
                                    'Unnamed: 4': 'Media programada',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Razaõ de despacho ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={480: data, 481: ' ', 482: ' '})
        valor10_v = valor10_re1.iloc[7:10]
        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\UsinCMRaz_maua3.xlsx')


def geracaotermicaT1et2(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores da mé-
    dia diária da geração termica T1 e T2.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\GerTerT1eT2_totalsin.xlsx'
    no = 2
    dado1 = data
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
          r'C:\Users\e806128\Desktop\1\GerTerT1eT2_sudeste.xlsx')
        add_inf2 = pd.read_excel(
          r'C:\Users\e806128\Desktop\1\GerTerT1eT2_sul.xlsx')
        add_inf3 = pd.read_excel(
          r'C:\Users\e806128\Desktop\1\GerTerT1eT2_nordeste.xlsx')
        add_inf4 = pd.read_excel(
          r'C:\Users\e806128\Desktop\1\GerTerT1eT2_norte.xlsx')
        add_inf5 = pd.read_excel(
          r'C:\Users\e806128\Desktop\1\GerTerT1eT2_totalsin.xlsx')
        no = no + 3
        valor9 = ipdo.iloc[690:700, 0:10]
        colunassh = [0, 1, 2, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={694: data})
        vari1 = valor10_re1.iloc[4:5]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\GerTerT1eT2_sudeste.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\GerTerT1eT2_sudeste.xlsx',
            index=False)
        valor9 = ipdo.iloc[690:700, 0:10]
        colunassh = [0, 1, 2, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={694: data})
        vari1 = valor10_re1.iloc[5:6]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\GerTerT1eT2_sul.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\GerTerT1eT2_sul.xlsx',
            index=False)
        valor9 = ipdo.iloc[690:700, 0:10]
        colunassh = [0, 1, 2, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={694: data})
        vari1 = valor10_re1.iloc[6:7]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\GerTerT1eT2_nordeste.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\GerTerT1eT2_nordeste.xlsx',
            index=False)
        valor9 = ipdo.iloc[690:700, 0:10]
        colunassh = [0, 1, 2, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={694: data})
        vari1 = valor10_re1.iloc[7:8]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\GerTerT1eT2_norte.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7)
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\GerTerT1eT2_norte.xlsx',
            index=False)

        valor9 = ipdo.iloc[690:700, 0:10]
        colunassh = [0, 1, 2, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                    'Unnamed: 2': 'Razão Despacho',
                                    'Unnamed: 3': 'Capacidade Instalada',
                                    'Unnamed: 4': 'Capacidade Disponível',
                                    'Unnamed: 5': 'Media Diária programada',
                                    'Unnamed: 6': 'Media Diária verificada ',
                                    'Unnamed: 7': 'Média diária Difer',
                                    'Unnamed: 8': 'Variação %',
                                    'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={694: data})
        vari1 = valor10_re1.iloc[9:10]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\GerTerT1eT2_totalsin.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7)
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\GerTerT1eT2_totalsin.xlsx',
            index=False)

    else:
        print('O arquivo não existe.')
        valor9 = ipdo.iloc[690:700, 0:10]
        colunassh = [0, 1, 2, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={694: data})
        valor10_v = valor10_re1.iloc[4:5]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\GerTerT1eT2_sudeste.xlsx')
        valor9 = ipdo.iloc[690:700, 0:10]
        colunassh = [0, 1, 2, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={695: data})
        valor10_v = valor10_re1.iloc[5:6]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\GerTerT1eT2_sul.xlsx')
        valor9 = ipdo.iloc[690:700, 0:10]
        colunassh = [0, 1, 2, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={696: data})
        valor10_v = valor10_re1.iloc[6:7]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\GerTerT1eT2_nordeste.xlsx')

        valor9 = ipdo.iloc[690:700, 0:10]
        colunassh = [0, 1, 2, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={697: data})
        valor10_v = valor10_re1.iloc[7:8]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\GerTerT1eT2_norte.xlsx')

        valor9 = ipdo.iloc[690:700, 0:10]
        colunassh = [0, 1, 2, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 2': 'Razão Despacho',
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 4': 'Capacidade Disponível',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada ',
                                'Unnamed: 7': 'Média diária Difer',
                                'Unnamed: 8': 'Variação %',
                                'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={699: data})
        valor10_v = valor10_re1.iloc[9:10]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\GerTerT1eT2_totalsin.xlsx')


def diferencasentrecapadidades(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores da mé-
    dia diária das usinas diferenciando as capacidades.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DiferCap_total.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_NorteFlum.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_Cubatão.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_JlacendaC.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_maracanaúI.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_maranhãoIII.xlsx')
        add_inf6 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_total.xlsx')
        no = no + 3
        valor9 = ipdo.iloc[703:713, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={707: data})
        vari1 = valor10_re1.iloc[4:5]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_NorteFlum.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_NorteFlum.xlsx',
            index=False)
        valor9 = ipdo.iloc[703:713, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={708: data})
        vari1 = valor10_re1.iloc[5:6]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_Cubatão.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_Cubatão.xlsx',
            index=False)
        valor9 = ipdo.iloc[703:713, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={709: data})
        vari1 = valor10_re1.iloc[6:7]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_JlacendaC.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_JlacendaC.xlsx',
            index=False)
        valor9 = ipdo.iloc[703:713, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={710: data})
        vari1 = valor10_re1.iloc[7:8]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_maracanaúI.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_maracanaúI.xlsx',
            index=False)
        valor9 = ipdo.iloc[703:713, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={711: data})
        vari1 = valor10_re1.iloc[8:9]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_maranhãoIII.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_maranhãoIII.xlsx',
            index=False)
        valor9 = ipdo.iloc[703:713, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={712: data})
        vari1 = valor10_re1.iloc[9:10]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_total.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf6.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf6 = add_inf6.rename(columns={'Unnamed: 0': ' '})
        add_inf6.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_total.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        valor9 = ipdo.iloc[703:713, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={707: data})
        vari1 = valor10_re1.iloc[4:5]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_NorteFlum.xlsx')
        valor9 = ipdo.iloc[703:713, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={708: data})
        vari1 = valor10_re1.iloc[5:6]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_Cubatão.xlsx')
        valor9 = ipdo.iloc[703:713, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={709: data})
        vari1 = valor10_re1.iloc[6:7]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_JlacendaC.xlsx')
        valor9 = ipdo.iloc[703:713, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={710: data})
        vari1 = valor10_re1.iloc[7:8]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_maracanaúI.xlsx')
        valor9 = ipdo.iloc[703:713, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={711: data})
        vari1 = valor10_re1.iloc[8:9]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_maranhãoIII.xlsx')
        valor9 = ipdo.iloc[703:713, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 6': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={712: data})
        vari1 = valor10_re1.iloc[9:10]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DiferCap_total.xlsx')


def deferencaentrecapacidadesoperativa(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores da mé-
    dia diária das usinas evidenciando as capacidades operativas.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DiEnCap_Total.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_AndraII.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Doatlant.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_PalmdeGo.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_goiania2.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Cuiaba.xlsx')
        add_inf6 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Daia.xlsx')
        add_inf7 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Figueira.xlsx')
        add_inf8 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Uruguai.xlsx')
        add_inf9 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_ValedAçú.xlsx')
        add_inf10 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_PBIII.xlsx')
        add_inf11 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Fortal.xlsx')
        add_inf12 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Apoena.xlsx')
        add_inf13 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Guarani.xlsx')
        add_inf14 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Petrol.xlsx')
        add_inf15 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_PuFerroI.xlsx')
        add_inf16 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Termom.xlsx')
        add_inf17 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_CristRoc.xlsx')
        add_inf18 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Jaraqui.xlsx')
        add_inf19 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Tambaqui.xlsx')
        add_inf20 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Xavantes.xlsx')
        add_inf21 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_ERBCan.xlsx')
        add_inf22 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Curumim.xlsx')
        add_inf23 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Total.xlsx')
        no = no + 3
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={738: data})
        vari1 = valor10_re1.iloc[4:5]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_AndraII.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_AndraII.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={739: data})
        vari1 = valor10_re1.iloc[5:6]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Doatlant.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Doatlant.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={740: data})
        vari1 = valor10_re1.iloc[6:7]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_PalmdeGo.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_PalmdeGo.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={741: data})
        vari1 = valor10_re1.iloc[7:8]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_goiania2.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_goiania2.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={742: data})
        vari1 = valor10_re1.iloc[8:9]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Cuiaba.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Cuiaba.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={743: data})
        vari1 = valor10_re1.iloc[9:10]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Daia.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf6.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf6 = add_inf6.rename(columns={'Unnamed: 0': ' '})
        add_inf6.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Daia.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={744: data})
        vari1 = valor10_re1.iloc[10:11]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Figueira.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf7.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf7 = add_inf7.rename(columns={'Unnamed: 0': ' '})
        add_inf7.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Figueira.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={745: data})
        vari1 = valor10_re1.iloc[11:12]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Uruguai.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf8.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf8 = add_inf8.rename(columns={'Unnamed: 0': ' '})
        add_inf8.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Uruguai.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={746: data})
        vari1 = valor10_re1.iloc[12:13]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_ValedAçú.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf9.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf9 = add_inf9.rename(columns={'Unnamed: 0': ' '})
        add_inf9.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_ValedAçú.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={747: data})
        vari1 = valor10_re1.iloc[13:14]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_PBIII.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf10.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4)
        add_inf10 = add_inf10.rename(columns={'Unnamed: 0': ' '})
        add_inf10.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_PBIII.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={748: data})
        vari1 = valor10_re1.iloc[14:15]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Fortal.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf11.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4)
        add_inf11 = add_inf11.rename(columns={'Unnamed: 0': ' '})
        add_inf11.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Fortal.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={749: data})
        vari1 = valor10_re1.iloc[15:16]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Apoena.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf12.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4)
        add_inf12 = add_inf12.rename(columns={'Unnamed: 0': ' '})
        add_inf12.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Apoena.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={750: data})
        vari1 = valor10_re1.iloc[16:17]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Guarani.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf13.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4)
        add_inf13 = add_inf13.rename(columns={'Unnamed: 0': ' '})
        add_inf13.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Guarani.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={751: data})
        vari1 = valor10_re1.iloc[17:18]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Petrol.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf14.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4)
        add_inf14 = add_inf14.rename(columns={'Unnamed: 0': ' '})
        add_inf14.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Petrol.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={752: data})
        vari1 = valor10_re1.iloc[18:19]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_PuFerroI.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf15.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4)
        add_inf15 = add_inf15.rename(columns={'Unnamed: 0': ' '})
        add_inf15.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_PuFerroI.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={753: data})
        vari1 = valor10_re1.iloc[19:20]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Termom.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf16.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4)
        add_inf16 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Termom.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={754: data})
        vari1 = valor10_re1.iloc[20:21]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_CristRoc.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf17.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4)
        add_inf17 = add_inf17.rename(columns={'Unnamed: 0': ' '})
        add_inf17.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_CristRoc.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={755: data})
        vari1 = valor10_re1.iloc[21:22]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Jaraqui.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf18.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4)
        add_inf18 = add_inf18.rename(columns={'Unnamed: 0': ' '})
        add_inf18.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Jaraqui.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={756: data})
        vari1 = valor10_re1.iloc[22:23]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Tambaqui.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf19.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4)
        add_inf19 = add_inf19.rename(columns={'Unnamed: 0': ' '})
        add_inf19.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Tambaqui.xlsx',
            index=False)

        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={757: data})
        vari1 = valor10_re1.iloc[23:24]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Xavantes.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf20.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4)
        add_inf20 = add_inf20.rename(columns={'Unnamed: 0': ' '})
        add_inf20.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Xavantes.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={758: data})
        vari1 = valor10_re1.iloc[24:25]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_ERBCan.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf21.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4)
        add_inf21 = add_inf21.rename(columns={'Unnamed: 0': ' '})
        add_inf21.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_ERBCan.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={759: data})
        vari1 = valor10_re1.iloc[25:26]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Curumim.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf22.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4)
        add_inf22 = add_inf22.rename(columns={'Unnamed: 0': ' '})
        add_inf22.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Curumim.xlsx',
            index=False)
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={760: data})
        vari1 = valor10_re1.iloc[26:27]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Total.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf23.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4)
        add_inf23 = add_inf23.rename(columns={'Unnamed: 0': ' '})
        add_inf23.to_excel(
            r'C:\Users\e806128\Desktop\1\DiEnCap_Total.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={738: data})
        valor10_v = valor10_re1.iloc[4:5]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_AndraII.xlsx')

        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={739: data})
        valor10_v = valor10_re1.iloc[5:6]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Doatlant.xlsx')
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={740: data})
        valor10_v = valor10_re1.iloc[6:7]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_PalmdeGo.xlsx')
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={741: data})
        valor10_v = valor10_re1.iloc[7:8]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_goiania2.xlsx')
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={742: data})
        valor10_v = valor10_re1.iloc[8:9]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Cuiaba.xlsx')

        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={743: data})
        valor10_v = valor10_re1.iloc[9:10]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Daia.xlsx')

        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={744: data})
        valor10_v = valor10_re1.iloc[10:11]
        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Figueira.xlsx')

        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={745: data})
        valor10_v = valor10_re1.iloc[11:12]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Uruguai.xlsx')
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={746: data})
        valor10_v = valor10_re1.iloc[12:13]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_ValedAçú.xlsx')
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                            'Unnamed: 2': 'Razão Despacho',
                            'Unnamed: 3': 'Capacidade Instalada',
                            'Unnamed: 4': 'Capacidade Disponível',
                            'Unnamed: 5': 'Capacidade disónível',
                            'Unnamed: 6': 'Media Diária verificada',
                            'Unnamed: 7': 'Diferança',
                            'Unnamed: 8': 'Variação %',
                            'Unnamed: 9': 'OBS'})
        valor10_re1 = valor10_re.rename(index={747: data})
        valor10_v = valor10_re1.iloc[13:14]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_PBIII.xlsx')

        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={748: data})
        valor10_v = valor10_re1.iloc[14:15]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Fortal.xlsx')

        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={749: data})
        valor10_v = valor10_re1.iloc[15:16]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Apoena.xlsx')

        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={750: data})
        valor10_v = valor10_re1.iloc[16:17]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Guarani.xlsx')

        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={751: data})
        valor10_v = valor10_re1.iloc[17:18]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Petrol.xlsx')

        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={752: data})
        valor10_v = valor10_re1.iloc[18:19]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_PuFerroI.xlsx')
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={753: data})
        valor10_v = valor10_re1.iloc[19:20]
        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Termom.xlsx')
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={754: data})
        valor10_v = valor10_re1.iloc[20:21]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_CristRoc.xlsx')

        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={755: data})
        valor10_v = valor10_re1.iloc[21:22]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Jaraqui.xlsx')
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={756: data})
        valor10_v = valor10_re1.iloc[22:23]
        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Tambaqui.xlsx')

        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={757: data})
        valor10_v = valor10_re1.iloc[23:24]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Xavantes.xlsx')

        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={758: data})
        valor10_v = valor10_re1.iloc[24:25]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_ERBCan.xlsx')
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={759: data})
        valor10_v = valor10_re1.iloc[25:26]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Curumim.xlsx')
        valor9 = ipdo.iloc[734:762, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={760: data})
        valor10_v = valor10_re1.iloc[26:27]
        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\DiEnCap_Total.xlsx')


def restriçãoemanutenção(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito da restrição da
    manutenção.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\ResOpMa_Total.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ResOpMa_GNA1.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ResOpMa_W.Arjona.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ResOpMa_Aparecid.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\ResOpMa_Total.xlsx')
        no = no + 3
        valor9 = ipdo.iloc[800:809, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={804: data})
        vari1 = valor10_re1.iloc[4:5]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ResOpMa_GNA1.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\ResOpMa_GNA1.xlsx',
            index=False)
        valor9 = ipdo.iloc[800:809, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={805: data})
        vari1 = valor10_re1.iloc[5:6]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ResOpMa_W.Arjona.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\ResOpMa_W.Arjona.xlsx',
            index=False)
        valor9 = ipdo.iloc[800:809, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={806: data})
        vari1 = valor10_re1.iloc[6:7]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ResOpMa_Aparecid.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
           r'C:\Users\e806128\Desktop\1\ResOpMa_Aparecid.xlsx',
           index=False)
        valor9 = ipdo.iloc[800:809, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={807: data})
        vari1 = valor10_re1.iloc[7:8]

        vari1.to_excel(r'C:\Users\e806128\Desktop\1\ResOpMa_Total.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\ResOpMa_Total.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        valor9 = ipdo.iloc[800:809, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={804: data})
        valor10_v = valor10_re1.iloc[4:5]
        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\ResOpMa_GNA1.xlsx')
        valor9 = ipdo.iloc[800:809, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={805: data})
        valor10_v = valor10_re1.iloc[5:6]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\ResOpMa_W.Arjona.xlsx')

        valor9 = ipdo.iloc[800:809, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={806: data})
        valor10_v = valor10_re1.iloc[6:7]
        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\ResOpMa_Aparecid.xlsx')
        valor9 = ipdo.iloc[800:809, 0:10]
        colunassh = [0, 1, 2, 4, 6, 8, 9]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 3': 'Capacidade Instalada',
                                'Unnamed: 5': 'Media Diária programada',
                                'Unnamed: 7': 'Media Diária verificada '
                          })
        valor10_re1 = valor10_re.rename(index={807: data})
        valor10_v = valor10_re1.iloc[7:8]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\ResOpMa_Total.xlsx')


def To554(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores da mé-
    dia diária das usinas com o total.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\To,5.5.4_total.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_manut.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_restricaoop.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_restopevaem.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_dresagrega.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_total.xlsx')
        no = no + 3
        valor9 = ipdo.iloc[840:851, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={845: data})
        vari1 = valor10_re1.iloc[5:6]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\To,5.5.4_manut.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_manut.xlsx',
            index=False)

        valor9 = ipdo.iloc[840:851, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={846: data})
        vari1 = valor10_re1.iloc[6:7]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_restricaoop.xlsx')

        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_restricaoop.xlsx',
            index=False)
        valor9 = ipdo.iloc[840:851, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={847: data})
        vari1 = valor10_re1.iloc[7:8]
        vari1.to_excel(r'C:\Users\e806128\Desktop\1\To,5.5.4_restopevaem.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_restopevaem.xlsx',
            index=False)
        valor9 = ipdo.iloc[840:851, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={848: data})
        vari1 = valor10_re1.iloc[8:9]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_dresagrega.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_dresagrega.xlsx',
            index=False)
        valor9 = ipdo.iloc[840:851, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={849: data})
        vari1 = valor10_re1.iloc[9:10]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_total.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_total.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        valor9 = ipdo.iloc[840:851, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)

        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={845: data})
        valor10_v = valor10_re1.iloc[5:6]

        valor10_v.to_excel(r'C:\Users\e806128\Desktop\1\To,5.5.4_manut.xlsx')

        valor9 = ipdo.iloc[840:851, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={846: data})
        valor10_v = valor10_re1.iloc[6:7]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_restricaoop.xlsx')
        valor9 = ipdo.iloc[840:851, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={847: data})
        valor10_v = valor10_re1.iloc[7:8]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_restopevaem.xlsx')
        valor9 = ipdo.iloc[840:851, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={848: data})
        valor10_v = valor10_re1.iloc[8:9]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_dresagrega.xlsx')
        valor9 = ipdo.iloc[840:851, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={849: data})
        valor10_v = valor10_re1.iloc[9:10]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\To,5.5.4_total.xlsx')


def diferencasentreCIeA(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito da diferenciação .
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DifCapInA_Total.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_tresl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_SykueI.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_PotigIII.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_Campos.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_SantaCruz.xlsx')
        add_inf6 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_Pirat.xlsx')
        add_inf7 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_TermonII.xlsx')
        add_inf8 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_Total.xlsx')
        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={857: data})
        vari1 = valor10_re1.iloc[1:2]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_tresl.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_tresl.xlsx',
            index=False)

        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={858: data})
        vari1 = valor10_re1.iloc[2:3]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_SykueI.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_SykueI.xlsx',
            index=False)
        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={859: data})
        vari1 = valor10_re1.iloc[3:4]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_PotigIII.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_PotigIII.xlsx',
            index=False)
        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={860: data})
        vari1 = valor10_re1.iloc[4:5]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_Campos.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_Campos.xlsx',
            index=False)
        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={861: data})
        vari1 = valor10_re1.iloc[5:6]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_SantaCruz.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_SantaCruz.xlsx',
            index=False)
        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={862: data})
        vari1 = valor10_re1.iloc[6:7]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_Pirat.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf6.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf6 = add_inf6.rename(columns={'Unnamed: 0': ' '})
        add_inf6.to_excel(
           r'C:\Users\e806128\Desktop\1\DifCapInA_Pirat.xlsx',
           index=False)
        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={863: data})
        vari1 = valor10_re1.iloc[7:8]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_TermonII.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf7.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf7 = add_inf7.rename(columns={'Unnamed: 0': ' '})
        add_inf7.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_TermonII.xlsx',
            index=False)
        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={872: data})
        vari1 = valor10_re1.iloc[16:17]
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_Total.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        add_inf8.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4)
        add_inf8 = add_inf8.rename(columns={'Unnamed: 0': ' '})
        add_inf8.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_Total.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={857: data})
        valor10_v = valor10_re1.iloc[1:2]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_tresl.xlsx')

        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={858: data})
        valor10_v = valor10_re1.iloc[2:3]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_SykueI.xlsx')
        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={859: data})
        valor10_v = valor10_re1.iloc[3:4]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_PotigIII.xlsx')
        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={860: data})
        valor10_v = valor10_re1.iloc[4:5]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_Campos.xlsx')
        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={861: data})
        valor10_v = valor10_re1.iloc[5:6]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_SantaCruz.xlsx')
        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={862: data})
        valor10_v = valor10_re1.iloc[6:7]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_Pirat.xlsx')
        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={863: data})
        valor10_v = valor10_re1.iloc[7:8]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_TermonII.xlsx')
        valor9 = ipdo.iloc[856:874, 0:10]
        colunassh = [0, 1, 2, 3, 4, 6, 8]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_re = valor10_d.rename(columns={
                                'Unnamed: 5': 'Capacidade Instalada',
                                'Unnamed: 7': 'Disponível',
                                'Unnamed: 9': 'Diferença '
                          })
        valor10_re1 = valor10_re.rename(index={872: data})
        valor10_v = valor10_re1.iloc[16:17]
        valor10_v.to_excel(
            r'C:\Users\e806128\Desktop\1\DifCapInA_Total.xlsx')


def demandasMaximasSin(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores da mé-
    dia das demandas máximas no sin.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\demMa_histórico.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\demMa_sin.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\demMa_norte.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\demMa_nordeste.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\demMa_SECO.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\demMa_SUL.xlsx')
        add_inf6 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\demMa_INTERC.xlsx')
        add_inf7 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\demMa_ITAIPU.xlsx')
        add_inf8 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\demMa_TERMOSE.xlsx')
        add_inf9 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\demMa_ATUAL.xlsx')
        add_inf10 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\demMa_histórico.xlsx')
        dado1 = data
        no = no + 3
        valor9 = ipdo.iloc[1080:1090, 10:13]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1081: 'Hidro Nac',
                                                1082: 'Itaip',
                                                1083: 'Termo Nuc',
                                                1084: 'Termo Conv',
                                                1085: 'Eólica ',
                                                1086: 'Solar',
                                                1087: ' Total SIN',
                                                1088: 'Interc. Inter',
                                                1089: 'Carga'})
        valor10_re1 = valor10_re.rename(index={'Unnamed: 12': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        vari1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_sin.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        dado8 = vari1.iloc[0, 6]
        dado9 = vari1.iloc[0, 7]
        dado10 = vari1.iloc[0, 8]

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           dado9,
                           dado10)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_sin.xlsx',
            index=False)

        valor9 = ipdo.iloc[1091:1098, 10:13]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={
                                                1092: 'Hidro',
                                                1093: 'Termo',
                                                1094: 'Eólica',
                                                1095: 'Solar',
                                                1096: 'Total Ger ',
                                                1097: 'Carga',
                                                })
        valor10_re1 = valor10_re.rename(index={'Unnamed: 12': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        vari1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_norte.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]

        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_norte.xlsx',
            index=False)
        valor9 = ipdo.iloc[1098:1105, 10:13]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1099: 'Hidro',
                                                1100: 'Termo',
                                                1101: 'Eólica',
                                                1102: 'Solar',
                                                1103: 'Total Ger',
                                                1104: 'Carga',
                                                })
        valor10_re1 = valor10_re.rename(index={'Unnamed: 12': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        vari1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_nordeste.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7)
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_nordeste.xlsx',
            index=False)
        valor9 = ipdo.iloc[1106:1113, 10:13]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1107: 'Hidro',
                                                1108: 'Termo',
                                                1109: 'Eólica',
                                                1110: 'Solar',
                                                1111: 'Total Ger',
                                                1112: 'Carga',
                                                })
        valor10_re1 = valor10_re.rename(index={'Unnamed: 12': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        vari1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_SECO.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]

        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7)
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_SECO.xlsx',
            index=False)
        valor9 = ipdo.iloc[1113:1120, 10:13]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1114: 'Hidro',
                                                1115: 'Termo',
                                                1116: 'Eólica',
                                                1117: 'Solar',
                                                1118: 'Total Ger',
                                                1119: 'Carga',
                                                })
        valor10_re1 = valor10_re.rename(index={'Unnamed: 12': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        vari1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_SUL.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]
        dado7 = vari1.iloc[0, 5]

        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7)
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_SUL.xlsx',
            index=False)
        valor9 = ipdo.iloc[1089:1094, 15:17]
        colunassh = [0]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1090: 'Interc N',
                                                1091: 'Interc NE',
                                                1092: 'Interc SE',
                                                1093: 'Interc S',
                                                })
        valor10_re1 = valor10_re.rename(index={'Unnamed: 16': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        vari1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_INTERC.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]

        add_inf6.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            )
        add_inf6 = add_inf6.rename(columns={'Unnamed: 0': ' '})
        add_inf6.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_INTERC.xlsx',
            index=False)
        valor9 = ipdo.iloc[1098:1101, 15:17]
        colunassh = [0]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1099: 'Elo 50 Hz',
                                                1100: 'Itaipu 60 Hz',
                                                })
        valor10_re1 = valor10_re.rename(index={'Unnamed: 16': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        vari1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_ITAIPU.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]

        add_inf7.loc[no] = (dado1,
                            dado2,
                            dado3)
        add_inf7 = add_inf7.rename(columns={'Unnamed: 0': ' '})
        add_inf7.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_ITAIPU.xlsx',
            index=False)
        valor9 = ipdo.iloc[1105:1108, 15:17]
        colunassh = [0]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1106: 'Termo Nuc',
                                                1107: 'Termo Conv',
                                                })
        valor10_re1 = valor10_re.rename(index={'Unnamed: 16': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        vari1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_TERMOSE.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]

        add_inf8.loc[no] = (dado1,
                            dado2,
                            dado3)
        add_inf8 = add_inf8.rename(columns={'Unnamed: 0': ' '})
        add_inf8.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_TERMOSE.xlsx',
            index=False)
        valor9 = ipdo.iloc[1121:1127, 10:16]

        colunassh = [0, 1, 3]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1121: 'Sul',
                                                1122: 'Sudeste',
                                                1123: 'Norte',
                                                1117: 'Solar',
                                                1124: 'Nordeste',
                                                1126: 'Brasil',
                                                })
        valor10_re1 = valor10_re.rename(index={'Unnamed: 12': data,
                                               'Unnamed: 14': ' ',
                                               'Unnamed: 15': ' '})
        vari12 = valor10_re1.iloc[0:3]
        colunasv = [4]
        vari1 = vari12.drop(vari12.columns[colunasv], axis=1)
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_ATUAL.xlsx')
        no = len(add_inf9) + 1
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]

        add_inf9.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6)
        add_inf9 = add_inf9.rename(columns={'Unnamed: 0': ' '})
        no = len(add_inf9) + 1
        vari11 = vari1
        vari11 = vari11.iloc[1:2]
        dado23 = vari11.iloc[0, 0]
        dado33 = vari11.iloc[0, 1]
        dado43 = vari11.iloc[0, 2]
        dado53 = vari11.iloc[0, 3]
        dado63 = vari11.iloc[0, 4]
        dado1 = (' ')
        add_inf9.loc[no] = (dado1,
                            dado23,
                            dado33,
                            dado43,
                            dado53,
                            dado63)
        add_inf9 = add_inf9.rename(columns={'Unnamed: 0': ' '})
        no = len(add_inf9) + 1
        vari12 = vari1
        vari12 = vari12.iloc[2:3]
        dado24 = vari12.iloc[0, 0]
        dado34 = vari12.iloc[0, 1]
        dado44 = vari12.iloc[0, 2]
        dado54 = vari12.iloc[0, 3]
        dado64 = vari12.iloc[0, 4]
        dado1 = (' ')
        add_inf9.loc[no] = (dado1,
                            dado24,
                            dado34,
                            dado44,
                            dado54,
                            dado64)
        add_inf9 = add_inf9.rename(columns={'Unnamed: 0': ' '})
        add_inf9.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_ATUAL.xlsx',
            index=False)

        valor9 = ipdo.iloc[1121:1127, 17:22]
        colunassh = [0, 1, 3]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1121: 'Sul',
                                                1122: 'Sudeste',
                                                1123: 'Norte',
                                                1117: 'Solar',
                                                1124: 'Nordeste ',
                                                1126: 'Brasil',
                                                })
        valor10_re1 = valor10_re.rename(index={'Unnamed: 12': data,
                                               'Unnamed: 19': ' ',
                                               'Unnamed: 21': ' '})
        vari12 = valor10_re1.iloc[0:3]
        colunasv = [4]
        vari1 = vari12.drop(vari12.columns[colunasv], axis=1)
        vari1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_histórico.xlsx')
        dado2 = vari1.iloc[0, 0]
        dado3 = vari1.iloc[0, 1]
        dado4 = vari1.iloc[0, 2]
        dado5 = vari1.iloc[0, 3]
        dado6 = vari1.iloc[0, 4]

        dado1 = data
        add_inf10.loc[no] = (dado1,
                             dado2,
                             dado3,
                             dado4,
                             dado5,
                             dado6
                             )
        add_inf10 = add_inf10.rename(columns={'Unnamed: 0': ' '})
        no = len(add_inf10) + 1
        vari11 = vari1
        vari11 = vari11.iloc[1:2]
        dado23 = vari11.iloc[0, 0]
        dado33 = vari11.iloc[0, 1]
        dado43 = vari11.iloc[0, 2]
        dado53 = vari11.iloc[0, 3]
        dado63 = vari11.iloc[0, 4]
        dado1 = (' ')
        add_inf10.loc[no] = (dado1,
                             dado23,
                             dado33,
                             dado43,
                             dado53,
                             dado63)
        add_inf10 = add_inf10.rename(columns={'Unnamed: 0': ' '})
        add_inf10.to_excel(
           r'C:\Users\e806128\Desktop\1\demMa_histórico.xlsx',
           index=False)

    else:
        print("O arquivo não existe.")
        valor9 = ipdo.iloc[1080:1090, 10:13]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1081: 'Hidro Nac',
                                                1082: 'Itaip',
                                                1083: 'Termo Nuc',
                                                1084: 'Termo Conv',
                                                1085: 'Eólica ',
                                                1086: 'Solar',
                                                1087: ' Total SIN',
                                                1088: 'Interc. Inter',
                                                1089: 'Carga'})
        valor10_re1 = valor10_re.rename(index={'Unnamed: 12': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        valor10_1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        valor10_1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_sin.xlsx')

        valor9 = ipdo.iloc[1091:1098, 10:13]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={
                                                1092: 'Hidro',
                                                1093: 'Termo',
                                                1094: 'Eólica',
                                                1095: 'Solar',
                                                1096: 'Total Ger ',
                                                1097: 'Carga',
                                                1087: ' Total SIN',
                                                1088: 'Interc. Inter',
                                                1089: 'Carga'})
        valor10_re1 = valor10_re.rename(index={'Unnamed: 12': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        valor10_1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        valor10_1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_norte.xlsx')
        valor9 = ipdo.iloc[1098:1105, 10:13]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1099: 'Hidro',
                                                1100: 'Termo',
                                                1101: 'Eólica',
                                                1102: 'Solar',
                                                1103: 'Total Ger',
                                                1104: 'Carga',
                                                1087: ' Total SIN',
                                                1088: 'Interc. Inter',
                                                1089: 'Carga'})
        valor10_re1 = valor10_re.rename(index={'Unnamed: 12': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        valor10_1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        valor10_1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_nordeste.xlsx')
        valor9 = ipdo.iloc[1106:1113, 10:13]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1107: 'Hidro',
                                                1108: 'Termo',
                                                1109: 'Eólica',
                                                1110: 'Solar',
                                                1111: 'Total Ger',
                                                1112: 'Carga',
                                                1087: ' Total SIN',
                                                1088: 'Interc. Inter',
                                                1089: 'Carga'})
        valor10_re1 = valor10_re.rename(index={'Unnamed: 12': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        valor10_1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        valor10_1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_SECO.xlsx')
        valor9 = ipdo.iloc[1113:1120, 10:13]
        colunassh = [0, 1]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1114: 'Hidro',
                                                1115: 'Termo',
                                                1116: 'Eólica',
                                                1117: 'Solar',
                                                1118: 'Total Ger',
                                                1119: 'Carga',
                                                1087: ' Total SIN',
                                                1088: 'Interc. Inter',
                                                1089: 'Carga'})
        valor10_re1 = valor10_re.rename(index={'Unnamed: 12': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        valor10_1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        valor10_1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_SUL.xlsx')
        valor9 = ipdo.iloc[1089:1094, 15:17]
        colunassh = [0]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1090: 'Interc N',
                                                1091: 'Interc NE',
                                                1092: 'Interc SE',
                                                1093: 'Interc S',
                                                1118: 'Total Ger ',
                                                1119: 'Carga',
                                                1087: ' Total SIN',
                                                1088: 'Interc. Inter',
                                                1089: 'Carga'})
        valor10_re1 = valor10_re.rename(index={'Unnamed: 16': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        valor10_1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        valor10_1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_INTERC.xlsx')

        valor9 = ipdo.iloc[1098:1101, 15:17]
        colunassh = [0]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1099: 'Elo 50 Hz',
                                                1100: 'Itaipu 60 Hz',
                                                1092: 'Interc SE',
                                                1093: 'Interc S',
                                                1118: 'Total Ger ',
                                                1119: 'Carga',
                                                1087: ' Total SIN',
                                                1088: 'Interc. Inter',
                                                1089: 'Carga'})
        valor10_re1 = valor10_re.rename(index={'Unnamed: 16': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        valor10_1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        valor10_1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_ITAIPU.xlsx')
        valor9 = ipdo.iloc[1105:1108, 15:17]
        colunassh = [0]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1106: 'Termo Nuc',
                                                1107: 'Termo Conv',
                                                1092: 'Interc SE',
                                                1093: 'Interc S',
                                                1118: 'Total Ger ',
                                                1119: 'Carga',
                                                1087: ' Total SIN',
                                                1088: 'Interc. Inter',
                                                1089: 'Carga'})
        valor10_re1 = valor10_re.rename(index={'Unnamed: 16': data})
        valor10_v = valor10_re1.iloc[0:1]
        colunasv = [0]
        valor10_1 = valor10_v.drop(valor10_v.columns[colunasv], axis=1)
        valor10_1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_TERMOSE.xlsx')
        valor9 = ipdo.iloc[1121:1127, 10:16]
        colunassh = [0, 1, 3]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1121: 'Sul',
                                                1122: 'Sudeste',
                                                1123: 'Norte',
                                                1117: 'Solar',
                                                1124: 'Nordeste',
                                                1126: 'Brasil',
                                                1087: ' Total SIN',
                                                1088: 'Interc. Inter',
                                                1089: 'Carga'})
        valor10_re1 = valor10_re.rename(index={'Unnamed: 12': data,
                                               'Unnamed: 14': ' ',
                                               'Unnamed: 15': ' '})
        valor10_12 = valor10_re1.iloc[0:3]
        colunasv = [4]
        valor10_1 = valor10_12.drop(valor10_12.columns[colunasv], axis=1)
        valor10_1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_ATUAL.xlsx')

        valor9 = ipdo.iloc[1121:1127, 17:22]
        colunassh = [0, 1, 3]
        valor10_d = valor9.drop(valor9.columns[colunassh], axis=1)
        valor10_d1 = valor10_d.transpose()
        valor10_re = valor10_d1.rename(columns={1121: 'Sul',
                                                1122: 'Sudeste',
                                                1123: 'Norte',
                                                1117: 'Solar',
                                                1124: 'Nordeste ',
                                                1126: 'Brasil',
                                                1087: ' Total SIN',
                                                1088: 'Interc. Inter',
                                                1089: 'Carga'})
        valor10_re1 = valor10_re.rename(index={'Unnamed: 12': data,
                                               'Unnamed: 19': data,
                                               'Unnamed: 21': ' '})
        valor10_12 = valor10_re1.iloc[0:3]
        colunasv = [4]
        valor10_1 = valor10_12.drop(valor10_12.columns[colunasv], axis=1)
        valor10_1.to_excel(
            r'C:\Users\e806128\Desktop\1\demMa_histórico.xlsx')


def submercado(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.
    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores da mé-
    dia diária do submercado.
    -------
    """
    dado1 = data
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\Sub_verificadodia.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Sub_verificadodia.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Máxima_Histórica.xlsx')
        no = no + 3
        carga = ipdo.iloc[1120:1125, 1:4]
        colunassc = [1, 2, 4, 3, 6]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1120: 'Sul',
                                           1121: 'Sudeste CO',
                                           1122: 'Norte',
                                           1123: 'NORDESTE',
                                           1124: 'SIN'})
        n_index1 = n_index.rename(index={'Unnamed: 3': data})
        form_tab = n_index1.iloc[1:2]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\Sub_verificadodia.xlsx')

        dado23 = form_tab.iloc[0, 0]
        dado33 = form_tab.iloc[0, 1]
        dado43 = form_tab.iloc[0, 2]
        dado53 = form_tab.iloc[0, 3]
        dado63 = form_tab.iloc[0, 4]

        add_inf.loc[no] = (dado1,
                           dado23,
                           dado33,
                           dado43,
                           dado53,
                           dado63)
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
           r'C:\Users\e806128\Desktop\1\Sub_verificadodia.xlsx',
           index=False)

        carga2 = ipdo.iloc[1120:1125, 1:9]
        colunassc = [1, 2, 4, 3, 6]
        carga3 = carga2.drop(carga2.columns[colunassc], axis=1)
        carga3_f = carga3.transpose()
        carga2_re = carga3_f.rename(columns={1120: 'Sul',
                                             1121: 'Sudeste CO',
                                             1122: 'Norte',
                                             1123: 'NORDESTE',
                                             1124: 'SIN'})
        carga2_re1 = carga2_re.rename(index={
            'Unnamed: 6': data,
            'Unnamed: 8': 'Data Verificação'})
        form_tab = carga2_re1.iloc[1:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\Máxima_Histórica.xlsx')
        dado23 = form_tab.iloc[0, 0]
        dado33 = form_tab.iloc[0, 1]
        dado43 = form_tab.iloc[0, 2]
        dado53 = form_tab.iloc[0, 3]
        dado63 = form_tab.iloc[0, 4]

        add_inf2.loc[no] = (dado1,
                            dado23,
                            dado33,
                            dado43,
                            dado53,
                            dado63)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        no = len(add_inf) + 1
        vari12 = form_tab
        vari12 = vari12.iloc[1:2]
        dado24 = vari12.iloc[0, 0]
        dado34 = vari12.iloc[0, 1]
        dado44 = vari12.iloc[0, 2]
        dado54 = vari12.iloc[0, 3]
        dado64 = vari12.iloc[0, 4]
        dado1 = (' ')
        add_inf2.loc[no] = (dado1,
                            dado24,
                            dado34,
                            dado44,
                            dado54,
                            dado64)
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
           r'C:\Users\e806128\Desktop\1\Máxima_Histórica.xlsx',
           index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1120:1125, 1:4]
        colunassc = [1, 2, 4, 3, 6]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1120: 'Sul',
                                           1121: 'Sudeste CO',
                                           1122: 'Norte',
                                           1123: 'NORDESTE',
                                           1124: 'SIN'})
        n_index1 = n_index.rename(index={'Unnamed: 3': data})
        form_tab = n_index1.iloc[1:2]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\Sub_verificadodia.xlsx')

        carga2 = ipdo.iloc[1120:1125, 1:9]
        colunassc = [1, 2, 4, 3, 6]
        carga3 = carga2.drop(carga2.columns[colunassc], axis=1)
        carga3_f = carga3.transpose()
        carga2_re = carga3_f.rename(columns={1120: 'Sul',
                                             1121: 'Sudeste CO',
                                             1122: 'Norte',
                                             1123: 'NORDESTE',
                                             1124: 'SIN'})
        carga2_re1 = carga2_re.rename(index={
            'Unnamed: 6': data,
            'Unnamed: 8': 'Data Verificação'})
        carga2_v = carga2_re1.iloc[1:3]

        carga2_v.to_excel(
            r'C:\Users\e806128\Desktop\1\Máxima_Histórica.xlsx')


def dadoshidraulicosSINRIOCoRUmba(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores dos da
    dos hidraulicos do rio Rio Corumba.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_vert.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_Aflu.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_defl.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_N.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumban_v.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_vert.xlsx')
        carga = ipdo.iloc[1130:1133, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1130: 'Corumbá IV',
                                           1131: 'Corumbá III',
                                           1132: 'Corumbá I',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_Aflu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_Aflu.xlsx',
            index=False)

        carga = ipdo.iloc[1130:1133, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1130: 'Corumbá IV',
                                           1131: 'Corumbá III ',
                                           1132: 'Corumbá I',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 16': data})
        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_defl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_defl.xlsx',
            index=False)
        carga = ipdo.iloc[1130:1133, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1130: 'Corumbá IV',
                                           1131: 'Corumbá III ',
                                           1132: 'Corumbá I',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_N.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_N.xlsx',
            index=False)

        carga = ipdo.iloc[1130:1133, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1130: 'Corumbá IV',
                                           1131: 'Corumbá III ',
                                           1132: 'Corumbá I',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 19': data})
        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumban_v.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumban_v.xlsx',
            index=False)

        carga = ipdo.iloc[1130:1133, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1130: 'Corumbá IV',
                                           1131: 'Corumbá III ',
                                           1132: 'Corumbá I',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 21': data})
        form_tab = n_index1.iloc[8:9]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_vert.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_vert.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1130:1133, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1130: 'Corumbá IV',
                                           1131: 'Corumbá III',
                                           1132: 'Corumbá I',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_Aflu.xlsx')

        carga = ipdo.iloc[1130:1133, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1130: 'Corumbá IV',
                                           1131: 'Corumbá III ',
                                           1132: 'Corumbá I',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 16': data})
        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_defl.xlsx')
        carga = ipdo.iloc[1130:1133, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1130: 'Corumbá IV',
                                           1131: 'Corumbá III ',
                                           1132: 'Corumbá I',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_N.xlsx')

        carga = ipdo.iloc[1130:1133, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1130: 'Corumbá IV',
                                           1131: 'Corumbá III ',
                                           1132: 'Corumbá I',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 19': data})
        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumban_v.xlsx')
        carga = ipdo.iloc[1130:1133, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1130: 'Corumbá IV',
                                           1131: 'Corumbá III ',
                                           1132: 'Corumbá I',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 21': data})
        form_tab = n_index1.iloc[8:9]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHiSinRCorumba_vert.xlsx')


def DHSRAraguai(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores dos da
    dos hidraulicos do rio Rio Araguai
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Volume.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Def.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Volume.xlsx')
        carga = ipdo.iloc[1133:1137, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1133: 'Nova Ponte',
                                           1134: 'Miranda',
                                           1135: 'Amador Aguiar 1',
                                           1136: 'Amador Aguiar 2',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})
        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Afl.xlsx',
            index=False)
        carga = ipdo.iloc[1133:1137, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1133: 'Nova Ponte',
                                           1134: 'Miranda ',
                                           1135: 'Amador Aguiar 1',
                                           1136: 'Amador Aguiar 2',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Def.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Def.xlsx',
            index=False)

        carga = ipdo.iloc[1133:1137, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                            1133: 'Nova Ponte',
                                            1134: 'Miranda ',
                                            1135: 'Amador Aguiar 1',
                                            1136: 'Amador Aguiar 2',
                                            })
        n_index1 = n_index.rename(index={'Unnamed: 17': data})
        form_tab = n_index1.iloc[4:5]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Nível.xlsx',
            index=False)

        carga = ipdo.iloc[1133:1137, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1133: 'Nova Ponte',
                                           1134: 'Miranda ',
                                           1135: 'Amador Aguiar 1',
                                           1136: 'Amador Aguiar 2',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Volume.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Volume.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1133:1137, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1133: 'Nova Ponte',
                                           1134: 'Miranda',
                                           1135: 'Amador Aguiar 1',
                                           1136: 'Amador Aguiar 2',
                                           1124: 'SIN'})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})
        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Afl.xlsx')
        carga = ipdo.iloc[1133:1137, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1133: 'Nova Ponte',
                                           1134: 'Miranda ',
                                           1135: 'Amador Aguiar 1',
                                           1136: 'Amador Aguiar 2',
                                           1124: 'SIN'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Def.xlsx')

        carga = ipdo.iloc[1133:1137, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
            1133: 'Nova Ponte',
            1134: 'Miranda ',
            1135: 'Amador Aguiar 1',
            1136: 'Amador Aguiar 2',
            1124: 'SIN'})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Nível.xlsx')

        carga = ipdo.iloc[1133:1137, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1133: 'Nova Ponte',
                                           1134: 'Miranda ',
                                           1135: 'Amador Aguiar 1',
                                           1136: 'Amador Aguiar 2',
                                           1124: 'SIN'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaDHSRAraguai_Volume.xlsx')


def DHSNRSMarcos(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores dos da
    dos hidraulicos do rio Rio Marcos.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcos_V.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcs_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcs_Def.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcs_N.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcos_V.xlsx')
        carga = ipdo.iloc[1137:1139, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1137: 'Batalha',
                                    1138: 'S. do Facão',
                                    })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcs_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcs_Afl.xlsx',
            index=False)

        carga = ipdo.iloc[1137:1139, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                        1137: 'Batalha',
                                        1138: 'S. do Facão',
                                        })
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcs_Def.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcs_Def.xlsx',
            index=False)

        carga = ipdo.iloc[1137:1139, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                        1137: 'Batalha',
                                        1138: 'S. do Facão',
                                        })
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcs_N.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcs_N.xlsx',
            index=False)

        carga = ipdo.iloc[1137:1139, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                        1137: 'Batalha',
                                        1138: 'S. do Facão',
                                        })
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcos_V.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcos_V.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1137:1139, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1137: 'Batalha',
                                    1138: 'S. do Facão'
                                    })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcs_Afl.xlsx')

        carga = ipdo.iloc[1137:1139, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                        1137: 'Batalha',
                                        1138: 'S. do Facão',
                                        })
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcs_Def.xlsx')

        carga = ipdo.iloc[1137:1139, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                        1137: 'Batalha',
                                        1138: 'S. do Facão',
                                        })
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcs_N.xlsx')

        carga = ipdo.iloc[1137:1139, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                        1137: 'Batalha',
                                        1138: 'S. do Facão',
                                        })
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRios.Marcos_V.xlsx')


def DHSRParanaiba(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores dos da
    dos hidraulicos do rio Rio Paranaiba.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DadHSinRParanaiba_V.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaiba_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaib_Defl.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaiba_N.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaiba_V.xlsx')
        carga = ipdo.iloc[1139:1143, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1139: 'Theodomiro C. Santiago',
                                           1140: 'Itumbiara ',
                                           1141: 'C. Dourada',
                                           1142: 'São Simão',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})
        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaiba_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaiba_Afl.xlsx',
            index=False)

        carga = ipdo.iloc[1139:1143, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1139: 'Theodomiro C. Santiago',
                                    1140: 'Itumbiara ',
                                    1141: 'C. Dourada',
                                    1142: 'São Simão',
                                    })
        n_index1 = n_index.rename(index={'Unnamed: 16': data})
        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaib_Defl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaib_Defl.xlsx',
            index=False)
        carga = ipdo.iloc[1139:1143, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1139: 'Theodomiro C. Santiago',
                                    1140: 'Itumbiara ',
                                    1141: 'C. Dourada',
                                    1142: 'São Simão',
                                    })
        n_index1 = n_index.rename(index={'Unnamed: 17': data})
        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaiba_N.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaiba_N.xlsx',
            index=False)
        carga = ipdo.iloc[1139:1143, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1139: 'Theodomiro C. Santiago',
                                    1140: 'Itumbiara ',
                                    1141: 'C. Dourada',
                                    1142: 'São Simão',
                                    })
        n_index1 = n_index.rename(index={'Unnamed: 19': data})
        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaiba_V.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaiba_V.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1139:1143, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1139: 'Theodomiro C. Santiago',
                                           1140: 'Itumbiara ',
                                           1141: 'C. Dourada',
                                           1142: 'São Simão',
                                           1124: 'SIN'})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaiba_Afl.xlsx')

        carga = ipdo.iloc[1139:1143, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1139: 'Theodomiro C. Santiago',
                                    1140: 'Itumbiara ',
                                    1141: 'C. Dourada',
                                    1142: 'São Simão',
                                    1124: 'SIN'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaib_Defl.xlsx')

        carga = ipdo.iloc[1139:1143, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1139: 'Theodomiro C. Santiago',
                                    1140: 'Itumbiara ',
                                    1141: 'C. Dourada',
                                    1142: 'São Simão',
                                    1124: 'SIN'})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaiba_N.xlsx')

        carga = ipdo.iloc[1139:1143, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1139: 'Theodomiro C. Santiago',
                                    1140: 'Itumbiara ',
                                    1141: 'C. Dourada',
                                    1142: 'São Simão',
                                    1124: 'SIN'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRParanaiba_V.xlsx')


def DHSRPardo(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores dos da
    dos hidraulicos do rio Rio Pardo.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Volume.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Afluen.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Defluê.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Volume.xlsx')
        carga = ipdo.iloc[1143:1146, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1143: 'Caconde',
                                    1144: 'E. Cunha ',
                                    1145: 'A. S. Oliveira',
                                    })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Afluen.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Afluen.xlsx',
            index=False)

        carga = ipdo.iloc[1143:1146, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1143: 'Caconde',
                                    1144: 'E. Cunha ',
                                    1145: 'A. S. Oliveira',
                                    })
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Defluê.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Defluê.xlsx',
            index=False)

        carga = ipdo.iloc[1143:1146, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1143: 'Caconde',
                                    1144: 'E. Cunha ',
                                    1145: 'A. S. Oliveira',
                                    })
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1143:1146, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1143: 'Caconde',
                                    1144: 'E. Cunha ',
                                    1145: 'A. S. Oliveira',
                                    })
        n_index1 = n_index.rename(index={'Unnamed: 19': data})
        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Volume.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Volume.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1143:1146, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1143: 'Caconde',
                                    1144: 'E. Cunha ',
                                    1145: 'A. S. Oliveira',
                                    1142: 'São Simão',
                                    1124: 'SIN'})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Afluen.xlsx')

        carga = ipdo.iloc[1143:1146, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1143: 'Caconde',
                                    1144: 'E. Cunha ',
                                    1145: 'A. S. Oliveira',
                                    1142: 'São Simão',
                                    1124: 'SIN'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Defluê.xlsx')

        carga = ipdo.iloc[1143:1146, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1143: 'Caconde',
                                    1144: 'E. Cunha ',
                                    1145: 'A. S. Oliveira',
                                    1142: 'São Simão',
                                    1124: 'SIN'})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Nível.xlsx')

        carga = ipdo.iloc[1143:1146, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1143: 'Caconde',
                                    1144: 'E. Cunha ',
                                    1145: 'A. S. Oliveira',
                                    1142: 'São Simão',
                                    1124: 'SIN'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRPardo_Volume.xlsx')


def DHriogrande(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores dos da
    dos hidraulicos do rio Rio Grande.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DadHSinRioG_Volume.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioG_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioG_Defluê.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioG_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioG_Volume.xlsx')
        no = no + 3
        carga = ipdo.iloc[1146:1158, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1146: 'Camargos',
                                           1147: 'Itutinga ',
                                           1148: 'Funil Grande',
                                           1149: 'Furnas',
                                           1150: 'M. Moraes',
                                           1151: 'L. C. Barreto ',
                                           1152: 'Jaguara ',
                                           1153: 'Igarapava ',
                                           1154: 'V. Grande ',
                                           1155: 'P. Colômbia ',
                                           1156:  'Marimbondo ',
                                           1157: 'A. Vermelha '})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})
        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioG_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        dado11 = form_tab.iloc[0, 9]
        dado12 = form_tab.iloc[0, 10]
        dado13 = form_tab.iloc[0, 11]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           dado9,
                           dado10,
                           dado11,
                           dado12,
                           dado13
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioG_Afluença.xlsx',
            index=False)

        carga = ipdo.iloc[1146:1158, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1146: 'Camargos',
                                           1147: 'Itutinga ',
                                           1148: 'Funil Grande',
                                           1149: 'Furnas',
                                           1150: 'M. Moraes',
                                           1151: 'L. C. Barreto ',
                                           1152: 'Jaguara ',
                                           1153: 'Igarapava ',
                                           1154: 'V. Grande ',
                                           1155: 'P. Colômbia ',
                                           1156:  'Marimbondo ',
                                           1157: 'A. Vermelha '})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})
        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioG_Defluê.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        dado11 = form_tab.iloc[0, 9]
        dado12 = form_tab.iloc[0, 10]
        dado13 = form_tab.iloc[0, 11]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9,
                            dado10,
                            dado11,
                            dado12,
                            dado13
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioG_Defluê.xlsx',
            index=False)
        carga = ipdo.iloc[1146:1158, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1146: 'Camargos',
                                           1147: 'Itutinga ',
                                           1148: 'Funil Grande',
                                           1149: 'Furnas',
                                           1150: 'M. Moraes',
                                           1151: 'L. C. Barreto ',
                                           1152: 'Jaguara ',
                                           1153: 'Igarapava ',
                                           1154: 'V. Grande ',
                                           1155: 'P. Colômbia ',
                                           1156:  'Marimbondo ',
                                           1157: 'A. Vermelha '})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})
        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioG_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        dado11 = form_tab.iloc[0, 9]
        dado12 = form_tab.iloc[0, 10]
        dado13 = form_tab.iloc[0, 11]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9,
                            dado10,
                            dado11,
                            dado12,
                            dado13
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioG_Nível.xlsx',
            index=False)

        carga = ipdo.iloc[1146:1158, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1146: 'Camargos',
                                           1147: 'Itutinga ',
                                           1148: 'Funil Grande',
                                           1149: 'Furnas',
                                           1150: 'M. Moraes',
                                           1151: 'L. C. Barreto ',
                                           1152: 'Jaguara ',
                                           1153: 'Igarapava ',
                                           1154: 'V. Grande ',
                                           1155: 'P. Colômbia ',
                                           1156:  'Marimbondo ',
                                           1157: 'A. Vermelha '})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioG_Volume.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        dado11 = form_tab.iloc[0, 9]
        dado12 = form_tab.iloc[0, 10]
        dado13 = form_tab.iloc[0, 11]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9,
                            dado10,
                            dado11,
                            dado12,
                            dado13
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioG_Volume.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1146:1158, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1146: 'Camargos',
                                           1147: 'Itutinga ',
                                           1148: 'Funil Grande',
                                           1149: 'Furnas',
                                           1150: 'M. Moraes',
                                           1151: 'L. C. Barreto ',
                                           1152: 'Jaguara ',
                                           1153: 'Igarapava ',
                                           1154: 'V. Grande ',
                                           1155: 'P. Colômbia ',
                                           1156:  'Marimbondo ',
                                           1157: 'A. Vermelha '})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioG_Afluença.xlsx')

        carga = ipdo.iloc[1146:1158, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1146: 'Camargos',
                                           1147: 'Itutinga ',
                                           1148: 'Funil Grande',
                                           1149: 'Furnas',
                                           1150: 'M. Moraes',
                                           1151: 'L. C. Barreto ',
                                           1152: 'Jaguara ',
                                           1153: 'Igarapava ',
                                           1154: 'V. Grande ',
                                           1155: 'P. Colômbia ',
                                           1156:  'Marimbondo ',
                                           1157: 'A. Vermelha '})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioG_Defluê.xlsx')

        carga = ipdo.iloc[1146:1158, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1146: 'Camargos',
                                           1147: 'Itutinga ',
                                           1148: 'Funil Grande',
                                           1149: 'Furnas',
                                           1150: 'M. Moraes',
                                           1151: 'L. C. Barreto ',
                                           1152: 'Jaguara ',
                                           1153: 'Igarapava ',
                                           1154: 'V. Grande ',
                                           1155: 'P. Colômbia ',
                                           1156:  'Marimbondo ',
                                           1157: 'A. Vermelha '})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]

        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DadHSinRioG_Nível.xlsx')

        carga = ipdo.iloc[1146:1158, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1146: 'Camargos',
                                           1147: 'Itutinga ',
                                           1148: 'Funil Grande',
                                           1149: 'Furnas',
                                           1150: 'M. Moraes',
                                           1151: 'L. C. Barreto ',
                                           1152: 'Jaguara ',
                                           1153: 'Igarapava ',
                                           1154: 'V. Grande ',
                                           1155: 'P. Colômbia ',
                                           1156:  'Marimbondo ',
                                           1157: 'A. Vermelha '})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioG_Volume.xlsx')


def DHSRVerde(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores dos da
    dos hidraulicos do rio Rio Verde.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DadHSinRioVerde_Vo.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioVerde_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioVerde_Def.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioVerde_Vo.xlsx')
        no = no + 3
        carga = ipdo.iloc[1158:1160, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1158: 'Salto Rio Verde',
                                           1159: 'Salto Rio Verdinho ',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})
        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioVerde_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioVerde_Afl.xlsx',
            index=False)

        carga = ipdo.iloc[1158:1160, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1158: 'Salto Rio Verde',
                                           1159: 'Salto Rio Verdinho ', })
        n_index1 = n_index.rename(index={'Unnamed: 16': data})
        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioVerde_Def.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]

        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioVerde_Def.xlsx',
            index=False)

        carga = ipdo.iloc[1158:1160, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1158: 'Salto Rio Verde',
                                           1159: 'Salto Rio Verdinho '})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})
        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioVerde_Vo.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]

        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioVerde_Vo.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1158:1160, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1158: 'Salto Rio Verde',
                                           1159: 'Salto Rio Verdinho ',
                                           1148: 'Funil Grande',
                                           1149: 'Furnas',
                                           1150: 'M. Moraes',
                                           1151: 'L. C. Barreto ',
                                           1152: 'Jaguara ',
                                           1153: 'Igarapava ',
                                           1154: 'V. Grande ',
                                           1155: 'P. Colômbia ',
                                           1156:  'Marimbondo ',
                                           1157: 'A. Vermelha '})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioVerde_Afl.xlsx')

        carga = ipdo.iloc[1158:1160, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1158: 'Salto Rio Verde',
                                           1159: 'Salto Rio Verdinho ', })
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioVerde_Def.xlsx')

        carga = ipdo.iloc[1158:1160, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1158: 'Salto Rio Verde',
                                           1159: 'Salto Rio Verdinho '})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioVerde_Vo.xlsx')


def DHSRclaro(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores dos da
    dos hidraulicos do rio Rio Claro.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_Volum.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_Aflu.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_De.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_Volum.xlsx')
        no = no + 3
        carga = ipdo.iloc[1160:1163, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1160: 'Caçu',
                                    1161: 'Barra dos Coqueiros ',
                                    1162: 'JLMG Pereira'})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_Aflu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_Aflu.xlsx',
            index=False)

        carga = ipdo.iloc[1160:1163, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_ = d_transp.rename(columns={
                                     1160: 'Caçu',
                                     1161: 'Barra dos Coqueiros ',
                                     1162: 'JLMG Pereira'})
        n_index1 = carga1_.rename(index={'Unnamed: 16': data})
        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_De.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_De.xlsx',
            index=False)

        carga = ipdo.iloc[1160:1163, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1160: 'Caçu',
                                    1161: 'Barra dos Coqueiros ',
                                    1162: 'JLMG Pereira'})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_Nível.xlsx',
            index=False)

        carga = ipdo.iloc[1160:1163, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1160: 'Caçu',
                                    1161: 'Barra dos Coqueiros ',
                                    1162: 'JLMG Pereira'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_Volum.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_Volum.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1160:1163, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1160: 'Caçu',
                                    1161: 'Barra dos Coqueiros ',
                                    1162: 'JLMG Pereira'})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_Aflu.xlsx')

        carga = ipdo.iloc[1160:1163, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_ = d_transp.rename(columns={
                                     1160: 'Caçu',
                                     1161: 'Barra dos Coqueiros ',
                                     1162: 'JLMG Pereira'})
        n_index1 = carga1_.rename(index={'Unnamed: 16': data})
        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_De.xlsx')

        carga = ipdo.iloc[1160:1163, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1160: 'Caçu',
                                    1161: 'Barra dos Coqueiros ',
                                    1162: 'JLMG Pereira'})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_Nível.xlsx')

        carga = ipdo.iloc[1160:1163, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1160: 'Caçu',
                                    1161: 'Barra dos Coqueiros ',
                                    1162: 'JLMG Pereira'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioClaro_Volum.xlsx')


def DadosHidraulicosSinRioCorrent(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores dos da
    dos hidraulicos do rio Rio cOrrente.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DadHSinRioCorrente_V.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorrente_Af.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorren_Defl.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorre_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorrente_V.xlsx')
        no = no + 3
        carga = ipdo.iloc[1163:1164, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1163: 'Espora',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})
        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorrente_Af.xlsx')
        dado2 = form_tab.iloc[0, 0]

        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorrente_Af.xlsx',
            index=False)
        carga = ipdo.iloc[1163:1164, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1163: 'Espora', })
        n_index1 = n_index.rename(index={'Unnamed: 16': data})
        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorren_Defl.xlsx')
        dado2 = form_tab.iloc[0, 0]

        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorren_Defl.xlsx',
            index=False)
        carga = ipdo.iloc[1163:1164, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1163: 'Espora', })
        n_index1 = n_index.rename(index={'Unnamed: 17': data})
        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorre_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]

        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorre_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1163:1164, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1163: 'Espora', })
        n_index1 = n_index.rename(index={'Unnamed: 19': data})
        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorrente_V.xlsx')
        dado2 = form_tab.iloc[0, 0]

        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorrente_V.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1163:1164, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1163: 'Espora',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorrente_Af.xlsx')

        carga = ipdo.iloc[1163:1164, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1163: 'Espora', })
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorren_Defl.xlsx')

        carga = ipdo.iloc[1163:1164, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1163: 'Espora', })
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorre_Nível.xlsx')
        carga = ipdo.iloc[1163:1164, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1163: 'Espora', })
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioCorrente_V.xlsx')


def DHSRioPiracicaba(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores dos da
    dos hidraulicos do rio Rio Piracicaba.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DadHSinRioPira_ve.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPirac_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPiraci_Defl.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPirac_Vo.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPira_ve.xlsx')
        no = no + 3
        carga = ipdo.iloc[1164:1166, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1164: 'G. Amorim',
                                           1165: 'Sá Carvalho',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPirac_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPirac_Afl.xlsx',
            index=False)

        carga = ipdo.iloc[1164:1166, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1164: 'G. Amorim',
                                           1165: 'Sá Carvalho'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPiraci_Defl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]

        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPiraci_Defl.xlsx',
            index=False)

        carga = ipdo.iloc[1164:1166, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1164: 'G. Amorim',
                                           1165: 'Sá Carvalho 1148'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPirac_Vo.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]

        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPirac_Vo.xlsx',
            index=False)

        carga = ipdo.iloc[1164:1166, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1164: 'G. Amorim',
                                           1165: 'Sá Carvalho',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 21': data})

        form_tab = n_index1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPira_ve.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]

        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPira_ve.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1164:1166, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1164: 'G. Amorim',
                                           1165: 'Sá Carvalho',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPirac_Afl.xlsx')

        carga = ipdo.iloc[1164:1166, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1164: 'G. Amorim',
                                           1165: 'Sá Carvalho'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPiraci_Defl.xlsx')

        carga = ipdo.iloc[1164:1166, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1164: 'G. Amorim',
                                           1165: 'Sá Carvalho'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPirac_Vo.xlsx')

        carga = ipdo.iloc[1164:1166, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1164: 'G. Amorim',
                                           1165: 'Sá Carvalho',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 21': data})

        form_tab = n_index1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioPira_ve.xlsx')


def DHSRStoantonio(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores dos da
    dos hidraulicos do rio Rio Antonio.
    -------
    """
    caminho_arquivo = (
        r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_vertimento.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_V.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_Ní.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_vertimento.xlsx')
        no = no + 3
        carga = ipdo.iloc[1166:1168, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1166: 'S. Grande',
                                           1167: 'P. Estrela',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})
        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_Afl.xlsx',
            index=False)
        carga = ipdo.iloc[1166:1168, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1166: 'S. Grande',
                                           1167: 'P. Estrela'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})
        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]

        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_Defluência.xlsx',
            index=False)
        carga = ipdo.iloc[1166:1168, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1166: 'S. Grande',
                                           1167: 'P. Estrela'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})
        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_V.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]

        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_V.xlsx',
            index=False)

        carga = ipdo.iloc[1166:1168, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1167: 'P. Estrela',
                                           1166: 'S. Grande '})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})
        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_Ní.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]

        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_Ní.xlsx',
            index=False)
        carga = ipdo.iloc[1166:1168, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1166: 'S. Grande',
                                    1167: 'P. Estrela',
                                    })
        n_index1 = n_index.rename(index={'Unnamed: 21': data})
        form_tab = n_index1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_vertimento.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]

        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_vertimento.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1166:1168, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1166: 'S. Grande',
                                           1167: 'P. Estrela',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_Afl.xlsx')

        carga = ipdo.iloc[1166:1168, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1166: 'S. Grande',
                                           1167: 'P. Estrela'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_Defluência.xlsx')

        carga = ipdo.iloc[1166:1168, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1166: 'S. Grande',
                                           1167: 'P. Estrela'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_V.xlsx')

        carga = ipdo.iloc[1166:1168, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1167: 'P. Estrela',
                                           1166: 'S. Grande '})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_Ní.xlsx')

        carga = ipdo.iloc[1166:1168, 12:22]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1166: 'S. Grande',
                                    1167: 'P. Estrela',
                                    })
        n_index1 = n_index.rename(index={'Unnamed: 21': data})

        form_tab = n_index1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioStoAnto_vertimento.xlsx')


def DadoHidraulicosSinRioDoce(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel, n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores dos da
    dos hidraulicos do rio Rio Doce.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_ve.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_Def.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_Vo.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_ve.xlsx')
        carga = ipdo.iloc[1168:1172, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1168: 'Risoleta Neves',
                                           1169: 'Baguari',
                                           1170: 'Aimorés',
                                           1171: 'Mascarenhas',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_Afl.xlsx',
            index=False)

        carga = ipdo.iloc[1168:1172, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1168: 'Risoleta Neves',
                                    1169: 'Baguari',
                                    1170: 'Aimorés',
                                    1171: 'Mascarenhas'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_Def.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_Def.xlsx',
            index=False)

        carga = ipdo.iloc[1168:1172, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1168: 'Risoleta Neves',
                                    1169: 'Baguari',
                                    1170: 'Aimorés',
                                    1171: 'Mascarenhas'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_Vo.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_Vo.xlsx',
            index=False)

        carga = ipdo.iloc[1168:1172, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1168: 'Risoleta Neves',
                                    1169: 'Baguari',
                                    1170: 'Aimorés',
                                    1171: 'Mascarenhas'})
        n_index1 = n_index.rename(index={'Unnamed: 21': data})

        form_tab = n_index1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_ve.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_ve.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1168:1172, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1168: 'Risoleta Neves',
                                           1169: 'Baguari',
                                           1170: 'Aimorés',
                                           1171: 'Mascarenhas',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_Afl.xlsx')

        carga = ipdo.iloc[1168:1172, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1168: 'Risoleta Neves',
                                    1169: 'Baguari',
                                    1170: 'Aimorés',
                                    1171: 'Mascarenhas'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_Def.xlsx')

        carga = ipdo.iloc[1168:1172, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1168: 'Risoleta Neves',
                                    1169: 'Baguari',
                                    1170: 'Aimorés',
                                    1171: 'Mascarenhas'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_Vo.xlsx')

        carga = ipdo.iloc[1168:1172, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={
                                    1168: 'Risoleta Neves',
                                    1169: 'Baguari',
                                    1170: 'Aimorés',
                                    1171: 'Mascarenhas'})
        n_index1 = n_index.rename(index={'Unnamed: 21': data})

        form_tab = n_index1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioDoce_ve.xlsx')


def DHRPinheiros(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores dos da
    dos hidraulicos do rio Rio Pinheiros.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_Vo.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_Af.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_De.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_N.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_Vo.xlsx')
        no = no + 3
        carga = ipdo.iloc[1178:1179, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1178: 'Billings'})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})
        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_Af.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_Af.xlsx',
            index=False)

        carga = ipdo.iloc[1178:1179, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1178: 'Billings'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_De.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_De.xlsx',
            index=False)
        carga = ipdo.iloc[1178:1179, 12:18]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1178: 'Billings '})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})
        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_N.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_N.xlsx',
            index=False)
        carga = ipdo.iloc[1178:1179, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1178: 'Billings',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 19': data})
        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_Vo.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_Vo.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1178:1179, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1178: 'Billings'})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})
        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_Af.xlsx')

        carga = ipdo.iloc[1178:1179, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1178: 'Billings'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_De.xlsx')

        carga = ipdo.iloc[1178:1179, 12:18]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1178: 'Billings '})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})
        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_N.xlsx')
        carga = ipdo.iloc[1178:1179, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1178: 'Billings',
                                           1165: 'Sá Carvalho 1148'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})
        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioPinhe_Vo.xlsx')


def RioGuarapiranga(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados hidraulicos do rio Rio Guarapiranga.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\RioGuara_Vol.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\RioGuara_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\RioGuara_Deflu.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Rio Guarapiranga.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\RioGuara_Vol.xlsx')
        no = no + 3
        carga = ipdo.iloc[1179:1180, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1179: 'Guarapiranga '})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\RioGuara_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\RioGuara_Afluença.xlsx',
            index=False)

        carga = ipdo.iloc[1179:1180, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1179: 'Guarapiranga'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\RioGuara_Deflu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\RioGuara_Deflu.xlsx',
            index=False)

        carga = ipdo.iloc[1179:1180, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1179: 'Guarapiranga '})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\Rio Guarapiranga.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\Rio Guarapiranga.xlsx',
            index=False)
        carga = ipdo.iloc[1179:1180, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1179: 'Guarapiranga '})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\RioGuara_Vol.xlsx',)
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\RioGuara_Vol.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1179:1180, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1179: 'Guarapiranga '})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\RioGuara_Afluença.xlsx')

        carga = ipdo.iloc[1179:1180, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1179: 'Guarapiranga'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\RioGuara_Deflu.xlsx')

        carga = ipdo.iloc[1179:1180, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1179: 'Guarapiranga '})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\Rio Guarapiranga.xlsx')
        carga = ipdo.iloc[1179:1180, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1179: 'Guarapiranga '})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\RioGuara_Vol.xlsx',)


def dHidraulicosSinRioTiete(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados hidraulicos do rio Rio Tiête.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DadHSinRioTiê_ve.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTi_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTiêt_De.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTiête_N.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioT_V.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTiê_ve.xlsx')
        no = no + 3
        carga = ipdo.iloc[1180:1188, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1180: 'Ponte Nova',
                                           1181: 'E. de Souza',
                                           1182: 'Barra Bonita',
                                           1183: 'Bariri',
                                           1184: 'Ibitinga',
                                           1185: 'Promissão ',
                                           1186: 'N. Avanhadava ',
                                           1187: 'Três Irmãos'})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})
        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTi_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           dado9
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTi_Afl.xlsx',
            index=False)
        carga = ipdo.iloc[1180:1188, 12:17]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1180: 'Ponte Nova',
                                           1181: 'E. de Souza',
                                           1182: 'Barra Bonita',
                                           1183: 'Bariri',
                                           1184: 'Ibitinga',
                                           1185: 'Promissão ',
                                           1186: 'N. Avanhadava ',
                                           1187: 'Três Irmãos'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})
        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTiêt_De.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTiêt_De.xlsx',
            index=False)
        carga = ipdo.iloc[1180:1188, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1180: 'Ponte Nova',
                                           1181: 'E. de Souza',
                                           1182: 'Barra Bonita',
                                           1183: 'Bariri',
                                           1184: 'Ibitinga',
                                           1185: 'Promissão ',
                                           1186: 'N. Avanhadava',
                                           1187: 'Três Irmãos'})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})
        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTiête_N.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTiête_N.xlsx',
            index=False)
        carga = ipdo.iloc[1180:1188, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1180: 'Ponte Nova',
                                           1181: 'E. de Souza',
                                           1182: 'Barra Bonita',
                                           1183: 'Bariri',
                                           1184: 'Ibitinga',
                                           1185: 'Promissão ',
                                           1186: 'N. Avanhadava ',
                                           1187: 'Três Irmãos'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})
        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioT_V.xlsx',)
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioT_V.xlsx',
            index=False)

        carga = ipdo.iloc[1180:1188, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1180: 'Ponte Nova',
                                           1181: 'E. de Souza',
                                           1182: 'Barra Bonita',
                                           1183: 'Bariri',
                                           1184: 'Ibitinga',
                                           1185: 'Promissão ',
                                           1186: 'N. Avanhadava ',
                                           1187: 'Três Irmãos'})
        n_index1 = n_index.rename(index={'Unnamed: 21': data})
        form_tab = n_index1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTiê_ve.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTiê_ve.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1180:1188, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1180: 'Ponte Nova',
                                           1181: 'E. de Souza',
                                           1182: 'Barra Bonita',
                                           1183: 'Bariri',
                                           1184: 'Ibitinga',
                                           1185: 'Promissão ',
                                           1186: 'N. Avanhadava ',
                                           1187: 'Três Irmãos'})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})
        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTi_Afl.xlsx')

        carga = ipdo.iloc[1180:1188, 12:17]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1180: 'Ponte Nova',
                                           1181: 'E. de Souza',
                                           1182: 'Barra Bonita',
                                           1183: 'Bariri',
                                           1184: 'Ibitinga',
                                           1185: 'Promissão ',
                                           1186: 'N. Avanhadava ',
                                           1187: 'Três Irmãos'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTiêt_De.xlsx')

        carga = ipdo.iloc[1180:1188, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1180: 'Ponte Nova',
                                           1181: 'E. de Souza',
                                           1182: 'Barra Bonita',
                                           1183: 'Bariri',
                                           1184: 'Ibitinga',
                                           1185: 'Promissão ',
                                           1186: 'N. Avanhadava',
                                           1187: 'Três Irmãos'})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTiête_N.xlsx')

        carga = ipdo.iloc[1180:1188, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1180: 'Ponte Nova',
                                           1181: 'E. de Souza',
                                           1182: 'Barra Bonita',
                                           1183: 'Bariri',
                                           1184: 'Ibitinga',
                                           1185: 'Promissão ',
                                           1186: 'N. Avanhadava ',
                                           1187: 'Três Irmãos'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioRioT_V.xlsx',)

        carga = ipdo.iloc[1180:1188, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1180: 'Ponte Nova',
                                           1181: 'E. de Souza',
                                           1182: 'Barra Bonita',
                                           1183: 'Bariri',
                                           1184: 'Ibitinga',
                                           1185: 'Promissão ',
                                           1186: 'N. Avanhadava ',
                                           1187: 'Três Irmãos'})
        n_index1 = n_index.rename(index={'Unnamed: 21': data})

        form_tab = n_index1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DadHSinRioTiê_ve.xlsx')


def DHidraulicosSinRioTibagi(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados hidraulicos do rio Rio Tibagi.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DaHSRioTibagi_Vo.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibagi_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibag_Defl.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibagi_N.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibagi_Vo.xlsx')
        no = no + 3
        carga = ipdo.iloc[1188:1189, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1188: 'Gov. J. Canet Jr',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibagi_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibagi_Afl.xlsx',
            index=False)
        carga = ipdo.iloc[1188:1189, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1188: 'Gov. J. Canet Jr'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibag_Defl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibag_Defl.xlsx',
            index=False)

        carga = ipdo.iloc[1188:1189, 12:18]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1188: 'Gov. J. Canet Jr'})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibagi_N.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibagi_N.xlsx',
            index=False)

        carga = ipdo.iloc[1188:1189, 12:22]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1188: 'Gov. J. Canet Jr'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibagi_Vo.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibagi_Vo.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1188:1189, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1188: 'Gov. J. Canet Jr',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibagi_Afl.xlsx')
        carga = ipdo.iloc[1188:1189, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1188: 'Gov. J. Canet Jr'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibag_Defl.xlsx')

        carga = ipdo.iloc[1188:1189, 12:18]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1188: 'Gov. J. Canet Jr'})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibagi_N.xlsx')

        carga = ipdo.iloc[1188:1189, 12:22]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1188: 'Gov. J. Canet Jr'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioTibagi_Vo.xlsx')


def DadosHSinRioParanapanema(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados hidraulicos do rio Rio Paranapanema.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DaHSRioParana.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Def.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParan_Nívl.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParanap_Vo.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana.xlsx')
        no = no + 3
        carga = ipdo.iloc[1189:1199, 12:16]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1189: 'Jurumirim',
                                           1190: 'Piraju',
                                           1191: 'Chavantes',
                                           1192: 'Ourinhos',
                                           1193: 'Salto Grande',
                                           1194: 'Canoas II ',
                                           1195: 'Canoas I',
                                           1196: 'Capivara',
                                           1197: 'Taquaruçu',
                                           1198: 'Rosana'})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})
        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaHSRioParana_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        dado11 = form_tab.iloc[0, 9]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           dado9,
                           dado10,
                           dado11
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Afl.xlsx',
            index=False)
        carga = ipdo.iloc[1189:1199, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1189: 'Jurumirim',
                                           1190: 'Piraju',
                                           1191: 'Chavantes',
                                           1192: 'Ourinhos',
                                           1193: 'Salto Grande',
                                           1194: 'Canoas II ',
                                           1195: 'Canoas I',
                                           1196: 'Capivara',
                                           1197: 'Taquaruçu',
                                           1198: 'Rosana'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})
        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Def.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        dado11 = form_tab.iloc[0, 9]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9,
                            dado10,
                            dado11
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Def.xlsx',
            index=False)
        carga = ipdo.iloc[1189:1199, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1189: 'Jurumirim',
                                           1190: 'Piraju',
                                           1191: 'Chavantes',
                                           1192: 'Ourinhos',
                                           1193: 'Salto Grande',
                                           1194: 'Canoas II ',
                                           1195: 'Canoas I',
                                           1196: 'Capivara',
                                           1197: 'Taquaruçu',
                                           1198: 'Rosana'})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})
        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaHSRioParan_Nívl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        dado11 = form_tab.iloc[0, 9]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9,
                            dado10,
                            dado11
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParan_Nívl.xlsx',
            index=False)
        carga = ipdo.iloc[1189:1199, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1189: 'Jurumirim',
                                           1190: 'Piraju',
                                           1191: 'Chavantes',
                                           1192: 'Ourinhos',
                                           1193: 'Salto Grande',
                                           1194: 'Canoas II ',
                                           1195: 'Canoas I',
                                           1196: 'Capivara',
                                           1197: 'Taquaruçu',
                                           1198: 'Rosana'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})
        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaHSRioParanap_Vo.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        dado11 = form_tab.iloc[0, 9]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9,
                            dado10,
                            dado11
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParanap_Vo.xlsx',
            index=False)
        carga = ipdo.iloc[1189:1199, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1189: 'Jurumirim',
                                           1190: 'Piraju',
                                           1191: 'Chavantes',
                                           1192: 'Ourinhos',
                                           1193: 'Salto Grande',
                                           1194: 'Canoas II ',
                                           1195: 'Canoas I',
                                           1196: 'Capivara',
                                           1197: 'Taquaruçu',
                                           1198: 'Rosana'})
        n_index1 = n_index.rename(index={'Unnamed: 21': data})
        form_tab = n_index1.iloc[8:9]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaHSRioParana.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        dado11 = form_tab.iloc[0, 9]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9,
                            dado10,
                            dado11
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana.xlsx',
            index=False)
    else:
        print("O arquivo não existe")
        carga = ipdo.iloc[1189:1199, 12:16]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1189: 'Jurumirim',
                                           1190: 'Piraju',
                                           1191: 'Chavantes',
                                           1192: 'Ourinhos',
                                           1193: 'Salto Grande',
                                           1194: 'Canoas II ',
                                           1195: 'Canoas I',
                                           1196: 'Capivara',
                                           1197: 'Taquaruçu',
                                           1198: 'Rosana'})
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaHSRioParana_Afl.xlsx')

        carga = ipdo.iloc[1189:1199, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1189: 'Jurumirim',
                                           1190: 'Piraju',
                                           1191: 'Chavantes',
                                           1192: 'Ourinhos',
                                           1193: 'Salto Grande',
                                           1194: 'Canoas II ',
                                           1195: 'Canoas I',
                                           1196: 'Capivara',
                                           1197: 'Taquaruçu',
                                           1198: 'Rosana'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaHSRioParana_Def.xlsx')

        carga = ipdo.iloc[1189:1199, 12:18]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1189: 'Jurumirim',
                                           1190: 'Piraju',
                                           1191: 'Chavantes',
                                           1192: 'Ourinhos',
                                           1193: 'Salto Grande',
                                           1194: 'Canoas II ',
                                           1195: 'Canoas I',
                                           1196: 'Capivara',
                                           1197: 'Taquaruçu',
                                           1198: 'Rosana'})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaHSRioParan_Nívl.xlsx')

        carga = ipdo.iloc[1189:1199, 12:20]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1189: 'Jurumirim',
                                           1190: 'Piraju',
                                           1191: 'Chavantes',
                                           1192: 'Ourinhos',
                                           1193: 'Salto Grande',
                                           1194: 'Canoas II ',
                                           1195: 'Canoas I',
                                           1196: 'Capivara',
                                           1197: 'Taquaruçu',
                                           1198: 'Rosana'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaHSRioParanap_Vo.xlsx')
        carga = ipdo.iloc[1189:1199, 12:22]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1189: 'Jurumirim',
                                           1190: 'Piraju',
                                           1191: 'Chavantes',
                                           1192: 'Ourinhos',
                                           1193: 'Salto Grande',
                                           1194: 'Canoas II ',
                                           1195: 'Canoas I',
                                           1196: 'Capivara',
                                           1197: 'Taquaruçu',
                                           1198: 'Rosana'})
        n_index1 = n_index.rename(index={'Unnamed: 21': data})

        form_tab = n_index1.iloc[8:9]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaHSRioParana.xlsx')


def DHSRParaná(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados hidraulicos do rio Rio Paraná.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DaHSRioParanaeve.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Af=.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_D.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Nív.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Vol.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParanaeve.xlsx')
        no = no + 3
        carga = ipdo.iloc[1199:1203, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1199: 'I. Solteira',
                                           1200: 'Jupiá',
                                           1201: 'P. Primavera',
                                           1202: 'Itaipu',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Af=.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Af=.xlsx',
            index=False)
        carga = ipdo.iloc[1199:1203, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1199: 'I. Solteira',
                                           1200: 'Jupiá',
                                           1201: 'P. Primavera',
                                           1202: 'Itaipu'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})
        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_D.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]

        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_D.xlsx',
            index=False)
        carga = ipdo.iloc[1199:1203, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1199: 'I. Solteira',
                                           1200: 'Jupiá',
                                           1201: 'P. Primavera',
                                           1202: 'Itaipu'})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Nív.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]

        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Nív.xlsx',
            index=False)
        carga = ipdo.iloc[1199:1203, 12:20]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1199: 'I. Solteira',
                                           1200: 'Jupiá',
                                           1201: 'P. Primavera',
                                           1202: 'Itaipu'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})
        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Vol.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]

        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Vol.xlsx',
            index=False)

        carga = ipdo.iloc[1199:1203, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1199: 'I. Solteira',
                                           1200: 'Jupiá',
                                           1201: 'P. Primavera',
                                           1202: 'Itaipu'})
        n_index1 = n_index.rename(index={'Unnamed: 21': data})

        form_tab = n_index1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParanaeve.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]

        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParanaeve.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1199:1203, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1199: 'I. Solteira',
                                           1200: 'Jupiá',
                                           1201: 'P. Primavera',
                                           1202: 'Itaipu',
                                           })
        n_index1 = n_index.rename(index={'Unnamed: 15': data})

        form_tab = n_index1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Af=.xlsx')

        carga = ipdo.iloc[1199:1203, 12:17]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1199: 'I. Solteira',
                                           1200: 'Jupiá',
                                           1201: 'P. Primavera',
                                           1202: 'Itaipu'})
        n_index1 = n_index.rename(index={'Unnamed: 16': data})

        form_tab = n_index1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_D.xlsx')

        carga = ipdo.iloc[1199:1203, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1199: 'I. Solteira',
                                           1200: 'Jupiá',
                                           1201: 'P. Primavera',
                                           1202: 'Itaipu'})
        n_index1 = n_index.rename(index={'Unnamed: 17': data})

        form_tab = n_index1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Nív.xlsx')

        carga = ipdo.iloc[1199:1203, 12:20]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1199: 'I. Solteira',
                                           1200: 'Jupiá',
                                           1201: 'P. Primavera',
                                           1202: 'Itaipu'})
        n_index1 = n_index.rename(index={'Unnamed: 19': data})

        form_tab = n_index1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParana_Vol.xlsx')

        carga = ipdo.iloc[1199:1203, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        n_index = d_transp.rename(columns={1199: 'I. Solteira',
                                           1200: 'Jupiá',
                                           1201: 'P. Primavera',
                                           1202: 'Itaipu'})
        n_index1 = n_index.rename(index={'Unnamed: 21': data})

        form_tab = n_index1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRioParanaeve.xlsx')


def DHSJaguari(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados hidraulicos do rio Rio Jagauri.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Nívl.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Aflu.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Defl.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Nívl.xlsx')
        no = no + 3
        carga = ipdo.iloc[1204:1205, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1204: 'Jaguari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Aflu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Aflu.xlsx',
            index=False)

        carga = ipdo.iloc[1204:1205, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1204: 'Jaguari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Defl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Defl.xlsx',
            index=False)

        carga = ipdo.iloc[1204:1205, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1204: 'Jaguari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1204:1205, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1204: 'Jaguari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Nívl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Nívl.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1204:1205, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1204: 'Jaguari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Aflu.xlsx')

        carga = ipdo.iloc[1204:1205, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1204: 'Jaguari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Defl.xlsx')

        carga = ipdo.iloc[1204:1205, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1204: 'Jaguari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Nível.xlsx')
        carga = ipdo.iloc[1204:1205, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1204: 'Jaguari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoJaguari_Nívl.xlsx')


def DHSRdPeixe(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados hidraulicos do rio Rio Peixe.
    -------
    """
    caminho_arquivo = (
        r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Volume.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Volume.xlsx')
        no = no + 3
        carga = ipdo.iloc[1204:1205, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1204: 'Picada'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Afluença.xlsx',
            index=False)
        carga = ipdo.iloc[1204:1205, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1204: 'Picada'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Defluência.xlsx',
            index=False)

        carga = ipdo.iloc[1204:1205, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1204: 'Picada'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Nível.xlsx', )
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Nível.xlsx',
            index=False)

        carga = ipdo.iloc[1204:1205, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1204: 'Picada'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        form_tab = carga1_re1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Volume.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Volume.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1204:1205, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1204: 'Picada'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Afluença.xlsx')
        carga = ipdo.iloc[1204:1205, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1204: 'Picada'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Defluência.xlsx')

        carga = ipdo.iloc[1204:1205, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1204: 'Picada'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Nível.xlsx', )

        carga = ipdo.iloc[1204:1205, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1204: 'Picada'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRdoPeixe_Volume.xlsx')


def DHSRParaibuna(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados hidraulicos do rio Rio Paraibuna.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DaHSinRioPar_vetim.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioPara_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRoPara_Defl.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioParaib_N.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioParai_Volume.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioPar_vetim.xlsx')
        no = no + 3
        carga = ipdo.iloc[1205:1214, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1205: 'Sobragi',
                                             1206: 'Lajes',
                                             1207: 'Fontes',
                                             1208: 'N.Peçanha',
                                             1209: 'P.Passos',
                                             1210: 'Tocos',
                                             1211: 'Transf. Tocos',
                                             1212: 'Santana',
                                             1213: 'Vigário'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioPara_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           dado9,
                           dado10
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioPara_Afl.xlsx',
            index=False)
        carga = ipdo.iloc[1205:1214, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1205: 'Sobragi',
                                             1206: 'Lajes',
                                             1207: 'Fontes',
                                             1208: 'N.Peçanha',
                                             1209: 'P.Passos',
                                             1210: 'Tocos',
                                             1211: 'Transf. Tocos',
                                             1212: 'Santana',
                                             1213: 'Vigário'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRoPara_Defl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9,
                            dado10
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRoPara_Defl.xlsx',
            index=False)
        carga = ipdo.iloc[1205:1214, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1205: 'Sobragi',
                                             1206: 'Lajes',
                                             1207: 'Fontes',
                                             1208: 'N.Peçanha',
                                             1209: 'P.Passos',
                                             1210: 'Tocos',
                                             1211: 'Transf. Tocos',
                                             1212: 'Santana',
                                             1213: 'Vigário'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioParaib_N.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9,
                            dado10
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioParaib_N.xlsx',
            index=False)
        carga = ipdo.iloc[1205:1214, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1205: 'Sobragi',
                                             1206: 'Lajes',
                                             1207: 'Fontes',
                                             1208: 'N.Peçanha',
                                             1209: 'P.Passos',
                                             1210: 'Tocos',
                                             1211: 'Transf. Tocos',
                                             1212: 'Santana',
                                             1213: 'Vigário'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioParai_Volume.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9,
                            dado10
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioParai_Volume.xlsx',
            index=False)
        carga = ipdo.iloc[1205:1214, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1205: 'Sobragi',
                                             1206: 'Lajes',
                                             1207: 'Fontes',
                                             1208: 'N.Peçanha',
                                             1209: 'P.Passos',
                                             1210: 'Tocos',
                                             1211: 'Transf. Tocos',
                                             1212: 'Santana',
                                             1213: 'Vigário'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioPar_vetim.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9,
                            dado10
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioPar_vetim.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1205:1214, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1205: 'Sobragi',
                                             1206: 'Lajes',
                                             1207: 'Fontes',
                                             1208: 'N.Peçanha',
                                             1209: 'P.Passos',
                                             1210: 'Tocos',
                                             1211: 'Transf. Tocos',
                                             1212: 'Santana',
                                             1213: 'Vigário'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioPara_Afl.xlsx')
        carga = ipdo.iloc[1205:1214, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1205: 'Sobragi',
                                             1206: 'Lajes',
                                             1207: 'Fontes',
                                             1208: 'N.Peçanha',
                                             1209: 'P.Passos',
                                             1210: 'Tocos',
                                             1211: 'Transf. Tocos',
                                             1212: 'Santana',
                                             1213: 'Vigário'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRoPara_Defl.xlsx')
        carga = ipdo.iloc[1205:1214, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1205: 'Sobragi',
                                             1206: 'Lajes',
                                             1207: 'Fontes',
                                             1208: 'N.Peçanha',
                                             1209: 'P.Passos',
                                             1210: 'Tocos',
                                             1211: 'Transf. Tocos',
                                             1212: 'Santana',
                                             1213: 'Vigário'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioParaib_N.xlsx')
        carga = ipdo.iloc[1205:1214, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1205: 'Sobragi',
                                             1206: 'Lajes',
                                             1207: 'Fontes',
                                             1208: 'N.Peçanha',
                                             1209: 'P.Passos',
                                             1210: 'Tocos',
                                             1211: 'Transf. Tocos',
                                             1212: 'Santana',
                                             1213: 'Vigário'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioParai_Volume.xlsx')
        carga = ipdo.iloc[1205:1214, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1205: 'Sobragi',
                                             1206: 'Lajes',
                                             1207: 'Fontes',
                                             1208: 'N.Peçanha',
                                             1209: 'P.Passos',
                                             1210: 'Tocos',
                                             1211: 'Transf. Tocos',
                                             1212: 'Santana',
                                             1213: 'Vigário'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRioPar_vetim.xlsx')


def DHSRParaibadSul(ipdo: pd.DataFrame):
    """
    data,colunassh,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados hidraulicos do rio Rio Paraiba do Sul.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DHSRPSevertimento.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Defluênc.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Volume.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPSevertimento.xlsx')
        no = no + 3
        carga = ipdo.iloc[1214:1222, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1214: 'Paraibuna',
                                             1215: 'S. Branca',
                                             1216: 'Funil',
                                             1217: 'S. Cecília',
                                             1218: 'Transf. Sta Cecília',
                                             1219: 'Anta',
                                             1220: 'Simplício',
                                             1221: 'Ilha dos Pombos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           dado9
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Afluença.xlsx',
            index=False)
        carga = ipdo.iloc[1214:1222, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1214: 'Paraibuna',
                                             1215: 'S. Branca',
                                             1216: 'Funil',
                                             1217: 'S. Cecília',
                                             1218: 'Transf. Sta Cecília',
                                             1219: 'Anta',
                                             1220: 'Simplício',
                                             1221: 'Ilha dos Pombos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Defluênc.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]

        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Defluênc.xlsx',
            index=False)

        carga = ipdo.iloc[1214:1222, 12:18]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1214: 'Paraibuna',
                                             1215: 'S. Branca',
                                             1216: 'Funil',
                                             1217: 'S. Cecília',
                                             1218: 'Transf. Sta Cecília',
                                             1219: 'Anta',
                                             1220: 'Simplício',
                                             1221: 'Ilha dos Pombos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        form_tab = carga1_re1.iloc[4:5]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]

        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Nível.xlsx',
            index=False)

        carga = ipdo.iloc[1214:1222, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1214: 'Paraibuna',
                                             1215: 'S. Branca',
                                             1216: 'Funil',
                                             1217: 'S. Cecília',
                                             1218: 'Transf. Sta Cecília',
                                             1219: 'Anta',
                                             1220: 'Simplício',
                                             1221: 'Ilha dos Pombos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Volume.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]

        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Volume.xlsx',
            index=False)
        carga = ipdo.iloc[1214:1222, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1214: 'Paraibuna',
                                             1215: 'S. Branca',
                                             1216: 'Funil',
                                             1217: 'S. Cecília',
                                             1218: 'Transf. Sta Cecília',
                                             1219: 'Anta',
                                             1220: 'Simplício',
                                             1221: 'Ilha dos Pombos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPSevertimento.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]

        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPSevertimento.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1214:1222, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1214: 'Paraibuna',
                                             1215: 'S. Branca',
                                             1216: 'Funil',
                                             1217: 'S. Cecília',
                                             1218: 'Transf. Sta Cecília',
                                             1219: 'Anta',
                                             1220: 'Simplício',
                                             1221: 'Ilha dos Pombos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Afluença.xlsx')
        carga = ipdo.iloc[1214:1222, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1214: 'Paraibuna',
                                             1215: 'S. Branca',
                                             1216: 'Funil',
                                             1217: 'S. Cecília',
                                             1218: 'Transf. Sta Cecília',
                                             1219: 'Anta',
                                             1220: 'Simplício',
                                             1221: 'Ilha dos Pombos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Defluênc.xlsx')

        carga = ipdo.iloc[1214:1222, 12:18]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1214: 'Paraibuna',
                                             1215: 'S. Branca',
                                             1216: 'Funil',
                                             1217: 'S. Cecília',
                                             1218: 'Transf. Sta Cecília',
                                             1219: 'Anta',
                                             1220: 'Simplício',
                                             1221: 'Ilha dos Pombos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        form_tab = carga1_re1.iloc[4:5]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Nível.xlsx')

        carga = ipdo.iloc[1214:1222, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1214: 'Paraibuna',
                                             1215: 'S. Branca',
                                             1216: 'Funil',
                                             1217: 'S. Cecília',
                                             1218: 'Transf. Sta Cecília',
                                             1219: 'Anta',
                                             1220: 'Simplício',
                                             1221: 'Ilha dos Pombos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPS_Volume.xlsx')
        carga = ipdo.iloc[1214:1222, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1214: 'Paraibuna',
                                             1215: 'S. Branca',
                                             1216: 'Funil',
                                             1217: 'S. Cecília',
                                             1218: 'Transf. Sta Cecília',
                                             1219: 'Anta',
                                             1220: 'Simplício',
                                             1221: 'Ilha dos Pombos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRPSevertimento.xlsx')


def dadosbacias(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados das bacias.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DB_Ger Prog.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Dadosbacias_Armaz.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Dadosbacis_ENAdia.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DB_ENAArmaz.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DB_ENA Bruta.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DB_Ger Verif.xlsx')
        add_inf6 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DB_Ger Prog.xlsx')
        no = no + 3
        carga = ipdo.iloc[1222:1231, 10:13]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1223: 'PNB',
                                             1224: 'GRA',
                                             1225: 'TIE',
                                             1226: 'PRN',
                                             1227: 'PAR',
                                             1228: 'SUL',
                                             1229: 'PRG',
                                             1230: 'DOC'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 12': data})
        form_tab1 = carga1_re1.iloc[0:1]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\Dadosbacias_Armaz.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           dado9
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\Dadosbacias_Armaz.xlsx',
            index=False)
        carga = ipdo.iloc[1222:1231, 10:15]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1223: 'PNB',
                                             1224: 'GRA',
                                             1225: 'TIE',
                                             1226: 'PRN',
                                             1227: 'PAR',
                                             1228: 'SUL',
                                             1229: 'PRG',
                                             1230: 'DOC'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 14': data})
        form_tab1 = carga1_re1.iloc[2:3]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\Dadosbacis_ENAdia.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\Dadosbacis_ENAdia.xlsx',
            index=False)
        carga = ipdo.iloc[1222:1231, 10:16]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1223: 'PNB',
                                             1224: 'GRA',
                                             1225: 'TIE',
                                             1226: 'PRN',
                                             1227: 'PAR',
                                             1228: 'SUL',
                                             1229: 'PRG',
                                             1230: 'DOC'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab1 = carga1_re1.iloc[3:4]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DB_ENAArmaz.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DB_ENAArmaz.xlsx',
            index=False)
        carga = ipdo.iloc[1222:1231, 10:17]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1223: 'PNB',
                                             1224: 'GRA',
                                             1225: 'TIE',
                                             1226: 'PRN',
                                             1227: 'PAR',
                                             1228: 'SUL',
                                             1229: 'PRG',
                                             1230: 'DOC'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab1 = carga1_re1.iloc[4:5]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DB_ENA Bruta.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DB_ENA Bruta.xlsx',
            index=False)

        carga = ipdo.iloc[1222:1231, 10:18]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1223: 'PNB',
                                             1224: 'GRA',
                                             1225: 'TIE',
                                             1226: 'PRN',
                                             1227: 'PAR',
                                             1228: 'SUL',
                                             1229: 'PRG',
                                             1230: 'DOC'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab1 = carga1_re1.iloc[5:6]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)

        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DB_Ger Verif.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DB_Ger Verif.xlsx',
            index=False)
        carga = ipdo.iloc[1222:1231, 10:20]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1223: 'PNB',
                                             1224: 'GRA',
                                             1225: 'TIE',
                                             1226: 'PRN',
                                             1227: 'PAR',
                                             1228: 'SUL',
                                             1229: 'PRG',
                                             1230: 'DOC'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab1 = carga1_re1.iloc[7:8]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)

        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DB_Ger Prog.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        add_inf6.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9
                            )
        add_inf6 = add_inf6.rename(columns={'Unnamed: 0': ' '})
        add_inf6.to_excel(
            r'C:\Users\e806128\Desktop\1\DB_Ger Prog.xlsx',
            index=False)
    else:
        print("O arquivo não existe")
        carga = ipdo.iloc[1222:1231, 10:13]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1223: 'PNB',
                                             1224: 'GRA',
                                             1225: 'TIE',
                                             1226: 'PRN',
                                             1227: 'PAR',
                                             1228: 'SUL',
                                             1229: 'PRG',
                                             1230: 'DOC'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 12': data})

        form_tab1 = carga1_re1.iloc[0:1]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\Dadosbacias_Armaz.xlsx')

        carga = ipdo.iloc[1222:1231, 10:15]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1223: 'PNB',
                                             1224: 'GRA',
                                             1225: 'TIE',
                                             1226: 'PRN',
                                             1227: 'PAR',
                                             1228: 'SUL',
                                             1229: 'PRG',
                                             1230: 'DOC'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 14': data})
        form_tab1 = carga1_re1.iloc[2:3]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)

        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\Dadosbacis_ENAdia.xlsx')

        carga = ipdo.iloc[1222:1231, 10:16]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1223: 'PNB',
                                             1224: 'GRA',
                                             1225: 'TIE',
                                             1226: 'PRN',
                                             1227: 'PAR',
                                             1228: 'SUL',
                                             1229: 'PRG',
                                             1230: 'DOC'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab1 = carga1_re1.iloc[3:4]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)

        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DB_ENAArmaz.xlsx')
        carga = ipdo.iloc[1222:1231, 10:17]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1223: 'PNB',
                                             1224: 'GRA',
                                             1225: 'TIE',
                                             1226: 'PRN',
                                             1227: 'PAR',
                                             1228: 'SUL',
                                             1229: 'PRG',
                                             1230: 'DOC'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab1 = carga1_re1.iloc[4:5]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)

        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DB_ENA Bruta.xlsx')

        carga = ipdo.iloc[1222:1231, 10:18]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1223: 'PNB',
                                             1224: 'GRA',
                                             1225: 'TIE',
                                             1226: 'PRN',
                                             1227: 'PAR',
                                             1228: 'SUL',
                                             1229: 'PRG',
                                             1230: 'DOC'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab1 = carga1_re1.iloc[5:6]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)

        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DB_Ger Verif.xlsx')

        carga = ipdo.iloc[1222:1231, 10:20]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1223: 'PNB',
                                             1224: 'GRA',
                                             1225: 'TIE',
                                             1226: 'PRN',
                                             1227: 'PAR',
                                             1228: 'SUL',
                                             1229: 'PRG',
                                             1230: 'DOC'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab1 = carga1_re1.iloc[7:8]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)

        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DB_Ger Prog.xlsx')


def DHSRioManso(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados hidráulicos do Rio Manso.
    -------
    """
    caminho_arquivo = (
        r'C:\Users\e806128\Desktop\1\DHSRM_volume.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRM_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRM_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRM_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRM_volume.xlsx')

        no = no + 3
        carga = ipdo.iloc[1233:1234, 12:16]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1233: 'Manso'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRM_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRM_Afluença.xlsx',
            index=False)

        carga = ipdo.iloc[1233:1234, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1233: 'Manso'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRM_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRM_Defluência.xlsx',
            index=False)
        carga = ipdo.iloc[1233:1234, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1233: 'Manso'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRM_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRM_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1233:1234, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1233: 'Manso'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRM_volume.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRM_volume.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1233:1234, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1233: 'Manso'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DHSRM_Afluença.xlsx')
        carga = ipdo.iloc[1233:1234, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1233: 'Manso'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DHSRM_Defluência.xlsx')
        carga = ipdo.iloc[1233:1234, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1233: 'Manso'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DHSRM_Nível.xlsx')
        carga = ipdo.iloc[1233:1234, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1233: 'Manso'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DHSRM_volume.xlsx')


def DHSRioItiquira(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos do rio Itiquira.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DaHSinRPE.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Volume.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRPE.xlsx')
        no = no + 3
        carga = ipdo.iloc[1234:1236, 12:16]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1234: 'Itiquira I',
                                             1235: 'Itiquira II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab = carga1_re1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Afluença.xlsx',
            index=False)
        carga = ipdo.iloc[1234:1236, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1234: 'Itiquira I',
                                             1235: 'Itiquira II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado3 = form_tab.iloc[0, 1]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Defluência.xlsx',
            index=False)
        carga = ipdo.iloc[1234:1236, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1234: 'Itiquira I',
                                             1235: 'Itiquira II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1234:1236, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1234: 'Itiquira I',
                                             1235: 'Itiquira II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Volume.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Volume.xlsx',
            index=False)
        carga = ipdo.iloc[1234:1236, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1234: 'Itiquira I',
                                             1235: 'Itiquira II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRPE.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRPE.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1234:1236, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1234: 'Itiquira I',
                                             1235: 'Itiquira II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Afluença.xlsx')
        carga = ipdo.iloc[1234:1236, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1234: 'Itiquira I',
                                             1235: 'Itiquira II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Defluência.xlsx')
        carga = ipdo.iloc[1234:1236, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1234: 'Itiquira I',
                                             1235: 'Itiquira II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Nível.xlsx')
        carga = ipdo.iloc[1234:1236, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1234: 'Itiquira I',
                                             1235: 'Itiquira II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRI_Volume.xlsx')
        carga = ipdo.iloc[1234:1236, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1234: 'Itiquira I',
                                             1235: 'Itiquira II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSinRPE.xlsx')


def DHSRioCorrentes(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Correntes.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DHSRiCevertimento.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiC_Afluença.xlsx')
        add_inf2 = pd.read_excel(
           r'C:\Users\e806128\Desktop\1\DHSRiC_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiC_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiC_volume.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCevertimento.xlsx')
        no = no + 3
        carga = ipdo.iloc[1236:1237, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1236: 'P.Pedra'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiC_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiC_Afluença.xlsx',
            index=False)
        carga = ipdo.iloc[1236:1237, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1236: 'P. Pedra'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiC_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiC_Defluência.xlsx',
            index=False)
        carga = ipdo.iloc[1236:1237, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1236: 'P. Pedra'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiC_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiC_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1236:1237, 12:22]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1236: 'P. Pedra'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiC_volume.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiC_volume.xlsx',
            index=False)

        carga = ipdo.iloc[1236:1237, 12:22]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1236: 'P. Pedra'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCevertimento.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCevertimento.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1236:1237, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1236: 'P.Pedra'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiC_Afluença.xlsx')
        carga = ipdo.iloc[1236:1237, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1236: 'P. Pedra'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiC_Defluência.xlsx')
        carga = ipdo.iloc[1236:1237, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1236: 'P. Pedra'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiC_Nível.xlsx')
        carga = ipdo.iloc[1236:1237, 12:22]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1236: 'P. Pedra'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiC_volume.xlsx')

        carga = ipdo.iloc[1236:1237, 12:22]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1236: 'P. Pedra'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCevertimento.xlsx')


def DHSRioJauru(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Jauru.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DHSRiJauruver.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauru_Afluen.xlsx')
        add_inf2 = pd.read_excel(
           r'C:\Users\e806128\Desktop\1\DHSRiJauru_Def.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauru_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauru_volume.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauruver.xlsx')
        no = no + 3
        carga = ipdo.iloc[1237:1238, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1237: 'Jauru'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauru_Afluen.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauru_Afluen.xlsx',
            index=False)
        carga = ipdo.iloc[1237:1238, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1237: 'Jauru'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauru_Def.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauru_Def.xlsx',
            index=False)
        carga = ipdo.iloc[1237:1238, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1237: ' Jauru'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauru_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauru_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1237:1238, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1237: 'Jauru'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauru_volume.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauru_volume.xlsx',
            index=False)
        carga = ipdo.iloc[1237:1238, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1237: 'Jauru', })
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauruver.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauruver.xlsx',
            index=False)
    else:
        print("O arquivo nao existe.")
        carga = ipdo.iloc[1237:1238, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1237: 'Jauru'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauru_Afluen.xlsx')
        carga = ipdo.iloc[1237:1238, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1237: 'Jauru'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauru_Def.xlsx')

        carga = ipdo.iloc[1237:1238, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1237: ' Jauru'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauru_Nível.xlsx')
        carga = ipdo.iloc[1237:1238, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1237: 'Jauru'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauru_volume.xlsx')
        carga = ipdo.iloc[1237:1238, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1237: 'Jauru', })
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJauruver.xlsx')


def DHSRioJordan(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Jordan.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DHSRiJordaoevert.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Aflu.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Defl.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Vol.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordaoevert.xlsx')
        no = no + 3
        carga = ipdo.iloc[1238:1241, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1238: 'Sta. Clara',
                                             1239: 'Fundão',
                                             1240: 'Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Aflu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Aflu.xlsx',
            index=False)
        carga = ipdo.iloc[1238:1241, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1238: 'Sta.Clara',
                                             1239: 'Fundão',
                                             1240: 'Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Defl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Defl.xlsx',
            index=False)
        carga = ipdo.iloc[1238:1241, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1238: 'Sta. Clara',
                                             1239: 'Fundão',
                                             1240: 'Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1238:1241, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1238: 'Sta. Clara',
                                             1239: 'Fundão',
                                             1240: 'Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Vol.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Vol.xlsx',
            index=False)
        carga = ipdo.iloc[1238:1241, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1238: 'Sta. Clara',
                                             1239: ' Fundão',
                                             1240: ' Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordaoevert.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordaoevert.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1238:1241, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1238: 'Sta. Clara',
                                             1239: 'Fundão',
                                             1240: 'Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Aflu.xlsx')
        carga = ipdo.iloc[1238:1241, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1238: 'Sta.Clara',
                                             1239: 'Fundão',
                                             1240: 'Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Defl.xlsx')
        carga = ipdo.iloc[1238:1241, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1238: 'Sta. Clara',
                                             1239: 'Fundão',
                                             1240: 'Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Nível.xlsx')
        carga = ipdo.iloc[1238:1241, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1238: 'Sta. Clara',
                                             1239: 'Fundão',
                                             1240: 'Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordao_Vol.xlsx')
        carga = ipdo.iloc[1238:1241, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1238: 'Sta. Clara',
                                             1239: ' Fundão',
                                             1240: ' Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJordaoevert.xlsx')


def DHSRioIguaçu(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Iguaçu.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DHSRiIguaçueverti.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Aflue.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Deflu.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçueverti.xlsx')

        no = no + 3
        carga = ipdo.iloc[1241:1248, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1241: 'G. B. Munhoz',
                                             1242: 'G. Ney Braga',
                                             1243: 'S. Santiago',
                                             1244: 'S. Osório',
                                             1245: 'Gov. José Richa',
                                             1246: 'Capanema',
                                             1247: ' Baixo Iguaçu'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Aflue.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Aflue.xlsx',
            index=False)
        carga = ipdo.iloc[1241:1248, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1241: 'G. B. Munhoz',
                                             1242: 'G. Ney Braga',
                                             1243: 'S. Santiago',
                                             1244: 'S. Osório',
                                             1245: 'Gov. José Richa',
                                             1246: 'Capanema',
                                             1247: ' Baixo Iguaçu'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Deflu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Deflu.xlsx',
            index=False)
        carga = ipdo.iloc[1241:1248, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1241: 'G. B. Munhoz',
                                             1242: ' G. Ney Braga',
                                             1243: 'S. Santiago',
                                             1244: 'S. Osório',
                                             1245: 'Gov. José Richa',
                                             1246: 'Capanema',
                                             1247: ' Baixo Iguaçu'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1241:1248, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1241: 'G. B. Munhoz',
                                             1242: 'G. Ney Braga',
                                             1243: 'S. Santiago',
                                             1244: 'S. Osório',
                                             1245: 'Gov. José Richa',
                                             1246: 'Capanema',
                                             1247: ' Baixo Iguaçu'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Volu.xlsx',
            index=False)
        carga = ipdo.iloc[1241:1248, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1241: 'G. B. Munhoz',
                                             1242: 'G. Ney Braga',
                                             1243: 'S. Santiago',
                                             1244: 'S. Osório',
                                             1245: 'Gov. José Richa',
                                             1246: 'Capanema',
                                             1247: ' Baixo Iguaçu'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçueverti.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçueverti.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1241:1248, 12:16]

        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1241: 'G. B. Munhoz',
                                             1242: 'G. Ney Braga',
                                             1243: 'S. Santiago',
                                             1244: 'S. Osório',
                                             1245: 'Gov. José Richa',
                                             1246: 'Capanema',
                                             1247: ' Baixo Iguaçu'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        carga1_f = carga1_re1.iloc[2:3]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Aflue.xlsx')

        carga = ipdo.iloc[1241:1248, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1241: 'G. B. Munhoz',
                                             1242: 'G. Ney Braga',
                                             1243: 'S. Santiago',
                                             1244: 'S. Osório',
                                             1245: 'Gov. José Richa',
                                             1246: 'Capanema',
                                             1247: ' Baixo Iguaçu'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        carga1_f = carga1_re1.iloc[3:4]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Deflu.xlsx')

        carga = ipdo.iloc[1241:1248, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1241: 'G. B. Munhoz',
                                             1242: ' G. Ney Braga',
                                             1243: 'S. Santiago',
                                             1244: 'S. Osório',
                                             1245: 'Gov. José Richa',
                                             1246: 'Capanema',
                                             1247: ' Baixo Iguaçu'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        carga1_f = carga1_re1.iloc[4:5]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Nível.xlsx')
        carga = ipdo.iloc[1241:1248, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1241: 'G. B. Munhoz',
                                             1242: 'G. Ney Braga',
                                             1243: 'S. Santiago',
                                             1244: 'S. Osório',
                                             1245: 'Gov. José Richa',
                                             1246: 'Capanema',
                                             1247: ' Baixo Iguaçu'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        carga1_f = carga1_re1.iloc[6:7]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçu_Volu.xlsx')
        carga = ipdo.iloc[1241:1248, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1241: 'G. B. Munhoz',
                                             1242: 'G. Ney Braga',
                                             1243: 'S. Santiago',
                                             1244: 'S. Osório',
                                             1245: 'Gov. José Richa',
                                             1246: 'Capanema',
                                             1247: ' Baixo Iguaçu'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        carga1_f = carga1_re1.iloc[8:9]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiIguaçueverti.xlsx')


def DHSrioCanoa(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Canoa.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DHSRiCanoasevert.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoa_Afluen.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoa_Defluê.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoa_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoas_Volum.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoasevert.xlsx')
        no = no + 3
        carga = ipdo.iloc[1248:1251, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1248: 'São Roque',
                                             1249: 'Garibaldi',
                                             1250: 'Campos Novos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoa_Afluen.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoa_Afluen.xlsx',
            index=False)
        carga = ipdo.iloc[1248:1251, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1248: 'São Roque',
                                             1249: 'Garibaldi',
                                             1250: 'Campos Novos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoa_Defluê.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoa_Defluê.xlsx',
            index=False)
        carga = ipdo.iloc[1248:1251, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1248: 'São Roque',
                                             1249: 'Garibaldi',
                                             1250: 'Campos Novos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoa_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoa_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1248:1251, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1248: 'São Roque',
                                             1249: 'Garibaldi',
                                             1250: 'Campos Novos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoas_Volum.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoas_Volum.xlsx',
            index=False)
        carga = ipdo.iloc[1248:1251, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1248: 'São Roque',
                                             1249: 'Garibaldi',
                                             1250: 'Campos Novos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoasevert.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoasevert.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1248:1251, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1248: 'São Roque',
                                             1249: 'Garibaldi',
                                             1250: 'Campos Novos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoa_Afluen.xlsx')
        carga = ipdo.iloc[1248:1251, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1248: 'São Roque',
                                             1249: 'Garibaldi',
                                             1250: 'Campos Novos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoa_Defluê.xlsx')
        carga = ipdo.iloc[1248:1251, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1248: 'São Roque',
                                             1249: 'Garibaldi',
                                             1250: 'Campos Novos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoa_Nível.xlsx')
        carga = ipdo.iloc[1248:1251, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1248: 'São Roque',
                                             1249: 'Garibaldi',
                                             1250: 'Campos Novos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoas_Volum.xlsx')
        carga = ipdo.iloc[1248:1251, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1248: 'São Roque',
                                             1249: 'Garibaldi',
                                             1250: 'Campos Novos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCanoasevert.xlsx')


def DHSRiopassoFundo(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Passo Fundo.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_verti.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Def.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Ní.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_verti.xlsx')
        no = no + 3
        carga = ipdo.iloc[1251:1254, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1251: 'Alzir Santos',
                                             1252: 'Passo Fundo',
                                             1253: 'Foz do Chapecó '})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Afl.xlsx',
            index=False)
        carga = ipdo.iloc[1251:1254, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1251: 'Alzir Santos',
                                             1252: 'Passo Fundo',
                                             1253: 'Foz do Chapecó'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Def.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Def.xlsx',
            index=False)
        carga = ipdo.iloc[1251:1254, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1251: 'Alzir Santos',
                                             1252: 'Passo Fundo',
                                             1253: 'Foz do Chapecó'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Ní.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Ní.xlsx',
            index=False)
        carga = ipdo.iloc[1251:1254, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1251: 'Alzir Santos',
                                             1252: 'Passo Fundo',
                                             1253: 'Foz do Chapecó '})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Volu.xlsx',
            index=False)

        carga = ipdo.iloc[1251:1254, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1251: 'Alzir Santos',
                                             1252: 'Passo Fundo',
                                             1253: 'Foz do Chapecó '})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_verti.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_verti.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1251:1254, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1251: 'Alzir Santos',
                                             1252: 'Passo Fundo',
                                             1253: 'Foz do Chapecó '})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Afl.xlsx')
        carga = ipdo.iloc[1251:1254, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1251: 'Alzir Santos',
                                             1252: 'Passo Fundo',
                                             1253: 'Foz do Chapecó'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Def.xlsx')
        carga = ipdo.iloc[1251:1254, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1251: 'Alzir Santos',
                                             1252: 'Passo Fundo',
                                             1253: 'Foz do Chapecó'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Ní.xlsx')
        carga = ipdo.iloc[1251:1254, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1251: 'Alzir Santos',
                                             1252: 'Passo Fundo',
                                             1253: 'Foz do Chapecó '})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_Volu.xlsx')

        carga = ipdo.iloc[1251:1254, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1251: 'Alzir Santos',
                                             1252: 'Passo Fundo',
                                             1253: 'Foz do Chapecó '})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPassoFun_verti.xlsx')


def DHSRChapeco(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Chapeco.
    -------
    """
    caminho_arquivo = (
        r'C:\Users\e806128\Desktop\1\DHSRiChapecoevert.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Aflu.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Defl.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Nívl.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Volume.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapecoevert.xlsx')
        no = no + 3
        carga = ipdo.iloc[1254:1255, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1254: 'Quebra Queixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Aflu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Aflu.xlsx',
            index=False)
        carga = ipdo.iloc[1254:1255, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1254: 'Quebra Queixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Defl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Defl.xlsx',
            index=False)

        carga = ipdo.iloc[1254:1255, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1254: 'Quebra Queixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Nívl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Nívl.xlsx',
            index=False)

        carga = ipdo.iloc[1254:1255, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1254: 'Quebra Queixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Volume.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Volume.xlsx',
            index=False)
        carga = ipdo.iloc[1254:1255, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1254: 'Quebra Queixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapecoevert.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapecoevert.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1254:1255, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1254: 'Quebra Queixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Aflu.xlsx')
        carga = ipdo.iloc[1254:1255, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1254: 'Quebra Queixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Defl.xlsx')

        carga = ipdo.iloc[1254:1255, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1254: 'Quebra Queixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Nívl.xlsx')

        carga = ipdo.iloc[1254:1255, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1254: 'Quebra Queixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapeco_Volume.xlsx')
        carga = ipdo.iloc[1254:1255, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1254: 'Quebra Queixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiChapecoevert.xlsx')


def DHSRioPelotas(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Pelotas.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DHSRiPelotasverto.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Aflu.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Defl.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Níve.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotasverto.xlsx')
        no = no + 3
        carga = ipdo.iloc[1255:1258, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1255: 'Barra Grande',
                                             1256: 'Machadinho ',
                                             1257: 'Itá'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Aflu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Aflu.xlsx',
            index=False)
        carga = ipdo.iloc[1255:1258, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1255: 'Barra Grande',
                                             1256: 'Machadinho ',
                                             1257: 'Itá'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Defl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Defl.xlsx',
            index=False)

        carga = ipdo.iloc[1255:1258, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1255: 'Barra Grande',
                                             1256: 'Machadinho ',
                                             1257: 'Itá'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Níve.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Níve.xlsx',
            index=False)

        carga = ipdo.iloc[1255:1258, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1255: 'Barra Grande',
                                             1256: 'Machadinho ',
                                             1257: 'Itá'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Volu.xlsx',
            index=False)
        carga = ipdo.iloc[1255:1258, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1255: 'Barra Grande',
                                             1256: 'Machadinho ',
                                             1257: 'Itá'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotasverto.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]

        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotasverto.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1255:1258, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1255: 'Barra Grande',
                                             1256: 'Machadinho ',
                                             1257: 'Itá'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Aflu.xlsx')
        carga = ipdo.iloc[1255:1258, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1255: 'Barra Grande',
                                             1256: 'Machadinho ',
                                             1257: 'Itá'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Defl.xlsx')

        carga = ipdo.iloc[1255:1258, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1255: 'Barra Grande',
                                             1256: 'Machadinho ',
                                             1257: 'Itá'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Níve.xlsx')

        carga = ipdo.iloc[1255:1258, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1255: 'Barra Grande',
                                             1256: 'Machadinho ',
                                             1257: 'Itá'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotas_Volu.xlsx')
        carga = ipdo.iloc[1255:1258, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1255: 'Barra Grande',
                                             1256: 'Machadinho ',
                                             1257: 'Itá'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiPelotasverto.xlsx')


def DHSRioJacui(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Jacui.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DHSRiJacuiverto.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJacui_Aflue.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJacui_Defluê.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJacui_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJacui_Volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJacuiverto.xlsx')
        no = no + 3
        carga = ipdo.iloc[1258:1263, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1258: 'Ernestina',
                                             1259: 'Passo Real',
                                             1260: 'Jacuí',
                                             1261: 'Itaúba ',
                                             1262: 'D. Francisca'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJacui_Aflue.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJacui_Aflue.xlsx',
            index=False)
        carga = ipdo.iloc[1258:1263, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1258: 'Ernestina',
                                             1259: 'Passo Real',
                                             1260: 'Jacuí',
                                             1261: 'Itaúba',
                                             1262: 'D. Francisca'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJacui_Defluê.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJacui_Defluê.xlsx',
            index=False)
        carga = ipdo.iloc[1258:1263, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1258: 'Ernestina',
                                             1259: 'Passo Real',
                                             1260: 'Jacuí',
                                             1261: 'Itaúba ',
                                             1262: 'D. Francisca'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJacui_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJacui_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1258:1263, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1258: 'Ernestina',
                                             1259: 'Passo Real',
                                             1260: 'Jacuí',
                                             1261: 'Itaúba ',
                                             1262: 'D. Francisca '})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJacui_Volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJacui_Volu.xlsx',
            index=False)
        carga = ipdo.iloc[1258:1263, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1258: 'Ernestina',
                                             1259: 'Passo Real',
                                             1260: 'Jacuí',
                                             1261: 'Itaúba ',
                                             1262: 'D. Francisca '})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJacuiverto.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiJacuiverto.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1258:1263, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1258: 'Ernestina',
                                             1259: 'Passo Real',
                                             1260: 'Jacuí',
                                             1261: 'Itaúba ',
                                             1262: 'D. Francisca'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        carga1_f = carga1_re1.iloc[2:3]
        carga1_f.to_excel(r'C:\Users\e806128\Desktop\1\DHSRiJacui_Aflue.xlsx')
        carga = ipdo.iloc[1258:1263, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1258: 'Ernestina',
                                             1259: 'Passo Real',
                                             1260: 'Jacuí',
                                             1261: 'Itaúba',
                                             1262: 'D. Francisca'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        carga1_f = carga1_re1.iloc[3:4]
        carga1_f.to_excel(r'C:\Users\e806128\Desktop\1\DHSRiJacui_Defluê.xlsx')

        carga = ipdo.iloc[1258:1263, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1258: 'Ernestina',
                                             1259: 'Passo Real',
                                             1260: 'Jacuí',
                                             1261: 'Itaúba ',
                                             1262: 'D. Francisca'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        carga1_f = carga1_re1.iloc[4:5]
        carga1_f.to_excel(r'C:\Users\e806128\Desktop\1\DHSRiJacui_Nível.xlsx')
        carga = ipdo.iloc[1258:1263, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1258: 'Ernestina',
                                             1259: 'Passo Real',
                                             1260: 'Jacuí',
                                             1261: 'Itaúba ',
                                             1262: 'D. Francisca '})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        carga1_f = carga1_re1.iloc[6:7]
        carga1_f.to_excel(r'C:\Users\e806128\Desktop\1\DHSRiJacui_Volu.xlsx')
        carga = ipdo.iloc[1258:1263, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1258: 'Ernestina',
                                             1259: 'Passo Real',
                                             1260: 'Jacuí',
                                             1261: 'Itaúba ',
                                             1262: 'D. Francisca '})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        carga1_f = carga1_re1.iloc[8:9]
        carga1_f.to_excel(r'C:\Users\e806128\Desktop\1\DHSRiJacuiverto.xlsx')


def DHSRioTaquari(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Taquari.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DHSRiTaquAntver.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_Aflu.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_Defl.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_Nív.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAntver.xlsx')
        no = no + 3
        carga = ipdo.iloc[1263:1266, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1263: 'Castro Alves',
                                             1264: 'Monte Claro',
                                             1265: '14 de Julho'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_Aflu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_Aflu.xlsx',
            index=False)
        carga = ipdo.iloc[1263:1266, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1263: 'Castro Alves',
                                             1264: 'Monte Claro',
                                             1265: '14 de Julho'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_Defl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_Defl.xlsx',
            index=False)
        carga = ipdo.iloc[1263:1266, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1263: 'Castro Alves',
                                             1264: 'Monte Claro',
                                             1265: '14 de Julho'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_Nív.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_Nív.xlsx',
            index=False)
        carga = ipdo.iloc[1263:1266, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1263: 'Castro Alves',
                                             1264: 'Monte Claro',
                                             1265: '14 de Julho'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1263:1266, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1263: 'Castro Alves',
                                             1264: 'Monte Claro',
                                             1265: '14 de Julho'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAntver.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAntver.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1263:1266, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1263: 'Castro Alves',
                                             1264: 'Monte Claro',
                                             1265: '14 de Julho'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        carga1_f = carga1_re1.iloc[2:3]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_Aflu.xlsx')
        carga = ipdo.iloc[1263:1266, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1263: 'Castro Alves',
                                             1264: 'Monte Claro',
                                             1265: '14 de Julho'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        carga1_f = carga1_re1.iloc[3:4]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_Defl.xlsx')

        carga = ipdo.iloc[1263:1266, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1263: 'Castro Alves',
                                             1264: 'Monte Claro',
                                             1265: '14 de Julho'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        carga1_f = carga1_re1.iloc[4:5]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_Nív.xlsx')
        carga = ipdo.iloc[1263:1266, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1263: 'Castro Alves',
                                             1264: 'Monte Claro',
                                             1265: '14 de Julho'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        carga1_f = carga1_re1.iloc[6:7]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAnt_volu.xlsx')
        carga = ipdo.iloc[1263:1266, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1263: 'Castro Alves',
                                             1264: 'Monte Claro',
                                             1265: '14 de Julho'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        carga1_f = carga1_re1.iloc[8:9]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiTaquAntver.xlsx')


def DHSRioCapivari(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Capivari.
    -------
    """
    caminho_arquivo = (
        r'C:\Users\e806128\Desktop\1\DHSRiCapivariv.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivari_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivari_Def.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivari_Ní.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivari_volu.xlsx')
        add_inf5 = pd.read_excel(
           r'C:\Users\e806128\Desktop\1\DHSRiCapivariv.xlsx')
        no = no + 3
        carga = ipdo.iloc[1266:1267, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1266: 'G. P. Souza'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivari_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivari_Afl.xlsx',
            index=False)
        carga = ipdo.iloc[1266:1267, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1266: 'G. P. Souza'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivari_Def.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivari_Def.xlsx',
            index=False)
        carga = ipdo.iloc[1266:1267, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1266: 'G. P. Souza'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivari_Ní.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
           r'C:\Users\e806128\Desktop\1\DHSRiCapivari_Ní.xlsx',
           index=False)
        carga = ipdo.iloc[1266:1267, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1266: 'G. P. Souza'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivari_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivari_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1266:1267, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1266: 'G. P. Souza'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivariv.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivariv.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1266:1267, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1266: 'G. P. Souza'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        carga1_f = carga1_re1.iloc[2:3]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivari_Afl.xlsx')
        carga = ipdo.iloc[1266:1267, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1266: 'G. P. Souza'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        carga1_f = carga1_re1.iloc[3:4]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivari_Def.xlsx')
        carga = ipdo.iloc[1266:1267, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1266: 'G. P. Souza'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        carga1_f = carga1_re1.iloc[4:5]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivari_Ní.xlsx')
        carga = ipdo.iloc[1266:1267, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1266: 'G. P. Souza'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        carga1_f = carga1_re1.iloc[6:7]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivari_volu.xlsx')
        carga = ipdo.iloc[1266:1267, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1266: 'G. P. Souza'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        carga1_f = carga1_re1.iloc[8:9]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DHSRiCapivariv.xlsx')


def dadosbacia84(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados da Bacia 84.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DaBaci84_GerProg.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_Arm.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_ENAdia.xlsx')
        add_inf3 = pd.read_excel(
           r'C:\Users\e806128\Desktop\1\DaBaci84_ENAArmaz.xlsx')
        add_inf4 = pd.read_excel(
           r'C:\Users\e806128\Desktop\1\DaBaci84_ENABruta.xlsx')
        add_inf5 = pd.read_excel(
          r'C:\Users\e806128\Desktop\1\DaBaci84_GerVerif.xlsx')
        add_inf6 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_GerProg.xlsx')
        no = no + 3
        carga = ipdo.iloc[1273:1279, 10:13]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1274: 'IGU',
                                             1275: 'JAC',
                                             1276: 'URY',
                                             1277: 'CAP',
                                             1278: 'PRG'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 12': data})

        form_tab1 = carga1_re1.iloc[0:1]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_Arm.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_Arm.xlsx',
            index=False)
        carga = ipdo.iloc[1273:1279, 10:15]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1274: 'IGU',
                                             1275: 'JAC',
                                             1276: 'URY',
                                             1277: 'CAP',
                                             1278: 'PRG'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 14': data})

        form_tab1 = carga1_re1.iloc[2:3]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_ENAdia.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_ENAdia.xlsx',
            index=False)
        carga = ipdo.iloc[1273:1279, 10:16]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1274: 'IGU',
                                             1275: 'JAC',
                                             1276: 'URY',
                                             1277: 'CAP',
                                             1278: 'PRG'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab1 = carga1_re1.iloc[3:4]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_ENAArmaz.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_ENAArmaz.xlsx',
            index=False)
        carga = ipdo.iloc[1273:1279, 10:17]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1274: 'IGU',
                                             1275: 'JAC',
                                             1276: 'URY',
                                             1277: 'CAP',
                                             1278: 'PRG'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab1 = carga1_re1.iloc[4:5]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_ENABruta.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_ENABruta.xlsx',
            index=False)
        carga = ipdo.iloc[1273:1279, 10:18]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1274: 'IGU',
                                             1275: 'JAC',
                                             1276: 'URY',
                                             1277: 'CAP',
                                             1278: 'PRG'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        form_tab1 = carga1_re1.iloc[5:6]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_GerVerif.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_GerVerif.xlsx',
            index=False)
        carga = ipdo.iloc[1273:1279, 10:20]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1274: 'IGU',
                                             1275: 'JAC',
                                             1276: 'URY',
                                             1277: 'CAP',
                                             1278: 'PRG'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab1 = carga1_re1.iloc[7:8]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_GerProg.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        add_inf6.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6
                            )
        add_inf6 = add_inf6.rename(columns={'Unnamed: 0': ' '})
        add_inf6.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_GerProg.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1273:1279, 10:13]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1274: 'IGU',
                                             1275: 'JAC',
                                             1276: 'URY',
                                             1277: 'CAP',
                                             1278: 'PRG'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 12': data})

        form_tab1 = carga1_re1.iloc[0:1]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaBaci84_Arm.xlsx')
        carga = ipdo.iloc[1273:1279, 10:15]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1274: 'IGU',
                                             1275: 'JAC',
                                             1276: 'URY',
                                             1277: 'CAP',
                                             1278: 'PRG'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 14': data})
        form_tab1 = carga1_re1.iloc[2:3]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_ENAdia.xlsx')
        carga = ipdo.iloc[1273:1279, 10:16]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1274: 'IGU',
                                             1275: 'JAC',
                                             1276: 'URY',
                                             1277: 'CAP',
                                             1278: 'PRG'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab1 = carga1_re1.iloc[3:4]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_ENAArmaz.xlsx')
        carga = ipdo.iloc[1273:1279, 10:17]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1274: 'IGU',
                                             1275: 'JAC',
                                             1276: 'URY',
                                             1277: 'CAP',
                                             1278: 'PRG'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab1 = carga1_re1.iloc[4:5]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_ENABruta.xlsx')
        carga = ipdo.iloc[1273:1279, 10:18]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1274: 'IGU',
                                             1275: 'JAC',
                                             1276: 'URY',
                                             1277: 'CAP',
                                             1278: 'PRG'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        form_tab1 = carga1_re1.iloc[5:6]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_GerVerif.xlsx')
        carga = ipdo.iloc[1273:1279, 10:20]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1274: 'IGU',
                                             1275: 'JAC',
                                             1276: 'URY',
                                             1277: 'CAP',
                                             1278: 'PRG'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab1 = carga1_re1.iloc[7:8]
        colunassx = [0]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci84_GerProg.xlsx')


def DadosHSRioTocantis(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Tocantins.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DaHSRTocatinsver.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatins_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatins_Def.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatins_Nív.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatins_volume.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatinsver.xlsx')
        no = no + 3
        carga = ipdo.iloc[1291:1298, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1291: 'Serra da Mesa',
                                             1292: 'Cana Brava',
                                             1293: 'São Salvador',
                                             1294: 'Peixe Angical',
                                             1295: 'Lajeado',
                                             1296: 'Estreito',
                                             1297: 'Tucuruí'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatins_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatins_Afl.xlsx',
            index=False)
        carga = ipdo.iloc[1291:1298, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1291: 'Serra da Mesa',
                                             1292: 'Cana Brava',
                                             1293: 'São Salvador',
                                             1294: 'Peixe Angical',
                                             1295: 'Lajeado',
                                             1296: 'Estreito',
                                             1297: 'Tucuruí'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatins_Def.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
           r'C:\Users\e806128\Desktop\1\DaHSRTocatins_Def.xlsx',
           index=False)
        carga = ipdo.iloc[1291:1298, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1291: 'Serra da Mesa',
                                             1292: 'Cana Brava',
                                             1293: 'São Salvador',
                                             1294: 'Peixe Angical',
                                             1295: 'Lajeado',
                                             1296: 'Estreito',
                                             1297: 'Tucuruí'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatins_Nív.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatins_Nív.xlsx',
            index=False)
        carga = ipdo.iloc[1291:1298, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1291: 'Serra da Mesa',
                                             1292: 'Cana Brava',
                                             1293: 'São Salvador',
                                             1294: 'Peixe Angical',
                                             1295: 'Lajeado',
                                             1296: 'Estreito',
                                             1297: 'Tucuruí'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatins_volume.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatins_volume.xlsx',
            index=False)
        carga = ipdo.iloc[1291:1298, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1291: 'Serra da Mesa',
                                             1292: 'Cana Brava',
                                             1293: 'São Salvador',
                                             1294: 'Peixe Angical',
                                             1295: 'Lajeado',
                                             1296: 'Estreito',
                                             1297: 'Tucuruí'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatinsver.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatinsver.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1291:1298, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1291: 'Serra da Mesa',
                                             1292: 'Cana Brava',
                                             1293: 'São Salvador',
                                             1294: 'Peixe Angical',
                                             1295: 'Lajeado',
                                             1296: 'Estreito',
                                             1297: 'Tucuruí'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        carga1_f = carga1_re1.iloc[2:3]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatins_Afl.xlsx')
        carga = ipdo.iloc[1291:1298, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1291: 'Serra da Mesa',
                                             1292: 'Cana Brava',
                                             1293: 'São Salvador',
                                             1294: 'Peixe Angical',
                                             1295: 'Lajeado',
                                             1296: 'Estreito',
                                             1297: 'Tucuruí'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        carga1_f = carga1_re1.iloc[3:4]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatins_Def.xlsx')
        carga = ipdo.iloc[1291:1298, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1291: 'Serra da Mesa',
                                             1292: 'Cana Brava',
                                             1293: 'São Salvador',
                                             1294: 'Peixe Angical',
                                             1295: 'Lajeado',
                                             1296: 'Estreito',
                                             1297: 'Tucuruí'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        carga1_f = carga1_re1.iloc[4:5]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatins_Nív.xlsx')
        carga = ipdo.iloc[1291:1298, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1291: 'Serra da Mesa',
                                             1292: 'Cana Brava',
                                             1293: 'São Salvador',
                                             1294: 'Peixe Angical',
                                             1295: 'Lajeado',
                                             1296: 'Estreito',
                                             1297: 'Tucuruí'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        carga1_f = carga1_re1.iloc[6:7]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatins_volume.xlsx')
        carga = ipdo.iloc[1291:1298, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1291: 'Serra da Mesa',
                                             1292: 'Cana Brava',
                                             1293: 'São Salvador',
                                             1294: 'Peixe Angical',
                                             1295: 'Lajeado',
                                             1296: 'Estreito',
                                             1297: 'Tucuruí'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        carga1_f = carga1_re1.iloc[8:9]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRTocatinsver.xlsx')


def dadosHSRioPetro(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Pretro.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DaHSRPretovert.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_Aflu.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_Defluê.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPretovert.xlsx')
        no = no + 3
        carga = ipdo.iloc[1298:1299, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1298: 'Queimado'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_Aflu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_Aflu.xlsx',
            index=False)
        carga = ipdo.iloc[1298:1299, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1298: 'Queimado'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_Defluê.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_Defluê.xlsx',
            index=False)

        carga = ipdo.iloc[1298:1299, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1298: 'Queimado'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1298:1299, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1298: 'Queimado'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1298:1299, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1298: 'Queimado'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPretovert.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPretovert.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1298:1299, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1298: 'Queimado'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_Aflu.xlsx')
        carga = ipdo.iloc[1298:1299, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1298: 'Queimado'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_Defluê.xlsx')

        carga = ipdo.iloc[1298:1299, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1298: 'Queimado'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_Nível.xlsx')
        carga = ipdo.iloc[1298:1299, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1298: 'Queimado'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPreto_volu.xlsx')
        carga = ipdo.iloc[1298:1299, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1298: 'Queimado'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRPretovert.xlsx')


def DHSRioSãoFrancisco(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio São Francisco.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DaHSRSFrancisco_vert.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFranci_Aflu.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFrancis_Def.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFranci_Ní.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFrancis_Volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFrancisco_vert.xlsx')
        no = no + 3
        carga = ipdo.iloc[1299:1306, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1299: 'Três Marias',
                                             1300: 'Sobradinho',
                                             1301: 'Luiz Gonzaga',
                                             1302: 'Apolônio Sales',
                                             1303: 'P. Afonso 4',
                                             1304: 'P. Afonso 1,2,3',
                                             1305: 'Xingó'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFranci_Aflu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFranci_Aflu.xlsx',
            index=False)
        carga = ipdo.iloc[1299:1306, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1299: 'Três Marias',
                                             1300: 'Sobradinho',
                                             1301: 'Luiz Gonzaga',
                                             1302: 'Apolônio Sales',
                                             1303: 'P. Afonso 4',
                                             1304: 'P. Afonso 1,2,3',
                                             1305: 'Xingó'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFrancis_Def.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFrancis_Def.xlsx',
            index=False)
        carga = ipdo.iloc[1299:1306, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1299: 'Três Marias',
                                             1300: 'Sobradinho',
                                             1301: 'Luiz Gonzaga',
                                             1302: 'Apolônio Sales',
                                             1303: 'P. Afonso 4',
                                             1304: 'P. Afonso 1,2,3',
                                             1305: 'Xingó'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFranci_Ní.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFranci_Ní.xlsx',
            index=False)
        carga = ipdo.iloc[1299:1306, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1299: 'Três Marias',
                                             1300: 'Sobradinho',
                                             1301: 'Luiz Gonzaga',
                                             1302: 'Apolônio Sales',
                                             1303: 'P. Afonso 4',
                                             1304: 'P. Afonso 1,2,3',
                                             1305: 'Xingó'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFrancis_Volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFrancis_Volu.xlsx',
            index=False)
        carga = ipdo.iloc[1299:1306, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1299: 'Três Marias',
                                             1300: 'Sobradinho',
                                             1301: 'Luiz Gonzaga',
                                             1302: 'Apolônio Sales',
                                             1303: 'P. Afonso 4',
                                             1304: 'P. Afonso 1,2,3',
                                             1305: 'Xingó'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFrancisco_vert.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFrancisco_vert.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1299:1306, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1299: 'Três Marias',
                                             1300: 'Sobradinho',
                                             1301: 'Luiz Gonzaga',
                                             1302: 'Apolônio Sales',
                                             1303: 'P. Afonso 4',
                                             1304: 'P. Afonso 1,2,3',
                                             1305: 'Xingó'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFranci_Aflu.xlsx')

        carga = ipdo.iloc[1299:1306, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1299: 'Três Marias',
                                             1300: 'Sobradinho',
                                             1301: 'Luiz Gonzaga',
                                             1302: 'Apolônio Sales',
                                             1303: 'P. Afonso 4',
                                             1304: 'P. Afonso 1,2,3',
                                             1305: 'Xingó'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFrancis_Def.xlsx')
        carga = ipdo.iloc[1299:1306, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1299: 'Três Marias',
                                             1300: 'Sobradinho',
                                             1301: 'Luiz Gonzaga',
                                             1302: 'Apolônio Sales',
                                             1303: 'P. Afonso 4',
                                             1304: 'P. Afonso 1,2,3',
                                             1305: 'Xingó'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFranci_Ní.xlsx')
        carga = ipdo.iloc[1299:1306, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1299: 'Três Marias',
                                             1300: 'Sobradinho',
                                             1301: 'Luiz Gonzaga',
                                             1302: 'Apolônio Sales',
                                             1303: 'P. Afonso 4',
                                             1304: 'P. Afonso 1,2,3',
                                             1305: 'Xingó'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFrancis_Volu.xlsx')
        carga = ipdo.iloc[1299:1306, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1299: 'Três Marias',
                                             1300: 'Sobradinho',
                                             1301: 'Luiz Gonzaga',
                                             1302: 'Apolônio Sales',
                                             1303: 'P. Afonso 4',
                                             1304: 'P. Afonso 1,2,3',
                                             1305: 'Xingó'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRSFrancisco_vert.xlsx')


def dhSrioJequitinhonha(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Jequitinhonha.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DaHSRJequitver.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_Aflu.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_Deflu.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_Nív.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequitver.xlsx')
        no = no + 3
        carga = ipdo.iloc[1306:1308, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1306: 'Irapé',
                                             1307: 'Itapebi'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_Aflu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_Aflu.xlsx',
            index=False)
        carga = ipdo.iloc[1306:1308, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1306: 'Irapé',
                                             1307: 'Itapebi'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_Deflu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_Deflu.xlsx',
            index=False)
        carga = ipdo.iloc[1306:1308, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1306: 'Irapé',
                                             1307: 'Itapebi'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_Nív.xlsx')

        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_Nív.xlsx',
            index=False)
        carga = ipdo.iloc[1306:1308, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1306: 'Irapé',
                                             1307: 'Itapebi'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1306:1308, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1306: 'Irapé',
                                             1307: 'Itapebi'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequitver.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequitver.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1306:1308, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1306: 'Irapé',
                                             1307: 'Itapebi'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_Aflu.xlsx')
        carga = ipdo.iloc[1306:1308, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1306: 'Irapé',
                                             1307: 'Itapebi'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_Deflu.xlsx')
        carga = ipdo.iloc[1306:1308, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1306: 'Irapé',
                                             1307: 'Itapebi'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_Nív.xlsx')
        carga = ipdo.iloc[1306:1308, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1306: 'Irapé',
                                             1307: 'Itapebi'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequit_volu.xlsx')
        carga = ipdo.iloc[1306:1308, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1306: 'Irapé',
                                             1307: 'Itapebi'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSRJequitver.xlsx')


def DhSRioParaguaçu(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Paraguaçu.
    -------
    """
    caminho_arquivo = r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_vert.xlsx'
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Def.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Ní.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_vert.xlsx')
        no = no + 3
        carga = ipdo.iloc[1308:1309, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1308: 'Pedra do Cavalo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Afl.xlsx',
            index=False)
        carga = ipdo.iloc[1308:1309, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1308: 'Pedra do Cavalo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Def.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Def.xlsx',
            index=False)

        carga = ipdo.iloc[1308:1309, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1308: 'Pedra do Cavalo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Ní.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Ní.xlsx',
            index=False)

        carga = ipdo.iloc[1308:1309, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1308: 'Pedra do Cavalo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Volu.xlsx',
            index=False)
        carga = ipdo.iloc[1308:1309, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1308: 'Pedra do Cavalo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_vert.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_vert.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1308:1309, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1308: 'Pedra do Cavalo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        carga1_f = carga1_re1.iloc[2:3]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Afl.xlsx')
        carga = ipdo.iloc[1308:1309, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1308: 'Pedra do Cavalo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        carga1_f = carga1_re1.iloc[3:4]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Def.xlsx')

        carga = ipdo.iloc[1308:1309, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1308: 'Pedra do Cavalo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        carga1_f = carga1_re1.iloc[4:5]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Ní.xlsx')

        carga = ipdo.iloc[1308:1309, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1308: 'Pedra do Cavalo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        carga1_f = carga1_re1.iloc[6:7]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_Volu.xlsx')
        carga = ipdo.iloc[1308:1309, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1308: 'Pedra do Cavalo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        carga1_f = carga1_re1.iloc[8:9]
        carga1_f.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaguaçu_vert.xlsx')


def DhSRioParnaiba(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Parnaiba.
    -------
    """
    caminho_arquivo = (
        r'C:\Users\e806128\Desktop\1\DaHSParnaíba_vertimento.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_volume.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_vertimento.xlsx')
        no = no + 3
        carga = ipdo.iloc[1309:1310, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1309: 'B. Esperança'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_Afluença.xlsx',
            index=False)
        carga = ipdo.iloc[1309:1310, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1309: 'B. Esperança'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_Defluência.xlsx',
            index=False)
        carga = ipdo.iloc[1309:1310, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1309: 'B. Esperança'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1309:1310, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1309: 'B. Esperança'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_volume.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_volume.xlsx',
            index=False)
        carga = ipdo.iloc[1309:1310, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1309: 'B. Esperança'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_vertimento.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_vertimento.xlsx',
            index=False)
    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1309:1310, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1309: 'B. Esperança'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_Afluença.xlsx')
        carga = ipdo.iloc[1309:1310, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1309: 'B. Esperança'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_Defluência.xlsx')
        carga = ipdo.iloc[1309:1310, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1309: 'B. Esperança'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_Nível.xlsx')
        carga = ipdo.iloc[1309:1310, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1309: 'B. Esperança'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_volume.xlsx')
        carga = ipdo.iloc[1309:1310, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1309: 'B. Esperança'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParnaíba_vertimento.xlsx')


def DhSRioAripuana(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Aripuanã.
    -------
    """
    caminho_arquivo = (
        r'C:\Users\e806128\Desktop\1\DaHSAripuanã_vertimento.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_vertimento.xlsx')
        no = no + 3
        carga = ipdo.iloc[1310:1311, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1310: 'Dardanelos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_Afluença.xlsx',
            index=False)
        carga = ipdo.iloc[1310:1311, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1310: 'Dardanelos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_Defluência.xlsx',
            index=False)
        carga = ipdo.iloc[1310:1311, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1310: 'Dardanelos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1310:1311, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1310: 'Dardanelos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1310:1311, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1310: 'Dardanelos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_vertimento.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_vertimento.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1310:1311, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1310: 'Dardanelos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_Afluença.xlsx')
        carga = ipdo.iloc[1310:1311, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1310: 'Dardanelos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_Defluência.xlsx')
        carga = ipdo.iloc[1310:1311, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1310: 'Dardanelos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_Nível.xlsx')
        carga = ipdo.iloc[1310:1311, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1310: 'Dardanelos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_volu.xlsx')
        carga = ipdo.iloc[1310:1311, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1310: 'Dardanelos'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSAripuanã_vertimento.xlsx')


def DhSRioComemoracao(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Comemoração.
    -------
    """
    caminho_arquivo = (
        r'C:\Users\e806128\Desktop\1\DaHSComemoracao_vertimento.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_vertimento.xlsx')
        no = no + 3
        carga = ipdo.iloc[1311:1312, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1311: 'Rondon II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_Afluença.xlsx',
            index=False)
        carga = ipdo.iloc[1311:1312, 12:17]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1311: 'Rondon II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_Defluência.xlsx',
            index=False)
        carga = ipdo.iloc[1311:1312, 12:18]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1311: 'Rondon II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1311:1312, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1311: 'Rondon II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1311:1312, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1311: 'Rondon II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_vertimento.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_vertimento.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1311:1312, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1311: 'Rondon II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_Afluença.xlsx', )
        carga = ipdo.iloc[1311:1312, 12:17]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1311: 'Rondon II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_Defluência.xlsx')
        carga = ipdo.iloc[1311:1312, 12:18]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1311: 'Rondon II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_Nível.xlsx', )
        carga = ipdo.iloc[1311:1312, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1311: 'Rondon II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_volu.xlsx')
        carga = ipdo.iloc[1311:1312, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1311: 'Rondon II'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSComemoracao_vertimento.xlsx')


def DhSRioJamari(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Jamari.
    -------
    """
    caminho_arquivo = (
        r'C:\Users\e806128\Desktop\1\DaHSJamari_vertimento.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_vertimento.xlsx')
        no = no + 3
        carga = ipdo.iloc[1312:1313, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1312: 'Samuel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_Afluença.xlsx',
            index=False)
        carga = ipdo.iloc[1312:1313, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1312: 'Samuel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_Defluência.xlsx',
            index=False)
        carga = ipdo.iloc[1312:1313, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1312: 'Samuel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1312:1313, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1312: 'Samuel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1312:1313, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1312: 'Samuel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_vertimento.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_vertimento.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1312:1313, 12:16]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1312: 'Samuel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_Afluença.xlsx')
        carga = ipdo.iloc[1312:1313, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1312: 'Samuel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab = carga1_re1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_Defluência.xlsx')
        carga = ipdo.iloc[1312:1313, 12:18]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1312: 'Samuel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_Nível.xlsx')
        carga = ipdo.iloc[1312:1313, 12:22]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1312: 'Samuel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_volu.xlsx')
        carga = ipdo.iloc[1312:1313, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1312: 'Samuel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSJamari_vertimento.xlsx')


def DhSRioGuapore(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Guaporé.
    -------
    """
    caminho_arquivo = (
        r'C:\Users\e806128\Desktop\1\DaHSGuapore_vertimento.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_vertimento.xlsx')
        no = no + 3
        carga = ipdo.iloc[1313:1314, 12:16]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1313: 'Guaporé'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_Afluença.xlsx',
            index=False)
        carga = ipdo.iloc[1313:1314, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1313: 'Guaporé'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_Defluência.xlsx',
            index=False)
        carga = ipdo.iloc[1313:1314, 12:18]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1313: 'Guaporé'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1313:1314, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1313: 'Guaporé'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1313:1314, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1313: 'Guaporé'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_vertimento.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_vertimento.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")

        carga = ipdo.iloc[1313:1314, 12:16]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1313: 'Guaporé'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_Afluença.xlsx')
        carga = ipdo.iloc[1313:1314, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1313: 'Guaporé'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_Defluência.xlsx')
        carga = ipdo.iloc[1313:1314, 12:18]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1313: 'Guaporé'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_Nível.xlsx')
        carga = ipdo.iloc[1313:1314, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1313: 'Guaporé'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_volu.xlsx')
        carga = ipdo.iloc[1313:1314, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1313: 'Guaporé'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSGuapore_vertimento.xlsx')


def DhSRioMadeira(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Madeira.
    -------
    """
    caminho_arquivo = (
        r'C:\Users\e806128\Desktop\1\DaHSmadeira_vertimento.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_vertimento.xlsx')
        no = no + 3
        carga = ipdo.iloc[1314:1316, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1314: 'Sto Antônio',
                                             1315: 'datairau'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_Afluença.xlsx',
            index=False)
        carga = ipdo.iloc[1314:1316, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1314: 'Sto Antônio',
                                             1315: 'datairau'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_Defluência.xlsx',
            index=False)
        carga = ipdo.iloc[1314:1316, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1314: 'Sto Antônio',
                                             1315: 'datairau'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1314:1316, 12:22]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1314: 'Sto Antônio',
                                             1315: 'datairau'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1314:1316, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1314: 'Sto Antônio',
                                             1315: 'datairau'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_vertimento.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_vertimento.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1314:1316, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1314: 'Sto Antônio',
                                             1315: 'datairau'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_Afluença.xlsx')
        carga = ipdo.iloc[1314:1316, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1314: 'Sto Antônio',
                                             1315: 'datairau'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_Defluência.xlsx')
        carga = ipdo.iloc[1314:1316, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1314: 'Sto Antônio',
                                             1315: 'datairau'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_Nível.xlsx')
        carga = ipdo.iloc[1314:1316, 12:22]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1314: 'Sto Antônio',
                                             1315: 'datairau'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaHSmadeira_volu.xlsx')
        carga = ipdo.iloc[1314:1316, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1314: 'Sto Antônio',
                                             1315: 'datairau'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSmadeira_vertimento.xlsx')


def DhSRioUatumã(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Uatumã.
    -------
    """
    caminho_arquivo = (
        r'C:\Users\e806128\Desktop\1\DaHSUatumã_vertimento.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_vertimento.xlsx')
        no = no + 3
        carga = ipdo.iloc[1316:1317, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1316: 'Balbina'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_Afluença.xlsx',
            index=False)
        carga = ipdo.iloc[1316:1317, 12:17]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1316: 'Balbina'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_Defluência.xlsx',
            index=False)
        carga = ipdo.iloc[1316:1317, 12:18]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1316: 'Balbina'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_Nível.xlsx',
            index=False)

        carga = ipdo.iloc[1316:1317, 12:22]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1316: 'Balbina'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1316:1317, 12:22]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1316: 'Balbina'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_vertimento.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_vertimento.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1316:1317, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1316: 'Balbina'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_Afluença.xlsx')
        carga = ipdo.iloc[1316:1317, 12:17]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1316: 'Balbina'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_Defluência.xlsx')
        carga = ipdo.iloc[1316:1317, 12:18]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1316: 'Balbina'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_Nível.xlsx')

        carga = ipdo.iloc[1316:1317, 12:22]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1316: 'Balbina'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})

        form_tab = carga1_re1.iloc[6:7]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_volu.xlsx')
        carga = ipdo.iloc[1316:1317, 12:22]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1316: 'Balbina'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUatumã_vertimento.xlsx')


def DhSRioAraguari(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Araguari.
    -------
    """
    caminho_arquivo = (
       r'C:\Users\e806128\Desktop\1\DaHSUAraguari_vertimento.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_vertimento.xlsx')
        no = no + 3
        carga = ipdo.iloc[1317:1320, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1317: 'Cachoeira Caldeirão',
                                             1318: 'Ferreira Gomes',
                                             1319: 'Coaracy Nunes'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_Afluença.xlsx',
            index=False)
        carga = ipdo.iloc[1317:1320, 12:17]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1317: 'Cachoeira Caldeirão',
                                             1318: 'Ferreira Gomes',
                                             1319: 'Coaracy Nunes'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_Defluência.xlsx',
            index=False)
        carga = ipdo.iloc[1317:1320, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1317: 'Cachoeira Caldeirão',
                                             1318: 'Ferreira Gomes',
                                             1319: 'Coaracy Nunes'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1317:1320, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1317: 'Cachoeira Caldeirão',
                                             1318: 'Ferreira Gomes',
                                             1319: 'Coaracy Nunes'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1317:1320, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1317: 'Cachoeira Caldeirão',
                                             1318: 'Ferreira Gomes',
                                             1319: 'Coaracy Nunes'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_vertimento.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_vertimento.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1317:1320, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1317: 'Cachoeira Caldeirão',
                                             1318: 'Ferreira Gomes',
                                             1319: 'Coaracy Nunes'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_Afluença.xlsx')
        carga = ipdo.iloc[1317:1320, 12:17]

        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1317: 'Cachoeira Caldeirão',
                                             1318: 'Ferreira Gomes',
                                             1319: 'Coaracy Nunes'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_Defluência.xlsx')
        carga = ipdo.iloc[1317:1320, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1317: 'Cachoeira Caldeirão',
                                             1318: 'Ferreira Gomes',
                                             1319: 'Coaracy Nunes'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_Nível.xlsx')
        carga = ipdo.iloc[1317:1320, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1317: 'Cachoeira Caldeirão',
                                             1318: 'Ferreira Gomes',
                                             1319: 'Coaracy Nunes'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_volu.xlsx')
        carga = ipdo.iloc[1317:1320, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1317: 'Cachoeira Caldeirão',
                                             1318: 'Ferreira Gomes',
                                             1319: 'Coaracy Nunes'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUAraguari_vertimento.xlsx')


def DhSRioTelesPires(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio TelesPires.
    -------
    """
    caminho_arquivo = (
       r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_vertimento.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_vertimento.xlsx')
        no = no + 3
        carga = ipdo.iloc[1320:1324, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1320: 'Sinop',
                                             1321: 'Colíder',
                                             1322: 'Teles Pires',
                                             1323: 'São Manoel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_Afluença.xlsx',
            index=False)
        carga = ipdo.iloc[1320:1324, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1320: 'Sinop',
                                             1321: 'Colíder',
                                             1322: 'Teles Pires',
                                             1323: 'São Manoel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_Defluência.xlsx',
            index=False)
        carga = ipdo.iloc[1320:1324, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1320: 'Sinop',
                                             1321: 'Colíder',
                                             1322: 'Teles Pires',
                                             1323: 'São Manoel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1320:1324, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1320: 'Sinop',
                                             1321: 'Colíder',
                                             1322: 'Teles Pires',
                                             1323: 'São Manoel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1320:1324, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1320: 'Sinop',
                                             1321: 'Colíder',
                                             1322: 'Teles Pires',
                                             1323: 'São Manoel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_vertimento.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_vertimento.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1320:1324, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1320: 'Sinop',
                                             1321: 'Colíder',
                                             1322: 'Teles Pires',
                                             1323: 'São Manoel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_Afluença.xlsx')
        carga = ipdo.iloc[1320:1324, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1320: 'Sinop',
                                             1321: 'Colíder',
                                             1322: 'Teles Pires',
                                             1323: 'São Manoel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_Defluência.xlsx')
        carga = ipdo.iloc[1320:1324, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1320: 'Sinop',
                                             1321: 'Colíder',
                                             1322: 'Teles Pires',
                                             1323: 'São Manoel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_Nível.xlsx')
        carga = ipdo.iloc[1320:1324, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1320: 'Sinop',
                                             1321: 'Colíder',
                                             1322: 'Teles Pires',
                                             1323: 'São Manoel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_volu.xlsx')
        carga = ipdo.iloc[1320:1324, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1320: 'Sinop',
                                             1321: 'Colíder',
                                             1322: 'Teles Pires',
                                             1323: 'São Manoel'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUTelesPires_vertimento.xlsx')


def DhSRioCuruaUna(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Curuá-Una.
    -------
    """
    caminho_arquivo = (
        r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_vertimento.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_vertimento.xlsx')
        no = no + 3
        carga = ipdo.iloc[1324:1325, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1324: 'Curuá-Una'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_Afluença.xlsx',
            index=False)
        carga = ipdo.iloc[1324:1325, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1324: 'Curuá-Una'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_Defluência.xlsx',
            index=False)
        carga = ipdo.iloc[1324:1325, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1324: 'Curuá-Una'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1324:1325, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1324: 'Curuá-Una'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1324:1325, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1324: 'Curuá-Una'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_vertimento.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_vertimento.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1324:1325, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1324: 'Curuá-Una'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_Afluença.xlsx')
        carga = ipdo.iloc[1324:1325, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1324: 'Curuá-Una'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]

        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_Defluência.xlsx')
        carga = ipdo.iloc[1324:1325, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1324: 'Curuá-Una'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_Nível.xlsx')
        carga = ipdo.iloc[1324:1325, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1324: 'Curuá-Una'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_volu.xlsx')
        carga = ipdo.iloc[1324:1325, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1324: 'Curuá-Una'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSCuruá-Una_vertimento.xlsx')


def DhSRioJari(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Jari.
    -------
    """
    caminho_arquivo = (
       r'C:\Users\e806128\Desktop\1\DaHSUJari_vertimento.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_Afluença.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_Defluência.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_vertimento.xlsx')
        no = no + 3

        carga = ipdo.iloc[1325:1326, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1325: 'Sto Antônio do Jari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_Afluença.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_Afluença.xlsx',
            index=False)
        carga = ipdo.iloc[1325:1326, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1325: 'Sto Antônio do Jari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_Defluência.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_Defluência.xlsx',
            index=False)
        carga = ipdo.iloc[1325:1326, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1325: 'Sto Antônio do Jari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1325:1326, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1325: 'Sto Antônio do Jari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1325:1326, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1325: 'Sto Antônio do Jari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_vertimento.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_vertimento.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1325:1326, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1325: 'Sto Antônio do Jari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_Afluença.xlsx')
        carga = ipdo.iloc[1325:1326, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1325: 'Sto Antônio do Jari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_Defluência.xlsx')
        carga = ipdo.iloc[1325:1326, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1325: 'Sto Antônio do Jari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_Nível.xlsx')
        carga = ipdo.iloc[1325:1326, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1325: 'Sto Antônio do Jari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_volu.xlsx')
        carga = ipdo.iloc[1325:1326, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        carga1_d = carga1.transpose()
        carga1_re = carga1_d.rename(columns={1325: 'Sto Antônio do Jari'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSUJari_vertimento.xlsx')


def DHSRParaopeba(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Paraopeba.
    -------
    """
    caminho_arquivo = (
       r'C:\Users\e806128\Desktop\1\DaHSParaopeba_verti.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_Afl.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_Del.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_Nív.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_verti.xlsx')
        no = no + 3

        carga = ipdo.iloc[1326:1327, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1326: 'Retiro Baixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_Afl.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf.loc[no] = (dado1,
                           dado2,
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_Afl.xlsx',
            index=False)
        carga = ipdo.iloc[1326:1327, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1326: 'Retiro Baixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_Del.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_Del.xlsx',
            index=False)
        carga = ipdo.iloc[1326:1327, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1326: 'Retiro Baixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_Nív.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_Nív.xlsx',
            index=False)
        carga = ipdo.iloc[1326:1327, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1326: 'Retiro Baixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1326:1327, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1326: 'Retiro Baixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_verti.xlsx')
        dado2 = form_tab.iloc[0, 0]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_verti.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1326:1327, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1326: 'Retiro Baixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_Afl.xlsx')
        carga = ipdo.iloc[1326:1327, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1326: 'Retiro Baixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_Del.xlsx')
        carga = ipdo.iloc[1326:1327, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1326: 'Retiro Baixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})

        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_Nív.xlsx')
        carga = ipdo.iloc[1326:1327, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1326: 'Retiro Baixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_volu.xlsx')
        carga = ipdo.iloc[1326:1327, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1326: 'Retiro Baixo'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})

        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSParaopeba_verti.xlsx')


def DHSRXingu(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados Hidráulicos Rio Xingu.
    -------
    """
    caminho_arquivo = (
       r'C:\Users\e806128\Desktop\1\DaHSXingu_vertime.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_Afluen.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_Defluên.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_Nível.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_volu.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_vertime.xlsx')
        no = no + 3

        carga = ipdo.iloc[1327:1331, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1327: 'Belo Monte',
                                             1328: 'Pimental',
                                             1329: 'Canal Pereira Barreto',
                                             1330: 'Desvio Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_Afluen.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_Afluen.xlsx',
            index=False)
        carga = ipdo.iloc[1327:1331, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1327: 'Belo Monte',
                                             1328: 'Pimental',
                                             1329: 'Canal Pereira Barreto',
                                             1330: 'Desvio Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_Defluên.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_Defluên.xlsx',
            index=False)
        carga = ipdo.iloc[1327:1331, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1327: 'Belo Monte',
                                             1328: 'Pimental',
                                             1329: 'Canal Pereira Barreto',
                                             1330: 'Desvio Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_Nível.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_Nível.xlsx',
            index=False)
        carga = ipdo.iloc[1327:1331, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1327: 'Belo Monte',
                                             1328: 'Pimental',
                                             1329: 'Canal Pereira Barreto',
                                             1330: 'Desvio Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_volu.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_volu.xlsx',
            index=False)
        carga = ipdo.iloc[1327:1331, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1327: 'Belo Monte',
                                             1328: 'Pimental',
                                             1329: 'Canal Pereira Barreto',
                                             1330: 'Desvio Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_vertime.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_vertime.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1327:1331, 12:16]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1327: 'Belo Monte',
                                             1328: 'Pimental',
                                             1329: 'Canal Pereira Barreto',
                                             1330: 'Desvio Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_Afluen.xlsx')
        carga = ipdo.iloc[1327:1331, 12:17]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1327: 'Belo Monte',
                                             1328: 'Pimental',
                                             1329: 'Canal Pereira Barreto',
                                             1330: 'Desvio Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_Defluên.xlsx')
        carga = ipdo.iloc[1327:1331, 12:18]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1327: 'Belo Monte',
                                             1328: 'Pimental',
                                             1329: 'Canal Pereira Barreto',
                                             1330: 'Desvio Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_Nível.xlsx')
        carga = ipdo.iloc[1327:1331, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1327: 'Belo Monte',
                                             1328: 'Pimental',
                                             1329: 'Canal Pereira Barreto',
                                             1330: 'Desvio Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[6:7]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_volu.xlsx')
        carga = ipdo.iloc[1327:1331, 12:22]
        carga1 = carga.drop(carga.columns[1], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1327: 'Belo Monte',
                                             1328: 'Pimental',
                                             1329: 'Canal Pereira Barreto',
                                             1330: 'Desvio Jordão'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 21': data})
        form_tab = carga1_re1.iloc[8:9]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaHSXingu_vertime.xlsx')


def dados85(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados da bacia 85.
    -------
    """
    caminho_arquivo = (
       r'C:\Users\e806128\Desktop\1\DaBaci85_GerPro.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_Armaz.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_ENAdia.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_ENAArma.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_ENABrut.xlsx')
        add_inf5 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_GerVeri.xlsx')
        add_inf6 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_GerPro.xlsx')
        no = no + 3

        carga = ipdo.iloc[1335:1339, 10:13]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1335: 'TOC',
                                             1336: 'SFR',
                                             1337: 'PBA',
                                             1338: 'AMA'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 12': data})
        form_tab = carga1_re1.iloc[0:1]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_Armaz.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_Armaz.xlsx',
            index=False)
        carga = ipdo.iloc[1335:1339, 10:15]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1335: 'TOC',
                                             1336: 'SFR',
                                             1337: 'PBA',
                                             1338: 'AMA'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 14': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_ENAdia.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_ENAdia.xlsx',
            index=False)
        carga = ipdo.iloc[1335:1339, 10:16]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1335: 'TOC',
                                             1336: 'SFR',
                                             1337: 'PBA',
                                             1338: 'AMA'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_ENAArma.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_ENAArma.xlsx',
            index=False)
        carga = ipdo.iloc[1335:1339, 10:17]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1335: 'TOC',
                                             1336: 'SFR',
                                             1337: 'PBA',
                                             1338: 'AMA'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_ENABrut.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_ENABrut.xlsx',
            index=False)
        carga = ipdo.iloc[1335:1339, 10:18]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1335: 'TOC',
                                             1336: 'SFR',
                                             1337: 'PBA',
                                             1338: 'AMA'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[5:6]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_GerVeri.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf5.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf5 = add_inf5.rename(columns={'Unnamed: 0': ' '})
        add_inf5.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_GerVeri.xlsx',
            index=False)
        carga = ipdo.iloc[1335:1339, 10:20]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1335: 'TOC',
                                             1336: 'SFR',
                                             1337: 'PBA',
                                             1338: 'AMA'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[7:8]
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_GerPro.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        add_inf6.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5
                            )
        add_inf6 = add_inf6.rename(columns={'Unnamed: 0': ' '})
        add_inf6.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_GerPro.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1335:1339, 10:13]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1335: 'TOC',
                                             1336: 'SFR',
                                             1337: 'PBA',
                                             1338: 'AMA'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 12': data})
        form_tab = carga1_re1.iloc[0:1]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaBaci85_Armaz.xlsx', )
        carga = ipdo.iloc[1335:1339, 10:15]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1335: 'TOC',
                                             1336: 'SFR',
                                             1337: 'PBA',
                                             1338: 'AMA'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 14': data})
        form_tab = carga1_re1.iloc[2:3]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaBaci85_ENAdia.xlsx')
        carga = ipdo.iloc[1335:1339, 10:16]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1335: 'TOC',
                                             1336: 'SFR',
                                             1337: 'PBA',
                                             1338: 'AMA'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})
        form_tab = carga1_re1.iloc[3:4]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaBaci85_ENAArma.xlsx')
        carga = ipdo.iloc[1335:1339, 10:17]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1335: 'TOC',
                                             1336: 'SFR',
                                             1337: 'PBA',
                                             1338: 'AMA'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})
        form_tab = carga1_re1.iloc[4:5]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaBaci85_ENABrut.xlsx')
        carga = ipdo.iloc[1335:1339, 10:18]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1335: 'TOC',
                                             1336: 'SFR',
                                             1337: 'PBA',
                                             1338: 'AMA'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 17': data})
        form_tab = carga1_re1.iloc[5:6]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaBaci85_GerVeri.xlsx')
        carga = ipdo.iloc[1335:1339, 10:20]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1335: 'TOC',
                                             1336: 'SFR',
                                             1337: 'PBA',
                                             1338: 'AMA'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 19': data})
        form_tab = carga1_re1.iloc[7:8]
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaBaci85_GerPro.xlsx')


def Dadosbacia86(ipdo: pd.DataFrame):
    """
    data,colunassc,transpose,to_excel,n_index.
    ----------
    ipdo : DataFrame da tabela.

    Retorna um arquivo .xlsx em formato de tabela a respeito dos valores
    dos dados da bacia 86.
    -------
    """
    caminho_arquivo = (
       r'C:\Users\e806128\Desktop\1\DaBaci85_N.xlsx')
    no = 2
    if os.path.exists(caminho_arquivo):
        print("O arquivo existe!")
        add_inf = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Dadosbacias86_SE.xlsx')
        add_inf2 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\Dadosbacias86_s.xlsx')
        add_inf3 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci86_NE.xlsx')
        add_inf4 = pd.read_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_N.xlsx')
        no = no + 3

        carga = ipdo.iloc[1403:1424, 10:13]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1405: 'PN',
                                             1406: 'GR',
                                             1407: 'TI',
                                             1408: 'PP',
                                             1409: 'PR',
                                             1410: 'PB',
                                             1411: 'PY',
                                             1412: 'DC',
                                             1413: 'JE',
                                             1414: 'IG',
                                             1415: 'JI',
                                             1416: 'RI',
                                             1417: 'CA',
                                             1418: 'SF',
                                             1419: 'PI',
                                             1420: 'PG',
                                             1421: 'TO',
                                             1422: 'AM',
                                             1423: 'OUTRAS'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 12': data})

        form_tab1 = carga1_re1.iloc[0:1]
        colunassx = [0, 1]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\Dadosbacias86_SE.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        dado11 = form_tab.iloc[0, 9]
        dado12 = form_tab.iloc[0, 10]
        dado13 = form_tab.iloc[0, 11]
        dado14 = form_tab.iloc[0, 12]
        dado15 = form_tab.iloc[0, 13]
        dado16 = form_tab.iloc[0, 14]
        dado17 = form_tab.iloc[0, 15]
        dado18 = form_tab.iloc[0, 16]
        dado19 = form_tab.iloc[0, 17]
        dado20 = form_tab.iloc[0, 18]
        add_inf.loc[no] = (dado1,
                           dado2,
                           dado3,
                           dado4,
                           dado5,
                           dado6,
                           dado7,
                           dado8,
                           dado9,
                           dado10,
                           dado11,
                           dado12,
                           dado13,
                           dado14,
                           dado15,
                           dado16,
                           dado17,
                           dado18,
                           dado19,
                           dado20
                           )
        add_inf = add_inf.rename(columns={'Unnamed: 0': ' '})
        add_inf.to_excel(
            r'C:\Users\e806128\Desktop\1\Dadosbacias86_SE.xlsx',
            index=False)
        carga = ipdo.iloc[1403:1424, 10:15]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1405: 'PN',
                                             1406: 'GR',
                                             1407: 'TI',
                                             1408: 'PP',
                                             1409: 'PR',
                                             1410: 'PB',
                                             1411: 'PY',
                                             1412: 'DC',
                                             1413: 'JE',
                                             1414: 'IG',
                                             1415: 'JI',
                                             1416: 'RI',
                                             1417: 'CA',
                                             1418: 'SF',
                                             1419: 'PI',
                                             1420: 'PG',
                                             1421: 'TO',
                                             1422: 'AM',
                                             1423: 'OUTRAS'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 14': data})

        form_tab1 = carga1_re1.iloc[2:3]
        colunassx = [0, 1]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\Dadosbacias86_s.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        dado11 = form_tab.iloc[0, 9]
        dado12 = form_tab.iloc[0, 10]
        dado13 = form_tab.iloc[0, 11]
        dado14 = form_tab.iloc[0, 12]
        dado15 = form_tab.iloc[0, 13]
        dado16 = form_tab.iloc[0, 14]
        dado17 = form_tab.iloc[0, 15]
        dado18 = form_tab.iloc[0, 16]
        dado19 = form_tab.iloc[0, 17]
        dado20 = form_tab.iloc[0, 18]
        add_inf2.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9,
                            dado10,
                            dado11,
                            dado12,
                            dado13,
                            dado14,
                            dado15,
                            dado16,
                            dado17,
                            dado18,
                            dado19,
                            dado20
                            )
        add_inf2 = add_inf2.rename(columns={'Unnamed: 0': ' '})
        add_inf2.to_excel(
            r'C:\Users\e806128\Desktop\1\Dadosbacias86_s.xlsx',
            index=False)
        carga = ipdo.iloc[1403:1424, 10:16]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1405: 'PN',
                                             1406: 'GR',
                                             1407: 'TI',
                                             1408: 'PP',
                                             1409: 'PR',
                                             1410: 'PB',
                                             1411: 'PY',
                                             1412: 'DC',
                                             1413: 'JE',
                                             1414: 'IG',
                                             1415: 'JI',
                                             1416: 'RI',
                                             1417: 'CA',
                                             1418: 'SF',
                                             1419: 'PI',
                                             1420: 'PG',
                                             1421: 'TO',
                                             1422: 'AM',
                                             1423: 'OUTRAS'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab1 = carga1_re1.iloc[3:4]
        colunassx = [0, 1]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaBaci86_NE.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        dado11 = form_tab.iloc[0, 9]
        dado12 = form_tab.iloc[0, 10]
        dado13 = form_tab.iloc[0, 11]
        dado14 = form_tab.iloc[0, 12]
        dado15 = form_tab.iloc[0, 13]
        dado16 = form_tab.iloc[0, 14]
        dado17 = form_tab.iloc[0, 15]
        dado18 = form_tab.iloc[0, 16]
        dado19 = form_tab.iloc[0, 17]
        dado20 = form_tab.iloc[0, 18]
        add_inf3.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9,
                            dado10,
                            dado11,
                            dado12,
                            dado13,
                            dado14,
                            dado15,
                            dado16,
                            dado17,
                            dado18,
                            dado19,
                            dado20
                            )
        add_inf3 = add_inf3.rename(columns={'Unnamed: 0': ' '})
        add_inf3.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci86_NE.xlsx',
            index=False)
        carga = ipdo.iloc[1403:1424, 10:17]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1405: 'PN',
                                             1406: 'GR',
                                             1407: 'TI',
                                             1408: 'PP',
                                             1409: 'PR',
                                             1410: 'PB',
                                             1411: 'PY',
                                             1412: 'DC',
                                             1413: 'JE',
                                             1414: 'IG',
                                             1415: 'JI',
                                             1416: 'RI',
                                             1417: 'CA',
                                             1418: 'SF',
                                             1419: 'PI',
                                             1420: 'PG',
                                             1421: 'TO',
                                             1422: 'AM',
                                             1423: 'OUTRAS'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab1 = carga1_re1.iloc[4:5]
        colunassx = [0, 1]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_N.xlsx')
        dado2 = form_tab.iloc[0, 0]
        dado3 = form_tab.iloc[0, 1]
        dado4 = form_tab.iloc[0, 2]
        dado5 = form_tab.iloc[0, 3]
        dado6 = form_tab.iloc[0, 4]
        dado7 = form_tab.iloc[0, 5]
        dado8 = form_tab.iloc[0, 6]
        dado9 = form_tab.iloc[0, 7]
        dado10 = form_tab.iloc[0, 8]
        dado11 = form_tab.iloc[0, 9]
        dado12 = form_tab.iloc[0, 10]
        dado13 = form_tab.iloc[0, 11]
        dado14 = form_tab.iloc[0, 12]
        dado15 = form_tab.iloc[0, 13]
        dado16 = form_tab.iloc[0, 14]
        dado17 = form_tab.iloc[0, 15]
        dado18 = form_tab.iloc[0, 16]
        dado19 = form_tab.iloc[0, 17]
        dado20 = form_tab.iloc[0, 18]
        add_inf4.loc[no] = (dado1,
                            dado2,
                            dado3,
                            dado4,
                            dado5,
                            dado6,
                            dado7,
                            dado8,
                            dado9,
                            dado10,
                            dado11,
                            dado12,
                            dado13,
                            dado14,
                            dado15,
                            dado16,
                            dado17,
                            dado18,
                            dado19,
                            dado20
                            )
        add_inf4 = add_inf4.rename(columns={'Unnamed: 0': ' '})
        add_inf4.to_excel(
            r'C:\Users\e806128\Desktop\1\DaBaci85_N.xlsx',
            index=False)

    else:
        print("O arquivo não existe.")
        carga = ipdo.iloc[1403:1424, 10:13]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1405: 'PN',
                                             1406: 'GR',
                                             1407: 'TI',
                                             1408: 'PP',
                                             1409: 'PR',
                                             1410: 'PB',
                                             1411: 'PY',
                                             1412: 'DC',
                                             1413: 'JE',
                                             1414: 'IG',
                                             1415: 'JI',
                                             1416: 'RI',
                                             1417: 'CA',
                                             1418: 'SF',
                                             1419: 'PI',
                                             1420: 'PG',
                                             1421: 'TO',
                                             1422: 'AM',
                                             1423: 'OUTRAS'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 12': data})

        form_tab1 = carga1_re1.iloc[0:1]
        colunassx = [0, 1]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\Dadosbacias86_SE.xlsx')
        carga = ipdo.iloc[1403:1424, 10:15]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1405: 'PN',
                                             1406: 'GR',
                                             1407: 'TI',
                                             1408: 'PP',
                                             1409: 'PR',
                                             1410: 'PB',
                                             1411: 'PY',
                                             1412: 'DC',
                                             1413: 'JE',
                                             1414: 'IG',
                                             1415: 'JI',
                                             1416: 'RI',
                                             1417: 'CA',
                                             1418: 'SF',
                                             1419: 'PI',
                                             1420: 'PG',
                                             1421: 'TO',
                                             1422: 'AM',
                                             1423: 'OUTRAS'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 14': data})

        form_tab1 = carga1_re1.iloc[2:3]
        colunassx = [0, 1]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\Dadosbacias86_s.xlsx')
        carga = ipdo.iloc[1403:1424, 10:16]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1405: 'PN',
                                             1406: 'GR',
                                             1407: 'TI',
                                             1408: 'PP',
                                             1409: 'PR',
                                             1410: 'PB',
                                             1411: 'PY',
                                             1412: 'DC',
                                             1413: 'JE',
                                             1414: 'IG',
                                             1415: 'JI',
                                             1416: 'RI',
                                             1417: 'CA',
                                             1418: 'SF',
                                             1419: 'PI',
                                             1420: 'PG',
                                             1421: 'TO',
                                             1422: 'AM',
                                             1423: 'OUTRAS'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 15': data})

        form_tab1 = carga1_re1.iloc[3:4]
        colunassx = [0, 1]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaBaci86_NE.xlsx')
        carga = ipdo.iloc[1403:1424, 10:17]
        colunassc = [0, 1]
        carga1 = carga.drop(carga.columns[colunassc], axis=1)
        d_transp = carga1.transpose()
        carga1_re = d_transp.rename(columns={1405: 'PN',
                                             1406: 'GR',
                                             1407: 'TI',
                                             1408: 'PP',
                                             1409: 'PR',
                                             1410: 'PB',
                                             1411: 'PY',
                                             1412: 'DC',
                                             1413: 'JE',
                                             1414: 'IG',
                                             1415: 'JI',
                                             1416: 'RI',
                                             1417: 'CA',
                                             1418: 'SF',
                                             1419: 'PI',
                                             1420: 'PG',
                                             1421: 'TO',
                                             1422: 'AM',
                                             1423: 'OUTRAS'})
        carga1_re1 = carga1_re.rename(index={'Unnamed: 16': data})

        form_tab1 = carga1_re1.iloc[4:5]
        colunassx = [0, 1]
        form_tab = form_tab1.drop(form_tab1.columns[colunassx], axis=1)
        form_tab.to_excel(r'C:\Users\e806128\Desktop\1\DaBaci85_N.xlsx')


if __name__ == "__main__":
    ipdo = pd.read_excel(r"C:\Users\e806128\Downloads\IPDO-13-03-2024.xlsm",
                         sheet_name='IPDO')
dadosprimeiro(ipdo)
intercambio(ipdo)
internacional(ipdo)
itaipu(ipdo)
nordeste(ipdo)
roraima(ipdo)
norte(ipdo)
sc(ipdo)
termos(ipdo)
carga(ipdo)
dados_hidrologicos(ipdo)
valoresMDUTT(ipdo)
ValMDUTTNorte(ipdo)
valoresMDUTTSUL(ipdo)
valoresMDTTNordeste(ipdo)
valormedioUsinaT1(ipdo)
valormediusinaT2(ipdo)
usinascommaisrazao(ipdo)
geracaotermicaT1et2(ipdo)
diferencasentrecapadidades(ipdo)
deferencaentrecapacidadesoperativa(ipdo)
restriçãoemanutenção(ipdo)
To554(ipdo)
diferencasentreCIeA(ipdo)
demandasMaximasSin(ipdo)
submercado(ipdo)
dadoshidraulicosSINRIOCoRUmba(ipdo)
DHSRAraguai(ipdo)
DHSNRSMarcos(ipdo)
DHSRParanaiba(ipdo)
DHSRPardo(ipdo)
DHriogrande(ipdo)
DHSRVerde(ipdo)
DHSRclaro(ipdo)
DadosHidraulicosSinRioCorrent(ipdo)
DHSRioPiracicaba(ipdo)
DHSRStoantonio(ipdo)
DadoHidraulicosSinRioDoce(ipdo)
DHRPinheiros(ipdo)
RioGuarapiranga(ipdo)
dHidraulicosSinRioTiete(ipdo)
DHidraulicosSinRioTibagi(ipdo)
DadosHSinRioParanapanema(ipdo)
DHSRParaná(ipdo)
DHSJaguari(ipdo)
DHSRdPeixe(ipdo)
DHSRParaibuna(ipdo)
DHSRParaibadSul(ipdo)
dadosbacias(ipdo)
DHSRioManso(ipdo)
DHSRioItiquira(ipdo)
DHSRioCorrentes(ipdo)
DHSRioJauru(ipdo)
DHSRioJordan(ipdo)
DHSRioIguaçu(ipdo)
DHSrioCanoa(ipdo)
DHSRiopassoFundo(ipdo)
DHSRChapeco(ipdo)
DHSRioPelotas(ipdo)
DHSRioJacui(ipdo)
DHSRioTaquari(ipdo)
DHSRioCapivari(ipdo)
dadosbacia84(ipdo)
DadosHSRioTocantis(ipdo)
dadosHSRioPetro(ipdo)
DHSRioSãoFrancisco(ipdo)
dhSrioJequitinhonha(ipdo)
DhSRioParnaiba(ipdo)
DhSRioAripuana(ipdo)
DhSRioComemoracao(ipdo)
DhSRioJamari(ipdo)
DhSRioGuapore(ipdo)
DhSRioMadeira(ipdo)
DhSRioUatumã(ipdo)
DhSRioAraguari(ipdo)
DhSRioTelesPires(ipdo)
DhSRioCuruaUna(ipdo)
DhSRioJari(ipdo)
DHSRParaopeba(ipdo)
DHSRXingu(ipdo)
dados85(ipdo)
Dadosbacia86(ipdo)
