import pandas as pd
import os
import time
caminho = r'C:\Users\e806128\Desktop\ipdos'
lista_arquivos = os.listdir(caminho)
lista_datas = []
nomes = []
ipdo1 = []
for arquivo in lista_arquivos:
    data = os.path.getmtime(F"{caminho}/{arquivo}")
    data = time.ctime(data)
    lista_datas.append((data, arquivo))
    df = pd.DataFrame(lista_datas)
    df1 = df.iloc[0, 0]
caminho = r"C:\Users\e806128\Desktop\ipdos"

caminhos_arquivos = []

if os.path.exists(caminho):
    arquivos_na_pasta = os.listdir(caminho)
    
    for arquivo in arquivos_na_pasta:
        caminho_completo = os.path.join(caminho, arquivo)
        if os.path.isfile(caminho_completo):
            caminhos_arquivos.append(caminho_completo)
print("Caminhos dos arquivos:")
for caminho in caminhos_arquivos:
    b,c,d,e,f,g,h,i,j = caminhos_arquivos
    
    b1 = pd.read_excel(b, sheet_name="IPDO")
    c1 = pd.read_excel(c, sheet_name="IPDO")
    d1 = pd.read_excel(d, sheet_name="IPDO")
    e1 = pd.read_excel(e, sheet_name="IPDO")
    f1 = pd.read_excel(f, sheet_name="IPDO")
    g1 = pd.read_excel(g, sheet_name="IPDO")
    h1 = pd.read_excel(h, sheet_name="IPDO")
    i1 = pd.read_excel(i, sheet_name="IPDO")
    j1 = pd.read_excel(j, sheet_name="IPDO")

    ipdo1 =[b1,c1,d1,e1,f1,g1,h1,i1,j1]
    
    b = os.path.basename(b).split('/')[-1]
    b = b.replace(".xlsm", "")
    c = os.path.basename(c).split('/')[-1]
    c = c.replace(".xlsm", "")
    d = os.path.basename(d).split('/')[-1]
    d = d.replace(".xlsm", "")
    e = os.path.basename(e).split('/')[-1]
    e = e.replace(".xlsm", "")
    f = os.path.basename(f).split('/')[-1]
    f = f.replace(".xlsm", "")
    g = os.path.basename(g).split('/')[-1]
    g = g.replace(".xlsm", "")
    h = os.path.basename(h).split('/')[-1]
    h = h.replace(".xlsm", "")
    i = os.path.basename(i).split('/')[-1]
    i = i.replace(".xlsm", "")
    j = os.path.basename(j).split('/')[-1]
    j = j.replace(".xlsm", "")

    print(b)
    nomes = [b,c,d,e,f,g,h,i,j]
    n = 0
    cont = 1 

    while cont < 9:
        for n, s in enumerate(zip(ipdo1,nomes)): 
            n = ipdo1[n]
            s = s[1]
            caminho_arquivo = r'C:\Users\e806128\Desktop\1\prograo_sin.xlsx'
            no = 2
            print("O arquivo não existe!")
            
            sin = n.iloc[6:16, 10:13]
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
            sin1_f.to_excel(r'C:\Users\e806128\Desktop\1\prograo_sin' + str(s) + '.xlsx')
            sin2 = n.iloc[6:16, 10:15]
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
            sin2_v.to_excel(r"C:\Users\e806128\Desktop\1\verifi_sin"+ str(s) +  '.xlsx') 
            cont += 1
