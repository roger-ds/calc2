import pandas as pd

# df = pd.read_excel('CV_142_Autopecas_cadastramento.xlsx')
# df =  df[['CEST', 'DESCRIÇÃO']]
# lista = df.values.tolist()
# for item in lista:
#     print(item[0], item[1])


# df = pd.read_excel('CV_142_Autopecas_cadastramento.xlsx')
# df =  df[['CEST', 'NCM/SH']]
# lista = df.values.tolist()
# for item in lista:
#     cest, ncm = str(item[0]).replace('.', ''), str(item[1]).replace('.', '' ).split(' ')
#     for item_ncm in ncm:
#         print(cest, item_ncm)


df = pd.read_excel('CV_142_Ferramentas_cadastramento.xlsx')
df =  df[['CEST', 'NCM/SH']]
df.dropna(inplace=True)
lista = df.values.tolist()
for item in lista:
    cest, ncm = str(item[0]).replace('.', ''), str(item[1]).replace('.', '' ).split(' ')
    for item_ncm in ncm:
        print(cest, item_ncm)