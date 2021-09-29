import pandas as pd

path1 = "\\\\atlas00\\Root\\Diffusion\\IT_Projects\\056-ETSS-Transformation\\056-ETSS-Transformation-Work\\0-AvantProjet\\19-PVID_Subchamber_Defect_FDC_PCM\\10-MetrologyProcessoperPVIDdescription\\50-META\\20210927_recap_ME_Endura.xlsx"
path2 = "\\\\atlas00\\Root\\Diffusion\\IT_Projects\\056-ETSS-Transformation\\056-ETSS-Transformation-Work\\0-AvantProjet\\19-PVID_Subchamber_Defect_FDC_PCM\\10-MetrologyProcessoperPVIDdescription\\50-META\\20210929_recap_output.xlsx"

xls = pd.ExcelFile(path1)
df1 = pd.read_excel(path1,'REP')
df2 = pd.read_excel(path1,'PVID')

df = pd.merge(df1, df2, how='outer', on=['TOOL', 'SLOT'])

df = df[['TOOL', 'SLOT','CHAMBRE','OPER','PVID', 'COMMENTARY']]

df['MOPER1'] = df['OPER']

for i in df.index:
    df.loc[i,'OPER'] = str(df.loc[i,'OPER'])

for i in df.index:
    if df.loc[i,'COMMENTARY'] != 'Pre-measurement':
        df.loc[i, 'MOPER1'] = ''

df['POPER'] = df['OPER']
for i in df.index:
    if df.loc[i, 'POPER'].find('CAR') != 3:
        df.loc[i, 'POPER'] = ''

df['MOPER2/MOPER'] = df['OPER']
for i in df.index:
    if df.loc[i,'MOPER1'] == df.loc[i, 'OPER'] or df.loc[i,'POPER'] == df.loc[i, 'OPER']:
        df.loc[i, 'MOPER2/MOPER'] = ''

for i in df.index:
    for j in df.index:
        if df.loc[i,'POPER'] == df.loc[i,'OPER'] and df.loc[j,'TOOL'] == df.loc[i,'TOOL'] and df.loc[j,'SLOT'] == df.loc[i,'SLOT'] and df.loc[j,'POPER'] == '':
            df.loc[j, 'POPER'] = df.loc[i,'POPER']

df.drop(columns=['OPER'])

df = df[['TOOL','CHAMBRE','SLOT','MOPER1','POPER','MOPER2/MOPER', 'PVID','COMMENTARY']]

# for i in df.index:
#     if df.loc[i,'MOPER1'] != '':
#         df.loc[i-1, 'MOPER1'] = df.loc[i,'MOPER1']
#         df.drop(df.index[i])

for i in df.index:
    if df.loc[i,'MOPER2/MOPER'] == 'nan':
        df.drop(df.index[i])


df.to_excel(path2 ,index = False)
