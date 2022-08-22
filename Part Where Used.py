#python 3.8

import pandas as pd

#read the Cognos downloaded BOMs - BOM Remove Depth report
Lamp_BOM = pd.read_excel(r'C:\Users\dwu\OneDrive - Lutron Electronics Co., Inc\Lamp pipeline\Latest BOMs\BOM_Lamp.xlsx',dtype = {'Current Component Item Number':str})
D3_LV_BOM = pd.read_excel(r'C:\Users\dwu\OneDrive - Lutron Electronics Co., Inc\Lamp pipeline\Latest BOMs\BOM_D3 PS LV.xlsx',dtype = {'Current Component Item Number':str})
D3_HV_BOM = pd.read_excel(r'C:\Users\dwu\OneDrive - Lutron Electronics Co., Inc\Lamp pipeline\Latest BOMs\BOM_D3 PS HV.xlsx',dtype = {'Current Component Item Number':str})
#Off Cognos BOMs, manually maintained
G2_Graze_BOM = pd.read_excel(r'C:\Users\dwu\OneDrive - Lutron Electronics Co., Inc\Lamp pipeline\Latest BOMs\G2 Graze_BOM.xlsx',dtype = {'Current Component Item Number':str})
G2_NGraze_BOM = pd.read_excel(r'C:\Users\dwu\OneDrive - Lutron Electronics Co., Inc\Lamp pipeline\Latest BOMs\G2 Non Graze_BOM.xlsx',dtype = {'Current Component Item Number':str})
N3_BOM = pd.read_excel(r'C:\Users\dwu\OneDrive - Lutron Electronics Co., Inc\Lamp pipeline\Latest BOMs\N3 BOM.xlsx',dtype = {'Current Component Item Number':str})


Lamp_BOM['Product'] = 'Lamp'
D3_LV_BOM['Product'] = 'D3_LV'
D3_HV_BOM['Product'] ='D3_HV'
G2_Graze_BOM['Product'] = 'G2 G'
G2_NGraze_BOM['Product'] = 'G2 NG'
N3_BOM['Product'] = 'N3'



Agg_BOM_1 = pd.concat([Lamp_BOM,D3_HV_BOM,D3_LV_BOM])
df1 = Agg_BOM_1.groupby(['Current Component Item Number','Current Component Item Description','Current Component Item Commodity Class Description'])['Product'].apply(','.join).reset_index()

Agg_BOM_2 = pd.concat([G2_NGraze_BOM,G2_Graze_BOM,N3_BOM])
df2 = Agg_BOM_2.groupby(['Current Component Item Number','Current Component Item Description'])['Product'].apply(', '.join).reset_index()


df = pd.concat([df1,df2])
df = df.groupby(['Current Component Item Number','Current Component Item Description'])['Product'].apply(', '.join).reset_index()



df.to_excel(r'C:\Users\dwu\OneDrive - Lutron Electronics Co., Inc\Lamp pipeline\Latest BOMs\Where Used.xlsx')



