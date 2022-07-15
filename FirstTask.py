import pandas as pd
import numpy as np
import json
from pandas import json_normalize
import datetime
import requests

file= r"C:\Users\DELL\Downloads\2022 07 12 Z_LISTE_INV.xlsx"
df=pd.read_excel(file)
#dropping the duplicate of columns
df=df.drop(df.columns[ [9,11,14,16] ],axis=1)

#rename columns
df.rename(columns = {'Doc.inven.':'inventory_doc',
                    'Article':'material',
                    'Désignation article':'designation',
                    'TyAr':'type',
                    'UQ':'unit',
                    'Mag.':'store',
                    'Fourn.':'supplier',
                    'Quantité théorique':'theoritical_quantity',
                    'Quantité saisie':'entred_quantity',
                    'écart enregistré':'deviation',
                    'Ecart (montant)':'deviation_cost',
                    'Dev..1':'dev',
                    'Div.':'division',
                    'Sup':'delete',
                    'Dte cptage':'date_catchment',
                    'Rectifié par':'corrected_by',
                    'Cpt':'catchment',
                    'Réf.inventaire':'refecrence_inventory',
                    'N° inventaire':'inventory_number',
                    'TyS':'Tys'},  inplace = True)


#Adding the year and week columns
#get current year and week
year=datetime.datetime.today().isocalendar()[0]
week=datetime.datetime.today().isocalendar()[1]
#insert year and week in first and second position
df.insert(0, 'year', year)
df.insert(1, 'week', week)

# Calculate KPI stock accuracy
df['stock_accuracy']=np.where(df['theoritical_quantity']!=0, (1-(df['theoritical_quantity']-df['entred_quantity']).abs()/df['theoritical_quantity'])*100, '')

#Number of rows in the dataframe
count=df.shape[0]


url = f'https://api.exchangerate.host/latest'
response = requests.get(url)
data = response.json()
df['rate']=df['dev'].map(data['rates'])



#currency conversion
df['deviation_cost_euro']=df['deviation_cost']*df['rate']

#new KPI   
df['type_of_measurement']=np.where( ( (df['unit']=='G') | (df['unit']=='KG') ), 'weighd','counted')
# Pour un article pesé : Stock accuracy = 100% si l’écart < 3% et l’écart < 250€ sinon la réf est considéré en écart
df['stock_accuracy']=np.where ( ( (df['type_of_measurement']=='weighd') & (df['deviation'] < 0.03) & (df['deviation_cost_euro']< 250)) , 1 , df['stock_accuracy'] )
# Pour un article compté : Stock accuracy = 100% si l’écart < 1% et l’écart < 250€ sinon la réf est considéré en écart
df['stock_accuracy']=np.where ( ( (df['type_of_measurement']=='counted') & (df['deviation'] < 0.01) & (df['deviation_cost_euro']< 250)) , 1 , df['stock_accuracy'] )

#deviation cost per division
cost_per_division =df.groupby(['year','week','division'])['deviation_cost_euro'].sum().reset_index() 
#total deviation cost euro
# total_deviation_cost=df['deviation_cost_euro'].sum() 
total_deviation_cost=df.agg(total_deviation_cost=('deviation_cost_euro','sum')) 
#Count per divison
count_per_division=df.groupby(['division']).agg(count=('week','count')).reset_index()
print(count_per_division)


# df.to_excel(r"C:\Users\DELL\Downloads\Final_data.xlsx")

