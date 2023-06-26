import pandas as pd
import numpy as np
import matplotlib as pl
import matplotlib.pyplot as plt
import io
import xlsxwriter

df = pd.read_csv("file_name_import.csv")

df.info()

#Gráficos: Separar data en bloques, armar tablas dinámicas y recién ahí graficar por pregunta.
g = df.drop(['Dimensión','Categoría','Promedio regional total'], axis=1)
g = g.fillna(0)
g.info()

groups = g['Pregunta'].unique()
groups

grouped = g.groupby(g['Pregunta'])

col = {}
  
for i in g['Pregunta'].unique():
    df_t = grouped.get_group(i)
    #df_t2 = df_t.drop(['Pregunta'], axis=1)
    df_t3 = df_t.pivot_table(columns="Respuesta")#, columns = "call_weekday", aggfunc = {"count_call":np.sum})
    df_t3.name='{}'.format(i)
    col[i] = df_t3


#importar y graficar cada dataframe

fig, ax = plt.subplots(figsize=(15,30))
workbook = xlsxwriter.Workbook('file_name_export.xlsx')
keys = list(col.keys())

for i in col: 
  ax = col[i].plot(kind='line')
  ax.set_title('{}'.format(i))
  #ax.set_xlabel('Pregunta')
  figname = '{}.jpg'.format(i)
  plt.legend(bbox_to_anchor=(1.04,1), loc="upper left")
  
  index = keys.index('{}'.format(i))
  wks1 = workbook.add_worksheet('{}'.format(index))
  wks1.write(0,0,'{}'.format(i))
  imgdata=io.BytesIO()
  plt.savefig(imgdata, format='png', bbox_inches='tight')
  wks1.insert_image(2,2, '', {'image_data': imgdata})

workbook.close()