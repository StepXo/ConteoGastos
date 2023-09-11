import pandas as pd
import numpy as np
import locale
import tkinter.messagebox as mb
import sys
# Establecer el idioma español
locale.setlocale(locale.LC_TIME,"es_MX")
datos = pd.read_excel('Datos en Bruto.xlsx', sheet_name=0)
horario = pd.read_excel('Datos en Bruto.xlsx', sheet_name=1)

#agregar el dia domingo al horario, esto para temas de practicidad
domingo = pd.DataFrame({"Dia": "Domingo"}, index=[0])
horario = pd.concat([horario,domingo],ignore_index=True,axis=0)
#ordenar los datos por si depronto estan des ordenados
datos = datos.sort_values(by="Fecha(dd/mm/YYYY)", ascending=True).reset_index(drop=True)
# Crear una máscara booleana para filtrar las fechas que están dentro del rango deseado
try:
 start_date = datos.loc[0,"Fecha(dd/mm/YYYY)"]
 end_date = datos.loc[len(datos["Fecha(dd/mm/YYYY)"])-1,"Fecha(dd/mm/YYYY)"]
except KeyError:
  # crear un botón que muestra una ventana emergente de error
  mb.showerror("Error", "EL DOCUMENTO NO PUEDE ESTAR VACÍO")
  sys.exit()

date_range = pd.date_range(start_date, end_date)
df=pd.DataFrame(date_range, columns=["Fecha"])
rango = len(date_range)//7 + 1
# Agregar dos columnas más con los nombres de los días y de los meses en español
df["Mes"] = df["Fecha"].dt.strftime("%B")
df["Dia"] = df["Fecha"].dt.strftime("%A %d")

#Re ordenar el dataframe horario de modo que inicie en el pimer dia del dataframe df
Semana = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
primer_dia = df.loc[0,"Dia"][:-3]

index = Semana.index(primer_dia)
Semana = Semana[index:]+Semana[:index]

horario.Dia = pd.Categorical(horario.Dia, Semana)

horario = horario.sort_values(by=["Dia"])
# Hacer un merge entre horario y new_df usando la columna de fecha
new_df = pd.DataFrame(np.tile(horario, (rango, 1)), columns=horario.columns)

df["Factor"] = new_df["Factor"]
df["Pasaje"] = df["Factor"].where(df["Factor"].isnull(), 2880) # asignar 1260 si Factor es nulo
df["Razon"] = df["Factor"].where(df["Factor"].isnull(), "Pasaje")
df = df.reindex(columns=["Fecha","Mes","Dia","Ingreso","Motivo","Pasaje","Factor","Gasto","Razon"])

for i in range(datos.shape[0]):
  if datos.loc[i,"signo"] == "-":
    df.loc[df.Fecha == datos.loc[i,"Fecha(dd/mm/YYYY)"], "Gasto"] = datos.loc[i,"Monto"]
    if df.loc[i,"Razon"] == "Pasaje":
      df.loc[df.Fecha == datos.loc[i,"Fecha(dd/mm/YYYY)"], "Razon"] = df.loc[i,"Razon"] + " y " + datos.loc[i,"Motivo"]
    else:
      df.loc[df.Fecha == datos.loc[i,"Fecha(dd/mm/YYYY)"], "Razon"] = datos.loc[i,"Motivo"]
  else:
    df.loc[df.Fecha == datos.loc[i,"Fecha(dd/mm/YYYY)"], "Ingreso"] = datos.loc[i,"Monto"]
    df.loc[df.Fecha == datos.loc[i,"Fecha(dd/mm/YYYY)"], "Motivo"] = datos.loc[i,"Motivo"]

df_aux = df[["Pasaje","Factor","Gasto"]]
df_aux = df_aux.fillna(0.0)
df["Total"] = df_aux['Pasaje'] * df_aux["Factor"] + df_aux["Gasto"]
df["GASTO TOTAL MES"] = None  

total_mes = df.groupby(['Mes'])['Total'].sum().reset_index()
for i in range(total_mes.shape[0]):
  df.loc[df.Mes == total_mes.loc[i,"Mes"], "GASTO TOTAL MES"] = total_mes.loc[i,"Total"]

df = df.dropna(subset=['Motivo','Razon'], how = "all")
print(df)
df.to_excel("seguimiento de gastos.xlsx", sheet_name="Gastos")