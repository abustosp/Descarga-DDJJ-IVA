import tabula
import pandas as pd

# Leer el Excel que posee las ubicaciones
Excel = pd.read_excel('Listado DDJJ IVA.xlsx')

# Reemplazar los \ por / en la columna 'Ubicación Descarga'
Excel['Ubicación Descarga'] = Excel['Ubicación Descarga'].astype(str).str.replace("\\", "/")

# hacer un for con los items del Excel
for i in range(len(Excel)):


    # Leer el PDF concatenando las columnas 'Ubicación Descarga' y 'DDJJ IVA'
    df = tabula.read_pdf((Excel['Ubicación Descarga'][i] + Excel['DDJJ IVA'][i] + ".pdf"), pages='all')

    Saldos = df[1]

    #Convertir la primera fila en el encabezado
    Saldos.columns = Saldos.iloc[0]

    #Eliminar la primera fila
    Saldos.drop(Saldos.index[0], inplace=True)

    Saldos['Importe'] = Saldos['Importe'].str.replace('$ ', '').astype(float)

    Saldos.to_excel(Excel['Ubicación Descarga'][i] + Excel['DDJJ IVA'][i] + ".xlsx", index=False)