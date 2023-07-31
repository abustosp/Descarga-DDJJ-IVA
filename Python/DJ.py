import tabula
import pandas as pd

# Leer el PDF 'DJF2002V230-178770024-1690580791043.pdf'
df = tabula.read_pdf('DJF2002V230-178770024-1690580791043.pdf', pages='all')

Saldos = df[1]

#Convertir la primera fila en el encabezado
Saldos.columns = Saldos.iloc[0]

#Eliminar la primera fila
Saldos.drop(Saldos.index[0], inplace=True)

Saldos['Importe'] = Saldos['Importe'].str.replace('$ ', '').astype(float)

Saldos.to_excel('DJF2002V230-178770024-1690580791043.xlsx', index=False)