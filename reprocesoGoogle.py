import pandas as pd
import google as google
from comunes import agregar_time_stamp

def pd_filter(row):
    if row['awsEsri-CUST-TipoGeo'] == 6 or row['awsEsri-CUST-TipoGeo'] == "-":
        return "x"

dfp = pd.read_excel("./Lotes a procesar/Reproceso Google.xlsx", index_col=0, sheet_name="LocatMejorCand")
dfc = pd.read_excel("./Lotes a procesar/Reproceso Google.xlsx", index_col=0, sheet_name="FinalCliente")

dfc = dfc[dfc["awsEsri-CUST-TipoGeo"].isin([1, 2, 3, 4, 5, 12, '1', '2', '3', '4', '5', '12'])]
dfc.to_excel("./Lotes procesados/LoteSoloLocation.xlsx")

dfGoogle = dfp[dfp['awsEsri-CUST-TipoGeo'].isin(['6', '-', 6, 'E'])]
dfGoogle.to_excel("./Lotes procesados/LoteReprocesarGoogle.xlsx")

google.procesar_lote_google("./Lotes procesados/LoteReprocesarGoogle.xlsx")










