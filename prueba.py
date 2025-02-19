import pandas as pd

df = pd.read_excel("Reproceso Google.xlsx", index_col=0, sheet_name="LocatMejorCand")

dfGoogle = df[not df['awsEsri-CUST-TipoGeo'].isin(['6', '-', 6])]

print(dfGoogle.info())