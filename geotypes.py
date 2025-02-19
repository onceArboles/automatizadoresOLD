import pandas as pd
from comunes import *
"""
def location_address_pe(un_archivo: str):
    dataframe = pd.read_excel(un_archivo)
    
    #lote_intermedio = dataframe[(dataframe["awsEsri-orden"] == 1)]
    #Esta hacea que solo se seleccione el primer candidato

    dataframe["Tipo-Geo"] = dataframe.apply(geotype_pandas_basico, axis=1)
    limpiar_pantalla()
    dataframe["Estado"] = dataframe.apply(estado_pandas_basico, axis=1)
    limpiar_pantalla()
    dataframe["Motivo"] = dataframe.apply(motivo_basico, axis=1)
    limpiar_pantalla()
    
    dataframe = dataframe.sort_values(by=['awsEsri-id', 'Tipo-Geo'])

    dataframe.to_clipboard()
    lote_final = pd.read_clipboard()
    
    
    #lote_final = lote_final.drop(columns=["awsEsri-singleLine", "awsEsri-Label",
    #                                              "awsEsri-Interpolated", "awsEsri-Categories", "awsEsri-Relevance", "awsEsri-Json"])
    
    lote_final = lote_final.rename(columns={"awsEsri-Latitud":"Latitud", "awsEsri-Longitud":"Longitud", "awsEsri-AddressNumber":"Altura","awsEsri-Street":"Calle",
                                            "awsEsri-Region":"Level2", "awsEsri-Subregion":"Level3", "awsEsri-Municipality":"Level4", "awsEsri-Neighborhood":"Level5",
                                            "awsEsri-PostalCode":"Código Postal", "awsEsri-id":"ID Cliente"})
    
    nombre_archivo = agregar_time_stamp('./Lotes procesados/AddressPE-LocationEsri - ')
        
    lote_final.to_excel(nombre_archivo)
    
    print("Se ha generado el archivo de normalización de Address PE utilizando AWS Location - ESRI")
    

def geotype_pandas_basico(fila):
    if (fila["awsEsri-Interpolated"] == False) & (fila["awsEsri-Categories"] == "AddressType") & (fila["awsEsri-Relevance"] >= 0.85):
        return "1"
    else:
        if (fila["awsEsri-Interpolated"] != False) & (fila["awsEsri-Categories"] == "AddressType") & (fila["awsEsri-Relevance"] >= 0.85):
            return "3"
        else: 
            if (fila["awsEsri-Categories"] == "StreetType") & (fila["awsEsri-Relevance"] >= 0.85):
                return "5"
            else:
                if (fila["awsEsri-Categories"] == "PointOfInterestType") & (fila["awsEsri-Relevance"] >= 0.85):
                    return "2"
                else:
                    if (fila["awsEsri-Categories"] == "MunicipalityType"):
                        return "6"
                    else:
                        if (fila["awsEsri-Categories"] == "IntersectionType") & (fila["awsEsri-Relevance"] >= 0.85):
                            return "4"
"""                        
def location_esri_PE(una_category: str, una_relevance: float, un_interpolated:bool):
    #El primer elemento de la tupla devuelta es el Geotype y el segundo corresponde al orden evaluado 
    if (un_interpolated == False) & (una_category == "AddressType") & (una_relevance >= 0.85):
        return ["1",1]
    else:
        if (un_interpolated == True) & (una_category == "AddressType") & (una_relevance >= 0.85):
            return ["3",2]
        else: 
            if (una_category == "StreetType") & (una_relevance >= 0.85):
                return ["5",4]
            else:
                if (una_category == "PointOfInterestType") & (una_relevance >= 0.85):
                    return ["2",9]
                else:
                    if (una_category == "MunicipalityType"):
                        return ["6",6]
                    else:
                        if (una_category == "IntersectionType") & (una_relevance >= 0.85):
                            return ["4",3]
                        else:
                            if (una_category == "NeighborhoodType") & (una_relevance >= 0.85):
                                return ["12",5]
                            else:
                                return ["-",99]
                        
def location_here_PE(una_category:str, un_interpolated:bool):
    if (un_interpolated == False) & (una_category == "AddressType"):
        return ["1",1]
    else:
        if (un_interpolated == True) & (una_category == "AddressType"):
            return ["3",2]
        else: 
            if (una_category == "StreetType"):
                return ["5",4]
            else:
                if (una_category == "PointOfInterestType"):
                    return ["2",9]
                else:
                    if (una_category == "MunicipalityType"):
                        return ["6",6]
                    else:
                        if (una_category == "IntersectionType"):
                            return ["4",3]
                        else:
                            if (una_category == "NeighborhoodType"):
                                return ["12",5]
                            else:
                                return ["-",99]

def google_all (un_type: str, una_route: str, un_street_number: str):
    
    if ("street_address" in un_type or "premise" in un_type or "subpremise" in un_type) and (una_route != None) and (un_street_number != None):
        return "1"
    else:
        if (("route" in un_type) and (un_street_number == None)):
            return "5"
        else:
            if (("route" in un_type) and (un_street_number != None)):
                return "3"
            else:
                if ("neighborhood" in un_type):
                      return "12"
                else:
                    if ("locality" in un_type or "postal_code" in un_type or "sublocality" in un_type):
                        return "6"
                    else:
                        if ("intersection" in un_type):
                                return "4"
                        else:
                            return "-"
    

    




"""                        
def estado_pandas_basico(fila):
    if (fila["Tipo-Geo"] != None):
        return "CO"
    else:
        return "NE"
    
"""
                    
    
    
    