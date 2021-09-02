"""
Breyner Santos Ortiz
EJECUCCIÓN DE REPORTES DE PROGRAMACIÓN DE LA PRODUCCIÓN 

"""
import pandas as pd
import numpy as np

STF = pd.read_excel("EXCEL/CUADRO COLECCION STF GROUP.xlsm","HOY", header=0)
EXPLO_REF = pd.read_excel("EXCEL/EXPLOSION_COL_MEX.xlsx","EXPLOSION REFERENCIA",
                          header=1,skiprows=2,skipfooter=1,
                         usecols=["Coleccion", "Referencia","ItemResumen","Necesidad"]) 
EXPLO = pd.read_excel("EXCEL/EXPLOSION_COL_MEX.xlsx","EXPLOSION",header=0)
EXPLO_NEW = pd.merge(EXPLO_REF,EXPLO)
EXPLO_NEW["Vector Disponible"] = (EXPLO_NEW["InventarioCol"] + 
                                 EXPLO_NEW["InventarioMex"] -
                                 EXPLO_NEW["Necesidad"] -
                                 EXPLO_NEW["NecesidadProgramadoMex"] -
                                 EXPLO_NEW["NecesidadProgramadoCol"])

def Lib_Col(baseDatos):
    filtro = ""
    lib = input("Referencias liberadas? : ")
    col = input("Ingrese la colecciòn a consultar: ")

    if lib == "0":
        filtro = baseDatos[baseDatos["Colección"]==col]
    else:
        filtro = baseDatos[baseDatos["Lib Dño"]==lib]
        filtro = filtro[filtro["Colección"]==col]
    filtro["Referencia"] = filtro["Referencia"].str.upper()
    referencias = filtro["Referencia"].to_numpy()
    return referencias

REFERENCIAS = Lib_Col(STF)

def ref_tela_ok(baseDatos,baseDatos2,array):
    filtro = baseDatos2[baseDatos2["Referencia"].isin(array)]
    filtro = filtro[filtro["ItemResumen"].str.contains("MT")]
    filtro = filtro[~filtro["ItemResumen"].str.contains("ENTRETELA|FORRO BOLSILLO")]
    filtro = filtro[["Coleccion","Referencia","ItemResumen","Necesidad"]]
    filtro = filtro.sort_values(by="Necesidad",ascending=False)
    return filtro.to_excel("REFERENCIAS CON TELA OK.xlsx","TELA PRINCIPAL OK")

def explo_ref_tela_ok(baseDatos, baseDatos2,array):
    filtro = baseDatos2[baseDatos2["Referencia"].isin(array)]
    filtro = filtro[filtro["ItemResumen"].str.contains("MT")]
    filtro = filtro[~filtro["ItemResumen"].str.contains("ENTRETELA|FORRO BOLSILLO")]
    filtro = filtro[filtro["Vector Disponible"]>-45]
    ref2 = filtro["Referencia"].to_numpy()
    filtro2 = baseDatos2[baseDatos2["Referencia"].isin(ref2)]
    filtro2 = filtro2[~filtro2["ItemResumen"].str.contains("ENTRETELA|GANCHO|SECURITY|BOLSA")]
    filtro2 = filtro2[["Referencia","ItemResumen","Necesidad","Vector Disponible","NecesidadNoProgramadoCol","InventarioCol","NecesidadProgramadoCol","InventarioTransitoCol",]]
    filtro3 = filtro2.groupby(["Referencia","ItemResumen"])
    return filtro3.first().to_excel("EXPLOSION REFERENCIAS.xlsx","EXPLOSION_REFERENCIAS")

def telas(baseDatos,baseDatos2,array):
    filtro = baseDatos2[baseDatos2["Referencia"].isin(array)]
    filtro = filtro[filtro["ItemResumen"].str.contains("MT")]
    filtro = filtro[~filtro["ItemResumen"].str.contains("ENTRETELA|FORRO BOLSILLO")]
    filtro = filtro[["ItemResumen","Necesidad","InventarioCol","InvTintoreriaCol","InvZonaFrancaCol","InventarioTransitoCol"]]
    filtro = filtro.groupby(by="ItemResumen").agg({"Necesidad":"sum","InventarioCol":"min","InvTintoreriaCol":"min","InvZonaFrancaCol":"min","InventarioTransitoCol":"min"})
    filtro = filtro.sort_values(by=(["Necesidad"]), ascending=False)
    return filtro.to_excel("ANALISIS TELAS COLECCION.xlsx","ANALISIS TELAS")

ref_tela_ok(STF,EXPLO_NEW,REFERENCIAS)
explo_ref_tela_ok(STF,EXPLO_NEW,REFERENCIAS)
telas(STF,EXPLO_NEW,REFERENCIAS)

