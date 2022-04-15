from datetime import datetime
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import pandas as pd

from constants import PATH

def get_trimestre(chaine):
    annee = chaine[:4]
    mois = int(chaine[5:7])
    vtrim = 0
    if 0 < mois <= 3:
        vtrim = 1
    elif 3 < mois <= 6:
        vtrim = 2
    elif 6 < mois <= 9:
        vtrim = 3
    else: vtrim = 4
    trim = annee+f"{vtrim}"
    return trim
def get_anneemois(chaine):
    annee = chaine[:4]
    mois = chaine[5:7]
    am = annee+mois
    return am
def get_ecretement(cout, branche):
    branches = [1,2,3,4,7]
    if branche in branches:
        criteres = {
            1: 30,
            2: 10,
            3: 40,
            4: 10,
            7: 50
        }
        cb = criteres[branche]*1000000
        vecrete = cout if cout < cb else cb
        return vecrete
    else: return cout
def deter(df, col):
    dc = df[col]
    vc = list(set(dc.to_list()))
    return sorted(vc)

def getData(source, filetype="csv"):
    if filetype == "csv":
        hs = pd.read_csv(source)
        df = pd.DataFrame(hs)
        return df
    elif filetype == "excel":
        h = pd.read_excel(source)
        df = pd.DataFrame(h)
        return df
    else:
        print("FILETYPE IS NOT RECOGNIZED...")
        return None


def getCouples(df, col1, col2):
    surv = deter(df, col1)
    decl = deter(df, col2)
    sude = surv
    sude.extend(decl)
    sude = sorted(list(set(sude)))
    couples = []
    for an_s in sude:
        andec = []
        indan = sude.index(an_s)
        andec = sude[indan:]
        for an_d in andec:
            couple = (an_s,an_d)
            couples.append(couple)
    return couples

def getTcd(couples, dm, col1, col2):
    lignes = []
    for couple in couples:
        an1 = couple[0]
        an2 = couple[1]
        # dataframe du couple
        dmc = dm.loc[(dm[col1] == an1) & (dm[col2] == an2)]
        nb_sinistres = dmc.shape[0]
        dmcq = dmc.groupby("CLE").sum()
        coutotal = dmcq["COUT TOTAL"].sum()
        ligne = {
            "CLE_TCD": "{}{}".format(an1,an2),
            col1 : an1,
            col2 : an2,
            "NBSINISTRE": nb_sinistres,
            "COUT_TOTAL": coutotal
        }
        lignes.append(ligne)
        tcd = pd.DataFrame(lignes)
    return tcd

def mul(liste):
    result = 1
    for v in liste:
        result *= v
    return result 

def get_cumul_sum(vl):
    vs = []
    for i in range(len(vl)):
        vc = vl[:i+1]
        vs.append(sum(vc))
    return vs

def get_cumul_mul(vl):
    vs = []
    for i in range(len(vl)):
        vc = vl[:i+1]
        vs.append(mul(vc))
    return vs

def get_factors(vl):
    vf = []
    for i in range(len(vl)-1):
        val = 0
        try:
            val = vl[i+1]/vl[i]
        except ZeroDivisionError:
            val = vl[i+1]/1
        vf.append(val)
    return vf

def getTriangleType1(df, col1, col2):
    lignes  = []
    survs = deter(df,col1)
    decls = deter(df,col2)
    for surv in survs:
        ds = df.loc[df[col1] == surv]
        ligne = {}
        for decl in decls:
            vs = ds.loc[ds[col2] == decl,"NBSINISTRE"].to_list()
            ligne[decl] = 0 if len(vs) == 0 else vs[0]
        lignes.append(ligne)
    triangle = pd.DataFrame(lignes)
    triangle.index = survs
    return triangle

def getTriangleType2(triangle1):
    nc = triangle1.shape[1]
    lignes = []
    for i in range(nc):
        nci = nc-i
        ligne = triangle1.iloc[i].to_list()[-nci:]
        ligne.extend([0 for _ in range(i)])
        lignes.append(ligne)
    triangle = pd.DataFrame(columns=triangle1.columns.to_list(), index=triangle1.index, data=lignes)
    return triangle

def getTriangleType3(triangle2):
    nc = triangle2.shape[1]
    lignes = []
    for i in range(nc):
        nci = nc-i
        ligne = triangle2.iloc[i].to_list()[:nci]
        ligne = get_cumul_sum(ligne)
        ligne.extend([0 for _ in range(i)])
        lignes.append(ligne)
    triangle = pd.DataFrame(columns=triangle2.columns.to_list(), index=triangle2.index, data=lignes)
    return triangle

def getTriangleFactor1(triangle3):
    nc = triangle3.shape[1]
    lignes = []
    fcols  = []
    for i in range(nc-1):
        nci = nc-i
        fcols.append("{}_to_{}".format(i+1,i+2))
        ligne = triangle3.iloc[i].to_list()[:nci]
        ligne = get_factors(ligne)
        ligne.extend([0 for _ in range(i)])
        lignes.append(ligne)
    triangle = pd.DataFrame(columns=fcols, index=triangle3.iloc[:nc-1].index, data=lignes)
    return triangle

def getTriangleFactor2(factor1):
    nc = factor1.shape[1]
    lignes = []
    for i in range(nc):
        nci = nc-i
        ligne = factor1.iloc[i].to_list()[:nci]
        ligne = get_cumul_mul(ligne)
        ligne.extend([0 for _ in range(i)])
        lignes.append(ligne)
    triangle = pd.DataFrame(columns=factor1.columns.to_list(), index=factor1.index, data=lignes)
    return triangle

def getFactorFinal(factor2):
    nl = factor2.shape[0]
    lignes = []
    ligne = []
    fcols = []
    for i in range(nl):
        fcols.append("{}".format(i+2))
        if i != nl-1:
            col1 = "{}_to_{}".format(i+1,i+2)
            col2 = "{}_to_{}".format(i+2,i+3)
            vcl1 = mul(factor2[col1].to_list()[:nl-i-1])
            vcl2 = mul(factor2[col2].to_list()[:nl-i-1])
            vci = 0
            try:
                vci = vcl2/vcl1
            except ZeroDivisionError:
                vci = vcl1/1
            ligne.append(vci)
        else:
            ligne.append(1)
    lignes.append(ligne)
    facteurs = pd.DataFrame(data=lignes, columns=fcols)
    return facteurs

def getallfactors(df, col1, col2):
    triangle1 = getTriangleType1(df, col1, col2)
    triangle2 = getTriangleType2(triangle1)
    triangle3 = getTriangleType3(triangle2)
    facteur1 = getTriangleFactor1(triangle3)
    facteur2 = getTriangleFactor2(facteur1)
    facteurs = getFactorFinal(facteur2)
    return facteurs

def getTriangleType4(triangle3, factors):
    nc = triangle3.shape[1]
    lignes = []
    facteurs = factors.iloc[0].to_list()
    for i in range(nc):
        nci = nc-i
        ligne = triangle3.iloc[i].to_list()[:nci]
        if i != 0:
            for j in range(i):
                k = -i + j
                ligne.append(round(ligne[-1]*facteurs[k]))
        lignes.append(ligne)
    triangle = pd.DataFrame(columns=triangle3.columns.to_list(), index=triangle3.index, data=lignes)
    return triangle


def prepare_data(file):
    S = 1000
    A = -S
    B = S
    
    #CHARGEMENT DE NOTRE DATAFRAME
    data = getData(file, "excel")
    #SELECTION DE NOS COLONNES
    acols = ["ANNEE","INTER","NO SINISTRE","BRANCHE","LIBELLE BRANCHE","SORT SINISTRES","LIBELLE SORT SINISTRES","DATESURV","DATEDECL","SAP","REGLEMENTS","ESTIMATION RECOURS","RECOURS ABOUTIS","COUT TOTAL"]
    df = data[acols] 
    #PREPARATION
    df["ANNEE"] = df["ANNEE"].astype(str)
    df["INTER"] = df["INTER"].astype(str)
    df["NO SINISTRE"] = df["NO SINISTRE"].astype(str)
    df["DATESURV"] = df["DATESURV"].astype(str)
    df["DATEDECL"] = df["DATEDECL"].astype(str)
    df["CLE"] = df["ANNEE"] + df["INTER"] + df["NO SINISTRE"]
    df["ANNEESURV"] = df["DATESURV"].apply(lambda x:x[:4])
    df["ANNEEDECL"] = df["DATEDECL"].apply(lambda x:x[:4])
    df["TRIMSURV"] = df["DATESURV"].apply(lambda x:get_trimestre(x))
    df["TRIMDECL"] = df["DATEDECL"].apply(lambda x:get_trimestre(x))
    df["MOISSURV"] = df["DATESURV"].apply(lambda x:get_anneemois(x))
    df["MOISDECL"] = df["DATEDECL"].apply(lambda x:get_anneemois(x))
    df = df.loc[(df["COUT TOTAL"] <= A) | (df["COUT TOTAL"] >= B)]
    cibranches = list(range(50,60))
    df.loc[df["BRANCHE"].isin(cibranches), "BRANCHE"] = 5

    return df

def traitementv4(data_of_branche):
    dsm = data_of_branche.copy()
    # TCD ANNUEL
    col1_a = "ANNEESURV"
    col2_a = "ANNEEDECL"
    couples_a = getCouples(dsm, col1_a, col2_a)
    tcd_a = getTcd(couples_a,dsm,col1_a,col2_a)
    #TCD TRIMESTRIEL
    col1_t = "TRIMSURV"
    col2_t = "TRIMDECL"
    couples_t = getCouples(dsm, col1_t, col2_t)
    tcd_t = getTcd(couples_t,dsm,col1_t,col2_t)
    #TCD MENSUEL
    col1_m = "MOISSURV"
    col2_m = "MOISDECL"
    couples_m = getCouples(dsm, col1_m, col2_m)
    tcd_m = getTcd(couples_m,dsm,col1_m,col2_m)

    #TRIANGLES && FACTEURS ANNUELS
    tri_a1 = getTriangleType1(tcd_a, col1_a, col2_a)
    tri_a2 = getTriangleType2(tri_a1)
    tri_a3 = getTriangleType3(tri_a2)
    fct_a1 = getTriangleFactor1(tri_a3)
    fct_a2 = getTriangleFactor2(fct_a1)
    facteurs_a = getFactorFinal(fct_a2)
    tri_a4 = getTriangleType4(tri_a3,facteurs_a)

    #TRIANGLES ET FACTEURS TRIMESTRIELS
    tri_t1 = getTriangleType1(tcd_t, col1_t, col2_t)
    tri_t2 = getTriangleType2(tri_t1)
    tri_t3 = getTriangleType3(tri_t2)
    fct_t1 = getTriangleFactor1(tri_t3)
    fct_t2 = getTriangleFactor2(fct_t1)
    facteurs_t = getFactorFinal(fct_t2)
    tri_t4 = getTriangleType4(tri_t3,facteurs_t)

    #TRIANGLES ET FACTEURS MENSUELS
    tri_m1 = getTriangleType1(tcd_m, col1_m, col2_m)
    tri_m2 = getTriangleType2(tri_m1)
    tri_m3 = getTriangleType3(tri_m2)
    fct_m1 = getTriangleFactor1(tri_m3)
    fct_m2 = getTriangleFactor2(fct_m1)
    facteurs_m = getFactorFinal(fct_m2)
    tri_m4 = getTriangleType4(tri_m3,facteurs_m)

    #ENREGISTREMENT
    nbranche = deter(dsm, "BRANCHE")[0]
    fnow = datetime.now().strftime("%d-%m-%Y_%H:%M:%S")
    filename = "sortie_sanlam_branche_{}_{}.xlsx".format(nbranche,fnow)
    filepath = PATH+filename
    writer = pd.ExcelWriter(filepath)

    # CALCUL ANNUEL
    tri_a1.to_excel(writer, "TRIANGLE_1_ANNUEL")
    tri_a2.to_excel(writer, "TRIANGLE_2_ANNUEL")
    tri_a3.to_excel(writer, "TRIANGLE_3_ANNUEL")
    tri_a4.to_excel(writer, "TRIANGLE_4_ANNUEL")
    fct_a1.to_excel(writer, "FACTEURS_1_ANNUEL")
    fct_a2.to_excel(writer, "FACTEURS_2_ANNUEL")
    facteurs_a.to_excel(writer, "FACTEURS_ANNUELS")

    # CALCUL TRIMESTRIEL
    tri_t1.to_excel(writer, "TRIANGLE_1_TRIMESTRIEL")
    tri_t2.to_excel(writer, "TRIANGLE_2_TRIMESTRIEL")
    tri_t3.to_excel(writer, "TRIANGLE_3_TRIMESTRIEL")
    tri_t4.to_excel(writer, "TRIANGLE_4_TRIMESTRIEL")
    fct_t1.to_excel(writer, "FACTEURS_1_TRIMESTRIEL")
    fct_t2.to_excel(writer, "FACTEURS_2_TRIMESTRIEL")
    facteurs_t.to_excel(writer, "FACTEURS_TRIMESTRIELS")

    # CALCUL MENSUEL
    tri_m1.to_excel(writer, "TRIANGLE_1_MENSUEL")
    tri_m2.to_excel(writer, "TRIANGLE_2_MENSUEL")
    tri_m3.to_excel(writer, "TRIANGLE_3_MENSUEL")
    tri_m4.to_excel(writer, "TRIANGLE_4_MENSUEL")
    fct_m1.to_excel(writer, "FACTEURS_1_MENSUEL")
    fct_m2.to_excel(writer, "FACTEURS_2_MENSUEL")
    facteurs_m.to_excel(writer, "FACTEURS_MENSUELS")

    writer.save()

    print("TRAITEMENT DE LA BRANCHE {} EFFECTUÉE AVEC SUCCÈS...".format(nbranche))

    return filepath, filename



