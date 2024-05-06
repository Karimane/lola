import pandas as pd

# Lecture
mb5b = pd.read_excel("mb5b.xlsx")
mb51 = pd.read_excel("mb51.xlsx")
mb51["Code mouvement"] = mb51["Code mouvement"].astype(str)

print (mb5b.columns)
print (mb51.columns)
print(mb51.head())

#Extraction
farin = mb5b[mb5b["Désignation article"].str.contains('farine', case=False)]
resultat = ("Farine.xlsx")
farin.to_excel(resultat, index=False)
print("Farine file has been created", resultat)

#Calcul Somme du stock initial mb5b
farin_df = pd.read_excel("Farine.xlsx")
print(farin_df.head())
print(farin_df.columns)
StockInitiale = farin_df["Stock initiale"].sum()
print("Le stock initial est :", StockInitiale) #2em col

#Traitement production mb51 
mb51.dropna(subset=["Désignation article"], inplace=True)
pattern = '|'.join(['101', '102'])# ajouter la recherche par code article
farinmb51 = mb51[mb51["Désignation article"].str.contains('farine', case=False)]
farinmb51 = mb51[mb51["Code mouvement"].str.contains(pattern, case=False)]
resultat2 = ("Farinemb51.xlsx")
farinmb51.to_excel(resultat2, index=False)
print("Farine file has been created", resultat2)

qte_df = pd.read_excel("Farinemb51.xlsx")
StockProd = qte_df["Quantité"].sum()
print("La quantité produite est :",StockProd) #1er col
resultat3= ("mb51FarineFinal.xlsx")
qte_df.to_excel(resultat3,index=False)
print("resultat finale created", resultat3)

#Traitement mb51 récup vente
mb51.dropna(subset=["Désignation article"], inplace=True)
pattern2 = '|'.join(['601', '602'])
ventemb51 = mb51[mb51["Désignation article"].str.contains('farine', case=False)]
ventemb51 = mb51[mb51["Code mouvement"].str.contains(pattern2, case=False)]
ventemb51["Quantité"] = ventemb51["Quantité"]
resultat4 = ("ventemb51.xlsx")
ventemb51.to_excel(resultat4, index=False)
print("vente file has been created", resultat4)

qtevente_df = pd.read_excel("ventemb51.xlsx")
Vente = qtevente_df["Quantité"].sum()
print("La quantité des ventes est :", Vente) # 4eme col
#Traitement  stock total 

Qtetotal = StockProd + StockInitiale
print("le stock totale est:",Qtetotal) #la 3em col
#Traitement  stock final
Qtefinal= Qtetotal - Vente
print("le stock final calculé est:",Qtefinal) # la 5 eme col
# récupération mb5b stock final
StockFinale = farin_df["Sold finale"].sum()
print("Le stock finale système est :", StockFinale)

Ma_Liste = []
Ma_Liste.append ({
 "Quantité Produite" : StockProd,
 "Stock initiale" : StockInitiale ,
 "Stock Totale" : Qtetotal,
 "Quantité des ventes" : Vente,
 "Stock final" : Qtefinal,

})
result_df = pd.DataFrame(Ma_Liste)
resultatFile = "final result.xlsx"
result_df.to_excel(resultatFile,index=False)
print("fichier final est crée ",result_df)









