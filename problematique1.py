import csv
# pip install xlsxwriter
import xlsxwriter

# Lecture du fichier
fichier = open(r"pays.csv")
lecture = csv.reader(fichier)  # Lecture du fichier en entier
listeDesPays = []

for ligne in lecture:
    print(ligne)  # List de str
    nomDuPays = ligne[len(ligne) - 1]
    # compteurCaractere = 0
    # for lettre in nomDuPays:
    #    if lettre != ' ':
    #        compteurCaractere += 1
    # if compteurCaractere >= 15:
    #    listeDesPays.append(nomDuPays)
    if len(nomDuPays) >= 15:
        print("Ce pays est ajouté")
        listeDesPays.append(nomDuPays)
    else:
        print("Celui la non")

fichier.close()

print(listeDesPays)

# Creation du fichier xls
fichierXLS = xlsxwriter.Workbook('pays15.xls')  # Création ou réecriture d'un fichier xls
feuilleDuFichierXLS = fichierXLS.add_worksheet()  # Ajout d'une feuille dans ce fichier

for index, pays in enumerate(listeDesPays):
    feuilleDuFichierXLS.write("A" + str(index), pays)

fichierXLS.close()
