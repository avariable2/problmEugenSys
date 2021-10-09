# pip install xlrd
import xlrd
import os
# pip install pydub
from pydub import AudioSegment, silence
import sys


class Found(Exception):
    pass


def match_target_amplitude(sound, target_dbfs):
    change_in_dbfs = target_dbfs - sound.dBFS
    return sound.apply_gain(change_in_dbfs)


def main():
    # Recuperation des noms des fichiers depuis le document donnée en parametre 1

    global list_audio_cut
    tab_titres_fichier = []

    fichier = xlrd.open_workbook("SD2_UnitAcknows_FR_AFN.xls")
    feuille_fichier_xls = fichier.sheet_by_index(0)  # Recupere la premiere feuille du fichier

    for i in range(feuille_fichier_xls.nrows):
        nom_catgerorie = feuille_fichier_xls.cell_value(i, 0)
        if nom_catgerorie == "InfAFN" and feuille_fichier_xls.cell_value(i, 8) != '':
            # Si le nom de la categorie est egale à celle donnée en 2eme paramètre et
            # que la cellule contenant le nom du fichier n'est pas vide
            tab_titres_fichier.append(feuille_fichier_xls.cell_value(i, 8))
    if len(tab_titres_fichier) == 0:
        print("La catégorie recherchée n'existe pas dans le fichier xls.")
        return None

    print("La catégorie à été trouvée. Début du traitement.")

    # Decoupage du fichier en plusieurs fichier
    song = AudioSegment.from_mp3("SD2_UnitAcknows_FR_AFN.mp3")
    song = match_target_amplitude(song, -20)

    # Le dB FS, sigle anglais pour Decibels relative to Full Scale, est une unité de niveau de signal audio. Elle
    # indique le rapport entre le niveau de ce signal et le niveau maximal.sur un multimedia
    dbfs = song.dBFS
    print(dbfs)

    min_decoupage = 100000  # Valeur arbitrairement grande pour essayer d'etre au plus pres
    list_secours = []
    try:
        for pmin_silence in range(400, 200, -50):
            for psilence_tresh in range(-160, 0, 10):
                list_audio_cut = silence.split_on_silence(song, min_silence_len=pmin_silence,
                                                          silence_thresh=psilence_tresh)
                print("[...]")
                if len(list_audio_cut) == len(tab_titres_fichier):
                    # Si on trouve le meme nombre de coupures que le nombre de fichiers attendus
                    print("Section approximative trouvée. Merci de verifier le résultat.")
                    print("Début de la création des fichiers.")
                    raise Found
                elif len(list_audio_cut) > len(tab_titres_fichier):
                    # Si on trouve plus de coupure que de nombre de fichier attendu on peut passer a l'iteration
                    # suivante
                    if min_decoupage > len(list_audio_cut):
                        min_decoupage = len(list_audio_cut)
                        list_secours = list_audio_cut
                    print(len(list_audio_cut))
                    break
        print("Aucun decoupage ne fonctionne en correspondance avec le nombre de découpages souhaités.")
        os.makedirs("decoupage_approximatif", exist_ok=True)
        for index, unAu in list_secours:
            unAu.export("decoupage_approximatif/Essai" + str(index).zfill(4) + ".ogg", format="ogg")
        print("Création d'un dossier alternatif contenant les fichiers. A vous de les verifier.")
    except Found:
        index = 0
        os.makedirs("acknows", exist_ok=True)
        for unAu in list_audio_cut:
            unAu.export("acknows/" + str(tab_titres_fichier[index]) + ".ogg", format="ogg")
            index += 1
        print("Fichiers créés.")

    return None


if __name__ == '__main__':
    main()
