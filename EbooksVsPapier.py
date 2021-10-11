# -*-coding:utf-8 -*

import os
import sys
import colorama
import requests
import re
from tkinter import *
from tkinter import filedialog

def run():
    accueil = '''       ----------------
      |Ebooks Vs Papier|
       ----------------
 Ce programme va identifier les ISBN d'ouvrages imprimés dans un fichier KBART valide
 (comme ceux qu'on trouve sur https://bacon.abes.fr/exporter.html).
 Il va ensuite chercher les notices correspondantes sur le Sudoc, et compter les 
 localisations dans le RCR indiqué par l'utilisateur.
 Cette information peut-être un outil d'aide à la décision dans les acquisitions de 
 bouquets de ressources électroniques.
'''
    print(accueil)
    os.system('pause')
    src = selection_kbart()
    if src:
        total_isbn, liste_lots_isbn = recup_isbn(src)
        if not total_isbn:
            print('Aucun ISBN papier identifié dans le fichier KBART.')
            os.system('pause')
        else:
            print('{} ISBN papier identifiés dans le fichier KBART.'.format(str(total_isbn)))
            rcr = choix_rcr()
            print('RCR sélectionné : {}.'.format(rcr))
            os.system('pause')
            clear_screen()
            avance_isbn, avance_ppn, avance_presences = 0, 0, 0
            affichage(total_isbn, avance_isbn, avance_ppn, avance_presences, rcr)
            for lot_isbn in liste_lots_isbn:
                nb_isbn, nb_ppn, nb_presences = interrogation_sudoc(rcr, lot_isbn)
                avance_isbn += nb_isbn
                avance_ppn += nb_ppn
                avance_presences += nb_presences
                affichage(total_isbn, avance_isbn, avance_ppn, avance_presences, rcr)
            print('Taux de notices papier localisées au RCR {} : {}%'.format(rcr, str(round(avance_presences/total_isbn*100, 2))))
            rapport = 'Fichier KBART analysé : {}.\n'.format(src)
            rapport += 'RCR de référence : {}.\n'.format(rcr)
            rapport += '{} ISBN papier identifiés dans le fichier.\n'.format(str(total_isbn))
            rapport += '{} notices Sudoc correspondantes.\n'.format(str(avance_ppn))
            rapport += "Attention, le nombre de notices peut dépasser le nombre d'ISBN (notices en doublon et ISBN erronés).\n"
            rapport += 'Si ce nombre de notices surestimé peut générer des localisations erronées, cela devrait rester marginal.\n'
            rapport += '{} localisations sur ces notices.\n'.format(str(avance_presences))
            rapport += 'Taux de notices papier localisées au RCR {} : {}%'.format(rcr, str(round(avance_presences/total_isbn*100, 2)))
            export_rapport(rapport)
            os.system('pause')

def selection_kbart():
    print('Sélectionner le fichier KBART à analyser')
    root = Tk()
    src = filedialog.askopenfilename(initialdir = os.path.expanduser("~/Desktop"),parent=root, title='Sélectionner le fichier KBART à analyser')
    root.destroy()
    return src

def recup_isbn(src):
    with open(src, 'r', encoding='utf-8', errors='replace') as fichier:
        kbart = fichier.readlines()
    liste_isbn = []
    for ligne in kbart:
        col = ligne.split('\t')
        if len(col) > 1 and len(col[1].replace('-', '')) == 10 and col[1].replace('-', '')[:9].isdecimal():
            isbn_10 = col[1].replace('-', '')
            isbn = '978' + isbn_10[:9] + str((10 - (38 + 3*int(isbn_10[0]) + int(isbn_10[1]) + 3*int(isbn_10[2]) + int(isbn_10[3]) + 3*int(isbn_10[4]) + int(isbn_10[5]) + 3*int(isbn_10[6]) + int(isbn_10[7]) + 3*int(isbn_10[8])) % 10) % 10)
            liste_isbn.append(isbn)
        elif len(col) > 1 and len(col[1].replace('-', '')) == 13 and col[1].replace('-', '').isdecimal():
            isbn = col[1].replace('-', '')
            liste_isbn.append(isbn)
    total_isbn = len(liste_isbn)
    liste_lots_isbn = [liste_isbn[x:x+100] for x in range(0, total_isbn, 100)]
    return total_isbn, liste_lots_isbn

def choix_rcr():
    rcr = '173002101'
    i = input('N° de RCR (par défaut 173002101, La Rochelle) : ')
    if i:
        rcr = i
    return rcr

def clear_screen():
    colorama.init()
    sys.stdout.write("\033[2J")

def interrogation_sudoc(rcr, lot_isbn):
    headers = {'User-Agent': 'Mozilla/5.0'}
    session = requests.Session()
    url_isbn2ppn = 'https://www.sudoc.fr/services/isbn2ppn/' + ','.join(lot_isbn)
    try:
        r = session.get(url_isbn2ppn, headers=headers)
        contenu_isbn2ppn = r.text
    except:
        contenu_isbn2ppn = ''
    liste_ppn_balises = re.findall('<ppn>.*?</ppn>', contenu_isbn2ppn, re.I)
    liste_ppn = []
    for e in liste_ppn_balises:
        ppn = re.sub('</?ppn>', '', e, re.I)
        liste_ppn.append(ppn)
    url_multiwhere = 'https://www.sudoc.fr/services/multiwhere/' + ','.join(liste_ppn)
    try:
        r = session.get(url_multiwhere, headers=headers)
        contenu_multiwhere = r.text
    except:
        contenu_multiwhere = ''
    liste_presences = re.findall(rcr, contenu_multiwhere, re.I)
    nb_isbn = len(lot_isbn)
    nb_ppn = len(liste_ppn)
    nb_presences = len(liste_presences)
    return nb_isbn, nb_ppn, nb_presences

def affichage(total_isbn, avance_isbn, avance_ppn, avance_presences, rcr):
    colorama.init()
    pb_width = 50
    i = int(avance_isbn/total_isbn*50)
    barre = '[' + "\u25A0" * i + " " * (pb_width-i) + ']'
    print(barre)
    chiffres_isbn = "ISBN analysés : {}/{} ({}%)".format(str(avance_isbn), str(total_isbn), str(int(avance_isbn/total_isbn*100)))
    print(chiffres_isbn)
    chiffres_sudoc = "{} notices Sudoc correspondantes / {} localisations au RCR {}".format(str(avance_ppn), str(avance_presences), rcr)
    print(chiffres_sudoc)
    sys.stdout.write("\033[1;1H")
    if avance_isbn == total_isbn:
        sys.stdout.write("\n\n\n")

def export_rapport(rapport):
    root = Tk()
    cible = filedialog.asksaveasfilename(initialdir=os.path.expanduser("~/Desktop"), parent=root, defaultextension='.txt', initialfile='rapport_EbooksVsPapier', title='Enregistrer un rapport')
    root.destroy()
    if cible:
        with open(cible, 'w', encoding='utf-8') as fichier:
            fichier.write(rapport)

if __name__ == "__main__":
    try:
        run()
    except Exception as erreur:
        print('Le programme a rencontré une erreur :')
        print(erreur)
        os.system('pause')