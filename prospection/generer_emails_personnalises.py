# -*- coding: utf-8 -*-
"""
Genere des brouillons d'email personnalises a partir d'un fichier de prospection
(Prospection_Logistique_Nord-Est_Reunion.xlsx ou Prospection_Public_Sante_Nord-Est_Reunion.xlsx)
et d'un modele defini dans modeles_email.py.

REGLE D'OR : ce script ne fait qu'ecrire des brouillons (.eml + .csv). Il n'envoie
jamais d'email lui-meme, ne se connecte a aucune boite mail. Chaque brouillon est
relu et envoye a la main.

Usage :
    python3 generer_emails_personnalises.py --lister
    python3 generer_emails_personnalises.py --categorie entreprise_prive --modele PREMIER_CONTACT
    python3 generer_emails_personnalises.py Prospection_Public_Sante_Nord-Est_Reunion.xlsx \\
        --categorie secteur_public --modele PREMIER_CONTACT_PUBLIC --statut "A contacter"

Options :
    --categorie   entreprise_prive (defaut) | secteur_public
    --modele      cle du modele (voir --lister ou l'onglet 01/02 de Modeles_Email_Prospection.xlsx)
    --statut      ne garder que les lignes ayant ce statut de prospection (repetable,
                  ou separe par des virgules) ; par defaut, toutes les lignes avec un email connu
    --sortie      dossier de sortie (defaut : emails_generes/ a cote du fichier source)
"""
import csv
import os
import re
import sys
from email.message import EmailMessage

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("openpyxl n'est pas installe.\n"
             "Lancez d'abord :  python3 -m pip install openpyxl")

from modeles_email import CONFIG, TEMPLATES, PLACEHOLDERS, personnaliser, signature_pour

# Colonnes (voir commun_prospection.colonnes() / generer_fichier_prospection.py)
C_NOM, C_SEG, C_ADR, C_CP, C_COM = 2, 3, 4, 5, 6
C_DG, C_FONC, C_TEL, C_MAIL, C_SITE = 11, 12, 13, 14, 15
C_POSTE, C_STATUT, C_NOTES = 16, 20, 22
LIGNE_ENTETE = 4

RE_EMAIL = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

FICHIER_PAR_DEFAUT = {
    "entreprise_prive": "Prospection_Logistique_Nord-Est_Reunion.xlsx",
    "secteur_public": "Prospection_Public_Sante_Nord-Est_Reunion.xlsx",
}


def est_vide(valeur):
    return valeur is None or str(valeur).strip() in PLACEHOLDERS


def email_valide(valeur):
    return bool(valeur) and not est_vide(valeur) and bool(RE_EMAIL.match(str(valeur).strip()))


def sluggify(texte, longueur=40):
    t = str(texte).strip().lower()
    t = re.sub(r"[^a-z0-9]+", "_", t)
    return t.strip("_")[:longueur] or "sans_nom"


def lister_modeles():
    for categorie, modeles in TEMPLATES.items():
        print("\n[%s]" % categorie)
        for cle, m in modeles.items():
            print("  %-32s %s" % (cle, m["nom"]))
            print("      %s" % m["quand"])


def valeurs_pour_ligne(ws, r):
    """Construit le dictionnaire de personnalisation pour une ligne du fichier."""
    nom = ws.cell(row=r, column=C_NOM).value
    segment = ws.cell(row=r, column=C_SEG).value
    commune = ws.cell(row=r, column=C_COM).value
    dirigeant = ws.cell(row=r, column=C_DG).value
    fonction = ws.cell(row=r, column=C_FONC).value
    poste = ws.cell(row=r, column=C_POSTE).value

    commune_ok = not est_vide(commune)
    dirigeant_ok = not est_vide(dirigeant)
    poste_ok = not est_vide(poste)
    fonction_ok = not est_vide(fonction)

    return dict(
        entreprise=nom or "",
        secteur="" if est_vide(segment) else segment,
        commune="" if not commune_ok else commune,
        commune_mention=("à %s" % commune) if commune_ok else "",
        fonction=fonction if fonction_ok else "responsable recrutement",
        poste=poste if poste_ok else "un(e) apprenti(e)",
        salutation=("Bonjour %s," % dirigeant) if dirigeant_ok else "Bonjour,",
    )


def creer_eml(chemin, destinataire, sujet, corps):
    msg = EmailMessage()
    msg["To"] = destinataire
    msg["From"] = "%s <%s>" % (CONFIG["mon_prenom_nom"], CONFIG["mon_email"])
    msg["Subject"] = sujet
    msg.set_content(corps)
    with open(chemin, "wb") as f:
        f.write(bytes(msg))


def traiter(chemin_src, categorie, cle_modele, statuts_filtre, dossier_sortie):
    modele = TEMPLATES[categorie][cle_modele]
    signature = signature_pour(categorie)

    wb = load_workbook(chemin_src, data_only=True)
    onglets = [n for n in wb.sheetnames if re.match(r"^0[1-7] ", n)]
    if not onglets:
        sys.exit("Aucun onglet de données trouvé (attendu : « 01 ... » à « 07 ... »).")

    os.makedirs(dossier_sortie, exist_ok=True)
    generes, journal = [], []

    for nom_onglet in onglets:
        ws = wb[nom_onglet]
        print("\n=== %s ===" % nom_onglet)
        for r in range(LIGNE_ENTETE + 1, ws.max_row + 1):
            nom = ws.cell(row=r, column=C_NOM).value
            if est_vide(nom) or str(nom).strip().startswith(("À prospecter", "A prospecter")):
                continue

            email = ws.cell(row=r, column=C_MAIL).value
            statut = ws.cell(row=r, column=C_STATUT).value or ""

            if not email_valide(email):
                journal.append([nom_onglet, r, nom, statut, "IGNOREE", "pas d'email connu"])
                continue
            if statuts_filtre and str(statut).strip() not in statuts_filtre:
                journal.append([nom_onglet, r, nom, statut, "IGNOREE",
                                 "statut hors filtre (%s)" % statut])
                continue

            valeurs = valeurs_pour_ligne(ws, r)
            valeurs["signature"] = signature
            sujet = personnaliser(modele["sujet"], valeurs)
            corps = personnaliser(modele["corps"], valeurs)

            nom_fichier = "%03d_%s.eml" % (len(generes) + 1, sluggify(nom))
            chemin_eml = os.path.join(dossier_sortie, nom_fichier)
            creer_eml(chemin_eml, str(email).strip(), sujet, corps)

            generes.append([nom_onglet, r, nom, email, sujet, corps])
            journal.append([nom_onglet, r, nom, statut, "GENEREE", nom_fichier])
            print("  ✓ %-46s -> %s" % (str(nom)[:46], nom_fichier))

    chemin_csv = os.path.join(dossier_sortie, "fusion_emails.csv")
    with open(chemin_csv, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Onglet", "Ligne", "Entreprise", "Email", "Sujet", "Corps"])
        w.writerows(generes)

    chemin_journal = os.path.join(dossier_sortie, "journal_emails.csv")
    with open(chemin_journal, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Onglet", "Ligne", "Entreprise", "Statut", "Résultat", "Détail"])
        w.writerows(journal)

    print("\n" + "=" * 62)
    print("Modèle utilisé     : [%s] %s" % (categorie, modele["nom"]))
    print("Emails générés     : %d" % len(generes))
    print("Lignes ignorées    : %d  (voir journal_emails.csv)" % (len(journal) - len(generes)))
    print("Dossier de sortie  : %s" % dossier_sortie)
    print("  -> ouvrez un .eml d'un double-clic : le brouillon s'affiche dans votre "
          "messagerie, prêt à relire et envoyer.")
    print("  -> ou importez fusion_emails.csv dans votre outil de publipostage habituel.")
    print("=" * 62)


def main():
    args = sys.argv[1:]

    def pop_option(nom, defaut=None):
        if nom in args:
            i = args.index(nom)
            valeur = args[i + 1] if i + 1 < len(args) else None
            del args[i:i + 2]
            return valeur
        return defaut

    if "--lister" in args:
        lister_modeles()
        return

    categorie = pop_option("--categorie", "entreprise_prive")
    cle_modele = pop_option("--modele")
    statut_brut = pop_option("--statut")
    sortie = pop_option("--sortie")
    fichiers = [a for a in args if not a.startswith("--")]

    if categorie not in TEMPLATES:
        sys.exit("Catégorie inconnue : %s (attendu : %s)" % (categorie, ", ".join(TEMPLATES)))
    if not cle_modele or cle_modele not in TEMPLATES[categorie]:
        print("Indiquez un modèle valide avec --modele. Modèles disponibles pour [%s] :\n"
              % categorie)
        for cle, m in TEMPLATES[categorie].items():
            print("  %-32s %s" % (cle, m["nom"]))
        sys.exit(1)

    ici = os.path.dirname(os.path.abspath(__file__))
    src = fichiers[0] if fichiers else os.path.join(ici, FICHIER_PAR_DEFAUT[categorie])
    if not os.path.exists(src):
        sys.exit("Fichier introuvable : %s" % src)

    statuts_filtre = None
    if statut_brut:
        statuts_filtre = {s.strip() for s in statut_brut.split(",") if s.strip()}

    dossier_sortie = sortie or os.path.join(
        os.path.dirname(os.path.abspath(src)), "emails_generes", categorie, cle_modele)

    print("Source  : %s" % src)
    print("Modèle  : [%s] %s" % (categorie, cle_modele))
    if statuts_filtre:
        print("Filtre statut : %s" % ", ".join(sorted(statuts_filtre)))
    print("Sortie  : %s" % dossier_sortie)

    traiter(src, categorie, cle_modele, statuts_filtre, dossier_sortie)


if __name__ == "__main__":
    main()
