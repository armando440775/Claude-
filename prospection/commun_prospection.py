# -*- coding: utf-8 -*-
"""
Briques communes aux fichiers de prospection apprentissage du bassin Nord-Est.

Utilise par :
    - generer_fichier_prospection.py   (entreprises privees de la logistique)
    - generer_fichier_public_sante.py  (collectivites, ecoles, sante)

Les deux classeurs partagent volontairement la MEME structure de colonnes :
le script completer_via_annuaire.py peut ainsi enrichir l'un comme l'autre.
"""
import re

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from urllib.parse import quote

# ---------------------------------------------------------------- Charte
BLEU   = "1F3864"
BLEU_C = "2E75B6"
GRIS   = "F2F2F2"
ORANGE = "ED7D31"
VERT   = "70AD47"
JAUNE  = "FFF2CC"
BLANC  = "FFFFFF"

F_TITRE   = Font(name="Calibri", size=16, bold=True, color=BLEU)
F_STITRE  = Font(name="Calibri", size=11, italic=True, color="595959")
F_ENTETE  = Font(name="Calibri", size=10, bold=True, color=BLANC)
F_CORPS   = Font(name="Calibri", size=10)
F_GRAS    = Font(name="Calibri", size=10, bold=True)
F_LIEN    = Font(name="Calibri", size=9, color="0563C1", underline="single")
F_H2      = Font(name="Calibri", size=12, bold=True, color=BLEU_C)

FILL_ENTETE = PatternFill("solid", fgColor=BLEU)
FILL_ALT    = PatternFill("solid", fgColor=GRIS)
FILL_JAUNE  = PatternFill("solid", fgColor=JAUNE)
FILL_ORANGE = PatternFill("solid", fgColor=ORANGE)
FILL_VERT   = PatternFill("solid", fgColor=VERT)

_th = Side(style="thin", color="BFBFBF")
BORD = Border(left=_th, right=_th, top=_th, bottom=_th)

AV = "A VERIFIER"
AC = "A COMPLETER"

# ---------------------------------------------------------------- Postes vises
P_OPLOG = "Opérateur logistique / Prépa. commandes"
P_MAGAS = "Magasinier / Gestion des stocks"
P_QUAI  = "Agent de quai / Réception-expédition"
P_TRANS = "Agent de transit / Déclarant douane jr"
P_ELOG  = "Employé logistique en magasin"


def terme_recherche(nom, commune):
    """Nettoie la raison sociale pour en faire un terme de recherche exploitable."""
    if nom.startswith("A prospecter") or nom.startswith("À prospecter"):
        return f"logistique transport {commune}"
    t = nom.split("(")[0]                      # retire l'acronyme entre parentheses
    t = re.split(r"\s[–—-]\s", t)[0]   # retire " - agence Reunion", " - site X"
    t = t.replace("’", " ").strip(" ,.")
    return f"{t} {commune}"


def lien_annuaire(nom, commune):
    return ("https://annuaire-entreprises.data.gouv.fr/rechercher?terme="
            + quote(terme_recherche(nom, commune)))


# ------------------------------------------------- Tables de reference INSEE
# Partagees par completer_via_annuaire.py et construire_fichier_sirene.py.
EFFECTIFS = {
    "NN": "Non renseigne (INSEE)", "00": "0 salarie", "01": "1 ou 2 salaries",
    "02": "3 a 5 salaries", "03": "6 a 9 salaries", "11": "10 a 19 salaries",
    "12": "20 a 49 salaries", "21": "50 a 99 salaries", "22": "100 a 199 salaries",
    "31": "200 a 249 salaries", "32": "250 a 499 salaries", "41": "500 a 999 salaries",
    "42": "1000 a 1999 salaries", "51": "2000 a 4999 salaries",
    "52": "5000 a 9999 salaries", "53": "10000 salaries et plus",
}

NAF = {
    "52.10A": "Entreposage et stockage frigorifique",
    "52.10B": "Entreposage et stockage non frigorifique",
    "52.24A": "Manutention portuaire",
    "52.24B": "Manutention non portuaire",
    "52.29A": "Messagerie, fret express",
    "52.29B": "Affretement et organisation des transports",
    "49.41A": "Transports routiers de fret interurbains",
    "49.41B": "Transports routiers de fret de proximite",
    "49.41C": "Location de camions avec chauffeur",
    "49.42Z": "Services de demenagement",
    "53.20Z": "Autres activites de poste et de courrier",
    "46.39A": "Commerce de gros de produits surgeles",
    "46.39B": "Commerce de gros alimentaire non specialise",
    "46.31Z": "Commerce de gros de fruits et legumes",
    "46.46Z": "Commerce de gros de produits pharmaceutiques",
    "46.73A": "Commerce de gros de bois et de materiaux de construction",
    "46.90Z": "Commerce de gros non specialise",
    "47.11D": "Supermarches",
    "47.11F": "Hypermarches",
    "47.52B": "Commerce de detail de materiaux de construction",
    "10.81Z": "Fabrication de sucre",
    # --- secteur public, scolaire et sanitaire
    "84.11Z": "Administration publique generale",
    "84.25C": "Services de secours et de lutte contre l'incendie",
    "85.20Z": "Enseignement primaire",
    "85.31Z": "Enseignement secondaire general",
    "85.32Z": "Enseignement secondaire technique ou professionnel",
    "85.42Z": "Enseignement superieur",
    "86.10Z": "Activites hospitalieres",
    "86.90E": "Autres activites para-medicales (dont dialyse)",
    "87.10A": "Hebergement medicalise pour personnes agees",
    "87.10B": "Hebergement medicalise pour enfants handicapes",
    "87.30A": "Hebergement social pour personnes agees",
    "87.30B": "Hebergement social pour handicapes physiques",
    "88.10A": "Aide a domicile",
    "88.10C": "Aide par le travail (ESAT)",
    "88.99B": "Action sociale sans hebergement n.c.a.",
    "56.29A": "Restauration collective sous contrat",
    "38.11Z": "Collecte des dechets non dangereux",
    "38.21Z": "Traitement et elimination des dechets non dangereux",
    "35.13Z": "Distribution d'electricite",
}

# ---------------------------------------------------------------- Colonnes
def colonnes(libelle_segment="Segment logistique"):
    """Colonnes du classeur. Seul l'intitule de la colonne C change d'un fichier
    a l'autre ; les index restent identiques pour completer_via_annuaire.py."""
    return [
        ("N°", 5),
        ("Raison sociale / Enseigne", 34),
        (libelle_segment, 26),
        ("Adresse", 38),
        ("CP", 7),
        ("Commune", 18),
        ("SIRET (14 chiffres)", 20),
        ("Code APE", 10),
        ("Libellé APE", 32),
        ("Effectif (tranche INSEE)", 22),
        ("Dirigeant (Nom Prénom)", 24),
        ("Fonction", 16),
        ("Téléphone", 16),
        ("Email contact", 30),
        ("Site web", 26),
        ("Poste apprenti visé", 30),
        ("Priorité", 9),
        ("Fiabilité donnée", 16),
        ("Vérifier SIRET / APE / Dirigeant", 34),
        ("Statut prospection", 20),
        ("Date dernier contact", 18),
        ("Notes", 40),
    ]


STATUTS = '"A contacter,Message laissé,RDV pris,Visite faite,Offre déposée,Candidat proposé,Contrat signé,Sans suite"'
PRIORITES = '"P1,P2,P3"'
FIABILITES = '"Vérifié INSEE,Vérifié,Source annuaire,A vérifier"'


def L(nom, seg, adr, cp, com, poste, prio, fiab, notes,
      siret=AV, ape=AV, lib_ape=AV, eff=AC, dg=AC, fonc="Dirigeant / DG",
      tel=AC, mail=AC, site=AC):
    """Une ligne du fichier. Tout champ non verifie reste a 'A VERIFIER'."""
    return dict(nom=nom, seg=seg, adr=adr, cp=cp, com=com, siret=siret, ape=ape,
                lib_ape=lib_ape, eff=eff, dg=dg, fonc=fonc, tel=tel, mail=mail,
                site=site, poste=poste, prio=prio, fiab=fiab, notes=notes)


# ---------------------------------------------------------------- Helpers
def entete_feuille(ws, titre, sous_titre, nb_cols):
    ws["A1"] = titre
    ws["A1"].font = F_TITRE
    ws["A2"] = sous_titre
    ws["A2"].font = F_STITRE
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nb_cols)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=nb_cols)
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 16


def feuille_commune(wb, nom_onglet, lignes, sous_titre,
                    prefixe_titre="PROSPECTION LOGISTIQUE", cols=None):
    cols = cols or colonnes()
    ws = wb.create_sheet(nom_onglet)
    entete_feuille(ws, f"{prefixe_titre} – {nom_onglet[3:].upper()}", sous_titre, len(cols))

    r = 4
    for i, (lib, larg) in enumerate(cols, start=1):
        c = ws.cell(row=r, column=i, value=lib)
        c.font, c.fill, c.border = F_ENTETE, FILL_ENTETE, BORD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = larg
    ws.row_dimensions[r].height = 34

    for j, e in enumerate(lignes, start=1):
        row = r + j
        vals = [j, e["nom"], e["seg"], e["adr"], e["cp"], e["com"], e["siret"], e["ape"],
                e["lib_ape"], e["eff"], e["dg"], e["fonc"], e["tel"], e["mail"],
                e["site"], e["poste"], e["prio"], e["fiab"],
                "Ouvrir la fiche entreprise", "A contacter", "", e["notes"]]
        for k, v in enumerate(vals, start=1):
            c = ws.cell(row=row, column=k, value=v)
            c.font, c.border = F_CORPS, BORD
            c.alignment = Alignment(vertical="top", wrap_text=(k in (2, 4, 9, 16, 22)))
            if j % 2 == 0:
                c.fill = FILL_ALT
        ws.cell(row=row, column=2).font = F_GRAS
        # SIRET et telephone en texte (ne jamais perdre les zeros de tete)
        ws.cell(row=row, column=7).number_format = "@"
        ws.cell(row=row, column=13).number_format = "@"
        ws.cell(row=row, column=21).number_format = "DD/MM/YYYY"
        # lien de verification
        lc = ws.cell(row=row, column=19)
        lc.hyperlink = lien_annuaire(e["nom"], e["com"])
        lc.font = F_LIEN
        # site web cliquable
        if str(e["site"]).startswith("http"):
            sc = ws.cell(row=row, column=15)
            sc.hyperlink = e["site"]
            sc.font = F_LIEN
        ws.row_dimensions[row].height = 46

    n = len(lignes)
    fin = r + n
    if n:
        dv_st = DataValidation(type="list", formula1=STATUTS, allow_blank=True)
        dv_pr = DataValidation(type="list", formula1=PRIORITES, allow_blank=True)
        dv_fi = DataValidation(type="list", formula1=FIABILITES, allow_blank=True)
        for dv, col in ((dv_st, "T"), (dv_pr, "Q"), (dv_fi, "R")):
            ws.add_data_validation(dv)
            dv.add(f"{col}{r+1}:{col}{r+400}")

    ws.auto_filter.ref = f"A{r}:{get_column_letter(len(cols))}{fin}"
    ws.freeze_panes = f"C{r+1}"
    ws.sheet_view.zoomScale = 90
    return ws, n


def onglet_pilotage(wb, compteurs, titre, sous_titre, libelle_total, objectifs,
                    libelle_volume="VOLUME PAR ONGLET", libelle_unite="Structures"):
    """Onglet de compteurs automatiques, alimente par la colonne T des onglets."""
    ws = wb.create_sheet("08 Pilotage")
    entete_feuille(ws, titre, sous_titre, 6)
    for i, w in enumerate([28, 14, 14, 14, 14, 14], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws["A4"] = libelle_volume
    ws["A4"].font = F_H2
    hdr = ["Onglet", libelle_unite, "À contacter", "RDV pris", "Contrats signés",
           "Taux de signature"]
    for i, h in enumerate(hdr, start=1):
        c = ws.cell(row=5, column=i, value=h)
        c.font, c.fill, c.border = F_ENTETE, FILL_ENTETE, BORD
        c.alignment = Alignment(horizontal="center", wrap_text=True)

    r = 6
    for nom_onglet, n in compteurs:
        q = f"'{nom_onglet}'!$T$5:$T$400"
        ws.cell(row=r, column=1, value=nom_onglet).font = F_CORPS
        ws.cell(row=r, column=2, value=n).font = F_CORPS
        ws.cell(row=r, column=3, value=f'=COUNTIF({q},"A contacter")')
        ws.cell(row=r, column=4, value=f'=COUNTIF({q},"RDV pris")')
        ws.cell(row=r, column=5, value=f'=COUNTIF({q},"Contrat signé")')
        ws.cell(row=r, column=6, value=f'=IF(B{r}=0,"",E{r}/B{r})')
        ws.cell(row=r, column=6).number_format = "0%"
        for k in range(1, 7):
            ws.cell(row=r, column=k).border = BORD
            ws.cell(row=r, column=k).font = F_CORPS
        r += 1

    tot = r
    ws.cell(row=tot, column=1, value=libelle_total).font = F_GRAS
    for k, col in ((2, "B"), (3, "C"), (4, "D"), (5, "E")):
        c = ws.cell(row=tot, column=k, value=f"=SUM({col}6:{col}{tot-1})")
        c.font = F_GRAS
    ws.cell(row=tot, column=6, value=f"=IF(B{tot}=0,\"\",E{tot}/B{tot})").font = F_GRAS
    ws.cell(row=tot, column=6).number_format = "0%"
    for k in range(1, 7):
        ws.cell(row=tot, column=k).border = BORD
        ws.cell(row=tot, column=k).fill = FILL_JAUNE

    r = tot + 3
    ws.cell(row=r, column=1, value="OBJECTIFS DE LA CAMPAGNE").font = F_H2
    r += 1
    for lib, val in objectifs:
        ws.cell(row=r, column=1, value=lib).font = F_CORPS
        ws.cell(row=r, column=2, value=val).font = F_GRAS
        r += 1
    ws.sheet_view.showGridLines = False
    return ws
