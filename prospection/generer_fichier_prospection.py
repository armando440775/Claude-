# -*- coding: utf-8 -*-
"""
Genere le fichier de prospection logistique - Bassin Nord-Est de La Reunion
(Saint-Denis -> Sainte-Rose, hauteurs incluses).

Cible : entreprises susceptibles de recruter un/une apprenti(e) sur les metiers
de la logistique (operateur logistique, preparateur de commandes, agent de quai,
magasinier, gestionnaire de stocks, agent de transit).

IMPORTANT - INTEGRITE DES DONNEES
Les raisons sociales et adresses proviennent de sources publiques (annuaires,
sites officiels des entreprises). Les champs SIRET / Code APE / Dirigeant /
Telephone / Email ne sont renseignes QUE lorsqu'ils ont ete verifies.
Sinon : "A VERIFIER" + lien direct vers l'Annuaire des Entreprises
(annuaire-entreprises.data.gouv.fr) pour completer en 2 clics.
Aucune donnee n'a ete inventee.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import re
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

def terme_recherche(nom, commune):
    """Nettoie la raison sociale pour en faire un terme de recherche exploitable."""
    if nom.startswith("A prospecter") or nom.startswith("\u00c0 prospecter"):
        return f"logistique transport {commune}"
    t = nom.split("(")[0]                      # retire l'acronyme entre parentheses
    t = re.split(r"\s[\u2013\u2014-]\s", t)[0]   # retire " - agence Reunion", " - site X"
    t = t.replace("\u2019", " ").strip(" ,.")
    return f"{t} {commune}"

def lien_annuaire(nom, commune):
    return ("https://annuaire-entreprises.data.gouv.fr/rechercher?terme="
            + quote(terme_recherche(nom, commune)))

# ---------------------------------------------------------------- Colonnes
COLS = [
    ("N°", 5),
    ("Raison sociale / Enseigne", 34),
    ("Segment logistique", 26),
    ("Adresse", 38),
    ("CP", 7),
    ("Commune", 18),
    ("SIRET (14 chiffres)", 20),
    ("Code APE", 10),
    ("Libellé APE", 32),
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
FIABILITES = '"Vérifié,Source annuaire,A vérifier"'

# ---------------------------------------------------------------- Donnees
# Chaque tuple : (raison sociale, segment, adresse, CP, commune, siret, ape,
#                 libelle_ape, dirigeant, fonction, tel, email, site,
#                 poste_apprenti, priorite, fiabilite, notes)
def L(nom, seg, adr, cp, com, poste, prio, fiab, notes,
      siret=AV, ape=AV, lib_ape=AV, dg=AC, fonc="Dirigeant / DG",
      tel=AC, mail=AC, site=AC):
    return dict(nom=nom, seg=seg, adr=adr, cp=cp, com=com, siret=siret, ape=ape,
                lib_ape=lib_ape, dg=dg, fonc=fonc, tel=tel, mail=mail, site=site,
                poste=poste, prio=prio, fiab=fiab, notes=notes)

P_OPLOG = "Opérateur logistique / Prépa. commandes"
P_MAGAS = "Magasinier / Gestion des stocks"
P_QUAI  = "Agent de quai / Réception-expédition"
P_TRANS = "Agent de transit / Déclarant douane jr"
P_ELOG  = "Employé logistique en magasin"

SAINT_DENIS = [
    L("SOLAM OUTRE-MER (SOLAM RÉUNION)", "Prestataire logistique (3PL)",
      "3 chemin des Écoliers", "97400", "Saint-Denis", P_OPLOG, "P1", "Source annuaire",
      "Gestion d'entrepôt, livraison, suivi des flux, expédition. Profil idéal : préparateur de commandes."),
    L("SOCIÉTÉ RÉUNIONNAISE TOUT TRANSIT (SRTT)", "Transit / Commission de transport",
      "106 rue Deux Canons", "97400", "Saint-Denis", P_TRANS, "P1", "Source annuaire",
      "Transport international + stockage, préparation de commandes, expédition."),
    L("CMR OI", "Transit multimodal + stockage",
      AV, "97400", "Saint-Denis", P_TRANS, "P1", "Source annuaire",
      "Routier, aérien, maritime + stockage et distribution."),
    L("RUNLOGISTIC", "Commissionnaire de transport",
      AV, "97400", "Saint-Denis", P_TRANS, "P1", "Source annuaire",
      "Site : runlogistic.re — organisateur de transport.", site="https://www.runlogistic.re/"),
    L("MADARUN", "Transport / Logistique",
      "27 rue Monseigneur Mondon", "97400", "Saint-Denis", P_OPLOG, "P2", "Source annuaire",
      "Transport particuliers et professionnels."),
    L("DEPOTAGE LOGISTIQUE RÉUNION (DLR)", "Dépotage conteneurs / Entreposage",
      AV, "97400", "Saint-Denis", P_QUAI, "P1", "Source annuaire",
      "SIRET relevé sur base publique — à reconfirmer sur l'Annuaire des Entreprises.",
      siret="91889408000012"),
    L("SLOI (Groupe SOMATRANS)", "Commission en douane / Transit",
      AV, "97400", "Saint-Denis", P_TRANS, "P1", "Source annuaire",
      "Filiale groupe SOMATRANS — commissionnaire en douane."),
    L("SFT GONDRAND FRÈRES – agence Réunion", "Transit international / Douane",
      AV, "97400", "Saint-Denis", P_TRANS, "P2", "Source annuaire",
      "Groupe national, process RH structuré : viser le siège + le responsable d'agence."),
    L("KOB DISTRIBUTION", "Grossiste alimentaire + entrepôt",
      "10 rue Charles Gounod", "97400", "Saint-Denis", P_OPLOG, "P1", "Source annuaire",
      "Entrepôt de stockage, clientèle CHR : fort besoin de préparation de commandes."),
    L("974 EMBALLAGE", "Négoce emballage / Dépôt",
      "ZI du Chaudron, 28 rue Gabriel de Kerveguen", "97490", "Sainte-Clotilde (Saint-Denis)",
      P_MAGAS, "P2", "Source annuaire", "Emballages et fournitures pour la restauration."),
    L("ADELIS", "Distribution froid négatif / Livraison",
      AV, "97400", "Saint-Denis", P_OPLOG, "P2", "Source annuaire",
      "Glaces artisanales, livraison sur toute l'île : logistique du froid."),
    L("SODIREL", "Distribution / Logistique de marques",
      AV, "97400", "Saint-Denis", P_OPLOG, "P2", "Source annuaire",
      "Distributeur multimarques, logistique intégrée."),
    L("RAVATE PROFESSIONNEL", "Négoce matériaux + plateforme",
      AV, "97400", "Saint-Denis", P_MAGAS, "P1", "A vérifier",
      "Filiale du groupe Ravate (siège Saint-Denis). Gros volumes = magasiniers/caristes."),
]

SAINTE_MARIE = [
    L("SIFA LOGISTICS – agence Réunion", "Transit aérien & maritime / Entrepôt",
      "ZA de Gillot – Aéroport Roland Garros", "97438", "Sainte-Marie", P_TRANS, "P1",
      "Source annuaire", "Entrepôt sur zone aéroportuaire. Ouv. lun-jeu 8h-16h30, ven 8h-16h.",
      site="https://sifalogistics.com/nos-agences/sifa-reunion/"),
    L("DSG TRANSIT REPIQUET", "Transit aérien",
      "Parc d'activités aéroportuaire, 44 rue Hélène Boucher – îlot 1", "97438", "Sainte-Marie",
      P_TRANS, "P1", "Source annuaire", "Agence air."),
    L("SNT – SOCIÉTÉ NOUVELLE DE TRANSPORT", "Transit & Logistique",
      "Gare de fret aérien (2e étage) – Aéroport Roland Garros", "97438", "Sainte-Marie",
      P_TRANS, "P1", "Source annuaire", "Sur la plateforme fret."),
    L("T-TRAM (TRANSIT TRANSPORTS ROUTIERS AÉRIENS ET MARITIMES)", "Transit multimodal",
      "Aérogare Fret Roland Garros", "97438", "Sainte-Marie", P_TRANS, "P1", "Source annuaire",
      "Multimodal — bon terrain pour un apprenti agent de transit."),
    L("CHRONOPOST – agence Saint-Denis de La Réunion", "Messagerie express / Colis",
      "40 rue Hélène Boucher", "97438", "Sainte-Marie", P_QUAI, "P1", "Source annuaire",
      "Tél. relevé en annuaire public, à reconfirmer. Horaires : 8h-12h30 / 13h30-17h30, sam 8h30-12h. "
      "Tri et quai = poste d'agent de quai en alternance.",
      tel="0262 48 34 34"),
    L("DHL INTERNATIONAL RÉUNION", "Messagerie express international",
      "12 rue Hélène Boucher – ZA Saint-Exupéry, zone aéroportuaire", "97438", "Sainte-Marie",
      P_QUAI, "P1", "Source annuaire", "Groupe international : passer par le responsable d'agence ET le service RH DOM."),
    L("TRANS EXPRESS RÉUNION", "Transport express / Messagerie",
      AV, "97438", "Sainte-Marie", P_QUAI, "P2", "Source annuaire", ""),
    L("OCX RÉUNION", "Transport / Logistique",
      AV, "97438", "Sainte-Marie", P_OPLOG, "P2", "Source annuaire", ""),
    L("SD BEAUSÉJOUR (supermarché)", "Grande distribution – réserve",
      "200 avenue Beau Pays", "97438", "Sainte-Marie", P_ELOG, "P2", "Source annuaire",
      "Réception marchandises / réserve : poste d'employé logistique."),
    L("U EXPRESS SAINTE-MARIE", "Grande distribution – réserve",
      "112 rue Roger Payet", "97438", "Sainte-Marie", P_ELOG, "P2", "Source annuaire",
      "Enseigne U : accords alternance fréquents."),
]

SAINTE_SUZANNE = [
    L("SOCIÉTÉ TRANS LOGISTICS RÉUNION", "Transport routier de fret",
      "85 rue Marchande", "97441", "Sainte-Suzanne", P_OPLOG, "P1", "Source annuaire", ""),
    L("MRUN TRANSPORT", "Transport routier de fret",
      "1B chemin Foutaque", "97441", "Sainte-Suzanne", P_QUAI, "P2", "Source annuaire", ""),
    L("TIAMONTINIA – TRANSPORT DE MARCHANDISES", "Transport routier de fret",
      "7 Bagatelle", "97441", "Sainte-Suzanne", P_QUAI, "P3", "Source annuaire",
      "TPE : vérifier la capacité à accueillir un apprenti (tuteur disponible)."),
    L("ALLON ROULE !", "Transport / Livraison",
      "19 chemin des Azalées", "97441", "Sainte-Suzanne", P_OPLOG, "P3", "Source annuaire", ""),
]

SAINT_ANDRE = [
    L("TRANSPORT LOGISTIQUE RÉUNION (TLR)", "Transport & logistique",
      "221 impasse du Centre", "97440", "Saint-André", P_OPLOG, "P1", "Vérifié",
      "SIRET du siège social relevé sur l'Annuaire des Entreprises. Entreprise récente : "
      "argument « construire votre équipe avec l'alternance ».",
      siret="94497358500010"),
    L("GLR LOGISTICS", "Prestataire logistique (SAS)",
      AV, "97440", "Saint-André", P_OPLOG, "P1", "Source annuaire", ""),
    L("SAM I TRANSPORTE", "Transport routier / Livraison",
      "250 rue Bois Rouge", "97440", "Saint-André", P_QUAI, "P2", "Source annuaire",
      "Livraisons sur toute La Réunion."),
    L("ZARLOC", "Location / Logistique",
      "248 chemin Fourchon", "97440", "Saint-André", P_MAGAS, "P3", "Source annuaire",
      "Vérifier l'activité exacte (code APE) avant l'appel."),
    L("TERALTA GRANULATS & BÉTON RÉUNION – site Saint-André", "Industrie / Flux matériaux",
      AV, "97440", "Saint-André", P_MAGAS, "P2", "Source annuaire",
      "Groupe structuré : passer par la RH régionale. Logistique de granulats et béton."),
    L("RUN MARKET SAINT-ANDRÉ", "Grande distribution – réserve",
      "CC La Cocoteraie", "97440", "Saint-André", P_ELOG, "P2", "Source annuaire", ""),
    L("SUPER U SAINT-ANDRÉ", "Grande distribution – réserve",
      AV, "97440", "Saint-André", P_ELOG, "P2", "Source annuaire",
      "Réception / réserve / drive : plusieurs postes possibles."),
]

BRAS_PANON = [
    L("CRÉOLE LOGISTIQUE TRANSPORT", "Transport de fret interurbain",
      AV, "97412", "Bras-Panon", P_OPLOG, "P1", "Vérifié",
      "SIRET relevé sur base publique (siège Bras-Panon), à reconfirmer. "
      "Classée INSEE en transport de fret interurbain.",
      siret="80216938300014"),
    L("AFA MARIE", "Transport routier",
      AV, "97412", "Bras-Panon", P_QUAI, "P3", "Source annuaire", ""),
]

SAINT_BENOIT = [
    L("TRANSPORTS CAMALON (GROUPE CAMALON)", "Transport de fret / Matières dangereuses",
      AV, "97470", "Saint-Benoît", P_OPLOG, "P1", "Source annuaire",
      "Créée en 1990. Denrées alimentaires, agrégats, ADR. Site : groupe-camalon.com — "
      "structure solide, très bon potentiel tuteur.",
      site="https://www.groupe-camalon.com/"),
    L("TRANS'COOP O.I", "Transport routier de marchandises",
      AV, "97470", "Saint-Benoît", P_QUAI, "P1", "Source annuaire",
      "Site : transcoopoi.re", site="https://www.transcoopoi.re/"),
    L("GROUPE HOSPITALIER EST RÉUNION (GHER)", "Logistique hospitalière / Magasin",
      "30 RN3 – ZAC Madeleine, Bras-Fusil", "97470", "Saint-Benoît", P_MAGAS, "P1",
      "Source annuaire",
      "Employeur public : apprentissage possible dans la fonction publique hospitalière. "
      "Magasin général, pharmacie, flux de médicaments et équipements."),
    L("CARREFOUR SAINT-BENOÎT", "Grande distribution – réserve",
      "6 chemin Goyaves", "97470", "Saint-Benoît", P_ELOG, "P1", "Source annuaire",
      "Hyper : réception, réserve, drive. Passer par le directeur de magasin + RH régionale."),
    L("SUPER U CHANE FAT", "Grande distribution – réserve",
      "1 rue Louis Brunet", "97470", "Saint-Benoît", P_ELOG, "P2", "Source annuaire",
      "Franchise indépendante : décision rapide, parler directement au dirigeant."),
]

EST_PROFOND = [
    L("À prospecter – ZA de Sainte-Rose", "Terrain à défricher",
      "Zone d'activités / commerces de la commune", "97439", "Sainte-Rose", P_ELOG, "P3",
      "A vérifier",
      "Aucune entreprise logistique pure identifiée en source ouverte. Cibler : supérettes, "
      "négoces de matériaux, transporteurs indépendants. Méthode : mairie + agence Pôle emploi "
      "/ France Travail de Saint-Benoît."),
    L("À prospecter – La Plaine-des-Palmistes", "Terrain à défricher",
      AV, "97431", "La Plaine-des-Palmistes", P_ELOG, "P3", "A vérifier",
      "Commerces de proximité, artisans, transport de matériaux. Faible densité logistique."),
    L("À prospecter – Salazie (hauteurs)", "Terrain à défricher",
      AV, "97433", "Salazie", P_ELOG, "P3", "A vérifier",
      "Hauteurs : agriculture (chouchou), transport de denrées, coopératives. "
      "Contacter la Chambre d'agriculture et les coopératives locales."),
    L("À prospecter – Saint-Denis hauteurs (La Montagne, Le Brûlé, Bois-de-Nèfles)",
      "Terrain à défricher", AV, "97400", "Saint-Denis (hauteurs)", P_ELOG, "P3", "A vérifier",
      "Peu de logistique en hauteur : viser les dépôts et commerces de proximité, "
      "et les entreprises du bas où les jeunes des hauteurs peuvent se rendre."),
]

ONGLETS = [
    ("01 Saint-Denis", SAINT_DENIS, "Saint-Denis / Sainte-Clotilde / Le Chaudron"),
    ("02 Sainte-Marie", SAINTE_MARIE, "Sainte-Marie – zone aéroportuaire Roland Garros, Duparc, Bel-Air"),
    ("03 Sainte-Suzanne", SAINTE_SUZANNE, "Sainte-Suzanne – Quartier Français, Bagatelle"),
    ("04 Saint-André", SAINT_ANDRE, "Saint-André – Cambuston, Bois-Rouge, La Cocoteraie"),
    ("05 Bras-Panon", BRAS_PANON, "Bras-Panon"),
    ("06 Saint-Benoît", SAINT_BENOIT, "Saint-Benoît – ZAC Bras-Fusil, Beaulieu"),
    ("07 Sainte-Rose & Hauteurs", EST_PROFOND, "Sainte-Rose, La Plaine-des-Palmistes, Salazie, hauteurs de Saint-Denis"),
]

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

def feuille_commune(wb, nom_onglet, lignes, sous_titre):
    ws = wb.create_sheet(nom_onglet)
    entete_feuille(ws, f"PROSPECTION LOGISTIQUE – {nom_onglet[3:].upper()}", sous_titre, len(COLS))

    r = 4
    for i, (lib, larg) in enumerate(COLS, start=1):
        c = ws.cell(row=r, column=i, value=lib)
        c.font, c.fill, c.border = F_ENTETE, FILL_ENTETE, BORD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = larg
    ws.row_dimensions[r].height = 34

    for j, e in enumerate(lignes, start=1):
        row = r + j
        vals = [j, e["nom"], e["seg"], e["adr"], e["cp"], e["com"], e["siret"], e["ape"],
                e["lib_ape"], e["dg"], e["fonc"], e["tel"], e["mail"], e["site"],
                e["poste"], e["prio"], e["fiab"],
                "Ouvrir la fiche entreprise", "A contacter", "", e["notes"]]
        for k, v in enumerate(vals, start=1):
            c = ws.cell(row=row, column=k, value=v)
            c.font, c.border = F_CORPS, BORD
            c.alignment = Alignment(vertical="top", wrap_text=(k in (2, 4, 9, 15, 21)))
            if j % 2 == 0:
                c.fill = FILL_ALT
        ws.cell(row=row, column=2).font = F_GRAS
        # SIRET en texte (ne jamais perdre les zeros de tete)
        ws.cell(row=row, column=7).number_format = "@"
        ws.cell(row=row, column=12).number_format = "@"
        ws.cell(row=row, column=20).number_format = "DD/MM/YYYY"
        # lien de verification
        lc = ws.cell(row=row, column=18)
        lc.hyperlink = lien_annuaire(e["nom"], e["com"])
        lc.font = F_LIEN
        # site web cliquable
        if str(e["site"]).startswith("http"):
            sc = ws.cell(row=row, column=14)
            sc.hyperlink = e["site"]
            sc.font = F_LIEN
        ws.row_dimensions[row].height = 46

    n = len(lignes)
    fin = r + n
    if n:
        dv_st = DataValidation(type="list", formula1=STATUTS, allow_blank=True)
        dv_pr = DataValidation(type="list", formula1=PRIORITES, allow_blank=True)
        dv_fi = DataValidation(type="list", formula1=FIABILITES, allow_blank=True)
        for dv, col in ((dv_st, "S"), (dv_pr, "P"), (dv_fi, "Q")):
            ws.add_data_validation(dv)
            dv.add(f"{col}{r+1}:{col}{r+400}")

    ws.auto_filter.ref = f"A{r}:{get_column_letter(len(COLS))}{fin}"
    ws.freeze_panes = f"C{r+1}"
    ws.sheet_view.zoomScale = 90
    return ws, n

# ---------------------------------------------------------------- Workbook
wb = Workbook()
wb.remove(wb.active)

# ---- Onglet 00 : mode d'emploi
ws = wb.create_sheet("00 Mode d'emploi")
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 40
ws.column_dimensions["C"].width = 100
ws["B2"] = "FICHIER DE PROSPECTION – LOGISTIQUE EN APPRENTISSAGE"
ws["B2"].font = Font(name="Calibri", size=18, bold=True, color=BLEU)
ws["B3"] = "Bassin Nord-Est de La Réunion : de Saint-Denis à Sainte-Rose, hauteurs comprises"
ws["B3"].font = F_STITRE

blocs = [
    ("À QUOI SERT CE FICHIER", ""),
    ("Objectif", "Identifier et suivre les entreprises du bassin Nord-Est susceptibles de signer un contrat "
                 "d'apprentissage sur un métier de la logistique (16-25 ans)."),
    ("Périmètre géographique", "Saint-Denis, Sainte-Clotilde, Sainte-Marie, Sainte-Suzanne, Saint-André, "
                 "Bras-Panon, Saint-Benoît, La Plaine-des-Palmistes, Sainte-Rose, Salazie et les hauteurs."),
    ("Un onglet = une commune", "Les onglets 01 à 07 suivent l'axe littoral Nord-Est, dans l'ordre de la route. "
                 "Vous pouvez faire une tournée terrain onglet par onglet."),
    ("", ""),
    ("⚠️ LIRE AVANT UTILISATION – FIABILITÉ DES DONNÉES", ""),
    ("Ce qui est fourni", "Raison sociale, adresse, commune, segment d'activité et poste d'apprenti pertinent. "
                 "Ces éléments proviennent de sources publiques (annuaires professionnels, sites officiels "
                 "des entreprises)."),
    ("Ce qui est à compléter", "SIRET, code APE, nom du dirigeant, téléphone et email portent la mention "
                 "« A VERIFIER » ou « A COMPLETER » lorsqu'ils n'ont pas pu être confirmés. "
                 "AUCUNE de ces données n'a été inventée : un SIRET ou un nom de dirigeant faux vous ferait "
                 "perdre votre crédibilité dès le premier appel."),
    ("Comment compléter en 2 clics", "Colonne « Vérifier SIRET / APE / Dirigeant » : chaque ligne contient un lien "
                 "pré-rempli vers l'Annuaire des Entreprises (annuaire-entreprises.data.gouv.fr). "
                 "Cliquez, puis recopiez le SIRET, le code APE et le nom du dirigeant. C'est gratuit, "
                 "officiel et à jour."),
    ("Emails et téléphones", "Ils ne figurent pas dans l'open data. Sources : site web de l'entreprise (page "
                 "Contact / Mentions légales), page LinkedIn de l'entreprise, ou standard téléphonique. "
                 "Astuce : demandez toujours « qui s'occupe des recrutements et de la taxe d'apprentissage ? »."),
    ("", ""),
    ("MÉTHODE DE TRAVAIL CONSEILLÉE", ""),
    ("1. Qualifier avant d'appeler", "Complétez SIRET + APE + dirigeant. Un appel préparé où vous citez le nom du "
                 "dirigeant et l'activité exacte double le taux de RDV."),
    ("2. Prioriser", "Colonne Priorité : P1 = logistique au cœur du métier, structure capable de tutorer. "
                 "P2 = besoin logistique réel mais secondaire. P3 = TPE ou à qualifier. Traitez les P1 en premier."),
    ("3. Tracer chaque contact", "Colonnes « Statut prospection » et « Date dernier contact » : à remplir après "
                 "CHAQUE échange. C'est ce qui fait la différence sur une campagne de 3 mois."),
    ("4. Suivre les chiffres", "Onglet « 08 Pilotage » : compteurs automatiques par commune et par statut."),
    ("", ""),
    ("ONGLETS OUTILS", ""),
    ("09 Codes APE à cibler", "La liste des codes APE/NAF qui trahissent un besoin logistique. Utilisez-la pour "
                 "élargir vous-même le fichier depuis l'Annuaire des Entreprises."),
    ("10 Script d'appel", "Trame d'appel à froid, argumentaire employeur et réponses aux objections les plus "
                 "fréquentes."),
]
r = 5
for titre, txt in blocs:
    if titre and not txt:
        ws.cell(row=r, column=2, value=titre).font = F_H2
        ws.cell(row=r, column=2).fill = FILL_JAUNE
        ws.cell(row=r, column=3).fill = FILL_JAUNE
        r += 1
        continue
    if not titre and not txt:
        r += 1
        continue
    a = ws.cell(row=r, column=2, value=titre); a.font = F_GRAS
    a.alignment = Alignment(vertical="top", wrap_text=True)
    b = ws.cell(row=r, column=3, value=txt); b.font = F_CORPS
    b.alignment = Alignment(vertical="top", wrap_text=True)
    ws.row_dimensions[r].height = max(15, 13 * (len(txt) // 95 + 1))
    r += 1
ws.sheet_view.showGridLines = False

# ---- Onglets communes
compteurs = []
for nom_onglet, lignes, st in ONGLETS:
    _, n = feuille_commune(wb, nom_onglet, lignes, st)
    compteurs.append((nom_onglet, n))

# ---- Onglet 08 : pilotage
ws = wb.create_sheet("08 Pilotage")
entete_feuille(ws, "PILOTAGE DE LA CAMPAGNE", "Compteurs automatiques – se mettent à jour quand vous remplissez les onglets communes", 6)
for i, w in enumerate([28, 14, 14, 14, 14, 14], start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

ws["A4"] = "VOLUME PAR COMMUNE"; ws["A4"].font = F_H2
hdr = ["Onglet", "Entreprises", "À contacter", "RDV pris", "Contrats signés", "Taux de signature"]
for i, h in enumerate(hdr, start=1):
    c = ws.cell(row=5, column=i, value=h)
    c.font, c.fill, c.border = F_ENTETE, FILL_ENTETE, BORD
    c.alignment = Alignment(horizontal="center", wrap_text=True)
r = 6
for nom_onglet, n in compteurs:
    q = f"'{nom_onglet}'!$S$5:$S$400"
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
ws.cell(row=tot, column=1, value="TOTAL BASSIN NORD-EST").font = F_GRAS
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
obj = [
    ("Appels à passer / semaine", "30"),
    ("RDV entreprise / semaine", "5"),
    ("Offres d'apprentissage à collecter", "à définir"),
    ("Ratio moyen observé", "≈ 10 appels → 2 RDV → 1 offre"),
]
for lib, val in obj:
    ws.cell(row=r, column=1, value=lib).font = F_CORPS
    ws.cell(row=r, column=2, value=val).font = F_GRAS
    r += 1
ws.sheet_view.showGridLines = False

# ---- Onglet 09 : codes APE
ws = wb.create_sheet("09 Codes APE à cibler")
entete_feuille(ws, "CODES APE / NAF À CIBLER", "Pour élargir vous-même le fichier depuis annuaire-entreprises.data.gouv.fr (filtre « activité »)", 4)
for i, w in enumerate([12, 62, 34, 14], start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
hdr = ["Code APE", "Libellé officiel", "Pourquoi c'est une cible apprentissage", "Priorité"]
for i, h in enumerate(hdr, start=1):
    c = ws.cell(row=4, column=i, value=h)
    c.font, c.fill, c.border = F_ENTETE, FILL_ENTETE, BORD
    c.alignment = Alignment(horizontal="center", wrap_text=True)
ape = [
    ("52.10A", "Entreposage et stockage frigorifique", "Cœur de cible : préparation de commandes, magasinage", "P1"),
    ("52.10B", "Entreposage et stockage non frigorifique", "Cœur de cible : entrepôts, plateformes", "P1"),
    ("52.24A", "Manutention portuaire", "Agent de quai, manutention (plutôt Le Port, à surveiller)", "P2"),
    ("52.24B", "Manutention non portuaire", "Agent de quai, chargement/déchargement", "P1"),
    ("52.29A", "Messagerie, fret express", "Tri, quai, expédition : très gros besoin en alternance", "P1"),
    ("52.29B", "Affrètement et organisation des transports", "Transit, exploitation transport", "P1"),
    ("49.41A", "Transports routiers de fret interurbains", "Exploitation, quai, gestion des tournées", "P1"),
    ("49.41B", "Transports routiers de fret de proximité", "Livraison, préparation de tournées", "P1"),
    ("49.41C", "Location de camions avec chauffeur", "Support logistique, planning", "P2"),
    ("49.42Z", "Services de déménagement", "Manutention, inventaire, gestion de stock", "P2"),
    ("53.20Z", "Autres activités de poste et de courrier", "Tri, distribution de colis", "P1"),
    ("46.xx", "Commerce de gros (négoce)", "Presque tous ont un entrepôt : magasinier, prépa. commandes", "P1"),
    ("47.11D/F", "Supermarchés / Hypermarchés", "Réception, réserve, drive : employé logistique", "P1"),
    ("47.52B", "Commerce de détail de matériaux de construction", "Dépôt, cour matériaux, magasinier-cariste", "P1"),
    ("10.xx / 11.xx", "Industries alimentaires et boissons", "Magasin matières premières, expédition produits finis", "P2"),
    ("86.10Z", "Activités hospitalières", "Magasin général, logistique de soins (apprentissage public)", "P2"),
    ("45.11Z / 45.31Z", "Commerce de véhicules / d'équipements automobiles", "Magasin pièces détachées : magasinier-vendeur", "P2"),
]
for j, (code, lib, why, p) in enumerate(ape, start=1):
    row = 4 + j
    for k, v in enumerate([code, lib, why, p], start=1):
        c = ws.cell(row=row, column=k, value=v)
        c.font, c.border = F_CORPS, BORD
        c.alignment = Alignment(vertical="top", wrap_text=True)
        if j % 2 == 0:
            c.fill = FILL_ALT
    ws.cell(row=row, column=1).font = F_GRAS
r = 4 + len(ape) + 2
ws.cell(row=r, column=1, value="MÉTHODE POUR ÉLARGIR LE FICHIER").font = F_H2
r += 1
for txt in [
    "1. Aller sur https://annuaire-entreprises.data.gouv.fr (gratuit, données officielles INSEE).",
    "2. Recherche avancée : saisir le code APE ci-dessus + département 974 + la commune visée.",
    "3. La fiche donne : raison sociale, SIRET, code APE, adresse, effectif, et le nom du dirigeant.",
    "4. Recopier dans l'onglet de la commune, puis chercher le téléphone/email sur le site de l'entreprise.",
    "5. Vérifier aussi que l'établissement est bien « en activité » avant de l'appeler.",
]:
    c = ws.cell(row=r, column=1, value=txt); c.font = F_CORPS
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    r += 1
ws.cell(row=r+1, column=1, value="Ouvrir l'Annuaire des Entreprises").font = F_LIEN
ws.cell(row=r+1, column=1).hyperlink = "https://annuaire-entreprises.data.gouv.fr/"
ws.sheet_view.showGridLines = False

# ---- Onglet 10 : script d'appel
ws = wb.create_sheet("10 Script d'appel")
entete_feuille(ws, "SCRIPT D'APPEL & ARGUMENTAIRE EMPLOYEUR", "Prospection apprentissage – métiers de la logistique", 2)
ws.column_dimensions["A"].width = 30
ws.column_dimensions["B"].width = 110
script = [
    ("H", "1. LE BARRAGE SECRÉTARIAT", ""),
    ("T", "Objectif", "Obtenir le nom + la ligne du décideur (gérant, responsable d'exploitation, RH)."),
    ("T", "Phrase", "« Bonjour, [Prénom Nom], je suis chargé des relations entreprises au CFA [nom]. "
                    "Je cherche la personne qui décide des recrutements en logistique, c'est bien "
                    "M./Mme [nom du dirigeant] ? »"),
    ("T", "Astuce", "Citer le nom du dirigeant (colonne J) fait sauter le barrage 8 fois sur 10. "
                    "D'où l'importance de qualifier AVANT d'appeler."),
    ("H", "2. L'ACCROCHE (30 secondes)", ""),
    ("T", "Phrase", "« Je ne vous appelle pas pour vous vendre quelque chose. Je forme des jeunes de 16 à 25 ans "
                    "aux métiers de la logistique sur le Nord-Est, et je place des apprentis chez des entreprises "
                    "comme la vôtre. Vous avez déjà accueilli un alternant ? »"),
    ("T", "Pourquoi ça marche", "On annonce qu'on n'est pas un commercial, on cite le territoire, "
                    "et on termine par une question ouverte."),
    ("H", "3. LES 5 ARGUMENTS QUI FONT SIGNER", ""),
    ("T", "① Le coût", "L'apprentissage est le dispositif le moins coûteux pour recruter. Les aides "
                    "à l'embauche et les exonérations s'appliquent (montants à vérifier chaque année "
                    "auprès de l'OPCO — ne jamais annoncer un chiffre non actualisé)."),
    ("T", "② La formation gratuite", "Le coût pédagogique est pris en charge par l'OPCO de l'entreprise. "
                    "Pour l'employeur, la formation ne sort pas de sa trésorerie."),
    ("T", "③ Le recrutement sans risque", "L'apprentissage, c'est une période d'essai longue et financée. "
                    "L'entreprise forme le salarié à SES process, SES logiciels, SES clients."),
    ("T", "④ La pénurie de main-d'œuvre", "Sur le Nord-Est, les caristes, préparateurs et agents de quai "
                    "sont difficiles à recruter. Former en interne, c'est sécuriser les 3 ans à venir."),
    ("T", "⑤ Le zéro-souci administratif", "« Le contrat, le Cerfa, l'OPCO, le suivi : c'est moi qui gère. "
                    "Vous, vous n'avez qu'à dire oui à un candidat. »"),
    ("H", "4. RÉPONSES AUX OBJECTIONS", ""),
    ("T", "« On n'a pas le temps de former »",
          "« Un apprenti est en entreprise 3 semaines sur 4. Il est opérationnel sur la préparation de commandes "
          "en quelques semaines. Et le tuteur, on le forme et on l'accompagne. »"),
    ("T", "« On n'a pas de budget »",
          "« Justement : c'est le contrat le moins cher du marché, et la formation est financée par votre OPCO. "
          "Je vous fais le calcul précis pour votre cas en 10 minutes. »"),
    ("T", "« Les jeunes ne sont pas fiables »",
          "« C'est pour ça que je fais un pré-recrutement : je vous présente 2 ou 3 candidats déjà testés "
          "sur la motivation et la ponctualité. Vous choisissez, ou vous ne choisissez personne. »"),
    ("T", "« On est trop petit »",
          "« Une entreprise de 3 personnes peut prendre un apprenti. Il faut juste un tuteur qui a le métier — "
          "et vous l'avez. »"),
    ("T", "« Rappelez-moi dans 6 mois »",
          "« Très bien. Je note [date précise]. En attendant je vous envoie une page qui résume les aides, "
          "et je vous préviens si j'ai un très bon profil qui habite à côté de chez vous. » "
          "→ noter la date dans la colonne T de l'onglet commune."),
    ("H", "5. LA CLÔTURE", ""),
    ("T", "Ne jamais raccrocher sans", "① un nom + un email direct, ② une date de relance précise, "
                    "③ si possible un RDV de 20 min sur site."),
    ("T", "Phrase de closing", "« Je passe dans votre secteur [jour]. Je vous montre 2 profils en 15 minutes, "
                    "vous me dites oui ou non. Ça vous va ? »"),
    ("H", "6. APRÈS L'APPEL – À FAIRE LE JOUR MÊME", ""),
    ("T", "Traçabilité", "Remplir « Statut prospection » + « Date dernier contact » + « Notes » dans l'onglet commune."),
    ("T", "Email de confirmation", "Envoyer sous 2h un mail court : ce qui a été dit, ce que vous envoyez, "
                    "la date de relance. Ça vous positionne comme un pro."),
]
r = 4
for typ, a, b in script:
    if typ == "H":
        c = ws.cell(row=r, column=1, value=a)
        c.font = F_H2; c.fill = FILL_JAUNE
        ws.cell(row=r, column=2).fill = FILL_JAUNE
        ws.row_dimensions[r].height = 20
    else:
        ca = ws.cell(row=r, column=1, value=a); ca.font = F_GRAS
        ca.alignment = Alignment(vertical="top", wrap_text=True)
        cb = ws.cell(row=r, column=2, value=b); cb.font = F_CORPS
        cb.alignment = Alignment(vertical="top", wrap_text=True)
        ws.row_dimensions[r].height = max(15, 14 * (len(b) // 105 + 1))
    r += 1
ws.sheet_view.showGridLines = False

# ---- Onglet 11 : sources
ws = wb.create_sheet("11 Sources & mise à jour")
entete_feuille(ws, "SOURCES DES DONNÉES", "Traçabilité et outils pour compléter / actualiser le fichier", 2)
ws.column_dimensions["A"].width = 46
ws.column_dimensions["B"].width = 92
src = [
    ("Annuaire des Entreprises (INSEE)", "https://annuaire-entreprises.data.gouv.fr/"),
    ("API Recherche d'entreprises (open data)", "https://recherche-entreprises.api.gouv.fr/"),
    ("Registre des transporteurs de marchandises", "https://www2.transports.developpement-durable.gouv.fr/registres/marchandises/99.pdf"),
    ("Pages Jaunes – La Réunion", "https://www.pagesjaunes.fr/annuaire/departement/reunion-974/transport-marchandises"),
    ("Kompass – Transports et logistique La Réunion", "https://fr.kompass.com/s/transports-et-logistique/10/r/la-reunion/fr_04/"),
    ("Annuaire des transitaires de La Réunion", "https://www.reunion-directory.com/professions/transit-douane.html"),
    ("CCI Réunion (fichiers entreprises, événements)", "https://www.reunion.cci.fr/"),
    ("France Travail / offres logistique 974", "https://candidat.francetravail.fr/offres/recherche"),
]
r = 4
ws.cell(row=r, column=1, value="SOURCES UTILISÉES POUR CONSTRUIRE CE FICHIER").font = F_H2
r += 1
for lib, url in src:
    ws.cell(row=r, column=1, value=lib).font = F_CORPS
    c = ws.cell(row=r, column=2, value=url)
    c.font = F_LIEN; c.hyperlink = url
    r += 1
r += 1
ws.cell(row=r, column=1, value="LIMITES CONNUES").font = F_H2
r += 1
for txt in [
    "• Les SIRET, codes APE et noms de dirigeants non confirmés sont marqués « A VERIFIER » : ils ne sont PAS renseignés au hasard.",
    "• Les emails et téléphones directs ne figurent dans aucune base ouverte : ils se collectent sur le site de l'entreprise ou par téléphone.",
    "• Sainte-Rose, La Plaine-des-Palmistes et Salazie comptent très peu d'entreprises logistiques : le potentiel y est surtout en commerce de proximité et transport agricole.",
    "• Le fichier est une base de départ : il doit vivre. Ajoutez une ligne à chaque entreprise rencontrée sur le terrain.",
]:
    c = ws.cell(row=r, column=1, value=txt); c.font = F_CORPS
    c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws.row_dimensions[r].height = max(15, 14 * (len(txt) // 130 + 1))
    r += 1
ws.sheet_view.showGridLines = False

out = "/home/user/Claude-/prospection/Prospection_Logistique_Nord-Est_Reunion.xlsx"
wb.save(out)
print("OK ->", out)
print("Onglets :", wb.sheetnames)
print("Total entreprises :", sum(n for _, n in compteurs))
