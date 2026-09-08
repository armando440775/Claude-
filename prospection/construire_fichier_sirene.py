# -*- coding: utf-8 -*-
"""
Construit un fichier de prospection EXHAUSTIF a partir de SIRENE (INSEE).

Contrairement au fichier qualitatif (generer_fichier_prospection.py), qui liste
des entreprises reperees une par une, celui-ci balaie TOUS les etablissements
actifs du perimetre dont le code APE correspond a un metier de la logistique.
Chaque ligne sort donc avec un SIRET, un code APE et un effectif officiels.

Deux sources possibles, au choix :

  1. L'API publique recherche-entreprises.api.gouv.fr (par defaut).
     Rien a telecharger, mais il faut un acces internet non filtre.

        python3 construire_fichier_sirene.py
        python3 construire_fichier_sirene.py --champ large
        python3 construire_fichier_sirene.py --perimetre 974

  2. Les fichiers Stock SIRENE telecharges sur data.gouv.fr (hors ligne).
     Telecharger et decompresser StockEtablissement_utf8.csv et
     StockUniteLegale_utf8.csv, puis :

        python3 construire_fichier_sirene.py --source stock \\
            --etablissements StockEtablissement_utf8.csv \\
            --unites StockUniteLegale_utf8.csv

INTEGRITE DES DONNEES
Toutes les lignes viennent de SIRENE : rien n'est invente, rien n'est devine.
Les telephones et emails ne figurent pas dans SIRENE : ils restent a completer.
"""
import argparse
import csv
import gzip
import io
import json
import os
import re
import sys
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    sys.exit("openpyxl manquant : python3 -m pip install openpyxl")
from openpyxl.styles import Alignment, Font

from commun_prospection import (
    AC, AV, BLEU, EFFECTIFS, F_CORPS, F_GRAS, F_H2, F_STITRE, FILL_JAUNE, L,
    NAF, P_ELOG, P_MAGAS, P_OPLOG, P_QUAI, P_TRANS, colonnes, entete_feuille,
    feuille_commune, onglet_pilotage,
)

API = "https://recherche-entreprises.api.gouv.fr/search"
PAUSE = 0.35
ICI = os.path.dirname(os.path.abspath(__file__))
FICHIER_QUALITATIF = os.path.join(ICI, "Prospection_Logistique_Nord-Est_Reunion.xlsx")

# ------------------------------------------------------------------ Cibles APE
# code APE -> (poste d'apprenti pertinent, coeur de cible ?)
NAF_COEUR = {
    "52.10A": (P_OPLOG, True), "52.10B": (P_OPLOG, True),
    "52.24A": (P_QUAI,  True), "52.24B": (P_QUAI,  True),
    "52.29A": (P_QUAI,  True), "52.29B": (P_TRANS, True),
    "49.41A": (P_OPLOG, True), "49.41B": (P_OPLOG, True), "49.41C": (P_MAGAS, True),
    "49.42Z": (P_MAGAS, True), "53.20Z": (P_QUAI,  True), "53.10Z": (P_QUAI, True),
}
NAF_LARGE = {
    "46.31Z": (P_OPLOG, False), "46.39A": (P_OPLOG, False), "46.39B": (P_OPLOG, False),
    "46.46Z": (P_OPLOG, False), "46.73A": (P_MAGAS, False), "46.90Z": (P_MAGAS, False),
    "47.11D": (P_ELOG,  False), "47.11F": (P_ELOG,  False), "47.52B": (P_MAGAS, False),
    "45.31Z": (P_MAGAS, False), "77.32Z": (P_MAGAS, False),
}

# ------------------------------------------------------------------ Perimetres
# Onglet -> communes (libelles SIRENE) et codes postaux
BASSIN = [
    ("01 Saint-Denis",           ["SAINT-DENIS"],              ["97400", "97490", "97417"]),
    ("02 Sainte-Marie",          ["SAINTE-MARIE"],             ["97438"]),
    ("03 Sainte-Suzanne",        ["SAINTE-SUZANNE"],           ["97441"]),
    ("04 Saint-André",           ["SAINT-ANDRE"],              ["97440"]),
    ("05 Bras-Panon",            ["BRAS-PANON"],               ["97412"]),
    ("06 Saint-Benoît",          ["SAINT-BENOIT"],             ["97470", "97437"]),
    ("07 Sainte-Rose & Hauteurs",["SAINTE-ROSE", "SALAZIE", "LA PLAINE-DES-PALMISTES"],
                                                               ["97439", "97433", "97431"]),
]
ZONES_974 = [
    ("01 Nord",  ["SAINT-DENIS", "SAINTE-MARIE", "SAINTE-SUZANNE"], []),
    ("02 Est",   ["SAINT-ANDRE", "BRAS-PANON", "SAINT-BENOIT", "SAINTE-ROSE", "SALAZIE",
                  "LA PLAINE-DES-PALMISTES"], []),
    ("03 Ouest", ["LE PORT", "LA POSSESSION", "SAINT-PAUL", "TROIS-BASSINS", "SAINT-LEU"], []),
    ("04 Sud",   ["LES AVIRONS", "L'ETANG-SALE", "SAINT-LOUIS", "CILAOS", "ENTRE-DEUX",
                  "LE TAMPON", "SAINT-PIERRE", "PETITE-ILE", "SAINT-JOSEPH",
                  "SAINT-PHILIPPE"], []),
]


def sans_accent(t):
    return "".join(c for c in unicodedata.normalize("NFD", str(t or ""))
                   if unicodedata.category(c) != "Mn")


def cle_commune(t):
    return re.sub(r"[^A-Z]", "", sans_accent(t).upper())


def cle_nom(t):
    t = sans_accent(t).upper()
    t = re.sub(r"\(.*?\)", " ", t)
    t = re.sub(r"\b(SARL|SAS|SASU|EURL|SA|SCI|STE|SOCIETE|GROUPE|AGENCE)\b", " ", t)
    return re.sub(r"[^A-Z0-9]", "", t)


def onglet_pour(commune, perimetre):
    table = BASSIN if perimetre == "bassin" else ZONES_974
    c = cle_commune(commune)
    for nom_onglet, communes, _cps in table:
        if any(c == cle_commune(x) for x in communes):
            return nom_onglet
    return None


def priorite(code_eff, coeur):
    gros = {"11", "12", "21", "22", "31", "32", "41", "42", "51", "52", "53"}
    petit = {"01", "02", "03"}
    if code_eff in gros:
        return "P1" if coeur else "P2"
    if code_eff in petit:
        return "P2" if coeur else "P3"
    if code_eff == "00":
        return "P3"
    return "P2" if coeur else "P3"          # effectif non renseigne


def adresse_lisible(et):
    """Reconstitue une adresse a partir des champs SIRENE."""
    if et.get("adresse"):
        return re.sub(r"\s+", " ", et["adresse"]).strip()
    bouts = [et.get("numeroVoieEtablissement"), et.get("indiceRepetitionEtablissement"),
             et.get("typeVoieEtablissement"), et.get("libelleVoieEtablissement")]
    return " ".join(str(b) for b in bouts if b) or AV


def dirigeant(res):
    for d in res.get("dirigeants") or []:
        if (d.get("type_dirigeant") or "").startswith("personne physique"):
            nom = " ".join(x for x in [(d.get("prenoms") or "").title(),
                                       (d.get("nom") or "").upper()] if x).strip()
            if nom:
                return nom, (d.get("qualite") or "Dirigeant").title()
    return AC, "Dirigeant / DG"


# ------------------------------------------------------------------ Source API
def api_get(params, essais=3):
    url = API + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={
        "User-Agent": "CFA-prospection-alternance/1.0 (donnees publiques)",
        "Accept": "application/json"})
    for i in range(essais):
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code == 429:
                time.sleep(2 * (i + 1)); continue
            if e.code >= 500 and i < essais - 1:
                time.sleep(1.5 * (i + 1)); continue
            raise
        except (urllib.error.URLError, TimeoutError):
            if i == essais - 1:
                raise
            time.sleep(1.5 * (i + 1))
    return {"results": []}


def collecter_api(cibles, perimetre):
    """Interroge l'API code APE par code APE, et retourne les etablissements."""
    lignes, vus = [], set()
    requetes = []
    if perimetre == "bassin":
        for _onglet, _communes, cps in BASSIN:
            for cp in cps:
                for naf in cibles:
                    requetes.append({"activite_principale": naf, "code_postal": cp})
    else:
        for naf in cibles:
            requetes.append({"activite_principale": naf, "departement": "974"})

    for i, base in enumerate(requetes, start=1):
        page, pages = 1, 1
        while page <= pages and page <= 400:
            p = dict(base, page=page, per_page=25, etat_administratif="A")
            try:
                data = api_get(p)
            except Exception as e:                                    # noqa: BLE001
                print("  ! echec sur %s (%s)" % (base, e))
                break
            pages = min(data.get("total_pages") or 1, 400)
            for res in data.get("results") or []:
                etabs = (res.get("matching_etablissements")
                         or ([res["siege"]] if res.get("siege") else []))
                for et in etabs:
                    ligne = ligne_depuis_api(res, et, cibles, perimetre)
                    if ligne and ligne["siret"] not in vus:
                        vus.add(ligne["siret"])
                        lignes.append(ligne)
            page += 1
            time.sleep(PAUSE)
        if i % 10 == 0 or i == len(requetes):
            print("  %d/%d requêtes — %d établissements" % (i, len(requetes), len(lignes)))
    return lignes


def ligne_depuis_api(res, et, cibles, perimetre):
    if (et.get("etat_administratif") or "A") != "A":
        return None
    ape = et.get("activite_principale") or res.get("activite_principale") or ""
    if ape not in cibles:
        return None
    commune = et.get("libelle_commune") or ""
    onglet = onglet_pour(commune, perimetre)
    if not onglet:
        return None
    eff = et.get("tranche_effectif_salarie") or res.get("tranche_effectif_salarie") or "NN"
    poste, coeur = cibles[ape]
    dg, fonc = dirigeant(res)
    return {
        "onglet": onglet, "siret": et.get("siret") or "",
        "nom": res.get("nom_complet") or res.get("nom_raison_sociale") or "",
        "ape": ape, "eff": eff, "commune": commune,
        "cp": et.get("code_postal") or "", "adresse": adresse_lisible(et),
        "poste": poste, "coeur": coeur, "dg": dg, "fonc": fonc,
        "siege": bool(et.get("est_siege")), "creation": et.get("date_creation") or "",
    }


# ------------------------------------------------------------------ Source stock
def ouvrir_csv(chemin):
    if chemin.endswith(".gz"):
        return io.TextIOWrapper(gzip.open(chemin, "rb"), encoding="utf-8")
    return io.open(chemin, encoding="utf-8", newline="")


def collecter_stock(chemin_etab, chemin_ul, cibles, perimetre):
    """Deux passes sur les fichiers Stock SIRENE : etablissements puis unites legales."""
    print("Lecture des établissements (fichier volumineux, patientez)…")
    lignes, sirens, lus = [], set(), 0
    with ouvrir_csv(chemin_etab) as f:
        for et in csv.DictReader(f):
            lus += 1
            if lus % 2000000 == 0:
                print("  %d millions de lignes lues…" % (lus // 1000000))
            cp = et.get("codePostalEtablissement") or ""
            if not cp.startswith("974"):
                continue
            if (et.get("etatAdministratifEtablissement") or "") != "A":
                continue
            ape = et.get("activitePrincipaleEtablissement") or ""
            if ape not in cibles:
                continue
            commune = et.get("libelleCommuneEtablissement") or ""
            onglet = onglet_pour(commune, perimetre)
            if not onglet:
                continue
            poste, coeur = cibles[ape]
            sirens.add(et.get("siren"))
            lignes.append({
                "onglet": onglet, "siret": et.get("siret") or "",
                "siren": et.get("siren"), "nom": "",
                "enseigne": et.get("enseigne1Etablissement")
                            or et.get("denominationUsuelleEtablissement") or "",
                "ape": ape,
                "eff": et.get("trancheEffectifsEtablissement") or "NN",
                "commune": commune, "cp": cp, "adresse": adresse_lisible(et),
                "poste": poste, "coeur": coeur, "dg": AC, "fonc": "Dirigeant / DG",
                "siege": (et.get("etablissementSiege") or "").lower() == "true",
                "creation": et.get("dateCreationEtablissement") or "",
            })
    print("  %d établissements retenus (%d lignes parcourues)" % (len(lignes), lus))

    print("Lecture des unités légales (pour les raisons sociales)…")
    noms = {}
    with ouvrir_csv(chemin_ul) as f:
        for ul in csv.DictReader(f):
            s = ul.get("siren")
            if s not in sirens:
                continue
            nom = (ul.get("denominationUniteLegale") or "").strip()
            if not nom:                       # entrepreneur individuel
                nom = " ".join(x for x in [(ul.get("prenom1UniteLegale") or "").title(),
                                           (ul.get("nomUniteLegale") or "").upper()] if x).strip()
            noms[s] = nom
            if len(noms) == len(sirens):
                break
    for l in lignes:
        l["nom"] = noms.get(l["siren"]) or l["enseigne"] or ("SIREN " + str(l["siren"]))
    return lignes


# ------------------------------------------------------------------ Classeur
def noms_deja_connus():
    """Raisons sociales du fichier qualitatif, pour signaler les doublons."""
    if not os.path.exists(FICHIER_QUALITATIF):
        return set()
    wb = load_workbook(FICHIER_QUALITATIF, data_only=True)
    connus = set()
    for onglet in wb.sheetnames:
        if not re.match(r"^0[1-7] ", onglet):
            continue
        ws = wb[onglet]
        for r in range(5, ws.max_row + 1):
            v = ws.cell(row=r, column=2).value
            if v:
                connus.add(cle_nom(v))
    return connus


def construire(lignes, perimetre, champ, sortie, source):
    connus = noms_deja_connus()
    table = BASSIN if perimetre == "bassin" else ZONES_974
    ordre_prio = {"P1": 0, "P2": 1, "P3": 2}
    doublons = 0

    par_onglet = {nom: [] for nom, _c, _cp in table}
    for l in sorted(lignes, key=lambda x: (ordre_prio[priorite(x["eff"], x["coeur"])],
                                           -list(EFFECTIFS).index(x["eff"])
                                           if x["eff"] in EFFECTIFS else 0,
                                           x["nom"])):
        deja = cle_nom(l["nom"]) in connus
        if deja:
            doublons += 1
        notes = []
        if deja:
            notes.append("DÉJÀ dans votre fichier de prospection qualitatif : "
                         "reportez-y vos notes plutôt que de repartir de zéro.")
        notes.append("Établissement %s, actif." % ("siège" if l["siege"] else "secondaire"))
        if l["creation"]:
            notes.append("Créé le %s." % l["creation"])
        notes.append("Extrait de SIRENE le %s." % date.today().strftime("%d/%m/%Y"))
        par_onglet[l["onglet"]].append(L(
            l["nom"], NAF.get(l["ape"], l["ape"]), l["adresse"], l["cp"],
            l["commune"].title(), l["poste"], priorite(l["eff"], l["coeur"]),
            "Vérifié INSEE", " ".join(notes),
            siret=l["siret"], ape=l["ape"], lib_ape=NAF.get(l["ape"], AV),
            eff=EFFECTIFS.get(l["eff"], l["eff"]), dg=l["dg"], fonc=l["fonc"]))

    COLS = colonnes("Activité (libellé NAF)")
    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("00 Mode d'emploi")
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 104
    ws["B2"] = "EXTRACTION SIRENE – ENTREPRISES DE LA LOGISTIQUE"
    ws["B2"].font = Font(name="Calibri", size=18, bold=True, color=BLEU)
    ws["B3"] = ("Source officielle INSEE · extraction du %s · %d établissements actifs"
                % (date.today().strftime("%d/%m/%Y"), len(lignes)))
    ws["B3"].font = F_STITRE
    blocs = [
        ("CE QUE CONTIENT CE FICHIER", ""),
        ("Une extraction, pas une sélection", "Toutes les entreprises du périmètre dont le code APE "
            "correspond à un métier de la logistique, sans exception et sans tri éditorial. "
            "C'est le complément du fichier qualitatif : ici la couverture, là-bas le repérage terrain."),
        ("Données garanties", "Raison sociale, SIRET, code APE, effectif, adresse et commune viennent "
            "directement de SIRENE. Aucune n'a été saisie ni devinée : la colonne « Fiabilité donnée » "
            "indique « Vérifié INSEE » sur toutes les lignes."),
        ("Ce qui manque, et pourquoi", "SIRENE ne contient NI téléphone NI email : ces champs restent "
            "à compléter. Le lien de la colonne S ouvre la fiche officielle de l'entreprise."),
        ("Doublons signalés", "%d ligne(s) existent déjà dans votre fichier de prospection qualitatif : "
            "la colonne Notes vous le dit, pour ne pas perdre les notes d'appel que vous y avez déjà."
            % doublons),
        ("", ""),
        ("COMMENT L'EXPLOITER", ""),
        ("Trier par priorité", "P1 = 10 salariés ou plus sur un code APE cœur de métier : capacité à "
            "tutorer avérée. P2 = 1 à 9 salariés, ou effectif non renseigné. P3 = 0 salarié déclaré : "
            "souvent une société sans personnel, à ne travailler qu'en dernier."),
        ("L'effectif d'abord", "Filtrez la colonne J. Une entreprise de 0 salarié ne prendra pas "
            "d'apprenti ; celles de 10 à 49 salariés sont votre cœur de cible."),
        ("Attention aux sièges", "Un groupe apparaît autant de fois qu'il a d'établissements dans le "
            "périmètre. Un seul appel au siège suffit : la colonne Notes précise siège ou secondaire."),
        ("Actualiser", "SIRENE bouge tous les mois. Relancez construire_fichier_sirene.py pour "
            "regénérer une extraction fraîche — voir l'onglet « 09 Paramètres »."),
    ]
    r = 5
    for titre, txt in blocs:
        if titre and not txt:
            ws.cell(row=r, column=2, value=titre).font = F_H2
            ws.cell(row=r, column=2).fill = FILL_JAUNE
            ws.cell(row=r, column=3).fill = FILL_JAUNE
            r += 1; continue
        if not titre and not txt:
            r += 1; continue
        a = ws.cell(row=r, column=2, value=titre); a.font = F_GRAS
        a.alignment = Alignment(vertical="top", wrap_text=True)
        b = ws.cell(row=r, column=3, value=txt); b.font = F_CORPS
        b.alignment = Alignment(vertical="top", wrap_text=True)
        ws.row_dimensions[r].height = max(15, 13 * (len(txt) // 100 + 1))
        r += 1
    ws.sheet_view.showGridLines = False

    compteurs = []
    for nom_onglet, communes, _cps in table:
        sous = ", ".join(c.title() for c in communes) + \
               " — tous les établissements actifs des codes APE ciblés"
        _, n = feuille_commune(wb, nom_onglet, par_onglet[nom_onglet], sous,
                               prefixe_titre="SIRENE", cols=COLS)
        compteurs.append((nom_onglet, n))

    onglet_pilotage(
        wb, compteurs, "PILOTAGE – EXTRACTION SIRENE",
        "Compteurs automatiques par onglet et par statut",
        "TOTAL EXTRACTION",
        [("Cible d'appels / semaine", "30"),
         ("Commencer par", "les P1 de l'onglet le plus proche de vous"),
         ("Rappel", "les doublons avec le fichier qualitatif sont signalés en colonne V")],
        libelle_volume="VOLUME PAR ONGLET", libelle_unite="Établissements")

    ws = wb.create_sheet("09 Paramètres")
    entete_feuille(ws, "PARAMÈTRES DE L'EXTRACTION",
                   "Pour reproduire exactement le même fichier, ou l'élargir", 3)
    for i, w in enumerate([30, 16, 62], start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    r = 4
    infos = [("Date d'extraction", date.today().strftime("%d/%m/%Y"), ""),
             ("Source", "API recherche-entreprises" if source == "api" else "Fichiers Stock SIRENE",
              "Données INSEE dans les deux cas"),
             ("Périmètre", "Bassin Nord-Est" if perimetre == "bassin" else "Département 974 entier", ""),
             ("Champ", "Cœur logistique" if champ == "coeur" else "Élargi (négoce, distribution)", ""),
             ("Établissements retenus", str(len(lignes)), "actifs uniquement"),
             ("Déjà connus", str(doublons), "présents dans le fichier qualitatif")]
    for lib, val, com in infos:
        ws.cell(row=r, column=1, value=lib).font = F_GRAS
        ws.cell(row=r, column=2, value=val).font = F_CORPS
        ws.cell(row=r, column=3, value=com).font = F_CORPS
        r += 1
    r += 1
    ws.cell(row=r, column=1, value="CODES APE INTERROGÉS").font = F_H2
    r += 1
    cibles = dict(NAF_COEUR)
    if champ == "large":
        cibles.update(NAF_LARGE)
    for code in sorted(cibles):
        poste, coeur = cibles[code]
        ws.cell(row=r, column=1, value=code).font = F_GRAS
        ws.cell(row=r, column=2, value="cœur" if coeur else "élargi").font = F_CORPS
        c = ws.cell(row=r, column=3, value="%s — poste visé : %s" % (NAF.get(code, ""), poste))
        c.font = F_CORPS; c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    r += 1
    ws.cell(row=r, column=1, value="COMMANDE UTILISÉE").font = F_H2
    ws.cell(row=r + 1, column=1,
            value="python3 construire_fichier_sirene.py --perimetre %s --champ %s"
                  % (perimetre, champ)).font = F_CORPS
    ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=3)
    ws.sheet_view.showGridLines = False

    wb.save(sortie)
    return len(lignes), doublons


# ------------------------------------------------------------------ Programme
def main():
    ap = argparse.ArgumentParser(description="Extraction SIRENE des entreprises de la logistique")
    ap.add_argument("--source", choices=["api", "stock"], default="api")
    ap.add_argument("--perimetre", choices=["bassin", "974"], default="bassin")
    ap.add_argument("--champ", choices=["coeur", "large"], default="coeur")
    ap.add_argument("--etablissements", help="StockEtablissement_utf8.csv (source stock)")
    ap.add_argument("--unites", help="StockUniteLegale_utf8.csv (source stock)")
    ap.add_argument("--sortie", default=os.path.join(ICI, "Prospection_SIRENE_Logistique_974.xlsx"))
    a = ap.parse_args()

    cibles = dict(NAF_COEUR)
    if a.champ == "large":
        cibles.update(NAF_LARGE)

    print("Extraction SIRENE — périmètre %s, champ %s, %d codes APE\n"
          % (a.perimetre, a.champ, len(cibles)))
    if a.source == "stock":
        if not a.etablissements or not a.unites:
            sys.exit("Source stock : indiquez --etablissements et --unites.\n"
                     "Fichiers à télécharger sur https://www.data.gouv.fr (base SIRENE, "
                     "« Stock Etablissement » et « Stock Unite Legale »).")
        lignes = collecter_stock(a.etablissements, a.unites, cibles, a.perimetre)
    else:
        try:
            lignes = collecter_api(cibles, a.perimetre)
        except urllib.error.URLError as e:
            sys.exit("\nImpossible de joindre l'API (%s).\n"
                     "Si votre réseau la bloque, utilisez la source « stock » : voir l'aide en tête "
                     "de ce fichier." % e)

    if not lignes:
        sys.exit("Aucun établissement trouvé. Vérifiez le périmètre et la connexion.")

    total, doublons = construire(lignes, a.perimetre, a.champ, a.sortie, a.source)
    print("\nOK -> %s" % a.sortie)
    print("     %d établissements actifs, %d déjà présents dans le fichier qualitatif" %
          (total, doublons))


if __name__ == "__main__":
    main()
