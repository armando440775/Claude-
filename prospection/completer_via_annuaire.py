# -*- coding: utf-8 -*-
"""
Complete automatiquement le fichier de prospection a partir de l'API officielle
de l'INSEE / DINUM : recherche-entreprises.api.gouv.fr (gratuite, sans cle).

Remplit, pour chaque entreprise du fichier :
    - le SIRET de l'etablissement de la commune
    - le code APE et son libelle
    - la tranche d'effectif salarie
    - le nom du dirigeant et sa qualite
    - l'adresse quand elle manque
et signale les etablissements FERMES (a ne pas appeler).

REGLE D'OR : le script ne remplit une case QUE s'il est sur de l'entreprise.
En cas de doute, il laisse "A VERIFIER" et ecrit la raison dans le journal.
Il ne modifie jamais le fichier d'origine : il ecrit une copie _COMPLETE.xlsx.

Usage :
    python3 completer_via_annuaire.py
    python3 completer_via_annuaire.py mon_fichier.xlsx
    python3 completer_via_annuaire.py --inplace      (ecrase le fichier d'origine)
"""
import csv
import json
import os
import re
import sys
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from difflib import SequenceMatcher

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("openpyxl n'est pas installe.\n"
             "Lancez d'abord :  python3 -m pip install openpyxl")

API = "https://recherche-entreprises.api.gouv.fr/search"
PAUSE = 0.45           # secondes entre deux appels (l'API limite a 7 appels/s)
SEUIL_SUR = 0.80       # score au-dela duquel on accepte sans commune identique
SEUIL_COMMUNE = 0.66   # score suffisant si la commune correspond
ECART_MINI = 0.10      # ecart mini avec le 2e candidat pour trancher

AV = "A VERIFIER"
AC = "A COMPLETER"
PLACEHOLDERS = {AV, AC, "", None, "A VÉRIFIER", "A COMPLÉTER"}

# Colonnes (voir generer_fichier_prospection.py)
C_NOM, C_ADR, C_CP, C_COM = 2, 4, 5, 6
C_SIRET, C_APE, C_LIBAPE, C_EFF = 7, 8, 9, 10
C_DG, C_FONC, C_FIAB, C_NOTES = 11, 12, 18, 22
LIGNE_ENTETE = 4

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

MOTS_VIDES = {
    "sarl", "sas", "sasu", "eurl", "sa", "sci", "scop", "snc", "ste", "societe",
    "groupe", "agence", "reunion", "la", "le", "les", "de", "du", "des", "d",
    "et", "oi", "ocean", "indien", "974", "site", "presence", "enseigne",
}


# ------------------------------------------------------------------ outils
def sans_accent(t):
    t = unicodedata.normalize("NFD", str(t))
    return "".join(c for c in t if unicodedata.category(c) != "Mn")


def normaliser(t):
    """Reduit un nom a ses mots significatifs, pour la comparaison."""
    t = sans_accent(t).lower()
    t = re.sub(r"\(.*?\)", " ", t)          # retire les parentheses
    t = re.split(r"\s[-–]\s", t)[0]          # coupe apres " - " (ex: "- agence Reunion")
    t = re.sub(r"[^a-z0-9]+", " ", t)
    mots = [m for m in t.split() if m and m not in MOTS_VIDES]
    return " ".join(mots) if mots else sans_accent(t).lower().strip()


def commune_base(t):
    """'Sainte-Clotilde (Saint-Denis)' -> 'sainte clotilde'."""
    return normaliser(re.sub(r"\(.*?\)", " ", str(t)))


def variantes_commune(t):
    """Toutes les facons de nommer la commune d'une ligne.

    'Sainte-Clotilde (Saint-Denis)' -> {'sainte clotilde', 'saint denis'}
    car l'INSEE rattache Sainte-Clotilde a la commune de Saint-Denis.
    """
    t = str(t or "")
    v = {commune_base(t)}
    for interne in re.findall(r"\((.*?)\)", t):
        v.add(normaliser(interne))
    return {x for x in v if x}


def meme_commune(libelle_insee, cellule_commune):
    return normaliser(libelle_insee or "") in variantes_commune(cellule_commune)


def est_vide(v):
    return v is None or str(v).strip() in {p for p in PLACEHOLDERS if p is not None} \
        or str(v).strip() == ""


# ------------------------------------------------------------------ API
def api_get(params, essais=3):
    url = API + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={
        "User-Agent": "CFA-prospection-alternance/1.0 (usage interne, donnees publiques)",
        "Accept": "application/json",
    })
    for i in range(essais):
        try:
            with urllib.request.urlopen(req, timeout=25) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code == 429:                      # trop d'appels : on souffle
                time.sleep(2 * (i + 1))
                continue
            if e.code >= 500 and i < essais - 1:
                time.sleep(1.5 * (i + 1))
                continue
            raise
        except (urllib.error.URLError, TimeoutError):
            if i == essais - 1:
                raise
            time.sleep(1.5 * (i + 1))
    return {"results": []}


def libelle_ape(code, resultat):
    for cle in ("libelle_activite_principale", "libelle_activite_principale_entreprise"):
        if resultat.get(cle):
            return resultat[cle]
    return NAF.get(code, AV)


def dirigeant_lisible(res):
    for d in res.get("dirigeants") or []:
        if (d.get("type_dirigeant") or "").startswith("personne physique"):
            nom = " ".join(x for x in [(d.get("prenoms") or "").title(),
                                       (d.get("nom") or "").upper()] if x).strip()
            if nom:
                return nom, (d.get("qualite") or "Dirigeant").title()
    for d in res.get("dirigeants") or []:      # personne morale (holding gerante)
        if d.get("denomination"):
            return d["denomination"], (d.get("qualite") or "Dirigeant").title()
    return None, None


def etablissement_pour(res, commune):
    """Prefere l'etablissement situe dans la commune de la ligne, sinon le siege."""
    siege = res.get("siege") or {}
    for et in (res.get("matching_etablissements") or []):
        if meme_commune(et.get("libelle_commune"), commune):
            return et
    return siege


def score_nom(nom_ligne, res):
    n = normaliser(nom_ligne)
    candidats = [res.get("nom_complet") or "", res.get("nom_raison_sociale") or ""]
    candidats += (res.get("siege") or {}).get("liste_enseignes") or []
    best = 0.0
    for c in candidats:
        if not c:
            continue
        cn = normaliser(c)
        r = SequenceMatcher(None, n, cn).ratio()
        if n and (n in cn or cn in n):        # inclusion = tres bon signe
            r = max(r, 0.90)
        best = max(best, r)
    return best


def chercher(nom, commune):
    """Retourne (resultat, etablissement, score, motif) ou (None, None, score, motif)."""
    requetes = [
        {"q": "%s %s" % (nom, re.sub(r"\(.*?\)", "", str(commune)).strip()),
         "departement": "974", "per_page": 10},
        {"q": nom, "departement": "974", "per_page": 10},
    ]
    vus, notes = {}, []
    for p in requetes:
        try:
            data = api_get(p)
        except Exception as e:                    # noqa: BLE001
            return None, None, 0.0, "erreur reseau : %s" % e
        for res in data.get("results") or []:
            s = score_nom(nom, res)
            if any(
                meme_commune((et or {}).get("libelle_commune"), commune)
                for et in ([res.get("siege")] + (res.get("matching_etablissements") or []))
            ):
                s = min(1.0, s + 0.12)
            cle = res.get("siren")
            if cle and s > vus.get(cle, (0.0, None))[0]:
                vus[cle] = (s, res)
        time.sleep(PAUSE)
        if vus and max(v[0] for v in vus.values()) >= SEUIL_SUR:
            break

    if not vus:
        return None, None, 0.0, "aucun resultat dans le 974"

    classes = sorted(vus.values(), key=lambda x: -x[0])
    meilleur, res = classes[0]
    second = classes[1][0] if len(classes) > 1 else 0.0
    et = etablissement_pour(res, commune)
    ici = meme_commune(et.get("libelle_commune"), commune)

    # Regle de securite : on ne renseigne JAMAIS un SIRET situe dans une autre
    # commune que celle de la ligne. Un homonyme au Sud vous ferait appeler
    # la mauvaise entreprise.
    if not ici:
        return None, None, meilleur, (
            "trouvee sous le nom \"%s\" mais a %s, pas a %s -> a verifier a la main"
            % (res.get("nom_complet"), (et.get("libelle_commune") or "?"), commune))

    if meilleur >= SEUIL_SUR and (meilleur - second >= ECART_MINI or meilleur >= 0.93):
        return res, et, meilleur, "correspondance sure (commune identique)"
    if meilleur >= SEUIL_COMMUNE and meilleur - second >= ECART_MINI:
        return res, et, meilleur, "correspondance probable (commune identique)"
    notes.append("candidats trop proches : %s (score %.2f, 2e a %.2f)"
                 % (res.get("nom_complet"), meilleur, second))
    return None, None, meilleur, "doute -> " + " ; ".join(notes)


# ------------------------------------------------------------------ traitement
def traiter(chemin_in, chemin_out, forcer=False):
    wb = load_workbook(chemin_in)
    onglets = [n for n in wb.sheetnames if re.match(r"^0[1-7] ", n)]
    journal, stats = [], {"traitees": 0, "completees": 0, "doutes": 0, "fermees": 0}

    for nom_onglet in onglets:
        ws = wb[nom_onglet]
        print("\n=== %s ===" % nom_onglet)
        for r in range(LIGNE_ENTETE + 1, ws.max_row + 1):
            nom = ws.cell(row=r, column=C_NOM).value
            commune = ws.cell(row=r, column=C_COM).value
            if not nom or str(nom).strip().startswith(("À prospecter", "A prospecter")):
                continue
            # Lignes "methode" (ex : "EHPAD du bassin - a lister") : rien a chercher
            type_structure = str(ws.cell(row=r, column=3).value or "").lower()
            if "à lister" in type_structure or "à défricher" in type_structure:
                print("  · %-46s ligne méthode, ignorée" % str(nom)[:46])
                continue
            manquants = [c for c in (C_SIRET, C_APE, C_LIBAPE, C_EFF, C_DG)
                         if est_vide(ws.cell(row=r, column=c).value)]
            if not manquants and not forcer:
                print("  · %-46s deja complete" % str(nom)[:46])
                continue

            stats["traitees"] += 1
            res, et, score, motif = chercher(str(nom), str(commune or ""))
            if not res:
                stats["doutes"] += 1
                print("  ? %-46s %s" % (str(nom)[:46], motif))
                journal.append([nom_onglet, r, nom, commune, "", "NON COMPLETE",
                                "%.2f" % score, motif])
                continue

            siret = (et or {}).get("siret") or ""
            ape = (et or {}).get("activite_principale") or res.get("activite_principale") or ""
            eff = (et or {}).get("tranche_effectif_salarie") \
                or res.get("tranche_effectif_salarie") or "NN"
            dg, qualite = dirigeant_lisible(res)
            ferme = ((et or {}).get("etat_administratif") or res.get("etat_administratif")) == "F"

            def poser(col, valeur, ecraser=False):
                if valeur in (None, ""):
                    return
                cell = ws.cell(row=r, column=col)
                if est_vide(cell.value) or ecraser:
                    cell.value = valeur

            poser(C_SIRET, siret, ecraser=forcer)
            ws.cell(row=r, column=C_SIRET).number_format = "@"
            poser(C_APE, ape, ecraser=forcer)
            poser(C_LIBAPE, libelle_ape(ape, res), ecraser=forcer)
            poser(C_EFF, EFFECTIFS.get(eff, eff), ecraser=True)
            if dg:
                poser(C_DG, dg, ecraser=forcer)
                poser(C_FONC, qualite)
            adr = (et or {}).get("adresse")
            if adr and est_vide(ws.cell(row=r, column=C_ADR).value):
                ws.cell(row=r, column=C_ADR).value = adr
            ws.cell(row=r, column=C_FIAB).value = "Vérifié INSEE"

            if ferme:
                stats["fermees"] += 1
                ancien = ws.cell(row=r, column=C_NOTES).value or ""
                ws.cell(row=r, column=C_NOTES).value = (
                    "⚠ ETABLISSEMENT FERME selon l'INSEE : ne pas appeler, "
                    "verifier s'il existe un autre etablissement actif. " + str(ancien))

            stats["completees"] += 1
            print("  ✓ %-46s %s | %s | %s%s"
                  % (str(nom)[:46], siret, ape, EFFECTIFS.get(eff, eff),
                     "  [FERME]" if ferme else ""))
            journal.append([nom_onglet, r, nom, commune, siret, "COMPLETEE",
                            "%.2f" % score, ("ETABLISSEMENT FERME - " if ferme else "") + motif])

    wb.save(chemin_out)
    dossier = os.path.dirname(os.path.abspath(chemin_out))
    chemin_journal = os.path.join(dossier, "journal_enrichissement.csv")
    with open(chemin_journal, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Onglet", "Ligne", "Entreprise", "Commune", "SIRET trouve",
                    "Resultat", "Score", "Commentaire"])
        w.writerows(journal)

    print("\n" + "=" * 62)
    print("Lignes traitees   : %d" % stats["traitees"])
    print("Completees INSEE  : %d" % stats["completees"])
    print("Laissees en doute : %d  (a faire a la main via le lien colonne S)"
          % stats["doutes"])
    print("Etablissements fermes reperes : %d" % stats["fermees"])
    print("Fichier   : %s" % chemin_out)
    print("Journal   : %s" % chemin_journal)
    print("=" * 62)


def main():
    args = [a for a in sys.argv[1:]]
    inplace = "--inplace" in args
    forcer = "--force" in args
    args = [a for a in args if not a.startswith("--")]
    ici = os.path.dirname(os.path.abspath(__file__))
    src = args[0] if args else os.path.join(
        ici, "Prospection_Logistique_Nord-Est_Reunion.xlsx")
    if not os.path.exists(src):
        sys.exit("Fichier introuvable : %s" % src)
    out = src if inplace else src.replace(".xlsx", "_COMPLETE.xlsx")
    print("Source : %s" % src)
    print("Sortie : %s\n" % out)
    try:
        traiter(src, out, forcer=forcer)
    except urllib.error.URLError as e:
        sys.exit("\nImpossible de joindre l'API (%s).\n"
                 "Verifiez votre connexion internet, ou le pare-feu de l'etablissement." % e)


if __name__ == "__main__":
    main()
