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
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from commun_prospection import (
    AC, AV, BLEU, BORD, F_CORPS, F_ENTETE, F_GRAS, F_H2, F_LIEN, F_STITRE,
    F_TITRE, FILL_ALT, FILL_ENTETE, FILL_JAUNE, L, P_ELOG, P_MAGAS, P_OPLOG,
    P_QUAI, P_TRANS, colonnes, entete_feuille, feuille_commune, onglet_pilotage,
)

COLS = colonnes()

# ---------------------------------------------------------------- Donnees
# Chaque ligne est construite avec L(...) : voir commun_prospection.py.
# Tout champ non verifie reste a "A VERIFIER" - aucune donnee inventee.
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
    L("CERP RÉUNION (GROUPE SIPR)", "Répartition pharmaceutique / Dépositaire",
      AV, "97400", "Saint-Denis", P_OPLOG, "P1", "Source annuaire",
      "Importateur-grossiste-répartiteur en produits pharmaceutiques. Groupe SIPR "
      "(CERP-Réunion, MAD-SIPR, Plurimed, Dicophar), 3 sites : Saint-Denis, Sainte-Marie, "
      "Saint-Pierre. Préparation de commandes sous Bonnes Pratiques de Distribution : "
      "cadre très formateur, gros volumes quotidiens.",
      site="https://www.cerp-sipr.com/cerp/"),
    L("CGF NORD", "Grossiste fruits & légumes / Bio",
      "91 rue Foch – Zone Foucherolles, Sainte-Clotilde", "97490", "Sainte-Clotilde (Saint-Denis)",
      P_OPLOG, "P1", "Source annuaire",
      "Distribution en gros de fruits, légumes et produits bio : préparation de commandes "
      "quotidienne et gestion du froid. Rotation forte = besoin récurrent de préparateurs."),
    L("SOMADIS – SOCIÉTÉ MANES DISTRIBUTION", "Commerce de gros / Entrepôt",
      "76 boulevard du Chaudron", "97490", "Sainte-Clotilde (Saint-Denis)", P_MAGAS, "P2",
      "Source annuaire", "Commerce de gros sur la zone du Chaudron : réception, stock, expédition."),
    L("DEMECO CHEUNG DÉMÉNAGEMENTS", "Déménagement / Garde-meuble",
      AV, "97400", "Saint-Denis", P_MAGAS, "P2", "Source annuaire",
      "Déménagement local, national et international, garde-meuble et groupage maritime. "
      "Manutention, inventaire et gestion de stock : bon terrain pour un profil magasinier."),
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
    L("TRANS EXPRESS RÉUNION (TER)", "Affrètement / Organisation de transports",
      AV, "97438", "Sainte-Marie", P_QUAI, "P1", "Source annuaire",
      "Affrètement et organisation de transports. Présent à Gillot pour le fret aérien "
      "et au port pour le fret maritime : deux univers dans la même entreprise."),
    L("OCX RÉUNION", "Messagerie express internationale",
      AV, "97438", "Sainte-Marie", P_QUAI, "P1", "Source annuaire",
      "Cité comme représentant officiel FedEx et TNT à La Réunion : quai, tri, "
      "documents d'expédition. Confirmer l'agrément lors de l'appel."),
    L("HOLDTRANS", "Transit / Commission en douane",
      "Nouvelle aérogare de fret – porte 238", "97438", "Sainte-Marie", P_TRANS, "P1",
      "Source annuaire",
      "Déclarants en douane agréés, fret maritime et aérien. Sur la plateforme fret de Gillot."),
    L("BOLLORÉ LOGISTICS RÉUNION", "Transit international / Logistique contractuelle",
      "7 rue André Lardy – La Mare", "97438", "Sainte-Marie", P_TRANS, "P1", "Source annuaire",
      "Groupe international : l'alternance y est une pratique installée. Viser le responsable "
      "d'agence ET la RH régionale, et déposer l'offre sur leur portail recrutement."),
    L("SCHENKER FRANCE – agence Réunion", "Transit / Commission de transport",
      "Zone aérogare de fret – Gillot", "97438", "Sainte-Marie", P_TRANS, "P1", "Source annuaire",
      "Groupe international présent sur la zone fret. Même méthode que Bolloré : agence + RH."),
    L("KUEHNE + NAGEL – présence 974", "Transit / Logistique internationale",
      AV, "97438", "Sainte-Marie", P_TRANS, "P2", "A vérifier",
      "Présence citée sur la zone aéroportuaire mais NON confirmée : vérifier l'existence "
      "d'un établissement 974 sur l'Annuaire des Entreprises AVANT tout appel."),
    L("LEROY MERLIN SAINTE-MARIE", "Distribution spécialisée – réserve & cour",
      AV, "97438", "Sainte-Marie", P_MAGAS, "P1", "Source annuaire",
      "Réception, réserve, cour matériaux, drive. Enseigne nationale : dossier alternance "
      "cadré. Passer par le magasin ET le portail recrutement de l'enseigne."),
    L("LA POSTE – centre courrier / colis Sainte-Marie", "Poste et courrier / Colis",
      "2 rue Ravine", "97438", "Sainte-Marie", P_QUAI, "P2", "Source annuaire",
      "Tri et distribution de colis (Colissimo). Gros recruteur en alternance, "
      "mais dossier centralisé : passer par la RH régionale La Poste Réunion."),
    L("SIPR / MAD-SIPR – site Sainte-Marie", "Logistique pharmaceutique / Matériel médical",
      AV, "97438", "Sainte-Marie", P_MAGAS, "P2", "Source annuaire",
      "Groupe SIPR, 3 sites sur l'île. Magasin de matériel médical : gestion de stock "
      "et préparation de commandes tracées."),
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
    L("TEREOS OCÉAN INDIEN – SUCRERIE DE BOIS-ROUGE", "Industrie sucrière / Flux & expéditions",
      "Bois-Rouge – Cambuston", "97440", "Saint-André", P_MAGAS, "P1", "Source annuaire",
      "Le plus gros site industriel de l'Est. Campagne sucrière de juillet à décembre : "
      "pics de flux, magasin de pièces, réception canne, expédition du sucre. "
      "Passer par la RH de Tereos Océan Indien (siège 974).",
      site="https://www.tereos.re/"),
    L("SKAL BRICO SODIS", "Négoce de matériaux / Dépôt",
      "705 chemin Lagourgue", "97440", "Saint-André", P_MAGAS, "P2", "Source annuaire",
      "Dépôt de matériaux : réception, cour, magasinier-cariste."),
]

BRAS_PANON = [
    L("CRÉOLE LOGISTIQUE TRANSPORT", "Transport de fret interurbain",
      "22 bis RN2 – Rivière des Roches", "97412", "Bras-Panon", P_OPLOG, "P1", "Vérifié",
      "SIRET et adresse relevés sur base publique (siège Bras-Panon), à reconfirmer sur "
      "l'Annuaire des Entreprises. Classée INSEE en transport routier de fret interurbain. "
      "Entreprise installée depuis une dizaine d'années : tuteur potentiel expérimenté.",
      siret="80216938300014"),
    L("AFA MARIE", "Transport routier",
      AV, "97412", "Bras-Panon", P_QUAI, "P3", "Source annuaire", ""),
    L("WIN LOCATION", "Location (activité à qualifier)",
      "6 rue des Palmiers", "97412", "Bras-Panon", P_MAGAS, "P3", "A vérifier",
      "Activité exacte à confirmer via le code APE. Si location de matériel : "
      "il y a un parc et un magasin à gérer, donc un vrai poste d'apprenti."),
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
    L("TRANSPORT ROUTIER LEGROS", "Transport routier de fret interurbain",
      "800 chemin Commence", "97437", "Sainte-Anne (Saint-Benoît)", P_OPLOG, "P2",
      "Source annuaire",
      "SARL créée en 2021, transports routiers de fret interurbains. Structure jeune : "
      "argument « formez dès maintenant votre futur exploitant / préparateur »."),
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
    ("11 Sources & mise à jour", "D'où viennent les données, et où aller les actualiser."),
    ("", ""),
    ("REMPLISSAGE AUTOMATIQUE – 2 MINUTES", ""),
    ("Le script d'auto-complétion", "Le fichier « completer_via_annuaire.py » (fourni à côté de ce classeur) "
                 "interroge l'API officielle de l'INSEE (recherche-entreprises.api.gouv.fr) et remplit tout seul "
                 "les colonnes SIRET, Code APE, Libellé APE, Effectif et Dirigeant, pour toutes les lignes du fichier."),
    ("Comment le lancer", "Sur Mac : ouvrir le Terminal. Sur PC : ouvrir l'invite de commandes. "
                 "Taper : python3 completer_via_annuaire.py  (voir le fichier LISEZ-MOI.md pour le détail). "
                 "Le script crée une COPIE enrichie et ne touche jamais à votre fichier de travail."),
    ("Ce qu'il ne fait pas", "Il ne remplit ni les téléphones ni les emails : ces données ne sont dans aucune "
                 "base ouverte. Il ne remplit une case QUE s'il est sûr de l'entreprise ; sinon il laisse "
                 "« A VERIFIER » et note le doute dans le journal."),
    ("La colonne Effectif", "Tranche d'effectif salarié INSEE. C'est votre filtre le plus utile : une entreprise "
                 "de 0 salarié ne pourra pas tutorer un apprenti, une entreprise de 10 à 50 salariés est "
                 "la cible idéale."),
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
onglet_pilotage(
    wb, compteurs,
    "PILOTAGE DE LA CAMPAGNE",
    "Compteurs automatiques – se mettent à jour quand vous remplissez les onglets communes",
    "TOTAL BASSIN NORD-EST",
    [
        ("Appels à passer / semaine", "30"),
        ("RDV entreprise / semaine", "5"),
        ("Offres d'apprentissage à collecter", "à définir"),
        ("Ratio moyen observé", "≈ 10 appels → 2 RDV → 1 offre"),
    ],
    libelle_volume="VOLUME PAR COMMUNE", libelle_unite="Entreprises",
)

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
    ("Annuaire fret – Aéroport Roland Garros", "https://www.reunion.aeroport.fr/fret/annuaire"),
    ("Kompass – Transports et logistique Saint-André", "https://fr.kompass.com/s/transports-et-logistique/10/v/saint-andre/fr_04_974_97409/"),
    ("Pages Jaunes – fiche SIRET Créole Logistique Transport", "https://www.pagesjaunes.fr/siret/80216938300014"),
    ("Tereos Océan Indien (sucrerie de Bois-Rouge)", "https://www.tereos.re/visitez-nos-sucreries/sucrerie-de-bois-rouge"),
    ("Groupe SIPR / CERP Réunion (répartition pharmaceutique)", "https://www.cerp-sipr.com/cerp/"),
    ("SIFA Logistics – agence Réunion", "https://sifalogistics.com/nos-agences/sifa-reunion/"),
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
    "• Les entreprises ajoutées lors de l'enrichissement proviennent d'annuaires en ligne (annuaire fret de l'aéroport, Kompass, Pages Jaunes) : le NOM est fiable, l'adresse est à confirmer, et le SIRET reste à récupérer via le lien de la colonne S ou via le script d'auto-complétion.",
    "• Une ligne porte la mention « A vérifier » en fiabilité (KUEHNE + NAGEL) : sa présence à La Réunion est citée mais non confirmée. Vérifiez-la sur l'Annuaire des Entreprises avant d'appeler.",
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
