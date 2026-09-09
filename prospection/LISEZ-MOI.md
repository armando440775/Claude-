# Prospection apprentissage — bassin Nord-Est de La Réunion

De Saint-Denis à Sainte-Rose, hauteurs comprises. Cible : structures susceptibles de
signer un **contrat d'apprentissage** sur un métier de la logistique (16-25 ans).

## Ce que contient ce dossier

**Deux fichiers de travail, même structure, même méthode :**

| Fichier | À quoi ça sert |
|---|---|
| `Prospection_Logistique_Nord-Est_Reunion.xlsx` | **Le privé.** 60 entreprises de la logistique, du transport et de la distribution. |
| `Prospection_Public_Sante_Nord-Est_Reunion.xlsx` | **Le public et la santé.** 39 structures : communes, CCAS, intercommunalités, Département, Région, lycées, CHU, cliniques, EHPAD. |

Double-cliquez, c'est tout. Une seule structure figure dans les deux fichiers — le **GHER**
à Saint-Benoît : ne passez pas deux appels.

**Les outils :**

| Fichier | À quoi ça sert |
|---|---|
| `completer_via_annuaire.py` | Remplit tout seul SIRET / code APE / effectif / dirigeant depuis l'INSEE. **Fonctionne sur les deux classeurs.** |
| `generer_fichier_prospection.py` | Régénère le classeur « privé » à zéro. |
| `generer_fichier_public_sante.py` | Régénère le classeur « public & santé » à zéro. |
| `commun_prospection.py` | Briques partagées par les deux générateurs (charte, colonnes, mise en page). Ne s'exécute pas seul. |
| `Modeles_Email_Prospection.xlsx` | **Le système d'email.** Tous les modèles d'email prêts à l'emploi (entreprises, secteur public, candidats). |
| `generer_emails_personnalises.py` | Génère un brouillon d'email personnalisé par entreprise à partir d'un des deux classeurs. |
| `modeles_email.py` | Vos coordonnées + le texte de tous les modèles. **C'est le fichier à personnaliser.** |
| `generer_modeles_email.py` | Régénère `Modeles_Email_Prospection.xlsx` à partir de `modeles_email.py`. |

⚠️ Les scripts `generer_*` **écrasent** le classeur et donc votre suivi de prospection.
Ne les lancez pas une fois que vous avez commencé à remplir les colonnes « Statut » et « Notes ».

## Les onglets du fichier « privé »

- **00 Mode d'emploi** — à lire en premier.
- **01 à 07** — un onglet par commune, dans l'ordre de la route littorale : Saint-Denis,
  Sainte-Marie, Sainte-Suzanne, Saint-André, Bras-Panon, Saint-Benoît, puis Sainte-Rose
  et les hauteurs. Vous pouvez faire une tournée terrain onglet par onglet.
- **08 Pilotage** — compteurs automatiques (à contacter / RDV / contrats signés) par commune.
- **09 Codes APE à cibler** — pour élargir vous-même le fichier depuis l'Annuaire des Entreprises.
- **10 Script d'appel** — trame d'appel à froid, 5 arguments employeur, réponses aux objections.
- **11 Sources & mise à jour** — d'où viennent les données.

## Les onglets du fichier « public & santé »

- **00 Mode d'emploi**, puis un onglet par famille : **01** communes et CCAS, **02**
  intercommunalités (CIREST, CINOR, syndicats), **03** Département / Région / Rectorat /
  Université / CROUS / SDIS, **04** lycées et collèges, **05** hôpital public (CHU Félix
  Guyon, GHER), **06** cliniques et santé privée, **07** EHPAD, dialyse et médico-social.
- **08 Pilotage** — mêmes compteurs automatiques.
- **09 Apprentissage dans le public** — **l'onglet à lire avant le premier appel** (voir
  ci-dessous).
- **10 Script d'appel (public)** — on ne parle pas à une mairie comme à un transporteur.
- **11 Sources & mise à jour**.

### Le piège du calendrier public

Dans le privé, on prospecte au printemps pour la rentrée. Dans la fonction publique
**territoriale**, le CNFPT ne finance les frais de formation que si la collectivité s'est
**recensée** pendant sa campagne annuelle, qui se déroule en début d'année civile (celle de
2026 était ouverte du 19 janvier au 20 mars). Une collectivité non recensée ne sera pas
financée, même avec un excellent candidat.

Conséquence : **vos rendez-vous d'octobre à décembre ne servent pas à signer, ils servent à
faire inscrire la collectivité au recensement de janvier.** Deux autres points à connaître :
seuls les niveaux 3 à 5 (CAP → BTS) sont financés — vos titres logistique sont dans la bonne
fourchette — et le métier visé doit figurer sur la liste des métiers en tension du CNFPT :
**vérifiez-le avant de promettre un financement à un DRH.** Le détail (pièces APF/APC, délais,
objections) est dans l'onglet 09.

Attention : montants, quotas et dates changent chaque année. Revérifiez sur cnfpt.fr avant
d'annoncer un chiffre.

## Remplir automatiquement les SIRET, APE, effectifs et dirigeants

Le script interroge l'API officielle **recherche-entreprises.api.gouv.fr** (gratuite, sans
inscription, données INSEE) et complète le fichier à votre place.

**Une seule fois — installer Python et la brique Excel :**

- **Mac** : Python est déjà là. Ouvrez *Terminal* (Applications → Utilitaires), puis tapez :
  ```
  python3 -m pip install openpyxl
  ```
- **PC Windows** : installez Python depuis <https://www.python.org/downloads/> en cochant
  **« Add Python to PATH »**, puis ouvrez *Invite de commandes* et tapez :
  ```
  py -m pip install openpyxl
  ```

**À chaque fois — lancer le remplissage :**

```
cd <le dossier prospection>
python3 completer_via_annuaire.py          # Mac — fichier « privé » par défaut
py completer_via_annuaire.py               # Windows

# pour le fichier public & santé, indiquez-le :
python3 completer_via_annuaire.py Prospection_Public_Sante_Nord-Est_Reunion.xlsx
```

Le script produit :

- `Prospection_Logistique_Nord-Est_Reunion_COMPLETE.xlsx` — la copie enrichie
  (**votre fichier d'origine n'est jamais modifié**) ;
- `journal_enrichissement.csv` — ligne par ligne, ce qui a été complété et ce qui ne l'a pas été.

Options : `--inplace` pour écraser le fichier d'origine, `--force` pour recalculer
même les lignes déjà renseignées.

### Ce que le script garantit

- Il ne remplit une case **que s'il est sûr de l'entreprise**. En cas de doute, il laisse
  « A VERIFIER » et explique pourquoi dans le journal.
- Il **refuse un homonyme situé dans une autre commune**. Un « WIN LOCATION » à Saint-Pierre
  ne sera jamais collé sur la ligne de Bras-Panon.
- Il signale en toutes lettres les **établissements fermés** selon l'INSEE : vous ne perdez
  plus un appel sur une entreprise qui n'existe plus.
- Il n'écrase jamais une donnée que vous avez saisie à la main.
- Il ignore les lignes « à lister » (autres lycées, EHPAD, ESAT…) : ce sont des méthodes de
  recherche, pas des structures.

### Ce que le script ne fait pas

Les **téléphones et emails** ne figurent dans aucune base ouverte. Ils se collectent sur le
site de l'entreprise (page Contact / Mentions légales), sur sa page LinkedIn, ou par téléphone.
Au standard, demandez toujours : « qui s'occupe des recrutements et de la taxe d'apprentissage ? »

## Fiabilité des données — à lire avant le premier appel

La colonne **« Fiabilité donnée »** vous dit à quoi vous fier :

| Niveau | Ce que ça veut dire |
|---|---|
| **Vérifié INSEE** | Complété par le script depuis la base officielle. Fiable. |
| **Vérifié** | SIRET relevé sur une base publique, à reconfirmer d'un clic. |
| **Source annuaire** | Nom et activité fiables (annuaire fret de l'aéroport, Kompass, Pages Jaunes) ; **adresse à confirmer**. |
| **A vérifier** | Piste non confirmée. Vérifiez l'existence de l'établissement **avant** d'appeler. |

**Aucune donnée n'a été inventée.** Un SIRET ou un nom de dirigeant faux vous ferait perdre
votre crédibilité dès la première phrase. Quand l'information n'était pas vérifiable, la case
porte « A VERIFIER » et la colonne S vous donne un lien pré-rempli vers l'Annuaire des
Entreprises pour la compléter en deux clics.

## Le système d'email — relancer sans tout retaper

Une fois les téléphones et emails collectés (voir ci-dessus), `Modeles_Email_Prospection.xlsx`
contient un modèle prêt à l'emploi pour chaque étape de la prospection : premier contact,
relance sans réponse, confirmation de rendez-vous, envoi d'un profil candidat, remerciement
après signature, etc. — un onglet pour le privé, un pour le secteur public (avec le rappel
de la campagne CNFPT), un pour les candidats de 16 à 25 ans (réponse à candidature, invitation
à une réunion d'information, mise en relation avec une entreprise…).

**Avant le premier envoi — une seule fois :** ouvrez `modeles_email.py` et complétez la
section `CONFIG` en haut du fichier (nom du CFA, votre nom, téléphone, email). Ces
informations remplissent automatiquement la signature de chaque email généré.

**Pour générer des brouillons personnalisés** à partir de votre fichier de prospection :

```
cd <le dossier prospection>
python3 generer_emails_personnalises.py --lister                       # voir tous les modèles
python3 generer_emails_personnalises.py --categorie entreprise_prive \
    --modele PREMIER_CONTACT --statut "A contacter"
```

Le script parcourt le classeur choisi (par défaut le fichier « privé » ; passez le nom du
fichier « public & santé » en premier argument pour l'autre classeur), ne retient que les
lignes où un **email a été renseigné** (colonne « Email contact »), et écrit :

- un fichier **`.eml`** par entreprise dans `emails_generes/<catégorie>/<modèle>/` — un
  double-clic l'ouvre comme brouillon dans Outlook, Mail ou Thunderbird, déjà rempli
  (destinataire, objet, texte), prêt à relire puis à envoyer à la main ;
- **`fusion_emails.csv`** — les mêmes emails en tableau, pour un outil de publipostage
  (Gmail, Outlook) si vous préférez cette méthode ;
- **`journal_emails.csv`** — la liste des lignes ignorées et pourquoi (pas d'email connu,
  statut hors filtre).

**Rien n'est envoyé automatiquement.** Le script ne se connecte à aucune boîte mail : il
prépare des brouillons, vous gardez la main sur chaque envoi.

L'option `--statut` filtre sur la colonne « Statut prospection » (ex. `"A contacter"`,
`"RDV pris"`, ou plusieurs séparés par une virgule) pour n'adresser le bon modèle qu'aux
bonnes lignes — l'onglet « Quand l'utiliser » de `Modeles_Email_Prospection.xlsx` indique
le statut adapté à chaque modèle.

Pour changer le texte d'un modèle : modifiez-le dans `modeles_email.py` (c'est le texte
réellement utilisé par la fusion), puis relancez `generer_modeles_email.py` pour que le
classeur de référence reflète le changement.

### Rappel RGPD / CNIL

La prospection professionnelle par email vers une adresse professionnelle est autorisée en
France sans accord préalable dès lors qu'elle concerne l'activité du destinataire, à condition
d'identifier clairement l'expéditeur et d'offrir un moyen simple de s'y opposer. C'est le rôle
de la mention « Si vous ne souhaitez plus recevoir de proposition de notre part, répondez STOP »
incluse dans la signature des emails entreprises et secteur public — ne la retirez pas.

## La colonne la plus utile : l'effectif

La tranche d'effectif INSEE est votre meilleur filtre de priorisation :

- **0 salarié** → ne pourra pas tutorer un apprenti, passez votre chemin ;
- **1 à 9 salariés** → possible si le dirigeant a le métier et du temps ;
- **10 à 49 salariés** → **la cible idéale** : assez structurée pour tutorer, assez petite
  pour décider vite ;
- **50 salariés et plus** → il faut passer par la RH, cycle plus long mais volumes supérieurs.
