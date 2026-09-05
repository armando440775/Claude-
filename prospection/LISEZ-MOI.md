# Prospection logistique — bassin Nord-Est de La Réunion

De Saint-Denis à Sainte-Rose, hauteurs comprises. Cible : entreprises susceptibles de
signer un **contrat d'apprentissage** sur un métier de la logistique (16-25 ans).

## Ce que contient ce dossier

| Fichier | À quoi ça sert |
|---|---|
| `Prospection_Logistique_Nord-Est_Reunion.xlsx` | **Votre fichier de travail.** 60 entreprises, 12 onglets. Double-cliquez, c'est tout. |
| `completer_via_annuaire.py` | Remplit tout seul SIRET / code APE / effectif / dirigeant depuis l'INSEE. |
| `generer_fichier_prospection.py` | Régénère le classeur à zéro (à n'utiliser que si vous voulez repartir du modèle). |

⚠️ `generer_fichier_prospection.py` **écrase** le classeur et donc votre suivi de prospection.
Ne le lancez pas une fois que vous avez commencé à remplir les colonnes « Statut » et « Notes ».

## Les onglets

- **00 Mode d'emploi** — à lire en premier.
- **01 à 07** — un onglet par commune, dans l'ordre de la route littorale : Saint-Denis,
  Sainte-Marie, Sainte-Suzanne, Saint-André, Bras-Panon, Saint-Benoît, puis Sainte-Rose
  et les hauteurs. Vous pouvez faire une tournée terrain onglet par onglet.
- **08 Pilotage** — compteurs automatiques (à contacter / RDV / contrats signés) par commune.
- **09 Codes APE à cibler** — pour élargir vous-même le fichier depuis l'Annuaire des Entreprises.
- **10 Script d'appel** — trame d'appel à froid, 5 arguments employeur, réponses aux objections.
- **11 Sources & mise à jour** — d'où viennent les données.

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
python3 completer_via_annuaire.py          # Mac
py completer_via_annuaire.py               # Windows
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

## La colonne la plus utile : l'effectif

La tranche d'effectif INSEE est votre meilleur filtre de priorisation :

- **0 salarié** → ne pourra pas tutorer un apprenti, passez votre chemin ;
- **1 à 9 salariés** → possible si le dirigeant a le métier et du temps ;
- **10 à 49 salariés** → **la cible idéale** : assez structurée pour tutorer, assez petite
  pour décider vite ;
- **50 salariés et plus** → il faut passer par la RH, cycle plus long mais volumes supérieurs.
