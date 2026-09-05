# Alternance 974 — mise en relation offre / demande

Application de mise en relation entre les **jeunes de 16 à 25 ans** que vous accompagnez et les
**employeurs du bassin Nord-Est** (Saint-Denis → Sainte-Rose, hauteurs comprises).

Un seul fichier, aucune installation, aucun compte, aucun abonnement.

## Démarrer en 30 secondes

1. Ouvrez le dossier `application`.
2. **Double-cliquez sur `Alternance_974.html`.**

Ça marche sur **Mac** comme sur **PC**, dans Safari, Chrome, Edge ou Firefox. Il n'y a rien à
installer et vous n'avez pas besoin d'internet : l'application fonctionne hors ligne.

Astuce : une fois ouverte, mettez la page en favori, ou glissez le fichier sur votre Dock (Mac) /
épinglez-le à la barre des tâches (Windows).

## Ce qu'elle contient déjà

Les **91 structures** de vos deux fichiers de prospection sont pré-chargées : entreprises de la
logistique et du transport, grande distribution, mais aussi communes, lycées, CHU, cliniques et
EHPAD. Vous n'avez rien à ressaisir.

Il ne vous reste qu'à **saisir vos candidats** : dès le premier enregistré, l'application vous
propose les employeurs les plus pertinents pour lui.

## Les six écrans

| Écran | À quoi il sert |
|---|---|
| **Tableau de bord** | Ce qui est ouvert, qui est à placer, et les rapprochements les plus évidents à traiter aujourd'hui. |
| **Offres & structures** | Les 91 employeurs, filtrables par commune, métier, secteur et statut. Une ligne = une fiche + les candidats compatibles. |
| **Candidats** | Vos jeunes : âge, commune, mobilité, métiers recherchés, diplôme préparé, disponibilité. |
| **Mise en relation** | Le cœur de l'outil. Un candidat → ses offres classées, ou une structure → vos candidats classés. |
| **Suivi des placements** | Le pipeline : Proposé → CV envoyé → Entretien planifié → Entretien réalisé → Contrat signé. |
| **Données & sauvegarde** | Export, restauration, export Excel (CSV), effacement. |

## Comment le score est calculé

Chaque rapprochement est noté sur 100, et **l'application affiche toujours pourquoi** — c'est ce
qui vous permet de le défendre devant un employeur ou devant le jeune.

| Critère | Points | Détail |
|---|---|---|
| Métier | 0 à 50 | 50 si le métier de l'offre fait partie de ceux que le jeune recherche, 30 pour la même famille (commerce ou logistique), 18 pour un métier passerelle. |
| Trajet | 0 à 30 | Distance le long de l'axe littoral Saint-Denis → Sainte-Rose, **pondérée par la mobilité du jeune**. Monter ou descendre des hauteurs compte comme un trajet supplémentaire. |
| Avancement | +6 à +12 | Une offre ferme déjà déposée passe devant une structure jamais contactée. |
| Priorité | +3 à +6 | Reprend la priorité P1/P2 de vos fichiers de prospection. |
| Alertes | −20 / −25 | Trajet impossible au quotidien, candidat hors 16-25 ans, ou mise en relation déjà créée. |

La mobilité est volontairement le deuxième critère le plus lourd : à La Réunion, un jeune de
Sainte-Rose sans véhicule ne fera pas Saint-Denis tous les matins, quel que soit l'intérêt du poste.
L'application le dit explicitement au lieu de proposer un placement qui échouera en trois semaines.

## Sauvegarde et données personnelles

**Tout reste sur votre ordinateur**, dans le navigateur avec lequel vous ouvrez le fichier. Rien
n'est envoyé sur internet, il n'y a ni serveur ni cloud.

Trois conséquences pratiques :

- **Exportez votre sauvegarde une fois par semaine** (écran « Données & sauvegarde » → *Exporter la
  sauvegarde*). C'est un fichier `.json` unique.
- **L'application n'est pas partagée entre collègues.** Si un collègue doit reprendre le suivi,
  envoyez-lui votre fichier de sauvegarde : il le charge avec *Restaurer une sauvegarde*.
- Si vous videz les données de navigation de votre navigateur, **vous perdez la base** si vous
  n'avez pas exporté.

**RGPD** : les fiches candidats contiennent des données personnelles de mineurs et de jeunes
majeurs. Ne partagez l'export qu'avec les personnes habilitées de votre CFA, et supprimez les
fiches devenues inutiles (bouton *Tout effacer*, ou suppression fiche par fiche).

## Mettre à jour la liste des structures

Les structures pré-chargées viennent des deux fichiers Excel du dossier `prospection`. Si vous les
enrichissez et voulez les réinjecter dans l'application :

```
cd application
python3 generer_application.py        # Mac
py generer_application.py             # Windows
```

Le script relit les deux classeurs et régénère `Alternance_974.html`. Vos candidats et votre suivi
ne sont pas touchés : ils vivent dans le navigateur, pas dans le fichier. Après régénération,
l'écran « Données » vous propose de *réintégrer* les structures nouvellement ajoutées.

## Fichiers du dossier

| Fichier | Rôle |
|---|---|
| `Alternance_974.html` | **L'application.** C'est le seul fichier dont vous avez besoin au quotidien. |
| `modele_application.html` | Le modèle sans données, utilisé par le générateur. |
| `generer_application.py` | Injecte les structures des fichiers de prospection dans le modèle. |
| `apercu-*.png` | Captures d'écran. |
