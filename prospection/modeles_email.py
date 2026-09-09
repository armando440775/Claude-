# -*- coding: utf-8 -*-
"""
Modeles d'email pour la prospection apprentissage du bassin Nord-Est.

Source unique de verite, utilisee par :
    - generer_modeles_email.py          (classeur de reference, lisible et imprimable)
    - generer_emails_personnalises.py   (fusion avec vos fichiers de prospection Excel)

A FAIRE UNE SEULE FOIS avant le premier envoi : completez la section CONFIG
ci-dessous avec vos coordonnees. Elle alimente automatiquement la signature
et l'objet de chaque email genere.
"""

# ---------------------------------------------------------------- CONFIG
# Vos informations : a personnaliser une bonne fois pour toutes.
CONFIG = dict(
    cfa_nom="Votre CFA",
    cfa_secteur="les métiers du commerce",
    cfa_zone="le bassin Nord-Est de La Réunion, de Saint-Denis à Sainte-Rose (hauteurs comprises)",
    mon_prenom_nom="Prénom NOM",
    ma_fonction="Responsable commercial et chargé(e) de Relations Entreprise",
    mon_tel="0262 00 00 00",
    mon_email="prenom.nom@votre-cfa.re",
    cfa_site="www.votre-cfa.re",
)

AV = "A VERIFIER"
AC = "A COMPLETER"
PLACEHOLDERS = {AV, AC, "", None, "A VÉRIFIER", "A COMPLÉTER"}

# ---------------------------------------------------------------- Signatures
# Bloc ajoute en fin de message. La mention "STOP" repond a l'obligation
# d'offrir un moyen simple de s'opposer a la prospection par email (RGPD/CNIL) ;
# gardez-la sur les emails entreprises et secteur public.
SIGNATURE_ENTREPRISE = """Bien cordialement,

{mon_prenom_nom}
{ma_fonction} — {cfa_nom}
{cfa_zone}
Tél. {mon_tel} | {mon_email} | {cfa_site}

Si vous ne souhaitez plus recevoir de proposition de notre part, répondez « STOP » à ce message."""

# Les candidats sont deja en contact avec vous (candidature, inscription) :
# pas de mention STOP necessaire.
SIGNATURE_CANDIDAT = """Bien cordialement,

{mon_prenom_nom}
{ma_fonction} — {cfa_nom}
Tél. {mon_tel} | {mon_email} | {cfa_site}"""


# ---------------------------------------------------------------- Modeles
# Chaque modele : nom (affiche dans le classeur de reference), quand (a quel
# moment / sur quel statut de prospection l'utiliser), sujet, corps.
# Variables disponibles : voir variables_disponibles() plus bas.
TEMPLATES = {

    # ============================================================ ENTREPRISES PRIVEES
    "entreprise_prive": {

        "PREMIER_CONTACT": dict(
            nom="1. Premier contact",
            quand="Statut « A contacter ». Envoi avant ou juste après le premier appel.",
            sujet="Recruter {poste} en alternance, sans frais de recrutement — {cfa_nom}",
            corps="""{salutation}

Je suis {ma_fonction} au sein de {cfa_nom}, centre de formation par apprentissage sur \
{cfa_zone}. Nous formons {cfa_secteur} et accompagnons chaque année des jeunes de 16 à \
25 ans vers un contrat d'apprentissage.

{entreprise} m'a semblé être une structure susceptible d'accueillir {poste} en \
alternance {commune_mention}. Le contrat d'apprentissage vous permet de former un jeune \
au poste tel que vous le pratiquez, sans frais de recrutement, avec une prise en charge \
de la formation par votre OPCO et des aides à l'embauche pouvant s'appliquer selon le \
profil du candidat.

Auriez-vous 15 minutes cette semaine ou la semaine prochaine pour en discuter ? Je peux \
aussi me déplacer si cela vous convient mieux.

{signature}""",
        ),

        "RELANCE_SANS_REPONSE": dict(
            nom="2. Relance sans réponse (J+7)",
            quand="Statut « A contacter » ou « Message laissé », une semaine sans retour.",
            sujet="Petite relance — apprenti(e) {poste} chez {entreprise}",
            corps="""{salutation}

Je me permets de revenir vers vous : je n'ai pas eu l'occasion d'échanger avec vous sur \
l'accueil {commune_mention} d'un(e) apprenti(e) {poste}.

Si le sujet vous intéresse, un simple appel de 10 minutes suffit pour voir si le profil \
et le calendrier vous conviennent. Sinon, dites-le-moi et je ne vous solliciterai plus \
sur ce sujet.

{signature}""",
        ),

        "CONFIRMATION_RDV": dict(
            nom="3. Confirmation de rendez-vous",
            quand="Après un accord téléphonique pour un RDV ou une visite.",
            sujet="Confirmation de notre rendez-vous — {cfa_nom} / {entreprise}",
            corps="""{salutation}

Je vous confirme notre rendez-vous du {date_rdv}{lieu_rdv_mention} au sujet de l'accueil \
{commune_mention} d'un(e) apprenti(e) {poste}.

Je viendrai avec une présentation rapide du contrat d'apprentissage (coût, démarches, \
calendrier) et, si nous en sommes déjà là, un ou deux profils de candidats correspondant \
au poste.

N'hésitez pas si vous souhaitez que j'associe une autre personne de {entreprise} à cet \
échange.

{signature}""",
        ),

        "APRES_RDV_ENVOI_CANDIDAT": dict(
            nom="4. Après RDV — envoi d'un profil candidat",
            quand="Statut « Visite faite », pour transmettre un ou plusieurs CV.",
            sujet="Suite à notre échange — profil(s) candidat pour {poste}",
            corps="""{salutation}

Merci pour le temps accordé lors de notre échange sur l'accueil {commune_mention} d'un(e) \
apprenti(e) {poste}.

Comme convenu, vous trouverez ci-joint le(s) profil(s) de {candidat_prenom} qui \
correspond(ent) à ce que vous recherchez. N'hésitez pas à me dire si vous souhaitez \
organiser un entretien ou une période d'immersion avant de vous engager.

{signature}""",
        ),

        "RELANCE_CANDIDAT_ENVOYE": dict(
            nom="5. Relance après envoi d'un profil",
            quand="Statut « Offre déposée » ou « Candidat proposé », sans retour depuis quelques jours.",
            sujet="Retour sur le profil transmis pour {poste} ?",
            corps="""{salutation}

Je reviens vers vous au sujet du profil de {candidat_prenom} transmis pour le poste \
d'apprenti(e) {poste}. Avez-vous pu l'examiner ? Je reste disponible pour organiser un \
entretien, ou pour vous proposer un autre profil si celui-ci ne correspond pas tout à \
fait à ce que vous recherchez.

{signature}""",
        ),

        "REMERCIEMENT_CONTRAT_SIGNE": dict(
            nom="6. Remerciement — contrat signé",
            quand="Statut « Contrat signé ». A envoyer dès la signature.",
            sujet="Merci pour votre confiance — {entreprise}",
            corps="""{salutation}

Merci d'avoir fait confiance à {cfa_nom} pour l'accueil de {candidat_prenom} en \
alternance sur le poste {poste}. C'est ce type de collaboration qui fait vivre \
l'apprentissage {commune_mention}.

Je reste votre interlocuteur(trice) tout au long du contrat : n'hésitez pas à me \
solliciter pour toute question de suivi, de planning école/entreprise ou \
administrative. Je vous recontacterai également avant l'échéance du contrat si vous \
souhaitez renouveler l'expérience.

{signature}""",
        ),

        "NURTURING_A_RAPPELER": dict(
            nom="7. Pas maintenant, à rappeler plus tard",
            quand="Statut « Sans suite » quand le refus est lié au calendrier (déjà pourvu, saison creuse).",
            sujet="On se recontacte pour la prochaine rentrée — {entreprise}",
            corps="""{salutation}

Je note que le moment n'est pas idéal pour accueillir {poste} en alternance chez \
{entreprise}. Je garde votre structure en tête et reviendrai vers vous en amont de la \
prochaine rentrée, sauf si vous préférez que je ne vous recontacte pas.

Si votre besoin évolue avant, n'hésitez pas à me solliciter directement.

{signature}""",
        ),
    },

    # ============================================================ SECTEUR PUBLIC
    "secteur_public": {

        "PREMIER_CONTACT_PUBLIC": dict(
            nom="1. Premier contact collectivité / établissement public",
            quand="Statut « A contacter ». Adapter le ton, moins commercial que le privé.",
            sujet="Accueillir un apprenti {poste} — {cfa_nom}",
            corps="""{salutation}

Je suis {ma_fonction} au sein de {cfa_nom}, centre de formation par apprentissage sur \
{cfa_zone}. Nous accompagnons des jeunes de 16 à 25 ans vers un contrat \
d'apprentissage, y compris dans la fonction publique territoriale.

{entreprise} pourrait être en mesure d'accueillir {poste} en apprentissage \
{commune_mention}. Je souhaiterais échanger avec vous sur ce sujet et, le cas échéant, \
sur le recensement annuel des besoins auprès du CNFPT qui conditionne la prise en \
charge des frais de formation.

Auriez-vous un créneau dans les prochaines semaines pour un échange téléphonique ou un \
rendez-vous ?

{signature}""",
        ),

        "RELANCE_RECENSEMENT_CNFPT": dict(
            nom="2. Rappel campagne de recensement CNFPT",
            quand="A envoyer en fin d'année civile, avant l'ouverture de la campagne CNFPT de janvier.",
            sujet="Ne manquez pas le recensement CNFPT pour financer votre apprenti {poste}",
            corps="""{salutation}

Un point important pour financer l'accueil {commune_mention} d'un(e) apprenti(e) \
{poste} : le CNFPT ne prend en charge les frais de formation que si {entreprise} \
s'est recensée pendant sa campagne annuelle, ouverte en tout début d'année civile.

Je vous propose d'échanger dès maintenant pour préparer ce recensement et vérifier que \
le métier visé figure bien sur la liste des métiers en tension du CNFPT. Un rendez-vous \
pris cette année ne sert pas encore à signer, mais à sécuriser le financement de \
l'année suivante.

{signature}""",
        ),

        "APRES_RDV_PUBLIC": dict(
            nom="3. Après rendez-vous (public)",
            quand="Statut « Visite faite ». Récapitule les prochaines étapes administratives.",
            sujet="Suite à notre échange — apprentissage {poste} à {entreprise}",
            corps="""{salutation}

Merci pour cet échange sur l'accueil {commune_mention} d'un(e) apprenti(e) {poste}. \
Comme convenu, je reviens vers vous avec les éléments nécessaires pour avancer sur ce \
dossier : vérification du métier sur la liste des métiers en tension du CNFPT, \
inscription au prochain recensement si ce n'est pas déjà fait, et, si vous le \
souhaitez, un ou deux profils de candidats.

{signature}""",
        ),
    },

    # ============================================================ CANDIDATS (16-25 ANS)
    "candidats": {

        "REPONSE_CANDIDATURE": dict(
            nom="1. Réponse à une candidature reçue",
            quand="Dès réception d'une candidature spontanée ou d'un formulaire.",
            sujet="Votre candidature chez {cfa_nom} — prochaines étapes",
            corps="""Bonjour {candidat_prenom},

Merci pour l'intérêt que vous portez à {cfa_nom} et à la formation en alternance sur \
{formation}.

Pour avancer sur votre dossier, je vous propose un premier échange (téléphonique ou sur \
place) afin de faire le point sur votre projet, votre disponibilité et les entreprises \
qui recrutent actuellement en apprentissage {cfa_zone}.

Dites-moi vos disponibilités dans les prochains jours et je reviens vers vous rapidement.

{signature}""",
        ),

        "INVITATION_INFO_COLLECTIVE": dict(
            nom="2. Invitation réunion d'information",
            quand="Pour convier un candidat repéré à une session collective ou portes ouvertes.",
            sujet="Invitation — réunion d'information {cfa_nom} le {date_info}",
            corps="""Bonjour {candidat_prenom},

{cfa_nom} organise une réunion d'information sur {formation} le {date_info}. Ce sera \
l'occasion de présenter le déroulé de la formation en alternance, les entreprises \
partenaires {cfa_zone} et les démarches pour signer un contrat d'apprentissage.

Merci de me confirmer votre présence par retour de mail afin que je vous réserve une \
place.

{signature}""",
        ),

        "CONFIRMATION_INSCRIPTION": dict(
            nom="3. Confirmation d'inscription",
            quand="Une fois le dossier d'inscription du candidat validé.",
            sujet="Confirmation de votre inscription — {formation}",
            corps="""Bonjour {candidat_prenom},

Votre inscription à la formation {formation} chez {cfa_nom} est confirmée. Prochaine \
étape : trouver l'entreprise qui vous accueillera en apprentissage. Je vous recontacte \
dès qu'un profil d'entreprise correspondant à votre projet se présente {cfa_zone}, et \
je reste disponible entre-temps si vous avez des pistes de votre côté.

{signature}""",
        ),

        "MISE_EN_RELATION_ENTREPRISE": dict(
            nom="4. Mise en relation avec une entreprise",
            quand="Quand un poste correspond au profil du candidat.",
            sujet="Une opportunité en alternance pour vous — {poste}",
            corps="""Bonjour {candidat_prenom},

{entreprise} recherche {poste} en alternance {commune_mention}, et votre profil \
correspond à ce qu'ils recherchent. Je leur ai transmis votre candidature.

Merci de rester joignable dans les prochains jours : ils pourraient vous contacter \
directement, ou je reviendrai vers vous pour organiser un entretien.

{signature}""",
        ),

        "RELANCE_CANDIDAT_SANS_NOUVELLES": dict(
            nom="5. Relance candidat sans nouvelles",
            quand="Candidat inscrit ou repéré qui ne répond plus depuis un moment.",
            sujet="Toujours partant(e) pour l'alternance ?",
            corps="""Bonjour {candidat_prenom},

Je n'ai plus de nouvelles depuis notre dernier échange au sujet de {formation}. Votre \
projet d'alternance tient-il toujours ? Si oui, dites-le-moi rapidement pour que je \
continue à vous proposer des entreprises {cfa_zone}. Si votre situation a changé, un \
petit mot suffit pour que je clôture votre dossier.

{signature}""",
        ),
    },
}


# ---------------------------------------------------------------- Helpers
def variables_disponibles():
    """Liste les variables utilisables dans un modele, pour l'onglet de reference
    du classeur genere par generer_modeles_email.py."""
    communes = [
        ("{salutation}", "« Bonjour Prénom Nom, » ou « Bonjour, » si le contact est inconnu"),
        ("{entreprise}", "Raison sociale / enseigne (colonne « Raison sociale / Enseigne »)"),
        ("{poste}", "Poste apprenti visé (colonne du même nom)"),
        ("{commune}", "Commune de l'établissement"),
        ("{commune_mention}", "« à Saint-André » automatiquement, ou vide si la commune est inconnue"),
        ("{secteur}", "Segment / type de structure"),
        ("{fonction}", "Fonction du contact chez l'entreprise"),
        ("{date_rdv}", "A completer a la main avant l'envoi (email de confirmation de RDV)"),
        ("{lieu_rdv_mention}", "A completer a la main, ex. « , dans vos locaux »"),
        ("{candidat_prenom}", "Prénom du candidat — emails candidats et emails « envoi de profil »"),
        ("{formation}", "Intitulé de la formation visée — emails candidats"),
        ("{date_info}", "Date de la prochaine réunion d'information — emails candidats"),
        ("{signature}", "Bloc de signature, rempli automatiquement à partir de CONFIG"),
    ]
    config_vars = [("{%s}" % k, "Personnalisation CFA, définie une fois dans CONFIG") for k in CONFIG]
    return communes + config_vars


class _Defaut(dict):
    """Laisse {variable} telle quelle si elle est absente, au lieu de planter :
    un oubli se repere immediatement a la lecture de l'email genere."""
    def __missing__(self, cle):
        return "{%s}" % cle


def personnaliser(gabarit, valeurs):
    """Remplace les {variables} d'un gabarit avec les valeurs fournies, complétées
    par les reglages de CONFIG."""
    fusion = dict(CONFIG)
    fusion.update(valeurs)
    return gabarit.format_map(_Defaut(fusion))


def signature_pour(categorie):
    """Renvoie le bon bloc de signature (avec mention STOP ou non) selon la categorie."""
    gabarit = SIGNATURE_CANDIDAT if categorie == "candidats" else SIGNATURE_ENTREPRISE
    return personnaliser(gabarit, {})
