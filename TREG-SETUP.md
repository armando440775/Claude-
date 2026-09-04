# Mise en route de treg

treg est un proxy unique vers ~2 900 endpoints d'API (SEO, SERP, enrichissement
de personnes et d'entreprises, reseaux sociaux, donnees web) : un seul token,
une seule URL, facturation a l'appel sur un solde prepaye.

Pour la prospection entreprises du bassin Nord-Est de La Reunion, les familles
d'endpoints utiles sont l'enrichissement societe (`*.companies.*`), la recherche
de contacts et d'emails professionnels (`treg.people.email.find`,
`*.people.search`) et la recherche semantique (`exa.*`).

---

## 1. Ce qui est deja fait dans ce depot

Le skill `treg` est installe dans `.claude/skills/treg/SKILL.md` (copie verbatim
de l'amont, voir `.claude/skills/treg/PROVENANCE.md`). Claude Code le charge
automatiquement des que vous lui demandez des donnees externes ou temps reel.

`.gitignore` protege deja `.treg/`, `token.json`, `.secret/` et `.env` : le token
ne doit jamais etre commite.

## 2. Ce qui reste a faire — en local, pas dans une session web

Le domaine `treg.to` est bloque par le proxy reseau des sessions Claude Code sur
le web. Les trois etapes ci-dessous doivent donc etre lancees depuis un terminal
sur votre machine :

```bash
curl -fsSL https://treg.to/install.sh | sh   # 1. le CLI
treg login                                   # 2. connexion (navigateur)
treg mcp install                             # 3. enregistre les outils treg dans l'agent
```

L'ordre compte : l'etape 3 ecrit le token produit par l'etape 2. Redemarrez
Claude Code apres l'etape 3 pour voir apparaitre les outils `catalog_search`,
`catalog_get`, `call`, `balance`, `my_tools`.

Variantes de connexion :

```bash
treg login --email vous@exemple.fr   # code a 6 chiffres par email, sans navigateur
treg login --token <token>           # non interactif (CI, agent) ; sinon TREG_TOKEN
```

L'installateur depose aussi une copie du skill dans le dossier skills personnel
de l'agent (`~/.claude/skills/treg/`). Elle fait doublon avec celle du depot :
supprimez-en une des deux.

## 3. Verification

```bash
treg --version                      # le CLI repond
treg balance                        # le token est valide, affiche le solde prepaye
treg catalog search "company enrichment"
treg catalog get <id-affiche>       # parametres + PRIX avant tout appel
```

Puis, dans Claude Code, verifiez que le skill repond :

> Cherche dans le catalogue treg un endpoint qui enrichit une entreprise a partir
> de son SIREN ou de son site web, et donne-moi son prix avant appel.

La reponse doit citer des ids d'endpoints et un prix. Si Claude ignore treg,
c'est que le skill n'est pas charge : verifiez la presence de
`.claude/skills/treg/SKILL.md` et relancez la session.

## 4. Points de vigilance

- **Toujours lire le prix** (`treg catalog get`) avant `treg call` : les ecarts de
  tarif entre fournisseurs d'une meme capacite vont jusqu'a 200x.
- **HTTP 402** = solde epuise ; rechargez dans le dashboard (Team -> Billing).
- **Ne jamais reessayer un 4xx chez un autre fournisseur** : c'est un probleme de
  parametres, et l'appel serait facture a chaque tentative.
- Le token vaut de l'argent reel : jamais dans un commit, jamais dans un ticket.
