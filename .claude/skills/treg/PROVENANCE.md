# Provenance du skill treg

- Source : https://github.com/superdesigndev/treg — `skills/treg/SKILL.md`
- Commit repris : `5e3cd520d8a80e3aae5d5576d28a3dd174ab5d5e` (2026-09-05)
- Version du skill : 0.15.0
- Licence amont : Apache-2.0

`SKILL.md` est une copie **verbatim** de l'amont. Ne pas l'editer a la main :
pour mettre a jour, recopier le fichier depuis le depot amont et mettre a jour
le commit ci-dessus.

## Pourquoi une copie dans le depot plutot que le plugin ?

Le plugin officiel s'installe avec :

    /plugin marketplace add superdesigndev/treg
    /plugin install treg@treg

Le plugin ne contient qu'un seul composant : ce `SKILL.md`. Le copier dans
`.claude/skills/treg/` donne le meme resultat, versionne dans le depot, donc
disponible pour toute personne qui clone le projet et pour les sessions
Claude Code sur le web (ou l'ajout de plugin n'est pas possible).

Si vous preferez le plugin, supprimez ce dossier pour eviter le doublon.
