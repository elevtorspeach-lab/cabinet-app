# Cabinet Avocat Offline Web

Version GitHub-ready de l'application web offline.

## Ce que contient cette copie

- Application web offline uniquement
- Aucun dossier client ni donnees metier inclus
- Aucun export, backup, release ou archive de distribution
- Assets, workers et librairies locales necessaires au fonctionnement

## Confidentialite

Cette copie est vide par defaut.
Les donnees utilisees par l'application sont stockees localement dans le navigateur (LocalStorage / IndexedDB) apres utilisation, et ne sont pas incluses dans ce dossier.

## Lancer en offline

1. Ouvrir `index.html` directement dans le navigateur
2. Ou utiliser `index.html?offline=1`

L'application detecte automatiquement le mode local/offline lorsqu'elle est ouverte depuis un fichier local.

## Connexion initiale

- Utilisateur par defaut: `manager`
- Mot de passe par defaut dans cette base source: `1234`

Pour une publication publique, il est recommande de changer ce mot de passe des le premier usage.

## Fichiers exclus volontairement

- `release/`
- `backups/`
- `exports/`
- `snapshot-*`
- `node_modules/`
- archives `.zip`

## Publication GitHub

Publiez le contenu de ce dossier comme depot source web offline propre.
