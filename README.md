# Suivi Usine — version statique (GitHub Pages)

Aucun serveur, aucune base de données : une page HTML qui lit `suivi.xlsx`
directement dans le navigateur (via SheetJS) et calcule les mêmes KPIs que
la version Flask.

## Déployer sur GitHub Pages

1. Crée un dépôt GitHub et pousse ce dossier tel quel :
   ```
   cd dashboard-static
   git init
   git add .
   git commit -m "Suivi Usine - dashboard statique"
   git branch -M main
   git remote add origin https://github.com/<ton-compte>/<ton-repo>.git
   git push -u origin main
   ```
2. Sur GitHub : **Settings → Pages → Source : Deploy from a branch → Branch : main / (root)**.
3. L'URL sera `https://<ton-compte>.github.io/<ton-repo>/` (quelques minutes
   pour la première publication).

⚠️ **Le dépôt sera public par défaut sur un compte GitHub gratuit** (donc
l'URL et son contenu — y compris `suivi.xlsx` — sont visibles par n'importe
qui ayant le lien). Si les données sont sensibles, utilise un compte
GitHub Pro/Team pour un dépôt privé avec Pages, ou repasse sur la version
Flask (mot de passe) hébergée ailleurs.

## Mettre à jour les données

Remplace `suivi.xlsx` par la nouvelle version et pousse :
```
git add suivi.xlsx
git commit -m "Mise à jour suivi.xlsx"
git push
```
GitHub Pages republie automatiquement en 1-2 minutes.

## Tester en local avant de pousser

Ouvrir `index.html` directement (double-clic) ne fonctionnera pas — les
navigateurs bloquent la lecture de fichiers locaux (`fetch`) depuis
`file://`. Sers le dossier avec un petit serveur local :
```
python -m http.server 8000
```
puis ouvre http://localhost:8000/

## Fichiers

```
index.html         La page (structure)
css/style.css       Mêmes couleurs que les autres versions
js/dashboard.js     Parsing suivi.xlsx + calcul des KPIs + graphiques (tout en JS)
js/xlsx.full.min.js SheetJS — lecture du fichier Excel dans le navigateur
js/echarts.min.js   Graphiques
suivi.xlsx           Les données — à remplacer pour mettre à jour
```

## Différences avec la version Flask (Downloads/dashboard)

- Pas de mot de passe, pas de connexion — si le dépôt est public, la page l'est aussi.
- Pas de bouton "Importer" — la mise à jour se fait via `git push`.
- Pas de "Renseigner la semaine prochaine" — rien n'est persisté côté
  serveur puisqu'il n'y a pas de serveur ; la semaine prochaine reste
  "Non renseigné" tant qu'elle n'est pas dans le fichier Excel.
- Filtres département/dates : identiques, calculés à la volée dans le navigateur.
