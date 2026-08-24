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
`file://` (erreur "Failed to fetch"). Lance `serve.bat` (double-clic, ou
`python -m http.server 8000` à la main), puis ouvre http://localhost:8000/

## Ajouter/mettre à jour les indicateurs manuels (appels, visite médicale)

Ces deux KPIs viennent de la feuille **"Indicateurs Manuels"** dans
`suivi.xlsx` — pas du code. Structure (même logique que Feuil3) :

| Indicateur              | S30 | S31 | S32 | ... |
|--------------------------|-----|-----|-----|-----|
| Appels telephoniques    | 12  | 8   | 15  | ... |
| Visite medicale         | 4   | 6   | 3   | ... |

Remplis la colonne de la semaine en cours (S32, etc.) chaque semaine ;
le dashboard affiche automatiquement la valeur de la semaine courante.
Une case vide affiche "Non renseigné". Si cette feuille est absente du
fichier (ancien `suivi.xlsx`), ces deux KPIs affichent juste "Non renseigné"
sans faire planter le reste.

 Si tu modifies `suivi.xlsx` avec un script Python (openpyxl) plutôt
qu'à la main dans Excel : ouvrir un classeur sans `data_only=True` puis le
resauvegarder **efface les valeurs mises en cache des cellules-formules**
(Réel/Écart de Feuil3 notamment). Édite plutôt directement dans Excel/LibreOffice,
ou si tu dois scripter, resauvegarde uniquement les cellules que tu modifies
et vérifie ensuite avec `data_only=True` que rien d'autre n'a été perdu.
