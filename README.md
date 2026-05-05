# Vérificateur de Compteurs — SNDE

Application React pour croiser n'importe quel fichier de prélèvement (Mellah, etc.)
avec le fichier de référence des abonnés actifs (Etat_AbnActif).

## Structure du projet

```
snde-verificateur/
├── index.html
├── package.json
├── vite.config.js
└── src/
    ├── main.jsx
    └── App.jsx        ← code principal
```

## Installation

```bash
# 1. Créer le dossier src/ et y déplacer main.jsx et App.jsx
mkdir src
mv main.jsx App.jsx src/

# 2. Installer les dépendances
npm install

# 3. Lancer en développement
npm run dev

# 4. Construire pour la production
npm run build
```

## Utilisation

1. **Charger le fichier référence** : glissez-déposez `Etat_AbnActif.xlsx`
   (ou tout autre fichier contenant la liste des abonnés actifs)

2. **Charger le fichier à vérifier** : glissez-déposez `Mellah_1_et_2.xlsx`
   (ou tout autre fichier de compteurs prélevés)

3. **Configurer les colonnes** :
   - Sélectionner la feuille si le fichier contient plusieurs feuilles
   - Sélectionner la colonne "numéro de compteur" dans chaque fichier
   - L'application détecte automatiquement la colonne la plus probable

4. **Lancer le croisement** : cliquer sur "Lancer le croisement"

5. **Résultats** :
   - Onglet **Actifs** : compteurs trouvés dans `Etat_AbnActif`, avec toutes les infos abonné
   - Onglet **Inactifs** : compteurs non trouvés (absent de la liste des actifs)
   - Recherche en temps réel dans tous les champs
   - Pagination (50 lignes par page)
   - Statistiques : total, actifs, inactifs, taux d'activation

## Notes techniques

- Utilise la bibliothèque **xlsx** (SheetJS) pour lire les fichiers Excel
- Détection automatique de la ligne d'en-tête (jusqu'à 10 lignes de décalage)
- Fonctionne avec `.xlsx` et `.xls`
- Tout le traitement est fait côté client (navigateur), aucune donnée n'est envoyée à un serveur
- Compatible avec n'importe quel fichier de compteurs, pas seulement Mellah
