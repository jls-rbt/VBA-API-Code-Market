# VBA-API-Code-Market

Code VBA pour interroger l'API Alpha Vantage et récupérer le dernier cours de l'action NVDA. Le module `NVDA_API.bas` inclut une extraction plus robuste du champ `05. price`, compatible avec les paramètres régionaux où la virgule est le séparateur décimal.

## Utilisation
1. Importez `NVDA_API.bas` dans votre projet VBA (Alt+F11 > `Fichier` > `Importer un fichier`).
2. Remplacez `VOTRE_CLE_API` par votre clé Alpha Vantage.
3. Exécutez la macro `RecupCoursNVDA` (Alt+F8) : la valeur s'écrit en cellule `C3` et la date/heure de récupération en `D3` de la feuille active.

## Notes
- La fonction `ParsePriceFromJson` remplace le point par le séparateur décimal de votre environnement avant la conversion en nombre pour éviter l'erreur `Type mismatch` sur les systèmes configurés avec une virgule.
