# Automatisation Excel avec openpyxl
## Tests manuels d'un programme de monitoring

Dans le  cadre de ma mission, je dois créer un outil de montoring de scrapping des documents administratifs.

### Sources
Les données sources sont des fichiers json récupérés via l'API de la base
de données. Ceux-ci sont transformés par le programme pour en tirer des indicateurs 
puis un tableau final est transmis dans une google sheet à une personne tierce qui effectue une vérification sur les sites internet identifiés. 
Pour obtenir des résultats, il faut beaucoup de collectes et le volume de données est assez important.

### Objectif
La phase actuelle du projet est d'implémenter des tests unitaires qui requièrent de contrôler l'input et l'output des différentes fonctions. 
Avec ce projet, mon but est de les visualiser sous Excel plutôt que sur la console, facilitant les tests manuels.

### Structure
Les opérations en amont ne pouvant être effectuées ou visualisées avec Excel sont dans leur propre module,
tandis que les interactions avec Excel sont rassemblées dans la fonction `main` 

LOAD AND TRANSFORM
- Transforme la liste de dictionnaires en dataframe
- Extrait du nested dictionnary les champs item scraped count et finish reason
- Exporte tel quel dans Excel

FILTER INSUFFICIENT COLLECTS
- Filtre les territoires ayant au moins 2 collectes
- Exporte tel quel dans Excel

PAIRS WITH LAST COLLECTS
- Cree un dataframe contenant uniquement les dernieres collectes de chaque territoire
- Merge avec l'ancien dataframe pour obtenir pour chaque collecte, la derniere collecte du territoire
- 2 nouvelles colonnes : item scraped count last et updated at last
- Exporte tel quel dans Excel

PROCESSED PAIRS WITH EXCEL
- Cree 3 nouvelles colonnes dans Excel
    - interval : nombre de jours entre une collecte et la derniere collecte d'un territoire
    - is stopped : si item scraped count est égal à 0
    - was active : si item scraped count est supérieur à 9

PROCESSED PAIRS WITH PYTHON
- La meme chose mais transformé dans Python

SUM SCRAPED ON CURRENTLY STOPPED COLLECTS 
- Cree une colonne calculant la somme cumulée des item scraped count par territoire
- Exporte tel quel dans Excel

EVER ACTIVE EXCEL
- Filtre les territoires pour lesquels il y a eu au moins une collecte active (was_active = TRUE)

EVER ACTIVE PYTHON
- Filtre les territoires pour lesquels il y a eu au moins une collecte active (was_active = TRUE)

GET CANDIDATE EXCEL
- Les collectes qui repondent aux criteres.

GET CANDIDATE PYTHON



