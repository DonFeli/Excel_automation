# Automatisation Excel avec openpyxl
## Tests manuels d'un programme de monitoring

Dans le  cadre de ma mission, je dois créer un outil de montoring de scrapping des documents administratifs.

### Sources
Les données sources sont des fichiers json récupérés via l'API de la base
de données *Pensieve*. Ceux-ci sont transformés par le programme pour en tirer des indicateurs 
puis un tableau final est transmis dans une google sheet à une personne tierce qui effectue une vérification sur les sites internet identifiés. 

### Objectif
La phase actuelle du projet est d'implémenter des tests unitaires qui requièrent de contrôler l'input et l'output des différentes fonctions. 
Avec ce projet, mon but est de les visualiser sous Excel plutôt que sur la console, facilitant les tests manuels.

### Structure
Les opérations en amont ne pouvant être effectuées ou visualisées avec Excel sont dans leur propre module,
tandis que les interactions avec Excel sont rassemblées dans la fonction `main` :



### CONSIGNES :

A partir d’un jeu de données que vous aurez sélectionné (format à votre discrétion), vous devrez développer un script R ou Python pour créer un ou des documents Excel.

Votre projet doit comporter au moins 4 des 6 éléments suivants :

1. Création de multiples fichiers Excel à partir d’une source de données unique
2. Transformation des données sources
3. Création de formules Excel.
4. Création d’un ou plusieurs tableaux de bord.
5. Modification de la mise en forme du fichier Excel
6. Insertion de graphiques ou d’images dans un fichier Excel.

Votre projet doit avoir du sens, vous devrez expliquer votre démarche dans un document séparé (sauf si vous utilisez un jupyter notebook, vous pouvez l’insérer directement dans le notebook).

Les documents à rendre sont : un fichier ZIP, contenant :
* Vos fichiers de données sources
* Votre script R ou votre Jupyter notebook python
* Une fiche explicative de votre démarche (1 page maximum)

### EVALUATION 
* Le jeu de données est pertinent et la démarche a du sens.
* Le nombre d’aspects contenu dans votre projet
* La lisibilité du code source, le fait qu’il soit commenté de manière compréhensible.
* Votre script fonctionne et produit bien un ou des fichiers Excel tels que décrit dans votre document d’explication de votre démarche.
* Votre script fonctionne sans modifications, notamment des chemins d’accès aux fichiers sources
* BONUS : utilisation de R ou Python pour réaliser des calculs statistiques ou de modélisation que vous exploitez ou présentez dans votre fichier Excel résultant de votre traitement.
* BONUS : le ou les fichiers Excel résultant de votre traitement ont une mise en forme très propre et agréable.