# HOW TO USE THE PROJECT

* Mets toi à la racine du projet avec un terminal (powershell, cmd etc... je conseil powershell) et tape: composer install (à faire qu'une fois, ca va installer tout ce qu'à besoin l'application)
* Ensuite pour lancer une commande dans l'app il faut se mettre à la raine du projet toujours avec un terminal et faire: php bin/console nom_de_la_commande
* Le nom de la commande se situe dans le fichier qui est dans src/Command/BourseSpreadsheetCommand.php
à la ligne 16 et qui est: 'app:bourse-spreadsheet' (ne pas mettre les guillemets)
* La commande sera donc: php bin/console app:bourse-spreadsheet
* Les fichiers xlxs seront dans le dossier public/spreadsheets/  
* Toutes les erreurs de requetes (les scraping qui ne sont pas passés) sont dans le fichier public/spreadsheets/errors.txt
    