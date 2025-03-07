# io-elements

Un outil de gestion d'éléments électrotechniques et électromécaniques.

# Logiciels requis

io-elements est une application Node.js. Plusieurs outils sont nécessaires pour faire fonctionner l'application :

* __git__ : gestionnaire de version disponible [ici pour Windows](https://git-scm.com/download/win "Télécharger git pour Windows")
* __Node.js__ : Le moteur d'exécution javascript basé sur le moteur Chrome V8, disponible [ici pour Windows](https://nodejs.org/en/ "Télécharger Node.js pour Windows"), choisir de préférence la version lts
* __npm__ : normalement installé avec Node.js, node.js est un moteur modulaire, npm est le gestionnaire de modules Node.js.

Tous ces logiciels sont légers, ne modifient pas l'installation de Windows.

## Installation de git

Ouvrir l'installeur Windows et laisser toutes les options d'installation par défaut, c'est suffisant : cliquer sur Ok tout le temps jusqu'à ce que l'installation démarre.

L'installeur git ne va pas chercher à installer des logiciels tiers.

## Installation de node.js

Lors de l'installation, l'installeur va demander d'installer les outils de compilation. Ce n'est pas nécessaire pour ce projet.

npm est installé en même temps que node.js.

# Installation de l'application sur un poste Windows

* Ouvrir une fenêtre du powershell, en mode utilisateur normal, le mode administrateur n'est pas nécessaire.
* se placer dans le répertoire qui contiendra le projet
* cloner le dépôt git : `git clone https://github.com/jledun/io-elements.git`
* aller dans le répertoire cloné : `cd io-elements`
* installer les dépendances : `npm install`
* C'est fait :-)

# Configuration des fichiers à ne pas enregistrer sur le dépot

Editer le fichier `\chemin\vers\le\projet\.gitignore` et ajouter à la fin du fichier :

```
*.xlsx
*.mdb
*.ldb
node_modules
files
```

Ainsi, les fichiers Excel et Access ne seront jamais indexés dans le dépot git.

# Mise à jour du code suite à une mise à jour du dépot git

## Savoir si le code local a été modifié ou non

* Ouvrir une fenêtre du powershell, en mode utilisateur normal, le mode administrateur n'est pas nécessaire.
* se placer dans le répertoire qui contient le projet
* exécuter la commande `git status`
* git affiche la liste des fichiers qui ont été modifiés ou en attente d'ajout à l'index ou rien si rien n'a été modifié.

## Le code local n'a pas été modifié

* Ouvrir une fenêtre du powershell, en mode utilisateur normal, le mode administrateur n'est pas nécessaire.
* se placer dans le répertoire qui contient le projet
* mettre à jour le code depuis le dépot git : `git pull`
* c'est fait :-)

## Le code local a été modifié

Ce document n'est pas un tuto pour utiliser git, il y en a plein librement disponibles sur Internet.

# Comment ça marche ?

Editer la liste des éléments dans la feuille de calcul 'Elements' incluse dans le fichier 'Elements.xlsx' : repartir d'une liste vide ou éditer une liste existante.

Quand la liste est complète, il faut mettre à jour la liste des éléments dans la feuille de calcul 'Asservissements' : cette application s'en charge pour vous :-)

Dans le powershell, se placer dans le répertoire contenant l'application ET le fichier 'Elements.xlsx' et entrer la commande `node io-engine.js`.

Accepter la licence et sélectionner le premier choix : "Créer ou mettre à jour la feuille 'Asservissements' et initialiser la feuille 'Cycles' dans le fichier 'Elements.xlsx'".

Après quelques secondes, l'opération se termine, vous pouvez compléter la feuille de calcul 'Asservissements' en saisissant les éléments en aval.

Une fois la feuille de calcul 'Asservissements' complète, dans le powershell, se placer dans le répertoire contenant l'application ET le fichier 'Elements.xlsx' et entrer la commande `node io-engine.js`.

Accepter la licence et sélectionner le deuxième choix : "Générer les chemins dans la feuille 'Cycles' de 'Elements.xlsx' et mettre à jour la base de données 'Cycles.mdb'".

La partie la plus longue est l'enregistrement du résultat dans la base de données, merci Windows.

Mais c'est tout, rien à faire de plus, passer à l'installation suivante :-)

# Que faire en cas de modification de process ?

Editer la liste des éléments dans la feuille de calcul 'Elements' incluse dans le fichier 'Elements.xlsx' pour ajouter ou retirer des éléments dans la liste.

Exécuter io-engine avec le premier choix pour mettre à jour les intitulés des lignes et colonnes de la feuille de calcul 'Asservissements', les données présentes à l'intérieur du tableau ne sont pas modifiées.

Mettre à jour les asservissements entre éléments en fonction des modifications de process intégrées.

Exécuter io-engine avec le deuxième choix et récupérez le nouveau résultat : le nouveau résultat écrase toujours le résultat précédent. Le résultat précédent est perdu, pensez à sauvegarder :-)

# Où sont les fichiers Excel ?

Formuler votre demande par email à [j.ledun@iosystems.fr](mailto:j.ledun@iosystems.fr)

# Licence MIT

Copyright 2025 Julien Ledun <j.ledun@iosystems.fr>

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

# Contributeurs

* Julien Ledun <j.ledun@iosystems.fr>

# TODO

* améliorer tout ça, l'interface, l'ergonomie, l'automatisation...
