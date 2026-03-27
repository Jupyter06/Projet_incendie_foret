# 🔥 Simulation de Feu de Forêt — Automate Cellulaire

[cite_start]Ce projet implémente un automate cellulaire modélisant la propagation d'un feu de forêt[cite: 3]. [cite_start]Il a été développé en appliquant strictement les principes de la Programmation Orientée Objet (POO) : Héritage, Encapsulation, Abstraction et Polymorphisme[cite: 31, 32, 33].

## 🌟 Fonctionnalités

* [cite_start]**Modèle Mathématique :** Utilisation du voisinage de Von Neumann (propagation aux 4 points cardinaux)[cite: 15].
* [cite_start]**États des cellules :** Vide, Arbre sain, Arbre en feu, Arbre brûlé[cite: 7, 8, 9, 10].
* **Environnement dynamique :** Intégration de la force et direction du vent, ainsi que de la météo (température et humidité) modifiant les probabilités de propagation.
* [cite_start]**Deux types de forêts :** Une grille standard avec des limites physiques, et une grille "Torique" où les bords opposés sont connectés[cite: 27].
* **Interface Graphique (Web) :** Visualisation en temps réel via Canvas HTML5 avec contrôles interactifs.
* **Pipeline Data / ML :** Outil de génération de données en batch pour entraîner des modèles d'Intelligence Artificielle.

## 📂 Architecture du Projet

Le projet est divisé en trois composants principaux :

1.  **Le Moteur de Simulation (`sim_feu.py`)**
    * Contient toute la logique POO en Python (`Arbre`, `Vent`, `Meteo`, `Milieu`, `ForetTorique`).
    * Peut être exécuté directement dans le terminal pour une démonstration console.

2.  **L'Interface Graphique (`Propagation_feu.html`)**
    * Application web autonome (HTML/CSS/JS).
    * Ouvrez simplement le fichier dans n'importe quel navigateur pour voir le feu se propager visuellement et ajuster les paramètres en direct.

3.  **Collecte de Données (`collecte_donnees.py`)**
    * Script permettant de lancer des centaines de simulations automatisées (Monte Carlo ou balayage systématique).
    * Génère un fichier Excel (`donnees_simulations.xlsx`) structurant les *Features* (vent, météo, densité) et les *Targets* (taux de destruction, durée) prêt pour le Machine Learning.

## 🚀 Comment utiliser ce projet

### Prérequis
Pour la partie Python (notamment la collecte de données), vous aurez besoin de `openpyxl` pour générer les fichiers Excel :
```bash
pip install openpyxl