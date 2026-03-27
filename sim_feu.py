"""
simulation_feu.py — Propagation de Feu : Automate Cellulaire
=============================================================
Interface graphique corrélative : simulation_feu.html (même dossier)

Architecture des classes (OOP) :
  Arbre         — cellule de la grille, machine à états (0-3)
  Vent          — facteur directionnel de propagation (produit scalaire)
  Meteo         — coefficient global humidité × température
  Milieu        — grille + orchestration (équivalent "Forest" de l'exercice)
  ForetTorique  — héritage de Milieu, voisinage torique (bords connectés)

Principes OOP appliqués :
  Encapsulation  — attributs privés (_etat, _grille…), accès via @property
  Héritage       — ForetTorique(Milieu) surcharge _coord_voisins()
  Abstraction    — méthodes publiques haut niveau (simuler, afficher, proportion)
  Polymorphisme  — _coord_voisins() : même signature, comportement différent
"""

import random
import math
from copy import deepcopy
from typing import List, Tuple, Optional


# ================================================================
#  ARBRE — cellule élémentaire de la grille
#  État encodé en entier : 0=vide, 1=sain, 2=en feu, 3=brûlé
# ================================================================

class Arbre:

    # Constantes d'état (encapsulation des valeurs magiques)
    VIDE   = 0
    SAIN   = 1
    FEU    = 2
    BRULEE = 3

    # Symboles pour l'affichage console (méthode requise par l'exercice)
    _SYMBOLES = {0: '.', 1: 'T', 2: 'F', 3: 'B'}

    def __init__(self, etat: int = VIDE) -> None:
        self._etat         = etat
        self._temps_en_feu = 0
        # Extension : durée de combustion variable (réalisme)
        self._duree_max    = random.randint(2, 4)

    # ---- Accesseurs (encapsulation) ----

    @property
    def etat(self) -> int:
        return self._etat

    @property
    def est_vide(self) -> bool:
        return self._etat == Arbre.VIDE

    @property
    def est_sain(self) -> bool:
        return self._etat == Arbre.SAIN

    @property
    def est_en_feu(self) -> bool:
        return self._etat == Arbre.FEU

    @property
    def est_brulee(self) -> bool:
        return self._etat == Arbre.BRULEE

    # ---- Méthode de l'exercice ----
    def symbole(self) -> str:
        """Retourne le symbole console : '.' / 'T' / 'F' / 'B'"""
        return Arbre._SYMBOLES.get(self._etat, '?')

    # ---- Mise à jour déterministe (Part 1) ----
    def update(self, voisin_en_feu: bool) -> None:
        """
        Applique la règle déterministe de l'exercice :
          sain + voisin en feu  → en feu
          en feu                → brûlé (1 étape)
          vide / brûlé          → absorbant (aucun changement)
        """
        if self._etat == Arbre.SAIN and voisin_en_feu:
            self._etat = Arbre.FEU
        elif self._etat == Arbre.FEU:
            self._etat = Arbre.BRULEE

    # ---- Extension : allumage probabiliste ----
    def mettre_en_feu(self) -> bool:
        """Allume l'arbre si sain. Retourne True si l'allumage a eu lieu."""
        if self._etat == Arbre.SAIN:
            self._etat         = Arbre.FEU
            self._temps_en_feu = 0
            return True
        return False

    # ---- Extension : combustion progressive ----
    def evoluer(self) -> None:
        """Fait progresser la combustion sur plusieurs cycles (réalisme)."""
        if self._etat == Arbre.FEU:
            self._temps_en_feu += 1
            if self._temps_en_feu >= self._duree_max:
                self._etat = Arbre.BRULEE

    def __repr__(self) -> str:
        etats = {0: 'VIDE', 1: 'SAIN', 2: 'FEU', 3: 'BRULEE'}
        return f"Arbre({etats.get(self._etat, '?')})"


# ================================================================
#  VENT — facteur directionnel de propagation
#  Le calcul repose sur le produit scalaire entre le vecteur vent
#  et le vecteur de propagation → alignement géométrique exact.
# ================================================================

class Vent:

    # Vecteurs unitaires des 8 directions cardinales et inter-cardinales
    VECTEURS: dict = {
        'N':  ( 0, -1),  'S':  ( 0,  1),
        'E':  ( 1,  0),  'O':  (-1,  0),
        'NE': ( 1, -1),  'NO': (-1, -1),
        'SE': ( 1,  1),  'SO': (-1,  1),
    }

    def __init__(self, direction: str = 'N', vitesse: int = 2) -> None:
        """
        direction : 'N', 'S', 'E', 'O', 'NE', 'NO', 'SE', 'SO'
        vitesse   : 0 (calme) → 3 (tempête)
        """
        if direction not in Vent.VECTEURS:
            raise ValueError(f"Direction invalide : {direction}")
        if not 0 <= vitesse <= 3:
            raise ValueError(f"Vitesse hors plage [0, 3] : {vitesse}")
        self.direction = direction
        self.vitesse   = vitesse

    def facteur_propagation(self, dx: int, dy: int) -> float:
        """
        Retourne un multiplicateur de probabilité pour propager vers (dx, dy).

        Produit scalaire vecteur_vent · vecteur_propagation normalisé :
          - aligné avec le vent  → facteur > 1 (propagation favorisée)
          - perpendiculaire      → facteur ≈ 1 (neutre)
          - contre le vent       → facteur < 1 (propagation freinée)
        """
        if self.vitesse == 0:
            return 1.0
        vx, vy  = Vent.VECTEURS[self.direction]
        dot     = dx * vx + dy * vy
        norme   = math.sqrt(dx**2 + dy**2)          # longueur vecteur propagation
        alignement = dot / norme if norme > 0 else 0 # ∈ [-√2, √2]
        boost   = (self.vitesse / 3) * alignement * 0.85
        return max(0.05, 1.0 + boost)

    def __repr__(self) -> str:
        return f"Vent(direction='{self.direction}', vitesse={self.vitesse})"


# ================================================================
#  MÉTÉO — coefficient global humidité × température
# ================================================================

class Meteo:

    def __init__(self, humidite: float = 40.0, temperature: float = 30.0) -> None:
        """
        humidite    : 0–100 %   — élevée freine la propagation
        temperature : °C        — élevée accélère la propagation
        """
        self.humidite    = humidite
        self.temperature = temperature

    def facteur_propagation(self) -> float:
        """
        Coefficient multiplicatif global :
          f_hum  = 1 − (humidité/100) × 0.75
          f_temp = 1 + (température − 20) / 50
          résultat = f_hum × f_temp  (minimum 0.05)
        """
        f_hum  = 1.0 - (self.humidite / 100.0) * 0.75
        f_temp = 1.0 + (self.temperature - 20.0) / 50.0
        return max(0.05, f_hum * f_temp)

    def __repr__(self) -> str:
        return f"Meteo(humidite={self.humidite}%, temperature={self.temperature}°C)"


# ================================================================
#  MILIEU (= "Forest" de l'exercice)
#  Gère la grille n×m, les règles de propagation, les statistiques.
#  Voisinage de Von Neumann : 4 directions cardinales (N, S, E, O).
# ================================================================

class Milieu:

    # Voisinage de Von Neumann — 4 directions cardinales uniquement
    VOISINS_VN: List[Tuple[int, int]] = [(0, -1), (0, 1), (1, 0), (-1, 0)]

    def __init__(
        self,
        nb_lignes:   int,
        nb_colonnes: int,
        p:           float,
        vent:        Vent,
        meteo:       Meteo,
    ) -> None:
        """
        nb_lignes, nb_colonnes : dimensions de la grille
        p    : probabilité qu'une cellule contienne un arbre sain ∈ [0, 1]
        vent : objet Vent (facteur directionnel)
        meteo: objet Meteo (coefficient global)
        """
        self.nb_lignes   = nb_lignes
        self.nb_colonnes = nb_colonnes
        self.vent        = vent
        self.meteo       = meteo
        self._grille: List[List[Arbre]] = []
        self.etape       = 0
        self.fini        = False
        self._initialiser(p)

    # ---- Initialisation ----

    def _initialiser(self, p: float) -> None:
        """Remplit la grille aléatoirement avec probabilité p pour un arbre sain."""
        self._grille = [
            [
                Arbre(Arbre.SAIN if random.random() < p else Arbre.VIDE)
                for _ in range(self.nb_colonnes)
            ]
            for _ in range(self.nb_lignes)
        ]
        self.etape = 0
        self.fini  = False

    # ---- Allumage du foyer ----

    def demarrer_foyer(self, x: int, y: int) -> None:
        """Allume un foyer à la position (colonne x, ligne y)."""
        cellule = self.get_cellule(x, y)
        if cellule and cellule.est_vide:
            self._grille[y][x] = Arbre(Arbre.SAIN)
        if cellule:
            self._grille[y][x].mettre_en_feu()

    def demarrer_foyer_centre(self) -> None:
        """Allume un foyer au centre de la grille."""
        self.demarrer_foyer(self.nb_colonnes // 2, self.nb_lignes // 2)

    def demarrer_foyer_aleatoire(self) -> None:
        """Allume un foyer sur un arbre sain tiré au hasard."""
        sains = [
            (x, y)
            for y in range(self.nb_lignes)
            for x in range(self.nb_colonnes)
            if self._grille[y][x].est_sain
        ]
        if sains:
            x, y = random.choice(sains)
            self.demarrer_foyer(x, y)

    # ---- Accès à la grille (abstraction) ----

    def get_cellule(self, x: int, y: int) -> Optional[Arbre]:
        """Retourne la cellule en (x, y) ou None si hors grille."""
        if 0 <= y < self.nb_lignes and 0 <= x < self.nb_colonnes:
            return self._grille[y][x]
        return None

    # ---- Coordonnées des voisins (Von Neumann) ----
    # Méthode "protégée" — surchargée dans ForetTorique

    def _coord_voisins(self, x: int, y: int) -> List[Tuple[int, int]]:
        """Retourne les coordonnées valides des 4 voisins Von Neumann de (x, y)."""
        return [
            (x + dx, y + dy)
            for dx, dy in Milieu.VOISINS_VN
            if 0 <= x + dx < self.nb_colonnes and 0 <= y + dy < self.nb_lignes
        ]

    # ---- Méthode requise : voisin en feu ? ----

    def voisin_en_feu(self, x: int, y: int) -> bool:
        """Retourne True si au moins un voisin Von Neumann de (x, y) est en feu."""
        return any(
            self._grille[ny][nx].est_en_feu
            for nx, ny in self._coord_voisins(x, y)
        )

    # ---- Méthode requise : proportion d'arbres sains ----

    def proportion(self) -> float:
        """Retourne la proportion de cases contenant un arbre sain."""
        total = self.nb_lignes * self.nb_colonnes
        return self.nb_sains() / total if total > 0 else 0.0

    # ---- Méthode requise : affichage console ----

    def afficher(self) -> str:
        """
        Retourne la représentation console de la grille avec symboles.
        Exemple de sortie :
            génération 3 :
            . T T . T
            T F T T .
            . T B F T
            proportion d'arbres sains : 0.52
        """
        lignes = [f"génération {self.etape} :"]
        for y in range(self.nb_lignes):
            lignes.append(' '.join(self._grille[y][x].symbole()
                                   for x in range(self.nb_colonnes)))
        lignes.append(f"proportion d'arbres sains : {self.proportion():.2f}")
        return '\n'.join(lignes)

    # ---- Méthode requise : simuler n générations ----

    def simuler(self, n: int) -> None:
        """Génère exactement n étapes de propagation (ou s'arrête si feu éteint)."""
        if n <= 0:
            raise ValueError("n doit être un entier strictement positif")
        for _ in range(n):
            if self.fini:
                break
            self.step()

    # ---- Étape de propagation ----
    # Règle base (exercice) + extension probabiliste Vent/Météo
    # Pattern automate cellulaire : collecte synchrone puis application

    def step(self) -> bool:
        """
        Calcule la prochaine génération.
        Retourne True si le feu continue, False s'il est éteint.

        Note sur la synchronicité :
          Les allumages sont collectés AVANT d'être appliqués,
          évitant qu'un arbre mis en feu à cette étape propage
          immédiatement dans la même étape (biais de parcours).
          Alternative : utiliser deepcopy(self._grille) comme
          suggéré par l'exercice — équivalent mais plus coûteux.
        """
        prob_base = 0.30
        a_allumer: List[Tuple[int, int]] = []

        for y in range(self.nb_lignes):
            for x in range(self.nb_colonnes):
                if not self._grille[y][x].est_en_feu:
                    continue
                for nx, ny in self._coord_voisins(x, y):
                    if not self._grille[ny][nx].est_sain:
                        continue
                    dx, dy    = nx - x, ny - y
                    f_vent    = self.vent.facteur_propagation(dx, dy)
                    f_meteo   = self.meteo.facteur_propagation()
                    prob      = prob_base * f_vent * f_meteo
                    if random.random() < prob:
                        a_allumer.append((nx, ny))

        # Application synchrone des allumages
        for x, y in a_allumer:
            self._grille[y][x].mettre_en_feu()

        # Évolution de chaque cellule (combustion progressive)
        for y in range(self.nb_lignes):
            for x in range(self.nb_colonnes):
                self._grille[y][x].evoluer()

        self.etape += 1
        self.fini   = self.nb_en_feu() == 0
        return not self.fini

    # ---- Comptages ----

    def nb_en_feu(self)  -> int: return self._compter(Arbre.FEU)
    def nb_sains(self)   -> int: return self._compter(Arbre.SAIN)
    def nb_brulees(self) -> int: return self._compter(Arbre.BRULEE)
    def nb_vides(self)   -> int: return self._compter(Arbre.VIDE)

    def _compter(self, etat: int) -> int:
        return sum(
            1
            for y in range(self.nb_lignes)
            for x in range(self.nb_colonnes)
            if self._grille[y][x].etat == etat
        )

    def __repr__(self) -> str:
        cls = type(self).__name__
        return (f"{cls}({self.nb_lignes}×{self.nb_colonnes}, "
                f"étape={self.etape}, en_feu={self.nb_en_feu()})")


# ================================================================
#  FORET TORIQUE — classe enfant de Milieu (héritage)
#
#  Polymorphisme : _coord_voisins() est surchargée.
#  Les bords de la grille se connectent (grille cylindrique/torique).
#    bord droit  ↔ bord gauche
#    bord bas    ↔ bord haut
#
#  Tout le reste (step, proportion, afficher, simuler…) est hérité
#  de Milieu sans modification — c'est l'intérêt de l'héritage.
# ================================================================

class ForetTorique(Milieu):

    def __init__(
        self,
        nb_lignes:   int,
        nb_colonnes: int,
        p:           float,
        vent:        Vent,
        meteo:       Meteo,
    ) -> None:
        super().__init__(nb_lignes, nb_colonnes, p, vent, meteo)

    def _coord_voisins(self, x: int, y: int) -> List[Tuple[int, int]]:
        """
        Voisinage Von Neumann torique : les coordonnées sont calculées
        modulo les dimensions — tous les voisins existent toujours.
        """
        return [
            ((x + dx) % self.nb_colonnes, (y + dy) % self.nb_lignes)
            for dx, dy in Milieu.VOISINS_VN
        ]


# ================================================================
#  POINT D'ENTRÉE — démonstration console
# ================================================================

def demo_console() -> None:
    """
    Exemple d'utilisation des classes en mode console.
    Lance une simulation et affiche chaque génération.
    """
    print("=" * 50)
    print("  Simulation de Propagation de Feu")
    print("  Interface graphique : simulation_feu.html")
    print("=" * 50)

    # Paramètres
    vent  = Vent(direction='NE', vitesse=2)
    meteo = Meteo(humidite=35, temperature=32)

    # Choix du type de grille
    print("\n[1] Grille normale (Milieu)")
    print("[2] Grille torique (ForetTorique)")
    choix = input("Choix (1/2, défaut=1) : ").strip()

    if choix == '2':
        foret = ForetTorique(nb_lignes=15, nb_colonnes=20, p=0.65,
                             vent=vent, meteo=meteo)
        print("→ ForetTorique instanciée (héritage de Milieu)\n")
    else:
        foret = Milieu(nb_lignes=15, nb_colonnes=20, p=0.65,
                       vent=vent, meteo=meteo)
        print("→ Milieu instancié\n")

    print(f"Paramètres : {vent} | {meteo}")

    # Démarrage du foyer
    foret.demarrer_foyer_aleatoire()
    print(f"Foyer allumé — {foret.nb_en_feu()} cellule(s) en feu\n")

    # Simulation génération par génération
    try:
        n_gen = int(input("Nombre de générations à simuler (défaut=10) : ") or 10)
    except ValueError:
        n_gen = 10

    print()
    for i in range(n_gen):
        print(foret.afficher())
        print()
        if foret.fini:
            print(f"🔥 Feu éteint à l'étape {foret.etape}")
            break
        foret.step()
    else:
        print(f"Simulation terminée — {n_gen} générations.")

    print(f"\nRésumé final :")
    print(f"  Étapes      : {foret.etape}")
    print(f"  Arbres sains  : {foret.nb_sains()}")
    print(f"  Arbres brûlés : {foret.nb_brulees()}")
    print(f"  Proportion    : {foret.proportion():.2%}")


def demo_rapide() -> None:
    """Démonstration rapide sans interaction — pour tests unitaires."""
    print("\n--- Démonstration rapide ---")

    vent  = Vent('S', 2)
    meteo = Meteo(50, 25)
    foret = Milieu(10, 10, 0.7, vent, meteo)
    foret.demarrer_foyer_centre()

    print(f"Initial : {foret}")
    foret.simuler(5)
    print(f"Après 5 étapes : {foret}")
    print(f"Proportion sains : {foret.proportion():.2%}")

    # Test ForetTorique
    torique = ForetTorique(10, 10, 0.7, vent, meteo)
    torique.demarrer_foyer_centre()
    torique.simuler(5)
    print(f"ForetTorique après 5 étapes : {torique}")

    # Test Arbre isolé
    a = Arbre(Arbre.SAIN)
    print(f"\nArbre sain : {a.symbole()}")
    a.update(voisin_en_feu=True)
    print(f"Après update(voisin_en_feu=True) : {a.symbole()}")
    a.update(voisin_en_feu=False)
    print(f"Après update() suivant : {a.symbole()}")


if __name__ == '__main__':
    import sys
    if '--demo' in sys.argv:
        demo_rapide()
    else:
        demo_console()