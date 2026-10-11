# Comparateur de prix des carburants

Outil Python qui trouve les stations-service les moins chères autour d'une ville, à partir des données ouvertes du ministère de l'Économie, **et qui vérifie que les prix affichés sont encore à jour**.

<img width="889" height="374" alt="image" src="https://github.com/user-attachments/assets/1214311e-3fbd-4feb-b674-5166f191f846" />

---

## Le problème

Les prix des carburants publiés en open data sont **déclarés par les stations elles-mêmes**. En testant une première version du programme, j'ai constaté qu'une station arrivait **1ère sur 4 carburants**… avec des prix qui n'avaient pas été mis à jour depuis **plus de 50 jours**.

Un classement naïf peut donc recommander des prix qui n'existent plus. Le programme sépare désormais les **prix fiables** des **prix douteux** avant d'établir le classement.

---

## Ce que fait le programme

| Étape | Description |
|---|---|
| 1. Saisie | L'utilisateur saisit une ville et un rayon en km. Les communes homonymes sont gérées (choix dans une liste). |
| 2. Géocodage | La ville est convertie en coordonnées GPS via l'API de la Géoplateforme (IGN). |
| 3. Collecte | Les prix de toutes les stations du rayon sont récupérés en temps réel, page par page (API paginée). |
| 4. Distances | La distance de chaque station est calculée avec la formule de haversine (calcul vectorisé avec numpy). |
| 5. Fiabilité | Chaque prix est daté : moins de 10 jours = fiable, 10 à 20 jours = à vérifier, plus de 20 jours = non représentatif. |
| 6. Restitution | Top 3 par carburant dans la console + rapport Excel mis en forme (prix fiables en vert, douteux en orange/rouge). |

Carburants couverts : Gazole, SP95, E10, SP98, E85.

---

## Installation et lancement

```bash
git clone https://github.com/<ton-pseudo>/comparateur-carburants.git
cd comparateur-carburants
pip install -r requirements.txt
python comparateur_carburants.py
```

Exemple :

```
Ville : Aubergenville
📍 Ville retenue : Aubergenville (78410)
Rayon en km (1 à 30) : 10

🏆 TOP 3 — PRIX FIABLES

⛽ Diesel (Gazole)
   1er | 2.239 €/L | 3.2 km | ...
```

Le rapport Excel est créé dans le dossier courant puis ouvert automatiquement (Windows, macOS et Linux).

---

## Choix techniques

| Choix | Pourquoi |
|---|---|
| **Format long** (`pd.melt`) | Les données arrivent avec une colonne par carburant. En passant à une ligne par couple (station, carburant), filtrer et classer devient beaucoup plus simple. |
| **Calcul vectorisé** (pandas / numpy) | Pas de boucle sur les lignes : le code reste lisible et passe à l'échelle (plusieurs milliers de stations). |
| **Tri sur 2 critères** | Prix croissant, puis distance : à prix égal, la station la plus proche passe devant. |
| **Seuils de fiabilité en constantes** | 10 et 20 jours sont des choix d'analyste, faciles à ajuster en haut du fichier. |
| **Gestion des erreurs** | `timeout` et `try/except` sur les appels réseau, validation des saisies, fichier Excel déjà ouvert : le programme ne plante pas. |
| **Code typé et documenté** | Type hints, docstrings au format Google, respect de PEP 8. |

---

## Limites

- Distance **à vol d'oiseau**, pas par la route.
- Les prix sont **déclarés** par les stations : le programme peut signaler un prix douteux, pas le corriger.
- Rayon limité à 30 km (la pagination de l'API est plafonnée à 10 000 résultats).

## Pistes d'amélioration

- Analyse à l'échelle de la France : part de prix périmés par département, écarts de prix entre départements.
- Carte interactive des stations.
- Application web (Streamlit).

---

## Sources des données

- [Prix des carburants en France — flux instantané (v2)](https://data.economie.gouv.fr/explore/dataset/prix-des-carburants-en-france-flux-instantane-v2/), ministère de l'Économie
- [Service de géocodage de la Géoplateforme](https://geoservices.ign.fr/documentation/services/services-geoplateforme/geocodage), IGN

---

## Auteur

**Anis NEFIKHA** — Étudiant en BUT Science des Données, IUT de Lille
https://www.linkedin.com/in/anis-nefikha/
