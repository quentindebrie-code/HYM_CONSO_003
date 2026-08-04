# Reporting des consommations clients — Hympyr Énergies

Outil Streamlit : on dépose un export de livraisons (Excel ou CSV), on récupère l'état des
consommations par site de livraison en **PDF** (document client, charte Hympyr) et en
**Excel** (classeur de suivi, entièrement formulé).

Tout tient dans `app.py`. Aucun import local, donc aucun risque d'erreur `ModuleNotFoundError`
au déploiement.

## Déploiement

Le dépôt doit contenir ces trois fichiers **à sa racine** :

```
app.py
requirements.txt
packages.txt
```

Sur Streamlit Cloud, indiquer `app.py` comme fichier principal. `packages.txt` installe les
bibliothèques système dont WeasyPrint a besoin (Pango, Cairo, HarfBuzz) ainsi que la police
Poppins ; sans elle, WeasyPrint bascule sur Arial et le PDF reste correct mais hors charte.

En local :

```bash
pip install -r requirements.txt
streamlit run app.py
```

Sur macOS, WeasyPrint réclame `brew install pango libffi` au préalable.

Si WeasyPrint ne peut pas se charger, l'application ne plante pas : l'import est différé et
l'export Excel reste disponible, avec un message précisant la cause.

## Parcours utilisateur

1. **Fichier source** — dépôt du `.xlsx`, `.xls` ou `.csv`. Le séparateur et l'encodage du CSV
   sont détectés automatiquement (`;`, `,`, tabulation, `|` — UTF-8, CP1252, Latin-1).
2. **Colonnes** — détection automatique par similarité de libellé. Site, désignation et
   quantité sont obligatoires ; date et n° de BL sont facultatifs mais conditionnent la vue
   mensuelle et le comptage des livraisons.
3. **Lignes à traiter** — l'outil liste les désignations présentes et décoche celles qui ne
   portent aucune quantité (références de bons de commande, annulations, commentaires).
   C'est un filtre par volume, pas par préfixe de libellé : il résiste aux variantes de saisie.
4. **Paramètres** — nom du client, période, unité. Deux tables éditables permettent de
   regrouper des libellés de sites ou de renommer des produits.
5. **Contrôles qualité** — voir ci-dessous.
6. **Aperçu** — indicateurs clés et tableaux par site, par produit, par mois.
7. **Génération** — téléchargement du PDF et du classeur Excel.

## Traitements automatiques

**Avoirs et refacturations.** Les lignes en quantité négative sont typées « Avoir ». Les lignes
positives dont le préfixe de n° de BL diffère du préfixe dominant sont typées
« Refacturation ». Les trois types sont compensés dans les totaux. L'outil vérifie si le volume
net après compensation est identique au volume des livraisons d'origine et le signale : si
c'est le cas, les régularisations ne portaient que sur le prix.

**Comptage des livraisons.** Une livraison = un bon de livraison. Un BL portant plusieurs
produits n'est compté qu'une fois ; les avoirs et refacturations ne génèrent pas de livraison
supplémentaire. Le rapprochement se fait sur le n° de BL privé de son préfixe de type.

**Normalisation des sites.** Casse, accents, apostrophes, traits d'union après Saint / Sainte.
Les regroupements restent manuels, via la table éditable, car fusionner deux libellés proches
est une décision métier.

## Contrôles qualité

| Contrôle | Ce qu'il détecte |
|---|---|
| Ratio constant entre colonnes de prix | Un ratio uniforme (ex. 1,20) sur tous les produits signe une TVA, pas une marge : la colonne est probablement TTC |
| Libellés de sites très proches | Doublons de saisie non regroupés (ex. recopie incrémentale Excel sur une année) |
| Dates illisibles | Lignes comptées dans les totaux mais absentes de la vue mensuelle |
| BL multi-produits | Lignes multiples rattachées à une seule livraison |
| Lignes écartées | Rappel du volume de lignes filtrées, à rapprocher du fichier source |

## Valorisation financière

Désactivée par défaut : la demande courante porte sur les volumes. Si elle est activée, l'outil
demande explicitement la colonne de prix et sa base (HT ou TTC 20 %), et convertit en
conséquence. Les prix sont interprétés en euros pour 1 000 unités.

**Ne jamais retenir une colonne de prix d'achat** : c'est une donnée interne qui n'a pas à
figurer dans un document client.

## Livrables produits

**PDF, 3 pages.** Page 1 : indicateurs clés, répartition par produit, saisonnalité mensuelle.
Page 2 : tableau des consommations par site, classé par volume décroissant. Page 3 :
méthodologie et périmètre, avec les contrôles de cohérence chiffrés.

**Classeur Excel, 5 onglets.** Synthèse · Consommations par site · Volumes par site et par
mois · Détail des livraisons · Méthodologie. Les onglets de synthèse agrègent le détail par
`SUMIFS` : le classeur reste vivant si une ligne du détail est corrigée.

## Organisation de `app.py`

Le fichier est découpé en cinq sections repérables par des bandeaux de commentaires :

1. Chargement, détection de schéma et nettoyage
2. Agrégats
3. Export Excel
4. Export PDF
5. Interface Streamlit

Les quatre premières sections ne dépendent pas de Streamlit : elles restent appelables depuis
un notebook ou un script batch.

## Limites connues

- Les prix sont supposés exprimés pour 1 000 unités de quantité.
- La vue mensuelle couvre une seule année civile ; un export à cheval sur deux exercices
  agrège les mois de même rang.
- Le regroupement des sites n'est pas persisté d'une session à l'autre.
- Le rapprochement avoir / refacturation repose sur le préfixe du n° de BL ; un plan de
  numérotation différent demanderait d'adapter la fonction `preparer`.
