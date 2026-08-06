"""Outil de reporting des consommations clients — Hympyr Énergies.

Dépose un export de livraisons (Excel ou CSV), l'outil produit l'état des
consommations par site de livraison au format PDF (pour le client) et Excel
(pour le suivi interne).

Lancement :  streamlit run app.py
"""

import datetime as dt
import difflib
import html
import io
import re
import unicodedata
from dataclasses import dataclass, field

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

VERT, FONCE = "#1A9E68", "#073D27"

# WeasyPrint dépend de bibliothèques système (Pango, Cairo). L'import est différé
# pour que l'export Excel reste disponible même si le PDF ne peut pas être généré.
_WEASY = None
_WEASY_ERREUR = None


def _weasyprint():
    global _WEASY, _WEASY_ERREUR
    if _WEASY is None and _WEASY_ERREUR is None:
        try:
            from weasyprint import HTML
            _WEASY = HTML
        except Exception as exc:  # ImportError, OSError (libs système absentes)
            _WEASY_ERREUR = str(exc)
    return _WEASY


# ============================================================================
# CHARGEMENT, DÉTECTION DE SCHÉMA ET NETTOYAGE
# ============================================================================

import difflib
import io
import re
import unicodedata
from dataclasses import dataclass, field

import pandas as pd

MOIS = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août",
        "Septembre", "Octobre", "Novembre", "Décembre"]

# Mots-clés utilisés pour la détection automatique des colonnes.
MOTS_CLES = {
    "date": ["date", "date bl", "date livraison", "date de livraison"],
    "bl": ["bl", "n bl", "n° b.l.", "no bl", "bon de livraison", "piece", "pièce",
           "num bl", "numero bl", "n° bl"],
    "site": ["site", "site livraison", "site de livraison", "lieu", "point de livraison",
             "adresse livraison", "chantier", "depot", "dépôt"],
    "designation": ["designation", "désignation", "libelle", "libellé", "produit",
                    "article", "reference", "référence"],
    "quantite": ["quantite", "quantité", "qte", "qté", "volume", "litres", "qty"],
}


# ---------------------------------------------------------------- utilitaires
def _sans_accent(txt: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", str(txt))
                   if unicodedata.category(c) != "Mn")


def _cle(txt: str) -> str:
    return re.sub(r"[^a-z0-9 ]", " ", _sans_accent(txt).lower()).strip()


def normaliser_site(libelle: str) -> str:
    """Remet en forme un libellé de site : casse, accents conservés, traits d'union."""
    s = re.sub(r"\s+", " ", str(libelle).strip())
    if not s:
        return "(non renseigné)"
    if s.isupper() or s.islower():
        mots = s.split(" ")
        petits = {"de", "du", "des", "la", "le", "les", "sur", "sous", "en", "et", "d", "l"}
        out = []
        for i, m in enumerate(mots):
            if "'" in m:
                a, b = m.split("'", 1)
                out.append(a.lower() + "'" + (b.capitalize() if b else ""))
            elif i > 0 and m.lower() in petits:
                out.append(m.lower())
            else:
                out.append(m.capitalize())
        s = " ".join(out)
    # Communes composées : les mots liés par des petits mots prennent un trait d'union.
    s = re.sub(r"\b(Saint|Sainte|St|Ste)\s+", lambda m: m.group(1) + "-", s)
    return s


# ---------------------------------------------------------------- chargement
def charger_fichier(fichier, feuille: str | None = None) -> tuple[pd.DataFrame, list[str]]:
    """Charge un .xlsx/.xls/.csv et renvoie (dataframe, liste des feuilles)."""
    nom = getattr(fichier, "name", str(fichier)).lower()
    if nom.endswith(".csv") or nom.endswith(".txt"):
        brut = fichier.read() if hasattr(fichier, "read") else open(fichier, "rb").read()
        texte = None
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                texte = brut.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if texte is None:
            raise ValueError("Encodage du CSV non reconnu.")
        premiere = texte.split("\n", 1)[0]
        sep = max([";", ",", "\t", "|"], key=premiere.count)
        df = pd.read_csv(io.StringIO(texte), sep=sep, dtype=str)
        return df, []
    xl = pd.ExcelFile(fichier)
    feuilles = list(xl.sheet_names)
    df = pd.read_excel(xl, sheet_name=feuille or feuilles[0], dtype=object)
    return df, feuilles


def deviner_colonnes(df: pd.DataFrame) -> dict:
    """Propose une correspondance colonne source -> rôle fonctionnel."""
    cols = list(df.columns)
    cles = {c: _cle(c) for c in cols}
    mapping: dict[str, str | None] = {r: None for r in MOTS_CLES}
    pris: set = set()
    for role, mots in MOTS_CLES.items():
        best, score = None, 0.0
        for c in cols:
            if c in pris:
                continue
            k = cles[c]
            s = max([difflib.SequenceMatcher(None, k, _cle(m)).ratio() for m in mots])
            if any(_cle(m) == k for m in mots):
                s = 1.0
            elif any(_cle(m) in k for m in mots if len(m) > 3):
                s = max(s, 0.85)
            if s > score:
                best, score = c, s
        if score >= 0.6:
            mapping[role] = best
            pris.add(best)
    # Colonnes numériques restantes = candidats prix
    prix = []
    for c in cols:
        if c in pris:
            continue
        serie = pd.to_numeric(df[c], errors="coerce")
        if serie.notna().sum() > len(df) * 0.5:
            prix.append(c)
    mapping["prix_candidats"] = prix
    return mapping


# ---------------------------------------------------------------- nettoyage
@dataclass
class Resultat:
    donnees: pd.DataFrame
    produits: list[str] = field(default_factory=list)
    sites: list[str] = field(default_factory=list)
    stats: dict = field(default_factory=dict)
    alertes: list[dict] = field(default_factory=list)


def inventaire_designations(df: pd.DataFrame, col_des: str, col_qte: str) -> pd.DataFrame:
    """Tableau des désignations présentes, avec volume et nombre de lignes.

    Sert à proposer automatiquement les lignes à conserver : celles qui portent
    effectivement une quantité.
    """
    t = pd.DataFrame({
        "Désignation": df[col_des].astype(str).str.strip(),
        "Quantité": pd.to_numeric(df[col_qte], errors="coerce").fillna(0),
    })
    inv = (t.groupby("Désignation")
             .agg(Lignes=("Quantité", "size"),
                  Volume=("Quantité", "sum"),
                  Volume_absolu=("Quantité", lambda s: s.abs().sum()))
             .reset_index()
             .sort_values("Volume_absolu", ascending=False))
    inv["À conserver"] = inv["Volume_absolu"] > 0
    return inv[["À conserver", "Désignation", "Lignes", "Volume"]].reset_index(drop=True)


def preparer(
    df: pd.DataFrame,
    mapping: dict,
    designations_retenues: list[str],
    regroupement_sites: dict[str, str] | None = None,
    renommage_produits: dict[str, str] | None = None,
    col_prix: str | None = None,
    diviseur_prix: float = 1.0,
    debut=None,
    fin=None,
) -> Resultat:
    """Construit le jeu de données propre servant de base à tous les livrables."""
    c_date, c_bl = mapping.get("date"), mapping.get("bl")
    c_site, c_des, c_qte = mapping["site"], mapping["designation"], mapping["quantite"]

    d = pd.DataFrame({
        "Site_source": df[c_site].astype(str).str.strip(),
        "Produit_source": df[c_des].astype(str).str.strip(),
        "Quantite": pd.to_numeric(df[c_qte], errors="coerce").fillna(0),
    })
    d["Date"] = (pd.to_datetime(df[c_date], errors="coerce", dayfirst=True)
                 if c_date else pd.NaT)
    d["BL"] = df[c_bl].astype(str).str.strip() if c_bl else ""
    if col_prix:
        d["PU"] = pd.to_numeric(df[col_prix], errors="coerce").fillna(0) / (diviseur_prix or 1.0)

    total_lignes = len(d)
    retenues = set(str(x).strip() for x in designations_retenues)
    d = d[d["Produit_source"].isin(retenues)].copy()
    lignes_exclues = total_lignes - len(d)

    # --- Filtre de période : il s'applique à TOUT le rapport, pas seulement aux
    # vues temporelles. Les lignes sans date exploitable sont écartées dès qu'un
    # filtre est actif, faute de pouvoir les situer.
    lignes_hors_periode = 0
    lignes_sans_date = 0
    if debut is not None or fin is not None:
        avant = len(d)
        garde = d["Date"].notna()
        lignes_sans_date = int((~garde).sum())
        if debut is not None:
            garde &= d["Date"] >= pd.Timestamp(debut)
        if fin is not None:
            garde &= d["Date"] <= pd.Timestamp(fin) + pd.Timedelta(days=1) \
                     - pd.Timedelta(seconds=1)
        d = d[garde].copy()
        lignes_hors_periode = avant - len(d)

    # --- Typologie des pièces
    d["Type"] = "Livraison"
    d.loc[d["Quantite"] < 0, "Type"] = "Avoir"
    if c_bl:
        prefixes = d.loc[d["Type"] == "Livraison", "BL"].str[:2].str.upper()
        dominant = prefixes.mode().iloc[0] if not prefixes.empty else ""
        refac = (d["Quantite"] > 0) & (d["BL"].str[:2].str.upper() != dominant)
        d.loc[refac, "Type"] = "Refacturation"

    # --- Normalisation des sites
    d["Site"] = d["Site_source"].map(normaliser_site)
    if regroupement_sites:
        d["Site"] = d["Site"].map(lambda s: regroupement_sites.get(s, s))
    d["Produit"] = d["Produit_source"]
    if renommage_produits:
        d["Produit"] = d["Produit"].map(lambda p: renommage_produits.get(p, p))

    # --- Axe temporel
    if d["Date"].notna().any():
        d["Mois_num"] = d["Date"].dt.month
        d["Mois"] = d["Mois_num"].map(lambda m: MOIS[int(m) - 1] if pd.notna(m) else "")
    else:
        d["Mois_num"], d["Mois"] = 0, ""

    # --- Comptage des livraisons : un BL = une livraison
    d = d.sort_values(["Site", "Date", "BL", "Produit"], na_position="last").reset_index(drop=True)
    d["Cpt_BL"] = 0
    liv = d["Type"] == "Livraison"
    if c_bl and d.loc[liv, "BL"].str.len().gt(2).any():
        cle_bl = d.loc[liv, "BL"].str[2:]
        d.loc[liv, "Cpt_BL"] = (~cle_bl.duplicated()).astype(int)
    else:
        d.loc[liv, "Cpt_BL"] = 1

    if "PU" in d:
        d["Montant"] = d["Quantite"] * d["PU"] / 1000

    produits = sorted(d["Produit"].unique(),
                      key=lambda p: -d.loc[d.Produit == p, "Quantite"].sum())
    sites = sorted(d["Site"].unique())

    stats = {
        "lignes_source": total_lignes,
        "lignes_exclues": lignes_exclues,
        "lignes_hors_periode": lignes_hors_periode,
        "lignes_sans_date": lignes_sans_date,
        "filtre_debut": debut,
        "filtre_fin": fin,
        "lignes_retenues": len(d),
        "lignes_livraison": int((d.Type == "Livraison").sum()),
        "lignes_avoir": int((d.Type == "Avoir").sum()),
        "lignes_refac": int((d.Type == "Refacturation").sum()),
        "nb_sites": len(sites),
        "nb_produits": len(produits),
        "nb_livraisons": int(d["Cpt_BL"].sum()),
        "volume": float(d["Quantite"].sum()),
        "volume_livraisons": float(d.loc[d.Type == "Livraison", "Quantite"].sum()),
        "date_min": d["Date"].min(),
        "date_max": d["Date"].max(),
    }
    if "Montant" in d:
        stats["montant"] = float(d["Montant"].sum())

    return Resultat(donnees=d, produits=produits, sites=sites, stats=stats)


# ---------------------------------------------------------------- diagnostics
def diagnostiquer(df_source: pd.DataFrame, mapping: dict, res: Resultat) -> list[dict]:
    """Contrôles de qualité sur les données. Renvoie une liste d'alertes."""
    a: list[dict] = []
    s = res.stats

    if s["lignes_exclues"]:
        a.append({"niveau": "info", "titre": "Lignes écartées",
                  "detail": f"{s['lignes_exclues']} lignes sur {s['lignes_source']} ne portent "
                            f"aucune quantité ou n'ont pas été retenues comme produit."})

    if s.get("lignes_hors_periode"):
        detail = (f"{s['lignes_hors_periode']} ligne(s) écartée(s) car hors de la période "
                  f"retenue. Le rapport entier — sites, produits, mois, détail — porte "
                  f"uniquement sur cette période.")
        if s.get("lignes_sans_date"):
            detail += (f" Dont {s['lignes_sans_date']} sans date exploitable, "
                       f"impossibles à situer.")
        a.append({"niveau": "info", "titre": "Filtre de période appliqué", "detail": detail})

    # Un avoir dont la facture d'origine est hors période fausse les totaux.
    if s["lignes_avoir"] or s["lignes_refac"]:
        d0 = res.donnees
        if d0["BL"].str.len().gt(2).any():
            cles_liv = set(d0.loc[d0.Type == "Livraison", "BL"].str[2:])
            orph = d0[(d0.Type != "Livraison") & (~d0["BL"].str[2:].isin(cles_liv))]
            if len(orph):
                a.append({
                    "niveau": "alerte", "titre": "Avoirs sans facture d'origine",
                    "detail": f"{len(orph)} avoir(s) ou refacturation(s) n'ont pas de "
                              f"livraison correspondante dans le périmètre retenu "
                              f"({orph['Quantite'].sum():+,.0f} unités). La facture "
                              f"d'origine est probablement hors période : les totaux "
                              f"s'en trouvent minorés ou majorés."})

    if s["lignes_avoir"]:
        ecart = s["volume"] - s["volume_livraisons"]
        detail = (f"{s['lignes_avoir']} avoir(s) et {s['lignes_refac']} refacturation(s) "
                  f"détectés et compensés dans les totaux.")
        detail += (" Le volume net est identique au volume des livraisons d'origine : "
                   "les régularisations ne portent pas sur les quantités."
                   if abs(ecart) < 0.5 else
                   f" Écart de {ecart:+,.0f} L entre volume net et livraisons d'origine : "
                   "les régularisations modifient les quantités.")
        a.append({"niveau": "info", "titre": "Avoirs et régularisations", "detail": detail})

    # Ratio constant entre deux colonnes de prix : signature d'une TVA, pas d'une marge.
    cands = mapping.get("prix_candidats") or []
    for i in range(len(cands)):
        for j in range(i + 1, len(cands)):
            x = pd.to_numeric(df_source[cands[i]], errors="coerce")
            y = pd.to_numeric(df_source[cands[j]], errors="coerce")
            m = (x > 0) & (y > 0)
            if m.sum() < 10:
                continue
            r = (y[m] / x[m]).round(4)
            part = r.value_counts(normalize=True)
            if part.iloc[0] >= 0.8 and part.index[0] != 1.0:
                a.append({
                    "niveau": "alerte", "titre": "Ratio constant entre deux colonnes de prix",
                    "detail": f"« {cands[j]} » = {part.index[0]:.2f} × « {cands[i]} » sur "
                              f"{part.iloc[0]:.0%} des lignes. Un ratio uniforme sur tous les "
                              f"produits évoque une TVA plutôt qu'une marge : vérifier la base "
                              f"HT/TTC avant toute valorisation."})

    # Libellés de sites proches : risque de doublon non regroupé.
    proches = []
    restants = list(res.sites)
    for site in list(res.sites):
        if site not in restants:
            continue
        restants.remove(site)
        voisins = difflib.get_close_matches(site, restants, n=3, cutoff=0.88)
        for v in voisins:
            proches.append(f"« {site} » / « {v} »")
            restants.remove(v)
    if proches:
        a.append({"niveau": "alerte", "titre": "Libellés de sites très proches",
                  "detail": "Regroupement possible : " + " · ".join(proches[:8])
                            + (" …" if len(proches) > 8 else "")
                            + ". À vérifier dans l'étape « Sites de livraison »."})

    d = res.donnees
    if d["Date"].isna().any():
        a.append({"niveau": "alerte", "titre": "Dates illisibles",
                  "detail": f"{int(d['Date'].isna().sum())} ligne(s) sans date exploitable : "
                            f"elles sont comptées dans les totaux mais absentes de la vue "
                            f"mensuelle."})

    if mapping.get("bl"):
        multi = d[d.Type == "Livraison"].groupby(d["BL"].str[2:])["Produit"].nunique()
        n_multi = int((multi > 1).sum())
        if n_multi:
            a.append({"niveau": "info", "titre": "Bons de livraison multi-produits",
                      "detail": f"{n_multi} bon(s) de livraison portent plusieurs produits. "
                                f"Ils ne sont comptés qu'une fois dans le nombre de livraisons."})

    if not a:
        a.append({"niveau": "ok", "titre": "Aucune anomalie détectée",
                  "detail": "Les contrôles automatiques n'ont rien relevé sur ce fichier."})
    return a


def tableau_sites(res: Resultat) -> pd.DataFrame:
    """Table éditable des libellés de sites (source -> libellé retenu)."""
    d = res.donnees
    t = (d.groupby(["Site_source", "Site"])
           .agg(Volume=("Quantite", "sum"), Lignes=("Quantite", "size"))
           .reset_index()
           .rename(columns={"Site_source": "Libellé source", "Site": "Libellé retenu"}))
    return t.sort_values("Libellé retenu").reset_index(drop=True)

# ============================================================================
# AGRÉGATS
# ============================================================================

import pandas as pd



def par_produit(res: Resultat) -> pd.DataFrame:
    d = res.donnees
    liv = d[d.Type == "Livraison"]
    t = pd.DataFrame(index=res.produits)
    t["Volume"] = d.groupby("Produit")["Quantite"].sum().reindex(res.produits).fillna(0)
    t["Part"] = t["Volume"] / t["Volume"].sum() if t["Volume"].sum() else 0
    t["Sites"] = liv.groupby("Produit")["Site"].nunique().reindex(res.produits).fillna(0)
    t["Lignes"] = liv.groupby("Produit").size().reindex(res.produits).fillna(0)
    if "Montant" in d:
        t["Montant"] = d.groupby("Produit")["Montant"].sum().reindex(res.produits).fillna(0)
        t["Prix moyen"] = (t["Montant"] / t["Volume"] * 1000).where(t["Volume"] != 0, 0)
    return t


def par_site(res: Resultat, tri: str = "volume") -> pd.DataFrame:
    d = res.donnees
    t = d.groupby("Site").agg(Livraisons=("Cpt_BL", "sum"), Volume=("Quantite", "sum"))
    pivot = d.pivot_table(index="Site", columns="Produit", values="Quantite",
                          aggfunc="sum", fill_value=0)
    pivot = pivot.reindex(columns=res.produits, fill_value=0)
    t = t.join(pivot)
    t["Part"] = t["Volume"] / t["Volume"].sum() if t["Volume"].sum() else 0
    t["Moyenne"] = (t["Volume"] / t["Livraisons"]).where(t["Livraisons"] != 0, 0)
    if "Montant" in d:
        t["Montant"] = d.groupby("Site")["Montant"].sum()
        t["Prix moyen"] = (t["Montant"] / t["Volume"] * 1000).where(t["Volume"] != 0, 0)
    return t.sort_values("Volume", ascending=False) if tri == "volume" else t.sort_index()


def mois_actifs(res: Resultat) -> list:
    """Mois couverts par les données, du premier au dernier observé.

    Sur un rapport filtré sur un semestre, afficher douze colonnes dont six vides
    est du bruit : on se limite à l'intervalle réellement livré.
    """
    m = res.donnees.loc[res.donnees["Mois_num"] > 0, "Mois_num"]
    if m.empty:
        return []
    return list(range(int(m.min()), int(m.max()) + 1))


def par_mois(res: Resultat) -> pd.Series:
    actifs = mois_actifs(res)
    if not actifs:
        return pd.Series(dtype=float)
    d = res.donnees[res.donnees["Mois_num"] > 0]
    s = d.groupby("Mois_num")["Quantite"].sum().reindex(actifs, fill_value=0)
    s.index = [MOIS[m - 1] for m in actifs]
    return s


def site_mois(res: Resultat) -> pd.DataFrame:
    actifs = mois_actifs(res)
    d = res.donnees[res.donnees["Mois_num"] > 0]
    p = d.pivot_table(index="Site", columns="Mois_num", values="Quantite",
                      aggfunc="sum", fill_value=0)
    p = p.reindex(columns=actifs, fill_value=0)
    p.columns = [MOIS[m - 1] for m in actifs]
    p = p.reindex(sorted(res.sites), fill_value=0)
    p["Total"] = p.sum(axis=1)
    return p


def prix_moyen_periode(res: Resultat, granularite: str = "Mois",
                       debut=None, fin=None, produits=None) -> pd.DataFrame:
    """Évolution du prix moyen pondéré sur une période.

    Le prix moyen d'une période n'est pas la moyenne des prix unitaires : c'est le
    montant total divisé par le volume total, sinon une petite livraison chère pèse
    autant qu'un plein camion. Les avoirs et refacturations sont inclus, puisqu'ils
    corrigent précisément le prix facturé.
    """
    d = res.donnees
    if "Montant" not in d.columns:
        return pd.DataFrame()
    d = d[d["Date"].notna()].copy()
    if debut is not None:
        d = d[d["Date"] >= pd.Timestamp(debut)]
    if fin is not None:
        d = d[d["Date"] <= pd.Timestamp(fin)]
    if produits:
        d = d[d["Produit"].isin(produits)]
    if d.empty:
        return pd.DataFrame()

    code = {"Mois": "M", "Semaine": "W", "Trimestre": "Q"}[granularite]
    freq = {"Mois": "MS", "Semaine": "W-MON", "Trimestre": "QS"}[granularite]
    d["Periode"] = d["Date"].dt.to_period(code).dt.start_time

    t = (d.groupby("Periode")
           .agg(Volume=("Quantite", "sum"), Montant=("Montant", "sum"),
                Livraisons=("Cpt_BL", "sum")))
    index = pd.date_range(t.index.min(), t.index.max(), freq=freq)
    t = t.reindex(index, fill_value=0)
    t.index.name = "Periode"
    t["Prix moyen"] = (t["Montant"] / t["Volume"] * 1000).where(t["Volume"] != 0)
    return t


def bornes_periode(ts, granularite: str):
    """Renvoie (premier jour, dernier jour) de la période commençant à ts."""
    code = {"Mois": "M", "Semaine": "W", "Trimestre": "Q"}[granularite]
    p = pd.Timestamp(ts).to_period(code)
    return p.start_time.date(), p.end_time.date()


def libelle_periode(ts, granularite: str) -> str:
    """Étiquette courte d'une période, adaptée à la granularité."""
    ts = pd.Timestamp(ts)
    if granularite == "Mois":
        return MOIS[ts.month - 1][:3] + " " + str(ts.year)[2:]
    if granularite == "Trimestre":
        return f"T{(ts.month - 1) // 3 + 1} {str(ts.year)[2:]}"
    return ts.strftime("%d/%m/%y")

# ============================================================================
# EXPORT EXCEL
# ============================================================================

import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter


XL_VERT = "1A9E68"
XL_FONCE = "073D27"
XL_CLAIR = "E8F5EF"
XL_GRIS = "F5F6F5"
XL_BLANC = "FFFFFF"
FONT = "Arial"

FMT_L = '#,##0 "L";-#,##0 "L";"-"'
FMT_PCT = '0.0%;-0.0%;"-"'
FMT_NB = '#,##0;-#,##0;"-"'
FMT_EUR = '#,##0.00 "€";-#,##0.00 "€";"-"'
FMT_DEC = '#,##0.00'

_xl_thin = Side(style="thin", color="D0D4D2")
XL_BORD = Border(left=_xl_thin, right=_xl_thin, top=_xl_thin, bottom=_xl_thin)
XL_D = "'Détail des livraisons'"


def _xl_base(ws):
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = XL_VERT


def _xl_titre(ws, texte, sous_titre):
    ws["A1"] = "HYMPYR ÉNERGIES"
    ws["A1"].font = Font(name=FONT, size=9, bold=True, color=XL_VERT)
    ws["A2"] = texte
    ws["A2"].font = Font(name=FONT, size=16, bold=True, color=XL_FONCE)
    ws["A3"] = sous_titre
    ws["A3"].font = Font(name=FONT, size=10, color="6B7280")
    ws.row_dimensions[1].height = 14
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 16
    ws.row_dimensions[4].height = 6


def _xl_entetes(ws, ligne, valeurs, largeurs):
    for i, (v, w) in enumerate(zip(valeurs, largeurs), start=1):
        c = ws.cell(row=ligne, column=i, value=v)
        c.font = Font(name=FONT, size=9, bold=True, color=XL_BLANC)
        c.fill = PatternFill("solid", fgColor=XL_FONCE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = XL_BORD
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[ligne].height = 32


def _xl_habiller(ws, r0, rt, ncol):
    for rr in range(r0, rt + 1):
        for col in range(1, ncol + 1):
            c = ws.cell(row=rr, column=col)
            c.border = XL_BORD
            c.font = Font(name=FONT, size=9, bold=(rr == rt),
                          color=XL_FONCE if rr == rt else "000000")
            if rr == rt:
                c.fill = PatternFill("solid", fgColor=XL_CLAIR)
            elif (rr - r0) % 2 == 1:
                c.fill = PatternFill("solid", fgColor=XL_GRIS)


def construire_classeur(res: Resultat, client: str, periode: str, unite: str = "L",
               libelle_montant: str = "Montant",
               prix_options: dict | None = None) -> bytes:
    d = res.donnees
    produits = res.produits
    sites = sorted(res.sites)
    ns, npr = len(sites), len(produits)
    argent = "Montant" in d.columns
    entete = f"Client : {client}  |  Période : {periode}"

    # Colonnes de l'onglet Détail : A..H fixes, I/J si valorisation
    COL_MOIS, COL_TYPE, COL_SITE, COL_PROD, COL_QTE, COL_CPT = "B", "D", "E", "F", "G", "H"
    COL_MNT = "J"

    wb = Workbook()

    # ------------------------------------------------------------ 1. Synthèse
    ws = wb.active
    ws.title = "Synthèse"
    _xl_base(ws)
    _xl_titre(ws, "État des consommations", f"{entete}  |  Volumes livrés")
    for col, w in zip("ABCDEFGH", [30, 16, 16, 16, 18, 4, 32, 18]):
        ws.column_dimensions[col].width = w

    rt_site = 6 + ns
    kpis = [("Volume total livré", f"='Consommations par site'!$G${rt_site}", FMT_L),
            ("Nombre de livraisons", f"='Consommations par site'!$B${rt_site}", FMT_NB),
            ("Sites livrés", ns, FMT_NB),
            ("Volume moyen par livraison", f"='Consommations par site'!$I${rt_site}", FMT_L)]
    ws.cell(row=6, column=1, value="CHIFFRES CLÉS").font = Font(
        name=FONT, size=10, bold=True, color=XL_FONCE)
    for i, (lib, val, fmt) in enumerate(kpis):
        col = 1 + i * 2
        c1 = ws.cell(row=7, column=col, value=lib)
        c1.font = Font(name=FONT, size=8, bold=True, color="6B7280")
        c2 = ws.cell(row=8, column=col, value=val)
        c2.font = Font(name=FONT, size=14, bold=True, color=XL_FONCE)
        c2.number_format = fmt
        for c in (c1, c2):
            c.alignment = Alignment(horizontal="center")
            c.fill = PatternFill("solid", fgColor=XL_CLAIR)
        ws.merge_cells(start_row=7, start_column=col, end_row=7, end_column=col + 1)
        ws.merge_cells(start_row=8, start_column=col, end_row=8, end_column=col + 1)
    ws.row_dimensions[7].height = 16
    ws.row_dimensions[8].height = 26

    ws.cell(row=11, column=1, value="RÉPARTITION PAR PRODUIT").font = Font(
        name=FONT, size=10, bold=True, color=XL_FONCE)
    hd = ["Produit", f"Volume ({unite})", "Part du volume", "Nb de sites livrés",
          "Nb de lignes de livraison"]
    lg = [30, 16, 16, 16, 18]
    if argent:
        hd += [f"{libelle_montant} (€)", "Prix moyen (€/1000 L)"]
        lg += [18, 18]
    _xl_entetes(ws, 12, hd, lg)
    tp = par_produit(res)
    lp = 13
    for i, p in enumerate(produits):
        rr = lp + i
        ws.cell(row=rr, column=1, value=p)
        ws.cell(row=rr, column=2,
                value=f'=SUMIFS({XL_D}!${COL_QTE}:${COL_QTE},{XL_D}!${COL_PROD}:${COL_PROD},$A{rr})')
        ws.cell(row=rr, column=3, value=f"=IFERROR($B{rr}/$B${lp + npr},0)")
        ws.cell(row=rr, column=4, value=int(tp.loc[p, "Sites"]))
        ws.cell(row=rr, column=5,
                value=f'=COUNTIFS({XL_D}!${COL_PROD}:${COL_PROD},$A{rr},'
                      f'{XL_D}!${COL_TYPE}:${COL_TYPE},"Livraison")')
        if argent:
            ws.cell(row=rr, column=6,
                    value=f'=SUMIFS({XL_D}!${COL_MNT}:${COL_MNT},'
                          f'{XL_D}!${COL_PROD}:${COL_PROD},$A{rr})')
            ws.cell(row=rr, column=7, value=f"=IFERROR($F{rr}/$B{rr}*1000,0)")
    rt = lp + npr
    ws.cell(row=rt, column=1, value="TOTAL")
    for col in ([2, 5] + ([6] if argent else [])):
        L = get_column_letter(col)
        ws.cell(row=rt, column=col, value=f"=SUM({L}{lp}:{L}{rt - 1})")
    ws.cell(row=rt, column=3, value=f"=IFERROR($B{rt}/$B${rt},0)")
    ws.cell(row=rt, column=4, value=ns)
    if argent:
        ws.cell(row=rt, column=7, value=f"=IFERROR($F{rt}/$B{rt}*1000,0)")
    _xl_habiller(ws, lp, rt, 7 if argent else 5)
    for rr in range(lp, rt + 1):
        ws.cell(row=rr, column=2).number_format = FMT_L
        ws.cell(row=rr, column=3).number_format = FMT_PCT
        ws.cell(row=rr, column=4).number_format = FMT_NB
        ws.cell(row=rr, column=5).number_format = FMT_NB
        if argent:
            ws.cell(row=rr, column=6).number_format = FMT_EUR
            ws.cell(row=rr, column=7).number_format = FMT_DEC

    # Le Top 10 est décalé à droite du tableau produits, dont la largeur dépend
    # de la présence ou non des colonnes de valorisation.
    ct = 9 if argent else 7
    LT = get_column_letter(ct)
    ws.column_dimensions[LT].width = 32
    ws.column_dimensions[get_column_letter(ct + 1)].width = 18
    top = par_site(res).head(10).index.tolist()
    ws.cell(row=11, column=ct, value="TOP 10 DES SITES PAR VOLUME").font = Font(
        name=FONT, size=10, bold=True, color=XL_FONCE)
    for i, v in enumerate(["Site de livraison", f"Volume ({unite})"]):
        c = ws.cell(row=12, column=ct + i, value=v)
        c.font = Font(name=FONT, size=9, bold=True, color=XL_BLANC)
        c.fill = PatternFill("solid", fgColor=XL_FONCE)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = XL_BORD
    for i, s in enumerate(top):
        rr = 13 + i
        ws.cell(row=rr, column=ct, value=s)
        ws.cell(row=rr, column=ct + 1,
                value=f'=SUMIFS({XL_D}!${COL_QTE}:${COL_QTE},'
                      f'{XL_D}!${COL_SITE}:${COL_SITE},${LT}{rr})')
        for col in (ct, ct + 1):
            c = ws.cell(row=rr, column=col)
            c.border = XL_BORD
            c.font = Font(name=FONT, size=9)
            if i % 2 == 1:
                c.fill = PatternFill("solid", fgColor=XL_GRIS)
        ws.cell(row=rr, column=ct + 1).number_format = FMT_L

    note = ("Document établi par Hympyr Énergies à partir des bons de livraison de la période. "
            + ("Voir l'onglet « Méthodologie » pour le périmètre et les règles de calcul."
               if argent else
               "Cet état porte exclusivement sur les volumes livrés ; il ne comporte aucune "
               "valorisation financière. Voir l'onglet « Méthodologie »."))
    ws.cell(row=26, column=1, value=note).font = Font(
        name=FONT, size=8, italic=True, color="6B7280")

    # -------------------------------------------- 2. Consommations par site
    ws = wb.create_sheet("Consommations par site")
    _xl_base(ws)
    _xl_titre(ws, "Consommations par site de livraison", f"{entete}  |  Volumes nets d'avoirs")
    hd = ["Site de livraison", "Nb de\nlivraisons"] + [f"{p}\n({unite})" for p in produits] \
         + [f"Volume total\n({unite})", "Part du\nvolume", f"Volume moyen\npar livraison ({unite})"]
    lg = [32, 12] + [14] * npr + [15, 11, 16]
    if argent:
        hd += [f"{libelle_montant}\n(€)", "Prix moyen\n(€/1000 L)"]
        lg += [16, 14]
    _xl_entetes(ws, 5, hd, lg)
    # Ligne 4 : libellés produits « en clair », invisibles, servant de critère aux SUMIFS.
    # Référencer une cellule évite tout problème de caractère joker ou de guillemet
    # dans un nom de produit écrit en dur dans la formule.
    for j, p in enumerate(produits):
        c = ws.cell(row=4, column=3 + j, value=p)
        c.font = Font(name=FONT, size=1, color=XL_BLANC)
    ws.row_dimensions[4].height = 3
    c_vol = get_column_letter(3 + npr)
    c_mnt = get_column_letter(6 + npr)
    r0 = 6
    for i, s in enumerate(sites):
        rr = r0 + i
        ws.cell(row=rr, column=1, value=s)
        ws.cell(row=rr, column=2,
                value=f'=SUMIFS({XL_D}!${COL_CPT}:${COL_CPT},{XL_D}!${COL_SITE}:${COL_SITE},$A{rr})')
        for j, p in enumerate(produits):
            L = get_column_letter(3 + j)
            ws.cell(row=rr, column=3 + j,
                    value=f'=SUMIFS({XL_D}!${COL_QTE}:${COL_QTE},'
                          f'{XL_D}!${COL_SITE}:${COL_SITE},$A{rr},'
                          f'{XL_D}!${COL_PROD}:${COL_PROD},{L}$4)')
        ws.cell(row=rr, column=3 + npr, value=f"=SUM(C{rr}:{get_column_letter(2 + npr)}{rr})")
        ws.cell(row=rr, column=4 + npr,
                value=f"=IFERROR(${c_vol}{rr}/${c_vol}${r0 + ns},0)")
        ws.cell(row=rr, column=5 + npr, value=f"=IFERROR(${c_vol}{rr}/$B{rr},0)")
        if argent:
            ws.cell(row=rr, column=6 + npr,
                    value=f'=SUMIFS({XL_D}!${COL_MNT}:${COL_MNT},'
                          f'{XL_D}!${COL_SITE}:${COL_SITE},$A{rr})')
            ws.cell(row=rr, column=7 + npr,
                    value=f"=IFERROR(${c_mnt}{rr}/${c_vol}{rr}*1000,0)")
    rt = r0 + ns
    ws.cell(row=rt, column=1, value=f"TOTAL {client.upper()}")
    for col in list(range(2, 4 + npr)) + ([6 + npr] if argent else []):
        L = get_column_letter(col)
        ws.cell(row=rt, column=col, value=f"=SUM({L}{r0}:{L}{rt - 1})")
    ws.cell(row=rt, column=4 + npr, value=f"=IFERROR(${c_vol}{rt}/${c_vol}${rt},0)")
    ws.cell(row=rt, column=5 + npr, value=f"=IFERROR(${c_vol}{rt}/$B{rt},0)")
    if argent:
        ws.cell(row=rt, column=7 + npr, value=f"=IFERROR(${c_mnt}{rt}/${c_vol}{rt}*1000,0)")
    ncol = (7 if argent else 5) + npr
    _xl_habiller(ws, r0, rt, ncol)
    for rr in range(r0, rt + 1):
        ws.cell(row=rr, column=2).number_format = FMT_NB
        for col in range(3, 4 + npr):
            ws.cell(row=rr, column=col).number_format = FMT_L
        ws.cell(row=rr, column=4 + npr).number_format = FMT_PCT
        ws.cell(row=rr, column=5 + npr).number_format = FMT_L
        if argent:
            ws.cell(row=rr, column=6 + npr).number_format = FMT_EUR
            ws.cell(row=rr, column=7 + npr).number_format = FMT_DEC
    ws.freeze_panes = "B6"
    ws.print_title_rows = "5:5"
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1

    # ------------------------------------------- 3. Volumes par site et mois
    mois_lib = list(par_mois(res).index)
    if mois_lib:
        nm = len(mois_lib)
        c_tot = 2 + nm
        ws = wb.create_sheet("Volumes par site et par mois")
        _xl_base(ws)
        _xl_titre(ws, "Volumes livrés par site et par mois",
               f"{entete}  |  En {unite}, tous produits confondus")
        _xl_entetes(ws, 5, ["Site de livraison"] + mois_lib + [f"Total ({unite})"],
                 [32] + [10] * nm + [13])
        r0 = 6
        for i, s_nom in enumerate(sites):
            rr = r0 + i
            ws.cell(row=rr, column=1, value=s_nom)
            for j in range(nm):
                L = get_column_letter(2 + j)
                ws.cell(row=rr, column=2 + j,
                        value=f'=SUMIFS({XL_D}!${COL_QTE}:${COL_QTE},'
                              f'{XL_D}!${COL_SITE}:${COL_SITE},$A{rr},'
                              f'{XL_D}!${COL_MOIS}:${COL_MOIS},{L}$5)')
            ws.cell(row=rr, column=c_tot,
                    value=f"=SUM(B{rr}:{get_column_letter(1 + nm)}{rr})")
        rt = r0 + ns
        ws.cell(row=rt, column=1, value="TOTAL")
        for col in range(2, c_tot + 1):
            L = get_column_letter(col)
            ws.cell(row=rt, column=col, value=f"=SUM({L}{r0}:{L}{rt - 1})")
        _xl_habiller(ws, r0, rt, c_tot)
        for rr in range(r0, rt + 1):
            for col in range(2, c_tot + 1):
                ws.cell(row=rr, column=col).number_format = FMT_NB
                if col == c_tot:
                    ws.cell(row=rr, column=col).font = Font(
                        name=FONT, size=9, bold=True,
                        color=XL_FONCE if rr == rt else "000000")
        ws.freeze_panes = "B6"
        ws.print_title_rows = "5:5"
        ws.page_setup.orientation = "landscape"
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1

    # ------------------------------------------ 3 bis. Évolution du prix moyen
    if argent and prix_options and prix_options.get("actif"):
        gr = prix_options.get("granularite", "Mois")
        filtre = prix_options.get("produits") or []
        tpx = prix_moyen_periode(res, gr, prix_options.get("debut"),
                                          prix_options.get("fin"), filtre)
        if len(tpx) >= 2:
            ws = wb.create_sheet("Évolution du prix moyen")
            _xl_base(ws)
            perim = ", ".join(filtre) if filtre else "tous produits confondus"
            _xl_titre(ws, "Évolution du prix moyen",
                   f"{entete}  |  Granularité : {gr.lower()}  |  {perim}")
            _xl_entetes(ws, 5, ["Période", "Du", "Au", f"Volume\n({unite})",
                             f"{libelle_montant}\n(€)", "Prix moyen\n(€/1000 L)",
                             "Nb de\nlivraisons", "Écart vs\nmoyenne"],
                     [16, 12, 12, 15, 16, 16, 12, 13])

            def _somme(colonne, ligne):
                """SUMIFS bornée en date, additionnée produit par produit si filtre."""
                base = (f'SUMIFS({XL_D}!${colonne}:${colonne},'
                        f'{XL_D}!$A:$A,">="&$B{ligne},{XL_D}!$A:$A,"<="&$C{ligne}')
                if not filtre:
                    return "=" + base + ")"
                return "=" + "+".join(
                    base + f',{XL_D}!${COL_PROD}:${COL_PROD},$J${6 + i})'
                    for i in range(len(filtre)))

            # Colonne technique J : libellés des produits filtrés, servant de critère.
            for i, p in enumerate(filtre):
                c = ws.cell(row=6 + i, column=10, value=p)
                c.font = Font(name=FONT, size=1, color=XL_BLANC)
            ws.column_dimensions["J"].width = 2

            r0 = 6
            for i, (ts, row) in enumerate(tpx.iterrows()):
                rr = r0 + i
                d1, d2 = bornes_periode(ts, gr)
                ws.cell(row=rr, column=1, value=libelle_periode(ts, gr))
                ws.cell(row=rr, column=2, value=d1).number_format = "DD/MM/YYYY"
                ws.cell(row=rr, column=3, value=d2).number_format = "DD/MM/YYYY"
                ws.cell(row=rr, column=4, value=_somme(COL_QTE, rr))
                ws.cell(row=rr, column=5, value=_somme(COL_MNT, rr))
                ws.cell(row=rr, column=6, value=f'=IFERROR($E{rr}/$D{rr}*1000,"")')
                ws.cell(row=rr, column=7, value=_somme(COL_CPT, rr))
                ws.cell(row=rr, column=8,
                        value=f'=IFERROR($F{rr}/$F${r0 + len(tpx)}-1,"")')
            rt = r0 + len(tpx)
            ws.cell(row=rt, column=1, value="MOYENNE PONDÉRÉE")
            for col in (4, 5, 7):
                L = get_column_letter(col)
                ws.cell(row=rt, column=col, value=f"=SUM({L}{r0}:{L}{rt - 1})")
            ws.cell(row=rt, column=6, value=f'=IFERROR($E{rt}/$D{rt}*1000,"")')
            ws.cell(row=rt, column=8, value=0)
            _xl_habiller(ws, r0, rt, 8)
            for rr in range(r0, rt + 1):
                ws.cell(row=rr, column=1).alignment = Alignment(horizontal="center")
                ws.cell(row=rr, column=4).number_format = FMT_L
                ws.cell(row=rr, column=5).number_format = FMT_EUR
                ws.cell(row=rr, column=6).number_format = FMT_DEC
                ws.cell(row=rr, column=7).number_format = FMT_NB
                ws.cell(row=rr, column=8).number_format = '+0.0%;-0.0%;0.0%'

            graphe = LineChart()
            graphe.title = "Prix moyen pondéré (€/1000 L)"
            graphe.height, graphe.width = 8.5, 20
            graphe.y_axis.title = "€ / 1000 L"
            graphe.legend = None
            donnees = Reference(ws, min_col=6, min_row=5, max_row=rt - 1)
            cats = Reference(ws, min_col=1, min_row=r0, max_row=rt - 1)
            graphe.add_data(donnees, titles_from_data=True)
            graphe.set_categories(cats)
            serie = graphe.series[0]
            serie.graphicalProperties.line.solidFill = XL_VERT
            serie.graphicalProperties.line.width = 28000
            serie.smooth = False
            ws.add_chart(graphe, "L5")

            ws.cell(row=rt + 2, column=1,
                    value="Prix moyen pondéré par les volumes : montant total divisé par "
                          "volume total. Une moyenne simple des prix unitaires donnerait "
                          "un résultat faussé par les petites livraisons.")
            ws.cell(row=rt + 2, column=1).font = Font(
                name=FONT, size=8, italic=True, color="6B7280")
            ws.freeze_panes = "A6"
            ws.print_title_rows = "5:5"
            ws.page_setup.orientation = "landscape"

    # ------------------------------------------------ 4. Détail des livraisons
    ws = wb.create_sheet("Détail des livraisons")
    _xl_base(ws)
    _xl_titre(ws, "Détail des livraisons", f"{entete}  |  Base de calcul du présent état")
    hd = ["Date", "Mois", "N° de BL", "Type de pièce", "Site de livraison", "Produit",
          f"Quantité\n({unite})", "Cpt.\nlivraison"]
    lg = [12, 13, 14, 16, 32, 20, 13, 10]
    if argent:
        hd += ["Prix unitaire\n(€/1000 L)", f"{libelle_montant}\n(€)"]
        lg += [15, 15]
    _xl_entetes(ws, 5, hd, lg)
    det = d.sort_values(["Site", "Date", "BL", "Produit"], na_position="last").reset_index(
        drop=True)
    r0 = 6
    for i, row in det.iterrows():
        rr = r0 + i
        if row["Date"] is not None and not (row["Date"] != row["Date"]):
            ws.cell(row=rr, column=1, value=row["Date"]).number_format = "DD/MM/YYYY"
        ws.cell(row=rr, column=2, value=row["Mois"])
        ws.cell(row=rr, column=3, value=str(row["BL"]))
        ws.cell(row=rr, column=4, value=row["Type"])
        ws.cell(row=rr, column=5, value=row["Site"])
        ws.cell(row=rr, column=6, value=row["Produit"])
        ws.cell(row=rr, column=7, value=float(row["Quantite"]))
        ws.cell(row=rr, column=8, value=int(row["Cpt_BL"]))
        if argent:
            ws.cell(row=rr, column=9, value=float(row["PU"]))
            ws.cell(row=rr, column=10, value=f"=G{rr}*I{rr}/1000")
    n = len(det)
    rt = r0 + n
    ws.cell(row=rt, column=1, value="TOTAL")
    for col in ([7, 8] + ([10] if argent else [])):
        L = get_column_letter(col)
        ws.cell(row=rt, column=col, value=f"=SUM({L}{r0}:{L}{rt - 1})")
    _xl_habiller(ws, r0, rt, 10 if argent else 8)
    for rr in range(r0, rt + 1):
        for col in (1, 2, 3, 4, 6, 8):
            ws.cell(row=rr, column=col).alignment = Alignment(horizontal="center")
        ws.cell(row=rr, column=7).number_format = FMT_NB
        ws.cell(row=rr, column=8).number_format = FMT_NB
        if argent:
            ws.cell(row=rr, column=9).number_format = FMT_DEC
            ws.cell(row=rr, column=10).number_format = FMT_EUR
    ws.freeze_panes = "A6"
    ws.auto_filter.ref = f"A5:{get_column_letter(10 if argent else 8)}{rt - 1}"
    ws.print_title_rows = "5:5"
    ws.page_setup.orientation = "landscape"

    # ------------------------------------------------------- 5. Méthodologie
    ws = wb.create_sheet("Méthodologie")
    _xl_base(ws)
    _xl_titre(ws, "Méthodologie et périmètre", "Règles de traitement appliquées aux données sources")
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 95
    s = res.stats
    blocs = [
        ("Objet", f"État des consommations par site de livraison pour le client {client}, "
                  f"sur la période {periode}."
                  + ("" if argent else " Le présent document porte exclusivement sur les "
                                        "volumes livrés et ne comporte aucune valorisation "
                                        "financière.")),
        ("Source des données",
         f"Extraction des mouvements de livraison issue du système de facturation Hympyr "
         f"Énergies : {s['lignes_source']} lignes brutes."),
        ("Lignes exclues du traitement",
         f"{s['lignes_exclues']} lignes ne portant aucune quantité (références de bons de "
         f"commande, lignes d'annulation, commentaires) ont été écartées."
         + (f" {s['lignes_hors_periode']} lignes supplémentaires ont été écartées car "
            f"hors de la période retenue : l'intégralité du présent état — sites, "
            f"produits, mois, détail — porte sur cette seule période."
            if s.get("lignes_hors_periode") else "")),
        ("Lignes retenues",
         f"{s['lignes_retenues']} lignes portant l'un des {npr} produits livrés : "
         + ", ".join(produits) + "."),
        ("Traitement des avoirs",
         f"{s['lignes_avoir']} avoir(s) et {s['lignes_refac']} refacturation(s) sont conservés "
         f"et compensés dans les totaux."
         + (" Le volume net après compensation est strictement identique au volume des "
            "livraisons d'origine : ces régularisations ne portent pas sur les quantités."
            if abs(s["volume"] - s["volume_livraisons"]) < 0.5 else "")),
        ("Comptage des livraisons",
         f"Une livraison = un bon de livraison. Un même BL portant plusieurs produits n'est "
         f"compté qu'une fois ; les avoirs et refacturations ne génèrent pas de livraison "
         f"supplémentaire. Total : {s['nb_livraisons']} livraisons sur {ns} sites."),
        ("Normalisation des sites",
         "Les libellés de sites ont été remis en forme (casse, accents, traits d'union) et, le "
         "cas échéant, regroupés manuellement lorsque plusieurs libellés désignaient le même "
         "site. Aucune modification de périmètre."),
        ("Contrôles de cohérence",
         f"{s['lignes_source']} = {s['lignes_exclues']} lignes exclues + "
         f"{s['lignes_retenues']} lignes retenues. "
         f"{s['lignes_retenues']} = {s['lignes_livraison']} lignes de livraison + "
         f"{s['lignes_avoir']} avoirs + {s['lignes_refac']} refacturations. "
         f"Somme des volumes par site = par mois = par produit = "
         f"{s['volume']:,.0f} {unite}.".replace(",", "\u202f")),
    ]
    if argent:
        blocs.insert(4, ("Base de valorisation",
                         f"{libelle_montant} = quantité × prix unitaire ÷ 1 000, sur la base de "
                         f"la colonne de prix retenue lors de la génération du présent état."))
    r = 6
    for lib, txt in blocs:
        c = ws.cell(row=r, column=1, value=lib)
        c.font = Font(name=FONT, size=9, bold=True, color=XL_FONCE)
        c.alignment = Alignment(vertical="top", wrap_text=True)
        c.fill = PatternFill("solid", fgColor=XL_CLAIR)
        c.border = XL_BORD
        c2 = ws.cell(row=r, column=2, value=txt)
        c2.font = Font(name=FONT, size=9)
        c2.alignment = Alignment(vertical="top", wrap_text=True)
        c2.border = XL_BORD
        ws.row_dimensions[r].height = max(30, 13 * (len(txt) // 95 + 1))
        r += 1

    for sh in wb.worksheets:
        sh.sheet_view.zoomScale = 100

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ============================================================================
# EXPORT PDF
# ============================================================================

import datetime as dt
import html



PDF_VERT, PDF_FONCE = "#1A9E68", "#073D27"
MOIS_C = ["Jan", "Fév", "Mar", "Avr", "Mai", "Juin", "Juil", "Août", "Sep", "Oct", "Nov", "Déc"]
MOIS_L = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
          "septembre", "octobre", "novembre", "décembre"]
PDF_PALETTE = ["#1A9E68", "#0F6B47", "#7FCBAA", "#C3E8D6", "#3FB985", "#05261A", "#A8DCC4",
           "#DDF2E8"]


def _pdf_n(x, d=0):
    return f"{x:,.{d}f}".replace(",", "\u202f").replace(".", ",")


def _pdf_e(x):
    return html.escape(str(x))


PDF_CSS = """
@page { size: A4; margin: 14mm 13mm 16mm 13mm;
  @bottom-left { content: "__PIED__"; font-family: Poppins, Arial, sans-serif;
    font-size: 7pt; color: #9AA0A6; }
  @bottom-right { content: counter(page) " / " counter(pages);
    font-family: Poppins, Arial, sans-serif; font-size: 7pt; color: #9AA0A6; } }
@page :first { margin-top: 0; }
* { box-sizing: border-box; }
body { font-family: Poppins, Arial, sans-serif; font-size: 8.5pt; color: #1F2421; margin: 0; }
.band { background: #073D27; color: #fff; padding: 16mm 13mm 12mm 13mm;
  margin: 0 -13mm 9mm -13mm; }
.band .bn { font-size: 8pt; letter-spacing: 2.5px; color: #7FCBAA; font-weight: 600; }
.band h1 { font-size: 24pt; margin: 5px 0 3px; font-weight: 700; line-height: 1.1; }
.band .sub { font-size: 11pt; color: #C3E8D6; font-weight: 300; }
.band .meta { font-size: 8pt; color: #8FB9A6; margin-top: 9px; }
h2 { font-size: 11pt; color: #073D27; margin: 0 0 6px; font-weight: 600;
  border-left: 3px solid #1A9E68; padding-left: 7px; }
h2.mt { margin-top: 15px; }
.kpis { display: flex; gap: 6px; margin-bottom: 14px; }
.kpi { flex: 1; background: #E8F5EF; border-radius: 4px; padding: 9px 4px; text-align: center; }
.kpi .v { font-size: 15pt; font-weight: 700; color: #073D27; line-height: 1.1; }
.kpi .k { font-size: 7pt; color: #4A6B5C; text-transform: uppercase; letter-spacing: .4px;
  margin-top: 3px; }
table { width: 100%; border-collapse: collapse; }
th { background: #073D27; color: #fff; font-size: 7pt; font-weight: 600; padding: 5px 4px;
  text-align: right; }
th:first-child { text-align: left; }
td { padding: 3.6px 4px; text-align: right; border-bottom: .4px solid #E3E7E5; font-size: 7.6pt; }
td.l { text-align: left; font-weight: 500; }
td.b { font-weight: 600; color: #073D27; }
tr:nth-child(even) td { background: #F7F9F8; }
tr.tot td { background: #E8F5EF !important; font-weight: 700; color: #073D27;
  border-top: 1px solid #1A9E68; border-bottom: none; }
.lg { font-size: 7.5pt; margin-right: 14px; color: #4A5550; }
.lg i { display: inline-block; width: 8px; height: 8px; border-radius: 2px; margin-right: 4px; }
.note { font-size: 7pt; color: #7A8580; margin-top: 5px; line-height: 1.45; }
.brk { page-break-before: always; }
.meth td { border-bottom: .4px solid #E3E7E5; vertical-align: top; line-height: 1.45;
  font-size: 7.6pt; text-align: left; }
.meth td:first-child { width: 30%; font-weight: 600; color: #073D27; }
.meth tr:nth-child(even) td { background: #F7F9F8; }
"""


def _pdf_svg_mois(serie) -> str:
    n = len(serie)
    if not n:
        return ""
    court = [MOIS_C[MOIS.index(nom)] for nom in serie.index]
    W, H, PAD = 700, 215, 26
    mx = serie.max()
    bw = (W - 2 * PAD) / n
    out = [f'<line x1="{PAD}" y1="{H - 30}" x2="{W - PAD}" y2="{H - 30}" stroke="#D0D4D2"/>']
    largeur = min(bw * 0.64, 46)
    for i, v in enumerate(serie):
        h = (v / mx) * (H - 58) if mx else 0
        x = PAD + i * bw + (bw - largeur) / 2
        y = H - 30 - h
        out.append(f'<rect x="{x:.1f}" y="{y:.1f}" width="{largeur:.1f}" '
                   f'height="{h:.1f}" rx="2" fill="{PDF_VERT}"/>')
        out.append(f'<text x="{x + largeur / 2:.1f}" y="{y - 5:.1f}" text-anchor="middle" '
                   f'font-size="8.5" fill="{PDF_FONCE}">{_pdf_n(v / 1000, 1)}</text>')
        out.append(f'<text x="{x + largeur / 2:.1f}" y="{H - 14:.1f}" text-anchor="middle" '
                   f'font-size="9" fill="#6B7280">{court[i]}</text>')
    return (f'<svg viewBox="0 0 {W} {H}" width="100%" font-family="Poppins">'
            + "".join(out) + "</svg>")


def _pdf_svg_produits(tp, produits, coul) -> str:
    total = tp["Volume"].sum()
    segs, x = [], 0.0
    for p in produits:
        part = tp.loc[p, "Volume"] / total if total else 0
        w = part * 700
        segs.append(f'<rect x="{x:.1f}" y="0" width="{w:.1f}" height="26" fill="{coul[p]}"/>')
        if part > 0.05:
            fonce = coul[p] in (PDF_PALETTE[0], PDF_PALETTE[1], PDF_PALETTE[5])
            segs.append(f'<text x="{x + w / 2:.1f}" y="17.5" text-anchor="middle" '
                        f'font-size="10" font-weight="600" '
                        f'fill="{"#fff" if fonce else PDF_FONCE}">{_pdf_n(part * 100, 1)} %</text>')
        x += w
    return ('<svg viewBox="0 0 700 32" width="100%" font-family="Poppins">'
            + "".join(segs) + "</svg>")


def _pdf_svg_prix(t, granularite: str) -> str:
    """Courbe du prix moyen pondéré, avec bande de volume en arrière-plan."""
    import pandas as pd
    W, H = 700, 240
    ML, MR, MT, MB = 54, 24, 18, 34
    pts = t[t["Prix moyen"].notna()]
    if len(pts) < 2:
        return ""

    vals = pts["Prix moyen"].tolist()
    lo, hi = min(vals), max(vals)
    marge = (hi - lo) * 0.18 or (hi * 0.05 or 1)
    lo, hi = lo - marge, hi + marge
    n = len(t)
    pas = (W - ML - MR) / max(n - 1, 1)
    vmax = t["Volume"].max() or 1

    def x(i):
        return ML + i * pas

    def y(v):
        return MT + (hi - v) / (hi - lo) * (H - MT - MB)

    out = []
    for k in range(5):
        v = lo + (hi - lo) * k / 4
        yy = y(v)
        out.append(f'<line x1="{ML}" y1="{yy:.1f}" x2="{W - MR}" y2="{yy:.1f}" '
                   f'stroke="#EDF1EF" stroke-width="1"/>')
        out.append(f'<text x="{ML - 7}" y="{yy + 3:.1f}" text-anchor="end" '
                   f'font-size="7.5" fill="#9AA0A6">{_pdf_n(v, 0)} €</text>')

    largeur = min(pas * 0.5, 22)
    for i, (_, r) in enumerate(t.iterrows()):
        h = (r["Volume"] / vmax) * (H - MT - MB) * 0.32
        if h > 0:
            out.append(f'<rect x="{x(i) - largeur / 2:.1f}" y="{H - MB - h:.1f}" '
                       f'width="{largeur:.1f}" height="{h:.1f}" fill="#E8F5EF"/>')

    coords = [(x(i), y(v)) for i, v in enumerate(t["Prix moyen"]) if pd.notna(v)]
    trace = " ".join(f"{'M' if k == 0 else 'L'}{cx:.1f},{cy:.1f}"
                     for k, (cx, cy) in enumerate(coords))
    out.append(f'<path d="{trace}" fill="none" stroke="{PDF_VERT}" '
               f'stroke-width="2.2" stroke-linejoin="round"/>')

    tv = t["Volume"].sum()
    moy = t["Montant"].sum() / tv * 1000 if tv else 0
    if lo < moy < hi:
        ym = y(moy)
        out.append(f'<line x1="{ML}" y1="{ym:.1f}" x2="{W - MR}" y2="{ym:.1f}" '
                   f'stroke="{PDF_FONCE}" stroke-width="1" stroke-dasharray="4,3"/>')
        out.append(f'<text x="{W - MR}" y="{ym - 5:.1f}" text-anchor="end" '
                   f'font-size="7.5" fill="{PDF_FONCE}" font-weight="600">'
                   f'moyenne {_pdf_n(moy, 0)} €</text>')

    # Étiquettes d'abscisse : une sur `saut`, plus la dernière si elle ne
    # chevauche pas la précédente.
    saut = max(1, n // 12)
    reperes = list(range(0, n, saut))
    if n - 1 - reperes[-1] >= saut * 0.7:
        reperes.append(n - 1)
    else:
        reperes[-1] = n - 1
    for i, (ts, r) in enumerate(t.iterrows()):
        if pd.notna(r["Prix moyen"]):
            out.append(f'<circle cx="{x(i):.1f}" cy="{y(r["Prix moyen"]):.1f}" '
                       f'r="2.6" fill="#fff" stroke="{PDF_VERT}" stroke-width="1.8"/>')
        if i in reperes:
            out.append(f'<text x="{x(i):.1f}" y="{H - 12}" text-anchor="middle" '
                       f'font-size="7.5" fill="#6B7280">'
                       f'{libelle_periode(ts, granularite)}</text>')

    out.append(f'<line x1="{ML}" y1="{H - MB}" x2="{W - MR}" y2="{H - MB}" '
               f'stroke="#D0D4D2"/>')
    return (f'<svg viewBox="0 0 {W} {H}" width="100%" font-family="Poppins">'
            + "".join(out) + "</svg>")


def construire_rapport(res: Resultat, client: str, periode: str, unite: str = "L",
               libelle_montant: str = "Montant",
               prix_options: dict | None = None) -> bytes:
    d = res.donnees
    produits = res.produits
    argent = "Montant" in d.columns
    coul = {p: PDF_PALETTE[i % len(PDF_PALETTE)] for i, p in enumerate(produits)}
    s = res.stats
    vol = s["volume"]
    nb_liv = s["nb_livraisons"]
    ns = s["nb_sites"]
    moy = vol / nb_liv if nb_liv else 0

    tp = par_produit(res)
    ts = par_site(res)
    tm = par_mois(res)
    a_mois = tm.sum() > 0

    # ---------------------------------------------------------------- KPI
    kpis = [(_pdf_n(vol), f"{unite} livrés"), (_pdf_n(nb_liv), "Livraisons"),
            (str(ns), "Sites livrés"), (_pdf_n(moy), f"{unite} / livraison")]
    if argent:
        kpis.append((_pdf_n(s["montant"], 0) + " €", libelle_montant))
    bloc_kpi = "".join(
        f'<div class="kpi"><div class="v">{_pdf_e(v)}</div><div class="k">{_pdf_e(k)}</div></div>'
        for v, k in kpis)

    # ------------------------------------------------------- produits
    th_prod = ["Produit", f"Volume livré ({unite})", "Part du volume", "Sites concernés",
               "Lignes de livraison"]
    if argent:
        th_prod += [f"{libelle_montant} (€)", "Prix moyen (€/1000 L)"]
    lignes_prod = ""
    for p in produits:
        r = tp.loc[p]
        lignes_prod += (f"<tr><td class='l'>{_pdf_e(p)}</td><td>{_pdf_n(r['Volume'])}</td>"
                        f"<td>{_pdf_n(r['Part'] * 100, 1)} %</td><td>{int(r['Sites'])}</td>"
                        f"<td>{int(r['Lignes'])}</td>")
        if argent:
            lignes_prod += (f"<td>{_pdf_n(r['Montant'], 2)} €</td>"
                            f"<td>{_pdf_n(r['Prix moyen'], 2)} €</td>")
        lignes_prod += "</tr>"
    tot_prod = (f"<tr class='tot'><td class='l'>TOTAL</td><td>{_pdf_n(vol)}</td><td>100,0 %</td>"
                f"<td>{ns}</td><td>{int(tp['Lignes'].sum())}</td>")
    if argent:
        tot_prod += (f"<td>{_pdf_n(s['montant'], 2)} €</td>"
                     f"<td>{_pdf_n(s['montant'] / vol * 1000 if vol else 0, 2)} €</td>")
    tot_prod += "</tr>"

    # ------------------------------------------------------- sites
    th_site = (["Site de livraison", "Livr."] + [f"{p} ({unite})" for p in produits]
               + [f"Volume total ({unite})", "Part", f"Moy. / livr. ({unite})"])
    if argent:
        th_site += [f"{libelle_montant} (€)"]
    lignes_site = ""
    for nom, r in ts.iterrows():
        lignes_site += f"<tr><td class='l'>{_pdf_e(nom)}</td><td>{int(r['Livraisons'])}</td>"
        for p in produits:
            lignes_site += f"<td>{_pdf_n(r[p]) if r[p] else '—'}</td>"
        lignes_site += (f"<td class='b'>{_pdf_n(r['Volume'])}</td>"
                        f"<td>{_pdf_n(r['Part'] * 100, 1)} %</td><td>{_pdf_n(r['Moyenne'])}</td>")
        if argent:
            lignes_site += f"<td class='b'>{_pdf_n(r['Montant'], 2)} €</td>"
        lignes_site += "</tr>"
    tot_site = f"<tr class='tot'><td class='l'>TOTAL</td><td>{nb_liv}</td>"
    for p in produits:
        tot_site += f"<td>{_pdf_n(tp.loc[p, 'Volume'])}</td>"
    tot_site += f"<td>{_pdf_n(vol)}</td><td>100,0 %</td><td>{_pdf_n(moy)}</td>"
    if argent:
        tot_site += f"<td>{_pdf_n(s['montant'], 2)} €</td>"
    tot_site += "</tr>"

    # ------------------------------------------------------- saisonnalité
    bloc_mois = ""
    if a_mois:
        i_max = int(tm.values.argmax())
        i_min = int(tm.values.argmin())
        nom_max = str(tm.index[i_max]).lower()
        nom_min = str(tm.index[i_min]).lower()
        bloc_mois = f"""
<h2 class="mt">Saisonnalité des volumes livrés</h2>
{_pdf_svg_mois(tm)}
<div class="note">Volumes en milliers de {unite}, tous produits et tous sites confondus.
Volume mensuel moyen : {_pdf_n(tm.mean())} {unite}.
Mois le plus fort : {nom_max} ({_pdf_n(tm.iloc[i_max])} {unite}) ;
mois le plus faible : {nom_min} ({_pdf_n(tm.iloc[i_min])} {unite}).</div>"""

    # ------------------------------------------------- évolution du prix moyen
    bloc_prix = ""
    if argent and prix_options and prix_options.get("actif"):
        gr = prix_options.get("granularite", "Mois")
        tpx = prix_moyen_periode(res, gr, prix_options.get("debut"),
                                          prix_options.get("fin"),
                                          prix_options.get("produits"))
        svg = _pdf_svg_prix(tpx, gr) if len(tpx) else ""
        if svg:
            valides = tpx[tpx["Prix moyen"].notna()]
            i_hi = valides["Prix moyen"].idxmax()
            i_lo = valides["Prix moyen"].idxmin()
            moy = tpx["Montant"].sum() / tpx["Volume"].sum() * 1000
            ampl = (valides["Prix moyen"].max() - valides["Prix moyen"].min()) / moy * 100
            filtre = prix_options.get("produits")
            perim = (", ".join(_pdf_e(p) for p in filtre) if filtre
                     else "tous produits confondus")
            bloc_prix = f"""
<div class="brk"></div>
<h2>Évolution du prix moyen</h2>
{svg}
<div class="note">Prix moyen pondéré par les volumes, en euros pour 1 000 {_pdf_e(unite)},
{perim}. Les barres claires rappellent le volume livré sur chaque période.
Moyenne de la période : {_pdf_n(moy, 2)} €.
Point haut : {libelle_periode(i_hi, gr)}
({_pdf_n(valides.loc[i_hi, 'Prix moyen'], 2)} €) ;
point bas : {libelle_periode(i_lo, gr)}
({_pdf_n(valides.loc[i_lo, 'Prix moyen'], 2)} €), soit une amplitude de
{_pdf_n(ampl, 1)} % de la moyenne.</div>"""

    # ------------------------------------------------------- méthodologie
    meth = [
        ("Objet", f"État des consommations par site de livraison pour le client {_pdf_e(client)}, "
                  f"sur la période {_pdf_e(periode)}."
                  + ("" if argent else " Le présent document porte exclusivement sur les "
                                        "volumes livrés et ne comporte aucune valorisation "
                                        "financière.")),
        ("Source", f"Mouvements de livraison extraits du système de facturation Hympyr "
                   f"Énergies ({s['lignes_source']} lignes brutes)."),
        ("Lignes exclues", f"{s['lignes_exclues']} lignes ne portant aucune quantité "
                           f"(références de bons de commande, annulations, commentaires)."
                           + (f" {s['lignes_hors_periode']} lignes écartées car hors de la "
                              f"période retenue : l'intégralité du présent état porte sur "
                              f"cette seule période."
                              if s.get("lignes_hors_periode") else "")),
        ("Lignes retenues", f"{s['lignes_retenues']} lignes portant l'un des "
                            f"{len(produits)} produits livrés : " + ", ".join(
                                _pdf_e(p) for p in produits) + "."),
        ("Avoirs et régularisations",
         f"{s['lignes_avoir']} avoir(s) et {s['lignes_refac']} refacturation(s) compensés dans "
         f"les totaux."
         + (" Ces régularisations portant uniquement sur le prix, le volume net est "
            "strictement identique au volume livré d'origine."
            if abs(s["volume"] - s["volume_livraisons"]) < 0.5 else "")),
        ("Comptage des livraisons",
         "Une livraison correspond à un bon de livraison. Un BL portant plusieurs produits "
         "n'est compté qu'une fois ; les avoirs ne génèrent pas de livraison supplémentaire."),
        ("Contrôles",
         f"{s['lignes_source']} = {s['lignes_exclues']} lignes exclues + "
         f"{s['lignes_retenues']} lignes retenues. La somme des volumes par site, par mois et "
         f"par produit s'établit dans chaque cas à {_pdf_n(vol)} {unite}."),
    ]
    if argent:
        meth.insert(4, ("Valorisation",
                        f"{_pdf_e(libelle_montant)} = quantité livrée × prix unitaire ÷ 1 000."))
    lignes_meth = "".join(f"<tr><td>{_pdf_e(k)}</td><td>{v}</td></tr>" for k, v in meth)

    pied = f"Hympyr Énergies — État des consommations — {client}"
    css = PDF_CSS.replace("__PIED__", pied.replace('"', "'"))
    legende = "".join(f'<span class="lg"><i style="background:{coul[p]}"></i>{_pdf_e(p)}</span>'
                      for p in produits)
    aujourdhui = dt.date.today().strftime("%d/%m/%Y")

    doc = f"""<html><head><meta charset="utf-8"><style>{css}</style></head><body>
<div class="band">
  <div class="bn">HYMPYR ÉNERGIES</div>
  <h1>État des consommations</h1>
  <div class="sub">{_pdf_e(client)} — Volumes livrés par site</div>
  <div class="meta">Période : {_pdf_e(periode)} · {ns} sites livrés ·
  Document établi le {aujourdhui}</div>
</div>

<div class="kpis">{bloc_kpi}</div>

<h2>Répartition par produit</h2>
{_pdf_svg_produits(tp, produits, coul)}
<div style="margin:2px 0 7px">{legende}</div>
<table>
<tr>{''.join(f'<th>{_pdf_e(t)}</th>' for t in th_prod)}</tr>
{lignes_prod}{tot_prod}
</table>
{bloc_mois}

<div class="brk"></div>
<h2>Consommations par site de livraison</h2>
<table>
<tr>{''.join(f'<th>{_pdf_e(t)}</th>' for t in th_site)}</tr>
{lignes_site}{tot_site}
</table>
<div class="note">Sites classés par volume total décroissant. Volumes nets des avoirs et
régularisations de la période.</div>
{bloc_prix}

<div class="brk"></div>
<h2>Méthodologie et périmètre</h2>
<table class="meth">{lignes_meth}</table>
<div class="note" style="margin-top:14px">Hympyr Énergies — Distribution de carburants et
combustibles depuis 1960. Ce document a une valeur d'information et ne se substitue pas aux
factures émises. Pour toute question relative au présent état, votre interlocuteur commercial
Hympyr se tient à votre disposition.</div>
</body></html>"""

    moteur = _weasyprint()
    if moteur is None:
        raise RuntimeError(
            "WeasyPrint indisponible sur cet environnement : "
            f"{_WEASY_ERREUR}. Le classeur Excel reste générable.")
    return moteur(string=doc).write_pdf()

# ============================================================================
# INTERFACE STREAMLIT
# ============================================================================


st.set_page_config(page_title="Reporting consommations — Hympyr Énergies",
                   page_icon="⛽", layout="wide")

st.markdown(f"""
<style>
  .stApp {{ background: #FBFCFB; }}
  h1, h2, h3 {{ color: {FONCE}; }}
  div[data-testid="stMetricValue"] {{ color: {FONCE}; font-weight: 700; }}
  .hym-band {{ background: {FONCE}; color: #fff; padding: 18px 22px; border-radius: 8px;
    margin-bottom: 18px; }}
  .hym-band .bn {{ font-size: 11px; letter-spacing: 2.5px; color: #7FCBAA;
    font-weight: 600; }}
  .hym-band h1 {{ color: #fff !important; font-size: 26px; margin: 4px 0 2px; }}
  .hym-band p {{ color: #C3E8D6; margin: 0; font-size: 14px; }}
  .stDownloadButton button {{ background: {VERT}; color: #fff; border: none;
    font-weight: 600; width: 100%; }}
  .stDownloadButton button:hover {{ background: {FONCE}; color: #fff; }}
</style>
<div class="hym-band">
  <div class="bn">HYMPYR ÉNERGIES</div>
  <h1>Reporting des consommations clients</h1>
  <p>Déposez l'export de livraisons, récupérez l'état par site en PDF et en Excel.</p>
</div>
""", unsafe_allow_html=True)


def nom_fichier(client: str, periode: str, ext: str) -> str:
    base = unicodedata.normalize("NFD", f"Etat_consommations_{periode}_{client}")
    base = "".join(c for c in base if unicodedata.category(c) != "Mn")
    base = re.sub(r"[^A-Za-z0-9]+", "_", base).strip("_")
    return f"{base}.{ext}"


# =====================================================================
# 1. Fichier
# =====================================================================
st.subheader("1 · Fichier source")
fichier = st.file_uploader(
    "Export des livraisons (.xlsx, .xls ou .csv)",
    type=["xlsx", "xls", "csv", "txt"],
    help="Le fichier doit contenir au minimum une colonne site de livraison, une colonne "
         "désignation produit et une colonne quantité.")

if fichier is None:
    st.info("En attente d'un fichier. L'outil détecte automatiquement les colonnes, écarte "
            "les lignes sans quantité (références de bons de commande, annulations), "
            "compense les avoirs et signale les anomalies avant génération.")
    st.stop()

try:
    df_source, feuilles = charger_fichier(fichier)
    if len(feuilles) > 1:
        feuille = st.selectbox("Feuille à traiter", feuilles, index=0)
        fichier.seek(0)
        df_source, _ = charger_fichier(fichier, feuille)
except Exception as exc:
    st.error(f"Lecture impossible : {exc}")
    st.stop()

df_source = df_source.dropna(how="all").dropna(axis=1, how="all")
st.success(f"{len(df_source)} lignes et {len(df_source.columns)} colonnes chargées.")
with st.expander("Aperçu des données brutes"):
    st.dataframe(df_source.head(30), use_container_width=True)

# =====================================================================
# 2. Colonnes
# =====================================================================
st.subheader("2 · Colonnes")
auto = deviner_colonnes(df_source)
cols = list(df_source.columns)
manquants = [r for r in ("site", "designation", "quantite") if not auto.get(r)]
if manquants:
    st.warning("Colonnes non reconnues automatiquement : " + ", ".join(manquants)
               + ". Sélectionnez-les ci-dessous.")
else:
    st.caption("Colonnes détectées automatiquement. Corrigez si nécessaire.")


def _sel(label, role, obligatoire=True):
    options = ([] if obligatoire else ["— aucune —"]) + cols
    val = auto.get(role)
    idx = options.index(val) if val in options else 0
    return st.selectbox(label, options, index=idx, key=f"col_{role}")


with st.expander("Correspondance des colonnes", expanded=bool(manquants)):
    c1, c2, c3 = st.columns(3)
    with c1:
        c_site = _sel("Site de livraison *", "site")
        c_des = _sel("Désignation produit *", "designation")
    with c2:
        c_qte = _sel("Quantité *", "quantite")
        c_date = _sel("Date", "date", obligatoire=False)
    with c3:
        c_bl = _sel("N° de bon de livraison", "bl", obligatoire=False)

mapping = {
    "site": c_site, "designation": c_des, "quantite": c_qte,
    "date": None if c_date == "— aucune —" else c_date,
    "bl": None if c_bl == "— aucune —" else c_bl,
    "prix_candidats": [c for c in auto.get("prix_candidats", [])
                       if c not in (c_site, c_des, c_qte, c_date, c_bl)],
}

# =====================================================================
# 3. Lignes à traiter
# =====================================================================
st.subheader("3 · Lignes à traiter")
inv = inventaire_designations(df_source, c_des, c_qte)
st.caption("Les désignations sans quantité (bons de commande, annulations, commentaires) sont "
           "décochées automatiquement. Décochez ou recochez selon vos besoins.")
inv_edit = st.data_editor(
    inv, use_container_width=True, hide_index=True, height=240,
    column_config={
        "À conserver": st.column_config.CheckboxColumn("Traiter", width="small"),
        "Désignation": st.column_config.TextColumn(disabled=True),
        "Lignes": st.column_config.NumberColumn(disabled=True, width="small"),
        "Volume": st.column_config.NumberColumn(disabled=True, format="%.0f"),
    }, key="inv")
retenues = inv_edit.loc[inv_edit["À conserver"], "Désignation"].tolist()
if not retenues:
    st.error("Aucune désignation sélectionnée.")
    st.stop()

# =====================================================================
# 4. Options
# =====================================================================
st.subheader("4 · Paramètres du rapport")
res_brut = preparer(df_source, mapping, retenues)
dmin, dmax = res_brut.stats["date_min"], res_brut.stats["date_max"]
annees = sorted(res_brut.donnees["Date"].dropna().dt.year.unique().tolist())

o1, o2 = st.columns([2, 3])
with o1:
    client = st.text_input("Nom du client", value="", placeholder="DÉPARTEMENT DU TARN")
    unite = st.text_input("Unité des quantités", value="L")
with o2:
    modes = ["Toute la période du fichier", "Une année civile", "Dates personnalisées"]
    mode = st.radio("Périmètre temporel", modes,
                    index=1 if len(annees) > 1 else 0, horizontal=True,
                    disabled=not annees,
                    help="Ce filtre s'applique à l'intégralité du rapport : sites, "
                         "produits, vue mensuelle, courbe de prix et détail des "
                         "livraisons.")
    debut = fin = None
    if not annees:
        st.caption("Aucune date exploitable : le filtre de période est indisponible.")
    elif mode == modes[1]:
        annee = st.selectbox("Année analysée", annees, index=len(annees) - 1)
        debut, fin = dt.date(annee, 1, 1), dt.date(annee, 12, 31)
    elif mode == modes[2]:
        bornes = st.date_input("Dates analysées", value=(dmin.date(), dmax.date()),
                               min_value=dmin.date(), max_value=dmax.date(),
                               format="DD/MM/YYYY")
        if isinstance(bornes, (tuple, list)) and len(bornes) == 2:
            debut, fin = bornes
        else:
            st.info("Sélectionnez une date de fin pour appliquer le filtre.")
            st.stop()

if debut and fin:
    defaut_periode = f"{debut.strftime('%d/%m/%Y')} – {fin.strftime('%d/%m/%Y')}"
    if (debut.month, debut.day, fin.month, fin.day) == (1, 1, 12, 31) \
            and debut.year == fin.year:
        defaut_periode = f"Année {debut.year}"
elif pd.notna(dmin):
    defaut_periode = f"{dmin.strftime('%d/%m/%Y')} – {dmax.strftime('%d/%m/%Y')}"
else:
    defaut_periode = str(dt.date.today().year)
periode = st.text_input("Intitulé de la période affiché sur les documents",
                        value=defaut_periode)

res0 = preparer(df_source, mapping, retenues, debut=debut, fin=fin)
if not len(res0.donnees):
    st.error("Aucune livraison sur la période retenue. Élargissez le périmètre.")
    st.stop()
if res0.stats["lignes_hors_periode"]:
    st.caption(f"Filtre actif : {res0.stats['lignes_hors_periode']} ligne(s) écartée(s) "
               f"hors période, {res0.stats['lignes_retenues']} conservée(s).")

with st.expander("Regrouper ou renommer les sites et les produits"):
    st.caption("Modifiez la colonne « Libellé retenu » pour corriger un libellé ou fusionner "
               "plusieurs lignes sous un même nom.")
    g1, g2 = st.columns([3, 2])
    with g1:
        t_sites = tableau_sites(res0)
        t_sites_edit = st.data_editor(
            t_sites, use_container_width=True, hide_index=True, height=260,
            column_config={
                "Libellé source": st.column_config.TextColumn(disabled=True),
                "Libellé retenu": st.column_config.TextColumn(width="medium"),
                "Volume": st.column_config.NumberColumn(disabled=True, format="%.0f"),
                "Lignes": st.column_config.NumberColumn(disabled=True, width="small"),
            }, key="sites")
    with g2:
        t_prod = pd.DataFrame({"Libellé source": res0.produits,
                               "Libellé retenu": res0.produits})
        t_prod_edit = st.data_editor(
            t_prod, use_container_width=True, hide_index=True, height=260,
            column_config={"Libellé source": st.column_config.TextColumn(disabled=True)},
            key="produits")

regroupement = {}
for _, r in t_sites_edit.iterrows():
    cible = str(r["Libellé retenu"]).strip()
    if cible:
        regroupement[normaliser_site(r["Libellé source"])] = cible
renommage = {str(r["Libellé source"]): str(r["Libellé retenu"]).strip()
             for _, r in t_prod_edit.iterrows() if str(r["Libellé retenu"]).strip()}

# --- Valorisation (désactivée par défaut)
col_prix, diviseur, libelle_montant = None, 1.0, "Montant"
prix_options = {"actif": False}
cands = mapping["prix_candidats"]
with st.expander("Valorisation financière (optionnelle)"):
    if not cands:
        st.caption("Aucune colonne de prix exploitable détectée dans le fichier.")
    else:
        actif = st.checkbox(
            "Inclure les montants dans le rapport", value=False,
            help="Par défaut, l'état porte uniquement sur les volumes. N'activez la "
                 "valorisation qu'après avoir vérifié la base HT/TTC de la colonne de prix.")
        if actif:
            p1, p2, p3 = st.columns(3)
            with p1:
                col_prix = st.selectbox("Colonne de prix", cands)
            with p2:
                base = st.radio("Base de la colonne", ["HT", "TTC (TVA 20 %)"])
                diviseur = 1.2 if base.startswith("TTC") else 1.0
            with p3:
                libelle_montant = st.text_input("Intitulé de la colonne montant",
                                                value="Montant HT")
            st.caption("Les prix sont interprétés en euros pour 1 000 unités. "
                       "Vérifiez qu'aucune colonne de coût d'achat interne n'est diffusée "
                       "au client.")

            st.divider()
            prix_options["actif"] = st.checkbox(
                "Ajouter la courbe d'évolution du prix moyen", value=True,
                help="Une page supplémentaire dans le PDF et un onglet dans le classeur.")
            if prix_options["actif"]:
                prix_options["debut"] = prix_options["fin"] = None
                q1, q2 = st.columns([1, 2])
                with q1:
                    prix_options["granularite"] = st.radio(
                        "Granularité", ["Mois", "Trimestre", "Semaine"], index=0)
                with q2:
                    prix_options["produits"] = st.multiselect(
                        "Produits analysés",
                        [renommage.get(p, p) for p in res0.produits], default=[],
                        help="Vide = tous les produits. Restreindre à un seul produit "
                             "donne une courbe lisible ; mélanger plusieurs énergies fait "
                             "surtout apparaître un effet de mix, pas un effet prix.")
                    st.caption("La courbe couvre le périmètre temporel défini à l'étape 4, "
                               "comme le reste du rapport.")

res = preparer(df_source, mapping, retenues, regroupement, renommage,
                       col_prix, diviseur, debut, fin)

# =====================================================================
# 5. Contrôles
# =====================================================================
st.subheader("5 · Contrôles qualité")
alertes = diagnostiquer(df_source, mapping, res)
for a in alertes:
    texte = f"**{a['titre']}** — {a['detail']}"
    if a["niveau"] == "alerte":
        st.warning(texte)
    elif a["niveau"] == "ok":
        st.success(texte)
    else:
        st.info(texte)

# =====================================================================
# 6. Aperçu
# =====================================================================
st.subheader("6 · Aperçu")
s = res.stats
k = st.columns(5)
k[0].metric(f"{unite} livrés", f"{s['volume']:,.0f}".replace(",", "\u202f"))
k[1].metric("Livraisons", f"{s['nb_livraisons']:,}".replace(",", "\u202f"))
k[2].metric("Sites livrés", s["nb_sites"])
moy = s["volume"] / s["nb_livraisons"] if s["nb_livraisons"] else 0
k[3].metric(f"{unite} / livraison", f"{moy:,.0f}".replace(",", "\u202f"))
k[4].metric("Montant" if "montant" in s else "Produits",
            f"{s['montant']:,.0f} €".replace(",", "\u202f") if "montant" in s
            else s["nb_produits"])

t1, t2, t3, t4 = st.tabs(["Par site", "Par produit", "Par mois", "Prix moyen"])
with t1:
    tsite = par_site(res)
    fmt = {c: "{:,.0f}" for c in tsite.columns if c != "Part"}
    fmt["Part"] = "{:.1%}"
    st.dataframe(tsite.style.format(fmt, na_rep="—"),
                 use_container_width=True, height=420)
with t2:
    st.dataframe(par_produit(res), use_container_width=True)
with t3:
    tm = par_mois(res)
    if tm.sum():
        st.bar_chart(tm, color=VERT, height=280)
    else:
        st.caption("Aucune date exploitable : la vue mensuelle est désactivée.")
with t4:
    if not prix_options.get("actif"):
        st.caption("Activez la valorisation financière et la courbe de prix dans les "
                   "paramètres pour afficher cette vue.")
    else:
        tpx = prix_moyen_periode(
            res, prix_options.get("granularite", "Mois"), prix_options.get("debut"),
            prix_options.get("fin"), prix_options.get("produits"))
        if tpx.empty or tpx["Prix moyen"].notna().sum() < 2:
            st.warning("Pas assez de périodes valorisées pour tracer une courbe.")
        else:
            vals = tpx["Prix moyen"].dropna()
            moy = tpx["Montant"].sum() / tpx["Volume"].sum() * 1000
            c = st.columns(4)
            c[0].metric("Prix moyen pondéré", f"{moy:,.2f} €".replace(",", "\u202f"))
            c[1].metric("Point haut", f"{vals.max():,.2f} €".replace(",", "\u202f"))
            c[2].metric("Point bas", f"{vals.min():,.2f} €".replace(",", "\u202f"))
            c[3].metric("Amplitude", f"{(vals.max() - vals.min()) / moy * 100:.1f} %")
            courbe = tpx["Prix moyen"].copy()
            courbe.index = [libelle_periode(i, prix_options["granularite"])
                            for i in courbe.index]
            st.line_chart(courbe, color=VERT, height=300)
            st.caption("Prix moyen pondéré par les volumes, en euros pour 1 000 "
                       f"{unite}. Une moyenne simple des prix unitaires donnerait un "
                       "résultat faussé par les petites livraisons.")

# =====================================================================
# 7. Génération
# =====================================================================
st.subheader("7 · Génération des documents")
if not client.strip():
    st.info("Renseignez le nom du client pour générer les documents.")
    st.stop()

if st.button("Générer le PDF et le classeur Excel", type="primary"):
    with st.spinner("Génération en cours…"):
        try:
            pdf = construire_rapport(res, client.strip(), periode, unite,
                                        libelle_montant, prix_options)
            xls = construire_classeur(res, client.strip(), periode, unite,
                                          libelle_montant, prix_options)
        except Exception as exc:
            st.error(f"Échec de la génération : {exc}")
            st.stop()
    st.session_state["pdf"] = pdf
    st.session_state["xls"] = xls
    st.session_state["nom"] = (nom_fichier(client, periode.replace(" ", ""), "pdf"),
                               nom_fichier(client, periode.replace(" ", ""), "xlsx"))

if "pdf" in st.session_state:
    n_pdf, n_xls = st.session_state["nom"]
    d1, d2 = st.columns(2)
    d1.download_button("Télécharger le rapport PDF (client)", st.session_state["pdf"],
                       file_name=n_pdf, mime="application/pdf")
    d2.download_button("Télécharger le classeur Excel (interne)", st.session_state["xls"],
                       file_name=n_xls,
                       mime="application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet")
    st.caption("Le classeur Excel est entièrement formulé : les onglets de synthèse "
               "s'actualisent si vous modifiez l'onglet « Détail des livraisons ».")
