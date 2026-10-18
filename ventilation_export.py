# -*- coding: utf-8 -*-
"""Ventilation EXPORT (DUM) : remplit automatiquement le modèle VENTI_EXPORT à partir des BL / factures.

Principe (identique au travail manuel dans le modèle) :
  1. chaque référence est cherchée dans la feuille « BASE DONNER » du modèle -> désignation précise (ex. « Conducteurs électriques »,
     « Parties d'armoires électriques ») ;
  2. la désignation donne le NGP via la feuille « NGPN » ;
  3. les lignes d'un même NGP forment un article de DUM (feuille ART1…ART20) : quantité, poids net, valeur devise ;
  4. les totaux et le destinataire / N° DUM alimentent la feuille « VENTILATION » ; les formules du modèle (VAL DHS, ANNEXE…) restent vivantes.
Les références absentes de BASE DONNER sont regroupées dans un article « À CLASSER » (jamais perdues, jamais devinées).
Le modèle (.xltx) est dans templates/VENTI_EXPORT.xltx ; l'Admin peut le remplacer depuis l'application."""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

import cleaning

DEFAULT_TEMPLATE = Path(__file__).resolve().parent / "templates" / "VENTI_EXPORT.xltx"
UNCLASSIFIED = "À CLASSER (référence introuvable dans BASE DONNER)"
MAX_ARTICLES, FIRST_LINE_ROW, LAST_LINE_ROW = 20, 3, 472        # feuilles ART1..ART20 : lignes 3 à 472
ORIGIN_NAME, ORIGIN_CODE = "MAROC", "MA"


def override_path() -> Path:
    import database
    return database.DATA_DIR / "templates" / "VENTI_EXPORT.xltx"


def get_template_bytes() -> bytes | None:
    """Modèle téléversé par l'Admin s'il existe, sinon le modèle livré avec l'application."""
    for p in (override_path(), DEFAULT_TEMPLATE):
        try:
            if p.is_file():
                return p.read_bytes()
        except OSError:
            continue
    return None


def save_template_override(data: bytes):
    p = override_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_bytes(data)


@dataclass
class Reference:
    base: dict[str, dict]                 # REF normalisée -> {designation, annexe, ngp}
    ngpn: dict[str, dict]                 # désignation (casefold) -> {designation, ngp, rubrique}
    ngp_to_name: dict[str, str]           # NGP -> désignation du modèle (1re ligne NGPN avec ce NGP)
    destinations: dict[str, str]          # destinataire (nom nettoyé) -> préfixe N° DUM
    dest_raw: dict[str, str] = field(default_factory=dict)   # nom nettoyé -> nom EXACT de la cellule ATTELIER (espaces finaux compris)


def norm_ref(x) -> str:
    return re.sub(r"\s+", "", str(x if x is not None else "")).upper()


def _ngp(v) -> str:
    if v is None or v == "":
        return ""
    try:
        return str(int(float(v)))
    except (TypeError, ValueError):
        return str(v).strip()


def load_reference(template_bytes: bytes) -> Reference:
    """Lit BASE DONNER, NGPN et ATTELIER dans le modèle (lecture seule, valeurs)."""
    wb = load_workbook(io.BytesIO(template_bytes), read_only=True, data_only=True)
    base: dict[str, dict] = {}
    for row in wb["BASE DONNER"].iter_rows(min_row=2, values_only=True):
        ref = norm_ref(row[2]) if len(row) > 2 else ""
        if ref and ref not in base:
            base[ref] = {"designation": str(row[9] or "").strip(), "annexe": str(row[10] or "").strip(), "ngp": _ngp(row[4])}
    ngpn: dict[str, dict] = {}
    ngp_to_name: dict[str, str] = {}
    for row in wb["NGPN"].iter_rows(min_row=2, values_only=True):
        if row[0] and row[1]:
            name, ngp = str(row[0]).strip(), _ngp(row[1])
            ngpn[name.casefold()] = {"designation": name, "ngp": ngp, "rubrique": str(row[2] or name).strip()}
            ngp_to_name.setdefault(ngp, name)
    destinations, dest_raw = {}, {}
    for r in wb["ATTELIER"].iter_rows(min_row=1, values_only=True):
        if r[0]:
            clean = str(r[0]).strip()
            destinations[clean] = str(r[1] or "").strip()
            dest_raw[clean] = str(r[0])          # « Allemagne » est écrit « Allemagne␠ » dans le modèle : Excel exige le nom exact
    wb.close()
    return Reference(base, ngpn, ngp_to_name, destinations, dest_raw)


def classify(ref: str, reference: Reference) -> dict | None:
    """Article de DUM d'une référence : {'ngp', 'name', 'rubrique'} ou None si introuvable."""
    info = reference.base.get(norm_ref(ref))
    if not info:
        return None
    j = info["designation"]
    if "conducteur" in j.casefold():                       # « CONDUCTEUR ( NB ) vérifie » -> « Conducteurs électriques »
        j = "Conducteurs électriques"
    row = reference.ngpn.get(j.casefold())
    if row:
        return {"ngp": row["ngp"], "name": row["designation"], "rubrique": row["rubrique"]}
    if info["ngp"]:
        return {"ngp": info["ngp"], "name": j or info["annexe"], "rubrique": info["annexe"] or j}
    return None


CONDUCTOR_RE = re.compile(r"(?:^|[\s(])H_[A-Z]|HARNAIS|FAISCEAU|C[ÂA]BL(?:AGE|E)\b|CONDUCTEUR", re.I)
CABINET_RE = re.compile(r"ARMOIRE|CABINET|COFFRET|PUPITRE|\bBAIE\b", re.I)


def guess_from_description(desc: str, reference: Reference) -> dict | None:
    """Suggestion (à vérifier) pour une référence absente de BASE DONNER, d'après sa description."""
    text = str(desc or "")
    for rx, name in ((CABINET_RE, "armoires électriques"), (CONDUCTOR_RE, "conducteurs électriques")):
        if rx.search(text) and name in reference.ngpn:
            row = reference.ngpn[name]
            return {"ngp": row["ngp"], "name": row["designation"], "rubrique": row["rubrique"]}
    return None


@dataclass
class Article:
    ngp: str
    designation: str                  # libellé DUM (rubrique annexe)
    template_name: str                # désignation à poser en A2 de la feuille ART (pour le VLOOKUP du NGP)
    lines: list[dict] = field(default_factory=list)

    @property
    def suggested(self) -> int: return sum(1 for l in self.lines if l.get("suggested"))
    @property
    def qty(self) -> float: return float(sum(l["qty"] for l in self.lines))
    @property
    def net(self) -> float: return float(sum(l["net"] for l in self.lines))
    @property
    def value(self) -> float: return float(sum(l["value"] for l in self.lines))


@dataclass
class Ventilation:
    articles: list[Article]
    totals: dict
    missing: list[dict]
    notes: list[str]

    def summary_df(self, rate: float = 0.0) -> pd.DataFrame:
        return pd.DataFrame([{
            "N°": i, "Désignation": a.designation, "NGP": a.ngp or "—", "Lignes": len(a.lines),
            "Quantité": round(a.qty, 3), "Poids net (kg)": round(a.net, 3), "Valeur devise": round(a.value, 2),
            "Valeur DH": int(-(-a.value * rate // 1)) if rate else None, "Origine": ORIGIN_CODE,
            "Source": ("BASE DONNER" if not a.suggested else f"BASE DONNER + {a.suggested} ligne(s) suggérée(s)") if a.ngp else "À classer",
        } for i, a in enumerate(self.articles, start=1)])


def build_ventilation(bl_df: pd.DataFrame, reference: Reference, guess: bool = True, manual: dict | None = None) -> Ventilation:
    """Regroupe les lignes d'export en articles de DUM. `bl_df` : tableau BL References (colonnes standard de l'export)."""
    num = lambda s: pd.to_numeric(s, errors="coerce")  # noqa: E731
    df = bl_df.copy().reset_index(drop=True)
    df["_bl"] = df["BL Number (Bon de Livraison)"].astype(str)
    inv_col = df["Invoice Number (Facture)"] if "Invoice Number (Facture)" in df else pd.Series([""] * len(df))
    df["_inv"] = inv_col.map(lambda v: str(v).strip() if str(v).strip() not in ("", "nan", "None") else "(sans facture)")
    df["_val"] = num(df["Prix unitaire total (Hors-taxe)"]).fillna(0.0)
    df["_qty"] = num(df["Quantity"]).fillna(0.0)
    df["_netbl"] = num(df["Total Net Weight of the BL"])
    # poids net du BL réparti sur ses lignes au prorata de la valeur (somme exacte)
    df["_net"] = 0.0
    for _bl, idx in df.groupby("_bl", sort=False).groups.items():
        idx = list(idx)
        total = float(df.loc[idx, "_netbl"].sum(skipna=True))
        if total > 0:
            for i, v in zip(idx, cleaning._allocate(total, df.loc[idx, "_val"].abs().tolist() or [1.0])):
                df.at[i, "_net"] = v
    if manual:                                                    # classement choisi par l'utilisateur : article -> désignation, facture -> article
        return _build_manual(df, reference, manual)
    articles: dict[str, Article] = {}
    missing: dict[str, dict] = {}
    for _, r in df.iterrows():
        ref = str(r["Reference"]).strip()
        art = classify(ref, reference)
        suggested = False
        if not art and guess:
            art = guess_from_description(r.get("Description", ""), reference)
            suggested = art is not None
        key = art["ngp"] if art else "?"
        if key not in articles:
            articles[key] = Article(art["ngp"], art["rubrique"], art["name"]) if art else Article("", UNCLASSIFIED, UNCLASSIFIED)
        articles[key].lines.append({"ref": ref, "desc": str(r.get("Description", "") or ""), "qty": float(r["_qty"]), "inv": r["_inv"],
                                    "net": float(r["_net"]), "value": float(r["_val"]), "bl": r["_bl"],
                                    "suggested": suggested})
        if not art or suggested:
            m = missing.setdefault(ref, {"Référence": ref, "Description": str(r.get("Description", "") or "")[:60], "Quantité": 0.0,
                                         "Valeur": 0.0, "Classement": (art["name"] + " (suggéré)") if suggested else "À classer"})
            m["Quantité"] += float(r["_qty"])
            m["Valeur"] += float(r["_val"])
    ordered = sorted(articles.values(), key=lambda a: (a.ngp == "", -a.value))
    totals = {"colis": float(num(df["Number of Packages (Nombre de Colis)"]).sum()),
              "brut": float(num(df["Total Gross Weight of the BL"]).sum()), "net": float(df["_netbl"].sum(skipna=True)),
              "valeur": float(df["_val"].sum()), "quantite": float(df["_qty"].sum()), "lignes": len(df)}
    notes = []
    if missing:
        n_sug = sum(1 for m in missing.values() if m["Classement"] != "À classer")
        notes.append(f"{len(missing)} référence(s) absente(s) de BASE DONNER"
                     + (f" ({n_sug} classée(s) par suggestion d'après la description : à vérifier)" if n_sug else "")
                     + " — ajoutez-les à la base du modèle pour un classement sûr.")
    if len(ordered) > MAX_ARTICLES:
        notes.append(f"Plus de {MAX_ARTICLES} articles : le modèle n'en accepte que {MAX_ARTICLES}.")
    cur = df["_currency"].dropna().astype(str).unique().tolist() if "_currency" in df else []
    if len([c for c in cur if c]) > 1:
        notes.append(f"Plusieurs devises dans le dossier ({', '.join(c for c in cur if c)}) : la ventilation additionne les montants tels quels.")
    return Ventilation(ordered, totals, list(missing.values()), notes)


def _totals(df: pd.DataFrame) -> dict:
    num = lambda s: pd.to_numeric(s, errors="coerce")  # noqa: E731
    return {"colis": float(num(df["Number of Packages (Nombre de Colis)"]).sum()),
            "brut": float(num(df["Total Gross Weight of the BL"]).sum()), "net": float(df["_netbl"].sum(skipna=True)),
            "valeur": float(df["_val"].sum()), "quantite": float(df["_qty"].sum()), "lignes": len(df)}


def _build_manual(df: pd.DataFrame, reference: Reference, manual: dict) -> Ventilation:
    """manual = {'articles': [désignation NGPN article 1, 2, …], 'assign': {n° facture: n° d'article (1..n)}}"""
    names = manual["articles"]
    arts: list[Article] = []
    for n in names:
        row = reference.ngpn.get(str(n).casefold())
        arts.append(Article(row["ngp"], row["rubrique"], row["designation"]) if row else Article("", str(n), str(n)))
    assign = manual.get("assign", {})
    notes: list[str] = []
    for _, r in df.iterrows():
        k = int(assign.get(r["_inv"], 1)) - 1
        k = min(max(k, 0), len(arts) - 1)
        arts[k].lines.append({"ref": str(r["Reference"]).strip(), "desc": str(r.get("Description", "") or ""), "qty": float(r["_qty"]),
                              "inv": r["_inv"], "net": float(r["_net"]), "value": float(r["_val"]), "bl": r["_bl"], "suggested": False})
    empty = [i + 1 for i, a in enumerate(arts) if not a.lines]
    if empty:
        notes.append(f"Article(s) {', '.join(map(str, empty))} sans aucune facture : ils seront laissés vides.")
    kept = [a for a in arts if a.lines]
    return Ventilation(kept, _totals(df), [], notes)


def invoice_table(vent: Ventilation) -> pd.DataFrame:
    """Une ligne par facture avec l'article proposé (celui qui porte la plus grande valeur de cette facture)."""
    rows: dict[str, dict] = {}
    for idx, art in enumerate(vent.articles, start=1):
        for l in art.lines:
            r = rows.setdefault(l["inv"], {"Facture": l["inv"], "BL": set(), "Lignes": 0, "Quantité": 0.0, "Valeur": 0.0, "_best": {}})
            r["BL"].add(l["bl"])
            r["Lignes"] += 1
            r["Quantité"] += l["qty"]
            r["Valeur"] += l["value"]
            r["_best"][idx] = r["_best"].get(idx, 0.0) + l["value"]
    out = []
    for r in rows.values():
        best = max(r["_best"], key=r["_best"].get) if r["_best"] else 1
        out.append({"Facture": r["Facture"], "BL": ", ".join(sorted(r["BL"])), "Lignes": r["Lignes"],
                    "Quantité": round(r["Quantité"], 3), "Valeur": round(r["Valeur"], 2), "Article": best})
    return pd.DataFrame(out)


def fill_template(template_bytes: bytes, vent: Ventilation, destinataire: str, rate: float, fret: float | None = None) -> bytes:
    """Remplit une copie du modèle (formules conservées : VAL DHS, ANNEXE, N° DUM… se calculent à l'ouverture dans Excel)."""
    wb = load_workbook(io.BytesIO(template_bytes), keep_links=False)       # keep_links=False : supprime le lien externe du taux
    v = wb["VENTILATION"]
    t = vent.totals
    v["C2"], v["C3"], v["C4"] = t["colis"], t["brut"], t["net"]
    v["C5"], v["C6"] = round(t["valeur"], 2), t["quantite"]
    v["I5"] = destinataire
    v["G8"] = float(rate)                                                    # remplace =[1]Feuil1!$B$2
    if fret is not None:
        v["L8"] = float(fret)
    desc = wb["DESCIPTION"]
    row_desc = 3
    for n, art in enumerate(vent.articles[:MAX_ARTICLES], start=1):
        ws = wb[f"ART{n}"]
        ws["A2"] = art.template_name
        ws["C2"] = "NB"
        ws["B4"] = "UE"
        for k, line in enumerate(art.lines[:LAST_LINE_ROW - FIRST_LINE_ROW + 1]):
            r = FIRST_LINE_ROW + k
            ws.cell(r, 4, line["ref"])
            ws.cell(r, 6, ORIGIN_NAME)
            ws.cell(r, 7, round(line["net"], 3))
            ws.cell(r, 8, line["qty"])
            ws.cell(r, 9, round(line["value"] / line["qty"], 4) if line["qty"] else 0)
            ws.cell(r, 10, round(line["value"], 2))
            desc.cell(row_desc, 2, line["ref"])
            desc.cell(row_desc, 4, ORIGIN_NAME)
            desc.cell(row_desc, 5, round(line["net"], 3))
            desc.cell(row_desc, 6, line["qty"])
            desc.cell(row_desc, 7, round(line["value"] / line["qty"], 4) if line["qty"] else 0)
            desc.cell(row_desc, 8, round(line["value"], 2))
            row_desc += 1
        v.cell(12 + n, 4, ORIGIN_CODE)                                       # ORGN : les exportations Alstom sont d'origine MAROC
        v.cell(12 + n, 11, "UE")                                             # Code accord (les formules du modèle ne le reprennent que pour l'article 1)
    wb.template = False
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
