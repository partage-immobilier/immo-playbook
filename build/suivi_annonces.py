#!/usr/bin/env python3
"""Génère 01-rechercher/outils/suivi_annonces.xlsx : tableau de suivi des annonces avec prix au m², rendement brut et écart DVF calculés par formules Excel. Cellules jaunes à saisir, bleues calculées, notes en ligne 2."""
import pathlib
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

ROOT = pathlib.Path(__file__).resolve().parent.parent
wb = Workbook(); ws = wb.active; ws.title = "Suivi des annonces"
blue = Font(color="1F3A5F", bold=True); inp = PatternFill("solid", fgColor="FFF4D6"); res = PatternFill("solid", fgColor="E4EBF3")
thin = Side(style="thin", color="B8C4D0"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
note_font = Font(italic=True, color="666666", size=9)

FIRST, LAST = 4, 203  # lignes de saisie : 200 annonces
# (en-tête, largeur, note, formule ou None, format)
COLS = [
    ("Date", 12, "Date où vous avez vu l'annonce", None, "DD/MM/YYYY"),
    ("Source", 13, "Leboncoin, SeLoger, PAP, agent, notaire, courrier", None, "@"),
    ("Ville", 16, "Commune, et quartier si utile", None, "@"),
    ("Type", 10, "Studio, T2, T3, maison, immeuble", None, "@"),
    ("Surface (m²)", 11, "Surface habitable, pas la surface au sol", None, "0.0"),
    ("Prix affiché (€)", 14, "Prix FAI (frais d'agence inclus)", None, "#,##0 €"),
    ("Prix au m² (€)", 13, "Calculé : prix ÷ surface", "=IF(AND(E{r}>0,F{r}>0),F{r}/E{r},\"\")", "#,##0 €"),
    ("Loyer estimé (€/mois)", 14, "Votre estimation prudente, hors charges (carte des loyers − 5 à 10 %)", None, "#,##0 €"),
    ("Rendement brut", 12, "Calculé : loyer × 12 ÷ prix affiché", "=IF(AND(F{r}>0,H{r}>0),H{r}*12/F{r},\"\")", "0.0%"),
    ("Travaux estimés (€)", 14, "Au m² (module 2) ou sur devis ; 0 si aucun", None, "#,##0 €"),
    ("Brute travaux compris", 13, "Calculé : loyer × 12 ÷ (prix + travaux)", "=IF(AND(F{r}>0,H{r}>0),H{r}*12/(F{r}+N(J{r})),\"\")", "0.0%"),
    ("Statut", 14, "Liste déroulante", None, "@"),
    ("Contact", 22, "Nom et téléphone de l'agent ou du vendeur", None, "@"),
    ("Prochaine action", 30, "Toujours avec une date : « relancer l'agent le 22/11 »", None, "@"),
    ("Médiane DVF (€/m²)", 13, "Repère du quartier, calculé une fois (app.dvf.etalab.gouv.fr)", None, "#,##0 €"),
    ("Écart vs DVF", 11, "Calculé : (prix au m² − médiane) ÷ médiane ; sous −15 % : appeler dans l'heure", "=IF(AND(ISNUMBER(G{r}),O{r}>0),G{r}/O{r}-1,\"\")", "0%"),
    ("Prix de revient au m²", 13, "Calculé : (prix + travaux) ÷ surface, à comparer à la médiane DVF des biens rénovés", "=IF(AND(E{r}>0,F{r}>0),(F{r}+N(J{r}))/E{r},\"\")", "#,##0 €"),
    ("Lien", 30, "Adresse de l'annonce", None, "@"),
    ("Notes", 40, "Ce que vous avez appris à l'appel et à la visite", None, "@"),
]
STATUTS = ["Vu", "Appelé", "Visite prévue", "Visité", "Offre faite", "Négociation", "Écarté", "Perdu", "Acquis"]

ws["A1"] = "Tableau de suivi des annonces (module 1). Cellules jaunes : à remplir. Cellules bleues : calculées. Une ligne par annonce, même écartée."
ws["A1"].font = blue
for i, (head, width, note, formula, fmt) in enumerate(COLS, 1):
    col = get_column_letter(i)
    ws.column_dimensions[col].width = width
    n = ws.cell(2, i, note); n.font = note_font; n.alignment = Alignment(wrap_text=True, vertical="top")
    h = ws.cell(3, i, head); h.font = Font(bold=True, color="FFFFFF"); h.fill = PatternFill("solid", fgColor="1F3A5F")
    h.alignment = Alignment(wrap_text=True, vertical="center"); h.border = box
    for r in range(FIRST, LAST + 1):
        c = ws.cell(r, i); c.border = box; c.number_format = fmt
        if formula:
            c.value = formula.format(r=r); c.fill = res
        else:
            c.fill = inp
ws.row_dimensions[2].height = 54; ws.row_dimensions[3].height = 32
ws.freeze_panes = "A4"
ws.auto_filter.ref = f"A3:{get_column_letter(len(COLS))}{LAST}"

dv = DataValidation(type="list", formula1='"' + ",".join(STATUTS) + '"', allow_blank=True)
dv.error = "Choisir un statut dans la liste"; ws.add_data_validation(dv); dv.add(f"L{FIRST}:L{LAST}")

# Mise en forme conditionnelle : rendement brut travaux compris et écart DVF
green = PatternFill("solid", fgColor="D9EAD3"); red = PatternFill("solid", fgColor="F4CCCC")
ws.conditional_formatting.add(f"K{FIRST}:K{LAST}", CellIsRule(operator="greaterThanOrEqual", formula=["'Paramètres'!$B$3"], fill=green))
ws.conditional_formatting.add(f"K{FIRST}:K{LAST}", CellIsRule(operator="lessThan", formula=["'Paramètres'!$B$4"], fill=red))
ws.conditional_formatting.add(f"P{FIRST}:P{LAST}", CellIsRule(operator="lessThanOrEqual", formula=["'Paramètres'!$B$5"], fill=green))
ws.conditional_formatting.add(f"P{FIRST}:P{LAST}", CellIsRule(operator="greaterThan", formula=["'Paramètres'!$B$6"], fill=red))

# Ligne d'exemple (annonce 2 du cours : T3 à rénover)
example = {
    1: "2026-11-15", 2: "Leboncoin", 3: "Ville A, centre", 4: "T3", 5: 62, 6: 89000, 8: 600, 10: 25000,
    12: "Appelé", 13: "Agence du Centre, M. Martin, 06 00 00 00 00",
    14: "Visite le 17/11 à 14 h ; demander PV d'AG et diagnostics avant", 15: 1700,
    18: "https://www.leboncoin.fr/...", 19: "Prix baissé de 95 000 à 89 000 €, pas de photo de la SdB, 2e étage sans ascenseur, charges 1 100 €/an",
}
from datetime import date
for col, val in example.items():
    ws.cell(FIRST, col, date(2026, 11, 15) if col == 1 else val)

# Feuille des paramètres de tri
p = wb.create_sheet("Paramètres")
p["A1"] = "Seuils de la grille de tri (module 1) : adaptez-les à votre objectif"; p["A1"].font = blue
rows = [
    (3, "Rendement brut travaux compris : seuil « bonne affaire » (vert)", 0.08, "0.0%", "Rente en ville moyenne : 8 % ; immeuble : 9 à 10 % ; patrimoine grande ville : 4 à 5 %"),
    (4, "Rendement brut travaux compris : seuil « à écarter » (rouge)", 0.05, "0.0%", "En dessous, le loyer ne couvrira pas le crédit sur 20 ans pour un objectif de rente"),
    (5, "Écart vs DVF : seuil « appeler dans l'heure » (vert)", -0.15, "0%", "Plus de 15 % sous la médiane des ventes réelles"),
    (6, "Écart vs DVF : seuil « trop cher » (rouge)", 0.10, "0%", "Plus de 10 % au-dessus, sauf bien rénové ou exceptionnel"),
    (8, "Fourchettes de travaux 2026 (€/m², à vérifier sur devis) : rafraîchissement", "150 à 300", "@", "Peinture, sols, petites reprises"),
    (9, "Rénovation moyenne", "500 à 900", "@", "Cuisine, salle de bains, électricité partielle, peinture, sols"),
    (10, "Rénovation lourde", "1 000 à 1 500", "@", "Tout refaire : réseaux, cloisons, menuiseries, isolation"),
]
for r, label, val, fmt, note in rows:
    p.cell(r, 1, label).border = box
    c = p.cell(r, 2, val); c.border = box; c.number_format = fmt; c.fill = inp
    p.cell(r, 3, note).font = note_font
p["A12"] = "Statuts disponibles : " + ", ".join(STATUTS); p["A12"].font = note_font
p.column_dimensions["A"].width = 64; p.column_dimensions["B"].width = 14; p.column_dimensions["C"].width = 80

out = ROOT / "01-rechercher" / "outils" / "suivi_annonces.xlsx"; wb.save(out); print("->", out.relative_to(ROOT))
