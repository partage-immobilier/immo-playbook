#!/usr/bin/env python3
"""Génère 00-preparer/outils/capacite_emprunt.xlsx : capacité d'emprunt (règle HCSF) et budget total. Toutes les cellules de résultat sont des formules Excel."""
import pathlib
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = pathlib.Path(__file__).resolve().parent.parent
wb = Workbook(); ws = wb.active; ws.title = "Capacité d'emprunt"
blue = Font(color="1F3A5F", bold=True); inp = PatternFill("solid", fgColor="FFF4D6"); res = PatternFill("solid", fgColor="E4EBF3")
thin = Side(style="thin", color="B8C4D0"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
def row(r, label, value=None, fmt=None, fill=None, note=""):
    ws.cell(r, 1, label).border = box
    c = ws.cell(r, 2, value); c.border = box
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if note: ws.cell(r, 3, note).font = Font(italic=True, color="666666")
ws["A1"] = "Capacité d'emprunt (règle HCSF : 35 % d'effort assurance comprise, 25 ans maximum ; en vigueur depuis janvier 2022, à vérifier)"; ws["A1"].font = blue
ws["A2"] = "Cellules jaunes : à remplir. Cellules bleues : calculées."; ws["A2"].font = Font(italic=True)
ws["A4"] = "1. Revenus"; ws["A4"].font = blue
row(5, "Salaires nets mensuels du foyer (avant impôt)", 2800, '#,##0 €', inp, "CDI hors période d'essai ; indépendants : moyenne des 2-3 derniers bilans")
row(6, "Loyers déjà perçus (mensuels, hors charges)", 0, '#,##0 €', inp)
row(7, "Loyer attendu du bien visé (mensuel, hors charges)", 750, '#,##0 €', inp, "Estimation d'agence ou annonces comparables")
row(8, "Part des loyers retenue par la banque", 0.70, '0%', inp, "Pratique courante : 70 %")
row(9, "Revenus retenus", "=B5+(B6+B7)*B8", '#,##0 €', res)
ws["A11"] = "2. Charges de crédit en cours"; ws["A11"].font = blue
row(12, "Mensualités de crédits en cours (immobilier, auto, conso)", 0, '#,##0 €', inp, "Un crédit conso de 250 € = environ 50 000 € de capacité en moins")
ws["A14"] = "3. Mensualité disponible"; ws["A14"].font = blue
row(15, "Taux d'effort maximal", 0.35, '0%', inp, "HCSF : 35 % assurance comprise")
row(16, "Mensualité maximale toutes charges comprises", "=B9*B15", '#,##0 €', res)
row(17, "Mensualité disponible pour le nouveau prêt (assurance comprise)", "=MAX(0,B16-B12)", '#,##0 €', res)
ws["A19"] = "4. Capital empruntable"; ws["A19"].font = blue
row(20, "Taux nominal annuel du prêt", 0.035, '0.00%', inp, "Taux du jour, à demander à la banque")
row(21, "Taux annuel de l'assurance emprunteur (sur capital initial)", 0.0025, '0.00%', inp, "0,10 % à 0,40 % selon âge et santé")
row(22, "Durée en années", 25, '0', inp, "25 maximum (27 avec travaux ≥ 10 % ou construction)")
row(23, "Mensualité hors assurance disponible", "=B17-B25*B21/12", '#,##0 €', res, "Itération simple : l'assurance est prise sur le capital calculé")
row(24, "Facteur d'actualisation (par euro de mensualité)", "=(1-(1+B20/12)^(-B22*12))/(B20/12)", '0.00', res)
row(25, "Capital empruntable", "=ROUND((B17*B24)/(1+B24*B21/12),-3)", '#,##0 €', res, "Résolution exacte : capital × (1/facteur + assurance/12) = mensualité disponible")
row(26, "Mensualité du prêt hors assurance", "=B25/B24", '#,##0 €', res)
row(27, "Assurance mensuelle", "=B25*B21/12", '#,##0 €', res)
row(28, "Mensualité totale (contrôle = mensualité disponible)", "=B26+B27", '#,##0 €', res)
row(29, "Taux d'effort après opération", "=(B12+B28)/B9", '0.0%', res, "Doit rester ≤ 35 %")
ws["A31"] = "5. Reste à vivre (indicatif)"; ws["A31"].font = blue
row(32, "Loyer actuel payé par le foyer (si locataire)", 0, '#,##0 €', inp)
row(33, "Nombre de personnes du foyer", 1, '0', inp)
row(34, "Reste à vivre mensuel", "=B5+B6-B12-B28-B32", '#,##0 €', res, "Seuils bancaires usuels : 800 à 1 200 € par personne")
row(35, "Reste à vivre par personne", "=B34/B33", '#,##0 €', res)
ws.column_dimensions["A"].width = 62; ws.column_dimensions["B"].width = 16; ws.column_dimensions["C"].width = 70

b = wb.create_sheet("Budget de l'opération")
b["A1"] = "Budget total d'une opération et apport nécessaire"; b["A1"].font = blue
def brow(r, label, value=None, fmt='#,##0 €', fill=None, note=""):
    b.cell(r, 1, label).border = box; c = b.cell(r, 2, value); c.border = box; c.number_format = fmt
    if fill: c.fill = fill
    if note: b.cell(r, 3, note).font = Font(italic=True, color="666666")
brow(3, "Prix d'achat négocié (frais d'agence inclus)", 150000, fill=inp)
brow(4, "Taux de frais de notaire", 0.08, '0.0%', inp, "Ancien : 7,5 à 8 % en 2026 ; neuf : 2 à 3 %")
brow(5, "Frais de notaire", "=B3*B4", fill=res)
brow(6, "Travaux", 20000, fill=inp, note="Devis ou fourchettes au m² (module 2)")
brow(7, "Mobilier et équipement (si meublé)", 5000, fill=inp)
brow(8, "Frais de dossier bancaire", 1000, fill=inp)
brow(9, "Garantie du prêt (caution ou hypothèque)", 2000, fill=inp, note="Module 5")
brow(10, "Coût total de l'opération", "=B3+B5+B6+B7+B8+B9", fill=res)
brow(12, "Part du prix financée par la banque", 0.90, '0%', inp, "100 % ou 110 % pour les très bons dossiers")
brow(13, "Travaux financés par la banque ?", "oui", '@', inp, "oui / non")
brow(14, "Montant emprunté", '=B3*B12+IF(B13="oui",B6,0)', fill=res)
brow(15, "Apport nécessaire", "=B10-B14", fill=res)
brow(17, "Épargne disponible", 28000, fill=inp)
brow(18, "Mensualité totale estimée (feuille 1, B28)", "='Capacité d''emprunt'!B28", fill=res)
brow(19, "Trésorerie de sécurité restante après apport", "=B17-B15", fill=res)
brow(20, "En mois de mensualité", "=IF(B18>0,B19/B18,0)", '0.0', res, "Viser au moins 3, idéalement 6")
brow(22, "Capital empruntable (feuille 1, B25)", "='Capacité d''emprunt'!B25", fill=res)
brow(23, "L'opération passe-t-elle en capacité ?", '=IF(B14<=B22,"oui","non : réduire le prix ou l\'emprunt")', '@', res)
b.column_dimensions["A"].width = 52; b.column_dimensions["B"].width = 18; b.column_dimensions["C"].width = 55
out = ROOT / "00-preparer" / "outils" / "capacite_emprunt.xlsx"; wb.save(out); print("->", out.relative_to(ROOT))
