#!/usr/bin/env python3
"""Génère 10-local-commercial/outils/analyse_local.xlsx : analyse d'un local commercial (rendement brut et net, DSCR,
cash-flow, prix plafond) et sensibilité à la vacance (3, 6, 12 mois) et à +20 % de travaux. Toutes les sorties sont des formules Excel."""
import pathlib
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side

ROOT = pathlib.Path(__file__).resolve().parent.parent
wb = Workbook(); ws = wb.active; ws.title = "Analyse du local"
blue = Font(color="1F3A5F", bold=True); inp = PatternFill("solid", fgColor="FFF4D6"); res = PatternFill("solid", fgColor="E4EBF3")
thin = Side(style="thin", color="B8C4D0"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
NOTE = Font(italic=True, color="666666")

def row(sheet, r, label, value=None, fmt='#,##0 €', fill=None, note=""):
    sheet.cell(r, 1, label).border = box
    c = sheet.cell(r, 2, value); c.border = box
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if note: sheet.cell(r, 3, note).font = NOTE

ws["A1"] = "Analyse d'un local commercial : rendement, DSCR, cash-flow et prix plafond (à jour au 15/11/2026, pratiques bancaires à vérifier)"; ws["A1"].font = blue
ws["A2"] = "Cellules jaunes : à remplir. Cellules bleues : calculées. Tous les loyers sont hors taxes et hors charges."; ws["A2"].font = Font(italic=True)

ws["A4"] = "1. Coût de l'opération"; ws["A4"].font = blue
row(ws, 5, "Prix d'achat négocié (frais d'agence inclus)", 150000, fill=inp)
row(ws, 6, "Taux de frais d'acquisition (notaire)", 0.08, '0.0%', inp, "Ancien : 7,5 à 8 % en 2026 ; 2 à 3 % si vente soumise à TVA (neuf ou option)")
row(ws, 7, "Frais d'acquisition", "=B5*B6", fill=res)
row(ws, 8, "Travaux à la charge du bailleur", 6000, fill=inp, note="Structure, façade, mise en conformité du bâti ; les aménagements sont au preneur")
row(ws, 9, "Frais bancaires et garantie", 2500, fill=inp, note="Dossier 0,5 à 1 % du prêt, hypothèque 1 à 1,5 %")
row(ws, 10, "Coût total de l'opération", "=B5+B7+B8+B9", fill=res)

ws["A12"] = "2. Loyer et charges"; ws["A12"].font = blue
row(ws, 13, "Loyer annuel HT hors charges (bail en cours ou loyer de marché prudent)", 14400, fill=inp, note="Local vide : loyer de marché après annonce test, pas le loyer espéré")
row(ws, 14, "Charges non refacturables annuelles", 1400, fill=inp, note="Assurance de l'immeuble, comptable, quote-part art. 606, honoraires de gestion (R145-35 C. com.)")
row(ws, 15, "Taxe foncière annuelle", 1900, fill=inp)
row(ws, 16, "Taxe foncière refacturée au preneur ?", "oui", '@', inp, "oui / non : refacturable seulement si le bail le prévoit (très courant)")
row(ws, 17, "Taxe foncière supportée par le bailleur", '=IF(B16="oui",0,B15)', fill=res)
row(ws, 18, "Vacance retenue en base (mois par an)", 1, '0.0', inp, "Au moins 1 mois par an ; davantage pour un local spécialisé")
row(ws, 19, "Loyer net annuel (après vacance, charges et taxe foncière supportée)", "=B13*(1-B18/12)-B14-B17", fill=res)
row(ws, 20, "Rendement brut (loyer / prix)", "=B13/B5", '0.0%', res, "Le chiffre des annonces : sert à trier, pas à décider")
row(ws, 21, "Rendement net (loyer net / coût total)", "=B19/B10", '0.0%', res, "Le seul chiffre comparable d'un local à l'autre")

ws["A23"] = "3. Financement"; ws["A23"].font = blue
row(ws, 24, "Part du coût total financée par la banque", 0.70, '0%', inp, "Crédit professionnel : 70 à 80 % courant, apport 20 à 30 % (pratique 2026)")
row(ws, 25, "Montant emprunté", "=B10*B24", fill=res)
row(ws, 26, "Apport", "=B10-B25", fill=res)
row(ws, 27, "Taux nominal annuel", 0.042, '0.00%', inp, "0,3 à 1 point au-dessus du crédit immobilier des particuliers")
row(ws, 28, "Taux annuel d'assurance (sur capital initial)", 0.003, '0.00%', inp)
row(ws, 29, "Durée en années", 15, '0', inp, "12 à 15 ans courant, 20 ans possible pour un bien bien loué")
row(ws, 30, "Mensualité hors assurance", "=B25*(B27/12)/(1-(1+B27/12)^(-B29*12))", fill=res)
row(ws, 31, "Annuité totale (capital, intérêts, assurance)", "=B30*12+B25*B28", fill=res)
row(ws, 32, "DSCR (loyer net / annuité)", "=IF(B31>0,B19/B31,0)", '0.00', res, "Attente courante des banques professionnelles : 1,2 à 1,3")
row(ws, 33, "Cash-flow annuel avant impôt", "=B19-B31", fill=res)
row(ws, 34, "Cash-flow mensuel avant impôt", "=B33/12", fill=res)
row(ws, 35, "DSCR minimal exigé par la banque", 1.25, '0.00', inp)
row(ws, 36, "Le dossier passe-t-il en DSCR ?", '=IF(B32>=B35,"oui","non : baisser le prix, allonger la durée ou augmenter l\'apport")', '@', res)

ws["A38"] = "4. Prix plafond (calcul à l'envers, à partir du loyer net)"; ws["A38"].font = blue
row(ws, 39, "Annuité maximale supportable (loyer net / DSCR exigé)", "=B19/B35", fill=res)
row(ws, 40, "Annuité par euro emprunté (crédit + assurance)", "=12*(B27/12)/(1-(1+B27/12)^(-B29*12))+B28", '0.0000', res)
row(ws, 41, "Capital empruntable maximal", "=B39/B40", fill=res)
row(ws, 42, "Coût total maximal (capital / part financée)", "=B41/B24", fill=res)
row(ws, 43, "Prix plafond (hors frais, travaux et frais bancaires)", "=MAX(0,(B42-B8-B9)/(1+B6))", fill=res, note="Au-delà, le DSCR exigé n'est plus respecté avec cette structure de financement")
row(ws, 44, "Écart entre le prix négocié et le prix plafond", "=B5-B43", fill=res, note="Positif : il faut négocier d'autant")

ws["A46"] = "5. Valeur de revente indicative (capitalisation du loyer)"; ws["A46"].font = blue
row(ws, 47, "Taux de rendement exigé par le marché pour ce type de local", 0.08, '0.0%', inp, "4,5 à 6 % métropole ; 7 à 9 % ville moyenne, preneur solide ; 9 à 12 % petite ville ou preneur fragile")
row(ws, 48, "Valeur indicative loué (loyer HT / taux)", "=B13/B47", fill=res)
row(ws, 49, "Décote si vendu vide", 0.30, '0%', inp, "20 à 40 %")
row(ws, 50, "Valeur indicative vide", "=B48*(1-B49)", fill=res)

ws.column_dimensions["A"].width = 68; ws.column_dimensions["B"].width = 18; ws.column_dimensions["C"].width = 80

# ---- Feuille de sensibilité -------------------------------------------------------------------------------
s = wb.create_sheet("Sensibilité")
s["A1"] = "Sensibilité du cash-flow et du DSCR à la vacance et aux travaux (formules liées à la feuille Analyse du local)"; s["A1"].font = blue
s["A2"] = "Chaque colonne reprend les hypothèses de la feuille 1 et ne change que le paramètre indiqué. Lignes jaunes modifiables."; s["A2"].font = Font(italic=True)
A = "'Analyse du local'!"
heads = ["Paramètre", "Base", "Vacance 3 mois", "Vacance 6 mois", "Vacance 12 mois", "Travaux +20 %", "Vacance 6 mois et travaux +20 %"]
for j, h in enumerate(heads, 1):
    c = s.cell(4, j, h); c.font = blue; c.border = box; c.fill = res
# paramètres de scénario (modifiables)
vac = [f"={A}B18", 3, 6, 12, f"={A}B18", 6]
trav = [1, 1, 1, 1, 1.2, 1.2]
labels = [
    ("Vacance (mois sur l'année)", '0.0'),
    ("Coefficient appliqué aux travaux", '0.00'),
    ("Travaux à la charge du bailleur", '#,##0 €'),
    ("Coût total de l'opération", '#,##0 €'),
    ("Montant emprunté", '#,##0 €'),
    ("Apport", '#,##0 €'),
    ("Loyer encaissé (après vacance)", '#,##0 €'),
    ("Loyer net (après charges et taxe foncière supportée)", '#,##0 €'),
    ("Annuité totale", '#,##0 €'),
    ("DSCR", '0.00'),
    ("Cash-flow annuel avant impôt", '#,##0 €'),
    ("Cash-flow mensuel avant impôt", '#,##0 €'),
    ("Rendement net", '0.0%'),
    ("Effort de trésorerie sur l'année (si cash-flow négatif)", '#,##0 €'),
]
for i, (lab, fmt) in enumerate(labels):
    r = 5 + i
    s.cell(r, 1, lab).border = box
    for j in range(6):
        col = chr(ord("B") + j)
        if i == 0: v, fill = vac[j], inp
        elif i == 1: v, fill = trav[j], inp
        elif i == 2: v, fill = f"={A}B8*{col}6", res
        elif i == 3: v, fill = f"={A}B5+{A}B7+{col}7+{A}B9", res
        elif i == 4: v, fill = f"={col}8*{A}B24", res
        elif i == 5: v, fill = f"={col}8-{col}9", res
        elif i == 6: v, fill = f"={A}B13*(1-{col}5/12)", res
        elif i == 7: v, fill = f"={col}11-{A}B14-{A}B17", res
        elif i == 8: v, fill = f"={col}9*12*({A}B27/12)/(1-(1+{A}B27/12)^(-{A}B29*12))+{col}9*{A}B28", res
        elif i == 9: v, fill = f"=IF({col}13>0,{col}12/{col}13,0)", res
        elif i == 10: v, fill = f"={col}12-{col}13", res
        elif i == 11: v, fill = f"={col}15/12", res
        elif i == 12: v, fill = f"={col}12/{col}8", res
        else: v, fill = f"=MAX(0,-{col}15)", res
        c = s.cell(r, j + 2, v); c.border = box; c.number_format = fmt; c.fill = fill
s.cell(20, 1, "Lecture : un local vide la première année coûte l'effort de trésorerie de la colonne « Vacance 12 mois » ; c'est la réserve à avoir, ou le différé d'amortissement à négocier.").font = NOTE
s.cell(21, 1, "Un DSCR qui passe sous 1,0 dans la colonne « Vacance 6 mois » signale un local qui ne supporte pas un changement de preneur sans apport de trésorerie.").font = NOTE
s.column_dimensions["A"].width = 58
for col in "BCDEFG": s.column_dimensions[col].width = 20

out = ROOT / "10-local-commercial" / "outils" / "analyse_local.xlsx"; wb.save(out); print("->", out.relative_to(ROOT))
