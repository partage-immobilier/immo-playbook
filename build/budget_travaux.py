#!/usr/bin/env python3
"""Génère 07-travaux/outils/budget_travaux.xlsx : budget de travaux par poste avec suivi des devis et des paiements,
et planning simple par lot. Toutes les cellules de résultat sont des formules Excel."""
import pathlib
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = pathlib.Path(__file__).resolve().parent.parent
wb = Workbook(); ws = wb.active; ws.title = "Budget"
blue = Font(color="1F3A5F", bold=True); inp = PatternFill("solid", fgColor="FFF4D6"); res = PatternFill("solid", fgColor="E4EBF3")
head = PatternFill("solid", fgColor="D5DEE9")
thin = Side(style="thin", color="B8C4D0"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
EUR = '#,##0 €'

ws["A1"] = "Budget travaux par poste : devis, montant retenu, acomptes, paiements, écart au budget"; ws["A1"].font = blue
ws["A2"] = "Cellules jaunes : à remplir. Cellules bleues : calculées. Une ligne par lot ; les totaux sont des formules."; ws["A2"].font = Font(italic=True)
ws["A3"] = "Fourchettes 2026 (à vérifier) : électricité 80-120 €/m², plomberie-salle de bain 4 000-8 000 € par pièce d'eau, isolation intérieure 50-90 €/m², peinture 25-40 €/m² peint, sols 30-80 €/m², cuisine 3 000-8 000 €."; ws["A3"].font = Font(italic=True, color="666666")

cols = ["Lot", "Descriptif (ce qui est demandé)", "Budget prévu", "Devis 1", "Devis 2", "Devis 3",
        "Retenu", "Acompte versé", "Payé (total, acompte compris)", "Reste à payer", "Écart au budget", "Écart en %", "Entreprise retenue", "Notes"]
HR = 5
for i, c in enumerate(cols, 1):
    cell = ws.cell(HR, i, c); cell.font = Font(bold=True); cell.fill = head; cell.border = box
    cell.alignment = Alignment(wrap_text=True, vertical="center")
ws.row_dimensions[HR].height = 32

lots = [
    ("Démolition, dépose, évacuation", "Dépose de l'existant, gravats en déchetterie", 1500),
    ("Maçonnerie, cloisons", "Cloisons, ouvertures, reprises", 3000),
    ("Électricité", "Tableau, circuits, points, Consuel", 6000),
    ("Plomberie, salle de bain", "Réseaux, douche, meuble vasque, faïence", 6000),
    ("Isolation, plâtrerie", "Doublage des murs extérieurs, plafonds", 2500),
    ("Menuiseries extérieures", "Fenêtres double vitrage, volets", 3600),
    ("Chauffage, eau chaude, VMC", "Radiateurs, chauffe-eau, VMC", 2500),
    ("Peinture", "Murs et plafonds, deux couches", 4500),
    ("Sols", "Stratifié ou vinyle, plinthes", 2500),
    ("Cuisine", "Meubles, plan de travail, évier, raccordements", 4500),
    ("Divers, nettoyage", "Nettoyage de fin de chantier, petites fournitures", 500),
]
FIRST = HR + 1
for r, (lot, desc, budget) in enumerate(lots, FIRST):
    ws.cell(r, 1, lot).fill = inp; ws.cell(r, 2, desc).fill = inp
    ws.cell(r, 3, budget).fill = inp
    for c in (4, 5, 6): ws.cell(r, c, None).fill = inp
    ws.cell(r, 7, None).fill = inp            # retenu : saisi (montant du devis signé)
    ws.cell(r, 8, None).fill = inp            # acompte versé
    ws.cell(r, 9, None).fill = inp            # payé total
    ws.cell(r, 10, f"=IF(G{r}=\"\",0,G{r}-I{r})").fill = res
    ws.cell(r, 11, f"=IF(G{r}=\"\",0,G{r}-C{r})").fill = res
    ws.cell(r, 12, f"=IF(OR(G{r}=\"\",C{r}=0),0,(G{r}-C{r})/C{r})").fill = res
    ws.cell(r, 13, None).fill = inp; ws.cell(r, 14, None).fill = inp
    for c in range(1, len(cols) + 1):
        cell = ws.cell(r, c); cell.border = box
        if c in (3, 4, 5, 6, 7, 8, 9, 10, 11): cell.number_format = EUR
        if c == 12: cell.number_format = '0.0%'
LAST = FIRST + len(lots) - 1

# ligne imprévus
r = LAST + 1
ws.cell(r, 1, "Imprévus").font = Font(bold=True); ws.cell(r, 2, "Provision, en % du budget des lots").fill = inp
ws.cell(r, 3, f"=ROUND(SUM(C{FIRST}:C{LAST})*B{r + 3},0)").fill = res
for c in range(1, len(cols) + 1):
    ws.cell(r, c).border = box
ws.cell(r, 3).number_format = EUR
IMPR = r

# totaux
r = LAST + 2
ws.cell(r, 1, "TOTAL").font = Font(bold=True)
ws.cell(r, 3, f"=SUM(C{FIRST}:C{LAST})+C{IMPR}")
for c, L in ((4, "D"), (5, "E"), (6, "F"), (7, "G"), (8, "H"), (9, "I"), (10, "J")):
    ws.cell(r, c, f"=SUM({L}{FIRST}:{L}{LAST})")
ws.cell(r, 11, f"=G{r}-SUM(C{FIRST}:C{LAST})")
ws.cell(r, 12, f"=IF(SUM(C{FIRST}:C{LAST})=0,0,K{r}/SUM(C{FIRST}:C{LAST}))")
for c in range(1, len(cols) + 1):
    cell = ws.cell(r, c); cell.border = box; cell.fill = res; cell.font = Font(bold=True)
    if 3 <= c <= 11: cell.number_format = EUR
ws.cell(r, 12).number_format = '0.0%'
ws.cell(r, 14, "Écart au budget = retenu moins budget prévu hors imprévus ; l'imprévu sert à l'absorber").font = Font(italic=True, color="666666")
TOT = r

# paramètres
r = TOT + 2
ws.cell(r, 1, "Paramètres").font = blue
ws.cell(r + 1, 1, "Taux d'imprévus sur le budget des lots"); ws.cell(r + 1, 2, 0.10).fill = inp; ws.cell(r + 1, 2).number_format = '0%'
ws.cell(r + 1, 3, "Conseillé : 10 %").font = Font(italic=True, color="666666")
ws.cell(r + 2, 1, "Acompte maximal conseillé à la commande"); ws.cell(r + 2, 2, 0.30).fill = inp; ws.cell(r + 2, 2).number_format = '0%'
ws.cell(r + 2, 3, "Au-delà de 30 %, vous financez l'artisan ; attestations d'assurance avant tout versement").font = Font(italic=True, color="666666")
ws.cell(r + 3, 1, "Retenue de garantie"); ws.cell(r + 3, 2, 0.05).fill = inp; ws.cell(r + 3, 2).number_format = '0%'
ws.cell(r + 3, 3, "5 % pendant un an après réception (loi du 16 juillet 1971), à prévoir au devis").font = Font(italic=True, color="666666")
# correction : la formule des imprévus pointe sur la cellule du taux
ws.cell(IMPR, 3, f"=ROUND(SUM(C{FIRST}:C{LAST})*B{r + 1},0)")

r = r + 5
ws.cell(r, 1, "Contrôles").font = blue
ws.cell(r + 1, 1, "Total des acomptes versés en % du retenu"); ws.cell(r + 1, 2, f"=IF(G{TOT}=0,0,H{TOT}/G{TOT})").fill = res; ws.cell(r + 1, 2).number_format = '0.0%'
ws.cell(r + 1, 3, f"=IF(B{r + 1}>B{TOT + 4},\"Attention : acomptes au-delà du plafond conseillé\",\"ok\")").font = Font(italic=True, color="666666")
ws.cell(r + 2, 1, "Retenue de garantie à conserver jusqu'à un an après réception"); ws.cell(r + 2, 2, f"=G{TOT}*B{TOT + 5}").fill = res; ws.cell(r + 2, 2).number_format = EUR
ws.cell(r + 3, 1, "Reste à payer en % du retenu"); ws.cell(r + 3, 2, f"=IF(G{TOT}=0,0,J{TOT}/G{TOT})").fill = res; ws.cell(r + 3, 2).number_format = '0.0%'
ws.cell(r + 4, 1, "Budget total (lots + imprévus) couvert par le retenu ?"); ws.cell(r + 4, 2, f"=IF(G{TOT}<=C{TOT},\"oui\",\"non : dépassement de \"&TEXT(G{TOT}-C{TOT},\"# ##0\")&\" €\")").fill = res

widths = [30, 38, 13, 12, 12, 12, 12, 13, 14, 13, 13, 9, 22, 40]
for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = ws.cell(FIRST, 3)

# ---------------- Feuille 2 : planning ----------------
p = wb.create_sheet("Planning")
p["A1"] = "Planning simple par lot : date de début, durée en jours ouvrés, fin calculée"; p["A1"].font = blue
p["A2"] = "Cellules jaunes : à remplir. Fin = jours ouvrés après le début (WORKDAY, hors week-ends). Le lot suivant commence en général après la fin du précédent."; p["A2"].font = Font(italic=True)
pcols = ["Lot", "Entreprise", "Début", "Durée (jours ouvrés)", "Fin calculée", "Début réel", "Fin réelle", "Retard (jours ouvrés)", "Remarques"]
for i, c in enumerate(pcols, 1):
    cell = p.cell(4, i, c); cell.font = Font(bold=True); cell.fill = head; cell.border = box; cell.alignment = Alignment(wrap_text=True)
p.row_dimensions[4].height = 30
p["C5"] = "Date de démarrage du chantier"; p["D5"] = None; p["D5"].fill = inp; p["D5"].number_format = 'DD/MM/YYYY'
p["E5"] = "Saisir une date ; les lots s'enchaînent à partir d'elle"; p["E5"].font = Font(italic=True, color="666666")
PF = 7
durations = [3, 8, 6, 8, 6, 2, 3, 6, 3, 2, 1]
for k, ((lot, _, _), d) in enumerate(zip(lots, durations)):
    r = PF + k
    p.cell(r, 1, lot).fill = inp; p.cell(r, 2, None).fill = inp
    if k == 0:
        p.cell(r, 3, "=IF(D5=\"\",\"\",D5)")
    else:
        p.cell(r, 3, f"=IF(E{r - 1}=\"\",\"\",WORKDAY(E{r - 1},1))")
    p.cell(r, 3).fill = res
    p.cell(r, 4, d).fill = inp
    p.cell(r, 5, f"=IF(C{r}=\"\",\"\",WORKDAY(C{r},D{r}-1))").fill = res
    p.cell(r, 6, None).fill = inp; p.cell(r, 7, None).fill = inp
    p.cell(r, 8, f"=IF(OR(G{r}=\"\",E{r}=\"\"),\"\",NETWORKDAYS(E{r},G{r})-1)").fill = res
    p.cell(r, 9, None).fill = inp
    for c in range(1, len(pcols) + 1):
        cell = p.cell(r, c); cell.border = box
        if c in (3, 5, 6, 7): cell.number_format = 'DD/MM/YYYY'
PL = PF + len(lots) - 1
r = PL + 1
p.cell(r, 1, "Durée totale (jours ouvrés)").font = Font(bold=True); p.cell(r, 4, f"=SUM(D{PF}:D{PL})").fill = res
p.cell(r, 5, f"=IF(D5=\"\",\"\",E{PL})").fill = res; p.cell(r, 5).number_format = 'DD/MM/YYYY'
p.cell(r + 1, 1, "Fin prévue du chantier (dernier lot)").font = Font(bold=True); p.cell(r + 1, 5, f"=E{PL}").fill = res; p.cell(r + 1, 5).number_format = 'DD/MM/YYYY'
p.cell(r + 2, 1, "Réception : fixer la date, préparer le PV, lister les réserves pièce par pièce").font = Font(italic=True, color="666666")
p.cell(r + 3, 1, "Si vous modifiez l'ordre des lots, remplacez la formule de début par une date ou par =WORKDAY(fin du lot précédent,1).").font = Font(italic=True, color="666666")
for i, w in enumerate([30, 22, 13, 12, 13, 13, 13, 12, 40], 1): p.column_dimensions[get_column_letter(i)].width = w

out = ROOT / "07-travaux" / "outils" / "budget_travaux.xlsx"; wb.save(out); print("->", out.relative_to(ROOT))
