#!/usr/bin/env python3
"""Génère 05-financer/outils/simulateur_pret.xlsx : mensualité et coût total avec comparaison 15/20/25 ans,
tableau d'amortissement sur 300 mois, comparaison amortissable / in fine. Toutes les sorties sont des formules Excel."""
import pathlib
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

ROOT = pathlib.Path(__file__).resolve().parent.parent
wb = Workbook()
blue = Font(color="1F3A5F", bold=True); note_font = Font(italic=True, color="666666")
inp = PatternFill("solid", fgColor="FFF4D6"); res = PatternFill("solid", fgColor="E4EBF3")
thin = Side(style="thin", color="B8C4D0"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
EUR, PCT, PCT2, INT = '#,##0 €', '0.00%', '0.0%', '0'

def cell(ws, r, c, value=None, fmt=None, fill=None):
    x = ws.cell(r, c, value); x.border = box
    if fmt: x.number_format = fmt
    if fill: x.fill = fill
    return x

def row(ws, r, label, value=None, fmt=None, fill=None, note="", note_col=3):
    cell(ws, r, 1, label); cell(ws, r, 2, value, fmt, fill)
    if note: ws.cell(r, note_col, note).font = note_font

# ---------------------------------------------------------------- Feuille 1
ws = wb.active; ws.title = "Mensualité et coût"
ws["A1"] = "Simulateur de prêt : mensualité, coût total et comparaison des durées"; ws["A1"].font = blue
ws["A2"] = "Cellules jaunes : à remplir. Cellules bleues : calculées. Notes en colonne C (bloc 1) et E (bloc 2)."; ws["A2"].font = Font(italic=True)
ws["A4"] = "1. Le prêt et votre situation"; ws["A4"].font = blue
row(ws, 5, "Capital emprunté", 200000, EUR, inp, "Prix + frais + travaux financés, moins l'apport")
row(ws, 6, "Taux annuel de l'assurance emprunteur (sur capital initial)", 0.0025, PCT, inp, "0,10 à 0,15 % en délégation pour un trentenaire, 0,25 à 0,35 % en contrat de banque")
row(ws, 7, "Frais de dossier et de garantie", 2500, EUR, inp, "Frais de dossier 0,5 à 1 % + garantie (hypothèque légale du prêteur de deniers 1 000 à 1 500 €, caution 2 000 à 2 800 €)")
row(ws, 8, "Revenus mensuels retenus par la banque", 3300, EUR, inp, "Salaires nets + 70 % des loyers existants et futur (module 0)")
row(ws, 9, "Autres mensualités de crédit en cours (assurance comprise)", 0, EUR, inp)
ws["A11"] = "2. Comparaison de trois durées (modifiez les durées et les taux en jaune)"; ws["A11"].font = blue
hdr = ["", "Durée 1", "Durée 2", "Durée 3", "Note"]
for c, h in enumerate(hdr, 1):
    x = cell(ws, 12, c, h); x.font = Font(bold=True); x.fill = PatternFill("solid", fgColor="E4EBF3")
cols = ["B", "C", "D"]
lines = [
    (13, "Durée (années)", [15, 20, 25], INT, inp, "25 ans maximum (HCSF, depuis janvier 2022), 27 avec différé travaux ou construction"),
    (14, "Taux nominal annuel", [0.032, 0.034, 0.035], PCT, inp, "Taux du jour, hors assurance ; plus la durée est longue, plus le taux est élevé"),
    (15, "Nombre de mensualités", ["={c}13*12"] * 3, INT, res, ""),
    (16, "Mensualité hors assurance", ["=$B$5*({c}14/12)/(1-(1+{c}14/12)^(-{c}15))"] * 3, EUR, res, "Formule de l'annuité constante"),
    (17, "Assurance mensuelle", ["=$B$5*$B$6/12"] * 3, EUR, res, "Sur capital initial (contrat de banque) ; une délégation sur capital restant dû coûte moins"),
    (18, "Mensualité totale", ["={c}16+{c}17"] * 3, EUR, res, "C'est le chiffre retenu dans le taux d'effort"),
    (19, "Total des intérêts", ["={c}16*{c}15-$B$5"] * 3, EUR, res, ""),
    (20, "Total de l'assurance", ["={c}17*{c}15"] * 3, EUR, res, ""),
    (21, "Frais de dossier et de garantie", ["=$B$7"] * 3, EUR, res, ""),
    (22, "Coût total du crédit", ["={c}19+{c}20+{c}21"] * 3, EUR, res, "Intérêts + assurance + frais : le seul chiffre qui compare deux offres"),
    (23, "Coût total en % du capital", ["={c}22/$B$5"] * 3, PCT2, res, ""),
    (24, "TAEG approché", ["=RATE({c}15,-{c}18,$B$5-$B$7)*12"] * 3, PCT, res, "À comparer au taux d'usure du trimestre (banque-france.fr)"),
    (25, "Taux d'effort après opération", ["=({c}18+$B$9)/$B$8"] * 3, PCT2, res, "Doit rester sous 35 % assurance comprise"),
    (26, "Passe sous 35 % ?", ['=IF({c}25<=0.35,"oui","non")'] * 3, '@', res, ""),
    (27, "Écart de mensualité avec la durée 1", ["=B18-B18", "=B18-C18", "=B18-D18"], EUR, res, "Ce que chaque allongement libère chaque mois"),
    (28, "Écart de coût total avec la durée 1", ["=B22-B22", "=C22-B22", "=D22-B22"], EUR, res, "Ce que chaque allongement coûte sur la durée"),
]
for r, label, vals, fmt, fill, note in lines:
    cell(ws, r, 1, label)
    for c, v in zip(cols, vals):
        cell(ws, r, ord(c) - 64, v.format(c=c) if isinstance(v, str) else v, fmt, fill)
    if note: ws.cell(r, 5, note).font = note_font
ws.column_dimensions["A"].width = 58; ws.column_dimensions["E"].width = 80
for c in "BCD": ws.column_dimensions[c].width = 15
ws.column_dimensions["C"].width = 15

# ---------------------------------------------------------------- Feuille 2
am = wb.create_sheet("Amortissement")
am["A1"] = "Tableau d'amortissement sur 300 mois (25 ans)"; am["A1"].font = blue
am["A2"] = "Les cellules jaunes reprennent par défaut la durée 3 de la feuille 1 ; remplacez-les par vos valeurs si besoin."; am["A2"].font = Font(italic=True)
row(am, 3, "Capital emprunté", "='Mensualité et coût'!B5", EUR, inp)
row(am, 4, "Taux nominal annuel", "='Mensualité et coût'!D14", PCT, inp)
row(am, 5, "Durée (années)", "='Mensualité et coût'!D13", INT, inp, "Au-delà de 25 ans, les lignes après le mois 300 ne sont pas affichées")
row(am, 6, "Taux annuel de l'assurance (sur capital initial)", "='Mensualité et coût'!B6", PCT, inp)
row(am, 7, "Différé partiel (mois, intérêts et assurance seuls)", 0, INT, inp, "0 sans différé ; 6 à 24 mois pendant des travaux (le capital restant dû ne baisse pas)")
row(am, 8, "Nombre de mensualités", "=B5*12", INT, res)
row(am, 9, "Mensualité hors assurance après le différé", "=B3*(B4/12)/(1-(1+B4/12)^(-(B8-B7)))", EUR, res, "Calculée sur la durée restante après le différé")
row(am, 10, "Assurance mensuelle", "=B3*B6/12", EUR, res)
row(am, 11, "Total des intérêts", "=SUM(C15:C314)", EUR, res)
row(am, 12, "Total de l'assurance", "=SUM(E15:E314)", EUR, res)
heads = ["Mois", "Capital restant dû début", "Intérêts", "Capital remboursé", "Assurance", "Mensualité totale", "Capital restant dû fin"]
for c, h in enumerate(heads, 1):
    x = cell(am, 14, c, h); x.font = Font(bold=True); x.fill = PatternFill("solid", fgColor="E4EBF3"); x.alignment = Alignment(wrap_text=True)
for i in range(300):
    r = 15 + i
    cell(am, r, 1, i + 1, INT)
    cell(am, r, 2, "=$B$3" if i == 0 else f"=G{r-1}", EUR, res)
    cell(am, r, 3, f"=IF(A{r}<=$B$8,B{r}*$B$4/12,0)", EUR, res)
    cell(am, r, 4, f"=IF(A{r}<=$B$7,0,IF(A{r}<=$B$8,MIN(B{r},$B$9-C{r}),0))", EUR, res)
    cell(am, r, 5, f"=IF(A{r}<=$B$8,$B$10,0)", EUR, res)
    cell(am, r, 6, f"=C{r}+D{r}+E{r}", EUR, res)
    cell(am, r, 7, f"=B{r}-D{r}", EUR, res)
am.column_dimensions["A"].width = 52; am.column_dimensions["C"].width = 70
for c in "BDEFG": am.column_dimensions[c].width = 18
am.freeze_panes = "A15"

# ---------------------------------------------------------------- Feuille 3
cf = wb.create_sheet("Amortissable vs in fine")
cf["A1"] = "Comparaison prêt amortissable / prêt in fine"; cf["A1"].font = blue
cf["A2"] = "Le in fine ne rembourse le capital qu'à la fin : il faut le reconstituer par une épargne nantie. Notes en colonne D."; cf["A2"].font = Font(italic=True)
row(cf, 4, "Capital emprunté", 200000, EUR, inp, note_col=4)
row(cf, 5, "Durée (années)", 20, INT, inp, note_col=4)
row(cf, 6, "Taux nominal du prêt amortissable", 0.035, PCT, inp, note_col=4)
row(cf, 7, "Taux nominal du prêt in fine", 0.038, PCT, inp, "Les banques majorent le in fine de 0,2 à 0,5 point", note_col=4)
row(cf, 8, "Taux annuel de l'assurance (sur capital initial)", 0.0025, PCT, inp, note_col=4)
row(cf, 9, "Rendement net du placement nanti (assurance-vie)", 0.025, PCT, inp, "Net de frais et de prélèvements sociaux, prudent", note_col=4)
row(cf, 10, "Taux marginal d'imposition + prélèvements sociaux", 0.472, PCT2, inp, "30 % + 17,2 % = 47,2 % ; 11 % + 17,2 % = 28,2 % ; en SCI à l'IS : 15 % ou 25 %", note_col=4)
row(cf, 11, "Loyer mensuel hors charges", 1200, EUR, inp, note_col=4)
row(cf, 12, "Nombre de mensualités", "=B5*12", INT, res, note_col=4)
for c, h in enumerate(["", "Amortissable", "In fine", "Note"], 1):
    x = cell(cf, 14, c, h); x.font = Font(bold=True); x.fill = PatternFill("solid", fgColor="E4EBF3")
lines3 = [
    (15, "Mensualité hors assurance", "=B4*(B6/12)/(1-(1+B6/12)^(-B12))", "=B4*B7/12", EUR, "In fine : intérêts seuls"),
    (16, "Assurance mensuelle", "=B4*B8/12", "=B4*B8/12", EUR, ""),
    (17, "Épargne mensuelle pour reconstituer le capital", 0, "=B4*(B9/12)/((1+B9/12)^B12-1)", EUR, "Versement constant qui atteint le capital à l'échéance au rendement indiqué"),
    (18, "Décaissement mensuel total", "=B15+B16+B17", "=C15+C16+C17", EUR, "Ce qui sort réellement chaque mois"),
    (19, "Total des intérêts", "=B15*B12-B4", "=C15*B12", EUR, ""),
    (20, "Total de l'assurance", "=B16*B12", "=C16*B12", EUR, ""),
    (21, "Économie d'impôt sur les intérêts (régime réel)", "=B19*B10", "=C19*B10", EUR, "Intérêts déductibles des loyers au réel (nu, meublé, SCI à l'IS)"),
    (22, "Coût net des intérêts après impôt", "=B19-B21", "=C19-C21", EUR, ""),
    (23, "Capital restant dû à l'échéance", 0, "=B4", EUR, "Le in fine doit être remboursé en une fois : épargne nantie ou revente"),
    (24, "Cash-flow mensuel hors épargne (loyer - mensualité - assurance)", "=B11-B15-B16", "=B11-C15-C16", EUR, "Le in fine affiche un cash-flow flatteur, l'épargne de reconstitution n'y est pas"),
    (25, "Cash-flow mensuel réel (épargne de reconstitution déduite)", "=B11-B18", "=B11-C18", EUR, ""),
    (26, "Verdict sur le coût net des intérêts", '=IF(B22<=C22,"amortissable moins cher","in fine moins cher")', '=IF(C22<B22,"in fine moins cher","amortissable moins cher")', '@', "Le in fine ne gagne que si le placement rapporte nettement plus que le taux du prêt"),
]
for r, label, va, vi, fmt, note in lines3:
    cell(cf, r, 1, label); cell(cf, r, 2, va, fmt, res); cell(cf, r, 3, vi, fmt, res)
    if note: cf.cell(r, 4, note).font = note_font
cf.column_dimensions["A"].width = 62; cf.column_dimensions["B"].width = 18; cf.column_dimensions["C"].width = 18; cf.column_dimensions["D"].width = 75

out = ROOT / "05-financer" / "outils" / "simulateur_pret.xlsx"; wb.save(out); print("->", out.relative_to(ROOT))
