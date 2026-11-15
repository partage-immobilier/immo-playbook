#!/usr/bin/env python3
"""Génère 02-visiter-et-chiffrer/outils/calculateur_rendement.xlsx : rendement brut, net, net-net, cash-flow mensuel selon le crédit, et TRI sur 10 ans avec une hypothèse de revente. Toutes les cellules de résultat sont des formules Excel. Cellules jaunes à saisir, bleues calculées, notes en colonne C."""
import pathlib
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

ROOT = pathlib.Path(__file__).resolve().parent.parent
wb = Workbook(); ws = wb.active; ws.title = "Calculateur"
blue = Font(color="1F3A5F", bold=True); inp = PatternFill("solid", fgColor="FFF4D6"); res = PatternFill("solid", fgColor="E4EBF3")
thin = Side(style="thin", color="B8C4D0"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
EUR, PCT, PCT2 = '#,##0 €', '0.0%', '0.00%'

def row(sheet, r, label, value=None, fmt=EUR, fill=None, note=""):
    sheet.cell(r, 1, label).border = box
    c = sheet.cell(r, 2, value); c.border = box; c.number_format = fmt
    if fill: c.fill = fill
    if note: sheet.cell(r, 3, note).font = Font(italic=True, color="666666")

def title(sheet, r, text):
    sheet.cell(r, 1, text).font = blue

ws["A1"] = "Calculateur de rendement (module 2) : brut, net, net-net, cash-flow et TRI. Exemple pré-rempli : le T3 de Mehdi."; ws["A1"].font = blue
ws["A2"] = "Cellules jaunes : à remplir. Cellules bleues : calculées. Les taux et fourchettes sont ceux de novembre 2026, à vérifier."; ws["A2"].font = Font(italic=True)

title(ws, 4, "1. Achat")
row(ws, 5, "Prix d'achat négocié (frais d'agence inclus)", 82000, EUR, inp, "Prix FAI ; si les honoraires s'ajoutent, les saisir en B6")
row(ws, 6, "Frais d'agence en plus (si non inclus dans le prix)", 0, EUR, inp)
row(ws, 7, "Taux de frais de notaire", 0.08, PCT, inp, "Ancien : 7,5 à 8 % en 2026 ; neuf : 2 à 3 %")
row(ws, 8, "Frais de notaire", "=(B5+B6)*B7", EUR, res)
row(ws, 9, "Travaux (devis ou fourchettes au m² du cours, imprévus compris)", 33600, EUR, inp, "Rafraîchissement 150 à 300 €/m², moyenne 500 à 900, lourde 1 000 à 1 500 (2026, à vérifier)")
row(ws, 10, "Mobilier et équipement (si meublé)", 0, EUR, inp)
row(ws, 11, "Frais de dossier bancaire et garantie du prêt", 1800, EUR, inp, "Module 5")
row(ws, 12, "Coût total de l'opération", "=B5+B6+B8+B9+B10+B11", EUR, res)
row(ws, 13, "Surface habitable (m²)", 59, '0.0', inp, "Mesurée à la visite, pas la surface au sol")
row(ws, 14, "Prix de revient au m² (prix + travaux, hors notaire)", "=IF(B13>0,(B5+B6+B9)/B13,0)", EUR, res, "À comparer à la médiane DVF des biens rénovés")
row(ws, 15, "Médiane DVF du quartier (€/m², facultatif)", 1700, EUR, inp, "app.dvf.etalab.gouv.fr, ventes de moins de deux ans")
row(ws, 16, "Écart du prix de revient par rapport à la médiane DVF", "=IF(B15>0,B14/B15-1,0)", PCT, res, "Positif : vous paierez plus cher que ce que vaut le bien à la revente")

title(ws, 18, "2. Loyer et charges")
row(ws, 19, "Loyer mensuel hors charges (estimation prudente)", 600, EUR, inp, "Médiane de dix annonces comparables, moins 5 à 10 %")
row(ws, 20, "Loyer annuel", "=B19*12", EUR, res)
row(ws, 21, "Vacance locative (semaines par an)", 4, '0', inp, "Ville moyenne : 4 semaines ; zone tendue : 2 ; immeuble : plus")
row(ws, 22, "Perte de loyer liée à la vacance", "=B20*B21/52", EUR, res)
row(ws, 23, "Charges de copropriété annuelles (total)", 1100, EUR, inp, "Relevé de charges ou dernier appel de fonds annuel")
row(ws, 24, "Part non récupérable sur le locataire", 0.40, PCT, inp, "Syndic, assurance de l'immeuble, gros entretien : souvent 30 à 50 %")
row(ws, 25, "Charges non récupérables", "=B23*B24", EUR, res)
row(ws, 26, "Taxe foncière (hors part ordures ménagères, récupérable)", 780, EUR, inp, "Dernier avis du vendeur")
row(ws, 27, "Assurance propriétaire non occupant (PNO)", 150, EUR, inp, "100 à 250 € par an")
row(ws, 28, "Gestion par une agence (% du loyer)", 0, PCT, inp, "0 si vous gérez vous-même ; 6 à 10 % sinon")
row(ws, 29, "Frais de gestion", "=B20*B28", EUR, res)
row(ws, 30, "Garantie loyers impayés (% du loyer)", 0.03, PCT, inp, "2,5 à 3,5 % ; 0 si caution solidaire ou Visale")
row(ws, 31, "Coût de la GLI", "=B20*B30", EUR, res)
row(ws, 32, "Entretien et petites réparations (% du prix par an)", 0.008, PCT, inp, "0,5 à 1 % du prix du bien")
row(ws, 33, "Provision d'entretien", "=B5*B32", EUR, res)
row(ws, 34, "Comptable (LMNP au réel, SCI à l'IS)", 0, EUR, inp, "400 à 1 000 € par an")
row(ws, 35, "Total des charges annuelles", "=B22+B25+B26+B27+B29+B31+B33+B34", EUR, res)
row(ws, 36, "Revenu net annuel (avant impôt et avant crédit)", "=B20-B35", EUR, res)
row(ws, 37, "Impôt annuel estimé sur les loyers", 0, EUR, inp, "LMNP réel : souvent 0 les premières années ; nu réel : (loyers − charges − intérêts) × (tranche + 17,2 %) ; SCI IS : 15 % puis 25 % après amortissements (2026)")
row(ws, 38, "Revenu net annuel après impôt", "=B36-B37", EUR, res)

title(ws, 40, "3. Rentabilités")
row(ws, 41, "Rendement brut sur le prix (le chiffre des annonces)", "=IF(B5>0,B20/B5,0)", PCT, res)
row(ws, 42, "Rendement brut sur le coût total", "=IF(B12>0,B20/B12,0)", PCT, res)
row(ws, 43, "Rendement net", "=IF(B12>0,B36/B12,0)", PCT, res, "Le seul chiffre qui compare deux biens ; indépendant du crédit et de l'impôt")
row(ws, 44, "Rendement net-net (après impôt)", "=IF(B12>0,B38/B12,0)", PCT, res)

title(ws, 46, "4. Crédit et cash-flow")
row(ws, 47, "Apport personnel", 12000, EUR, inp, "Souvent au moins les frais de notaire et de garantie")
row(ws, 48, "Montant emprunté", "=MAX(0,B12-B47)", EUR, res)
row(ws, 49, "Taux nominal annuel", 0.034, PCT2, inp, "Taux du jour")
row(ws, 50, "Durée (années)", 20, '0', inp, "25 ans maximum pour un particulier (HCSF, depuis janvier 2022)")
row(ws, 51, "Taux annuel de l'assurance emprunteur (sur capital initial)", 0.0025, PCT2, inp, "0,10 à 0,40 %")
row(ws, 52, "Mensualité hors assurance", "=IF(B48>0,-PMT(B49/12,B50*12,B48),0)", EUR, res)
row(ws, 53, "Assurance mensuelle", "=B48*B51/12", EUR, res)
row(ws, 54, "Mensualité totale", "=B52+B53", EUR, res)
row(ws, 55, "Cash-flow mensuel avant impôt", "=B36/12-B54", EUR, res, "Négatif = effort d'épargne mensuel")
row(ws, 56, "Cash-flow mensuel après impôt", "=B38/12-B54", EUR, res)
row(ws, 57, "Couverture de l'annuité par le revenu net (DSCR)", "=IF(B54>0,B36/(B54*12),0)", '0.00', res, "Les banques aiment voir 1,1 ou plus (pratique, pas une règle)")
row(ws, 58, "Capital remboursé la première année", "=IF(B48>0,-CUMPRINC(B49/12,B50*12,B48,1,12,0),0)", EUR, res, "L'épargne forcée que le cash-flow ne montre pas")

title(ws, 60, "5. Le même bien avec un autre crédit : le cash-flow dépend du financement, pas du bien")
ws.cell(61, 1, "Durée (années)").border = box
for col, years in zip("BCD", (15, 20, 25)):
    c = ws[f"{col}61"]; c.value = years; c.fill = inp; c.border = box; c.number_format = '0'
    m = ws[f"{col}62"]; m.value = f"=IF($B$48>0,-PMT($B$49/12,{col}61*12,$B$48),0)+$B$48*$B$51/12"; m.fill = res; m.border = box; m.number_format = EUR
    f = ws[f"{col}63"]; f.value = f"=$B$36/12-{col}62"; f.fill = res; f.border = box; f.number_format = EUR
    k = ws[f"{col}64"]; k.value = f"=IF($B$48>0,-CUMPRINC($B$49/12,{col}61*12,$B$48,1,12,0)/12,0)"; k.fill = res; k.border = box; k.number_format = EUR
ws.cell(62, 1, "Mensualité totale (assurance comprise)").border = box
ws.cell(63, 1, "Cash-flow mensuel avant impôt").border = box
ws.cell(64, 1, "Capital remboursé par mois la première année").border = box
ws.cell(65, 1, "Même rendement net (B43) dans les trois cas : seul le crédit change.").font = Font(italic=True, color="666666")
ws.column_dimensions["A"].width = 62; ws.column_dimensions["B"].width = 16; ws.column_dimensions["C"].width = 16; ws.column_dimensions["D"].width = 16

t = wb.create_sheet("TRI 10 ans")
t["A1"] = "TRI (taux de rentabilité interne) sur 10 ans avec une hypothèse de revente"; t["A1"].font = blue
t["A2"] = "Le TRI mesure le rendement annuel de l'argent réellement sorti de votre poche (apport, efforts mensuels), compte tenu de ce que vous récupérez à la revente."; t["A2"].font = Font(italic=True)
row(t, 4, "Évolution annuelle du prix du bien (hypothèse)", 0.01, PCT, inp, "0 % = prudent ; testez − 1 % et + 2 % : si le TRI ne tient qu'avec une hausse, vous pariez")
row(t, 5, "Évolution annuelle du loyer (hypothèse, indexation IRL)", 0.01, PCT, inp, "Loyers F et G gelés depuis août 2022")
row(t, 6, "Frais de revente (% du prix de revente)", 0.05, PCT, inp, "Agence 4 à 6 % ; 0 si vente entre particuliers")
row(t, 7, "Impôt estimé sur la plus-value à la revente", 0, EUR, inp, "Particulier : 19 % + 17,2 % après abattements (module 9) ; LMNP : amortissements réintégrés depuis 2025 ; SCI IS : sur la valeur amortie")
row(t, 8, "Valeur du bien aujourd'hui après travaux", "=Calculateur!B5+Calculateur!B6+Calculateur!B9", EUR, res, "Par défaut prix + travaux ; remplacez par votre estimation (médiane DVF rénové × surface)")
row(t, 9, "Prix de revente estimé à l'année 10", "=B8*(1+B4)^10", EUR, res)
row(t, 10, "Capital restant dû après 10 ans", "=IF(Calculateur!B48>0,IF(Calculateur!B50*12>120,-FV(Calculateur!B49/12,120,-Calculateur!B52,Calculateur!B48,0),0),0)", EUR, res, "Ce que vous remboursez à la banque le jour de la vente")
row(t, 11, "Frais de revente", "=B9*B6", EUR, res)
row(t, 12, "Produit net de la revente", "=B9-B10-B11-B7", EUR, res)

hdr = ["Année", "Loyers nets après impôt", "Annuité de crédit (assurance comprise)", "Cash-flow annuel", "Flux pour le TRI"]
for i, h in enumerate(hdr, 1):
    c = t.cell(14, i, h); c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F3A5F"); c.border = box
    c.alignment = Alignment(wrap_text=True, vertical="center")
t.row_dimensions[14].height = 32
# Année 0 : l'apport sort de votre poche
t.cell(15, 1, 0).border = box
for col in (2, 3, 4):
    c = t.cell(15, col, 0); c.border = box; c.number_format = EUR; c.fill = res
c = t.cell(15, 5, "=-Calculateur!B47"); c.border = box; c.number_format = EUR; c.fill = res
for y in range(1, 11):
    r = 15 + y
    t.cell(r, 1, y).border = box
    c = t.cell(r, 2, f"=Calculateur!$B$38*(1+$B$5)^({y}-1)"); c.number_format = EUR; c.fill = res; c.border = box
    c = t.cell(r, 3, "=Calculateur!$B$54*12"); c.number_format = EUR; c.fill = res; c.border = box
    c = t.cell(r, 4, f"=B{r}-C{r}"); c.number_format = EUR; c.fill = res; c.border = box
    flux = f"=D{r}+$B$12" if y == 10 else f"=D{r}"
    c = t.cell(r, 5, flux); c.number_format = EUR; c.fill = res; c.border = box
t.cell(25, 6, "← année 10 : cash-flow + produit net de la revente").font = Font(italic=True, color="666666")

row(t, 27, "TRI sur 10 ans", '=IFERROR(IRR(E15:E25),"sans objet")', PCT, res, "Fonction TRI d'Excel. « Sans objet » : aucun flux négatif (apport nul et cash-flow positif), vous n'avez rien investi")
row(t, 28, "Total des efforts (apport + cash-flows négatifs)", '=-SUMIF(E15:E25,"<0")', EUR, res, "L'argent réellement sorti de votre poche en dix ans")
row(t, 29, "Gain net cumulé sur 10 ans", "=SUM(E15:E25)", EUR, res, "Tout ce qui entre moins tout ce qui sort, revente comprise")
row(t, 30, "Capital remboursé par le crédit en 10 ans", "=Calculateur!B48-B10", EUR, res, "L'épargne forcée financée en grande partie par le locataire")
row(t, 32, "Rappel : rendement net du bien (indépendant du crédit)", "=Calculateur!B43", PCT, res)
t.column_dimensions["A"].width = 58; t.column_dimensions["B"].width = 20; t.column_dimensions["C"].width = 22
t.column_dimensions["D"].width = 18; t.column_dimensions["E"].width = 18; t.column_dimensions["F"].width = 40

out = ROOT / "02-visiter-et-chiffrer" / "outils" / "calculateur_rendement.xlsx"; wb.save(out); print("->", out.relative_to(ROOT))
