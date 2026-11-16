#!/usr/bin/env python3
"""Génère 09-montage-fiscal-et-structures/outils/simulateur_structures.xlsx.

Feuille 1 : comparaison sur 15 ans de trois véhicules pour le même bien :
  - nom propre, location nue au régime réel (déficit foncier),
  - nom propre, LMNP au réel (amortissements, art. 39 C CGI),
  - SCI à l'IS (amortissements, IS 15 % / 25 %, CRL ou CFE, double imposition à la revente).
Feuille 2 : plan d'amortissement par composant d'une SCI à l'IS.

Cellules jaunes : à saisir. Cellules bleues : calculées. Toutes les sorties sont des formules Excel.
Règles fiscales à la date du 15 novembre 2026, à vérifier avant toute décision.
"""
import pathlib
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = pathlib.Path(__file__).resolve().parent.parent
wb = Workbook()
ws = wb.active
ws.title = "Comparaison 15 ans"

blue = Font(color="1F3A5F", bold=True)
inp = PatternFill("solid", fgColor="FFF4D6")
res = PatternFill("solid", fgColor="E4EBF3")
head = PatternFill("solid", fgColor="D5DEE9")
thin = Side(style="thin", color="B8C4D0")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
EUR = '#,##0 €'
PCT = '0.0%'


def prow(r, label, value=None, fmt=EUR, fill=None, note=""):
    ws.cell(r, 1, label).border = box
    c = ws.cell(r, 2, value)
    c.border = box
    c.number_format = fmt
    if fill:
        c.fill = fill
    if note:
        ws.cell(r, 3, note).font = Font(italic=True, color="666666")


# ---------------------------------------------------------------- paramètres
ws["A1"] = "Simulateur : nom propre location nue (réel) vs LMNP réel vs SCI à l'IS, sur 15 ans"
ws["A1"].font = blue
ws["A2"] = ("Cellules jaunes : à remplir. Cellules bleues : calculées. Règles à la date du 15/11/2026, "
            "à vérifier : LF 2025 (amortissements LMNP réintégrés dans la plus-value), IS 15 % jusqu'à 42 500 € puis 25 %, "
            "PS 17,2 %, déficit foncier 10 700 €, CRL 2,5 %.")
ws["A2"].font = Font(italic=True)
ws["A2"].alignment = Alignment(wrap_text=True)
ws.merge_cells("A2:N2")
ws.row_dimensions[2].height = 30

ws["A4"] = "1. Le bien"; ws["A4"].font = blue
prow(5, "Prix d'achat (frais d'agence inclus)", 150000, EUR, inp)
prow(6, "Part du terrain dans le prix (non amortissable)", 0.15, PCT, inp, "10 à 20 % en général, plus en centre-ville ; pratique comptable, à justifier")
prow(7, "Taux de frais de notaire", 0.08, PCT, inp, "Ancien : 7,5 à 8 % en 2026")
prow(8, "Frais de notaire", "=B5*B7", EUR, res)
prow(9, "Travaux (déductibles en nue, amortis en meublé et à l'IS)", 20000, EUR, inp, "Travaux d'amélioration ; pas de construction ni d'agrandissement en nue")
prow(10, "Mobilier (meublé uniquement, payé sur fonds propres)", 6000, EUR, inp)
prow(11, "Loyer annuel hors charges en location nue", 11400, EUR, inp)
prow(12, "Loyer annuel hors charges en meublé", 13800, EUR, inp, "Souvent 15 à 25 % au-dessus du loyer nu")
prow(13, "Charges non récupérables annuelles (taxe foncière, PNO, copropriété, gestion)", 2200, EUR, inp)
prow(14, "Revalorisation annuelle des loyers et charges", 0.01, PCT, inp)
prow(15, "Revalorisation annuelle de la valeur du bien", 0.01, PCT, inp, "Sert au calcul de revente en année 15")

ws["A17"] = "2. Le financement"; ws["A17"].font = blue
prow(18, "Apport (hors mobilier)", 15000, EUR, inp)
prow(19, "Montant emprunté", "=B5+B8+B9-B18", EUR, res, "Prix + notaire + travaux, moins l'apport")
prow(20, "Taux nominal annuel", 0.035, '0.00%', inp)
prow(21, "Assurance emprunteur (taux annuel sur capital initial)", 0.0025, '0.00%', inp)
prow(22, "Durée du prêt en années", 25, '0', inp, "25 ans maximum (HCSF, janvier 2022)")
prow(23, "Mensualité hors assurance", "=-PMT(B20/12,B22*12,B19)", EUR, res)

ws["A25"] = "3. La fiscalité du foyer et des véhicules"; ws["A25"].font = blue
prow(26, "Tranche marginale d'imposition (TMI)", 0.30, PCT, inp, "0, 11, 30, 41 ou 45 % (barème 2026)")
prow(27, "Prélèvements sociaux", 0.172, PCT, inp, "17,2 % (2026)")
prow(28, "Plafond d'imputation du déficit foncier sur le revenu global", 10700, EUR, inp, "Art. 156 I 3° CGI ; 21 400 € pour travaux énergétiques payés 2023-2025, à vérifier")
prow(29, "Taux réduit d'IS", 0.15, PCT, inp, "Art. 219 I b CGI")
prow(30, "Seuil du taux réduit", 42500, EUR, inp)
prow(31, "Taux normal d'IS", 0.25, PCT, inp)
prow(32, "Flat tax sur dividendes (PFU)", 0.314, PCT, inp, "12,8 % IR + 18,6 % PS depuis 2026, art. 200 A CGI")
prow(33, "Expert-comptable LMNP (par an)", 600, EUR, inp)
prow(34, "Expert-comptable SCI à l'IS (par an)", 1300, EUR, inp)
prow(35, "CFE annuelle en meublé", 350, EUR, inp, "Base minimum communale ; due en meublé, pas en location nue d'habitation (art. 1447 CGI)")
prow(36, "CRL (SCI à l'IS, location nue, immeuble > 15 ans)", 0.025, PCT, inp, "Art. 234 nonies CGI")
prow(37, "Immeuble achevé depuis plus de 15 ans ? (oui/non)", "oui", '@', inp)
prow(38, "SCI à l'IS exploitée en meublé ? (oui/non)", "non", '@', inp, "non = location nue (CRL) ; oui = meublé (CFE, mobilier amorti)")
prow(39, "Durée d'amortissement du bâti (années)", 30, '0', inp, "Pratique : 25 à 40 ans hors terrain (voir feuille 2 pour le détail par composant)")
prow(40, "Durée d'amortissement des travaux (années)", 10, '0', inp)
prow(41, "Durée d'amortissement du mobilier (années)", 7, '0', inp, "5 à 10 ans")
prow(42, "Amortissement annuel du bâti", "=B5*(1-B6)/B39", EUR, res)
prow(43, "Amortissement annuel des travaux", "=B9/B40", EUR, res)
prow(44, "Amortissement annuel du mobilier", "=B10/B41", EUR, res)

ws.column_dimensions["A"].width = 64
ws.column_dimensions["B"].width = 16
ws.column_dimensions["C"].width = 70

# ---------------------------------------------------------------- tableau annuel
HDR = 47
FIRST = 48
N = 15
LAST = FIRST + N - 1

cols = [
    # (clé, en-tête, groupe)
    ("an", "Année", "Commun"),
    ("loy_n", "Loyer nu", "Commun"),
    ("loy_m", "Loyer meublé", "Commun"),
    ("chg", "Charges non récup.", "Commun"),
    ("int", "Intérêts", "Commun"),
    ("ass", "Assurance", "Commun"),
    ("cap", "Capital remboursé", "Commun"),
    ("mens", "Annuité totale", "Commun"),
    # nue
    ("n_rhi", "Résultat hors intérêts", "Nue réel"),
    ("n_cf", "Charges financières", "Nue réel"),
    ("n_glob", "Déficit imputé revenu global", "Nue réel"),
    ("n_cree", "Déficit reportable créé", "Nue réel"),
    ("n_stk0", "Report début", "Nue réel"),
    ("n_net", "Revenu net avant report", "Nue réel"),
    ("n_use", "Report utilisé", "Nue réel"),
    ("n_imp", "Revenu imposable", "Nue réel"),
    ("n_stk1", "Report fin", "Nue réel"),
    ("n_tax", "Impôt (IR+PS) net d'économie", "Nue réel"),
    ("n_cf1", "Trésorerie de l'année", "Nue réel"),
    ("n_cum", "Trésorerie cumulée", "Nue réel"),
    # lmnp
    ("l_r0", "Résultat avant amort.", "LMNP réel"),
    ("l_am", "Amortissement de l'année", "LMNP réel"),
    ("l_dstk0", "Déficit début", "LMNP réel"),
    ("l_astk0", "Amort. différé début", "LMNP réel"),
    ("l_base", "Base après déficit", "LMNP réel"),
    ("l_amd", "Amort. déduit", "LMNP réel"),
    ("l_dstk1", "Déficit fin", "LMNP réel"),
    ("l_astk1", "Amort. différé fin", "LMNP réel"),
    ("l_imp", "Revenu imposable", "LMNP réel"),
    ("l_tax", "Impôt (IR+PS)", "LMNP réel"),
    ("l_cf1", "Trésorerie de l'année", "LMNP réel"),
    ("l_cum", "Trésorerie cumulée", "LMNP réel"),
    ("l_amcum", "Amort. déduits cumulés", "LMNP réel"),
    # sci is
    ("s_loy", "Loyer retenu", "SCI IS"),
    ("s_crl", "CRL", "SCI IS"),
    ("s_cfe", "CFE", "SCI IS"),
    ("s_am", "Amortissements", "SCI IS"),
    ("s_rf", "Résultat fiscal", "SCI IS"),
    ("s_dstk0", "Déficit début", "SCI IS"),
    ("s_imp", "Résultat imposable", "SCI IS"),
    ("s_dstk1", "Déficit fin", "SCI IS"),
    ("s_is", "IS", "SCI IS"),
    ("s_rn", "Résultat net comptable", "SCI IS"),
    ("s_cf1", "Trésorerie société de l'année", "SCI IS"),
    ("s_cum", "Trésorerie cumulée (vue associé)", "SCI IS"),
    ("s_dist", "Bénéfices distribuables cumulés", "SCI IS"),
    ("s_net", "Cumul net associé si tout distribué", "SCI IS"),
]
L = {}
for i, (k, h, g) in enumerate(cols, start=1):
    L[k] = get_column_letter(i)

ws.cell(HDR - 1, 1, "4. Tableau annuel (toutes les cellules sont des formules)").font = blue
# en-têtes de groupe
cur = None
for i, (k, h, g) in enumerate(cols, start=1):
    c = ws.cell(HDR - 2, i, g if g != cur else "")
    cur = g
    c.font = Font(bold=True, color="1F3A5F", size=9)
    c = ws.cell(HDR, i, h)
    c.font = Font(bold=True, size=9)
    c.fill = head
    c.border = box
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if i > 1:
        ws.column_dimensions[get_column_letter(i)].width = 14
ws.row_dimensions[HDR].height = 42


def f(k, r):
    return f"{L[k]}{r}"


for idx in range(N):
    r = FIRST + idx
    p = r - 1  # ligne précédente
    first = idx == 0
    F = {}
    F["an"] = idx + 1
    F["loy_n"] = f"=$B$11*(1+$B$14)^({f('an', r)}-1)"
    F["loy_m"] = f"=$B$12*(1+$B$14)^({f('an', r)}-1)"
    F["chg"] = f"=$B$13*(1+$B$14)^({f('an', r)}-1)"
    F["int"] = f"=IF({f('an', r)}>$B$22,0,-CUMIPMT($B$20/12,$B$22*12,$B$19,({f('an', r)}-1)*12+1,{f('an', r)}*12,0))"
    F["ass"] = f"=IF({f('an', r)}>$B$22,0,$B$19*$B$21)"
    F["cap"] = f"=IF({f('an', r)}>$B$22,0,-CUMPRINC($B$20/12,$B$22*12,$B$19,({f('an', r)}-1)*12+1,{f('an', r)}*12,0))"
    F["mens"] = f"={f('int', r)}+{f('ass', r)}+{f('cap', r)}"
    # --- nue réel
    F["n_rhi"] = f"={f('loy_n', r)}-{f('chg', r)}-IF({f('an', r)}=1,$B$9,0)"
    F["n_cf"] = f"={f('int', r)}+{f('ass', r)}"
    F["n_glob"] = f"=IF({f('n_rhi', r)}<0,MIN($B$28,-{f('n_rhi', r)}),0)"
    F["n_cree"] = (f"=IF({f('n_rhi', r)}<0,(-{f('n_rhi', r)}-{f('n_glob', r)})+{f('n_cf', r)},"
                   f"MAX(0,{f('n_cf', r)}-{f('n_rhi', r)}))")
    F["n_stk0"] = "0" if first else f"={f('n_stk1', p)}"
    F["n_net"] = f"=IF({f('n_rhi', r)}<0,0,MAX(0,{f('n_rhi', r)}-{f('n_cf', r)}))"
    F["n_use"] = f"=MIN({f('n_stk0', r)},{f('n_net', r)})"
    F["n_imp"] = f"={f('n_net', r)}-{f('n_use', r)}"
    F["n_stk1"] = f"={f('n_stk0', r)}-{f('n_use', r)}+{f('n_cree', r)}"
    F["n_tax"] = f"={f('n_imp', r)}*($B$26+$B$27)-{f('n_glob', r)}*$B$26"
    F["n_cf1"] = f"={f('loy_n', r)}-{f('chg', r)}-{f('mens', r)}-{f('n_tax', r)}"
    F["n_cum"] = (f"=-$B$18+{f('n_cf1', r)}" if first else f"={f('n_cum', p)}+{f('n_cf1', r)}")
    # --- LMNP réel
    F["l_r0"] = (f"={f('loy_m', r)}-{f('chg', r)}-$B$33-$B$35-{f('int', r)}-{f('ass', r)}"
                 f"-IF({f('an', r)}=1,$B$8,0)")
    F["l_am"] = (f"=IF({f('an', r)}<=$B$39,$B$42,0)+IF({f('an', r)}<=$B$40,$B$43,0)"
                 f"+IF({f('an', r)}<=$B$41,$B$44,0)")
    F["l_dstk0"] = "0" if first else f"={f('l_dstk1', p)}"
    F["l_astk0"] = "0" if first else f"={f('l_astk1', p)}"
    F["l_base"] = f"=IF({f('l_r0', r)}<0,0,MAX(0,{f('l_r0', r)}-{f('l_dstk0', r)}))"
    F["l_amd"] = f"=MIN({f('l_am', r)}+{f('l_astk0', r)},{f('l_base', r)})"
    F["l_dstk1"] = (f"=IF({f('l_r0', r)}<0,{f('l_dstk0', r)}-{f('l_r0', r)},"
                    f"{f('l_dstk0', r)}-MIN({f('l_dstk0', r)},{f('l_r0', r)}))")
    F["l_astk1"] = f"={f('l_am', r)}+{f('l_astk0', r)}-{f('l_amd', r)}"
    F["l_imp"] = f"={f('l_base', r)}-{f('l_amd', r)}"
    F["l_tax"] = f"={f('l_imp', r)}*($B$26+$B$27)"
    F["l_cf1"] = f"={f('loy_m', r)}-{f('chg', r)}-$B$33-$B$35-{f('mens', r)}-{f('l_tax', r)}"
    F["l_cum"] = (f"=-$B$18-$B$10+{f('l_cf1', r)}" if first else f"={f('l_cum', p)}+{f('l_cf1', r)}")
    F["l_amcum"] = (f"={f('l_amd', r)}" if first else f"={f('l_amcum', p)}+{f('l_amd', r)}")
    # --- SCI IS
    F["s_loy"] = f"=IF($B$38=\"oui\",{f('loy_m', r)},{f('loy_n', r)})"
    F["s_crl"] = f"=IF(AND($B$38<>\"oui\",$B$37=\"oui\"),{f('s_loy', r)}*$B$36,0)"
    F["s_cfe"] = f"=IF($B$38=\"oui\",$B$35,0)"
    F["s_am"] = (f"=IF({f('an', r)}<=$B$39,$B$42,0)+IF({f('an', r)}<=$B$40,$B$43,0)"
                 f"+IF(AND($B$38=\"oui\",{f('an', r)}<=$B$41),$B$44,0)")
    F["s_rf"] = (f"={f('s_loy', r)}-{f('chg', r)}-$B$34-{f('s_crl', r)}-{f('s_cfe', r)}-{f('int', r)}"
                 f"-{f('ass', r)}-IF({f('an', r)}=1,$B$8,0)-{f('s_am', r)}")
    F["s_dstk0"] = "0" if first else f"={f('s_dstk1', p)}"
    F["s_imp"] = f"=MAX(0,{f('s_rf', r)}-{f('s_dstk0', r)})"
    F["s_dstk1"] = (f"=IF({f('s_rf', r)}<0,{f('s_dstk0', r)}-{f('s_rf', r)},"
                    f"{f('s_dstk0', r)}-MIN({f('s_dstk0', r)},{f('s_rf', r)}))")
    F["s_is"] = f"=MIN({f('s_imp', r)},$B$30)*$B$29+MAX(0,{f('s_imp', r)}-$B$30)*$B$31"
    F["s_rn"] = f"={f('s_rf', r)}-{f('s_is', r)}"
    F["s_cf1"] = (f"={f('s_loy', r)}-{f('chg', r)}-$B$34-{f('s_crl', r)}-{f('s_cfe', r)}"
                  f"-{f('mens', r)}-{f('s_is', r)}")
    F["s_cum"] = (f"=-$B$18-IF($B$38=\"oui\",$B$10,0)+{f('s_cf1', r)}" if first
                  else f"={f('s_cum', p)}+{f('s_cf1', r)}")
    F["s_dist"] = (f"={f('s_rn', r)}" if first else f"={f('s_dist', p)}+{f('s_rn', r)}")
    F["s_net"] = (f"={f('s_cum', r)}-$B$32*MAX(0,MIN({f('s_dist', r)},"
                  f"{f('s_cum', r)}+$B$18+IF($B$38=\"oui\",$B$10,0)))")
    for i, (k, h, g) in enumerate(cols, start=1):
        c = ws.cell(r, i, F[k])
        c.border = box
        c.number_format = '0' if k == "an" else EUR
        c.fill = res
        c.font = Font(size=9)

# ---------------------------------------------------------------- synthèse
S = LAST + 3
ws.cell(S, 1, "5. Synthèse sur 15 ans").font = blue
hdrs = ["", "Nue réel (nom propre)", "LMNP réel (nom propre)", "SCI à l'IS", "Note"]
for j, h in enumerate(hdrs, start=1):
    c = ws.cell(S + 1, j, h)
    c.font = Font(bold=True)
    c.fill = head
    c.border = box
ws.column_dimensions["D"].width = 16
ws.column_dimensions["E"].width = 60


def srow(r, label, fn, fl, fs, note="", fmt=EUR):
    ws.cell(r, 1, label).border = box
    for j, v in enumerate([fn, fl, fs], start=2):
        c = ws.cell(r, j, v)
        c.border = box
        c.number_format = fmt
        c.fill = res
    if note:
        ws.cell(r, 5, note).font = Font(italic=True, color="666666")


rng = lambda k: f"{L[k]}{FIRST}:{L[k]}{LAST}"
r = S + 2
srow(r, "Impôt cumulé sur les loyers (15 ans)", f"=SUM({rng('n_tax')})", f"=SUM({rng('l_tax')})",
     f"=SUM({rng('s_is')})", "SCI : IS seul ; la flat tax à la distribution est comptée plus bas")
r += 1
srow(r, "Trésorerie nette cumulée fin d'année 15 (apport compris)", f"={f('n_cum', LAST)}",
     f"={f('l_cum', LAST)}", f"={f('s_cum', LAST)}",
     "SCI : argent dans la société, pas dans votre poche (cash trap)")
r += 1
srow(r, "Dont, pour la SCI, net pour l'associé si tout est distribué", "", "", f"={f('s_net', LAST)}",
     "Dividendes limités aux bénéfices distribuables ; apport remboursable sans impôt via compte courant")
r += 2
ws.cell(r, 1, "Revente à la fin de l'année 15 (hypothèses simplifiées)").font = blue
r += 1
R_PRIX = r
srow(r, "Prix de revente estimé", "=$B$5*(1+$B$15)^15", "=$B$5*(1+$B$15)^15", "=$B$5*(1+$B$15)^15",
     "Prix × revalorisation annuelle sur 15 ans")
r += 1
R_CRD = r
srow(r, "Capital restant dû", f"=$B$19-SUM({rng('cap')})", f"=$B$19-SUM({rng('cap')})",
     f"=$B$19-SUM({rng('cap')})")
r += 1
R_PV = r
srow(r, "Plus-value imposable brute",
     f"=MAX(0,B{R_PRIX}-$B$5*(1+MAX($B$7,0.075))-$B$5*0.15)",
     f"=MAX(0,C{R_PRIX}-($B$5*(1+MAX($B$7,0.075))+$B$5*0.15-{f('l_amcum', LAST)}))",
     f"=MAX(0,D{R_PRIX}-($B$5+$B$9+IF($B$38=\"oui\",$B$10,0)-SUM({rng('s_am')})))",
     "Particuliers : forfait notaire 7,5 % et travaux 15 % (art. 150 VB CGI) ; LMNP : amortissements déduits "
     "retranchés (LF 2025) ; IS : prix moins valeur nette comptable")
r += 1
R_TAX = r
srow(r, "Impôt sur la plus-value",
     f"=B{R_PV}*(1-MIN(1,10*0.06))*0.19+B{R_PV}*(1-MIN(1,10*0.0165))*$B$27",
     f"=C{R_PV}*(1-MIN(1,10*0.06))*0.19+C{R_PV}*(1-MIN(1,10*0.0165))*$B$27",
     f"=MIN(D{R_PV},$B$30)*$B$29+MAX(0,D{R_PV}-$B$30)*$B$31",
     "Particuliers après 15 ans : abattement IR 60 % (6 %/an de la 6e à la 15e année), PS 16,5 % ; "
     "surtaxe au-delà de 50 000 € non modélisée. IS : taux réduit puis normal")
r += 1
R_NET = r
srow(r, "Net de revente après impôt et remboursement du prêt",
     f"=B{R_PRIX}-B{R_CRD}-B{R_TAX}", f"=C{R_PRIX}-C{R_CRD}-C{R_TAX}",
     f"=D{R_PRIX}-D{R_CRD}-D{R_TAX}", "SCI : encore dans la société à ce stade")
r += 1
R_BONI = r
srow(r, "SCI : flat tax à la liquidation (boni distribué)", "", "",
     f"=$B$32*MAX(0,D{R_NET}+{f('s_cum', LAST)})",
     "31,4 % sur tout ce qui dépasse les apports (deuxième étage d'imposition)")
r += 1
R_TOT = r
srow(r, "Résultat total net sur 15 ans (trésorerie cumulée + revente nette)",
     f"=B{S+3}+B{R_NET}", f"=C{S+3}+C{R_NET}",
     f"=D{S+3}+D{R_NET}-D{R_BONI}",
     "Ce que vous avez gagné, net de tout, par rapport à votre mise de départ")
r += 1
srow(r, "Classement (1 = meilleur)", f"=RANK(B{R_TOT},$B${R_TOT}:$D${R_TOT})",
     f"=RANK(C{R_TOT},$B${R_TOT}:$D${R_TOT})", f"=RANK(D{R_TOT},$B${R_TOT}:$D${R_TOT})",
     "Change avec la TMI, la durée et le prix de revente : testez plusieurs hypothèses", fmt='0')
r += 2
ws.cell(r, 1, ("Limites : surtaxe plus-value > 50 000 € non modélisée ; déficit foncier limité à 10 ans de report non "
               "modélisé ; frais de notaire passés en charge l'année 1 en LMNP et SCI IS (option) ; CSG déductible non "
               "modélisée ; pas d'IRA ; revente supposée exactement en année 15.")).font = Font(italic=True, color="666666")

ws.freeze_panes = ws.cell(FIRST, 2)

# ============================================================ feuille 2
a = wb.create_sheet("Amortissement SCI IS")
a["A1"] = "Plan d'amortissement par composant (SCI à l'IS ou LMNP au réel)"
a["A1"].font = blue
a["A2"] = ("Le terrain ne s'amortit pas. Chaque composant a sa durée. Les pourcentages et durées sont des pratiques "
           "comptables courantes (à faire valider par votre expert-comptable), pas des règles légales.")
a["A2"].font = Font(italic=True)
a.merge_cells("A2:H2")


def arow(r, label, value=None, fmt=EUR, fill=None, note=""):
    a.cell(r, 1, label).border = box
    c = a.cell(r, 2, value)
    c.border = box
    c.number_format = fmt
    if fill:
        c.fill = fill
    if note:
        a.cell(r, 3, note).font = Font(italic=True, color="666666")


arow(4, "Prix d'achat du bien (hors frais)", "='Comparaison 15 ans'!B5", EUR, inp, "Repris de la feuille 1, modifiable")
arow(5, "Frais d'acquisition (notaire, agence si non inclus)", "='Comparaison 15 ans'!B8", EUR, inp)
arow(6, "Frais d'acquisition : charge immédiate (oui) ou amortis avec le gros œuvre (non) ?", "oui", '@', inp,
     "Option comptable (art. 38 quinquies annexe III CGI) : en charge, ils creusent le déficit de l'année 1")
arow(7, "Travaux", "='Comparaison 15 ans'!B9", EUR, inp)
arow(8, "Mobilier", "='Comparaison 15 ans'!B10", EUR, inp, "0 en location nue")

a["A10"] = "Composants du bien"; a["A10"].font = blue
heads = ["Composant", "Part du prix", "Base amortissable", "Durée (ans)", "Dotation annuelle", "Note"]
for j, h in enumerate(heads, start=1):
    c = a.cell(11, j, h); c.font = Font(bold=True); c.fill = head; c.border = box
comps = [
    ("Terrain (non amortissable)", 0.15, 0, "10 à 20 % ; plus en ville"),
    ("Gros œuvre (structure, murs, toiture)", 0.40, 50, "40 à 80 ans selon les cabinets"),
    ("Façades, étanchéité", 0.10, 30, "20 à 30 ans"),
    ("Installations générales et techniques (électricité, plomberie, chauffage)", 0.20, 20, "15 à 25 ans"),
    ("Agencements intérieurs (cuisine, sols, cloisons)", 0.15, 15, "10 à 15 ans"),
]
CR0 = 12
for i, (name, part, dur, note) in enumerate(comps):
    r = CR0 + i
    a.cell(r, 1, name).border = box
    c = a.cell(r, 2, part); c.number_format = PCT; c.fill = inp; c.border = box
    base = f"=$B$4*B{r}" + (f"+IF($B$6=\"oui\",0,$B$5)" if i == 1 else "")
    c = a.cell(r, 3, base); c.number_format = EUR; c.fill = res; c.border = box
    c = a.cell(r, 4, dur); c.number_format = '0'; c.fill = inp; c.border = box
    c = a.cell(r, 5, f"=IF(D{r}=0,0,C{r}/D{r})"); c.number_format = EUR; c.fill = res; c.border = box
    a.cell(r, 6, note).font = Font(italic=True, color="666666")
rT = CR0 + len(comps)
a.cell(rT, 1, "Travaux").border = box
c = a.cell(rT, 3, "=$B$7"); c.number_format = EUR; c.fill = res; c.border = box
c = a.cell(rT, 4, 10); c.fill = inp; c.border = box
c = a.cell(rT, 5, f"=IF(D{rT}=0,0,C{rT}/D{rT})"); c.number_format = EUR; c.fill = res; c.border = box
a.cell(rT, 6, "10 à 15 ans ; c'est la fin de cette ligne qui crée la « falaise » de l'année 11").font = Font(italic=True, color="666666")
rM = rT + 1
a.cell(rM, 1, "Mobilier").border = box
c = a.cell(rM, 3, "=$B$8"); c.number_format = EUR; c.fill = res; c.border = box
c = a.cell(rM, 4, 7); c.fill = inp; c.border = box
c = a.cell(rM, 5, f"=IF(D{rM}=0,0,C{rM}/D{rM})"); c.number_format = EUR; c.fill = res; c.border = box
a.cell(rM, 6, "5 à 10 ans").font = Font(italic=True, color="666666")
rS = rM + 1
a.cell(rS, 1, "Contrôle : total des parts du prix (doit faire 100 %)").border = box
c = a.cell(rS, 2, f"=SUM(B{CR0}:B{CR0+len(comps)-1})"); c.number_format = PCT; c.fill = res; c.border = box
c = a.cell(rS, 5, f"=SUM(E{CR0}:E{rM})"); c.number_format = EUR; c.fill = res; c.border = box
a.cell(rS, 4, "Dotation totale année 1").font = Font(bold=True)

# tableau annuel
Y0 = rS + 3
a.cell(Y0 - 1, 1, "Plan annuel : dotation par composant, total et valeur nette comptable").font = blue
lines = list(range(CR0, rM + 1))  # toutes les lignes amortissables (terrain inclus, dotation 0)
yh = ["Année"] + [a.cell(rr, 1).value.split(" (")[0][:28] for rr in lines] + ["Dotation totale", "Amortissements cumulés", "Valeur nette comptable"]
for j, h in enumerate(yh, start=1):
    c = a.cell(Y0, j, h); c.font = Font(bold=True, size=9); c.fill = head; c.border = box
    c.alignment = Alignment(wrap_text=True, vertical="top")
    a.column_dimensions[get_column_letter(j)].width = 15
a.row_dimensions[Y0].height = 40
NY = 50
nL = len(lines)
colTot = get_column_letter(nL + 2)
colCum = get_column_letter(nL + 3)
colVNC = get_column_letter(nL + 4)
for y in range(1, NY + 1):
    r = Y0 + y
    a.cell(r, 1, y).border = box
    for j, rr in enumerate(lines, start=2):
        c = a.cell(r, j, f"=IF(AND($D${rr}>0,A{r}<=$D${rr}),$E${rr},0)")
        c.number_format = EUR; c.fill = res; c.border = box; c.font = Font(size=9)
    c = a.cell(r, nL + 2, f"=SUM(B{r}:{get_column_letter(nL+1)}{r})"); c.number_format = EUR; c.fill = res; c.border = box
    c = a.cell(r, nL + 3, f"={colTot}{r}" if y == 1 else f"={colCum}{r-1}+{colTot}{r}"); c.number_format = EUR; c.fill = res; c.border = box
    c = a.cell(r, nL + 4, f"=SUM($C${CR0}:$C${rM})-{colCum}{r}"); c.number_format = EUR; c.fill = res; c.border = box
a.column_dimensions["A"].width = 60
a.column_dimensions["F"].width = 50
a.freeze_panes = a.cell(Y0 + 1, 2)

out = ROOT / "09-montage-fiscal-et-structures" / "outils" / "simulateur_structures.xlsx"
out.parent.mkdir(parents=True, exist_ok=True)
wb.save(out)
print("->", out.relative_to(ROOT))
