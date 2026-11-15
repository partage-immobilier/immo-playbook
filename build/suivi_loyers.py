#!/usr/bin/env python3
"""Génère 08-mettre-en-location/outils/suivi_loyers.xlsx : suivi des loyers et des charges par lot sur 12 mois
(loyer dû, encaissé, retard, charges, régularisation, taxe foncière, cash par lot et total). Toutes les sorties sont des formules Excel."""
import pathlib
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = pathlib.Path(__file__).resolve().parent.parent
NLOTS = 4
wb = Workbook()
blue = Font(color="1F3A5F", bold=True); inp = PatternFill("solid", fgColor="FFF4D6"); res = PatternFill("solid", fgColor="E4EBF3")
head = PatternFill("solid", fgColor="D5DEE9")
thin = Side(style="thin", color="B8C4D0"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
EUR = '#,##0 €'; DATE = 'DD/MM/YYYY'

# ---------------- Feuille 1 : paramètres des lots ----------------
pa = wb.active; pa.title = "Lots"
pa["A1"] = "Paramètres des lots (une colonne par lot). Cellules jaunes : à remplir. Cellules bleues : calculées."; pa["A1"].font = blue
pa["A2"] = "Loyers et provisions mensuels hors charges / charges ; montants annuels pour la taxe foncière, l'assurance, les charges réelles. Repères 2026 : GLI 2,5 à 3,5 % du loyer CC, gestion 6 à 8 % TTC, PNO 100 à 250 €/an."; pa["A2"].font = Font(italic=True, color="666666")
pa["A4"] = "Année de suivi"; pa["B4"] = 2026; pa["B4"].fill = inp
pa["A5"] = "Jour d'échéance du loyer dans le mois"; pa["B5"] = 5; pa["B5"].fill = inp
pa["A6"] = "Premier mois suivi (1er du mois)"; pa["B6"] = "=DATE(B4,1,1)"; pa["B6"].fill = res; pa["B6"].number_format = DATE
pa["C6"] = "Modifier la formule pour un suivi qui ne commence pas en janvier (ex. =DATE(B4,7,1))"; pa["C6"].font = Font(italic=True, color="666666")

params = [
    ("Nom du lot", ["Lot 1", "Lot 2", "Lot 3", "Lot 4"], '@', "Appartement, chambre de colocation..."),
    ("Mode (nue, meublé, colocation, courte durée)", ["meublé", "", "", ""], '@', ""),
    ("Loyer mensuel hors charges", [650, 0, 0, 0], EUR, "Montant du bail"),
    ("Provision ou forfait de charges mensuel", [60, 0, 0, 0], EUR, "Forfait possible en meublé, sinon provision avec régularisation annuelle"),
    ("Dépôt de garantie encaissé", [1300, 0, 0, 0], EUR, "1 mois HC en nue, 2 mois HC en meublé (loi de 1989)"),
    ("Charges récupérables réelles de l'année", [900, 0, 0, 0], EUR, "Décret 87-713 : eau, entretien des communs, ordures... à saisir en fin d'année"),
    ("Charges de copropriété non récupérables de l'année", [500, 0, 0, 0], EUR, "Honoraires de syndic, gros travaux, assurance immeuble"),
    ("Taxe foncière de l'année (TEOM comprise)", [950, 0, 0, 0], EUR, "Avis reçu en septembre-octobre"),
    ("dont taxe d'enlèvement des ordures ménagères (TEOM)", [150, 0, 0, 0], EUR, "Récupérable sur le locataire, déjà incluse dans les charges récupérables si vous la refacturez"),
    ("Assurance propriétaire non occupant (annuelle)", [150, 0, 0, 0], EUR, ""),
    ("Assurance loyers impayés (taux sur loyer CC, 0 si aucune)", [0.03, 0, 0, 0], '0.0%', "Non cumulable avec une caution, sauf étudiant ou apprenti"),
    ("Honoraires de gestion (taux sur encaissements, 0 si direct)", [0, 0, 0, 0], '0.0%', "6 à 8 % TTC en agence, 20 à 25 % en conciergerie courte durée"),
    ("Mensualité de crédit affectée au lot (assurance comprise)", [520, 0, 0, 0], EUR, "Au prorata si un prêt couvre plusieurs lots"),
    ("Entretien et petites réparations de l'année", [300, 0, 0, 0], EUR, "Chauffe-eau, robinetterie, remise en état"),
]
PR0 = 8
pa.cell(PR0 - 1, 1, "Paramètre").font = Font(bold=True)
for j in range(NLOTS):
    c = pa.cell(PR0 - 1, 2 + j, f"Lot {j + 1}"); c.font = Font(bold=True); c.fill = head; c.border = box
pa.cell(PR0 - 1, 2 + NLOTS, "Note").font = Font(bold=True)
for i, (label, vals, fmt, note) in enumerate(params):
    r = PR0 + i
    pa.cell(r, 1, label).border = box
    for j in range(NLOTS):
        c = pa.cell(r, 2 + j, vals[j]); c.fill = inp; c.border = box; c.number_format = fmt
    if note: pa.cell(r, 2 + NLOTS, note).font = Font(italic=True, color="666666")
pa.column_dimensions["A"].width = 58
for j in range(NLOTS): pa.column_dimensions[get_column_letter(2 + j)].width = 14
pa.column_dimensions[get_column_letter(2 + NLOTS)].width = 80
# lignes des paramètres (pour les références)
P = {name: PR0 + i for i, (name, *_ ) in enumerate(params)}
R_NOM, R_LOYER, R_PROV, R_CHREC, R_CHNR, R_TF, R_TEOM, R_PNO, R_GLI, R_GEST, R_CREDIT, R_ENT = (
    P["Nom du lot"], P["Loyer mensuel hors charges"], P["Provision ou forfait de charges mensuel"],
    P["Charges récupérables réelles de l'année"], P["Charges de copropriété non récupérables de l'année"],
    P["Taxe foncière de l'année (TEOM comprise)"], P["dont taxe d'enlèvement des ordures ménagères (TEOM)"],
    P["Assurance propriétaire non occupant (annuelle)"], P["Assurance loyers impayés (taux sur loyer CC, 0 si aucune)"],
    P["Honoraires de gestion (taux sur encaissements, 0 si direct)"], P["Mensualité de crédit affectée au lot (assurance comprise)"],
    P["Entretien et petites réparations de l'année"])

# ---------------- Feuille 2 : suivi mensuel ----------------
su = wb.create_sheet("Suivi mensuel")
su["A1"] = "Suivi mensuel des loyers par lot : loyer dû, encaissé, retard, charges. Cellules jaunes : à remplir à chaque encaissement."; su["A1"].font = blue
su["A2"] = "Saisir le montant encaissé et la date d'encaissement. Le retard se calcule par rapport à l'échéance (jour fixé dans la feuille Lots) et par rapport à aujourd'hui tant que rien n'est encaissé."; su["A2"].font = Font(italic=True, color="666666")
cols = ["Mois", "Échéance", "Loyer HC dû", "Charges dues", "Total dû", "Encaissé", "Date d'encaissement", "Reste dû (cumul)", "Jours de retard", "Observations (relance, commandement, régularisation...)"]
BLOCK = 12 + 5  # en-tête lot, en-tête colonnes, 12 mois, total, ligne vide
blocks = []
for j in range(NLOTS):
    col = get_column_letter(2 + j)  # colonne du lot dans la feuille Lots
    top = 4 + j * BLOCK
    su.cell(top, 1, f"=\"Lot \"&{j + 1}&\" : \"&Lots!{col}{R_NOM}").font = blue
    for i, c in enumerate(cols, 1):
        cell = su.cell(top + 1, i, c); cell.font = Font(bold=True); cell.fill = head; cell.border = box; cell.alignment = Alignment(wrap_text=True)
    su.row_dimensions[top + 1].height = 30
    first = top + 2
    for m in range(12):
        r = first + m
        su.cell(r, 1, f"=EDATE(Lots!$B$6,{m})").number_format = 'MMM YYYY'
        su.cell(r, 2, f"=DATE(YEAR(A{r}),MONTH(A{r}),Lots!$B$5)").number_format = DATE
        su.cell(r, 3, f"=Lots!{col}${R_LOYER}").number_format = EUR
        su.cell(r, 4, f"=Lots!{col}${R_PROV}").number_format = EUR
        su.cell(r, 5, f"=C{r}+D{r}").number_format = EUR
        su.cell(r, 6, None).fill = inp; su.cell(r, 6).number_format = EUR
        su.cell(r, 7, None).fill = inp; su.cell(r, 7).number_format = DATE
        prev = f"H{r - 1}" if m > 0 else "0"
        su.cell(r, 8, f"={prev}+E{r}-F{r}").number_format = EUR
        su.cell(r, 9, f"=IF(F{r}>=E{r},IF(G{r}=\"\",0,MAX(0,G{r}-B{r})),IF(TODAY()>B{r},TODAY()-B{r},0))")
        su.cell(r, 10, None).fill = inp
        for c in (1, 2, 3, 4, 5, 8, 9): su.cell(r, c).fill = res
        for c in range(1, 11): su.cell(r, c).border = box
    last = first + 11
    t = last + 1
    su.cell(t, 1, "Total de l'année").font = Font(bold=True)
    for c, L in ((3, "C"), (4, "D"), (5, "E"), (6, "F")):
        su.cell(t, c, f"=SUM({L}{first}:{L}{last})").number_format = EUR
    su.cell(t, 8, f"=H{last}").number_format = EUR
    su.cell(t, 9, f"=COUNTIF(I{first}:I{last},\">0\")")
    su.cell(t, 10, "Reste dû = impayés cumulés en fin d'année ; colonne I en total = nombre de mois en retard").font = Font(italic=True, color="666666")
    for c in range(1, 11):
        su.cell(t, c).border = box; su.cell(t, c).fill = res; su.cell(t, c).font = Font(bold=True)
    blocks.append((first, last, t))
for i, w in enumerate([12, 12, 12, 12, 12, 12, 14, 14, 11, 50], 1): su.column_dimensions[get_column_letter(i)].width = w
su.freeze_panes = "A4"

# ---------------- Feuille 3 : synthèse annuelle ----------------
sy = wb.create_sheet("Synthèse annuelle")
sy["A1"] = "Synthèse annuelle par lot et total : régularisation des charges, taxe foncière, cash avant impôt. Toutes les cellules sont calculées."; sy["A1"].font = blue
sy["A2"] = "Le cash avant impôt ne tient pas compte de l'impôt (revenus fonciers ou BIC) ni des amortissements : voir modules 7 et 9."; sy["A2"].font = Font(italic=True, color="666666")
sy.cell(4, 1, "Poste").font = Font(bold=True)
for j in range(NLOTS):
    col = get_column_letter(2 + j)
    c = sy.cell(4, 2 + j, f"=Lots!{col}{R_NOM}"); c.font = Font(bold=True); c.fill = head; c.border = box
c = sy.cell(4, 2 + NLOTS, "Total"); c.font = Font(bold=True); c.fill = head; c.border = box
sy.cell(4, 3 + NLOTS, "Note").font = Font(bold=True)

lines = []
def add(label, fml, note="", fmt=EUR, bold=False):
    lines.append((label, fml, note, fmt, bold))
# fml est une fonction j -> formule pour le lot j (col = lettre du lot dans Lots, blk = (first,last,total) dans Suivi)
add("Loyers et charges dus sur l'année", lambda j, col, b: f"='Suivi mensuel'!E{b[2]}", "12 × (loyer + provision)")
add("Loyers et charges encaissés", lambda j, col, b: f"='Suivi mensuel'!F{b[2]}", "Saisi mois par mois dans la feuille Suivi")
add("Impayés en fin d'année", lambda j, col, b: f"='Suivi mensuel'!H{b[2]}", "Reste dû cumulé")
add("Taux d'encaissement", lambda j, col, b: f"=IF(B{{r1}}=0,0,B{{r2}}/B{{r1}})", "Objectif : 100 %", fmt='0.0%')
add("Provisions de charges encaissées (part charges)", lambda j, col, b: f"=IF('Suivi mensuel'!E{b[2]}=0,0,'Suivi mensuel'!F{b[2]}*'Suivi mensuel'!D{b[2]}/'Suivi mensuel'!E{b[2]})", "Part des encaissements affectée aux charges, au prorata")
add("Charges récupérables réelles", lambda j, col, b: f"=Lots!{col}{R_CHREC}", "Décret 87-713 ; inclut la TEOM si vous la refacturez")
add("Régularisation des charges (+ : à réclamer au locataire, - : à rembourser)", lambda j, col, b: f"=IF(Lots!{col}{R_PROV}=0,0,B{{r6}}-B{{r5}})", "Pas de régularisation en forfait (meublé) : mettre les charges réelles égales au forfait, ou ignorer cette ligne", bold=True)
add("Taxe foncière (TEOM comprise)", lambda j, col, b: f"=Lots!{col}{R_TF}", "")
add("dont TEOM récupérée sur le locataire", lambda j, col, b: f"=Lots!{col}{R_TEOM}", "Comptée dans les charges récupérables réelles")
add("Charges de copropriété non récupérables", lambda j, col, b: f"=Lots!{col}{R_CHNR}", "")
add("Assurance PNO", lambda j, col, b: f"=Lots!{col}{R_PNO}", "")
add("Assurance loyers impayés", lambda j, col, b: f"=Lots!{col}{R_GLI}*B{{r1}}", "Taux × loyers CC dus")
add("Honoraires de gestion", lambda j, col, b: f"=Lots!{col}{R_GEST}*B{{r2}}", "Taux × encaissements")
add("Entretien et réparations", lambda j, col, b: f"=Lots!{col}{R_ENT}", "")
add("Mensualités de crédit (12 mois)", lambda j, col, b: f"=Lots!{col}{R_CREDIT}*12", "")
add("Cash avant impôt", lambda j, col, b: f"=B{{r2}}+B{{r7}}-B{{r6}}-B{{r8}}+B{{r9}}-B{{r10}}-B{{r11}}-B{{r12}}-B{{r13}}-B{{r14}}-B{{r15}}", "Encaissé + régularisation − charges récupérables − TF + TEOM − copro − PNO − GLI − gestion − entretien − crédit", bold=True)
add("Cash mensuel moyen", lambda j, col, b: f"=B{{r16}}/12", "", bold=True)
add("Rendement des encaissements (cash / loyers dus)", lambda j, col, b: f"=IF(B{{r1}}=0,0,B{{r16}}/B{{r1}})", "", fmt='0.0%')

R1 = 5
rows = {f"r{i + 1}": R1 + i for i in range(len(lines))}
for i, (label, fml, note, fmt, bold) in enumerate(lines):
    r = R1 + i
    sy.cell(r, 1, label).border = box
    if bold: sy.cell(r, 1).font = Font(bold=True)
    for j in range(NLOTS):
        col = get_column_letter(2 + j); L = get_column_letter(2 + j)
        f = fml(j, col, blocks[j])
        # remplacer les références B{rN} par la colonne du lot
        for k, v in rows.items():
            f = f.replace("B{" + k + "}", f"{L}{v}")
        c = sy.cell(r, 2 + j, f); c.fill = res; c.border = box; c.number_format = fmt
        if bold: c.font = Font(bold=True)
    tcol = 2 + NLOTS; TL = get_column_letter(tcol); first_lot = get_column_letter(2); last_lot = get_column_letter(1 + NLOTS)
    if fmt == EUR:
        tf = f"=SUM({first_lot}{r}:{last_lot}{r})"
    elif label.startswith("Taux d'encaissement"):
        tf = f"=IF({TL}{rows['r1']}=0,0,{TL}{rows['r2']}/{TL}{rows['r1']})"
    else:
        tf = f"=IF({TL}{rows['r1']}=0,0,{TL}{rows['r16']}/{TL}{rows['r1']})"
    c = sy.cell(r, tcol, tf); c.fill = res; c.border = box; c.number_format = fmt; c.font = Font(bold=True)
    if note: sy.cell(r, tcol + 1, note).font = Font(italic=True, color="666666")
sy.column_dimensions["A"].width = 62
for j in range(NLOTS + 1): sy.column_dimensions[get_column_letter(2 + j)].width = 14
sy.column_dimensions[get_column_letter(3 + NLOTS)].width = 75

out = ROOT / "08-mettre-en-location" / "outils" / "suivi_loyers.xlsx"; wb.save(out); print("->", out.relative_to(ROOT))
