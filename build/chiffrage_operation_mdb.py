#!/usr/bin/env python3
"""Génère 11-marchand-de-biens/outils/chiffrage_operation_mdb.xlsx : chiffrage d'une opération de marchand de biens
(coût de revient détaillé, financement, marge brute, IS, marge nette, TRI sur la durée de portage) avec une colonne
« scénario dégradé » (+15 % travaux, +6 mois, -5 % prix). Toutes les sorties sont des formules Excel."""
import pathlib
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side

ROOT = pathlib.Path(__file__).resolve().parent.parent
wb = Workbook(); ws = wb.active; ws.title = "Chiffrage"
blue = Font(color="1F3A5F", bold=True); inp = PatternFill("solid", fgColor="FFF4D6"); res = PatternFill("solid", fgColor="E4EBF3")
thin = Side(style="thin", color="B8C4D0"); box = Border(left=thin, right=thin, top=thin, bottom=thin); NOTE = Font(italic=True, color="666666")

# Colonnes : A libellé, B scénario de base, C scénario dégradé, D note
def row(r, label, base=None, degr=None, fmt='#,##0 €', fill_b=None, fill_c=None, note=""):
    ws.cell(r, 1, label).border = box
    b = ws.cell(r, 2, base); b.border = box; b.number_format = fmt
    if fill_b: b.fill = fill_b
    c = ws.cell(r, 3, degr); c.border = box; c.number_format = fmt
    if fill_c: c.fill = fill_c
    if note: ws.cell(r, 4, note).font = NOTE

ws["A1"] = "Chiffrage d'une opération de marchand de biens (à jour au 15/11/2026 ; taux et pratiques bancaires à vérifier)"; ws["A1"].font = blue
ws["A2"] = "Cellules jaunes : à remplir. Cellules bleues : calculées. Colonne C : scénario dégradé (travaux +15 %, délai +6 mois, prix de revente -5 %), paramètres modifiables en lignes 5 à 7."; ws["A2"].font = Font(italic=True)
for j, h in enumerate(["Poste", "Scénario de base", "Scénario dégradé", "Notes"], 1):
    c = ws.cell(3, j, h); c.font = blue; c.fill = res; c.border = box

ws["A4"] = "0. Paramètres du scénario dégradé"; ws["A4"].font = blue
row(5, "Majoration des travaux", 0, 0.15, '0%', inp, inp, "Pratique : +15 % sur un bâti ancien sans diagnostic structure")
row(6, "Mois de portage supplémentaires", 0, 6, '0', inp, inp)
row(7, "Baisse du prix de revente", 0, 0.05, '0%', inp, inp)

ws["A9"] = "1. Acquisition"; ws["A9"].font = blue
row(10, "Prix d'achat négocié", 150000, "=B10", fill_b=inp, fill_c=res)
row(11, "Taux de frais d'acquisition", 0.028, "=B11", '0.0%', inp, res, "2,5 à 3 % avec l'engagement de revendre sous 5 ans (art. 1115 CGI : 0,715 % au lieu de 5,8 %) ; 7,5 à 8 % sans")
row(12, "Frais d'acquisition", "=B10*B11", "=C10*C11", fill_b=res, fill_c=res)

ws["A14"] = "2. Travaux et honoraires"; ws["A14"].font = blue
row(15, "Travaux TTC (devis signés)", 60000, "=B15*(1+C5)", fill_b=inp, fill_c=res, note="TVA au bon taux : 10 % amélioration, 5,5 % énergie, 20 % sinon ; non récupérable si revente exonérée")
row(16, "Taux d'imprévus sur les travaux", 0.10, "=B16", '0%', inp, res, "10 % minimum, 15 % sur bâti ancien")
row(17, "Imprévus", "=B15*B16", "=C15*C16", fill_b=res, fill_c=res)
row(18, "Honoraires (architecte, géomètre, diagnostics, bureau d'études, notaire division)", 6000, "=B18", fill_b=inp, fill_c=res)
row(19, "Assurances (dommages-ouvrage, RC professionnelle, multirisque)", 3500, "=B19", fill_b=inp, fill_c=res, note="DO : 2 à 5 % des travaux, obligatoire si travaux relevant de la décennale (art. L242-1 C. assurances)")

ws["A21"] = "3. Financement et portage"; ws["A21"].font = blue
row(22, "Durée totale de portage (mois, de l'acte à l'encaissement)", 10, "=B22+C6", '0', inp, res, "Règle de prudence courante : 12 mois maximum")
row(23, "Base finançable (prix + frais + travaux + imprévus + honoraires + assurances)", "=B10+B12+B15+B17+B18+B19", "=C10+C12+C15+C17+C18+C19", fill_b=res, fill_c=res)
row(24, "Part financée par la banque", 0.75, "=B24", '0%', inp, res, "70 à 80 % courant ; apport 20 à 30 %, davantage pour une première opération")
row(25, "Montant du crédit d'opération", "=B23*B24", "=C23*C24", fill_b=res, fill_c=res, note="Le dégradé suppose que le surcoût est financé dans les mêmes proportions ; en pratique il est souvent sur fonds propres")
row(26, "Apport (fonds propres immobilisés)", "=B23-B25", "=C23-C25", fill_b=res, fill_c=res)
row(27, "Taux d'intérêt annuel du crédit", 0.06, "=B27", '0.00%', inp, res, "4,5 à 7 % en 2026 selon dossier (à vérifier)")
row(28, "Intérêts sur la durée de portage", "=B25*B27*B22/12", "=C25*C27*C22/12", fill_b=res, fill_c=res, note="In fine ou capitalisés : intérêts simples sur le capital débloqué")
row(29, "Frais de dossier, hypothèque et mainlevée", 5000, "=B29", fill_b=inp, fill_c=res, note="Dossier 0,5 à 1 % ; hypothèque 1 à 1,5 % ; mainlevée quelques centaines d'euros")
row(30, "Taxe foncière annuelle", 1200, "=B30", fill_b=inp, fill_c=res)
row(31, "Charges de portage mensuelles (copropriété, énergie, assurance prolongée)", 100, "=B31", fill_b=inp, fill_c=res)
row(32, "Taxe foncière au prorata et charges de portage", "=B30*B22/12+B31*B22", "=C30*C22/12+C31*C22", fill_b=res, fill_c=res)

ws["A34"] = "4. Revente"; ws["A34"].font = blue
row(35, "Surface vendue (m²)", 100, "=B35", '0', inp, res)
row(36, "Prix de revente au m² (premier quartile des ventes DVF comparables)", 3000, "=B36*(1-C7)", fill_b=inp, fill_c=res, note="Jamais la médiane des annonces")
row(37, "Prix de revente", "=B35*B36", "=C35*C36", fill_b=res, fill_c=res)
row(38, "Taux de commission de revente", 0.04, "=B38", '0.0%', inp, res, "3 à 5 % ; 0 si vente directe")
row(39, "Commission de revente", "=B37*B38", "=C37*C38", fill_b=res, fill_c=res)
row(40, "Vente soumise à la TVA sur le prix total ? (immeuble neuf ou rendu neuf)", "non", "=B40", '@', inp, res, "oui / non. Art. 257, I, 2, 2° CGI : si oui, TVA 20 % collectée sur le prix et TVA sur travaux récupérable")
row(41, "TVA récupérable sur les travaux si vente soumise à TVA (estimation)", 10000, "=B41*(1+C5)", fill_b=inp, fill_c=res, note="TVA incluse dans la ligne 15 ; ne compte que si la ligne 40 est « oui »")
row(42, "Prix de revente hors taxes", '=IF(B40="oui",B37/1.2,B37)', '=IF(C40="oui",C37/1.2,C37)', fill_b=res, fill_c=res)

ws["A44"] = "5. Résultat"; ws["A44"].font = blue
row(45, "Coût de revient complet", '=B10+B12+B15+B17+B18+B19+B28+B29+B32+B39-IF(B40="oui",B41,0)', '=C10+C12+C15+C17+C18+C19+C28+C29+C32+C39-IF(C40="oui",C41,0)', fill_b=res, fill_c=res)
row(46, "Marge brute", "=B42-B45", "=C42-C45", fill_b=res, fill_c=res)
row(47, "Marge brute en % du coût de revient", "=IF(B45>0,B46/B45,0)", "=IF(C45>0,C46/C45,0)", '0.0%', res, res, "Banques : souvent 15 à 20 % minimum (pratique)")
row(48, "Seuil d'IS à taux réduit", 42500, "=B48", fill_b=inp, fill_c=res, note="15 % jusqu'à 42 500 € de bénéfice, 25 % au-delà (2026), sous conditions")
row(49, "IS", "=IF(B46<=0,0,MIN(B46,B48)*0.15+MAX(0,B46-B48)*0.25)", "=IF(C46<=0,0,MIN(C46,C48)*0.15+MAX(0,C46-C48)*0.25)", fill_b=res, fill_c=res, note="Calcul sur l'opération seule ; l'IS réel se calcule sur le résultat de l'exercice")
row(50, "Marge nette après IS", "=B46-B49", "=C46-C49", fill_b=res, fill_c=res)
row(51, "Marge nette en % du coût de revient", "=IF(B45>0,B50/B45,0)", "=IF(C45>0,C50/C45,0)", '0.0%', res, res, "Règle de prudence courante : 20 % minimum (pratique, à adapter)")
row(52, "Rendement des fonds propres sur la période", "=IF(B26>0,B50/B26,0)", "=IF(C26>0,C50/C26,0)", '0.0%', res, res)
row(53, "TRI annualisé sur fonds propres", "=IF(AND(B26>0,B50>-B26),(1+B50/B26)^(12/B22)-1,-1)", "=IF(AND(C26>0,C50>-C26),(1+C50/C26)^(12/C22)-1,-1)", '0.0%', res, res, "Fonds propres immobilisés pendant la durée de portage, marge nette récupérée à la revente")
row(54, "Règle de prudence respectée (marge nette ≥ 20 % et délai ≤ 12 mois) ?", '=IF(AND(B51>=0.2,B22<=12),"oui","non")', '=IF(AND(C51>=0.2,C22<=12),"oui","non")', '@', res, res)
row(55, "Coût de chaque mois de portage supplémentaire", "=B25*B27/12+B30/12+B31", "=C25*C27/12+C30/12+C31", fill_b=res, fill_c=res, note="Intérêts + taxe foncière + charges : ce que coûte un mois de retard")

ws["A57"] = "6. Prix d'achat maximal pour la règle de prudence (scénario de base)"; ws["A57"].font = blue
row(58, "Taux d'IS retenu pour l'approximation", 0.15, None, '0%', inp, None, "15 % si la marge brute reste sous 42 500 € ; 25 % sinon")
row(59, "Coût de revient maximal pour 20 % de marge nette", "=B42/(1+0.2/(1-B58))", None, fill_b=res, note="Résout (prix HT - coût) × (1 - IS) = 20 % × coût")
row(60, "Coûts hors prix et hors postes proportionnels au prix", "=B15+B17+B18+B19+B29+B32+B39-IF(B40=\"oui\",B41,0)", None, fill_b=res)
row(61, "Coefficient des postes proportionnels au prix (frais d'acquisition + intérêts sur la part financée)", "=1+B11+(1+B11)*B24*B27*B22/12", None, '0.0000', res, None, "Chaque euro de prix coûte 1 € + frais + intérêts sur la part empruntée")
row(62, "Prix d'achat maximal (indicatif)", "=MAX(0,(B59-B60-(B15+B17+B18+B19)*B24*B27*B22/12)/B61)", None, fill_b=res, note="Au-delà, la marge nette passe sous 20 % du coût de revient avec ces hypothèses")

ws.column_dimensions["A"].width = 78; ws.column_dimensions["B"].width = 20; ws.column_dimensions["C"].width = 20; ws.column_dimensions["D"].width = 80

out = ROOT / "11-marchand-de-biens" / "outils" / "chiffrage_operation_mdb.xlsx"; wb.save(out); print("->", out.relative_to(ROOT))
