#!/usr/bin/env python3
"""Génère 05-financer/outils/dossier_bancaire.pptx : modèle de dossier de financement en dix diapositives,
à remplir puis à exporter en PDF et à placer en tête des pièces. Les champs entre crochets sont à remplacer ;
les consignes sont dans les notes de chaque diapositive (elles ne s'impriment pas).

Usage : python build/dossier_bancaire.py
"""
import pathlib
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = ROOT / "05-financer" / "outils" / "dossier_bancaire.pptx"

GREEN = RGBColor(0x2F, 0x4B, 0x2F); DARK = RGBColor(0x17, 0x30, 0x1F); LIME = RGBColor(0xC8, 0xE0, 0x5A)
CREAM = RGBColor(0xF6, 0xF7, 0xEC); GREY = RGBColor(0x6A, 0x7A, 0x6A); WHITE = RGBColor(0xFF, 0xFF, 0xFF)
HEAD = RGBColor(0xE8, 0xEE, 0xDF); LINE = RGBColor(0xB9, 0xC4, 0xB9); TEXT = RGBColor(0x1E, 0x2A, 0x22)
FONT = "Calibri"
W, H = Inches(13.333), Inches(7.5)
MX = Inches(0.7)                      # marge latérale
CW = W - 2 * MX                       # largeur utile

prs = Presentation()
prs.slide_width, prs.slide_height = W, H
BLANK = prs.slide_layouts[6]
state = {"n": 0}


def text(slide, x, y, w, h, value, size=14, bold=False, color=TEXT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False):
    """Zone de texte ; `value` peut être une liste de paragraphes."""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame; tf.word_wrap = True; tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05); tf.margin_top = tf.margin_bottom = Inches(0.03)
    for i, line in enumerate(value if isinstance(value, list) else [value]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align; p.space_after = Pt(4)
        r = p.add_run(); r.text = line
        r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic; r.font.color.rgb = color; r.font.name = FONT
    return box


def rect(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None: s.line.fill.background()
    else: s.line.color.rgb = line; s.line.width = Pt(0.75)
    s.shadow.inherit = False
    return s


def new_slide(title, subtitle, notes):
    """Diapositive de contenu : bandeau de titre, sous-titre, numéro de page, consignes en notes."""
    state["n"] += 1
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, W, Inches(0.12), LIME)
    text(s, MX, Inches(0.35), CW - Inches(1), Inches(0.6), title, size=26, color=DARK)
    text(s, MX, Inches(0.95), CW, Inches(0.4), subtitle, size=12, color=GREY, italic=True)
    rect(s, MX, Inches(1.4), CW, Pt(1.2), LIME)
    text(s, MX, H - Inches(0.45), Inches(8), Inches(0.3), "Dossier de financement · [Nom de l'emprunteur ou de la société] · [mois année]", size=9, color=GREY)
    text(s, W - MX - Inches(1), H - Inches(0.45), Inches(1), Inches(0.3), str(state["n"]), size=9, color=GREY, align=PP_ALIGN.RIGHT)
    s.notes_slide.notes_text_frame.text = notes
    return s


def table(slide, x, y, w, rows, widths, header=True, size=11, row_h=0.36, first_col_bold=True):
    """Tableau sobre : en-tête vert clair, première colonne en gras, filets fins."""
    gt = slide.shapes.add_table(len(rows), len(rows[0]), x, y, w, Inches(row_h * len(rows)))
    t = gt.table
    tbl = gt._element.graphic.graphicData.tbl
    tbl.tblPr.set("bandRow", "0"); tbl.tblPr.set("firstRow", "0")
    total = sum(widths)
    for i, cw in enumerate(widths): t.columns[i].width = int(w * cw / total)
    for r, row in enumerate(rows):
        t.rows[r].height = Inches(row_h)
        for c, val in enumerate(row):
            cell = t.cell(r, c); cell.fill.solid()
            is_head = header and r == 0
            cell.fill.fore_color.rgb = HEAD if is_head else (CREAM if c == 0 and first_col_bold else WHITE)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Inches(0.08); cell.margin_top = cell.margin_bottom = Inches(0.03)
            tf = cell.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; run = p.add_run(); run.text = val
            run.font.size = Pt(size); run.font.name = FONT; run.font.color.rgb = DARK if is_head else TEXT
            run.font.bold = is_head or (c == 0 and first_col_bold) or val.startswith("Total") or val.startswith("Coût total")
    return t


# ---------------------------------------------------------------- 1. Page de garde
s = prs.slides.add_slide(BLANK); state["n"] += 1
rect(s, 0, 0, W, H, DARK)
rect(s, MX, Inches(2.2), Inches(1.2), Pt(4), LIME)
text(s, MX, Inches(2.4), CW, Inches(1), "DOSSIER DE FINANCEMENT", size=40, color=WHITE)
text(s, MX, Inches(3.4), CW, Inches(0.6), "Acquisition d'un [type de bien] à [ville], destiné à la location [nue, meublée, commerciale]", size=18, color=LIME)
text(s, MX, Inches(4.7), CW, Inches(1.2), ["[Prénom Nom] ou [Nom de la SCI], représentée par [Prénom Nom], gérant", "[téléphone] · [adresse électronique]", "[date de remise du dossier]"], size=14, color=WHITE)
s.notes_slide.notes_text_frame.text = ("Page de garde. Une ligne dit ce que vous achetez et pourquoi. Pas de logo de banque, pas de photo de vous. "
    "Exportez le fichier en PDF une fois rempli et placez-le en tête des pièces, dans le même fichier.")

# ---------------------------------------------------------------- 2. Synthèse
s = new_slide("L'opération en une page", "Ce que le comité de crédit doit retenir s'il ne lit que cette diapositive.",
    "Le comité lit la première page de contenu, puis cherche les chiffres. Donnez-lui la décision à prendre : combien, pour quoi, "
    "avec quel loyer en face. Les quatre chiffres doivent être identiques à ceux des diapositives suivantes et des pièces.")
kw = (CW - Inches(0.6)) / 4
for i, (lab, val, sub) in enumerate([
    ("COÛT TOTAL", "[xxx xxx] €", "prix, frais, travaux"),
    ("MONTANT DEMANDÉ", "[xxx xxx] €", "sur [20] ans, apport [xx xxx] €"),
    ("LOYER ANNUEL", "[xx xxx] €", "hors charges, [justifié par...]"),
    ("TAUX D'EFFORT APRÈS OPÉRATION", "[xx] %", "loyers retenus à 70 %"),
]):
    x = MX + i * (kw + Inches(0.2))
    rect(s, x, Inches(1.75), kw, Inches(1.75), CREAM, LINE)
    rect(s, x, Inches(1.75), Inches(0.07), Inches(1.75), LIME)
    text(s, x + Inches(0.2), Inches(1.85), kw - Inches(0.3), Inches(0.35), lab, size=10, color=GREEN, bold=True)
    text(s, x + Inches(0.2), Inches(2.25), kw - Inches(0.3), Inches(0.7), val, size=28, color=DARK)
    text(s, x + Inches(0.2), Inches(2.95), kw - Inches(0.3), Inches(0.4), sub, size=11, color=GREY)
text(s, MX, Inches(3.85), CW, Inches(0.4), "Pourquoi ce projet tient", size=15, color=GREEN, bold=True)
text(s, MX, Inches(4.3), CW, Inches(2.4), [
    "•  Le bien : [type, surface, adresse], acheté [x] % sous le prix des ventes récentes du quartier ([source : ventes DVF jointes]).",
    "•  Le loyer : [montant] € par mois, [bail en cours depuis ... / estimation d'agence et annonces comparables jointes].",
    "•  L'emprunteur : [situation professionnelle, ancienneté], épargne de [montant] € conservée après l'opération.",
    "•  Le rendement : [x] % brut sur le coût total ; le loyer couvre [x] % de la mensualité.",
    "•  Le calendrier : compromis signé le [date], condition suspensive de prêt jusqu'au [date], acte prévu le [date].",
], size=13)

# ---------------------------------------------------------------- 3. Le projet
s = new_slide("Présentation du projet", "Le bien, son prix, son loyer. Des faits vérifiables dans les pièces, pas d'adjectifs.",
    "Chaque ligne renvoie à une pièce : compromis, devis, estimation locative, bail. Rendement brut = loyer annuel hors charges ÷ coût total. "
    "Pour un local commercial, ajoutez le type de bail (3/6/9), l'échéance, l'indexation (ILC ou ILAT) et l'activité du preneur. "
    "Ajoutez une ou deux photos à droite si la place le permet.")
table(s, MX, Inches(1.65), Inches(7.4), [
    ["Poste", "Valeur"],
    ["Adresse", "[adresse complète]"],
    ["Type de bien, surface", "[appartement T3, immeuble de x lots, local], [xx] m²"],
    ["État, année, DPE", "[à rénover / bon état], [année], DPE [lettre]"],
    ["Prix affiché", "[xxx xxx] €"],
    ["Prix négocié", "[xxx xxx] €"],
    ["Frais de notaire", "[xx xxx] €"],
    ["Travaux (devis joints)", "[xx xxx] €"],
    ["Frais de garantie et de dossier", "[x xxx] €"],
    ["Coût total de l'opération", "[xxx xxx] €"],
    ["Loyer hors charges", "[xxx] € par mois, soit [xx xxx] € par an"],
    ["Mode de location", "[nue / meublée / colocation / bail commercial 3/6/9]"],
    ["Rendement brut sur coût total", "[x,x] %"],
], [3, 4.4], size=11, row_h=0.37)
rect(s, MX + Inches(7.7), Inches(1.65), CW - Inches(7.7), Inches(2.6), CREAM, LINE)
text(s, MX + Inches(7.7), Inches(1.65), CW - Inches(7.7), Inches(2.6), "[Photo de la façade ou plan]", size=12, color=GREY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(s, MX + Inches(7.7), Inches(4.45), CW - Inches(7.7), Inches(0.35), "Le marché, en trois lignes", size=13, color=GREEN, bold=True)
text(s, MX + Inches(7.7), Inches(4.85), CW - Inches(7.7), Inches(1.8), [
    "•  Ventes comparables : [x xxx] €/m² ([source]).",
    "•  Loyers comparables : [xx] €/m² ([source]).",
    "•  Demande locative : [employeurs, transports, délai de relocation].",
], size=12)

# ---------------------------------------------------------------- 4. La demande de prêt
s = new_slide("Demande de prêt", "Une demande précise obtient une réponse précise.",
    "Écrivez ce que vous demandez, pas « à définir ». La durée longue préserve la capacité d'emprunt suivante. Le différé partiel se cale "
    "sur la durée des travaux et de la mise en location. La garantie réelle sur le bien financé (hypothèque légale spéciale du prêteur de "
    "deniers) cloisonne chaque opération. L'exonération d'indemnités de remboursement anticipé se demande ici, pas après l'offre.")
table(s, MX, Inches(1.75), CW, [
    ["Élément", "Demande", "Commentaire"],
    ["Montant emprunté", "[xxx xxx] €", "[prix + frais + travaux − apport]"],
    ["Apport", "[xx xxx] €", "[origine : épargne visible sur les relevés joints]"],
    ["Durée", "[20 / 25] ans", "[durée maximale souhaitée]"],
    ["Taux", "Fixe, à proposer", ""],
    ["Différé", "[partiel de x mois / aucun]", "[durée des travaux et de la mise en location]"],
    ["Garantie proposée", "[hypothèque sur le bien financé]", "[et caution personnelle des associés si société]"],
    ["Assurance emprunteur", "[quotités par tête]", "[contrat groupe à la signature, délégation possible ensuite]"],
    ["Clauses demandées", "[exonération d'IRA, modulation, report]", ""],
], [2.6, 3.6, 5.7], size=12, row_h=0.47)
text(s, MX, Inches(6.2), CW, Inches(0.5), "Plan de financement : [coût total xxx xxx €] = [apport xx xxx €] + [prêt xxx xxx €].", size=13, color=GREEN, bold=True)

# ---------------------------------------------------------------- 5. Les emprunteurs
s = new_slide("Les emprunteurs ou les associés", "La banque prête à des personnes, même quand c'est une société qui signe.",
    "Une colonne par emprunteur ou par associé caution. Profession : indiquez le type de contrat et l'ancienneté. "
    "Un dirigeant salarié de sa société indique son contrat de travail et son salaire, pas le chiffre d'affaires de la société. "
    "Expérience immobilière : nombre de biens, années de gestion, aucun impayé de crédit.")
table(s, MX, Inches(1.75), CW, [
    ["", "Emprunteur ou associé 1", "Emprunteur ou associé 2"],
    ["Nom, prénom", "[NOM Prénom]", "[NOM Prénom]"],
    ["Date de naissance", "[jj/mm/aaaa]", "[jj/mm/aaaa]"],
    ["Situation familiale", "[célibataire, marié sous le régime de..., pacsé], [x] enfant(s)", "[...]"],
    ["Profession", "[métier], [CDI] depuis [année] chez [employeur]", "[...]"],
    ["Revenu net mensuel", "[x xxx] €", "[x xxx] €"],
    ["Logement", "[propriétaire de sa résidence principale / locataire]", "[...]"],
    ["Adresse", "[adresse]", "[adresse]"],
    ["Expérience immobilière", "[x biens loués depuis aaaa, gestion directe]", "[...]"],
], [2.4, 4.75, 4.75], size=12, row_h=0.47)

# ---------------------------------------------------------------- 6. Le patrimoine
s = new_slide("Le patrimoine", "Ce que vous possédez, ce que vous devez, ce qui reste. Un tableau par emprunteur ou par associé.",
    "Valeur moins dette égale net. Séparez l'épargne disponible (livrets) de l'épargne longue (PEA, assurance-vie, épargne salariale) : "
    "la banque aime voir une épargne qui ne sera pas dépensée demain. Chaque ligne a son justificatif dans les pièces. "
    "Dupliquez la diapositive pour le second emprunteur.")
table(s, MX, Inches(1.75), CW, [
    ["[NOM Prénom]", "Valeur", "Capital restant dû", "Net"],
    ["Résidence principale", "[xxx xxx] €", "[xxx xxx] €", "[xxx xxx] €"],
    ["Biens locatifs détenus en direct", "[xxx xxx] €", "[xxx xxx] €", "[xxx xxx] €"],
    ["Parts de sociétés immobilières", "[xxx xxx] €", "[xxx xxx] €", "[xxx xxx] €"],
    ["Livrets (épargne disponible)", "[xx xxx] €", "", "[xx xxx] €"],
    ["PEA, compte-titres", "[xx xxx] €", "", "[xx xxx] €"],
    ["Assurance-vie, épargne salariale", "[xx xxx] €", "", "[xx xxx] €"],
    ["Total", "[xxx xxx] €", "[xxx xxx] €", "[xxx xxx] €"],
], [4.6, 2.4, 2.4, 2.4], size=12, row_h=0.47)
text(s, MX, Inches(5.75), CW, Inches(0.8), ["Épargne mensuelle régulière : [xxx] € par mois depuis [année].",
     "Épargne conservée après l'opération : [xx xxx] €, soit [x] mois de mensualités."], size=13, color=GREEN)

# ---------------------------------------------------------------- 7. La structure
s = new_slide("La structure (si l'achat se fait en société)", "Qui détient quoi, et qui s'engage. À supprimer pour un achat en nom propre.",
    "Un schéma simple : les associés, la société, le bien. La banque aime la simplicité : une société civile détenue directement par des "
    "personnes physiques se lit en dix secondes. Une holding intercalée ne passe que si elle a des bilans et de la trésorerie à montrer. "
    "Indiquez le régime fiscal, le gérant et qui se porte caution.")
bw, bh = Inches(3.2), Inches(0.9)
cx = W / 2
def node(x, y, label, sub, fill=CREAM, col=TEXT):
    r = rect(s, x, y, bw, bh, fill, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, x, y + Inches(0.08), bw, Inches(0.4), label, size=14, bold=True, color=col, align=PP_ALIGN.CENTER)
    text(s, x, y + Inches(0.45), bw, Inches(0.4), sub, size=11, color=GREY if fill != DARK else LIME, align=PP_ALIGN.CENTER)
    return r
def link(x1, y1, x2, y2):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2); c.line.color.rgb = GREEN; c.line.width = Pt(1.5)
xa, xb = cx - bw - Inches(0.9), cx + Inches(0.9)
node(xa, Inches(1.8), "[Associé 1]", "[50] % des parts · caution")
node(xb, Inches(1.8), "[Associé 2]", "[50] % des parts · caution")
node(cx - bw / 2, Inches(3.5), "[Nom de la SCI]", "SCI à l'[IS / IR] · gérant : [Prénom Nom]", DARK, WHITE)
node(cx - bw / 2, Inches(5.2), "[Le bien financé]", "[adresse] · hypothèque au profit de la banque")
link(xa + bw / 2, Inches(1.8) + bh, cx - Inches(0.6), Inches(3.5))
link(xb + bw / 2, Inches(1.8) + bh, cx + Inches(0.6), Inches(3.5))
link(cx, Inches(3.5) + bh, cx, Inches(5.2))
text(s, MX, Inches(6.3), CW, Inches(0.5), "Société immatriculée le [date] (ou en cours d'immatriculation). Comptes ouverts à [banque]. Statuts et Kbis joints.", size=12, color=GREY, align=PP_ALIGN.CENTER)

# ---------------------------------------------------------------- 8. Revenus et charges
s = new_slide("Revenus et charges", "Le calcul du banquier, fait avant lui : loyers retenus à 70 %, charges à 100 %.",
    "Refaites le calcul comme la banque : salaires à 100 %, loyers existants et futur à 70 %, toutes les mensualités à 100 %, assurance comprise. "
    "Taux d'effort = charges ÷ revenus retenus (plafond de 35 %, règle du HCSF en vigueur depuis 2022, à vérifier). "
    "Le prêt de la résidence principale se compte à 50 % par tête s'il est commun. Un tableau par emprunteur, ou un tableau du foyer.")
table(s, MX, Inches(1.75), CW, [
    ["Revenus mensuels", "À 100 %", "Retenu", "Charges mensuelles", "À 100 %"],
    ["Salaire net ou revenu d'activité", "[x xxx] €", "[x xxx] €", "Prêt de la résidence principale ou loyer", "[xxx] €"],
    ["Loyers actuels", "[x xxx] €", "[70 % : xxx] €", "Prêts locatifs en cours", "[xxx] €"],
    ["Loyer futur du projet", "[xxx] €", "[70 % : xxx] €", "Mensualité du projet, assurance comprise", "[xxx] €"],
    ["Autres revenus réguliers", "[xxx] €", "[xxx] €", "Autres crédits, pensions versées", "[xxx] €"],
    ["Total", "[x xxx] €", "[x xxx] €", "Total", "[x xxx] €"],
], [3.3, 1.5, 1.7, 3.9, 1.5], size=12, row_h=0.5)
kw = (CW - Inches(0.4)) / 3
for i, (lab, val) in enumerate([("TAUX D'EFFORT AVANT", "[xx] %"), ("TAUX D'EFFORT APRÈS", "[xx] %"), ("RESTE À VIVRE APRÈS", "[x xxx] €")]):
    x = MX + i * (kw + Inches(0.2))
    rect(s, x, Inches(5.05), kw, Inches(1.2), CREAM, LINE); rect(s, x, Inches(5.05), Inches(0.07), Inches(1.2), LIME)
    text(s, x + Inches(0.2), Inches(5.12), kw - Inches(0.3), Inches(0.3), lab, size=10, color=GREEN, bold=True)
    text(s, x + Inches(0.2), Inches(5.45), kw - Inches(0.3), Inches(0.7), val, size=26, color=DARK)

# ---------------------------------------------------------------- 9. Ce que la banque y gagne
s = new_slide("Ce que je propose à la banque", "Un prêt se rembourse ; une relation rapporte. Dites ce que vous apportez.",
    "Un conseiller a des objectifs commerciaux sur l'année : assurances, épargne, placements, comptes professionnels. Ne proposez que ce que "
    "vous êtes prêt à souscrire et à garder. Le prêt seul rapporte peu à la banque ; l'ensemble de la relation rapporte davantage, et c'est "
    "ce total que le comité regarde. Chiffrez quand vous le pouvez.")
table(s, MX, Inches(1.75), CW, [
    ["Ce que j'apporte", "Détail", "Quand"],
    ["Domiciliation", "[salaires / loyers de l'opération / compte de la société]", "[à la signature]"],
    ["Assurance emprunteur", "[contrat de la banque, quotités]", "[à l'offre]"],
    ["Assurances du bien", "[propriétaire non occupant, multirisque de l'immeuble]", "[à l'acte]"],
    ["Épargne et placements", "[versement de x € sur un PEA ou une assurance-vie, x € par mois]", "[à l'acte, puis mensuel]"],
    ["Comptes et moyens de paiement", "[compte de la société, carte, banque en ligne]", "[à la création]"],
    ["Prévoyance, autres contrats", "[habitation, automobile, prévoyance]", "[à l'échéance des contrats actuels]"],
], [3, 6.4, 2.5], size=12, row_h=0.5)
text(s, MX, Inches(5.5), CW, Inches(1), ["Perspective : [un autre projet dans les x mois, à étudier avec vous si cette opération se déroule bien].",
     "Sur [20] ans, cette opération représente pour la banque environ [xx xxx] € d'intérêts et [x xxx] € d'assurances et de services."], size=13, color=GREEN)

# ---------------------------------------------------------------- 10. Les pièces
s = new_slide("Pièces jointes", "Dans le même fichier, à la suite, dans cet ordre. Un dossier complet au premier envoi gagne trois semaines.",
    "Listez ce qui est réellement joint, et seulement cela. Toutes les pages de chaque document, lisibles, de moins de trois mois quand une durée "
    "est demandée. La check-list détaillée du module donne la liste complète par pochette.")
col = (CW - Inches(0.4)) / 2
for i, (titre, items) in enumerate([
    ("Documents personnels (pour chaque emprunteur et chaque caution)", [
        "Pièce d'identité, justificatif de domicile, livret de famille",
        "Deux derniers avis d'imposition",
        "Trois derniers bulletins de salaire et celui de décembre",
        "Contrat de travail ou attestation d'employeur",
        "Indépendants : trois derniers bilans ; micro : déclarations URSSAF",
        "Trois derniers relevés de tous les comptes",
        "Relevés d'épargne (livrets, PEA, assurance-vie)",
        "Taxes foncières, tableaux d'amortissement des prêts en cours",
        "Baux et dernières quittances des biens loués",
    ]),
    ("Documents du projet", [
        "Compromis de vente et ses annexes (diagnostics)",
        "Estimation de la valeur locative ou annonces comparables",
        "Bail en cours et dernières quittances si le bien est loué",
        "Devis des travaux",
        "Trois derniers procès-verbaux d'assemblée générale si copropriété",
        "Photos, plan, ventes comparables du quartier",
        "Société : statuts, Kbis, bénéficiaires effectifs",
        "Société : procès-verbal autorisant l'achat et l'emprunt",
        "Société : bilans et relevés si elle existe déjà",
    ]),
]):
    x = MX + i * (col + Inches(0.4))
    rect(s, x, Inches(1.75), col, Inches(4.7), CREAM, LINE)
    text(s, x + Inches(0.2), Inches(1.85), col - Inches(0.4), Inches(0.4), titre, size=13, color=GREEN, bold=True)
    text(s, x + Inches(0.2), Inches(2.55), col - Inches(0.4), Inches(3.8), ["☐  " + it for it in items], size=12)

prs.core_properties.author = prs.core_properties.last_modified_by = "Immo Playbook"; prs.core_properties.title = "Dossier de financement"
OUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUT)
print("écrit :", OUT.relative_to(ROOT), "·", state["n"], "diapositives")
