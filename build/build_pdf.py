#!/usr/bin/env python3
"""Génère les PDF d'un module : cours, mémo, quiz, exercice (mode « cours ») et les outils de `outils/*.md`
(mode « support » : check-lists, guides, modèles, fiches de référence, trames).

Usage : python build/build_pdf.py 00-preparer [autre-module ...]   (sans argument : tous les modules)
Sortie : pdf/<module>-<document>.pdf

Conventions Markdown des outils (voir BRIEF) :
  ligne 1 `# Titre`, ligne 2 `Type : Checklist|Guide|Modèle|Référence|Trame`, ligne 3 phrase d'usage,
  bloc `>` dont la première ligne en gras est le titre de l'encadré, sections `##` numérotées automatiquement,
  ligne en italique sous un `##` = sous-titre, items `- [ ] texte` = cases à cocher, `**Objet :** ...` = barre d'objet,
  `### Entrée` = fiche de référence, ligne `Locataire type : ... · ...` = ligne de métadonnées.
"""
import sys, re, pathlib, datetime
import fitz, markdown

ROOT = pathlib.Path(__file__).resolve().parent.parent
BRAND = "IMMO · PLAYBOOK"
VERSION = "v1.0 · " + datetime.date.today().strftime("%m/%Y")

GREEN = "#2f4b2f"; DARK = "#17301f"; LIME = "#c8e05a"; CREAM = "#f6f7ec"; GREY = "#6a7a6a"

CSS_COMMON = f"""
body {{ font-family: sans-serif; font-size: 10pt; line-height: 1.45; color: #1e2a22; }}
p {{ margin: 0 0 6pt 0; }}
a {{ color: {GREEN}; }}
table {{ border-collapse: collapse; width: 100%; margin: 6pt 0 10pt 0; font-size: 9pt; }}
th, td {{ border-width: 0.5pt; border-style: solid; border-color: #b9c4b9; padding: 3pt 5pt; vertical-align: top; text-align: left; }}
th {{ background-color: #e8eedf; }}
li {{ margin-bottom: 3pt; }}
hr {{ border: 0; border-top: 0.5pt solid #b9c4b9; margin: 10pt 0; }}
code {{ font-family: monospace; font-size: 9pt; }}
.type {{ font-size: 7.5pt; color: #556; margin: 0 0 1pt 0; }}
h1 {{ font-size: 19pt; margin: 0 0 5pt 0; color: {DARK}; font-weight: normal; }}
.purpose {{ color: #333; margin: 0 0 8pt 0; }}
.box {{ background-color: {CREAM}; border-left: 3px solid {LIME}; padding: 6pt 10pt 2pt 10pt; margin: 6pt 0 12pt 0; }}
.boxh {{ font-size: 7.5pt; color: #4a6b35; margin: 0 0 4pt 0; }}
blockquote {{ background-color: {CREAM}; border-left: 3px solid {LIME}; padding: 6pt 10pt 2pt 10pt; margin: 6pt 0 10pt 0; color: #3b4a3b; font-size: 9.5pt; }}
"""

CSS_COURS = CSS_COMMON + f"""
h2 {{ font-size: 13pt; color: {GREEN}; font-weight: bold; margin: 16pt 0 4pt 0; border-bottom: 1.2pt solid {LIME}; padding-bottom: 2pt; }}
h3 {{ font-size: 11pt; color: {GREEN}; margin: 11pt 0 3pt 0; }}
"""

CSS_SUPPORT = CSS_COMMON + f"""
h2 {{ font-size: 11pt; color: {GREEN}; font-weight: normal; margin: 14pt 0 2pt 0; }}
h3 {{ font-size: 9.5pt; color: {GREEN}; font-weight: normal; margin: 10pt 0 2pt 0; }}
.sub {{ font-style: italic; color: #666; font-size: 9pt; margin: 0 0 6pt 0; }}
.chk {{ list-style: none; margin: 0 0 4pt 0; padding-left: 0; }}
.chk li {{ margin: 0 0 3pt 0; }}
.objet {{ background-color: #dcea6a; padding: 4pt 8pt; margin: 8pt 0 6pt 0; }}
.lbl {{ font-size: 7.5pt; color: {GREEN}; }}
.meta {{ font-size: 8.5pt; color: {GREY}; margin: 0 0 8pt 0; }}
"""

def spaced(text: str, html: bool = True) -> str:
    """Capitales espacées, à la manière d'un intertitre de support : I M M O.  En HTML, les espaces entre mots
    sont insécables pour ne pas être fusionnés avec l'espacement des lettres."""
    words = text.upper().split()
    gap = "&nbsp;&nbsp;&nbsp;" if html else "   "
    return gap.join(" ".join(w) for w in words)

def md_html(text: str) -> str:
    return markdown.markdown(text, extensions=["tables", "sane_lists"])

def support_html(md_text: str) -> str:
    """Transforme un outil Markdown en HTML stylé (type, usage, encadré, sections numérotées, cases, objets)."""
    lines = md_text.strip("\n").splitlines()
    title = lines[0].lstrip("# ").strip() if lines and lines[0].startswith("#") else "Document"
    body = lines[1:]
    doc_type = "Support"
    if body and re.match(r"^Type\s*:", body[0]):
        doc_type = body[0].split(":", 1)[1].strip(); body = body[1:]
    while body and not body[0].strip(): body = body[1:]
    purpose = ""
    if body and not body[0].startswith(("#", ">", "-", "|", "*")):
        purpose = body[0].strip(); body = body[1:]
    html = md_html("\n".join(body))
    # encadré : première citation
    def box(m):
        inner = m.group(1)
        inner = re.sub(r"^\s*<p><strong>(.*?)</strong>\s*:?\s*(.*?)</p>", lambda k: f'<p class="boxh">{spaced(k.group(1))}</p>' + (f"<p>{k.group(2)}</p>" if k.group(2).strip() else ""), inner, count=1, flags=re.S)
        return f'<div class="box">{inner}</div>'
    html = re.sub(r"<blockquote>(.*?)</blockquote>", box, html, flags=re.S)
    # sections numérotées
    counter = [0]
    def h2(m):
        counter[0] += 1
        return f'<h2>{counter[0]:02d}&nbsp;&nbsp;{spaced(re.sub("<.*?>", "", m.group(1)))}</h2>'
    html = re.sub(r"<h2>(.*?)</h2>", h2, html, flags=re.S)
    html = re.sub(r"<h3>(.*?)</h3>", lambda m: f"<h3>{re.sub('<.*?>', '', m.group(1)).upper()}</h3>", html, flags=re.S)
    # sous-titre italique juste après un titre
    html = re.sub(r"(</h[23]>)\s*<p><em>(.*?)</em></p>", r'\1<p class="sub">\2</p>', html, flags=re.S)
    # cases à cocher
    html = re.sub(r"<li>\s*\[ \]\s*(.*?)</li>", r'<li class="chk">&#9744;&nbsp; \1</li>', html, flags=re.S)
    html = re.sub(r"<ul>(\s*<li class=\"chk\">)", r'<ul class="chk">\1', html)
    # barre objet
    html = re.sub(r"<p><strong>Objet\s*:?</strong>\s*:?\s*(.*?)</p>", r'<div class="objet"><span class="lbl">O B J E T</span>&nbsp;&nbsp;&nbsp; \1</div>', html, flags=re.S)
    # métadonnées
    html = re.sub(r"<p>((?:Locataire type|Point de vigilance|Profil|Risque|Délai|Source)\s*:.*?)</p>", r'<p class="meta">\1</p>', html, flags=re.S)
    head = f'<p class="type">{spaced(doc_type)}</p><h1>{title}</h1>' + (f'<p class="purpose">{purpose}</p>' if purpose else "")
    return head + html

def render(html: str, css: str, pdf_path: pathlib.Path, footer_left: str):
    story = fitz.Story(html=html, user_css=css)
    writer = fitz.DocumentWriter(str(pdf_path))
    mediabox = fitz.paper_rect("a4")
    where = mediabox + (50, 62, -50, -62)
    more = 1
    while more:
        dev = writer.begin_page(mediabox)
        more, _ = story.place(where)
        story.draw(dev)
        writer.end_page()
    writer.close()
    doc = fitz.open(str(pdf_path)); n = doc.page_count
    grey = (0.45, 0.5, 0.45); line = (0.72, 0.78, 0.72)
    for i, page in enumerate(doc, 1):
        w, h = page.rect.width, page.rect.height
        brand = spaced(BRAND, html=False)
        page.insert_text((w - 50 - fitz.get_text_length(brand, fontsize=7), 40), brand, fontsize=7, color=grey)
        page.draw_line((50, h - 44), (w - 50, h - 44), color=line, width=0.5)
        page.insert_text((50, h - 32), f"[ {footer_left} ]", fontsize=7, color=grey)
        right = f"{VERSION} · p. {i}/{n}"
        page.insert_text((w - 50 - fitz.get_text_length(right, fontsize=7), h - 32), right, fontsize=7, color=grey)
    doc.save(str(pdf_path), incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP); doc.close()
    return n

def module_title(module_dir: pathlib.Path) -> str:
    readme = module_dir / "README.md"
    if readme.exists():
        first = readme.read_text(encoding="utf-8").splitlines()[0]
        return first.lstrip("# ").strip()
    return module_dir.name

def build(module_dir: pathlib.Path):
    out = ROOT / "pdf"; out.mkdir(exist_ok=True)
    title = module_title(module_dir)
    for name in ["cours", "memo", "quiz", "exercice"]:
        src = module_dir / f"{name}.md"
        if src.exists():
            dest = out / f"{module_dir.name}-{name}.pdf"
            n = render(md_html(src.read_text(encoding="utf-8")), CSS_COURS, dest, f"{title} · {name}")
            print(f"-> {dest.relative_to(ROOT)} ({n} p.)")
    for src in sorted((module_dir / "outils").glob("*.md")):
        dest = out / f"{module_dir.name}-outil-{src.stem}.pdf"
        n = render(support_html(src.read_text(encoding="utf-8")), CSS_SUPPORT, dest, title)
        print(f"-> {dest.relative_to(ROOT)} ({n} p.)")

if __name__ == "__main__":
    mods = [ROOT / m for m in sys.argv[1:]] or sorted(p for p in ROOT.glob("[01][0-9]-*") if p.is_dir())
    for m in mods: build(m)
