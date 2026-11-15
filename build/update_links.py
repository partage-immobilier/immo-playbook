#!/usr/bin/env python3
"""Met à jour les liens de téléchargement des PDF et Excel : tableau « Téléchargements » dans le README racine
et section « Fichiers à télécharger » dans le README de chaque module. À lancer après build_pdf.py."""
import pathlib, re
ROOT = pathlib.Path(__file__).resolve().parent.parent
PDF = ROOT / "pdf"
LABELS = {"cours": "Cours", "memo": "Mémo", "quiz": "Quiz", "exercice": "Exercice corrigé"}

def module_files(mod: pathlib.Path):
    docs = {k: PDF / f"{mod.name}-{k}.pdf" for k in LABELS}
    docs = {k: v for k, v in docs.items() if v.exists()}
    outils = sorted(PDF.glob(f"{mod.name}-outil-*.pdf"))
    xlsx = sorted(p for p in (mod / "outils").glob("*") if p.suffix in (".xlsx", ".pptx")) if (mod / "outils").exists() else []
    return docs, outils, xlsx

def pretty(stem: str) -> str:
    return stem.replace("_", " ").capitalize()

def replace_section(text: str, header: str, body: str) -> str:
    pat = re.compile(rf"\n## {re.escape(header)}\n.*?(?=\n## |\Z)", re.S)
    block = f"\n## {header}\n\n{body.strip()}\n"
    return pat.sub(block, text, count=1) if pat.search(text) else text.rstrip("\n") + "\n" + block

mods = sorted(p for p in ROOT.glob("[01][0-9]-*") if p.is_dir())
rows = ["| Module | Cours | Mémo | Quiz | Exercice | Outils (PDF) | Excel et PowerPoint |", "|---|---|---|---|---|---|---|"]
for mod in mods:
    docs, outils, xlsx = module_files(mod)
    title = (mod / "README.md").read_text(encoding="utf-8").splitlines()[0].lstrip("# ").strip()
    cell = lambda k: f"[PDF](pdf/{docs[k].name})" if k in docs else "à venir"
    o = ", ".join(f"[{pretty(p.stem.split('-outil-',1)[1])}](pdf/{p.name})" for p in outils) or ("sans objet" if docs else "à venir")
    x = ", ".join(f"[{pretty(p.stem)}]({mod.name}/outils/{p.name})" for p in xlsx) or ("à venir" if not docs else "sans objet")
    rows.append(f"| [{title}]({mod.name}/) | {cell('cours')} | {cell('memo')} | {cell('quiz')} | {cell('exercice')} | {o} | {x} |")
    # section du module
    lines = []
    for k, p in docs.items(): lines.append(f"- [{LABELS[k]} (PDF)](../pdf/{p.name})")
    for p in outils: lines.append(f"- [{pretty(p.stem.split('-outil-',1)[1])} (PDF)](../pdf/{p.name})")
    for p in xlsx: lines.append(f"- [{pretty(p.stem)} ({'PowerPoint' if p.suffix == '.pptx' else 'Excel'})](outils/{p.name})")
    if lines:
        r = mod / "README.md"; r.write_text(replace_section(r.read_text(encoding="utf-8"), "Fichiers à télécharger", "\n".join(lines)), encoding="utf-8")
readme = ROOT / "README.md"
body = "Cliquez sur un PDF puis sur le bouton « Download » (ou « Raw ») de GitHub pour l'enregistrer. Les Excel et le PowerPoint se téléchargent de la même façon.\n\n" + "\n".join(rows)
readme.write_text(replace_section(readme.read_text(encoding="utf-8"), "Téléchargements", body), encoding="utf-8")
print("liens mis à jour :", len(mods), "modules")
