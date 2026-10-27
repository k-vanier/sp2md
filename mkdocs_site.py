"""Projet MkDocs (thème Material) prêt à publier, construit à partir de la copie brute.

Résultat dans <copie>/mkdocs/ :
    mkdocs.yml   configuration (réécrite à chaque génération)
    docs/        pages Markdown + fichiers d'origine (site autonome)
    site/        site HTML statique, à déposer tel quel sur un serveur web
"""
import logging
import posixpath
import re
import shutil
from pathlib import Path

import yaml

from commun import ecrire_texte, existe, log, lp
from convertir import Convertisseur

LANGUES = {1036: "fr", 1033: "en", 1031: "de", 1034: "es", 3082: "es", 1040: "it", 1043: "nl", 2070: "pt", 1046: "pt"}
# Bibliothèques techniques de SharePoint : converties, mais absentes du menu.
BIBLIOTHEQUES_TECHNIQUES = {"siteassets", "style library", "formservertemplates", "sitecollectionimages", "sitecollectiondocuments"}

LISEZMOI = """# Site MkDocs généré par sp2md

Copie du site SharePoint {url}, faite le {date}.

- `site/` : le site HTML, prêt à publier. Ouvrez `site/index.html` pour le consulter sur ce poste.
- `docs/` : les pages Markdown et les fichiers d'origine.
- `mkdocs.yml` : la configuration. Elle est réécrite à chaque génération par sp2md.

## Publier

Attention : le contenu vient d'un SharePoint d'entreprise. Vérifiez qu'il peut être rendu public
avant de le publier ailleurs que sur un serveur interne.

1. **Serveur web (interne ou public)** : copier le contenu de `site/` à la racine du site
   (IIS, Apache, nginx, partage web…). Aucune base de données ni langage serveur nécessaire.
2. **GitHub Pages** : déposer ce dossier dans un dépôt GitHub, puis, avec Python installé :
   `pip install -r requirements.txt` puis `mkdocs gh-deploy`.
3. **Netlify** : glisser-déposer le dossier `site/` sur https://app.netlify.com/drop.

## Modifier puis reconstruire

Après avoir modifié des pages dans `docs/` : `pip install -r requirements.txt` puis `mkdocs build`
(ou `mkdocs serve` pour un aperçu sur http://127.0.0.1:8000).
"""


def generer_mkdocs(dossier: Path, documents=True, construire=True) -> Path:
    projet = Path(dossier) / "mkdocs"
    docs = projet / "docs"
    # docs/ est entièrement généré : on repart de zéro pour ne pas publier de pages supprimées entre-temps.
    shutil.rmtree(lp(docs), ignore_errors=True)
    conv = Convertisseur(dossier, documents=documents, sortie=docs, mkdocs=True)
    conv.lancer()

    webs = {w[0]: w for w in conv.webs}
    racine = webs[""][1]
    langue = LANGUES.get(racine.get("Language"), "fr")
    date = (conv.site.get("Date") or "")[:10]
    config = {
        "site_name": racine["Title"],
        "site_description": racine.get("Description") or f"Copie du site SharePoint {racine['Url']}",
        "copyright": f"Contenu issu de {racine['Url']}" + (f", copié le {date}" if date else ""),
        "docs_dir": "docs",
        "site_dir": "site",
        # Pages en .html : le site fonctionne aussi ouvert en local, et « X.docx.md » ne gêne pas « X.docx ».
        "use_directory_urls": False,
        "theme": {
            "name": "material",
            "language": langue,
            "font": False,  # pas d'appel à Google Fonts : le site marche sur un réseau interne
            "features": ["navigation.tabs", "navigation.sections", "navigation.top", "navigation.footer",
                         "search.suggest", "search.highlight", "toc.follow", "content.code.copy"],
            "palette": [
                {"media": "(prefers-color-scheme: light)", "scheme": "default", "primary": "indigo", "accent": "indigo",
                 "toggle": {"icon": "material/brightness-7", "name": "Mode sombre"}},
                {"media": "(prefers-color-scheme: dark)", "scheme": "slate", "primary": "indigo", "accent": "indigo",
                 "toggle": {"icon": "material/brightness-4", "name": "Mode clair"}},
            ],
        },
        "plugins": [{"search": {"lang": langue}}],
        "markdown_extensions": ["tables", {"toc": {"permalink": True, "permalink_title": "Lien vers cette section"}}, "admonition", "attr_list", "md_in_html",
                                "footnotes", "def_list", "pymdownx.superfences", "pymdownx.tilde", "pymdownx.details"],
        # Documents convertis et pages de bibliothèque : accessibles par les liens, pas tous dans le menu.
        "not_in_nav": "*.md",
        "nav": nav_web(conv, webs, "", docs),
    }
    entete = "# Généré par sp2md : ce fichier est réécrit à chaque génération.\n"
    ecrire_texte(projet / "mkdocs.yml", entete + yaml.safe_dump(config, allow_unicode=True, sort_keys=False, width=1000))
    ecrire_texte(projet / "requirements.txt", "mkdocs-material\n")
    ecrire_texte(projet / ".gitignore", "site/\n")
    ecrire_texte(projet / "LISEZMOI.md", LISEZMOI.format(url=racine["Url"], date=date or "?"))
    log(f"Projet MkDocs : {projet / 'mkdocs.yml'}")
    if construire:
        construire_site(projet)
    return projet


def nav_web(conv: Convertisseur, webs, web_rel, docs: Path):
    """Menu d'un site : navigation SharePoint, autres pages, documents, listes, sous-sites, plan du site."""
    _, web, navigation, listes = webs[web_rel]
    index = posixpath.join(web_rel, "index.md")
    sru = web["ServerRelativeUrl"].rstrip("/") + "/"
    present = lambda c: c and existe(docs / c)  # noqa: E731
    utilises = {index}

    def noeud(n):
        """Lien de la navigation SharePoint -> entrée du menu MkDocs (page locale ou adresse externe)."""
        titre = re.sub(r"\s+", " ", n.get("Title") or "").strip()
        url = n.get("Url") or ""
        cible = None
        if url and "linkless.header" not in url:
            local = conv.chemin_local(url, sru)
            if present(local):
                cible = local
            else:
                _, _, absolu = conv.analyser(url, sru)
                cible = absolu if absolu.startswith("http") else None
        enfants = [e for e in (noeud(c) for c in n.get("Children") or []) if e]
        if cible:
            utilises.add(cible)
        if enfants:
            if cible and cible != index:
                enfants.insert(0, {titre: cible})
            return {titre or "Section": enfants}
        if not cible or cible == index or not titre:
            return None
        return {titre: cible}

    barre = navigation.get("QuickLaunch") or navigation.get("TopNavigationBar") or []
    menu_sp = [e for e in (noeud(n) for n in barre) if e]

    pages, bibliotheques, autres = [], [], []
    for liste, donnees in listes:
        racine = conv.rel(liste["RootFolder"]["ServerRelativeUrl"]) or ""
        if conv.est_pages(liste, donnees):
            for el in conv.fichiers(donnees):
                r = conv.rel(el["FileRef"])
                cible = conv.cibles.get(r.lower()) if r else None
                if present(cible) and cible not in utilises:
                    pages.append({el.get("Title") or posixpath.splitext(el["FileLeafRef"])[0]: cible})
        elif liste.get("ItemCount") and posixpath.basename(racine).lower() not in BIBLIOTHEQUES_TECHNIQUES:
            cible = conv.cibles.get(racine.lower())
            if present(cible) and cible not in utilises:
                (bibliotheques if liste["BaseType"] == 1 else autres).append({liste["Title"]: cible})

    menu = [{"Accueil": index}] + menu_sp
    titre_pages = "Autres pages" if menu_sp else "Pages"
    for titre, entrees in ((titre_pages, pages), ("Documents", bibliotheques), ("Listes", autres)):
        if entrees:
            menu.append({titre: sorted(entrees, key=lambda e: next(iter(e)).lower())})
    sous = [s for s in web.get("SousSites") or [] if s is not None and s in webs]
    if sous:
        menu.append({"Sous-sites": [{webs[s][1]["Title"]: nav_web(conv, webs, s, docs)} for s in sous]})
    plan = conv.plans.get(web_rel)
    if plan and plan != index:
        menu.append({"Plan du site": plan})
    return menu


class _VersJournal(logging.Handler):
    def __init__(self):
        super().__init__(logging.INFO)
        self.avertissements = 0

    def emit(self, record):
        if record.levelno >= logging.WARNING:
            self.avertissements += 1
            log(f"  mkdocs : {record.getMessage()}")


def construire_site(projet: Path):
    from mkdocs.commands.build import build
    from mkdocs.config import load_config

    log("Construction du site HTML (MkDocs)...")
    journal = logging.getLogger("mkdocs")
    relais = _VersJournal()
    journal.addHandler(relais)
    journal.propagate = False
    try:
        config = load_config(config_file=str(projet / "mkdocs.yml"))
        config.plugins.on_startup(command="build", dirty=False)
        try:
            build(config)
        finally:
            config.plugins.on_shutdown()
    finally:
        journal.removeHandler(relais)
        journal.propagate = True
    log(f"Site prêt à publier : {projet / 'site'} ({relais.avertissements} avertissement(s) MkDocs)")
