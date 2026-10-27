"""Construit la version Markdown à partir de la copie brute (dossier brut/)."""
import hashlib
import json
import os
import posixpath
import re
import shutil
from html import unescape
from pathlib import Path
from urllib.parse import quote, unquote, urlparse

from bs4 import BeautifulSoup
from markdownify import MarkdownConverter

from commun import DOSSIER_META, ecrire_texte, existe, lire_json, log, lp
from telecharger import MODELE_PAGES

CONVERTIBLES = {".docx", ".pptx", ".xlsx", ".xlsm", ".xls", ".pdf", ".html", ".htm", ".csv", ".txt", ".json", ".xml",
                ".msg", ".epub"}
# Fichiers texte affichés tels quels dans un bloc de code.
TEXTES_BRUTS = {".pem", ".pub", ".log", ".ini", ".cfg", ".conf", ".sh", ".ps1", ".bat", ".sql", ".yml", ".yaml",
                ".properties"}
IMAGES = {".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".bmp", ".ico"}

WP_MARKDOWN = "1ef5ed11-ce7b-44be-bc5e-4abd55101d16"
WP_SEPARATEUR = "2161a1c6-db61-4731-b97c-3cdb303f7cbb"
WP_ESPACEUR = "8654b779-4886-46d4-8ffb-b5ed960ee986"
WP_TITRE = "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788"      # zone de titre / bannière de page
WP_CODE = "7b317bca-c919-4982-af2f-8399173e5a1e"
WP_BOUTON = "0f087d7f-520e-42b7-89c0-496aaf979d58"
WP_FICHIER = "b7dd04e1-19ce-4b24-9132-b60a1c2b910d"   # « Fichier et média »

# Propriétés d'une page que SharePoint recopie dans son fichier .aspx.
CHAMPS_ASPX = {"CanvasContent1", "WikiField", "PublishingPageContent", "LayoutWebpartsContent", "BannerImageUrl",
               "Description", "Title"}
CHAMPS_CONTENU = ("CanvasContent1", "WikiField", "PublishingPageContent")
# MkDocs ne sait pas écrire au-delà de 260 caractères sous Windows : chemins plus longs raccourcis.
LIMITE_CHEMIN = 80

# Colonnes de vue qui n'apportent rien dans un tableau.
COLONNES_IGNOREES = {"DocIcon", "Edit", "SelectTitle", "_UIVersionString", "AppAuthor", "AppEditor"}
ALIAS_COLONNES = {"LinkTitle": "Title", "LinkTitleNoMenu": "Title", "LinkFilename": "FileLeafRef",
                  "LinkFilenameNoMenu": "FileLeafRef"}


def nettoyer(md: str) -> str:
    md = re.sub(r"[ \t]+\n", "\n", md)
    return re.sub(r"\n{3,}", "\n\n", md).strip() + "\n"


def echapper_cellule(v) -> str:
    return str(v).replace("|", "\\|").replace("\r", "").replace("\n", "<br>")


class ConvertisseurHtml(MarkdownConverter):
    """Listes imbriquées indentées de 4 espaces : lues de la même façon par CommonMark et par MkDocs."""

    def convert_li(self, el, text, parent_tags):
        md = super().convert_li(el, text, parent_tags)
        puce = md.split(" ", 1)[0] + " "
        largeur = max(4, len(puce))
        lignes = md.split("\n")
        lignes[0] = puce.ljust(largeur) + lignes[0][len(puce):]
        decalage = " " * len(puce)
        lignes[1:] = [" " * largeur + l[len(puce):] if l.startswith(decalage) else l for l in lignes[1:]]
        return "\n".join(lignes)


class Convertisseur:
    def __init__(self, dossier_site: Path, documents=True, sortie=None, mkdocs=False):
        """mkdocs=True : site autonome (originaux copiés à côté, page d'accueil SharePoint en index.md)."""
        self.dossier = Path(dossier_site)
        self.brut = self.dossier / "brut"
        self.md = Path(sortie) if sortie else self.dossier / "markdown"
        self.mkdocs = mkdocs
        self.plans = {}    # web_rel -> page « plan du site » de ce site
        self.meta = self.brut / DOSSIER_META
        site = lire_json(self.meta / "site.json")
        self.site = site
        self.racine_sru = site["ServerRelativeUrl"].rstrip("/")
        u = urlparse(site["Url"])
        self.origine = f"{u.scheme}://{u.netloc}"
        self.documents = documents
        self.cibles = {}   # chemin relatif (minuscule) -> fichier .md ou image dans le dossier de sortie
        self.webs = []     # (web_rel, web, navigation, [(liste, données)])
        self.markitdown = None
        self.stats = {"pages": 0, "documents": 0, "images": 0, "listes": 0, "echecs": 0}

    # ---------- chemins et liens ----------

    def rel(self, sru: str):
        sru = unquote(sru).rstrip("/")
        r = self.racine_sru
        if not r:
            return sru.lstrip("/")
        if sru.lower() == r.lower():
            return ""
        if sru.lower().startswith(r.lower() + "/"):
            return sru[len(r) + 1:]
        return None

    def cible(self, rel: str):
        c = self.cibles.get(rel.lower())
        if c is None and rel.lower().endswith(".aspx"):
            # AllItems.aspx, Forms/…, page par défaut : on remonte vers la liste ou le site.
            parent = rel
            while c is None and parent:
                parent = posixpath.dirname(parent)
                c = self.cibles.get(parent.lower())
        return c

    def relatif(self, depuis_md: str, vers: str) -> str:
        """Lien relatif entre deux chemins exprimés depuis le dossier du site."""
        r = posixpath.relpath(vers, posixpath.dirname(depuis_md) or ".")
        return quote(r, safe="/._-~")

    def analyser(self, href: str, source_sru: str):
        """Lien SharePoint -> (rel, ancre, url absolue) ; rel vaut None pour un lien hors du site."""
        u = urlparse(unescape(href))
        if u.scheme and u.netloc.lower() != urlparse(self.origine).netloc.lower():
            return None, "", href
        # Lien de partage « /:u:/r/sites/… » : même fichier que « /sites/… ».
        chemin = re.sub(r"^/:[a-z]{1,2}:/[a-z]/", "/", u.path)
        if not u.scheme and not chemin.startswith("/"):
            chemin = posixpath.normpath(posixpath.join(posixpath.dirname(source_sru), unquote(chemin)))
        absolu = self.origine + quote(unquote(chemin), safe="/") + (f"?{u.query}" if u.query else "")
        return self.rel(chemin), (f"#{u.fragment}" if u.fragment else ""), absolu

    def chemin_local(self, href: str, source_sru: str):
        """Fichier du dossier de sortie visé par un lien SharePoint (page, document, liste), ou None."""
        rel, _, _ = self.analyser(href, source_sru)
        return self.cible(rel) if rel is not None else None

    def lien(self, href: str, depuis_md: str, source_sru: str) -> str:
        if not href or href.startswith(("#", "mailto:", "tel:", "javascript:", "data:")):
            return href
        rel, ancre, absolu = self.analyser(href, source_sru)
        if rel is None:
            return absolu
        c = self.cible(rel)
        if c is not None:
            return self.relatif(depuis_md, c) + ancre
        if rel and existe(self.chemin_brut(rel)):
            return self.vers_brut(depuis_md, rel)
        return absolu

    def chemin_brut(self, rel: str) -> Path:
        """Fichier de brut/ ; retrouve aussi un nom abîmé par un transfert (zip…) : « é » devenu « Ã© »."""
        p = self.brut / rel
        if not existe(p):
            for codage in ("cp1252", "latin-1"):
                try:
                    autre = self.brut / rel.encode("utf-8").decode(codage)
                except UnicodeDecodeError:
                    continue
                if existe(autre):
                    return autre
        return p

    def local(self, rel: str) -> str:
        """Chemin d'un fichier dans le dossier de sortie (raccourci en MkDocs s'il est trop long)."""
        if not self.mkdocs or len(rel) <= LIMITE_CHEMIN:
            return rel
        empreinte = hashlib.sha1(rel.lower().encode("utf-8")).hexdigest()[:10]
        # Pas de « _ » en tête : GitHub Pages (Jekyll) et d'autres hébergeurs ignorent ces dossiers.
        return f"sp2md-fichiers/{empreinte}/{posixpath.basename(rel)[-60:]}"

    def vers_brut(self, depuis_md: str, rel: str) -> str:
        """Lien vers un fichier d'origine : dans brut/, ou copié à côté en mode MkDocs (site autonome)."""
        src = self.chemin_brut(rel)
        if not self.mkdocs:
            return self.relatif(depuis_md, Path(os.path.relpath(src, self.md)).as_posix())
        rel = self.local(rel)
        dest = self.md / rel
        if not existe(dest) or os.path.getmtime(lp(dest)) < os.path.getmtime(lp(src)):
            os.makedirs(lp(dest.parent), exist_ok=True)
            shutil.copy2(lp(src), lp(dest))
        return self.relatif(depuis_md, rel)

    def html_vers_md(self, html: str, depuis_md: str, source_sru: str) -> str:
        soup = BeautifulSoup(html or "", "html.parser")
        # Espaces insécables : en début de ligne, ils transformeraient le paragraphe en bloc de code.
        for texte in soup.find_all(string=lambda s: "\xa0" in s):
            if not texte.find_parent(["pre", "code"]):
                texte.replace_with(texte.replace("\xa0", " "))
        # Image insérée dans le texte : un bloc « imagePlugin » qui ne contient pas de balise <img>.
        for bloc in soup.find_all(attrs={"data-imageurl": True}):
            image = soup.new_tag("img", src=bloc["data-imageurl"], alt=bloc.get("data-alttext") or "")
            paragraphe = soup.new_tag("p")
            if bloc.get("data-linkurl"):
                lien = soup.new_tag("a", href=bloc["data-linkurl"])
                lien.append(image)
                paragraphe.append(lien)
            else:
                paragraphe.append(image)
            if bloc.get("data-captiontext"):
                legende = soup.new_tag("em")
                legende.string = bloc["data-captiontext"]
                paragraphe.extend([soup.new_tag("br"), legende])
            bloc.replace_with(paragraphe)
        # Tableau d'une seule case (encadré de mise en valeur) : rendu en citation.
        for table in soup.find_all("table"):
            cases = table.find_all(["td", "th"])
            if len(cases) == 1:
                citation = soup.new_tag("blockquote")
                citation.extend(list(cases[0].contents))
                table.replace_with(citation)
        for a in soup.find_all("a", href=True):
            a["href"] = self.lien(a["href"], depuis_md, source_sru)
        for img in soup.find_all("img", src=True):
            img["src"] = self.lien(img["src"].split("?")[0], depuis_md, source_sru)
        return ConvertisseurHtml(heading_style="ATX", bullets="-").convert(str(soup)).strip()

    # ---------- passe 1 : inventaire ----------

    def charger_web(self, web_rel: str):
        meta = self.meta / web_rel
        if not existe(meta / "web.json"):
            log.erreur(f"métadonnées absentes pour le sous-site « {web_rel} »")
            return
        web = lire_json(meta / "web.json")
        navigation = lire_json(meta / "navigation.json") if existe(meta / "navigation.json") else {}
        listes = []
        for liste in lire_json(meta / "lists.json"):
            f = meta / "listes" / f"{liste['Id']}.json"
            if existe(f):
                listes.append((liste, lire_json(f)))
        for liste, donnees in listes:
            if self.est_pages(liste, donnees):
                for el in self.fichiers(donnees):
                    self.completer_page(el)
        self.webs.append((web_rel, web, navigation, listes))
        index = posixpath.join(web_rel, "index.md")
        self.cibles[web_rel.lower()] = index
        # En MkDocs, la page d'accueil SharePoint devient index.md et le plan du site passe à côté.
        accueil = self.page_accueil(web_rel, web, listes) if self.mkdocs else None
        self.plans[web_rel] = posixpath.join(web_rel, "plan-du-site.md") if accueil else index

        for liste, donnees in listes:
            racine = self.rel(liste["RootFolder"]["ServerRelativeUrl"]) or ""
            if self.est_pages(liste, donnees):
                self.cibles[racine.lower()] = index
                for el in self.fichiers(donnees):
                    r = self.rel(el["FileRef"])
                    if r and r.lower().endswith(".aspx"):
                        self.cibles[r.lower()] = index if r == accueil else self.local(r[:-5]) + ".md"
            elif liste["BaseType"] == 1:
                self.cibles[racine.lower()] = posixpath.join(racine, "index.md")
                for el in self.fichiers(donnees):
                    r = self.rel(el["FileRef"])
                    if r is None:
                        continue
                    ext = posixpath.splitext(r)[1].lower()
                    if ext in IMAGES:
                        self.cibles[r.lower()] = self.local(r)
                    elif (ext in CONVERTIBLES or ext in TEXTES_BRUTS) and self.documents:
                        self.cibles[r.lower()] = self.local(r) + ".md"
            else:
                self.cibles[racine.lower()] = racine + ".md"
        for s in web.get("SousSites") or []:
            if s is not None:
                self.charger_web(s)

    def completer_page(self, el):
        """Copie faite sans le contenu des pages : on le relit dans le fichier .aspx de brut/."""
        rel = self.rel(el["FileRef"])
        if rel and rel.lower().endswith(".aspx") and not any(el.get(c) for c in CHAMPS_CONTENU):
            fichier = self.chemin_brut(rel)
            if existe(fichier):
                with open(lp(fichier), "rb") as f:
                    octets = f.read()
                try:
                    texte = octets.decode("utf-8-sig")
                except UnicodeDecodeError:  # certaines pages anciennes sont en Windows-1252
                    texte = octets.decode("cp1252", errors="replace")
                for nom, valeur in re.findall(r'<mso:(\w+) msdt:dt="string">(.*?)</mso:\1>', texte, re.S):
                    if nom in CHAMPS_ASPX and not el.get(nom):
                        valeur = unescape(valeur)
                        el[nom] = {"Url": valeur.split(",")[0].strip()} if nom == "BannerImageUrl" else valeur
        if not el.get("Title"):
            el["Title"] = self.titre_bandeau(el) or posixpath.splitext(el["FileLeafRef"])[0].replace("-", " ")

    @staticmethod
    def titre_bandeau(el):
        """Titre saisi dans la zone de titre de la page."""
        for champ in ("LayoutWebpartsContent", "CanvasContent1"):
            soup = BeautifulSoup(el.get(champ) or "", "html.parser")
            for balise in soup.find_all(attrs={"data-sp-webpartdata": True}) + soup.find_all(attrs={"data-sp-controldata": True}):
                try:
                    d = json.loads(balise.get("data-sp-webpartdata") or balise.get("data-sp-controldata"))
                except ValueError:
                    continue
                titre = (d.get("properties") or {}).get("title") if d.get("id") == WP_TITRE else None
                if titre and titre.strip():
                    return titre.strip()
        return None

    def page_accueil(self, web_rel, web, listes):
        """Page d'accueil du site : celle déclarée par SharePoint, sinon Home/Accueil/Default.aspx (si non vide)."""
        pages = {}
        for liste, donnees in listes:
            if self.est_pages(liste, donnees):
                for el in self.fichiers(donnees):
                    r = self.rel(el["FileRef"])
                    if r and r.lower().endswith(".aspx") and any(el.get(c) for c in CHAMPS_CONTENU):
                        pages[r.lower()] = r
        candidats = [posixpath.join(web_rel, web["WelcomePage"])] if web.get("WelcomePage") else []
        candidats += [r for r in pages.values() if posixpath.basename(r).lower() in ("home.aspx", "accueil.aspx", "default.aspx")]
        return next((pages[c.lower()] for c in candidats if c.lower() in pages), None)

    @staticmethod
    def fichiers(donnees):
        return [e for e in donnees["elements"] if e.get("FSObjType") == 0 and e.get("FileRef")]

    @staticmethod
    def est_pages(liste, donnees):
        if liste["BaseTemplate"] in (MODELE_PAGES, 850):
            return True
        return any("CanvasContent1" in e or "PublishingPageContent" in e for e in donnees["elements"][:5])

    # ---------- passe 2 : génération ----------

    def lancer(self):
        log(f"Conversion en Markdown vers {self.md}")
        self.charger_web("")
        for web_rel, web, navigation, listes in self.webs:
            for liste, donnees in listes:
                try:
                    if self.est_pages(liste, donnees):
                        for el in self.fichiers(donnees):
                            self.page(el)
                    elif liste["BaseType"] == 1:
                        self.bibliotheque(liste, donnees)
                    else:
                        self.liste(liste, donnees)
                except Exception as e:
                    log.erreur(f"conversion de « {liste['Title']} » : {e}")
            self.index_web(web_rel, web, navigation, listes)
        s = self.stats
        log(f"Markdown terminé : {s['pages']} pages, {s['documents']} documents convertis, "
            f"{s['images']} images, {s['listes']} listes, {s['echecs']} échecs de conversion.")

    def pied(self, sru):
        return f"\n\n---\n*Source : <{self.origine}{quote(sru, safe='/')}>*\n"

    def page(self, el):
        rel = self.rel(el["FileRef"])
        if not rel or not rel.lower().endswith(".aspx"):
            return
        dest = self.cibles.get(rel.lower()) or rel[:-5] + ".md"
        sru = el["FileRef"]
        titre = el.get("Title") or posixpath.splitext(el["FileLeafRef"])[0]
        morceaux = [f"# {titre}"]
        if el.get("Description"):
            morceaux.append("> " + el["Description"].replace("\n", "\n> "))
        banniere = el.get("BannerImageUrl")
        if isinstance(banniere, dict) and banniere.get("Url") and "_layouts/" not in banniere["Url"].lower():
            image = self.lien(banniere["Url"].split("?")[0], dest, sru)
            if not image.startswith("http"):  # images de banque d'images Microsoft : décoratives, non copiées
                morceaux.append(f"![]({image})")
        if el.get("CanvasContent1"):
            morceaux.append(self.canvas(el["CanvasContent1"], dest, sru))
        for champ in ("WikiField", "PublishingPageContent"):
            if el.get(champ):
                morceaux.append(self.html_vers_md(el[champ], dest, sru))
        ecrire_texte(self.md / dest, nettoyer("\n\n".join(m for m in morceaux if m) + self.pied(sru)))
        self.stats["pages"] += 1

    def canvas(self, html, dest, sru) -> str:
        soup = BeautifulSoup(html, "html.parser")
        controles = []
        for i, c in enumerate(soup.find_all(attrs={"data-sp-canvascontrol": True})):
            try:
                cd = json.loads(c.get("data-sp-controldata") or "{}")
            except ValueError:
                cd = {}
            pos = cd.get("position") or {}
            cle = (pos.get("zoneIndex", 0), pos.get("sectionIndex", 0), pos.get("controlIndex", 0), i)
            controles.append((cle, c, cd))
        controles.sort(key=lambda x: x[0])
        zones = {}  # zoneIndex -> [section repliable ou None, morceaux]
        for (zone, *_), c, cd in controles:
            courante = zones.setdefault(zone, [None, []])
            groupe = cd.get("zoneGroupMetadata") or {}
            if groupe.get("displayName"):
                courante[0] = groupe
            rte = c.find(attrs={"data-sp-rte": True})
            if rte is not None:
                courante[1].append(self.html_vers_md(rte.decode_contents(), dest, sru))
                continue
            wp = c.find(attrs={"data-sp-webpartdata": True})
            if wp is not None:
                try:
                    donnees = json.loads(wp["data-sp-webpartdata"])
                except ValueError:
                    continue
                courante[1].append(self.webpart(donnees, dest, sru, wp.find(attrs={"data-sp-htmlproperties": True})))
        sortie = []
        for groupe, morceaux in zones.values():
            contenu = "\n\n".join(m for m in morceaux if m)
            sortie.append(self.section_repliable(groupe, contenu) if groupe else contenu)
        return "\n\n".join(s for s in sortie if s)

    def section_repliable(self, groupe, contenu) -> str:
        """Section repliable SharePoint : bloc dépliable en MkDocs, titre simple en Markdown."""
        titre = self.libelle(groupe["displayName"])
        if self.mkdocs:
            marque = "???+" if groupe.get("isExpanded") else "???"
            corps = "\n".join("    " + l if l.strip() else "" for l in contenu.split("\n"))
            return f'{marque} note "{titre.replace(chr(34), chr(39))}"\n\n{corps}'
        niveau = min(max(int(groupe.get("headingLevel") or 3), 2), 6)
        return f"{'#' * niveau} {titre}\n\n{contenu}"

    def webpart(self, wp, dest, sru, props_html=None) -> str:
        titre = wp.get("title") or "Composant"
        ident = (wp.get("id") or "").lower()
        props = wp.get("properties") or {}
        spc = wp.get("serverProcessedContent") or {}
        textes = dict(spc.get("searchablePlainTexts") or {})
        htmls = dict(spc.get("htmlStrings") or {})
        images = dict(spc.get("imageSources") or {})
        liens = dict(spc.get("links") or {})
        # Certaines pages ne gardent ces valeurs que dans data-sp-htmlproperties.
        for el in props_html.find_all(attrs={"data-sp-prop-name": True}) if props_html else []:
            cle = el["data-sp-prop-name"]
            if el.name == "a" and el.get("href"):
                liens.setdefault(cle, el["href"])
            elif el.name == "img" and el.get("src"):
                images.setdefault(cle, el["src"])
            elif el.has_attr("data-sp-searchableplaintext"):
                textes.setdefault(cle, el.get_text())
            elif el.decode_contents().strip():
                htmls.setdefault(cle, el.decode_contents())
        for cle in ("baseUrl", "wopiurl"):
            liens.pop(cle, None)

        if ident == WP_SEPARATEUR:
            return "---"
        if ident in (WP_ESPACEUR, WP_TITRE):  # le titre de la page est déjà en tête
            return ""
        if ident == WP_MARKDOWN:
            return textes.get("code") or props.get("code") or ""
        if ident == WP_CODE:
            return self.bloc_code(unescape(textes.get("code") or ""), props.get("language"))
        if ident == WP_BOUTON and liens.get("linkUrl"):
            return f"**[{self.libelle(textes.get('label') or titre)}]({self.lien(liens['linkUrl'], dest, sru)})**"
        if ident == WP_FICHIER:
            url = liens.get("serverRelativeUrl") or props.get("file")
            if url:
                nom = textes.get("title") or posixpath.basename(unquote(urlparse(url).path))
                return f"Fichier : [{self.libelle(nom)}]({self.lien(url, dest, sru)})"

        sortie = []
        if textes.get("title"):
            sortie.append(f"### {textes.pop('title').strip()}")
        for cle, src in images.items():
            # Les vignettes des éléments (items[0]…, content[0]…) ne sont pas affichées seules.
            if src and "[" not in cle and "/_layouts/" not in src.lower():
                alt = self.libelle(props.get("altText") or "")
                image = f"![{alt}]({self.lien(src.split('?')[0], dest, sru)})"
                if props.get("linkUrl"):
                    image = f"[{image}]({self.lien(props['linkUrl'], dest, sru)})"
                sortie.append(image)
        if textes.get("captionText"):
            sortie.append(f"*{textes.pop('captionText')}*")
        for h in htmls.values():
            sortie.append(self.html_vers_md(h, dest, sru))
        puces = []
        for cle, url in liens.items():
            if not url:
                continue
            m = re.match(r"^(.*?\[\d+\])", cle)
            libelle = textes.pop(m.group(1) + ".title", None) if m else None
            puce = f"- [{self.libelle(libelle or url)}]({self.lien(url, dest, sru)})"
            description = textes.pop(m.group(1) + ".description", None) if m else None
            puces.append(puce + (f" — {description.strip()}" if description and description.strip() else ""))
        # Textes restants des éléments (boutons « En savoir plus »…) : ignorés ; les autres sont gardés.
        sortie.extend(t.strip() for c, t in textes.items() if "[" not in c and isinstance(t, str) and t.strip())
        if puces:
            sortie.append("\n".join(puces))
        if props.get("embedCode"):
            sortie.append(f"```html\n{props['embedCode']}\n```")
        if not sortie or (len(sortie) == 1 and sortie[0].startswith("###")):
            sortie.append(f"> *Composant « {titre} » non converti (contenu dynamique).*")
        return "\n\n".join(sortie)

    @staticmethod
    def libelle(texte: str) -> str:
        """Texte d'un lien Markdown : une ligne, sans crochets."""
        return re.sub(r"\s+", " ", texte).replace("[", "(").replace("]", ")").strip()

    @staticmethod
    def bloc_code(code: str, langage=None) -> str:
        # L'extrait de code SharePoint double souvent chaque saut de ligne.
        if "\n\n" in code and not re.search(r"[^\n]\n[^\n]", code):
            code = code.replace("\n\n", "\n")
        langage = re.sub(r"[^a-z0-9+#-]", "", (langage or "").lower())
        cloture = "````" if "```" in code else "```"
        return f"{cloture}{langage}\n{code.strip(chr(10))}\n{cloture}"

    def bibliotheque(self, liste, donnees):
        racine = self.rel(liste["RootFolder"]["ServerRelativeUrl"]) or ""
        index = posixpath.join(racine, "index.md")
        lignes = [f"# {liste['Title']}"]
        if liste.get("Description"):
            lignes.append(liste["Description"])
        arbre, dossiers = [], set()
        for el in sorted(donnees["elements"], key=lambda e: (e.get("FileRef") or "").lower()):
            rel = self.rel(el.get("FileRef") or "")
            if not rel or rel.lower().startswith(racine.lower() + "/forms/"):
                continue
            parties = rel[len(racine):].strip("/").split("/")
            est_dossier = el.get("FSObjType") == 1
            # Chaque dossier parent est affiché une fois, même s'il n'est pas dans la liste.
            for i in range(len(parties) if est_dossier else len(parties) - 1):
                chemin = "/".join(parties[: i + 1]).lower()
                if chemin not in dossiers:
                    dossiers.add(chemin)
                    arbre.append(f"{'    ' * i}- **{parties[i]}/**")
            if est_dossier:
                continue
            retrait = "    " * (len(parties) - 1)
            nom = el.get("FileLeafRef") or parties[-1]
            if not existe(self.chemin_brut(rel)):
                arbre.append(f"{retrait}- {nom} *(non téléchargé)*")
                continue
            c = self.convertir_fichier(rel)
            original = self.vers_brut(index, rel)
            if c and c != self.local(rel):
                arbre.append(f"{retrait}- [{nom}]({self.relatif(index, c)}) ([original]({original}))")
            elif c:
                arbre.append(f"{retrait}- [{nom}]({self.relatif(index, c)})")
            else:
                arbre.append(f"{retrait}- [{nom}]({original})")
        lignes.append("\n".join(arbre) if arbre else "*Bibliothèque vide.*")
        ecrire_texte(self.md / index, nettoyer("\n\n".join(lignes)))

    def convertir_fichier(self, rel):
        """Copie l'image ou convertit le document ; renvoie le chemin dans markdown/ ou None."""
        cible = self.cibles.get(rel.lower())
        src = self.chemin_brut(rel)
        if cible is None or not existe(src):
            return None
        dest = self.md / cible
        if existe(dest) and os.path.getmtime(lp(dest)) >= os.path.getmtime(lp(src)):
            return cible
        ext = posixpath.splitext(rel)[1].lower()
        try:
            if ext in IMAGES:
                os.makedirs(lp(dest.parent), exist_ok=True)
                shutil.copy2(lp(src), lp(dest))
                self.stats["images"] += 1
            else:
                texte = self.texte_document(src, rel)
                original = self.vers_brut(cible, rel)
                ecrire_texte(dest, nettoyer(
                    f"# {posixpath.basename(rel)}\n\n*Document d'origine : [{posixpath.basename(rel)}]({original})*\n\n{texte}"))
                self.stats["documents"] += 1
            return cible
        except Exception as e:
            self.stats["echecs"] += 1
            log.erreur(f"conversion de {rel} : {str(e).splitlines()[0] if str(e) else type(e).__name__}")
            self.cibles.pop(rel.lower(), None)
            return None

    def texte_document(self, src, rel) -> str:
        """Texte d'un document, gardé en cache : Markdown simple et MkDocs ne le convertissent qu'une fois."""
        cache = self.meta / "_texte" / (rel + ".md")
        if existe(cache) and os.path.getmtime(lp(cache)) >= os.path.getmtime(lp(src)):
            with open(lp(cache), encoding="utf-8") as f:
                return f.read()
        ext = posixpath.splitext(rel)[1].lower()
        if ext in TEXTES_BRUTS:
            with open(lp(src), "rb") as f:
                octets = f.read()
            try:
                texte = octets.decode("utf-8")
            except UnicodeDecodeError:
                texte = octets.decode("cp1252", errors="replace")
            texte = self.bloc_code(texte.replace("\r\n", "\n"))
        else:
            if self.markitdown is None:
                from markitdown import MarkItDown
                self.markitdown = MarkItDown()
            if ext == ".xlsm":  # classeur Excel à macros : lu comme un .xlsx
                from markitdown import StreamInfo
                texte = self.markitdown.convert(lp(src), stream_info=StreamInfo(extension=".xlsx")).text_content or ""
            else:
                texte = self.markitdown.convert(lp(src)).text_content or ""
        # Binaire déguisé (zip renommé en .txt…) : on retire les caractères de contrôle.
        texte = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", texte)
        ecrire_texte(cache, texte)
        log(f"  converti {rel}")
        return texte

    def valeur(self, v, dest, sru):
        if v is None:
            return ""
        if isinstance(v, bool):
            return "Oui" if v else "Non"
        if isinstance(v, dict):
            if "Url" in v:
                return f"[{v.get('Description') or v['Url']}]({self.lien(v['Url'], dest, sru)})"
            if "Label" in v:
                return v["Label"]
            return ", ".join(str(x) for x in v.values() if x not in (None, ""))
        if isinstance(v, list):
            return ", ".join(self.valeur(x, dest, sru) for x in v)
        s = str(v)
        if "<" in s and ">" in s:
            s = BeautifulSoup(s, "html.parser").get_text(" ", strip=True)
        if re.match(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$", s):
            s = s.replace("T", " ").rstrip("Z")
        return s

    def liste(self, liste, donnees):
        racine = self.rel(liste["RootFolder"]["ServerRelativeUrl"]) or ""
        dest = racine + ".md"
        sru = liste["RootFolder"]["ServerRelativeUrl"] + "/"
        noms = {"Attachments": "Pièces jointes"}
        noms.update({c["InternalName"]: c["Title"] for c in donnees["champs"]})
        colonnes = [ALIAS_COLONNES.get(c, c) for c in donnees.get("colonnes") or [] if c not in COLONNES_IGNOREES]
        if not colonnes:
            colonnes = [c["InternalName"] for c in donnees["champs"]
                        if not c.get("Hidden") and not c.get("ReadOnlyField")]
        colonnes = list(dict.fromkeys(colonnes))
        lignes = [f"# {liste['Title']}"]
        if liste.get("Description"):
            lignes.append(liste["Description"])
        elements = donnees["elements"]
        if not elements:
            lignes.append("*Liste vide.*")
        else:
            tableau = ["| " + " | ".join(echapper_cellule(noms.get(c, c)) for c in colonnes) + " |",
                       "|" + "---|" * len(colonnes)]
            for el in elements:
                cellules = []
                for c in colonnes:
                    if c == "Attachments":
                        pj = el.get("AttachmentFiles") or []
                        v = ", ".join(f"[{p['FileName']}]({self.lien(p['ServerRelativeUrl'], dest, sru)})" for p in pj)
                    else:
                        v = self.valeur(el.get(c, el.get(c + "Id")), dest, sru)
                    cellules.append(echapper_cellule(v))
                tableau.append("| " + " | ".join(cellules) + " |")
            lignes.append("\n".join(tableau))
        ecrire_texte(self.md / dest, nettoyer("\n\n".join(lignes)))
        self.stats["listes"] += 1

    def navigation(self, noeuds, dest, sru, prof=0):
        lignes = []
        for n in noeuds:
            url = n.get("Url") or ""
            titre = n.get("Title") or url
            lien = self.lien(url, dest, sru) if url and url != "http://linkless.header/" else ""
            lignes.append("    " * prof + (f"- [{titre}]({lien})" if lien else f"- {titre}"))
            lignes += self.navigation(n.get("Children") or [], dest, sru, prof + 1)
        return lignes

    def index_web(self, web_rel, web, navigation, listes):
        dest = self.plans.get(web_rel) or posixpath.join(web_rel, "index.md")
        sru = web["ServerRelativeUrl"].rstrip("/") + "/"
        lignes = [f"# Plan du site {web['Title']}" if dest.endswith("plan-du-site.md") else f"# {web['Title']}"]
        if web.get("Description"):
            lignes.append(web["Description"])
        for nom, titre in (("QuickLaunch", "Navigation"), ("TopNavigationBar", "Navigation supérieure")):
            nav = self.navigation(navigation.get(nom) or [], dest, sru)
            if nav:
                lignes += [f"## {titre}", "\n".join(nav)]
        pages, bibliotheques, autres = [], [], []
        for liste, donnees in listes:
            if self.est_pages(liste, donnees):
                for el in sorted(self.fichiers(donnees), key=lambda e: e["FileLeafRef"].lower()):
                    r = self.rel(el["FileRef"])
                    if r and r.lower().endswith(".aspx"):
                        titre = el.get("Title") or el["FileLeafRef"][:-5]
                        pages.append(f"- [{titre}]({self.relatif(dest, self.cibles.get(r.lower()) or r[:-5] + '.md')})")
            else:
                racine = self.rel(liste["RootFolder"]["ServerRelativeUrl"]) or ""
                cible = self.cibles.get(racine.lower())
                if cible:
                    ligne = f"- [{liste['Title']}]({self.relatif(dest, cible)}) ({liste['ItemCount']} éléments)"
                    (bibliotheques if liste["BaseType"] == 1 else autres).append(ligne)
        for titre, contenu in (("Pages", pages), ("Bibliothèques de documents", bibliotheques), ("Listes", autres)):
            if contenu:
                lignes += [f"## {titre}", "\n".join(contenu)]
        sous = [s for s in web.get("SousSites") or [] if s is not None]
        if sous:
            lignes.append("## Sous-sites")
            lignes.append("\n".join(f"- [{posixpath.basename(s)}]({self.relatif(dest, posixpath.join(s, 'index.md'))})"
                                    for s in sous))
        ecrire_texte(self.md / dest, nettoyer("\n\n".join(lignes)))
