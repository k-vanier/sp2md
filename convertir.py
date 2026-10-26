"""Construit la version Markdown à partir de la copie brute (dossier brut/)."""
import json
import os
import posixpath
import re
import shutil
from pathlib import Path
from urllib.parse import quote, unquote, urlparse

from bs4 import BeautifulSoup
from markdownify import markdownify

from commun import DOSSIER_META, ecrire_texte, existe, lire_json, log, lp
from telecharger import MODELE_PAGES

CONVERTIBLES = {".docx", ".pptx", ".xlsx", ".xls", ".pdf", ".html", ".htm", ".csv", ".txt", ".json", ".xml", ".msg", ".epub"}
IMAGES = {".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".bmp", ".ico"}

WP_MARKDOWN = "1ef5ed11-ce7b-44be-bc5e-4abd55101d16"
WP_SEPARATEUR = "2161a1c6-db61-4731-b97c-3cdb303f7cbb"
WP_ESPACEUR = "8654b779-4886-46d4-8ffb-b5ed960ee986"

# Colonnes de vue qui n'apportent rien dans un tableau.
COLONNES_IGNOREES = {"DocIcon", "Edit", "SelectTitle", "_UIVersionString", "AppAuthor", "AppEditor"}
ALIAS_COLONNES = {"LinkTitle": "Title", "LinkTitleNoMenu": "Title", "LinkFilename": "FileLeafRef",
                  "LinkFilenameNoMenu": "FileLeafRef"}


def nettoyer(md: str) -> str:
    md = re.sub(r"[ \t]+\n", "\n", md)
    return re.sub(r"\n{3,}", "\n\n", md).strip() + "\n"


def echapper_cellule(v) -> str:
    return str(v).replace("|", "\\|").replace("\r", "").replace("\n", "<br>")


class Convertisseur:
    def __init__(self, dossier_site: Path, documents=True):
        self.dossier = Path(dossier_site)
        self.brut = self.dossier / "brut"
        self.md = self.dossier / "markdown"
        self.meta = self.brut / DOSSIER_META
        site = lire_json(self.meta / "site.json")
        self.site = site
        self.racine_sru = site["ServerRelativeUrl"].rstrip("/")
        u = urlparse(site["Url"])
        self.origine = f"{u.scheme}://{u.netloc}"
        self.documents = documents
        self.cibles = {}   # chemin relatif (minuscule) -> fichier .md ou image dans markdown/
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

    def lien(self, href: str, depuis_md: str, source_sru: str) -> str:
        if not href or href.startswith(("#", "mailto:", "tel:", "javascript:", "data:")):
            return href
        u = urlparse(href)
        if u.scheme and u.netloc.lower() != urlparse(self.origine).netloc.lower():
            return href
        chemin = u.path
        if not u.scheme and not chemin.startswith("/"):
            chemin = posixpath.normpath(posixpath.join(posixpath.dirname(source_sru), unquote(chemin)))
        absolu = self.origine + quote(unquote(chemin), safe="/") + (f"?{u.query}" if u.query else "")
        rel = self.rel(chemin)
        if rel is None:
            return absolu
        ancre = f"#{u.fragment}" if u.fragment else ""
        c = self.cible(rel)
        if c is not None:
            return self.relatif("markdown/" + depuis_md, "markdown/" + c) + ancre
        if rel and existe(self.brut / rel):
            return self.relatif("markdown/" + depuis_md, "brut/" + rel)
        return absolu

    def html_vers_md(self, html: str, depuis_md: str, source_sru: str) -> str:
        soup = BeautifulSoup(html or "", "html.parser")
        for a in soup.find_all("a", href=True):
            a["href"] = self.lien(a["href"], depuis_md, source_sru)
        for img in soup.find_all("img", src=True):
            img["src"] = self.lien(img["src"].split("?")[0], depuis_md, source_sru)
        return markdownify(str(soup), heading_style="ATX", bullets="-").strip()

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
        self.webs.append((web_rel, web, navigation, listes))
        self.cibles[web_rel.lower()] = posixpath.join(web_rel, "index.md")

        for liste, donnees in listes:
            racine = self.rel(liste["RootFolder"]["ServerRelativeUrl"]) or ""
            if self.est_pages(liste, donnees):
                self.cibles[racine.lower()] = posixpath.join(web_rel, "index.md")
                for el in self.fichiers(donnees):
                    r = self.rel(el["FileRef"])
                    if r and r.lower().endswith(".aspx"):
                        self.cibles[r.lower()] = r[:-5] + ".md"
            elif liste["BaseType"] == 1:
                self.cibles[racine.lower()] = posixpath.join(racine, "index.md")
                for el in self.fichiers(donnees):
                    r = self.rel(el["FileRef"])
                    if r is None:
                        continue
                    ext = posixpath.splitext(r)[1].lower()
                    if ext in IMAGES:
                        self.cibles[r.lower()] = r
                    elif ext in CONVERTIBLES and self.documents:
                        self.cibles[r.lower()] = r + ".md"
            else:
                self.cibles[racine.lower()] = racine + ".md"
        for s in web.get("SousSites") or []:
            if s is not None:
                self.charger_web(s)

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
        dest = rel[:-5] + ".md"
        sru = el["FileRef"]
        titre = el.get("Title") or posixpath.splitext(el["FileLeafRef"])[0]
        morceaux = [f"# {titre}"]
        if el.get("Description"):
            morceaux.append("> " + el["Description"].replace("\n", "\n> "))
        banniere = el.get("BannerImageUrl")
        if isinstance(banniere, dict) and banniere.get("Url") and "_layouts/" not in banniere["Url"]:
            morceaux.append(f"![]({self.lien(banniere['Url'].split('?')[0], dest, sru)})")
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
            controles.append((cle, c))
        controles.sort(key=lambda x: x[0])
        sortie = []
        for _, c in controles:
            rte = c.find(attrs={"data-sp-rte": True})
            if rte is not None:
                sortie.append(self.html_vers_md(rte.decode_contents(), dest, sru))
                continue
            wp = c.find(attrs={"data-sp-webpartdata": True})
            if wp is not None:
                try:
                    donnees = json.loads(wp["data-sp-webpartdata"])
                except ValueError:
                    continue
                sortie.append(self.webpart(donnees, dest, sru, wp.find(attrs={"data-sp-htmlproperties": True})))
        return "\n\n".join(s for s in sortie if s)

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
        liens.pop("baseUrl", None)

        if ident == WP_SEPARATEUR:
            return "---"
        if ident == WP_ESPACEUR:
            return ""
        if ident == WP_MARKDOWN:
            return textes.get("code") or props.get("code") or ""

        sortie = []
        if textes.get("title"):
            sortie.append(f"### {textes.pop('title')}")
        for src in images.values():
            if src:
                alt = (props.get("altText") or "").replace("]", "")
                sortie.append(f"![{alt}]({self.lien(src.split('?')[0], dest, sru)})")
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
            puces.append(f"- [{(libelle or url).replace(']', '')}]({self.lien(url, dest, sru)})")
        sortie.extend(t for t in textes.values() if isinstance(t, str) and t.strip())
        if puces:
            sortie.append("\n".join(puces))
        if props.get("embedCode"):
            sortie.append(f"```html\n{props['embedCode']}\n```")
        if not sortie or (len(sortie) == 1 and sortie[0].startswith("###")):
            sortie.append(f"> *Composant « {titre} » non converti (contenu dynamique).*")
        return "\n\n".join(sortie)

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
                    arbre.append(f"{'  ' * i}- **{parties[i]}/**")
            if est_dossier:
                continue
            retrait = "  " * (len(parties) - 1)
            nom = el.get("FileLeafRef") or parties[-1]
            original = self.relatif("markdown/" + index, "brut/" + rel)
            c = self.convertir_fichier(rel)
            if c:
                arbre.append(f"{retrait}- [{nom}]({self.relatif(index, c)}) ([original]({original}))")
            else:
                arbre.append(f"{retrait}- [{nom}]({original})")
        lignes.append("\n".join(arbre) if arbre else "*Bibliothèque vide.*")
        ecrire_texte(self.md / index, nettoyer("\n\n".join(lignes)))

    def convertir_fichier(self, rel):
        """Copie l'image ou convertit le document ; renvoie le chemin dans markdown/ ou None."""
        cible = self.cibles.get(rel.lower())
        src = self.brut / rel
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
                if self.markitdown is None:
                    from markitdown import MarkItDown
                    self.markitdown = MarkItDown()
                texte = self.markitdown.convert(lp(src)).text_content or ""
                original = self.relatif("markdown/" + cible, "brut/" + rel)
                ecrire_texte(dest, nettoyer(
                    f"# {posixpath.basename(rel)}\n\n*Document d'origine : [{posixpath.basename(rel)}]({original})*\n\n{texte}"))
                self.stats["documents"] += 1
                log(f"  converti {rel}")
            return cible
        except Exception as e:
            self.stats["echecs"] += 1
            log.erreur(f"conversion de {rel} : {str(e).splitlines()[0] if str(e) else type(e).__name__}")
            self.cibles.pop(rel.lower(), None)
            return None

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
            lignes.append("  " * prof + (f"- [{titre}]({lien})" if lien else f"- {titre}"))
            lignes += self.navigation(n.get("Children") or [], dest, sru, prof + 1)
        return lignes

    def index_web(self, web_rel, web, navigation, listes):
        dest = posixpath.join(web_rel, "index.md")
        sru = web["ServerRelativeUrl"].rstrip("/") + "/"
        lignes = [f"# {web['Title']}"]
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
                        pages.append(f"- [{titre}]({self.relatif(dest, r[:-5] + '.md')})")
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
