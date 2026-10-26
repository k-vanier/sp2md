"""Remplit un site SharePoint de test avec un contenu varié, pour essayer sp2md.

Tout ce qui est créé porte le préfixe « sp2md » : bibliothèque, liste, pages,
liens de navigation, sous-site. Relancer le script met à jour sans dupliquer.
"""
import html
import io
import json
import struct
import time
import uuid
import zipfile
import zlib
from urllib.parse import quote
from xml.sax.saxutils import escape

from commun import log
from telecharger import ErreurSP, chemin_api

VERBOSE = "application/json;odata=verbose"
CT_PAGE_SITE = "0x0101009D1CB255DA76424F860D91F20E6C4118"
APPLI_PAGES = "b6917cb1-93a0-4b97-a84d-7cf49975d4ec"

WP_IMAGE = "d1d91016-032f-456d-98a4-721247c305e8"
WP_LIENS = "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd"
WP_SEPARATEUR = "2161a1c6-db61-4731-b97c-3cdb303f7cbb"
WP_MARKDOWN = "1ef5ed11-ce7b-44be-bc5e-4abd55101d16"
WP_ACTUALITES = "8c88f208-6c77-4bdb-86a0-0c47b4316588"

BIBLIOTHEQUE = "sp2md Documents"
LISTE = "sp2md Contacts"
SOUS_SITE = "sp2md-sous-site"


# ---------- écriture via l'API REST ----------

class Ecrivain:
    """Appels REST en écriture sur un site, avec le jeton anti-falsification (digest)."""

    def __init__(self, requete, url_web):
        self.r = requete
        self.web = url_web.rstrip("/")
        self._digest, self._quand = None, 0

    def _verifier(self, rep, url):
        if rep.status < 400:
            return
        try:
            detail = rep.json()["error"]["message"]["value"]
        except Exception:
            detail = rep.text()[:200]
        raise ErreurSP(rep.status, url, detail)

    def digest(self):
        if not self._digest or time.time() - self._quand > 1200:
            url = f"{self.web}/_api/contextinfo"
            rep = self.r.post(url, headers={"Accept": VERBOSE})
            self._verifier(rep, url)
            self._digest = rep.json()["d"]["GetContextWebInformation"]["FormDigestValue"]
            self._quand = time.time()
        return self._digest

    def get(self, chemin):
        url = f"{self.web}/_api/{chemin}"
        rep = self.r.get(url, headers={"Accept": VERBOSE})
        self._verifier(rep, url)
        return rep.json()["d"]

    def post(self, chemin, corps=None, octets=None, merge=False):
        url = f"{self.web}/_api/{chemin}"
        entetes = {"Accept": VERBOSE, "X-RequestDigest": self.digest()}
        donnees = octets
        if corps is not None:
            entetes["Content-Type"] = VERBOSE
            donnees = json.dumps(corps)
        if merge:
            entetes.update({"X-HTTP-Method": "MERGE", "IF-MATCH": "*"})
        for essai in range(5):
            rep = self.r.post(url, headers=entetes, data=donnees, timeout=300_000)
            if rep.status not in (429, 503):
                break
            time.sleep(10 * (essai + 1))
        self._verifier(rep, url)
        texte = rep.text()
        return json.loads(texte).get("d") if texte.strip() else None


def lit(valeur: str) -> str:
    """Littéral OData encodé pour l'URL."""
    return chemin_api(valeur)


def u(chemin: str) -> str:
    """Chemin encodé pour un lien HTML."""
    return quote(chemin, safe="/:")


# ---------- fichiers de test générés sans dépendance ----------

def fichier_docx(titre, paragraphes, tableau=None) -> bytes:
    def para(texte, gras=False, taille=None):
        rpr = ("<w:b/>" if gras else "") + (f'<w:sz w:val="{taille}"/>' if taille else "")
        return f'<w:p><w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{escape(texte)}</w:t></w:r></w:p>'

    corps = para(titre, gras=True, taille=36) + "".join(para(p) for p in paragraphes)
    if tableau:
        grille = "".join('<w:gridCol w:w="3000"/>' for _ in tableau[0])
        lignes = "".join("<w:tr>" + "".join(f"<w:tc>{para(c, gras=(i == 0))}</w:tc>" for c in ligne) + "</w:tr>"
                         for i, ligne in enumerate(tableau))
        corps += (f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>{grille}</w:tblGrid>'
                  f"{lignes}</w:tbl>{para('')}")
    tampon = io.BytesIO()
    with zipfile.ZipFile(tampon, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                   "</Types>")
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                   "</Relationships>")
        z.writestr("word/document.xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                   f"<w:body>{corps}<w:sectPr/></w:body></w:document>")
    return tampon.getvalue()


def fichier_pdf(titre, lignes) -> bytes:
    def txt(s):
        return s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")

    flux = f"BT /F1 20 Tf 72 760 Td ({txt(titre)}) Tj /F1 12 Tf"
    for l in lignes:
        flux += f" 0 -22 Td ({txt(l)}) Tj"
    flux = (flux + " ET").encode("cp1252", "replace")
    objets = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
        b"<< /Length %d >>\nstream\n" % len(flux) + flux + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
    ]
    sortie = bytearray(b"%PDF-1.4\n")
    positions = []
    for i, o in enumerate(objets, 1):
        positions.append(len(sortie))
        sortie += b"%d 0 obj\n" % i + o + b"\nendobj\n"
    xref = len(sortie)
    sortie += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objets) + 1)
    sortie += b"".join(b"%010d 00000 n \n" % p for p in positions)
    sortie += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objets) + 1, xref)
    return bytes(sortie)


def fichier_png(largeur=320, hauteur=120) -> bytes:
    lignes = bytearray()
    for y in range(hauteur):
        lignes.append(0)
        for x in range(largeur):
            bande = (x // 40 + y // 40) % 2
            lignes += bytes((20 + x * 60 // largeur, 90 + 60 * bande, 200 - y))

    def bloc(type_, donnees):
        return struct.pack(">I", len(donnees)) + type_ + donnees + struct.pack(">I", zlib.crc32(type_ + donnees) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + bloc(b"IHDR", struct.pack(">IIBBBBB", largeur, hauteur, 8, 2, 0, 0, 0))
            + bloc(b"IDAT", zlib.compress(bytes(lignes))) + bloc(b"IEND", b""))


def fichier_xlsx() -> bytes:
    from openpyxl import Workbook
    classeur = Workbook()
    f = classeur.active
    f.title = "Budget"
    for ligne in (["Poste", "Prévu", "Réalisé"], ["Formation", 12000, 9800], ["Matériel", 30000, 31250], ["Audit", 8000, 8000]):
        f.append(ligne)
    classeur.create_sheet("Notes").append(["Classeur généré par sp2md pour les tests"])
    tampon = io.BytesIO()
    classeur.save(tampon)
    return tampon.getvalue()


def fichier_pptx():
    try:
        from pptx import Presentation
    except ImportError:
        return None
    pres = Presentation()
    diapo = pres.slides.add_slide(pres.slide_layouts[0])
    diapo.shapes.title.text = "Présentation de test sp2md"
    diapo.placeholders[1].text = "Deux diapositives pour vérifier la conversion"
    diapo = pres.slides.add_slide(pres.slide_layouts[1])
    diapo.shapes.title.text = "Objectifs"
    diapo.placeholders[1].text = "Copier le site\nConvertir en Markdown\nVérifier les liens"
    tampon = io.BytesIO()
    pres.save(tampon)
    return tampon.getvalue()


def fichier_zip() -> bytes:
    tampon = io.BytesIO()
    with zipfile.ZipFile(tampon, "w") as z:
        z.writestr("lisez-moi.txt", "Archive de test : elle doit rester telle quelle (non convertie).")
    return tampon.getvalue()


# ---------- contenu des pages (format HTML stocké par SharePoint) ----------

def attr(obj) -> str:
    return html.escape(json.dumps(obj, ensure_ascii=False), quote=True)


def position(zone, section, ordre, facteur):
    return {"zoneIndex": zone, "sectionIndex": section, "controlIndex": ordre, "sectionFactor": facteur, "layoutIndex": 1}


def ctl_texte(contenu, zone, section=1, ordre=1, facteur=12):
    cd = {"controlType": 4, "displayMode": 2, "id": str(uuid.uuid4()), "position": position(zone, section, ordre, facteur),
          "emphasis": {}, "editorType": "CKEditor"}
    return (f'<div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="{attr(cd)}">'
            f'<div data-sp-rte="">{contenu}</div></div>')


def ctl_webpart(wp_id, titre, proprietes, spc, zone, section=1, ordre=1, facteur=12):
    instance = str(uuid.uuid4())
    cd = {"controlType": 3, "webPartId": wp_id, "id": instance, "position": position(zone, section, ordre, facteur),
          "displayMode": 2, "emphasis": {}}
    spc = {"htmlStrings": {}, "searchablePlainTexts": {}, "imageSources": {}, "links": {}, **spc}
    donnees = {"id": wp_id, "instanceId": instance, "title": titre, "description": titre,
               "serverProcessedContent": spc, "dataVersion": "1.0", "properties": proprietes}
    props_html = "".join(
        f'<div data-sp-prop-name="{html.escape(k)}" data-sp-searchableplaintext="true">{html.escape(v)}</div>'
        for k, v in spc["searchablePlainTexts"].items())
    props_html += "".join(f'<a data-sp-prop-name="{html.escape(k)}" href="{html.escape(v)}"></a>' for k, v in spc["links"].items())
    props_html += "".join(f'<img data-sp-prop-name="{html.escape(k)}" src="{html.escape(v)}" />' for k, v in spc["imageSources"].items())
    return (f'<div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="{attr(cd)}">'
            f'<div data-sp-webpart="" data-sp-webpartdataversion="1.0" data-sp-webpartdata="{attr(donnees)}">'
            f'<div data-sp-componentid="">{wp_id}</div><div data-sp-htmlproperties="">{props_html}</div></div></div>')


def wp_image(url_image, ids, zone):
    props = {"imageSourceType": 2, "altText": "Logo de test sp2md", "overlayText": "", "fixAspectRatio": False,
             "captionText": "Image hébergée dans la bibliothèque de test", "alignment": "Center", **ids}
    spc = {"searchablePlainTexts": {"captionText": props["captionText"]}, "imageSources": {"imageSource": url_image}}
    return ctl_webpart(WP_IMAGE, "Image", props, spc, zone)


def wp_liens(url_web, liens, zone):
    items = [{"id": i + 1, "thumbnailType": 3, "sourceItem": {"itemType": 2, "fileExtension": ""}} for i in range(len(liens))]
    props = {"items": items, "isMigrated": True, "layoutId": "List", "shouldShowThumbnail": True, "hideWebPartWhenEmpty": True,
             "dataProviderId": "QuickLinks", "listLayoutOptions": {"showDescription": False, "showIcon": True},
             "buttonLayoutOptions": {"showDescription": False, "buttonTreatment": 2, "iconPositionType": 2,
                                     "textAlignmentVertical": 2, "textAlignmentHorizontal": 2, "linesOfText": 2},
             "waffleLayoutOptions": {"iconSize": 1, "onlyShowThumbnail": False}, "imageWidth": 100}
    textes = {"title": "Liens utiles"}
    adresses = {"baseUrl": url_web}
    for i, (titre, url) in enumerate(liens):
        textes[f"items[{i}].title"] = titre
        adresses[f"items[{i}].sourceItem.url"] = url
    return ctl_webpart(WP_LIENS, "Liens rapides", props, {"searchablePlainTexts": textes, "links": adresses}, zone)


def wp_separateur(zone):
    return ctl_webpart(WP_SEPARATEUR, "Séparateur", {"length": 100, "weight": 1}, {}, zone)


def wp_markdown(code, zone):
    props = {"displayPreview": True, "lineWrapping": True, "miniMap": {"enabled": False}, "previewState": "Show", "theme": "Monokai"}
    return ctl_webpart(WP_MARKDOWN, "Markdown", props, {"searchablePlainTexts": {"code": code}}, zone)


def wp_actualites(zone):
    props = {"layoutId": "FeaturedNews", "dataProviderId": "news", "showChrome": True, "newsDataSourceProp": 1,
             "newsSiteList": [], "renderItemsSliderValue": 4}
    return ctl_webpart(WP_ACTUALITES, "Actualités", props, {"searchablePlainTexts": {"title": "Actualités"}}, zone)


def canvas(*controles) -> str:
    return "<div>" + "".join(controles) + "</div>"


# ---------- opérations SharePoint ----------

def obtenir_liste(e: Ecrivain, titre, modele, description):
    chemin = (f"web/lists/GetByTitle('{lit(titre)}')"
              "?$select=Id,ItemCount,ListItemEntityTypeFullName,RootFolder/ServerRelativeUrl&$expand=RootFolder")
    try:
        return e.get(chemin)
    except ErreurSP as err:
        if err.statut != 404:
            raise
    log(f"  création de « {titre} »")
    e.post("web/lists", {"__metadata": {"type": "SP.List"}, "Title": titre, "BaseTemplate": modele,
                         "Description": description, "AllowContentTypes": True, "ContentTypesEnabled": False})
    return e.get(chemin)


def deposer(e: Ecrivain, dossier_sru, nom, octets):
    d = e.post(f"web/GetFolderByServerRelativePath(decodedurl='{lit(dossier_sru)}')"
               f"/Files/AddUsingPath(decodedurl='{lit(nom)}',overwrite=true)", octets=octets)
    log(f"  déposé {dossier_sru}/{nom}")
    return d


def creer_dossier(e: Ecrivain, sru):
    e.post(f"web/folders/add('{lit(sru)}')")


def creer_page(e: Ecrivain, pages, ct_page, nom, titre, description, contenu):
    sru = f"{pages['RootFolder']['ServerRelativeUrl']}/{nom}"
    try:
        e.post(f"web/GetFolderByServerRelativePath(decodedurl='{lit(pages['RootFolder']['ServerRelativeUrl'])}')"
               f"/Files/AddTemplateFile(urlOfFile='{lit(sru)}',templateFileType=3)")
    except ErreurSP as err:
        if "exist" not in str(err).lower():
            raise
    fichier = f"web/GetFileByServerRelativePath(decodedurl='{lit(sru)}')"
    try:
        e.post(f"{fichier}/CheckOut()")
    except ErreurSP:
        pass  # pas d'extraction obligatoire, ou déjà extraite
    id_element = e.get(f"{fichier}/ListItemAllFields?$select=Id")["Id"]
    e.post(f"web/lists(guid'{pages['Id']}')/items({id_element})", merge=True, corps={
        "__metadata": {"type": pages["ListItemEntityTypeFullName"]},
        "Title": titre, "Description": description, "ContentTypeId": ct_page,
        "ClientSideApplicationId": APPLI_PAGES, "PageLayoutType": "Article", "PromotedState": 0,
        "CanvasContent1": contenu,
    })
    for action in ("CheckIn(comment='sp2md',checkintype=1)", "Publish(comment='sp2md')"):
        try:
            e.post(f"{fichier}/{action}")
        except ErreurSP:
            pass  # bibliothèque sans extraction ni brouillons
    log(f"  page {sru}")
    return sru


def etape(nom, fonction, *args):
    try:
        return fonction(*args)
    except ErreurSP as err:
        log.erreur(f"{nom} : {err}")
    return None


def remplir_web(e: Ecrivain, web, parent=None):
    """Bibliothèque, fichiers, liste et pages sur un site ; renvoie les URL utiles aux autres sites."""
    srv = web["ServerRelativeUrl"].rstrip("/")
    absolu = web["Url"].rstrip("/")
    racine_absolue = absolu[: len(absolu) - len(srv)] if srv else absolu

    log("Bibliothèque de documents")
    biblio = obtenir_liste(e, BIBLIOTHEQUE, 101, "Bibliothèque de test créée par sp2md")
    rb = biblio["RootFolder"]["ServerRelativeUrl"]
    for d in ("Dossier A", "Dossier A/Sous-dossier B"):
        etape(f"dossier {d}", creer_dossier, e, f"{rb}/{d}")

    fichiers = [
        ("", "Guide utilisateur.docx", fichier_docx("Guide utilisateur", [
            "Ce document Word sert à vérifier la conversion de sp2md.",
            "Il contient des accents (é, è, à, ç), un tableau et plusieurs paragraphes."],
            [["Étape", "Responsable", "Délai"], ["Copie du site", "sp2md", "10 min"], ["Relecture", "Moi", "1 jour"]])),
        ("Dossier A", "Procédure & qualité (v2).docx", fichier_docx("Procédure qualité", [
            "Nom de fichier avec esperluette, parenthèses et accents.",
            "1. Préparer  2. Contrôler  3. Archiver"])),
        ("", "Budget 2026.xlsx", fichier_xlsx()),
        ("", "Rapport annuel.pdf", fichier_pdf("Rapport annuel (test)", [
            "Document PDF généré par sp2md.", "Deuxième ligne du rapport.", "Fin du document."])),
        ("Dossier A/Sous-dossier B", "L'été au bureau.txt", "Fichier texte avec une apostrophe dans le nom.\n".encode("utf-8")),
        ("", "données.csv", "Nom;Ville;Effectif\nAtelier;Lyon;12\nSiège;Paris;40\n".encode("utf-8")),
        ("", "archives.zip", fichier_zip()),
        ("", "logo-sp2md.png", fichier_png()),
        ("", "Présentation.pptx", fichier_pptx()),
    ]
    deposes = {}
    for dossier, nom, octets in fichiers:
        if octets is None:
            continue
        d = etape(f"dépôt de {nom}", deposer, e, f"{rb}/{dossier}".rstrip("/"), nom, octets)
        if d:
            deposes[nom] = d

    log("Liste")
    liste = etape("liste", remplir_liste, e)
    resultat = {"biblio": rb, "liste": liste}

    log("Pages")
    sp = e.get("web/lists?$filter=BaseTemplate%20eq%20119"
               "&$select=Id,ListItemEntityTypeFullName,RootFolder/ServerRelativeUrl&$expand=RootFolder")["results"]
    if not sp:
        log.erreur("pas de bibliothèque « Pages du site » sur ce site : pages non créées")
        return resultat
    pages = sp[0]
    ct_page = CT_PAGE_SITE
    try:
        for ct in e.get(f"web/lists(guid'{pages['Id']}')/ContentTypes?$select=StringId")["results"]:
            if ct["StringId"].startswith(CT_PAGE_SITE):
                ct_page = ct["StringId"]
                break
    except ErreurSP:
        pass
    rp = u(pages["RootFolder"]["ServerRelativeUrl"])
    resultat["pages"] = pages["RootFolder"]["ServerRelativeUrl"]
    logo = u(f"{rb}/logo-sp2md.png")
    doc = lambda chemin: u(f"{rb}/{chemin}")  # noqa: E731
    ids_image = {}
    if "logo-sp2md.png" in deposes:
        ids_image = {"siteId": e.get("site?$select=Id")["Id"], "webId": web["Id"], "listId": biblio["Id"],
                     "uniqueId": deposes["logo-sp2md.png"].get("UniqueId", "")}

    if parent is None:
        ete = doc("Dossier A/Sous-dossier B/L'été au bureau.txt")
        accueil = canvas(
            ctl_texte(
                "<h2>Bienvenue sur le site de test</h2>"
                "<p>Cette page a été créée par <strong>sp2md</strong> pour vérifier la conversion en <em>Markdown</em>.</p>"
                f'<p>Voir aussi la <a href="{rp}/sp2md-Procedure.aspx">page Procédure</a>, '
                f'le <a href="{racine_absolue}{doc("Guide utilisateur.docx")}">guide utilisateur</a> '
                f'et le <a href="{doc("Budget 2026.xlsx")}">budget</a>.</p>'
                "<h3>Points clés</h3><ul><li>Un premier point</li><li>Un deuxième point avec <strong>du gras</strong></li></ul>"
                "<ol><li>Étape un</li><li>Étape deux</li></ol>"
                "<table><tbody><tr><td><strong>Service</strong></td><td><strong>Contact</strong></td></tr>"
                "<tr><td>Qualité</td><td>Jean Dupont</td></tr><tr><td>Production</td><td>Marie Martin</td></tr></tbody></table>",
                zone=1),
            wp_image(logo, ids_image, zone=2),
            wp_liens(absolu, [
                ("Guide utilisateur", f"{racine_absolue}{doc('Guide utilisateur.docx')}"),
                ("Rapport annuel (PDF)", f"{racine_absolue}{doc('Rapport annuel.pdf')}"),
                ("Liste des contacts", f"{racine_absolue}{u(liste)}" if liste else absolu),
                ("Site de Microsoft", "https://www.microsoft.com/fr-fr"),
            ], zone=3),
            wp_separateur(zone=4),
            wp_markdown("## Bloc Markdown\n\nTexte saisi dans le composant **Markdown**.\n\n- élément A\n- élément B", zone=5),
            wp_actualites(zone=6),
        )
        etape("page Accueil", creer_page, e, pages, ct_page, "sp2md-Accueil.aspx", "sp2md Accueil",
              "Page d'accueil du site de test", accueil)
        procedure = canvas(
            ctl_texte("<h2>Colonne de gauche</h2><p>Section sur deux colonnes.</p>"
                      f'<p>Retour à l\'<a href="{rp}/sp2md-Accueil.aspx">accueil</a>.</p>', zone=1, section=1, facteur=6),
            ctl_texte("<h2>Colonne de droite</h2>"
                      f'<p>Fichier avec apostrophe : <a href="{ete}">'
                      "L'été au bureau</a>.</p>"
                      f'<p>Lien vers le sous-site : <a href="{u(srv)}/{SOUS_SITE}">sous-site</a>.</p>', zone=1, section=2, facteur=6),
            ctl_texte("<h3>Section pleine largeur</h3><p>Dernier paragraphe de la page.</p>", zone=2),
        )
        etape("page Procédure", creer_page, e, pages, ct_page, "sp2md-Procedure.aspx", "sp2md Procédure",
              "Page sur deux colonnes", procedure)
    else:
        contenu = canvas(ctl_texte(
            "<h2>Page du sous-site</h2>"
            f'<p>Retour à l\'<a href="{u(parent)}">accueil du site parent</a>.</p>'
            f'<p>Document local : <a href="{doc("Guide utilisateur.docx")}">guide du sous-site</a>.</p>', zone=1))
        etape("page du sous-site", creer_page, e, pages, ct_page, "sp2md-Sous-site.aspx", "sp2md Sous-site",
              "Page créée dans le sous-site", contenu)
    return resultat


CHAMPS_LISTE = [
    '<Field Type="Text" DisplayName="Courriel" Name="Courriel" StaticName="Courriel"/>',
    '<Field Type="URL" DisplayName="Site web" Name="SiteWeb" StaticName="SiteWeb" Format="Hyperlink"/>',
    '<Field Type="Boolean" DisplayName="Actif" Name="Actif" StaticName="Actif"><Default>1</Default></Field>',
    '<Field Type="DateTime" DisplayName="Arrivée" Name="Arrivee" StaticName="Arrivee" Format="DateOnly"/>',
    '<Field Type="Choice" DisplayName="Service" Name="Service" StaticName="Service"><CHOICES>'
    "<CHOICE>Qualité</CHOICE><CHOICE>Production</CHOICE><CHOICE>Direction</CHOICE></CHOICES></Field>",
    '<Field Type="Number" DisplayName="Poste" Name="Poste" StaticName="Poste" Decimals="0"/>',
    '<Field Type="Note" DisplayName="Notes" Name="Notes" StaticName="Notes" RichText="TRUE" RichTextMode="FullHtml" NumLines="6"/>',
]


def remplir_liste(e: Ecrivain):
    liste = obtenir_liste(e, LISTE, 100, "Liste de test créée par sp2md")
    api = f"web/lists(guid'{liste['Id']}')"
    existants = {c["InternalName"] for c in e.get(f"{api}/fields?$select=InternalName")["results"]}
    for xml in CHAMPS_LISTE:
        nom = xml.split('Name="')[1].split('"')[0]
        if nom not in existants:
            e.post(f"{api}/fields/CreateFieldAsXml", {"parameters": {
                "__metadata": {"type": "SP.XmlSchemaFieldCreationInformation"}, "SchemaXml": xml, "Options": 24}})
    vue = e.get(f"{api}/DefaultView/ViewFields")
    if "Attachments" not in (vue.get("Items") or {}).get("results", []):
        etape("colonne Pièces jointes", e.post, f"{api}/DefaultView/ViewFields/AddViewField('Attachments')")
    if liste["ItemCount"] == 0:
        type_el = liste["ListItemEntityTypeFullName"]

        def url(u, texte):
            return {"__metadata": {"type": "SP.FieldUrlValue"}, "Url": u, "Description": texte}

        elements = [
            {"Title": "Jean Dupont", "Courriel": "jean.dupont@exemple.fr", "SiteWeb": url("https://www.exemple.fr", "Exemple"),
             "Actif": True, "Arrivee": "2024-03-01T00:00:00Z", "Service": "Qualité", "Poste": 1234,
             "Notes": "<p>Responsable <strong>qualité</strong>.<br>Deuxième ligne | avec une barre verticale.</p>"},
            {"Title": "Marie Martin", "Courriel": "marie.martin@exemple.fr", "Actif": False,
             "Arrivee": "2025-09-15T00:00:00Z", "Service": "Production", "Poste": 5678},
            {"Title": "Équipe Direction", "Service": "Direction", "Notes": "<p>Élément sans courriel.</p>"},
        ]
        for i, champs in enumerate(elements):
            d = e.post(f"{api}/items", {"__metadata": {"type": type_el}, **champs})
            if i == 0:
                e.post(f"{api}/items({d['Id']})/AttachmentFiles/add(FileName='cv-jean-dupont.txt')",
                       octets="CV de test en pièce jointe.\n".encode("utf-8"))
        log(f"  {len(elements)} éléments ajoutés")
    return liste["RootFolder"]["ServerRelativeUrl"]


def ajouter_navigation(e: Ecrivain, web, liens):
    barre = "TopNavigationBar" if web.get("WebTemplate") == "SITEPAGEPUBLISHING" else "QuickLaunch"
    existants = {n["Title"] for n in e.get(f"web/navigation/{barre}?$select=Title")["results"]}
    for titre, url, externe in liens:
        if titre not in existants:
            e.post(f"web/navigation/{barre}", {"__metadata": {"type": "SP.NavigationNode"},
                                               "Title": titre, "Url": url, "IsExternal": externe})
    log(f"  liens ajoutés à la navigation ({barre})")


def remplir(requete, web_connu):
    e = Ecrivain(requete, web_connu["Url"])
    web = e.get("web?$select=Id,Title,Url,ServerRelativeUrl,Language,WebTemplate")
    srv = web["ServerRelativeUrl"].rstrip("/")
    log(f"Remplissage du site « {web['Title']} »")
    r = remplir_web(e, web)
    accueil = f"{r.get('pages') or srv + '/SitePages'}/sp2md-Accueil.aspx"

    log("Navigation")
    liens = [("sp2md Accueil", u(accueil), False), ("sp2md Documents", u(r["biblio"]), False)]
    if r.get("liste"):
        liens.append(("sp2md Contacts", u(r["liste"]), False))
    liens.append(("Microsoft (lien externe)", "https://www.microsoft.com/fr-fr", True))
    etape("navigation", ajouter_navigation, e, web, liens)

    log("Sous-site")
    url_sous = f"{web['Url'].rstrip('/')}/{SOUS_SITE}"
    try:
        Ecrivain(requete, url_sous).get("web?$select=Id")
        log("  le sous-site existe déjà")
    except ErreurSP:
        try:
            e.post("web/webs/add", {"parameters": {
                "__metadata": {"type": "SP.WebCreationInformation"}, "Title": "sp2md Sous-site", "Url": SOUS_SITE,
                "Description": "Sous-site de test créé par sp2md", "Language": web["Language"],
                "WebTemplate": "STS#3", "UseSamePermissionsAsParentSite": True}})
            log(f"  sous-site créé : {url_sous}")
        except ErreurSP as err:
            log.erreur(f"création du sous-site refusée (souvent désactivée par l'administrateur) : {err}")
            url_sous = None
    if url_sous:
        es = Ecrivain(requete, url_sous)
        sous = es.get("web?$select=Id,Title,Url,ServerRelativeUrl,Language,WebTemplate")
        remplir_web(es, sous, parent=accueil)

    racine = web["Url"][: len(web["Url"]) - len(srv)] if srv else web["Url"]
    log(f"Site de test prêt : {racine}{u(accueil)}")
