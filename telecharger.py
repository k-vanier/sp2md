"""Copie brute d'un site SharePoint Online via son API REST.

L'authentification est celle du navigateur (Edge) piloté par Playwright :
l'utilisateur se connecte lui-même, on réutilise ensuite ses cookies.
"""
import re
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, urlparse

from commun import DOSSIER_META, ecrire_json, ecrire_octets, existe, log, taille

ENTETES_JSON = {"Accept": "application/json;odata=nometadata"}

# Segments qui marquent la fin de l'URL du site dans une URL de page ou de document.
SEGMENTS_FIN_SITE = re.compile(
    r"/(SitePages|Pages|Lists|Shared%20Documents|Shared Documents|Documents%20partages|"
    r"Documents partages|SiteAssets|Forms|_layouts|_api)(/|$)",
    re.IGNORECASE,
)

MODELE_PAGES = 119
CHAMPS_PAGES = ["CanvasContent1", "WikiField", "Description", "BannerImageUrl", "PublishingPageContent"]


class ErreurSP(Exception):
    def __init__(self, statut, url, detail=""):
        super().__init__(f"{statut} sur {url} {detail}".strip())
        self.statut = statut


def url_web_probable(url: str) -> str:
    """Devine l'URL du site à partir de n'importe quelle URL SharePoint."""
    u = urlparse(url.strip())
    chemin = u.path
    m = SEGMENTS_FIN_SITE.search(chemin)
    if m:
        chemin = chemin[: m.start()]
    elif chemin.lower().endswith(".aspx"):
        chemin = chemin.rsplit("/", 1)[0]
    return f"{u.scheme or 'https'}://{u.netloc}{chemin.rstrip('/')}"


def chemin_api(sru: str) -> str:
    """Encode un chemin relatif serveur pour GetFileByServerRelativePath(decodedurl='...')."""
    return quote(sru.replace("'", "''"), safe="/")


class ClientSP:
    def __init__(self, requete):
        self.requete = requete
        self.nb_requetes = 0

    def _get(self, url, binaire=False, essais=8):
        derniere = None
        for essai in range(essais):
            try:
                r = self.requete.get(url, headers=None if binaire else ENTETES_JSON, timeout=600_000)
            except Exception as e:  # coupure réseau, délai dépassé…
                derniere = e
                time.sleep(min(60, 2 ** essai))
                continue
            self.nb_requetes += 1
            if r.status in (429, 503):
                attente = r.headers.get("retry-after", "")
                attente = int(attente) if attente.isdigit() else min(120, 5 * (essai + 1))
                log(f"  SharePoint demande de ralentir, pause de {attente} s")
                time.sleep(attente)
                continue
            if r.status >= 400:
                raise ErreurSP(r.status, url, r.text()[:200].replace("\n", " "))
            if binaire:
                return r.body()
            try:
                return r.json()
            except Exception:
                raise ErreurSP("réponse non JSON (session expirée ?)", url)
        raise ErreurSP("abandon après plusieurs essais", url, str(derniere or ""))

    def json(self, url, essais=8):
        return self._get(url, essais=essais)

    def octets(self, url):
        return self._get(url, binaire=True)

    def tout(self, url):
        """Suit la pagination OData et renvoie tous les éléments."""
        resultat = []
        while url:
            d = self.json(url)
            resultat.extend(d.get("value", []))
            url = d.get("odata.nextLink")
        return resultat


def attendre_connexion(page, client, url_web, delai_s=900):
    """Attend que l'utilisateur soit connecté et que l'API du site réponde."""
    log("Connectez-vous dans la fenêtre du navigateur (MFA compris). J'attends...")
    fin = time.time() + delai_s
    candidat = url_web
    while time.time() < fin:
        try:
            ctx = page.evaluate(
                "() => window._spPageContextInfo ? _spPageContextInfo.webAbsoluteUrl : null"
            )
            if ctx and urlparse(ctx).netloc.lower() == urlparse(url_web).netloc.lower():
                candidat = ctx
        except Exception:
            pass  # page en cours de navigation
        try:
            return client.json(
                f"{candidat}/_api/web?$select=Title,Url,ServerRelativeUrl,Description,Id", essais=1
            )
        except Exception:
            page.wait_for_timeout(3000)
    raise RuntimeError("Connexion non détectée après 15 minutes.")


class Aspirateur:
    def __init__(self, client: ClientSP, web_racine: dict, dossier_brut: Path, sous_sites=True):
        self.c = client
        self.brut = dossier_brut
        self.meta = dossier_brut / DOSSIER_META
        self.racine_sru = web_racine["ServerRelativeUrl"].rstrip("/")
        self.sous_sites = sous_sites
        self.nb_fichiers = 0
        self.nb_ignores = 0
        self.octets = 0

    def rel(self, sru: str):
        """Chemin relatif au site racine, ou None si hors du site."""
        r = self.racine_sru
        if not r:
            return sru.lstrip("/")
        if sru.lower() == r.lower():
            return ""
        if sru.lower().startswith(r.lower() + "/"):
            return sru[len(r) + 1:]
        return None

    def lancer(self, url_web, web):
        ecrire_json(self.meta / "site.json", {
            "Url": web["Url"],
            "ServerRelativeUrl": web["ServerRelativeUrl"],
            "Title": web["Title"],
            "Date": datetime.now().isoformat(timespec="seconds"),
        })
        self.web(url_web)
        log(f"Téléchargement terminé : {self.nb_fichiers} fichiers téléchargés, "
            f"{self.nb_ignores} déjà à jour, {self.octets / 1e6:.1f} Mo, {self.c.nb_requetes} requêtes.")

    def web(self, url_web):
        web = self.c.json(f"{url_web}/_api/web?$select=Title,Url,ServerRelativeUrl,Description,Id,Language,WebTemplate")
        web_rel = self.rel(web["ServerRelativeUrl"].rstrip("/")) or ""
        meta = self.meta / web_rel
        log(f"=== Site « {web['Title']} » ({web['Url']})")
        try:
            web["WelcomePage"] = self.c.json(f"{url_web}/_api/web/RootFolder?$select=WelcomePage").get("WelcomePage")
        except ErreurSP:
            pass

        navigation = {}
        for nom in ("QuickLaunch", "TopNavigationBar"):
            try:
                navigation[nom] = self.c.tout(f"{url_web}/_api/web/navigation/{nom}?$expand=Children")
            except ErreurSP as e:
                log(f"  navigation {nom} indisponible ({e.statut})")
        ecrire_json(meta / "navigation.json", navigation)

        listes = self.c.tout(
            f"{url_web}/_api/web/lists?$filter=Hidden%20eq%20false"
            "&$select=Id,Title,BaseType,BaseTemplate,ItemCount,Description,RootFolder/ServerRelativeUrl"
            "&$expand=RootFolder"
        )
        ecrire_json(meta / "lists.json", listes)
        for liste in listes:
            try:
                self.liste(url_web, liste, meta)
            except ErreurSP as e:
                log.erreur(f"liste « {liste['Title']} » : {e}")

        sous_sites = []
        if self.sous_sites:
            try:
                sous_sites = self.c.tout(f"{url_web}/_api/web/webs?$select=Title,Url,ServerRelativeUrl")
            except ErreurSP as e:
                log.erreur(f"sous-sites de {url_web} : {e}")
        web["SousSites"] = [self.rel(s["ServerRelativeUrl"]) for s in sous_sites]
        ecrire_json(meta / "web.json", web)
        for s in sous_sites:
            try:
                self.web(s["Url"])
            except ErreurSP as e:
                log.erreur(f"sous-site {s['Url']} : {e}")

    def liste(self, url_web, liste, meta):
        api = f"{url_web}/_api/web/lists(guid'{liste['Id']}')"
        log(f"- {liste['Title']} ({liste['ItemCount']} éléments)")
        champs = self.c.tout(f"{api}/fields?$select=Title,InternalName,TypeAsString,Hidden,ReadOnlyField")
        noms = {c["InternalName"] for c in champs}
        try:
            vue = self.c.json(f"{api}/DefaultView/ViewFields")
            colonnes = vue.get("Items") or []
            if isinstance(colonnes, dict):
                colonnes = colonnes.get("results", [])
        except ErreurSP:
            colonnes = []

        if liste["BaseType"] == 1:  # bibliothèque de documents
            base = "Id,FileRef,FileLeafRef,FSObjType"
            pages = ",".join(["Title", "Modified"] + [n for n in CHAMPS_PAGES if n in noms])
            # Du plus complet au plus simple : une colonne refusée ne doit pas faire perdre le contenu des pages.
            essais = [f"$select={base},{pages},File/Length&$expand=File", f"$select={base},{pages}", f"$select={base}"]
            for i, requete in enumerate(essais):
                try:
                    elements = self.c.tout(f"{api}/items?{requete}&$top=2000")
                    break
                except ErreurSP as e:
                    if i == len(essais) - 1:
                        raise
                    log(f"  requête simplifiée ({e.statut})")
            for el in elements:
                el["File_x0020_Size"] = (el.pop("File", None) or {}).get("Length", el.get("File_x0020_Size"))
        else:
            try:
                elements = self.c.tout(f"{api}/items?$expand=AttachmentFiles&$top=2000")
            except ErreurSP:
                elements = self.c.tout(f"{api}/items?$top=2000")

        ecrire_json(meta / "listes" / f"{liste['Id']}.json", {
            "liste": liste, "champs": champs, "colonnes": colonnes, "elements": elements,
        })

        for el in elements:
            if liste["BaseType"] == 1 and el.get("FSObjType") == 0 and el.get("FileRef"):
                t = el.get("File_x0020_Size")
                self.fichier(url_web, el["FileRef"], int(t) if str(t or "").isdigit() else None)
            for pj in el.get("AttachmentFiles") or []:
                self.fichier(url_web, pj["ServerRelativeUrl"], None)

    def fichier(self, url_web, sru, taille_attendue):
        rel = self.rel(sru)
        if rel is None:
            return
        dest = self.brut / rel
        if existe(dest) and (taille_attendue is None or taille(dest) == taille_attendue):
            self.nb_ignores += 1
            return
        try:
            donnees = self.c.octets(
                f"{url_web}/_api/web/GetFileByServerRelativePath(decodedurl='{chemin_api(sru)}')/$value"
            )
        except ErreurSP as e:
            log.erreur(f"fichier {sru} : {e}")
            return
        ecrire_octets(dest, donnees)
        self.nb_fichiers += 1
        self.octets += len(donnees)
        log(f"    {rel} ({len(donnees) / 1024:.0f} Ko)")
