"""sp2md : télécharge un site SharePoint Online et en fait une version Markdown.

Usage :
    sp2md                         (menu avec toutes les options)
    sp2md https://tenant.sharepoint.com/sites/MonSite
    sp2md --markdown-seulement MonSite
    sp2md --remplir-test https://tenant.sharepoint.com/sites/TestSp2md
"""
import argparse
import os
import re
import shutil
import sys
from contextlib import contextmanager
from pathlib import Path
from urllib.parse import urlparse

from commun import DOSSIER_META, existe, lire_json, log

PROFIL = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "sp2md" / "profil-navigateur"


def nom_dossier(web) -> str:
    sru = web["ServerRelativeUrl"].strip("/")
    nom = sru.rsplit("/", 1)[-1] if sru else urlparse(web["Url"]).netloc.split(".")[0]
    return re.sub(r'[<>:"/\\|?*]', "_", nom) or "site"


def ouvrir_navigateur(pw):
    derniere = None
    for canal in ("msedge", "chrome", None):
        try:
            return pw.chromium.launch_persistent_context(
                str(PROFIL), channel=canal, headless=False, no_viewport=True,
                args=["--start-maximized"],
            )
        except Exception as e:
            derniere = e
    raise RuntimeError(f"Impossible de lancer Edge ou Chrome : {derniere}")


@contextmanager
def session_sharepoint(url):
    """Ouvre le navigateur, attend la connexion, fournit (contexte, client, web)."""
    from playwright.sync_api import sync_playwright
    from telecharger import ClientSP, attendre_connexion, url_web_probable

    with sync_playwright() as pw:
        ctx = ouvrir_navigateur(pw)
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=120_000)
            client = ClientSP(ctx.request)
            yield ctx, client, attendre_connexion(page, client, url_web_probable(url))
        finally:
            ctx.close()


# ---------- actions ----------

def copier(url, sortie: Path, markdown=True, documents=True, sous_sites=True, mkdocs=True):
    from telecharger import Aspirateur

    with session_sharepoint(url) as (_, client, web):
        dossier = sortie / nom_dossier(web)
        log.ouvrir(dossier / "sp2md.log")
        log(f"Connecté. Site : « {web['Title']} » -> {dossier}")
        Aspirateur(client, web, dossier / "brut", sous_sites).lancer(web["Url"], web)
    if markdown:
        convertir(dossier, documents, mkdocs)
    bilan(dossier, markdown, markdown and mkdocs)


def refaire_markdown(dossier: Path, documents=True, mkdocs=True):
    if not existe(dossier / "brut" / DOSSIER_META / "site.json"):
        raise RuntimeError(f"{dossier} n'est pas une copie faite par sp2md (brut\\{DOSSIER_META}\\site.json absent).")
    log.ouvrir(dossier / "sp2md.log")
    convertir(dossier, documents, mkdocs)
    bilan(dossier, True, mkdocs)


def convertir(dossier: Path, documents: bool, mkdocs: bool):
    from convertir import Convertisseur
    Convertisseur(dossier, documents=documents).lancer()
    if mkdocs:
        from mkdocs_site import generer_mkdocs
        generer_mkdocs(dossier, documents)


def bilan(dossier: Path, markdown: bool, mkdocs=False):
    log(f"Terminé. Copie brute : {dossier / 'brut'}")
    if markdown:
        log(f"          Markdown    : {dossier / 'markdown' / 'index.md'}")
    if mkdocs:
        log(f"          Site MkDocs : {dossier / 'mkdocs' / 'site' / 'index.html'} (à publier : dossier mkdocs\\site)")
    if log.erreurs:
        log(f"{len(log.erreurs)} erreur(s), détail dans {dossier / 'sp2md.log'}")


def remplir_site_test(url):
    from remplir_test import remplir

    log.fermer()
    with session_sharepoint(url) as (ctx, _, web):
        print(f"\nLe site « {web['Title']} » ({web['Url']}) va recevoir du contenu de test :\n"
              "bibliothèque, liste, pages, liens de navigation et un sous-site, tous nommés « sp2md ».\n"
              "Rien d'existant n'est modifié ni supprimé.")
        if not oui_non("Continuer ?", defaut=False):
            log("Annulé.")
            return
        remplir(ctx.request, web)
    if log.erreurs:
        log(f"{len(log.erreurs)} erreur(s) :")
        for err in log.erreurs:
            log(f"  - {err}")


def oublier_connexion():
    if not PROFIL.exists():
        print("Aucune connexion mémorisée.")
    elif oui_non("Effacer la connexion mémorisée (il faudra se reconnecter) ?", defaut=False):
        shutil.rmtree(PROFIL, ignore_errors=True)
        print("Connexion oubliée.")


# ---------- menu ----------

def oui_non(question, defaut=True) -> bool:
    rep = input(f"{question} ({'O/n' if defaut else 'o/N'}) : ").strip().lower()
    return defaut if not rep else rep in ("o", "oui", "y", "yes")


def demander_url():
    url = input("URL du site SharePoint (vide pour revenir) : ").strip().strip('"')
    if not url:
        return None
    if not url.lower().startswith("http"):
        url = "https://" + url
    if ".sharepoint." not in urlparse(url).netloc.lower():
        if not oui_non(f"« {urlparse(url).netloc} » ne ressemble pas à SharePoint Online. Continuer quand même ?", defaut=False):
            return None
    return url


def copies_existantes(dossier: Path):
    copies = []
    for d in sorted(p for p in dossier.iterdir() if p.is_dir()) if dossier.is_dir() else []:
        site = d / "brut" / DOSSIER_META / "site.json"
        if existe(site):
            try:
                info = lire_json(site)
                copies.append((d, f"{info.get('Title', d.name)} - {info.get('Url', '')} (copié le {info.get('Date', '?')[:10]})"))
            except Exception:
                copies.append((d, d.name))
    return copies


def choisir_copie(travail: Path):
    copies = copies_existantes(travail)
    if copies:
        print(f"\nCopies trouvées dans {travail} :")
        for i, (_, texte) in enumerate(copies, 1):
            print(f"  {i}. {texte}")
    else:
        print(f"\nAucune copie trouvée dans {travail}.")
    rep = input("Numéro, ou chemin d'un dossier de copie (vide pour revenir) : ").strip().strip('"')
    if not rep:
        return None
    if rep.isdigit() and 1 <= int(rep) <= len(copies):
        return copies[int(rep) - 1][0]
    return Path(rep)


def menu():
    travail = Path.cwd()
    while True:
        print("\n" + "=" * 60)
        print(" sp2md - SharePoint Online vers Markdown")
        print(f" Dossier de travail : {travail}")
        print("=" * 60)
        print(" 1. Copier un site et le convertir (Markdown + site MkDocs)")
        print(" 2. Copier un site seulement (sans conversion)")
        print(" 3. Reconvertir une copie existante (Markdown + site MkDocs)")
        print(" 4. Remplir un site de test")
        print(" 5. Changer le dossier de travail")
        print(" 6. Oublier la connexion (changer de compte)")
        print(" 0. Quitter")
        try:
            choix = input("Votre choix : ").strip()
        except (KeyboardInterrupt, EOFError):
            return
        if choix == "0":
            return
        try:
            if choix in ("1", "2"):
                url = demander_url()
                if not url:
                    continue
                markdown = choix == "1"
                documents = markdown and oui_non("Convertir aussi Word, Excel, PowerPoint et PDF en Markdown ?")
                mkdocs = markdown and oui_non("Créer aussi le site MkDocs prêt à publier ?")
                sous_sites = oui_non("Inclure les sous-sites ?")
                copier(url, travail, markdown, documents, sous_sites, mkdocs)
            elif choix == "3":
                dossier = choisir_copie(travail)
                if not dossier:
                    continue
                documents = oui_non("Convertir aussi Word, Excel, PowerPoint et PDF en Markdown ?")
                refaire_markdown(dossier, documents, oui_non("Créer aussi le site MkDocs prêt à publier ?"))
            elif choix == "4":
                url = demander_url()
                if not url:
                    continue
                remplir_site_test(url)
            elif choix == "5":
                rep = input(f"Nouveau dossier de travail [{travail}] : ").strip().strip('"')
                if rep:
                    nouveau = Path(rep).expanduser().resolve()
                    if not nouveau.exists() and oui_non(f"{nouveau} n'existe pas. Le créer ?"):
                        nouveau.mkdir(parents=True)
                    if nouveau.is_dir():
                        travail = nouveau
                continue
            elif choix == "6":
                oublier_connexion()
                continue
            else:
                print("Choix inconnu.")
                continue
        except KeyboardInterrupt:
            log("Interrompu. Relancez la même opération pour reprendre : les fichiers déjà copiés sont conservés.")
        except Exception as e:
            log(f"ÉCHEC : {e}")
        finally:
            log.fermer()
        try:
            input("\nAppuyez sur Entrée pour revenir au menu...")
        except (KeyboardInterrupt, EOFError):
            return


# ---------- ligne de commande ----------

def main():
    if len(sys.argv) == 1:
        menu()
        return
    ap = argparse.ArgumentParser(prog="sp2md", description="Copie un site SharePoint Online en local et le convertit en Markdown. "
                                 "Sans argument, un menu propose toutes les options.")
    ap.add_argument("url", nargs="?", help="URL du site (ou d'une de ses pages)")
    ap.add_argument("--sortie", default=".", help="dossier où créer la copie (par défaut : dossier courant)")
    ap.add_argument("--sans-markdown", action="store_true", help="télécharger seulement")
    ap.add_argument("--sans-documents", action="store_true", help="ne pas convertir les documents Office/PDF en Markdown")
    ap.add_argument("--sans-sous-sites", action="store_true", help="ne pas descendre dans les sous-sites")
    ap.add_argument("--sans-mkdocs", action="store_true", help="ne pas créer le site MkDocs")
    ap.add_argument("--markdown-seulement", metavar="DOSSIER", help="refaire la conversion d'une copie existante")
    ap.add_argument("--remplir-test", metavar="URL", help="ajouter du contenu de test dans un site (pour essayer sp2md)")
    args = ap.parse_args()

    try:
        if args.remplir_test:
            remplir_site_test(args.remplir_test)
        elif args.markdown_seulement:
            refaire_markdown(Path(args.markdown_seulement), not args.sans_documents, not args.sans_mkdocs)
        elif args.url:
            copier(args.url, Path(args.sortie).resolve(), not args.sans_markdown,
                   not args.sans_documents, not args.sans_sous_sites, not args.sans_mkdocs)
        else:
            ap.error("indiquez l'URL du site, ou lancez sp2md sans argument pour le menu.")
    except KeyboardInterrupt:
        log("Interrompu. Relancez la même commande pour reprendre : les fichiers déjà copiés sont conservés.")
        sys.exit(130)
    except Exception as e:
        log(f"ÉCHEC : {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
