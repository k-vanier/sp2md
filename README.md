# sp2md — SharePoint Online vers Markdown et MkDocs

Copie un site SharePoint Online en local, en fait une version Markdown, puis un site MkDocs
(thème Material) prêt à publier.

## Utilisation

Copier `dist\sp2md.exe` (un seul fichier) sur le poste et le lancer (double-clic) : un menu propose tout.

```
 1. Copier un site et le convertir (Markdown + site MkDocs)
 2. Copier un site seulement (sans conversion)
 3. Reconvertir une copie existante (Markdown + site MkDocs)
 4. Remplir un site de test
 5. Changer le dossier de travail
 6. Oublier la connexion (changer de compte)
 0. Quitter
```

Les choix 1 et 2 demandent l'URL, puis s'il faut convertir les documents, créer le site MkDocs et inclure les sous-sites.
Le choix 3 liste les copies trouvées dans le dossier de travail. Après chaque opération, on revient au menu.
La copie est créée dans le dossier de travail : au départ, le dossier d'où l'on lance l'exe.

Sans passer par le menu, en ligne de commande :

```bat
sp2md.exe https://tenant.sharepoint.com/sites/MonSite
```

1. Edge s'ouvre (Chrome à défaut) : se connecter à Microsoft 365, MFA compris.
   Le profil est gardé dans `%LOCALAPPDATA%\sp2md\profil-navigateur`, la connexion est mémorisée.
2. L'application détecte la connexion et télécharge tout via l'API REST de SharePoint :
   pages, bibliothèques, listes, pièces jointes, navigation, sous-sites.
3. La version Markdown est générée à côté, puis le site MkDocs.

Relancer la même commande reprend là où on s'est arrêté (les fichiers déjà copiés et de même taille sont ignorés).

## Résultat

```
MonSite/
  sp2md.log
  brut/                 copie fidèle (arborescence SharePoint)
    _sp2md/             métadonnées JSON (contenu des pages, listes, navigation)
    Shared Documents/...
    SitePages/*.aspx
  markdown/
    index.md            accueil : navigation, pages, bibliothèques, listes, sous-sites
    SitePages/*.md      pages converties
    Shared Documents/index.md, *.docx.md, *.pdf.md, images
    Lists/*.md          listes en tableaux
  mkdocs/
    mkdocs.yml          configuration (thème Material, menu, recherche) — réécrite à chaque génération
    docs/               pages + fichiers d'origine : le site ne dépend pas de brut/
    site/               SITE HTML PRÊT À PUBLIER
    LISEZMOI.md         comment publier
```

Dans `markdown/`, les liens internes pointent vers les `.md` ; les fichiers non convertibles (zip, etc.) pointent vers `brut/`.

## Site MkDocs

- La page d'accueil du site SharePoint devient l'accueil du site MkDocs ; l'ancien index devient « Plan du site ».
- Menu : Accueil, Pages, Documents, Listes, Sous-sites, Plan du site. Les documents convertis sont accessibles
  depuis la page de leur bibliothèque, avec un lien vers le fichier d'origine.
- Recherche intégrée, mode sombre, pages en `.html` : le site s'ouvre aussi en local (`site/index.html`).
- Aucune ressource externe (pas de Google Fonts) : il fonctionne sur un réseau interne.

Pour publier : copier le contenu de `mkdocs/site/` sur un serveur web (IIS, Apache, nginx…), ou sur
GitHub Pages / Netlify (voir `mkdocs/LISEZMOI.md`). Vérifier d'abord que le contenu peut être rendu public.

## Options

| Option | Effet |
|---|---|
| `--sortie DOSSIER` | créer la copie ailleurs que dans le dossier courant |
| `--sans-markdown` | télécharger seulement |
| `--sans-documents` | ne pas convertir Word/Excel/PowerPoint/PDF en Markdown |
| `--sans-sous-sites` | ignorer les sous-sites |
| `--sans-mkdocs` | ne pas créer le site MkDocs |
| `--markdown-seulement MonSite` | refaire la conversion à partir de `MonSite\brut` sans retélécharger |

## Remplir un site de test

Pour essayer sp2md sans toucher à un vrai site : créer un site vide (SharePoint > « + Créer un site »), puis :

```bat
sp2md.exe --remplir-test https://tenant.sharepoint.com/sites/TestSp2md
```

Après connexion et confirmation, le script ajoute, tous nommés « sp2md » :

- une bibliothèque « sp2md Documents » : Word, Excel, PowerPoint, PDF, CSV, ZIP, PNG,
  dossiers imbriqués, noms avec accents, `&`, parenthèses et apostrophe ;
- une liste « sp2md Contacts » : colonnes texte, lien, oui/non, date, choix, nombre, texte riche, pièce jointe ;
- deux pages modernes : texte (titres, listes, tableau, liens), image, liens rapides, séparateur,
  bloc Markdown, Actualités (dynamique), section sur deux colonnes ;
- des liens dans la navigation du site ;
- un sous-site « sp2md-sous-site » avec le même contenu (si l'administrateur autorise les sous-sites).

Rien d'existant n'est modifié ni supprimé ; relancer met à jour sans dupliquer.
Il faut le droit de modifier le site (propriétaire ou membre).

## Limites

- Composants dynamiques (Actualités, Contenu mis en avant, Power BI, Forms…) : remplacés par une mention.
- Mise en page (colonnes, couleurs) perdue ; la structure est conservée.
- On ne voit que ce que le compte connecté peut voir.

## Construire l'exe

```bat
build.bat
```

Produit `dist\sp2md.exe` (un seul fichier, ~140 Mo, aucun Python requis).
