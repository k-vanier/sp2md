"""Outils partagés : chemins longs Windows, JSON, journal."""
import json
import os
import sys
from datetime import datetime
from pathlib import Path

DOSSIER_META = "_sp2md"


def lp(chemin) -> str:
    """Chemin utilisable même au-delà de 260 caractères sous Windows."""
    s = os.path.abspath(str(chemin))
    if os.name == "nt" and len(s) > 240 and not s.startswith("\\\\?\\"):
        s = "\\\\?\\" + s
    return s


def existe(chemin) -> bool:
    return os.path.exists(lp(chemin))


def taille(chemin) -> int:
    return os.path.getsize(lp(chemin))


def ecrire_octets(chemin, donnees: bytes):
    os.makedirs(lp(Path(chemin).parent), exist_ok=True)
    tmp = lp(str(chemin) + ".part")
    with open(tmp, "wb") as f:
        f.write(donnees)
    os.replace(tmp, lp(chemin))


def ecrire_texte(chemin, texte: str):
    ecrire_octets(chemin, texte.encode("utf-8"))


def ecrire_json(chemin, obj):
    ecrire_texte(chemin, json.dumps(obj, ensure_ascii=False, indent=2))


def lire_json(chemin):
    with open(lp(chemin), encoding="utf-8") as f:
        return json.load(f)


class Journal:
    def __init__(self):
        self.fichier = None
        self.erreurs = []

    def ouvrir(self, chemin):
        self.fermer()
        os.makedirs(lp(Path(chemin).parent), exist_ok=True)
        self.fichier = open(lp(chemin), "a", encoding="utf-8")

    def fermer(self):
        """Ferme le fichier journal et oublie les erreurs (nouvelle opération)."""
        if self.fichier:
            self.fichier.close()
        self.fichier = None
        self.erreurs = []

    def __call__(self, msg: str):
        ligne = f"{datetime.now():%H:%M:%S} {msg}"
        try:
            print(ligne, flush=True)
        except UnicodeEncodeError:
            print(ligne.encode(sys.stdout.encoding or "ascii", "replace").decode(), flush=True)
        if self.fichier:
            self.fichier.write(ligne + "\n")
            self.fichier.flush()

    def erreur(self, msg: str):
        self.erreurs.append(msg)
        self("ERREUR " + msg)


log = Journal()
