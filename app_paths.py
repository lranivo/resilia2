"""
Chemins et configuration de Resilia2.

Fonctionne en mode script (python resilia2.py) comme en exécutable
PyInstaller (resilia2.exe) : les données de l'utilisateur (fichiers
traités, journaux, diagnostics, configuration) vivent dans
%LOCALAPPDATA%\\Resilia2, jamais à côté de l'exécutable.
"""

import json
import os
import sys
from pathlib import Path

NOM_APP = "Resilia2"

CONFIG_DEFAUT = {
    "port": 8765,
    "theme": "clair",
    "ouvrir_navigateur": True,
    # False : Chrome (Playwright) tourne sans fenêtre visible.
    # True  : fenêtre visible, utile pour diagnostiquer un problème.
    "navigateur_visible": False,
    "performance": {
        # "natif" : programme intégré (veille + activité Teams)
        # "vlc"   : lecture silencieuse VLC sans interface (veille seule)
        "mode": "natif",
        "simuler_activite_teams": True,
        "intervalle_secondes": 60,
        "chemin_vlc": "",
    },
}


def est_gele():
    return bool(getattr(sys, "frozen", False))


def dossier_code():
    return Path(__file__).resolve().parent


def dossier_ressources():
    """Dossier des ressources embarquées (web/, icône)."""
    return Path(getattr(sys, "_MEIPASS", dossier_code()))


def dossier_web():
    return dossier_ressources() / "web"


def dossier_donnees():
    surcharge = os.environ.get("RESILIA_DATA")

    if surcharge:
        base = Path(surcharge)
    else:
        racine = os.environ.get("LOCALAPPDATA")
        base = (
            Path(racine) / NOM_APP
            if racine
            else Path.home() / ".local" / "share" / NOM_APP.lower()
        )

    base.mkdir(parents=True, exist_ok=True)
    return base


def _sous_dossier(nom):
    dossier = dossier_donnees() / nom
    dossier.mkdir(parents=True, exist_ok=True)
    return dossier


def dossier_travail():
    return _sous_dossier("travail")


def dossier_diagnostics():
    return _sous_dossier("diagnostics")


def dossier_logs():
    return _sous_dossier("logs")


def fichier_service():
    """Fichier écrit par le service : port et PID réels."""
    return dossier_donnees() / "service.json"


def commande_enfant(*args):
    """Relance ce même programme avec d'autres arguments."""
    if est_gele():
        return [sys.executable, *args]

    return [sys.executable, str(dossier_code() / "resilia2.py"), *args]


def _fusion(defaut, lu):
    resultat = dict(defaut)

    for cle, valeur in (lu or {}).items():
        if isinstance(valeur, dict) and isinstance(defaut.get(cle), dict):
            resultat[cle] = _fusion(defaut[cle], valeur)
        else:
            resultat[cle] = valeur

    return resultat


def charger_config():
    chemin = dossier_donnees() / "config.json"

    try:
        lu = json.loads(chemin.read_text(encoding="utf-8"))
    except Exception:
        lu = {}

    config = _fusion(CONFIG_DEFAUT, lu)

    if not chemin.exists():
        sauver_config(config)

    return config


def sauver_config(config):
    chemin = dossier_donnees() / "config.json"

    try:
        chemin.write_text(
            json.dumps(config, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
    except Exception:
        pass
