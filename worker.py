"""
Worker : exécute UN scénario dans son propre processus.

Lancé par le service avec :
    resilia2.exe --worker <id_scenario> --file <classeur.xlsx> --log <log>

Un processus par scénario permet de l'arrêter proprement (le service tue
l'arbre de processus, Chrome compris) sans arrêter l'interface web.
"""

import os
import sys
import traceback
from pathlib import Path

from app_paths import dossier_diagnostics
from scenarios import charger_fonction


def _appliquer_chemins(fichier):
    """
    Les modules métier cherchent resil.xlsx à côté de leur code : on les
    redirige vers le classeur chargé dans l'interface web, et leurs
    diagnostics vers le dossier de données.
    """

    import check_wasac
    import isolation_client
    import salto_common

    for module in (check_wasac, isolation_client, salto_common):
        if hasattr(module, "EXCEL_FILE"):
            module.EXCEL_FILE = fichier

    for module in (isolation_client, salto_common):
        if hasattr(module, "DIAGNOSTIC_DIR"):
            module.DIAGNOSTIC_DIR = dossier_diagnostics()


def lancer_worker(id_scenario, fichier, fichier_log):
    flux = open(fichier_log, "a", encoding="utf-8", buffering=1)
    sys.stdout = flux
    sys.stderr = flux
    os.environ["PYTHONIOENCODING"] = "utf-8"

    print(f"[RESILIA] Scénario {id_scenario} | {fichier}")

    try:
        _appliquer_chemins(Path(fichier))
        fonction = charger_fonction(id_scenario)
        fonction()

    except BaseException as erreur:
        traceback.print_exc()
        print(
            f"[RESILIA][ERREUR FATALE] {type(erreur).__name__}: {erreur}"
        )
        return 1

    print("[RESILIA] Scénario terminé.")
    return 0
