"""
Resilia2 : point d'entrée unique (resilia2.exe).

    resilia2.exe                       -> fenêtre « Resilia Launcher »
    resilia2.exe --service             -> serveur web local
    resilia2.exe --performance         -> maintien en éveil du poste
    resilia2.exe --worker <id> --file <classeur> --log <journal>
                                       -> exécute un scénario

Le lanceur relance ce même programme avec ces arguments : un seul
exécutable à distribuer.
"""

import sys
import traceback
from datetime import datetime


def _option(arguments, nom):
    if nom in arguments:
        index = arguments.index(nom)

        if index + 1 < len(arguments):
            return arguments[index + 1]

    return None


def _journaliser_crash():
    try:
        from app_paths import dossier_logs

        with open(
            dossier_logs() / "crash.log", "a", encoding="utf-8"
        ) as fichier:
            fichier.write(f"\n=== {datetime.now():%Y-%m-%d %H:%M:%S} ===\n")
            fichier.write(traceback.format_exc())
    except Exception:
        pass


def main():
    arguments = sys.argv[1:]

    if arguments and arguments[0] == "--service":
        from service import run_service
        run_service()
        return 0

    if arguments and arguments[0] == "--performance":
        from performance import run_performance
        run_performance()
        return 0

    if arguments and arguments[0] == "--worker":
        from worker import lancer_worker

        id_scenario = arguments[1] if len(arguments) > 1 else ""
        fichier = _option(arguments, "--file")
        journal = _option(arguments, "--log")

        if not (id_scenario and fichier and journal):
            return 2

        return lancer_worker(id_scenario, fichier, journal)

    from launcher import run_launcher
    run_launcher()
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except BaseException:
        _journaliser_crash()
        raise
