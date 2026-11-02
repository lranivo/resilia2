"""
Resilia Service : serveur web local (http://localhost:<port>).

- sert l'interface web (web/index.html) ;
- reçoit le classeur Resil_aammjj.xlsx ;
- exécute les scénarios choisis, un par un, chacun dans un processus
  worker (arrêtable) ;
- calcule la progression à partir des logs des workers.

Sécurité : le serveur n'écoute que sur 127.0.0.1, vérifie l'en-tête Host
(anti DNS-rebinding) et exige un en-tête personnalisé sur toutes les
requêtes POST (anti CSRF : un site web quelconque ne peut pas déclencher
un scénario ni fermer Excel).
"""

import io
import json
import os
import re
import signal
import socket
import subprocess
import sys
import threading
import time
from collections import deque
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

from app_paths import (
    charger_config,
    commande_enfant,
    dossier_logs,
    dossier_ressources,
    dossier_travail,
    dossier_web,
    fichier_service,
)
from scenarios import (
    PAR_ID,
    SCENARIOS,
    SuiviProgression,
    colonnes_manquantes,
)

TAILLE_MAX_FICHIER = 50 * 1024 * 1024
EXTENSIONS_AUTORISEES = (".xlsx", ".xlsm")
EN_TETES_MODELE = [
    "DPS/ACT", "SU", "Etat Résil", "Injoignabilité R1",
    "Statut", "Injoignabilité R2", "Architecture",
]

CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


# ============================================================
# PROCESSUS : ARRÊT D'UN ARBRE, EXCEL EN ARRIÈRE-PLAN
# ============================================================

def tuer_arbre(processus):
    """Arrête un processus et tous ses descendants (Chrome compris)."""

    if processus is None or processus.poll() is not None:
        return

    try:
        if os.name == "nt":
            subprocess.run(
                ["taskkill", "/PID", str(processus.pid), "/T", "/F"],
                stdin=subprocess.DEVNULL,
                capture_output=True,
                creationflags=CREATE_NO_WINDOW,
            )
        else:
            os.killpg(processus.pid, signal.SIGKILL)

    except Exception:
        try:
            processus.kill()
        except Exception:
            pass


def _pids_excel():
    resultat = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
        stdin=subprocess.DEVNULL,
        capture_output=True,
        text=True,
        errors="replace",
        creationflags=CREATE_NO_WINDOW,
    )

    return [
        int(pid)
        for pid in re.findall(r'"EXCEL\.EXE","(\d+)"', resultat.stdout)
    ]


def _pids_avec_fenetre_visible():
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    pids = set()

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def rappel(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            pid = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            pids.add(pid.value)

        return True

    user32.EnumWindows(rappel, 0)
    return pids


def fermer_excel_arriere_plan():
    """
    Ferme les instances Excel SANS fenêtre visible (celles que les
    scénarios ouvrent en arrière-plan et qui restent parfois bloquées).
    Les Excel visibles ne sont jamais fermés : ils peuvent contenir des
    classeurs non enregistrés.
    """

    if os.name != "nt":
        return {"fermes": 0, "visibles": 0,
                "message": "Disponible uniquement sous Windows."}

    pids = _pids_excel()
    visibles = _pids_avec_fenetre_visible()
    fermes = 0
    gardes = 0

    for pid in pids:
        if pid in visibles:
            gardes += 1
            continue

        retour = subprocess.run(
            ["taskkill", "/PID", str(pid), "/F"],
            stdin=subprocess.DEVNULL,
            capture_output=True,
            creationflags=CREATE_NO_WINDOW,
        )

        if retour.returncode == 0:
            fermes += 1

    if not pids:
        message = "Aucun processus Excel en cours."
    else:
        message = f"{fermes} processus Excel en arrière-plan fermé(s)."

        if gardes:
            message += (
                f" {gardes} Excel visible(s) conservé(s) "
                "(fermez-les vous-même)."
            )

    return {"fermes": fermes, "visibles": gardes, "message": message}


# ============================================================
# CLASSEUR EXCEL (en-têtes, modèle)
# ============================================================

def lire_en_tetes(chemin):
    from openpyxl import load_workbook

    classeur = load_workbook(chemin, read_only=True, data_only=True)

    try:
        feuille = classeur.active
        premiere = next(
            feuille.iter_rows(min_row=1, max_row=1, values_only=True), ()
        )

    finally:
        classeur.close()

    return [
        str(valeur).strip()
        for valeur in premiere
        if valeur is not None and str(valeur).strip()
    ]


def generer_modele():
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    classeur = Workbook()
    feuille = classeur.active
    feuille.title = "Resil"
    feuille.append(EN_TETES_MODELE)

    for cellule in feuille[1]:
        cellule.font = Font(bold=True, color="FFFFFF")
        cellule.fill = PatternFill("solid", fgColor="FF7900")
        cellule.alignment = Alignment(horizontal="center")

    for lettre, largeur in zip("ABCDEFG", (16, 16, 14, 20, 14, 20, 24)):
        feuille.column_dimensions[lettre].width = largeur

    feuille.freeze_panes = "A2"

    flux = io.BytesIO()
    classeur.save(flux)
    return flux.getvalue()


# ============================================================
# APPLICATION : ÉTAT ET EXÉCUTION DES SCÉNARIOS
# ============================================================

class Application:
    def __init__(self):
        self.verrou = threading.RLock()
        self.journal = deque(maxlen=400)
        self.fichier = None  # {"nom", "chemin", "taille", "en_tetes"}
        self.etat = "inactif"
        self.ids = []
        self.resultats = {}
        self.courant = None
        self.suivi = None
        self.debut = None
        self.fin = None
        self.message = ""
        self._arret = threading.Event()
        self._processus = None
        self._fil = None

    # ---------- fichier ----------

    def deposer_fichier(self, nom, contenu):
        with self.verrou:
            if self.etat == "en_cours":
                raise PermissionError(
                    "Un scénario est en cours : impossible de changer "
                    "de fichier."
                )

        nom = Path(nom).name
        nom = re.sub(r"[^\w.\- ()]", "_", nom, flags=re.UNICODE)

        if not nom.lower().endswith(EXTENSIONS_AUTORISEES):
            raise ValueError("Format attendu : .xlsx ou .xlsm")

        chemin = dossier_travail() / nom
        chemin.write_bytes(contenu)

        try:
            en_tetes = lire_en_tetes(chemin)
        except Exception as erreur:
            chemin.unlink(missing_ok=True)
            raise ValueError(
                f"Fichier Excel illisible : {erreur}"
            ) from erreur

        with self.verrou:
            self.fichier = {
                "nom": nom,
                "chemin": str(chemin),
                "taille": len(contenu),
                "en_tetes": en_tetes,
            }
            self.etat = "inactif"
            self.resultats = {}
            self.courant = None
            self.suivi = None
            self.debut = None
            self.fin = None
            self.message = ""
            self._ajouter_journal(
                f"Fichier chargé : {nom} ({len(en_tetes)} colonnes)"
            )

    # ---------- lancement / arrêt ----------

    def lancer(self, ids):
        with self.verrou:
            if self.etat == "en_cours":
                raise PermissionError("Un scénario est déjà en cours.")

            if not self.fichier:
                raise ValueError("Chargez d'abord un fichier Excel.")

            demandes = [
                s["id"] for s in SCENARIOS
                if s["id"] in ids and s["disponible"]
            ]

            if not demandes:
                raise ValueError("Aucun scénario valide sélectionné.")

            for id_scenario in demandes:
                manquantes = colonnes_manquantes(
                    id_scenario, self.fichier["en_tetes"]
                )

                if manquantes:
                    raise ValueError(
                        f"{PAR_ID[id_scenario]['libelle']} : colonne(s) "
                        f"manquante(s) : {', '.join(manquantes)}"
                    )

            self.ids = demandes
            self.resultats = {
                i: {"etat": "attente", "fait": 0, "total": None,
                    "erreurs": 0, "ignores": 0, "unite": "lignes"}
                for i in demandes
            }
            self.etat = "en_cours"
            self.courant = None
            self.suivi = None
            self.debut = time.time()
            self.fin = None
            self.message = ""
            self._arret.clear()
            self.journal.clear()

            self._fil = threading.Thread(
                target=self._executer, daemon=True
            )
            self._fil.start()

    def arreter(self):
        with self.verrou:
            if self.etat != "en_cours":
                return

            self.message = "Arrêt demandé..."

        self._arret.set()
        tuer_arbre(self._processus)

    # ---------- exécution ----------

    def _ajouter_journal(self, ligne):
        self.journal.append(ligne)

    def _executer(self):
        chemin = Path(self.fichier["chemin"])
        statut_final = "termine"

        for id_scenario in self.ids:
            if self._arret.is_set():
                statut_final = "arrete"
                break

            libelle = PAR_ID[id_scenario]["libelle"]
            suivi = SuiviProgression()

            with self.verrou:
                self.courant = id_scenario
                self.suivi = suivi
                self.resultats[id_scenario]["etat"] = "en_cours"
                self._ajouter_journal(f"=== {libelle} ===")

            journal_worker = (
                dossier_logs()
                / f"{datetime.now():%Y%m%d_%H%M%S}_{id_scenario}.log"
            )
            journal_worker.write_text("", encoding="utf-8")

            options = {"start_new_session": True} if os.name != "nt" else {}

            try:
                processus = subprocess.Popen(
                    commande_enfant(
                        "--worker", id_scenario,
                        "--file", str(chemin),
                        "--log", str(journal_worker),
                    ),
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=CREATE_NO_WINDOW,
                    **options,
                )

            except Exception as erreur:
                self._terminer_scenario(id_scenario, suivi, "erreur")
                self._ajouter_journal(f"[RESILIA][ERREUR] {erreur}")
                statut_final = "erreur"
                break

            self._processus = processus
            code = self._suivre(processus, journal_worker, suivi)
            self._processus = None

            if self._arret.is_set():
                self._terminer_scenario(id_scenario, suivi, "arrete")
                statut_final = "arrete"
                break

            if code != 0:
                self._terminer_scenario(id_scenario, suivi, "erreur")
                statut_final = "erreur"
                self.message = (
                    f"{libelle} : arrêt sur erreur (voir le journal)."
                )
                break

            self._terminer_scenario(id_scenario, suivi, "termine")

        if statut_final == "arrete":
            # Le worker a été tué : on nettoie l'Excel resté en arrière-plan.
            try:
                fermer_excel_arriere_plan()
            except Exception:
                pass

        with self.verrou:
            self.etat = statut_final
            self.courant = None
            self.fin = time.time()

            if statut_final == "arrete":
                self.message = "Exécution arrêtée."
            elif statut_final == "termine":
                self.message = "Tous les scénarios sont terminés."

    def _suivre(self, processus, journal_worker, suivi):
        """Lit le log du worker en continu jusqu'à la fin du processus."""

        reste = ""

        with open(journal_worker, "r", encoding="utf-8",
                  errors="replace") as flux:
            while True:
                morceau = flux.read()

                if morceau:
                    reste += morceau
                    *lignes, reste = reste.split("\n")

                    for ligne in lignes:
                        self._traiter_ligne(ligne, suivi)

                    continue

                if processus.poll() is not None:
                    # Dernière lecture après la fin du processus.
                    morceau = flux.read()

                    if morceau:
                        reste += morceau

                    for ligne in reste.split("\n"):
                        self._traiter_ligne(ligne, suivi)

                    break

                time.sleep(0.2)

        return processus.returncode

    def _traiter_ligne(self, ligne, suivi):
        ligne = ligne.rstrip("\r")

        if not ligne.strip():
            return

        with self.verrou:
            suivi.lire(ligne)
            self._ajouter_journal(ligne)

    def _terminer_scenario(self, id_scenario, suivi, etat):
        with self.verrou:
            resultat = self.resultats[id_scenario]
            resultat["etat"] = etat
            resultat["fait"] = suivi.fait
            resultat["total"] = suivi.total
            resultat["erreurs"] = suivi.erreurs
            resultat["ignores"] = suivi.ignores
            resultat["unite"] = suivi.unite

    # ---------- état pour l'interface ----------

    def etat_json(self):
        with self.verrou:
            en_tetes = self.fichier["en_tetes"] if self.fichier else []

            scenarios = []

            for s in SCENARIOS:
                resultat = self.resultats.get(s["id"], {})

                if (
                    s["id"] == self.courant
                    and self.suivi is not None
                ):
                    fait = self.suivi.fait
                    total = self.suivi.total
                    erreurs = self.suivi.erreurs
                    ignores = self.suivi.ignores
                    unite = self.suivi.unite
                else:
                    fait = resultat.get("fait", 0)
                    total = resultat.get("total")
                    erreurs = resultat.get("erreurs", 0)
                    ignores = resultat.get("ignores", 0)
                    unite = resultat.get("unite", "lignes")

                scenarios.append({
                    "id": s["id"],
                    "libelle": s["libelle"],
                    "disponible": s["disponible"],
                    "manquantes": (
                        colonnes_manquantes(s["id"], en_tetes)
                        if self.fichier and s["disponible"] else []
                    ),
                    "etat": resultat.get("etat", ""),
                    "fait": fait,
                    "total": total,
                    "erreurs": erreurs,
                    "ignores": ignores,
                    "unite": unite,
                })

            # Pourcentage global : chaque scénario sélectionné pèse autant.
            pourcentage = 0.0
            nombre = len(self.ids)

            if nombre:
                somme = 0.0

                for id_scenario in self.ids:
                    resultat = self.resultats[id_scenario]

                    if resultat["etat"] == "termine":
                        somme += 1.0
                    elif id_scenario == self.courant and self.suivi:
                        somme += self.suivi.fraction()

                pourcentage = 100.0 * somme / nombre

                if self.etat == "termine":
                    pourcentage = 100.0

            courant = None
            fait = 0
            total = 0
            ignores = 0
            unite = "lignes"

            if self.courant and self.suivi:
                courant = PAR_ID[self.courant]["libelle"]
                fait = self.suivi.fait
                total = self.suivi.total or 0
                ignores = self.suivi.ignores
                unite = self.suivi.unite

            fin = self.fin or time.time()
            ecoule = int(fin - self.debut) if self.debut else 0

            return {
                "fichier": (
                    {"nom": self.fichier["nom"],
                     "taille": self.fichier["taille"],
                     "en_tetes": en_tetes}
                    if self.fichier else None
                ),
                "scenarios": scenarios,
                "execution": {
                    "etat": self.etat,
                    "courant": courant,
                    "nombre": nombre,
                    "termines": sum(
                        1 for i in self.ids
                        if self.resultats[i]["etat"] == "termine"
                    ),
                    "pourcentage": round(pourcentage, 1),
                    "fait": fait,
                    "total": total,
                    "ignores": ignores,
                    "unite": unite,
                    "ecoule": ecoule,
                    "message": self.message,
                },
                "peut_telecharger": bool(
                    self.fichier and self.etat != "en_cours"
                ),
                "journal": list(self.journal)[-250:],
            }


# ============================================================
# SERVEUR HTTP
# ============================================================

class Gestionnaire(BaseHTTPRequestHandler):
    server_version = "Resilia2"
    protocol_version = "HTTP/1.1"

    def log_message(self, *arguments):
        pass

    # ---------- utilitaires ----------

    @property
    def app(self):
        return self.server.app

    def _hotes_autorises(self):
        port = self.server.server_address[1]
        return {f"127.0.0.1:{port}", f"localhost:{port}"}

    def _repondre(self, code, contenu=b"", type_contenu="text/plain",
                  en_tetes=None):
        if isinstance(contenu, str):
            contenu = contenu.encode("utf-8")

        self.send_response(code)
        self.send_header("Content-Type", f"{type_contenu}; charset=utf-8"
                         if type_contenu.startswith("text/")
                         or type_contenu == "application/json"
                         else type_contenu)
        self.send_header("Content-Length", str(len(contenu)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")

        for cle, valeur in (en_tetes or {}).items():
            self.send_header(cle, valeur)

        self.end_headers()

        if self.command != "HEAD":
            self.wfile.write(contenu)

    def _json(self, donnees, code=200):
        self._repondre(
            code,
            json.dumps(donnees, ensure_ascii=False),
            "application/json",
        )

    def _erreur(self, code, message):
        # Le corps de la requête n'a pas forcément été lu : on referme la
        # connexion pour ne pas désynchroniser le protocole HTTP/1.1.
        self.close_connection = True
        self._json({"ok": False, "erreur": message}, code)

    def _controle_hote(self):
        if self.headers.get("Host", "") not in self._hotes_autorises():
            self._erreur(403, "Hôte non autorisé.")
            return False

        return True

    def _controle_post(self):
        if self.headers.get("X-Resilia") != "1":
            self._erreur(403, "En-tête de sécurité manquant.")
            return False

        origine = self.headers.get("Origin")

        if origine:
            autorises = {f"http://{h}" for h in self._hotes_autorises()}

            if origine not in autorises:
                self._erreur(403, "Origine non autorisée.")
                return False

        return True

    def _corps(self, limite):
        longueur = int(self.headers.get("Content-Length") or 0)

        if longueur > limite:
            raise ValueError("Fichier trop volumineux (50 Mo maximum).")

        return self.rfile.read(longueur) if longueur else b""

    # ---------- GET ----------

    def do_GET(self):
        if not self._controle_hote():
            return

        chemin = urlparse(self.path).path

        try:
            if chemin in ("/", "/index.html"):
                page = (dossier_web() / "index.html").read_bytes()
                self._repondre(200, page, "text/html")

            elif chemin in ("/favicon.ico", "/resilia.ico"):
                self._icone()

            elif chemin == "/api/ping":
                self._json({"ok": True})

            elif chemin == "/api/state":
                self._json(self.app.etat_json())

            elif chemin == "/api/template":
                nom = f"Resil_{datetime.now():%y%m%d}.xlsx"
                self._repondre(
                    200,
                    generer_modele(),
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet",
                    {"Content-Disposition":
                     f'attachment; filename="{nom}"'},
                )

            elif chemin == "/api/download":
                self._telecharger()

            else:
                self._erreur(404, "Introuvable.")

        except Exception as erreur:
            self._erreur(500, f"{type(erreur).__name__}: {erreur}")

    def _icone(self):
        """Icône de l'onglet du navigateur (resilia.ico)."""

        for dossier in (dossier_ressources(), Path(__file__).resolve().parent):
            chemin = dossier / "resilia.ico"

            if chemin.is_file():
                self._repondre(200, chemin.read_bytes(), "image/x-icon")
                return

        self._erreur(404, "Icône introuvable.")

    def _telecharger(self):
        fichier = self.app.fichier

        if not fichier:
            self._erreur(404, "Aucun fichier chargé.")
            return

        if self.app.etat == "en_cours":
            self._erreur(409, "Un scénario est en cours.")
            return

        donnees = Path(fichier["chemin"]).read_bytes()
        nom = fichier["nom"]

        self._repondre(
            200,
            donnees,
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet",
            {"Content-Disposition": f'attachment; filename="{nom}"'},
        )

    # ---------- POST ----------

    def do_POST(self):
        if not self._controle_hote() or not self._controle_post():
            return

        chemin = urlparse(self.path).path

        try:
            if chemin == "/api/upload":
                nom = unquote(self.headers.get("X-Filename", ""))
                contenu = self._corps(TAILLE_MAX_FICHIER)

                if not nom or not contenu:
                    raise ValueError("Fichier vide ou sans nom.")

                self.app.deposer_fichier(nom, contenu)
                self._json({"ok": True})

            elif chemin == "/api/run":
                corps = json.loads(self._corps(65536) or b"{}")
                self.app.lancer(list(corps.get("scenarios", [])))
                self._json({"ok": True})

            elif chemin == "/api/stop":
                self.app.arreter()
                self._json({"ok": True})

            elif chemin == "/api/kill-excel":
                if self.app.etat == "en_cours":
                    raise PermissionError(
                        "Un scénario est en cours : arrêtez-le d'abord."
                    )

                self._json({"ok": True, **fermer_excel_arriere_plan()})

            else:
                self._erreur(404, "Introuvable.")

        except PermissionError as erreur:
            self._erreur(409, str(erreur))

        except ValueError as erreur:
            self._erreur(400, str(erreur))

        except Exception as erreur:
            self._erreur(500, f"{type(erreur).__name__}: {erreur}")


class Serveur(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = False

    def __init__(self, adresse, app):
        super().__init__(adresse, Gestionnaire)
        self.app = app


def _port_libre(port):
    for candidat in range(port, port + 30):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as test:
            try:
                test.bind(("127.0.0.1", candidat))
                return candidat
            except OSError:
                continue

    raise RuntimeError("Aucun port libre trouvé.")


def run_service():
    journal = open(
        dossier_logs() / "service.log", "a", encoding="utf-8", buffering=1
    )
    sys.stdout = journal
    sys.stderr = journal

    config = charger_config()
    port = _port_libre(int(config.get("port", 8765)))
    app = Application()
    serveur = Serveur(("127.0.0.1", port), app)

    fichier_service().write_text(
        json.dumps({"pid": os.getpid(), "port": port}), encoding="utf-8"
    )
    print(f"[SERVICE] Interface disponible sur http://localhost:{port}")

    try:
        serveur.serve_forever()

    finally:
        try:
            tuer_arbre(app._processus)
        except Exception:
            pass

        try:
            fichier_service().unlink(missing_ok=True)
        except Exception:
            pass
