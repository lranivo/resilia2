"""
Resilia Launcher : petite fenêtre de pilotage.

Deux cartes :
  - Resilia Service  : démarre / arrête le serveur web local
                       (http://localhost:<port>) et ouvre l'interface ;
  - Performance Mode : démarre / arrête le maintien en éveil du poste.

Chaque service est un processus séparé (ce même programme relancé avec
--service ou --performance) : le PID affiché est celui du processus lancé.
"""

import json
import os
import subprocess
import threading
import tkinter as tk
import urllib.request
import webbrowser
from tkinter import messagebox

from app_paths import (
    charger_config,
    commande_enfant,
    dossier_donnees,
    dossier_ressources,
    fichier_service,
    sauver_config,
)
from service import tuer_arbre

CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0

THEMES = {
    "clair": {
        "fond": "#eef0f6", "carte": "#ffffff", "bord": "#e3e7ef",
        "titre": "#1e293b", "entete": "#94a3b8", "pid": "#9ca3af",
        "vert": "#16a34a", "vert_bord": "#8fe3b0", "ambre": "#d97706",
        "gris": "#a8afbd", "bouton": "#ffffff", "bouton_bord": "#e3e7ef",
        "bouton_txt": "#1f2937", "stop_fond": "#fde9e9",
        "stop_bord": "#f3b9b9", "stop_txt": "#dc2626",
        "off_txt": "#a8afbd", "pied": "#b3bccb",
    },
    
    "sombre": {
        "fond": "#050505",
        "carte": "#111111",
        "bord": "#005c38",

        "titre": "#e8e8e8",
        "entete": "#525252",
        "pid": "#555555",

        "vert": "#00e887",
        "vert_bord": "#005c38",
        "ambre": "#f0b84b",

        "gris": "#666666",

        "bouton": "#252525",
        "bouton_bord": "#292929",
        "bouton_txt": "#d4d4d4",

        "stop_fond": "#291414",
        "stop_bord": "#682323",
        "stop_txt": "#ff4545",

        "off_txt": "#666666",
        "pied": "#383838",
    },

}


class Processus:
    """Un service lancé en arrière-plan (ce même programme + un argument)."""

    def __init__(self, argument):
        self.argument = argument
        self.processus = None

    @property
    def vivant(self):
        return self.processus is not None and self.processus.poll() is None

    @property
    def pid(self):
        return self.processus.pid if self.vivant else None

    def demarrer(self):
        if self.vivant:
            return

        options = {"start_new_session": True} if os.name != "nt" else {}

        self.processus = subprocess.Popen(
            commande_enfant(self.argument),
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=CREATE_NO_WINDOW,
            **options,
        )

    def arreter(self):
        tuer_arbre(self.processus)


class Lanceur:
    LARGEUR = 420
    HAUTEUR = 352

    def __init__(self):
        self.config = charger_config()
        self.nom_theme = (
            self.config.get("theme")
            if self.config.get("theme") in THEMES else "clair"
        )

        self.service = Processus("--service")
        self.performance = Processus("--performance")
        self.port = None
        self._ping_en_cours = False
        self._resultat_ping = None
        self._interface_ouverte = False
        self._signature = None

        self.racine = tk.Tk()
        self.racine.title("Resilia Launcher")
        self.racine.resizable(False, False)

        self.k = max(1.0, self.racine.winfo_fpixels("1i") / 96.0)

        try:
            self.racine.iconbitmap(
                str(dossier_ressources() / "resilia.ico")
            )
        except Exception:
            pass

        self.canvas = tk.Canvas(
            self.racine,
            width=self.s(self.LARGEUR),
            height=self.s(self.HAUTEUR),
            highlightthickness=0,
            bd=0,
        )
        self.canvas.pack()

        self.racine.protocol("WM_DELETE_WINDOW", self.quitter)
        self.dessiner()
        self.racine.after(300, self.tick)

    # ---------- géométrie ----------

    def s(self, valeur):
        return int(round(valeur * self.k))

    @property
    def c(self):
        return THEMES[self.nom_theme]

    @staticmethod
    def police(taille, graisse="normal"):
        return ("Segoe UI", taille, graisse)

    def rectangle_arrondi(self, x1, y1, x2, y2, rayon, **options):
        s = self.s
        x1, y1, x2, y2, r = s(x1), s(y1), s(x2), s(y2), s(rayon)
        points = [
            x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
            x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
            x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
        ]
        return self.canvas.create_polygon(points, smooth=True, **options)

    def texte(self, x, y, contenu, taille, graisse="normal",
              couleur=None, ancre="w", etiquettes=()):
        return self.canvas.create_text(
            self.s(x), self.s(y), text=contenu, anchor=ancre,
            font=self.police(taille, graisse),
            fill=couleur or self.c["titre"], tags=etiquettes,
        )

    # ---------- états ----------

    def etat_service(self):
        if not self.service.vivant:
            return "STOPPED"

        return "RUNNING" if self.port else "STARTING"

    def etat_performance(self):
        return "RUNNING" if self.performance.vivant else "STOPPED"

    # ---------- dessin ----------

    def bouton(self, etiquette, x1, y1, x2, y2, libelle, style, action):
        c = self.c

        if style == "rouge":
            fond, bord, txt = c["stop_fond"], c["stop_bord"], c["stop_txt"]
        elif style == "inactif":
            fond, bord, txt = c["bouton"], c["bouton_bord"], c["off_txt"]
        else:
            fond, bord, txt = c["bouton"], c["bouton_bord"], c["bouton_txt"]

        self.rectangle_arrondi(
            x1, y1, x2, y2, 8, fill=fond, outline=bord,
            tags=(etiquette,),
        )
        self.texte(
            (x1 + x2) / 2, (y1 + y2) / 2, libelle, 9, "normal", txt,
            "center", (etiquette,),
        )

        if style != "inactif":
            self.canvas.tag_bind(
                etiquette, "<Button-1>", lambda _e: action()
            )
            self.canvas.tag_bind(
                etiquette, "<Enter>",
                lambda _e: self.canvas.config(cursor="hand2"),
            )
            self.canvas.tag_bind(
                etiquette, "<Leave>",
                lambda _e: self.canvas.config(cursor=""),
            )

    def carte(self, cle, titre, etat, pid, y, action_start, action_stop):
        c = self.c

        actif = etat in ("RUNNING", "STARTING")
        couleur_etat = (
            c["vert"] if etat == "RUNNING"
            else c["ambre"] if etat == "STARTING" else c["off_txt"]
        )
        bord = c["vert_bord"] if etat == "RUNNING" else c["bord"]

        self.rectangle_arrondi(
            30, y, 390, y + 102, 14, fill=c["carte"], outline=bord,
            width=1,
        )

        # pastille + titre
        rayon = self.s(5)
        cx, cy = self.s(52), self.s(y + 30)
        self.canvas.create_oval(
            cx - rayon, cy - rayon, cx + rayon, cy + rayon,
            fill=couleur_etat if actif else c["gris"], outline="",
        )
        self.texte(66, y + 30, titre, 12, "bold")
        self.texte(47, y + 58, etat, 8, "bold", couleur_etat)
        self.texte(
            47, y + 80, f"PID: {pid}" if pid else "PID: -", 8,
            "normal", c["pid"],
        )

        self.bouton(
            f"{cle}_start", 282, y + 14, 374, y + 46, "Start", "normal",
            action_start,
        )
        self.bouton(
            f"{cle}_stop", 282, y + 56, 374, y + 88, "Stop",
            "rouge" if actif else "inactif", action_stop,
        )

    def dessiner(self):
        c = self.c
        self.canvas.delete("all")
        self.canvas.config(bg=c["fond"])

        # en-tête
        self.texte(30, 52, "CONTROL PANEL", 8, "bold", c["entete"])

        soleil = "☀" if self.nom_theme == "clair" else "☾"
        self.texte(
            338, 52, soleil, 13, "normal", c["entete"], "center",
            ("theme",),
        )
        self.texte(
            374, 52, "▼", 9, "normal", c["entete"], "center", ("menu",),
        )

        for etiquette, action in (
            ("theme", self.basculer_theme), ("menu", self.menu),
        ):
            self.canvas.tag_bind(
                etiquette, "<Button-1>", lambda _e, a=action: a()
            )
            self.canvas.tag_bind(
                etiquette, "<Enter>",
                lambda _e: self.canvas.config(cursor="hand2"),
            )
            self.canvas.tag_bind(
                etiquette, "<Leave>",
                lambda _e: self.canvas.config(cursor=""),
            )

        etat_service = self.etat_service()
        etat_performance = self.etat_performance()

        self.carte(
            "service", "Resilia Service", etat_service,
            self.service.pid, 80, self.demarrer_service,
            self.arreter_service,
        )
        self.carte(
            "performance", "Performance Mode", etat_performance,
            self.performance.pid, 196, self.demarrer_performance,
            self.arreter_performance,
        )

        self.texte(
            self.LARGEUR / 2, 322, "RESILIA • OBS", 7, "bold", c["pied"],
            "center",
        )

        self._signature = self.signature()

    def signature(self):
        return (
            self.etat_service(), self.service.pid,
            self.etat_performance(), self.performance.pid,
            self.nom_theme,
        )

    # ---------- actions ----------

    def demarrer_service(self):
        if self.service.vivant:
            self.ouvrir_interface()
            return

        try:
            fichier_service().unlink(missing_ok=True)
        except Exception:
            pass

        self.port = None
        self._interface_ouverte = False

        try:
            self.service.demarrer()
        except Exception as erreur:
            messagebox.showerror(
                "Resilia Launcher",
                f"Impossible de démarrer le service :\n{erreur}",
            )

        self.dessiner()

    def arreter_service(self):
        if self.service.vivant:
            self.service.arreter()

        self.port = None
        self._interface_ouverte = False
        self.dessiner()

    def demarrer_performance(self):
        if not self.performance.vivant:
            try:
                self.performance.demarrer()
            except Exception as erreur:
                messagebox.showerror(
                    "Resilia Launcher",
                    f"Impossible de démarrer le mode performance :\n"
                    f"{erreur}",
                )

        self.dessiner()

    def arreter_performance(self):
        if self.performance.vivant:
            self.performance.arreter()

        self.dessiner()

    def ouvrir_interface(self):
        if self.port:
            webbrowser.open(f"http://localhost:{self.port}")

    def ouvrir_donnees(self):
        dossier = str(dossier_donnees())

        try:
            if os.name == "nt":
                os.startfile(dossier)  # noqa: S606
            else:
                subprocess.Popen(["xdg-open", dossier])
        except Exception:
            pass

    def basculer_theme(self):
        self.nom_theme = "sombre" if self.nom_theme == "clair" else "clair"
        self.config["theme"] = self.nom_theme
        sauver_config(self.config)
        self.dessiner()

    def menu(self):
        menu = tk.Menu(self.racine, tearoff=0)
        menu.add_command(
            label="Ouvrir l'interface web",
            command=self.ouvrir_interface,
            state="normal" if self.port else "disabled",
        )
        menu.add_command(
            label="Ouvrir le dossier des données (journaux, diagnostics)",
            command=self.ouvrir_donnees,
        )
        menu.add_separator()
        menu.add_command(label="Quitter", command=self.quitter)

        x = self.racine.winfo_rootx() + self.s(374)
        y = self.racine.winfo_rooty() + self.s(64)

        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def quitter(self):
        if self.service.vivant or self.performance.vivant:
            if not messagebox.askyesno(
                "Resilia Launcher",
                "Des services sont en cours.\n"
                "Les arrêter (un scénario en cours sera interrompu) "
                "et quitter ?",
            ):
                return

        self.service.arreter()
        self.performance.arreter()
        self.racine.destroy()

    # ---------- boucle ----------

    def _tester_service(self):
        port = None

        try:
            infos = json.loads(
                fichier_service().read_text(encoding="utf-8")
            )
            url = f"http://127.0.0.1:{int(infos['port'])}/api/ping"

            with urllib.request.urlopen(url, timeout=1) as reponse:
                if json.loads(reponse.read()).get("ok"):
                    port = int(infos["port"])

        except Exception:
            port = None

        self._resultat_ping = port
        self._ping_en_cours = False

    def tick(self):
        if not self.service.vivant:
            self.port = None
            self._resultat_ping = None
            self._interface_ouverte = False

        else:
            # 1) lire d'abord le résultat de la vérification précédente
            if self.port is None and self._resultat_ping:
                self.port = self._resultat_ping
                self._resultat_ping = None

            # 2) puis, si besoin, en relancer une
            if self.port is None and not self._ping_en_cours:
                self._ping_en_cours = True
                threading.Thread(
                    target=self._tester_service, daemon=True
                ).start()

        if (
            self.port
            and not self._interface_ouverte
            and self.config.get("ouvrir_navigateur", True)
        ):
            self._interface_ouverte = True
            self.ouvrir_interface()

        if self.signature() != self._signature:
            self.dessiner()

        self.racine.after(700, self.tick)

    def lancer(self):
        self.racine.mainloop()


def _instance_unique():
    """Évite d'ouvrir deux lanceurs (Windows)."""

    if os.name != "nt":
        return None

    import ctypes

    mutex = ctypes.windll.kernel32.CreateMutexW(
        None, False, "Resilia2LauncherMutex"
    )

    if ctypes.windll.kernel32.GetLastError() == 183:
        return False

    return mutex


def run_launcher():
    if os.name == "nt":
        import ctypes

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

    mutex = _instance_unique()

    if mutex is False:
        racine = tk.Tk()
        racine.withdraw()
        messagebox.showinfo(
            "Resilia Launcher", "Resilia Launcher est déjà ouvert."
        )
        racine.destroy()
        return

    Lanceur().lancer()
    del mutex
