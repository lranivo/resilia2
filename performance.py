"""
Performance Mode : garde le poste utilisable pendant les longs traitements.

Deux modes (config.json -> "performance" -> "mode") :

  "natif" (défaut)
      * demande à Windows de ne pas se mettre en veille et de ne pas
        éteindre l'écran (SetThreadExecutionState, la même API que les
        lecteurs vidéo ; aucun droit administrateur) ;
      * option « simuler_activite_teams » : appui périodique sur la touche
        F15 (sans effet dans les applications) pour que Windows et Teams
        voient une activité et ne passent pas en « Absent ».

  "vlc"
      * lance VLC sans interface sur un fichier audio silencieux en
        boucle ; VLC empêche alors la mise en veille tant qu'il lit.
        Ne maintient pas le statut Teams.

Remarque : l'activité simulée réinitialise aussi les minuteries
d'inactivité (verrouillage automatique inclus). Ne laissez pas le poste
sans surveillance dans un lieu non sécurisé, et vérifiez que cet usage
est compatible avec la politique informatique de la société.
"""

import os
import subprocess
import sys
import time
import wave

from app_paths import charger_config, dossier_donnees, dossier_logs

ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002
VK_F15 = 0x7E
KEYEVENTF_KEYUP = 0x0002
CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0


def _garder_eveille():
    if os.name != "nt":
        return

    import ctypes

    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
    )


def _battement_activite():
    if os.name != "nt":
        return

    import ctypes

    user32 = ctypes.windll.user32
    user32.keybd_event(VK_F15, 0, 0, 0)
    user32.keybd_event(VK_F15, 0, KEYEVENTF_KEYUP, 0)


def _trouver_vlc(chemin_config):
    candidats = [chemin_config] if chemin_config else []

    for variable in ("ProgramFiles", "ProgramFiles(x86)"):
        racine = os.environ.get(variable)

        if racine:
            candidats.append(
                os.path.join(racine, "VideoLAN", "VLC", "vlc.exe")
            )

    for chemin in candidats:
        if chemin and os.path.isfile(chemin):
            return chemin

    return None


def _creer_silence():
    chemin = dossier_donnees() / "silence.wav"

    if not chemin.exists():
        with wave.open(str(chemin), "wb") as fichier:
            fichier.setnchannels(1)
            fichier.setsampwidth(2)
            fichier.setframerate(8000)
            fichier.writeframes(b"\x00\x00" * 8000 * 5)

    return chemin


def _mode_vlc(configuration):
    vlc = _trouver_vlc(configuration.get("chemin_vlc", ""))

    if not vlc:
        print("[PERFORMANCE] VLC introuvable : bascule en mode natif.")
        return False

    silence = _creer_silence()
    print(f"[PERFORMANCE] VLC : {vlc}")

    processus = subprocess.Popen(
        [vlc, "--intf", "dummy", "--quiet", "--loop", str(silence)],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=CREATE_NO_WINDOW,
    )

    try:
        processus.wait()
    finally:
        if processus.poll() is None:
            processus.terminate()

    return True


def run_performance():
    sys.stdout = sys.stderr = open(
        dossier_logs() / "performance.log", "a", encoding="utf-8",
        buffering=1,
    )

    configuration = charger_config()["performance"]
    mode = str(configuration.get("mode", "natif")).lower()

    print(f"[PERFORMANCE] Démarrage (mode {mode}).")

    if mode == "vlc" and _mode_vlc(configuration):
        return

    intervalle = max(10, int(configuration.get("intervalle_secondes", 60)))
    simuler = bool(configuration.get("simuler_activite_teams", True))

    print(
        f"[PERFORMANCE] Veille bloquée ; activité simulée : "
        f"{'oui' if simuler else 'non'} (toutes les {intervalle}s)."
    )

    while True:
        _garder_eveille()

        if simuler:
            _battement_activite()

        time.sleep(intervalle)
