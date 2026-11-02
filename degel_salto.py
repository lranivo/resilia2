"""
Module DÉGEL : dégèle la production dans Salto (PEGA) pour chaque DPS de
la colonne DPS/ACT, puis écrit « Dégelé » dans la colonne Statut.
"""

import re
import time

from isolation_client import chrono
from salto_common import (
    ouvrir_autres_activites,
    search_dps,
    executer_sur_excel,
)
from salto_utils import (
    attendre_page_stable,
    clic_lent,
    ecouteurs_diagnostic,
    trouver_attache,
    trouver_dans_frames,
    trouver_partout,
)


# ============================================================
# PARAMÈTRES DU DÉGEL
# ============================================================

STATUT_DEGELE = "Dégelé"

# Réglages de rapidité (chemin optimiste rapide ; les vérifications et
# les nouvelles tentatives rattrapent les rares échecs).
STABILITE_AVANT_CLIC_MS = 500
PAUSE_SURVOL_MS = 150
ATTENTE_APRES_DISPARITION_MS = 500

RE_DEGELER_PRODUCTION = re.compile(
    r"D[eé]geler\s+(?:une|la)\s+prod\w+", re.IGNORECASE
)
RE_TRAITER_EVENEMENT = re.compile(
    r"Traiter\s+l\W?\s*[eé]v[eé]nement", re.IGNORECASE
)
RE_BOUTON_DEGELER = re.compile(r"^\s*D[eé]geler\s*$", re.IGNORECASE)
RE_BOUTON_NON = re.compile(r"^\s*Non\s*$", re.IGNORECASE)


# ============================================================
# SALTO : ÉTAPES DU DÉGEL
# ============================================================


def cliquer_degeler_production(page):
    """
    Étape 5 : « Dégeler une production », avec vérification.

    Le panneau ne reste affiché que quelques millisecondes et un clic
    peut partir pendant un rafraîchissement PEGA sans rien déclencher.
    Après chaque clic, on vérifie donc que la liste de l'événement
    (lien « Traiter l'événement ») s'affiche ; sinon on recommence :
    tentative 1 = clic souris sur l'élément visible, tentatives 2 à 4 =
    clic JavaScript sur l'élément présent dans le DOM (sans dépendre de
    l'affichage ni du survol).
    """

    builders = [lambda f: f.get_by_text(RE_DEGELER_PRODUCTION)]

    attendre_page_stable(page, STABILITE_AVANT_CLIC_MS)

    derniere_erreur = None

    for tentative in range(1, 5):
        try:
            if tentative == 1:
                _, element = trouver_dans_frames(
                    page, builders, 60000, "Dégeler une production"
                )
                element.click(timeout=5000)
                methode = "clic"

            else:
                _, element = trouver_attache(
                    page, builders, 15000, "Dégeler une production"
                )
                element.evaluate("el => el.click()")
                methode = "clic JavaScript"

            print(
                f"[SALTO] « Dégeler une production » : {methode} "
                f"({tentative}/4)."
            )

            localiser_lien_traiter(page, 8000)
            print("[SALTO] Liste de l'événement affichée.")
            return

        except Exception as error:
            derniere_erreur = error

            # La liste est peut-être apparue malgré l'erreur.
            try:
                localiser_lien_traiter(page, 3000)
                print("[SALTO] Liste de l'événement affichée.")
                return
            except Exception:
                pass

            print(
                f"[SALTO] Tentative {tentative}/4 sans effet "
                f"({type(error).__name__}), nouvelle tentative..."
            )

            attendre_page_stable(page, 500, 4000)

    raise RuntimeError(
        f"« Dégeler une production » n'a pas affiché l'événement : "
        f"{derniere_erreur}"
    )


BUILDERS_BOUTON_DEGELER = [
    lambda f: f.get_by_role("button", name=RE_BOUTON_DEGELER),
    lambda f: f.locator(
        "input[type='button'][value='Dégeler'], "
        "input[type='submit'][value='Dégeler']"
    ),
    lambda f: f.locator("button").filter(has_text=RE_BOUTON_DEGELER),
]


def localiser_lien_traiter(page, timeout_ms):
    """
    Lien « Traiter l'événement » : uniquement celui dont l'action
    openAssignment concerne le gel de production (FREEZEPRODUCTION),
    pour ne pas traiter par erreur un autre événement de la liste.
    """

    _, lien = trouver_dans_frames(
        page,
        [
            lambda f: f.locator(
                "a[onclick*='openAssignment'][onclick*='FREEZEPRODUCTION']"
            ).filter(has_text=RE_TRAITER_EVENEMENT),
        ],
        timeout_ms=timeout_ms,
        label="Traiter l'événement (FREEZEPRODUCTION)",
    )

    return lien


def ouvrir_traitement_evenement(page):
    """
    Étape 6 : ouvre « Traiter l'événement » et retourne le bouton
    « Dégeler » du formulaire ouvert (voir _ouvrir_traitement_evenement).
    """

    with ecouteurs_diagnostic(page) as etat:
        return _ouvrir_traitement_evenement(page, etat)


def _ouvrir_traitement_evenement(page, etat):
    """
    Avant chaque clic : page stable + survol (comme un humain), pour que
    les objets JavaScript de la grille PEGA soient initialisés.
    Si une erreur JS survient pendant le clic (page pas prête), on attend
    la stabilité puis on réessaie la même méthode avant de changer.
    """

    strategies = [
        ("clic", lambda lien: lien.click(timeout=10000)),
        ("clic lent", lambda lien: clic_lent(page, lien)),
        ("clic JavaScript", lambda lien: lien.evaluate("el => el.click()")),
    ]

    for numero, (nom, action) in enumerate(strategies, start=1):
        derniere = numero == len(strategies)

        for essai in range(1, 4):
            try:
                lien = localiser_lien_traiter(
                    page, 60000 if numero == 1 and essai == 1 else 10000
                )

            except RuntimeError:
                # Le lien a disparu : le formulaire est peut-être en
                # cours d'ouverture.
                _, bouton = trouver_partout(
                    page, BUILDERS_BOUTON_DEGELER, 30000, "Bouton Dégeler"
                )
                return bouton

            attendre_page_stable(page, STABILITE_AVANT_CLIC_MS, 5000)

            try:
                lien.scroll_into_view_if_needed()
                lien.hover(timeout=5000)
                page.wait_for_timeout(PAUSE_SURVOL_MS)
            except Exception:
                pass

            etat["erreur_js"] = False
            action(lien)
            page.wait_for_timeout(200)

            suffixe = f" (essai {essai}/3)" if essai > 1 else ""
            print(f"[SALTO] « Traiter l'événement » : {nom}{suffixe}.")

            if etat["erreur_js"]:
                delai = 3000
            else:
                delai = 30000 if derniere else 10000

            try:
                _, bouton = trouver_partout(
                    page,
                    BUILDERS_BOUTON_DEGELER,
                    delai,
                    "Bouton Dégeler",
                )
                return bouton

            except RuntimeError:
                pass

            if etat["erreur_js"]:
                print(
                    "[SALTO] Erreur JS pendant le clic : page pas prête, "
                    "attente puis nouvel essai..."
                )
                attendre_page_stable(page, 1500, 8000)
                continue

            print(
                f"[SALTO] Le formulaire ne s'ouvre pas après le {nom}, "
                "méthode suivante..."
            )
            break

    raise RuntimeError(
        "« Traiter l'événement » cliqué, mais le formulaire "
        "(bouton Dégeler) ne s'ouvre pas."
    )


def cliquer_bouton_degeler(page, bouton):
    """
    Étape 7 : bouton « Dégeler » du formulaire d'affectation.
    Page stable + survol avant le clic (formulaire piloté par le
    JavaScript PEGA, qui s'initialise après l'affichage).
    """

    attendre_page_stable(page, STABILITE_AVANT_CLIC_MS, 5000)

    try:
        bouton.hover(timeout=5000)
        page.wait_for_timeout(PAUSE_SURVOL_MS)
    except Exception:
        pass

    bouton.click(timeout=10000)
    print("[SALTO] Bouton « Dégeler » cliqué.")


BUILDERS_BOUTON_NON = [
    lambda f: f.locator(
        "button.pzbutton[name^='MESPostpone']"
    ).filter(has_text=RE_BOUTON_NON),
    lambda f: f.get_by_role("button", name=RE_BOUTON_NON),
]


def attendre_disparition(page, bouton, timeout_ms):
    """True si le bouton disparaît (ou si son frame est fermé)."""

    deadline = time.monotonic() + timeout_ms / 1000

    while time.monotonic() < deadline:
        try:
            if not bouton.is_visible():
                return True

        except Exception:
            return True

        page.wait_for_timeout(150)

    return False


def cliquer_non_et_verifier(page):
    """
    Étape 8 : bouton « Non », avec vérification qu'il disparaît
    (environ 5 s à la main).

    Le bouton est piloté par le JavaScript PEGA (doFormSubmit) : un clic
    trop précoce ne déclenche rien. Avant le clic on attend la stabilité
    de la page et on survole le bouton. Si le bouton ne disparaît pas
    dans les 15 s (ou si une erreur JS survient), on réessaie, puis on
    change de méthode de clic.
    """

    strategies = [
        ("clic", lambda bouton: bouton.click(timeout=10000)),
        ("clic lent", lambda bouton: clic_lent(page, bouton)),
        ("clic JavaScript", lambda bouton: bouton.evaluate("el => el.click()")),
    ]

    deja_clique = False

    with ecouteurs_diagnostic(page) as etat:
        for numero, (nom, action) in enumerate(strategies, start=1):
            for essai in range(1, 3):
                try:
                    _, bouton = trouver_partout(
                        page,
                        BUILDERS_BOUTON_NON,
                        8000 if deja_clique else 60000,
                        "Bouton Non",
                    )

                except RuntimeError:
                    if deja_clique:
                        # Le bouton a disparu : traitement finalisé.
                        page.wait_for_timeout(ATTENTE_APRES_DISPARITION_MS)
                        return

                    raise

                attendre_page_stable(page, STABILITE_AVANT_CLIC_MS, 5000)

                try:
                    bouton.hover(timeout=5000)
                    page.wait_for_timeout(PAUSE_SURVOL_MS)
                except Exception:
                    pass

                etat["erreur_js"] = False
                action(bouton)
                deja_clique = True
                page.wait_for_timeout(200)

                suffixe = f" (essai {essai}/2)" if essai > 1 else ""
                print(f"[SALTO] Bouton « Non » : {nom}{suffixe}.")

                if attendre_disparition(page, bouton, 15000):
                    page.wait_for_timeout(ATTENTE_APRES_DISPARITION_MS)
                    return

                if etat["erreur_js"] and essai == 1:
                    print(
                        "[SALTO] Erreur JS pendant le clic : page pas "
                        "prête, attente puis nouvel essai..."
                    )
                    attendre_page_stable(page, 1500, 8000)
                    continue

                print(
                    f"[SALTO] Le bouton « Non » est toujours affiché "
                    f"après le {nom}, méthode suivante..."
                )
                break

    raise RuntimeError(
        "Le bouton « Non » est toujours affiché : "
        "le dégel n'a pas été finalisé."
    )


def degeler_dps(page, dps_id):
    """Enchaîne les étapes 3 à 8 pour un DPS."""

    with chrono(f"{dps_id} | recherche"):
        search_dps(page, dps_id)

    with chrono(f"{dps_id} | Autres activités"):
        ouvrir_autres_activites(page)

    with chrono(f"{dps_id} | Dégeler une production"):
        cliquer_degeler_production(page)

    with chrono(f"{dps_id} | Traiter l'événement"):
        bouton_degeler = ouvrir_traitement_evenement(page)

    with chrono(f"{dps_id} | Dégeler"):
        cliquer_bouton_degeler(page, bouton_degeler)

    with chrono(f"{dps_id} | Non"):
        cliquer_non_et_verifier(page)


def run_degel_salto():
    executer_sur_excel(
        nom="DEGEL_SALTO",
        action="dégeler",
        traiter_dps=degeler_dps,
        statut_ok=STATUT_DEGELE,
    )


if __name__ == "__main__":
    run_degel_salto()