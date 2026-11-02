"""
Module GEL : gèle la production dans Salto (PEGA) pour chaque DPS de la
colonne DPS/ACT, puis écrit « Gelé » dans la colonne Statut.

Seules les lignes à isoler (« Etat Résil » = NOK) sont traitées.
"""

import re
import time
from datetime import datetime, timedelta

from etat_resil import MODE_A_ISOLER
from isolation_client import chrono
from salto_common import (
    ouvrir_autres_activites,
    search_dps,
    executer_sur_excel,
)
from salto_utils import (
    attendre_dans_frame,
    attendre_page_stable,
    ecouteurs_diagnostic,
    trouver_attache,
    trouver_dans_frames,
)


# ============================================================
# PARAMÈTRES DU GEL
# ============================================================

DUREE_GEL_JOURS = 7
ORIGINE = "Interne-Production"
CAUSE = "En attente validation config matériels"
STATUT_GELE = "Gelé"

RE_AJOUTER_TRAVAIL = re.compile(r"Ajouter\s+un\s+travail", re.IGNORECASE)
RE_GELER_PRODUCTION = re.compile(r"Geler\s+la\s+prod\w+", re.IGNORECASE)

# Réglages de rapidité (chemin optimiste rapide ; les vérifications et
# les nouvelles tentatives rattrapent les rares échecs).
STABILITE_AVANT_CLIC_MS = 500
PAUSE_SURVOL_MS = 150
ATTENTE_APRES_FERMETURE_MS = 500


# ============================================================
# SALTO : ÉTAPES DU GEL
# ============================================================


def formulaire_gel_ouvert(page, timeout_ms):
    """True si le champ « date de suivi » du formulaire est affiché."""

    try:
        trouver_dans_frames(
            page,
            [lambda f: f.locator("input[name$='pFollowUpDateTime']")],
            timeout_ms=timeout_ms,
            label="Champ date de suivi",
        )
        return True

    except RuntimeError:
        return False


def declencher_geler(page, methode, premier):
    """
    Survole « Ajouter un travail » puis déclenche « Geler la production ».
    Méthodes : « clic », « clic lent », « clic JavaScript » (ce dernier
    ne dépend pas de l'affichage du menu, qui se referme très vite).
    """

    _, ajouter = trouver_dans_frames(
        page,
        [lambda f: f.get_by_text(RE_AJOUTER_TRAVAIL)],
        timeout_ms=30000 if premier else 10000,
        label="Ajouter un travail",
    )

    ajouter.hover(timeout=5000)
    page.wait_for_timeout(PAUSE_SURVOL_MS)

    if methode == "clic JavaScript":
        _, geler = trouver_attache(
            page,
            [lambda f: f.get_by_text(RE_GELER_PRODUCTION)],
            timeout_ms=5000,
            label="Geler la production",
        )
        geler.evaluate("el => el.click()")
        return

    try:
        _, geler = trouver_dans_frames(
            page,
            [lambda f: f.get_by_text(RE_GELER_PRODUCTION)],
            timeout_ms=4000,
            label="Geler la production (survol)",
        )

    except RuntimeError:
        # Certains menus PEGA s'ouvrent au clic plutôt qu'au survol.
        ajouter.click(timeout=5000)

        _, geler = trouver_dans_frames(
            page,
            [lambda f: f.get_by_text(RE_GELER_PRODUCTION)],
            timeout_ms=8000,
            label="Geler la production (clic)",
        )

    geler.hover(timeout=5000)

    if methode == "clic lent":
        page.wait_for_timeout(300)
        geler.click(
            position={"x": 8, "y": 6}, delay=150, timeout=5000
        )
    else:
        geler.click(timeout=5000)


def cliquer_geler_production(page):
    """
    Étape 5 : « Geler la production », avec vérification.

    Le menu se referme très vite et un clic peut partir avant que le
    JavaScript PEGA soit prêt, sans rien déclencher. Après chaque clic,
    on vérifie donc que le formulaire s'ouvre (champ date visible) ;
    sinon on réessaie, puis on change de méthode (clic, clic lent,
    clic JavaScript).
    """

    methodes = ("clic", "clic lent", "clic JavaScript")
    premier = True

    with ecouteurs_diagnostic(page) as etat:
        attendre_page_stable(page, STABILITE_AVANT_CLIC_MS, 8000)

        for numero, methode in enumerate(methodes, start=1):
            # Le formulaire s'est peut-être ouvert avec retard.
            if numero > 1 and formulaire_gel_ouvert(page, 500):
                print("[SALTO] Formulaire de gel affiché.")
                return

            for essai in range(1, 3):
                etat["erreur_js"] = False

                try:
                    declencher_geler(page, methode, premier)
                    premier = False

                except Exception as error:
                    premier = False
                    resume = (
                        str(error).splitlines()[0] if str(error) else ""
                    )
                    print(
                        f"[SALTO] « Geler la production » : {methode} "
                        f"en échec ({type(error).__name__}: {resume})"
                    )
                    attendre_page_stable(page, 500, 4000)
                    break

                page.wait_for_timeout(200)

                suffixe = f" (essai {essai}/2)" if essai > 1 else ""
                print(
                    f"[SALTO] « Geler la production » : "
                    f"{methode}{suffixe}."
                )

                delai = 3000 if etat["erreur_js"] else 10000

                if formulaire_gel_ouvert(page, delai):
                    print("[SALTO] Formulaire de gel affiché.")
                    return

                if etat["erreur_js"] and essai == 1:
                    print(
                        "[SALTO] Erreur JS pendant le clic : page pas "
                        "prête, attente puis nouvel essai..."
                    )
                    attendre_page_stable(page, 1500, 8000)
                    continue

                print(
                    f"[SALTO] Le formulaire ne s'ouvre pas après le "
                    f"{methode}, méthode suivante..."
                )
                break

    raise RuntimeError(
        "« Geler la production » cliqué, mais le formulaire de gel "
        "ne s'ouvre pas."
    )


def remplir_formulaire_gel(page, date_suivi, commentaire):
    """Étapes 6 à 9 : date, origine, cause, commentaire."""

    # Étape 6 : date de suivi (J+7)
    frame, champ_date = trouver_dans_frames(
        page,
        [lambda f: f.locator("input[name$='pFollowUpDateTime']")],
        timeout_ms=30000,
        label="Champ date de suivi",
    )

    champ_date.click(timeout=10000)
    champ_date.fill(date_suivi)
    champ_date.press("Tab")
    print(f"[SALTO] Date de suivi : {date_suivi}")

    # Étape 7 : origine
    origine = attendre_dans_frame(
        page,
        frame,
        [lambda f: f.locator("#SelectOrigin")],
        timeout_ms=30000,
        label="Liste Origine",
    )

    origine.select_option(value=ORIGINE)
    print(f"[SALTO] Origine : {ORIGINE}")

    # Étape 8 : cause (liste dynamique, chargée après le choix d'origine)
    cause = attendre_dans_frame(
        page,
        frame,
        [lambda f: f.locator("#SelectCause")],
        timeout_ms=30000,
        label="Liste Cause",
    )

    deadline = time.monotonic() + 20

    while time.monotonic() < deadline:
        option = cause.locator(f"option[value='{CAUSE}']")

        if option.count() > 0:
            break

        page.wait_for_timeout(200)

    else:
        raise RuntimeError(
            f"Option de cause introuvable après chargement : {CAUSE}"
        )

    cause.select_option(value=CAUSE)
    print(f"[SALTO] Cause : {CAUSE}")

    # Étape 9 : commentaire
    commentaire_champ = attendre_dans_frame(
        page,
        frame,
        [lambda f: f.locator("textarea[name$='ppyDescription']")],
        timeout_ms=30000,
        label="Champ commentaire",
    )

    commentaire_champ.click(timeout=10000)
    commentaire_champ.fill(commentaire)
    commentaire_champ.press("Tab")
    print("[SALTO] Commentaire renseigné.")

    return frame, champ_date


BUILDERS_VALIDER = [
    lambda f: f.get_by_role(
        "button",
        name=re.compile(r"^\s*Valider\s*$", re.IGNORECASE),
    ),
    lambda f: f.locator(
        "input[type='button'][value='Valider'], "
        "input[type='submit'][value='Valider']"
    ),
    lambda f: f.get_by_text("Valider", exact=True),
]


def formulaire_ferme(champ_date):
    try:
        return not champ_date.is_visible()
    except Exception:
        # Élément ou frame disparu : formulaire fermé.
        return True


def attendre_fermeture(page, champ_date, timeout_ms):
    deadline = time.monotonic() + timeout_ms / 1000

    while time.monotonic() < deadline:
        if formulaire_ferme(champ_date):
            return True

        page.wait_for_timeout(150)

    return False


def valider_gel(page, frame, champ_date):
    """
    Étape 10 : clic sur « Valider », puis vérification que le formulaire
    est bien fermé. Page stable + survol avant le clic ; en cas
    d'échec, nouvelle tentative par clic JavaScript.
    """

    with ecouteurs_diagnostic(page) as etat:
        for essai, methode in enumerate(("clic", "clic JavaScript"), 1):
            if essai > 1 and formulaire_ferme(champ_date):
                break

            try:
                bouton = attendre_dans_frame(
                    page,
                    frame,
                    BUILDERS_VALIDER,
                    timeout_ms=30000 if essai == 1 else 8000,
                    label="Bouton Valider",
                )

            except RuntimeError:
                if essai > 1 or formulaire_ferme(champ_date):
                    break

                raise

            attendre_page_stable(page, STABILITE_AVANT_CLIC_MS, 5000)

            try:
                bouton.hover(timeout=5000)
                page.wait_for_timeout(PAUSE_SURVOL_MS)
            except Exception:
                pass

            etat["erreur_js"] = False

            if methode == "clic":
                bouton.click(timeout=10000)
            else:
                bouton.evaluate("el => el.click()")

            print(f"[SALTO] « Valider » : {methode}.")

            if attendre_fermeture(
                page,
                champ_date,
                3000 if etat["erreur_js"] else 15000,
            ):
                break

            print(
                "[SALTO] Le formulaire est toujours affiché, "
                "nouvelle tentative..."
            )

        else:
            raise RuntimeError(
                "Le formulaire de gel est toujours affiché après "
                "« Valider » (erreur de validation PEGA ?)."
            )

        if not formulaire_ferme(champ_date):
            raise RuntimeError(
                "Le formulaire de gel est toujours affiché après "
                "« Valider » (erreur de validation PEGA ?)."
            )

        page.wait_for_timeout(ATTENTE_APRES_FERMETURE_MS)


def geler_dps(page, dps_id):
    """Enchaîne les étapes 3 à 10 pour un DPS."""

    aujourdhui = datetime.now()
    date_jour = aujourdhui.strftime("%d/%m/%Y")
    date_suivi = (
        aujourdhui + timedelta(days=DUREE_GEL_JOURS)
    ).strftime("%d/%m/%Y")

    commentaire = (
        f"Isolation client pour {DUREE_GEL_JOURS} jours "
        f"suite demande de résiliation du {date_jour}"
    )

    with chrono(f"{dps_id} | recherche"):
        search_dps(page, dps_id)

    with chrono(f"{dps_id} | Autres activités"):
        ouvrir_autres_activites(page)

    with chrono(f"{dps_id} | Geler la production"):
        cliquer_geler_production(page)

    with chrono(f"{dps_id} | formulaire"):
        frame, champ_date = remplir_formulaire_gel(
            page, date_suivi, commentaire
        )

    with chrono(f"{dps_id} | validation"):
        valider_gel(page, frame, champ_date)


def run_gel_salto():
    executer_sur_excel(
        nom="GEL_SALTO",
        action="geler",
        traiter_dps=geler_dps,
        statut_ok=STATUT_GELE,
        filtre_etat=MODE_A_ISOLER,
    )


if __name__ == "__main__":
    run_gel_salto()