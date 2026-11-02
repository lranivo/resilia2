"""
Module VALIDATION : valide le plan de fabrication dans Salto (PEGA) pour
chaque DPS de la colonne DPS/ACT, puis écrit « Validé » dans la colonne
Statut.

Étapes par DPS :
  1. recherche du DPS ;
  2. clic sur « Modifier le plan de fabrication » ;
  3. clic sur « Valider le plan » ;
  4. clic sur « Valider ».
Chaque clic est vérifié (l'étape suivante doit s'afficher) avant de
continuer.

Seules les lignes dont « Etat Résil » vaut OK sont traitées.
"""

import re

from etat_resil import MODE_OK
from isolation_client import chrono
from salto_common import search_dps, executer_sur_excel
from salto_utils import (
    attendre_absence,
    cliquer_avec_verification,
    trouver_partout,
)


# ============================================================
# PARAMÈTRES DE LA VALIDATION
# ============================================================

STATUT_VALIDE = "Validé"

RE_MODIFIER_PLAN = re.compile(
    r"Modifier\s+le\s+plan\s+de\s+fabrication", re.IGNORECASE
)
RE_BOUTON_VALIDER_PLAN = re.compile(
    r"^\s*Valider\s+le\s+plan\s*$", re.IGNORECASE
)
RE_BOUTON_VALIDER = re.compile(r"^\s*Valider\s*$", re.IGNORECASE)


# ============================================================
# LOCALISATION DES ÉLÉMENTS
# ============================================================

# Même comportement que « Traiter l'événement » : lien de la liste dont
# l'action ouvre l'affectation (openAssignment).
BUILDERS_MODIFIER_PLAN = [
    lambda f: f.locator("a[onclick*='openAssignment']").filter(
        has_text=RE_MODIFIER_PLAN
    ),
    lambda f: f.locator("a").filter(has_text=RE_MODIFIER_PLAN),
    lambda f: f.get_by_text(RE_MODIFIER_PLAN),
]

BUILDERS_VALIDER_PLAN = [
    lambda f: f.get_by_role("button", name=RE_BOUTON_VALIDER_PLAN),
    lambda f: f.locator(
        "input[type='button'][value='Valider le plan'], "
        "input[type='submit'][value='Valider le plan']"
    ),
    lambda f: f.locator("button").filter(has_text=RE_BOUTON_VALIDER_PLAN),
]

# Le texte doit être exactement « Valider » (et non « Valider le plan »).
BUILDERS_VALIDER = [
    lambda f: f.get_by_role("button", name=RE_BOUTON_VALIDER),
    lambda f: f.locator(
        "input[type='button'][value='Valider'], "
        "input[type='submit'][value='Valider']"
    ),
    lambda f: f.locator("button").filter(has_text=RE_BOUTON_VALIDER),
]


def trouver(page, builders, timeout_ms, label):
    """Retourne le locator visible (dans n'importe quel onglet / frame)."""

    _, locator = trouver_partout(page, builders, timeout_ms, label)
    return locator


# ============================================================
# SALTO : ÉTAPES DE LA VALIDATION
# ============================================================

def valider_dps(page, dps_id):
    """Enchaîne la recherche puis les trois clics vérifiés."""

    with chrono(f"{dps_id} | recherche"):
        search_dps(page, dps_id)

    # Étape 4 : « Modifier le plan de fabrication » -> « Valider le plan »
    # doit apparaître.
    with chrono(f"{dps_id} | Modifier le plan de fabrication"):
        cliquer_avec_verification(
            page,
            localiser=lambda t: trouver(
                page, BUILDERS_MODIFIER_PLAN, t,
                "Modifier le plan de fabrication",
            ),
            verifier=lambda t: trouver_partout(
                page, BUILDERS_VALIDER_PLAN, t, "Bouton Valider le plan"
            ),
            label="Modifier le plan de fabrication",
            timeout_localisation_ms=90000,
        )

    # Étape 5 : « Valider le plan » -> « Valider » doit apparaître.
    with chrono(f"{dps_id} | Valider le plan"):
        cliquer_avec_verification(
            page,
            localiser=lambda t: trouver(
                page, BUILDERS_VALIDER_PLAN, t, "Bouton Valider le plan"
            ),
            verifier=lambda t: trouver_partout(
                page, BUILDERS_VALIDER, t, "Bouton Valider"
            ),
            label="Valider le plan",
        )

    # Étape 6 : « Valider » -> le bouton doit disparaître (validation
    # prise en compte, comme pour « Non » dans le dégel).
    with chrono(f"{dps_id} | Valider"):
        cliquer_avec_verification(
            page,
            localiser=lambda t: trouver(
                page, BUILDERS_VALIDER, t, "Bouton Valider"
            ),
            verifier=lambda t: attendre_absence(
                page, BUILDERS_VALIDER, t, "Bouton Valider"
            ),
            label="Valider",
            timeout_verification_ms=15000,
        )

    page.wait_for_timeout(500)


def run_validation_salto():
    executer_sur_excel(
        nom="VALIDATION_SALTO",
        action="valider",
        traiter_dps=valider_dps,
        statut_ok=STATUT_VALIDE,
        filtre_etat=MODE_OK,
    )


if __name__ == "__main__":
    run_validation_salto()
