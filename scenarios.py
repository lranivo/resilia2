"""
Catalogue des scénarios de Resilia2 et suivi de la progression.

La progression est calculée en lisant les lignes de log que les modules
affichent déjà (aucune modification des modules métier) :

    [FILTRE] 5 ligne(s) ignorée(s) (...)         -> lignes écartées (Etat Résil)
    [EXCEL] 12 DPS à geler.                      -> total (lignes à traiter)
    [EXCEL] Ligne 5 | DPS-123 | Gelé             -> une ligne traitée
    [GEL_SALTO][ERREUR] Ligne 6 | DPS-124 | ...  -> une ligne en erreur

Si le format de ces lignes change dans un module, adapter les expressions
régulières ci-dessous.
"""

import re


# ============================================================
# CATALOGUE (l'ordre est l'ordre d'exécution quand plusieurs
# scénarios sont cochés : colonne de gauche, puis de droite)
# ============================================================

SCENARIOS = [
    {
        "id": "checking",
        "libelle": "Checking",
        "disponible": True,
        "colonnes": ["su", "etat résil", "injoignabilité r1",
                     "injoignabilité r2"],
    },
    {
        "id": "isolation",
        "libelle": "Isolation",
        "disponible": True,
        "colonnes": ["su", "statut", "etat résil"],
    },
    {
        "id": "suppression_feuillet",
        "libelle": "Suppression Feuillet",
        "disponible": False,
        "colonnes": [],
    },
    {
        "id": "checking_dps_nok",
        "libelle": "Checking DPS NOK",
        "disponible": False,
        "colonnes": [],
    },
    {
        "id": "validation_pf",
        "libelle": "Validation PF",
        "disponible": True,
        "colonnes": ["dps/act", "statut", "etat résil"],
    },
    {
        "id": "gel",
        "libelle": "Gel de gamme",
        "disponible": True,
        "colonnes": ["dps/act", "statut", "etat résil"],
    },
    {
        "id": "degel",
        "libelle": "Dégel",
        "disponible": True,
        "colonnes": ["dps/act", "statut"],
    },
]

PAR_ID = {s["id"]: s for s in SCENARIOS}


def charger_fonction(id_scenario):
    """
    Imports explicites (et non dynamiques) pour que PyInstaller embarque
    chaque module dans l'exécutable.
    """

    if id_scenario == "checking":
        from check_wasac import run_check_wasac_excel
        return run_check_wasac_excel

    if id_scenario == "isolation":
        from isolation_client import run_isolation_client
        return run_isolation_client

    if id_scenario == "validation_pf":
        from validation_salto import run_validation_salto
        return run_validation_salto

    if id_scenario == "gel":
        from gel_salto import run_gel_salto
        return run_gel_salto

    if id_scenario == "degel":
        from degel_salto import run_degel_salto
        return run_degel_salto

    raise ValueError(f"Scénario indisponible : {id_scenario}")


def colonnes_manquantes(id_scenario, en_tetes):
    """Colonnes requises par le scénario et absentes du fichier."""

    presentes = {str(e).strip().casefold() for e in en_tetes}

    return [
        c for c in PAR_ID[id_scenario]["colonnes"] if c not in presentes
    ]


# ============================================================
# SUIVI DE PROGRESSION (lecture des logs)
# ============================================================

RE_TOTAL = re.compile(
    r"^\[EXCEL\]\s+(\d+)\s+(SU|DPS|réseau)", re.IGNORECASE
)
RE_LIGNE_TRAITEE = re.compile(r"^\[EXCEL\]\s+Ligne\s+\d+\s+\|")
RE_LIGNE_ERREUR = re.compile(r"\[ERREUR\]\s+Ligne\s+\d+")

# Lignes écartées par le filtre « Etat Résil » (annoncées avant le total).
RE_IGNOREES = re.compile(r"^\[FILTRE\]\s+(\d+)\s+ligne", re.IGNORECASE)

UNITES = {"su": "SU", "dps": "DPS", "réseau": "réseaux"}


class SuiviProgression:
    def __init__(self):
        self.total = None
        self.fait = 0
        self.erreurs = 0
        self.ignores = 0
        self.unite = "lignes"

    def lire(self, ligne):
        """Met à jour les compteurs d'après une ligne de log."""

        ligne = ligne.strip()

        correspondance_ignorees = RE_IGNOREES.match(ligne)

        if correspondance_ignorees:
            self.ignores = int(correspondance_ignorees.group(1))
            return

        # Tant que le total n'est pas annoncé, les lignes « ignorée » ou
        # « déjà … » ne comptent pas.
        if self.total is None:
            correspondance = RE_TOTAL.match(ligne)

            if correspondance:
                self.total = int(correspondance.group(1))
                self.unite = UNITES.get(
                    correspondance.group(2).casefold(), "lignes"
                )

            return

        if RE_LIGNE_ERREUR.search(ligne):
            self.fait += 1
            self.erreurs += 1

        elif RE_LIGNE_TRAITEE.match(ligne) and "ignor" not in ligne:
            self.fait += 1

    def fraction(self, termine=False):
        if termine:
            return 1.0

        if not self.total:
            return 0.0

        return min(self.fait / self.total, 1.0)
