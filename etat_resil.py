"""
Filtrage des lignes Excel selon la colonne « Etat Résil » (remplie par le
scénario Checking) :

  - Validation PF            -> uniquement les lignes « OK » ;
  - Isolation / Gel de gamme -> uniquement les lignes à isoler (« NOK »).

Un seul classeur suffit donc : plus besoin de séparer les OK et les NOK.

Les lignes écartées sont journalisées AVANT l'annonce du total
(« [EXCEL] N ... à traiter »), pour que la barre de progression ne compte
que les lignes réellement traitées. La ligne « [FILTRE] N ligne(s)
ignorée(s) » est lue par scenarios.py pour l'affichage dans l'interface.
"""

ETAT_OK = "OK"

# Etats considérés comme « à isoler / à geler ».
# Par prudence, seules les lignes « NOK » sont isolées : une ligne
# « ERREUR » (checking échoué) ou vide (checking pas encore fait) n'est
# jamais isolée par erreur. Pour les inclure, ajouter "ERREUR" et/ou ""
# dans cet ensemble.
ETATS_A_ISOLER = {"NOK"}

MODE_OK = "ok"          # Validation PF
MODE_A_ISOLER = "nok"   # Isolation, Gel de gamme


def normaliser_etat(valeur):
    """'  nok ' -> 'NOK' ; cellule vide -> ''."""

    if valeur is None:
        return ""

    return str(valeur).strip().upper()


def doit_traiter(etat, mode):
    """True si une ligne avec cet Etat Résil est concernée par le mode."""

    etat = normaliser_etat(etat)

    if mode == MODE_OK:
        return etat == ETAT_OK

    if mode == MODE_A_ISOLER:
        return etat in ETATS_A_ISOLER

    raise ValueError(f"Mode de filtrage inconnu : {mode!r}")


def libelle_mode(mode):
    return "Etat Résil = OK" if mode == MODE_OK else "Etat Résil = NOK"


def annoncer_ignorees(ignorees, mode):
    """
    ignorees : liste de (ligne_excel, reference, etat_normalise).
    Affiche le détail puis la ligne de synthèse lue par l'interface.
    """

    if not ignorees:
        return

    for row, reference, etat in ignorees:
        print(
            f"[EXCEL] Ligne {row} | {reference} ignoré "
            f"(Etat Résil : {etat or 'vide'})."
        )

    repartition = {}

    for _, _, etat in ignorees:
        cle = etat or "vide"
        repartition[cle] = repartition.get(cle, 0) + 1

    detail = ", ".join(
        f"{cle} : {nombre}" for cle, nombre in sorted(repartition.items())
    )

    print(
        f"[FILTRE] {len(ignorees)} ligne(s) ignorée(s) "
        f"({libelle_mode(mode)} requis ; trouvé -> {detail})."
    )
