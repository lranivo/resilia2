import re
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import pythoncom
import win32com.client

from playwright.sync_api import sync_playwright

from browser import launch_browser, close_browser
from etat_resil import (
    MODE_A_ISOLER,
    annoncer_ignorees,
    doit_traiter,
    normaliser_etat,
)
from check_wasac import (
    open_wasac,
    find_reseau_field,
    rechercher_acquitter,
    click_text,
    MAX_REFERENCES,
    texte_disponible,
)


EXCEL_FILE = Path(__file__).resolve().parent / "resil.xlsx"

# Attente après "Confirmer ce traitement" (laisser la requête partir
# avant la navigation suivante). À réduire avec prudence.
ATTENTE_APRES_CONFIRMATION_MS = 3000


# ============================================================
# SCRIPT D'ISOLATION
# ============================================================

def generer_script_isolation():
    date_jour = datetime.now().strftime("%d/%m/%Y")

    return f"""conf t
interface GigabitEthernet 0/0
shut
description shut LAN suite demande de resil du {date_jour}
exit
interface GigabitEthernet 0/1
shut
description shut LAN suite demande de resil du {date_jour}
exit
interface GigabitEthernet 0/2
shut
description shut LAN suite demande de resil du {date_jour}
exit
interface GigabitEthernet 0/3
shut
description shut LAN suite demande de resil du {date_jour}
exit
int dot11radio 0/0
shut
description shut LAN suite demande de resil du {date_jour}
exit
int dot11radio 1/0
shut
description shut LAN suite demande de resil du {date_jour}
end

write mem

!
"""


@contextmanager
def chrono(label):
    """Affiche le temps passé dans une étape."""

    debut = time.monotonic()
    try:
        yield
    finally:
        print(f"[TEMPS] {label} : {time.monotonic() - debut:.1f}s")


def attendre_resultat_recherche(page, max_ms=3000, marge_ms=500):
    """
    Remplace l'ancienne pause fixe de 3 s après la recherche.
    Dès que le menu du feuillet est chargé (feuillet=<ref> dans l'URL
    d'un frame), on continue après une courte marge. Dans le pire cas,
    l'attente est la même qu'avant (max_ms).
    """

    deadline = time.monotonic() + max_ms / 1000

    while time.monotonic() < deadline:
        for frame in list(page.frames):
            try:
                if re.search(r"[?&]feuillet=[^&]+", frame.url):
                    page.wait_for_timeout(marge_ms)
                    return
            except Exception:
                continue

        page.wait_for_timeout(100)


# ============================================================
# RECHERCHE WASAC
# ============================================================

def rechercher_su(page, su):
    print(f"[WASAC] Recherche de {su}...")

    with chrono("    ouverture WASAC"):
        open_wasac(page)

    with chrono("    saisie du réseau"):
        champ = find_reseau_field(page, timeout=90000)
        champ.fill(su)
        champ.press("Enter")

    with chrono("    attente du résultat"):
        attendre_resultat_recherche(page)
        rechercher_acquitter(page)

    print(f"[WASAC] SU {su} affiché.")


# ============================================================
# FORMULAIRE DE TRAITEMENT CLI
# ============================================================

def selectionner_traitement_cli(page):
    deadline = time.monotonic() + 15
    frame_traitement = None

    while time.monotonic() < deadline:
        for frame in page.frames:
            try:
                if frame.locator("#traitement_list_chosen").count() > 0:
                    frame_traitement = frame
                    break
            except Exception:
                continue

        if frame_traitement is not None:
            break

        page.wait_for_timeout(200)

    if frame_traitement is None:
        raise RuntimeError("Formulaire de traitement introuvable.")

    menu = frame_traitement.locator(
        "#traitement_list_chosen a.chosen-single"
    )
    menu.wait_for(state="visible", timeout=10000)
    menu.click()

    option_cli = frame_traitement.locator(
        "#traitement_list_chosen .chosen-results li"
    ).filter(has_text="CLI (Exécution de commandes en CLI)")

    option_cli.wait_for(state="visible", timeout=10000)
    option_cli.click()

    bouton_valider = frame_traitement.locator("#valider_traitement")
    bouton_valider.wait_for(state="visible", timeout=10000)
    bouton_valider.click()


def trouver_textarea_script(page, timeout=20000):
    deadline = time.monotonic() + timeout / 1000

    while time.monotonic() < deadline:
        for frame in page.frames:
            try:
                champ = frame.locator("#script_code_script").first

                if champ.count() > 0 and champ.is_visible():
                    return champ
            except Exception:
                continue

        page.wait_for_timeout(200)

    raise RuntimeError("Champ de script CLI introuvable.")


# ============================================================
# ISOLATION D'UN SU À FEUILLET SIMPLE
# ============================================================

def appliquer_traitement_isolation(page, su):
    print(f"[ISOLATION] Traitement du feuillet {su}...")

    with chrono("    ouverture formulaire CLI"):
        click_text(page, "Traitement", timeout=10000)
        page.wait_for_timeout(500)
        selectionner_traitement_cli(page)

    with chrono("    saisie du script"):
        champ_script = trouver_textarea_script(page)
        champ_script.fill(generer_script_isolation())

    with chrono("    confirmation"):
        click_text(
            page,
            "Confirmer ce traitement",
            timeout=15000,
        )

    print(f"[ISOLATION] Confirmation déclenchée pour {su}.")

    page.wait_for_timeout(ATTENTE_APRES_CONFIRMATION_MS)
    return True


# Éléments du menu WASAC qui prouvent que la page est chargée.
INDICATEURS_MENU = (
    "Exploitation",
    re.compile(r"Synth[eè]se\s+des\s+informations", re.IGNORECASE),
)


def texte_visible(page, texte):
    """True si le texte est visible dans la page ou l'un de ses frames."""

    for frame in list(page.frames):
        try:
            locator = frame.get_by_text(texte, exact=True).first

            if locator.count() > 0 and locator.is_visible():
                return True

        except Exception:
            continue

    return False


def signature_pages(page):
    """Empreinte légère : URL + taille du texte de chaque frame."""

    parts = []

    for frame in list(page.frames):
        try:
            taille = frame.evaluate(
                "document.body ? document.body.innerText.length : 0"
            )
        except Exception:
            taille = -1

        parts.append((frame.url, taille))

    return tuple(parts)


def traitement_disponible(page, timeout=12000, stabilite=1000):
    """
    Attente adaptative du bouton Traitement.

    - Traitement visible                      -> True immédiatement.
    - Menu chargé (Exploitation / Synthèse)
      ET contenu inchangé pendant `stabilite` ms
      sans Traitement                         -> False rapidement
                                                 (absence confirmée).
    - Page pas encore chargée                 -> on continue d'attendre
                                                 jusqu'à `timeout` ms.

    Aucune pause fixe : le cas normal ne perd pas de temps.
    """

    deadline = time.monotonic() + timeout / 1000
    derniere_signature = None
    stable_depuis = None

    while time.monotonic() < deadline:
        if texte_visible(page, "Traitement"):
            print("[DEBUG] Traitement trouvé.")
            return True

        if any(texte_visible(page, t) for t in INDICATEURS_MENU):
            signature = signature_pages(page)
            maintenant = time.monotonic()

            if signature != derniere_signature:
                derniere_signature = signature
                stable_depuis = maintenant

            elif maintenant - stable_depuis >= stabilite / 1000:
                print(
                    "[DEBUG] Traitement absent : page chargée "
                    "et stable."
                )
                return False

        page.wait_for_timeout(200)

    print("[DEBUG] Traitement absent : délai dépassé.")
    return False


# ============================================================
# ISOLATION D'UN SU À RÉFÉRENCES MULTIPLES
# ============================================================

DIAGNOSTIC_DIR = Path(__file__).resolve().parent / "diagnostics"


def sauvegarder_diagnostic(page, reference):
    """
    Sauvegarde une capture d'écran, le texte visible et la liste des
    frames quand le bouton Traitement est introuvable.
    """

    try:
        DIAGNOSTIC_DIR.mkdir(exist_ok=True)
        horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = DIAGNOSTIC_DIR / f"{reference}_{horodatage}"

        try:
            page.screenshot(path=str(base) + ".png")
        except Exception as error:
            print(f"[DIAG] Capture impossible : {error}")

        lignes = ["FRAMES :"]
        lignes += [f"  - {frame.url}" for frame in page.frames]
        lignes += ["", "TEXTE VISIBLE :", texte_disponible(page)]

        Path(str(base) + ".txt").write_text(
            "\n".join(lignes), encoding="utf-8"
        )

        print(f"[DIAG] Diagnostic enregistré : {base}.png / .txt")

    except Exception as error:
        print(f"[DIAG] Échec du diagnostic : {error}")


def traiter_references_multiples(page, su, references):
    """
    Isole chaque référence extraite, une par une, dans un onglet dédié.
    Retourne le statut à écrire sur la ligne du SU mère.
    """

    reussies = []
    echecs = []

    for numero, reference in enumerate(references, start=1):
        print(f"\n[ISOLATION] {su} -> R{numero} : {reference}")

        debut_ref = time.monotonic()
        page_ref = page.context.new_page()

        try:
            rechercher_su(page_ref, reference)

            if not traitement_disponible(page_ref, timeout=10000):
                print(
                    f"[ISOLATION] Traitement absent pour {reference}, "
                    "nouvelle tentative..."
                )
                rechercher_su(page_ref, reference)

                if not traitement_disponible(page_ref, timeout=10000):
                    sauvegarder_diagnostic(page_ref, reference)
                    raise RuntimeError(
                        f"Bouton Traitement introuvable pour {reference}."
                    )

            appliquer_traitement_isolation(page_ref, reference)
            reussies.append(reference)

        except Exception as error:
            print(
                f"[ISOLATION][ERREUR] {reference} : "
                f"{type(error).__name__}: {error}"
            )
            echecs.append(reference)

        finally:
            print(
                f"[TEMPS] {reference} | TOTAL référence : "
                f"{time.monotonic() - debut_ref:.1f}s"
            )
            try:
                page_ref.close()
            except Exception:
                pass

    total = len(references)

    print(
        f"[ISOLATION] {su} | OK={reussies} | KO={echecs}"
    )

    if len(reussies) == total:
        return "Isolé"

    if reussies:
        return (
            f"Partiel ({len(reussies)}/{total}) "
            f"KO: {', '.join(echecs)}"
        )

    return "ERREUR"


# ============================================================
# EXTRACTION RAPIDE DES RÉFÉRENCES (SU À RÉFÉRENCES MULTIPLES)
# ============================================================

REGEX_SYNTHESE = re.compile(
    r"Synth[eè]se\s+des\s+informations", re.IGNORECASE
)
REGEX_ETAT = re.compile(r"RESILIATION\s+EN\s+COURS", re.IGNORECASE)
REGEX_REFERENCE = re.compile(r"\bCD[0O][A-Z0-9]{7}\b", re.IGNORECASE)


def lire_contenus_resiliation(page, timeout=15000):
    """
    Lit le texte de chaque frame avec evaluate() (échec immédiat si le
    frame est vide ou détaché, contrairement à inner_text qui attend 5 s).
    Retourne dès que RESILIATION EN COURS apparaît ET que le contenu
    est stable sur deux lectures espacées de 500 ms.
    """

    deadline = time.monotonic() + timeout / 1000
    precedent = None

    while time.monotonic() < deadline:
        contenus = []

        for frame in list(page.frames):
            try:
                contenu = frame.evaluate(
                    "document.body ? document.body.innerText : ''"
                )
            except Exception:
                continue

            if contenu and REGEX_ETAT.search(contenu):
                contenus.append(contenu)

        if contenus:
            if contenus == precedent:
                return contenus

            precedent = contenus
            page.wait_for_timeout(500)
            continue

        page.wait_for_timeout(250)

    return []


def extraire_references_rapide(page):
    """
    Même logique que extraire_references_multiples (check_wasac),
    sans pause fixe de 2 s ni inner_text à 5 s par frame vide.
    """

    print("[WASAC] Recherche de Synthèse des informations...")

    with chrono("    clic Synthèse des informations"):
        click_text(page, REGEX_SYNTHESE, timeout=30000)

    with chrono("    lecture des frames"):
        contenus = lire_contenus_resiliation(page)

    if not contenus:
        raise RuntimeError(
            "Impossible de lire les routeurs dans "
            "Synthèse des informations."
        )

    references = []

    for contenu in contenus:
        lignes = [
            ligne.strip()
            for ligne in contenu.splitlines()
            if ligne.strip()
        ]

        for index, ligne in enumerate(lignes):
            if not REGEX_ETAT.search(ligne):
                continue

            zone = " ".join(lignes[max(0, index - 3): index + 1])
            matches = REGEX_REFERENCE.findall(zone)

            if not matches:
                continue

            reference = matches[-1].upper()

            if reference.startswith("CDO"):
                reference = "CD0" + reference[3:]

            if reference not in references:
                references.append(reference)

    if not references:
        raise RuntimeError(
            "Aucune référence avec l'état RESILIATION EN COURS "
            "n'a été trouvée."
        )

    if len(references) > MAX_REFERENCES:
        print(
            f"[WASAC][ATTENTION] {len(references)} références "
            f"détectées. Seules les {MAX_REFERENCES} premières "
            f"seront traitées."
        )

    references = references[:MAX_REFERENCES]

    print(f"[WASAC] Références à vérifier : {references}")

    return references


def traiter_su_isolation(page, su):
    """
    Retourne le statut à écrire dans Excel pour le SU mère.
    - Bouton Traitement présent : SU à feuillet simple.
    - Bouton Traitement absent  : SU à références multiples.
    """

    print(f"\n[ISOLATION] Début du SU {su}")
    debut_su = time.monotonic()

    with chrono(f"{su} | recherche WASAC"):
        rechercher_su(page, su)

    with chrono(f"{su} | détection Traitement"):
        present = traitement_disponible(page)

    if present:
        with chrono(f"{su} | traitement CLI"):
            appliquer_traitement_isolation(page, su)

        print(f"[TEMPS] {su} | TOTAL : {time.monotonic() - debut_su:.1f}s")
        return "Isolé"

    print(
        f"[ISOLATION] Bouton Traitement absent pour {su} : "
        "SU à références multiples."
    )

    with chrono(f"{su} | extraction des références"):
        references = extraire_references_rapide(page)

    statut = traiter_references_multiples(page, su, references)

    print(f"[TEMPS] {su} | TOTAL : {time.monotonic() - debut_su:.1f}s")
    return statut


# ============================================================
# EXCEL
# ============================================================

def trouver_colonnes(ws):
    entetes = {}
    derniere_colonne = ws.UsedRange.Columns.Count

    for col in range(1, derniere_colonne + 1):
        valeur = ws.Cells(1, col).Value

        if valeur is not None:
            entetes[str(valeur).strip().casefold()] = col

    colonnes = {}

    for cle in ("su", "statut", "etat résil"):
        if cle not in entetes:
            raise RuntimeError(f"Colonne Excel introuvable : {cle}")

        colonnes[cle] = entetes[cle]

    return colonnes


def run_isolation_client():
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Fichier Excel introuvable : {EXCEL_FILE}"
        )

    excel = None
    wb = None
    browser = None

    pythoncom.CoInitialize()

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(EXCEL_FILE))
        ws = wb.ActiveSheet

        colonnes = trouver_colonnes(ws)
        col_su = colonnes["su"]
        col_statut = colonnes["statut"]
        col_etat = colonnes["etat résil"]

        derniere_ligne = ws.Cells(
            ws.Rows.Count,
            col_su,
        ).End(-4162).Row

        lignes = []
        ignorees = []

        for row in range(2, derniere_ligne + 1):
            valeur = ws.Cells(row, col_su).Value

            if valeur is None:
                continue

            su = str(valeur).strip().upper()

            if not su.startswith("CD0"):
                continue

            # Seules les lignes à isoler (Etat Résil = NOK) sont traitées.
            etat = ws.Cells(row, col_etat).Value

            if not doit_traiter(etat, MODE_A_ISOLER):
                ignorees.append((row, su, normaliser_etat(etat)))
                continue

            lignes.append((row, su))

        # Avant l'annonce du total : la progression ne compte que les
        # lignes réellement traitées.
        annoncer_ignorees(ignorees, MODE_A_ISOLER)

        print(f"[EXCEL] {len(lignes)} SU à traiter.")

        if not lignes:
            return

        with sync_playwright() as playwright:
            try:
                browser, context, page = launch_browser(playwright)

                for row, su in lignes:
                    try:
                        statut = traiter_su_isolation(page, su)

                        ws.Cells(row, col_statut).Value = statut
                        print(
                            f"[EXCEL] Ligne {row} | {su} | {statut}"
                        )

                    except Exception as error:
                        print(
                            f"[ERREUR] Ligne {row} | {su} | "
                            f"{type(error).__name__}: {error}"
                        )
                        ws.Cells(row, col_statut).Value = "ERREUR"

                    finally:
                        wb.Save()

            finally:
                if browser:
                    close_browser(browser)

        print("[ISOLATION] Traitement terminé.")

    finally:
        if wb is not None:
            try:
                wb.Save()
            finally:
                wb.Close(SaveChanges=True)

        if excel is not None:
            excel.Quit()

        pythoncom.CoUninitialize()


if __name__ == "__main__":
    run_isolation_client()