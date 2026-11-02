import re
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import pythoncom
import win32com.client

from playwright.sync_api import (
    sync_playwright,
    Error as PlaywrightError,
    TimeoutError as PlaywrightTimeoutError,
)

from browser import launch_browser, close_browser


EXCEL_FILE = Path(__file__).resolve().parent / "resil.xlsx"

RESEAU_ID = ""

WASAC_URL = (
    "https://wasac.sso.infra.ftgroup/perl/wasac-new/loginsso.cgi"
)

SEUIL_JOURS = 7
MAX_REFERENCES = 2


MOIS_FR = {
    "janv": 1,
    "jan": 1,
    "févr": 2,
    "févr": 2,
    "fév": 2,
    "mars": 3,
    "avr": 4,
    "avril": 4,
    "mai": 5,
    "juin": 6,
    "juil": 7,
    "juillet": 7,
    "août": 8,
    "sept": 9,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "déc": 12,
    "dec": 12,
}


# ============================================================
# OUTILS WASAC
# ============================================================

def open_wasac(page, max_attempts=3):
    """Ouvre WASAC avec plusieurs tentatives si la navigation échoue."""

    for attempt in range(1, max_attempts + 1):
        try:
            print(
                f"[WASAC] Ouverture, tentative "
                f"{attempt}/{max_attempts}..."
            )

            page.goto(
                WASAC_URL,
                wait_until="domcontentloaded",
                timeout=120000,
            )

            print(f"[WASAC] URL après navigation : {page.url}")
            return

        except PlaywrightError as error:
            print(
                f"[WASAC] Échec navigation "
                f"{attempt}/{max_attempts} : {error}"
            )

            if attempt == max_attempts:
                raise

            time.sleep(3 * attempt)

    raise RuntimeError("Impossible d'ouvrir WASAC.")


def find_visible(page, text, timeout=60000, exact=True):
    """Recherche un texte visible dans la page ou ses iframes."""

    deadline = time.monotonic() + timeout

    while time.monotonic() < deadline:
        contexts = [page] + [
            frame for frame in page.frames
            if frame != page.main_frame
        ]

        for context in contexts:
            try:
                locator = context.get_by_text(
                    text, exact=exact
                ).first

                if locator.count() > 0 and locator.is_visible():
                    return locator

            except Exception:
                continue

        page.wait_for_timeout(250)

    raise PlaywrightTimeoutError(
        f"Texte visible introuvable : {text!r} "
        f"après {timeout} ms"
    )


def find_reseau_field(page, timeout=90000):
    """Recherche le champ de recherche du réseau dans les iframes."""

    deadline = time.monotonic() + timeout

    while time.monotonic() < deadline:
        contexts = [page] + [
            frame for frame in page.frames
            if frame != page.main_frame
        ]

        for context in contexts:
            for selector in (
                "#nom_reseau_TB",
                'input[name="nom_reseau"]',
            ):
                try:
                    locator = context.locator(selector).first

                    if locator.count() > 0 and locator.is_visible():
                        print(
                            f"[WASAC] Champ trouvé : {selector} "
                            f"| contexte : {context.url}"
                        )
                        return locator

                except Exception:
                    continue

        page.wait_for_timeout(250)

    raise PlaywrightTimeoutError(
        "Champ nom_reseau_TB introuvable dans la page et les iframes."
    )


def click_text(page, text, timeout=60000):
    """Attend un texte visible puis clique dessus."""

    locator = find_visible(
        page, text, timeout=timeout, exact=True
    )

    locator.click(timeout=15000)
    print(f"[WASAC] Clic effectué : {text}")

    return locator


def texte_disponible(page):
    """Récupère le texte visible de la page et de ses iframes."""

    textes = []

    contexts = [page] + [
        frame for frame in page.frames
        if frame != page.main_frame
    ]

    for context in contexts:
        try:
            contenu = context.locator("body").inner_text(
                timeout=3000
            )

            if contenu.strip():
                textes.append(contenu)

        except Exception:
            continue

    return "\n".join(textes)


def rechercher_acquitter(page):
    """Clique sur Acquitter si le bouton est présent."""

    for frame in list(page.frames):
        try:
            bouton = frame.get_by_text(
                "Acquitter", exact=True
            ).first

            if bouton.count() > 0 and bouton.is_visible():
                bouton.click(timeout=5000)

                print("[WASAC] Acquitter cliqué.")
                page.wait_for_timeout(2000)
                return

        except Exception:
            continue

    print("[WASAC] Pas d'Acquitter, poursuite.")


@contextmanager
def chrono(label):
    """Affiche le temps passé dans une étape."""

    debut = time.monotonic()
    try:
        yield
    finally:
        print(f"[TEMPS] {label} : {time.monotonic() - debut:.1f}s")


def attendre_resultat_recherche(page, max_ms=5000, marge_ms=500):
    """
    Remplace la pause fixe après la recherche : dès que le menu du
    feuillet est chargé (feuillet=<réf> dans l'URL d'un frame), on
    continue après une courte marge. Pire cas : max_ms (comme avant).
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


def exploitation_disponible(page, timeout=12000, stabilite=1000):
    """
    Attente adaptative d'Exploitation.

    - Exploitation visible                    -> True immédiatement.
    - Menu chargé (Synthèse visible) et page
      stable `stabilite` ms sans Exploitation -> False rapidement.
    - Page pas encore chargée                 -> on attend jusqu'à
                                                 `timeout` ms.
    """
    print("[DEBUG] Entrée dans exploitation_disponible()")

    deadline = time.monotonic() + timeout / 1000
    derniere_signature = None
    stable_depuis = None

    while time.monotonic() < deadline:
        if texte_visible(page, "Exploitation"):
            print("[DEBUG] Exploitation trouvé.")
            return True

        if any(texte_visible(page, t) for t in INDICATEURS_MENU):
            signature = signature_pages(page)
            maintenant = time.monotonic()

            if signature != derniere_signature:
                derniere_signature = signature
                stable_depuis = maintenant

            elif maintenant - stable_depuis >= stabilite / 1000:
                print(
                    "[DEBUG] Exploitation absent : page chargée "
                    "et stable."
                )
                return False

        page.wait_for_timeout(200)

    print("[DEBUG] Exploitation absent : délai dépassé.")
    return False



# ============================================================
# EXTRACTION DES RÉFÉRENCES MULTIPLES
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


def extraire_references_multiples(page):
    """
    Ouvre Synthèse des informations et récupère les références
    dont l'état est RESILIATION EN COURS.

    Aucun fichier temporaire ni presse-papiers n'est utilisé.
    Lecture rapide : pas de pause fixe, pas d'inner_text à 5 s.
    """

    print(
        "[WASAC] Exploitation introuvable : "
        "recherche de Synthèse des informations..."
    )

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

    # Une référence doit appartenir à la même ligne que l'état
    # ou se trouver juste avant celui-ci dans le texte extrait.
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

            # La référence recherchée est généralement la plus
            # proche de l'état RESILIATION EN COURS.
            reference = matches[-1].upper()

            # Corriger un O en troisième caractère après CD si besoin.
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


def verifier_dernier_ping(page, seuil_jours=7, timeout=60):
    """Lit Dernier Ping OK et retourne (ancienneté, résultat)."""

    date_regex = re.compile(
        r"(\d{1,2})\s+"
        r"(janv?\.?|févr?\.?|mars|avr\.?|avril|mai|juin|"
        r"juil?\.?|juillet|août|sept?\.?|sep\.?|oct\.?|"
        r"nov\.?|déc\.?|dec\.?)\s+"
        r"(\d{4})\s+"
        r"(\d{1,2}):(\d{2}):(\d{2})",
        re.IGNORECASE,
    )

    print("[PING] Recherche de la date Dernier Ping OK...")

    deadline = time.monotonic() + timeout
    date_ping = None

    while time.monotonic() < deadline:
        contexts = [page] + [
            frame for frame in page.frames
            if frame != page.main_frame
        ]

        for context in contexts:
            try:
                contenu = context.locator("body").inner_text(
                    timeout=3000
                )

                section = re.search(
                    r"Dernier\s+Ping\s+OK"
                    r"(.*?)"
                    r"Premier\s+Ping\s+OK",
                    contenu,
                    re.IGNORECASE | re.DOTALL,
                )

                if not section:
                    continue

                match_date = date_regex.search(section.group(1))

                if not match_date:
                    continue

                jour, mois, annee, heure, minute, seconde = (
                    match_date.groups()
                )

                mois_cle = mois.lower().replace(".", "")
                mois_num = MOIS_FR.get(mois_cle)

                if mois_num is None:
                    continue

                date_ping = datetime(
                    int(annee),
                    mois_num,
                    int(jour),
                    int(heure),
                    int(minute),
                    int(seconde),
                )

                break

            except Exception:
                continue

        if date_ping is not None:
            break

        page.wait_for_timeout(1000)

    if date_ping is None:
        raise RuntimeError(
            "Impossible de lire la date Dernier Ping OK dans WASAC."
        )

    maintenant = datetime.now()
    secondes = int((maintenant - date_ping).total_seconds())

    if secondes < 0:
        raise RuntimeError(
            "La date du dernier ping est dans le futur."
        )

    resultat = (
        "NOK"
        if secondes < seuil_jours * 86400
        else "OK"
    )

    jours = secondes // 86400
    heures = (secondes % 86400) // 3600
    minutes = (secondes % 3600) // 60

    anciennete_texte = (
        f"{jours}j {heures:02d}h {minutes:02d}m"
    )

    print(f"[PING] Dernier Ping OK : {date_ping}")
    print(f"[PING] Ancienneté : {anciennete_texte}")
    print(f"[PING] Résultat : {resultat}")

    return anciennete_texte, resultat


# ============================================================
# CHECKING CLASSIQUE D'UN FEUILLET
# ============================================================

def rechercher_reseau(page, reseau_id):
    """Recherche une référence dans l'onglet WASAC fourni."""

    open_wasac(page)

    champ = find_reseau_field(page, timeout=90000)
    champ.fill(reseau_id)
    champ.press("Enter")

    print(f"[WASAC] Recherche envoyée : {reseau_id}")

    with chrono("    attente du résultat"):
        attendre_resultat_recherche(page, max_ms=5000)
        rechercher_acquitter(page)


def controle_classique(page, reseau_id):
    """
    Contrôle une référence qui doit présenter Exploitation.
    Retourne (ancienneté, résultat).
    """

    rechercher_reseau(page, reseau_id)

    if not exploitation_disponible(page):
        raise RuntimeError(
            f"Exploitation introuvable pour {reseau_id}. "
            "Le feuillet nécessite un traitement particulier."
        )

    click_text(page, "Exploitation", timeout=15000)
    page.wait_for_timeout(1000)

    click_text(page, "Informations", timeout=60000)

    find_visible(
        page,
        "Polling SNMP",
        timeout=120000,
        exact=True,
    )

    print("[WASAC] Section Polling SNMP détectée.")

    # Conserve le comportement existant : la date du dernier ping
    # est lue dans la page après l'apparition de Polling SNMP.
    return verifier_dernier_ping(
        page,
        seuil_jours=SEUIL_JOURS,
        timeout=60,
    )


# ============================================================
# TRAITEMENT D'UN SU A LES RÉFÉRENCES MULTIPLES
# ============================================================


def traiter_reseau(page, reseau_id):
    """
    Traite un SU :
    - classique : contrôle direct du ping ;
    - multiple : extraction des références, puis contrôle dans
      un nouvel onglet pour chaque référence.
    """

    print(f"\n[EXCEL] Traitement du réseau {reseau_id}")

    # Recherche initiale : une seule fois.
    rechercher_reseau(page, reseau_id)

    print("[DEBUG] Vérification du type de feuillet...")

    # Cette vérification ne doit être effectuée qu'une seule fois.
    if exploitation_disponible(page):
        print("[WASAC] Feuillet classique détecté.")

        click_text(page, "Exploitation", timeout=15000)
        page.wait_for_timeout(500)

        click_text(page, "Informations", timeout=30000)

        find_visible(
            page,
            "Polling SNMP",
            timeout=60000,
            exact=True,
        )

        anciennete, resultat = verifier_dernier_ping(
            page,
            seuil_jours=SEUIL_JOURS,
            timeout=30,
        )

        # Ordre attendu par l'Excel : (R1, R2, Etat Résil).
        # SU simple : une seule ancienneté (R1), R2 vide.
        return anciennete, "", resultat

    print("[WASAC] Feuillet multiple détecté.")

    # Extraction unique des références.
    references = extraire_references_multiples(page)

    resultats = []

    for numero, reference in enumerate(references, start=1):
        print(f"\n[WASAC] Contrôle R{numero} : {reference}")

        page_check = page.context.new_page()

        try:
            anciennete, resultat = controle_classique(
                page_check,
                reference,
            )

            resultats.append({
                "reference": reference,
                "anciennete": anciennete,
                "resultat": resultat,
                "erreur": None,
            })

        except Exception as error:
            print(
                f"[WASAC][ERREUR] {reference} : "
                f"{type(error).__name__}: {error}"
            )

            resultats.append({
                "reference": reference,
                "anciennete": "",
                "resultat": "",
                "erreur": str(error),
            })

        finally:
            try:
                page_check.close()
            except Exception:
                pass

    r1 = resultats[0]["anciennete"] if resultats else ""
    r2 = resultats[1]["anciennete"] if len(resultats) > 1 else ""

    statuts = [r["resultat"] for r in resultats if r["resultat"]]
    erreurs = [r for r in resultats if r["erreur"] is not None]

    if "NOK" in statuts:
        etat = "NOK"
    elif erreurs or len(statuts) != len(references):
        etat = "ERREUR"
    elif statuts and all(s == "OK" for s in statuts):
        etat = "OK"
    else:
        etat = "ERREUR"

    print(f"[EXCEL] R1={r1} | R2={r2} | Etat={etat}")

    return r1, r2, etat



# ============================================================
# EXCEL : REPÉRAGE DES COLONNES
# ============================================================

def trouver_colonnes_com(ws):
    """Retrouve les colonnes Excel à partir des en-têtes de la ligne 1."""

    entetes = {}

    derniere_colonne = ws.UsedRange.Columns.Count

    for col in range(1, derniere_colonne + 1):
        valeur = ws.Cells(1, col).Value

        if valeur is not None:
            cle = str(valeur).strip().casefold()
            entetes[cle] = col

    colonnes_requises = {
        "su": "SU",
        "etat résil": "Etat Résil",
        "injoignabilité r1": "Injoignabilité R1",
        "injoignabilité r2": "Injoignabilité R2",
    }

    colonnes = {}

    for cle, nom in colonnes_requises.items():
        if cle not in entetes:
            raise RuntimeError(
                f"Colonne Excel introuvable : {nom}"
            )

        colonnes[nom] = entetes[cle]

    return colonnes


# ============================================================
# EXCEL : TRAITEMENT AUTOMATIQUE
# ============================================================

def run_check_wasac_excel():
    """Traite les SU commençant par CD0 dans resil.xlsx."""

    excel_file = EXCEL_FILE

    if not excel_file.exists():
        raise FileNotFoundError(
            f"Fichier Excel introuvable : {excel_file}"
        )

    excel = None
    wb = None
    browser = None

    pythoncom.CoInitialize()

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(excel_file))
        ws = wb.ActiveSheet

        print(f"[EXCEL] Fichier ouvert : {excel_file.name}")
        print(f"[EXCEL] Onglet utilisé : {ws.Name}")

        colonnes = trouver_colonnes_com(ws)

        col_su = colonnes["SU"]
        col_etat = colonnes["Etat Résil"]
        col_r1 = colonnes["Injoignabilité R1"]
        col_r2 = colonnes["Injoignabilité R2"]

        derniere_ligne = ws.Cells(
            ws.Rows.Count, col_su
        ).End(-4162).Row  # xlUp

        lignes = []

        for row in range(2, derniere_ligne + 1):
            valeur = ws.Cells(row, col_su).Value

            if valeur is None:
                continue

            reseau_id = str(valeur).strip().upper()

            if reseau_id.startswith("CD0"):
                lignes.append((row, reseau_id))

        print(f"[EXCEL] {len(lignes)} réseau(x) à traiter.")

        if not lignes:
            return

        with sync_playwright() as playwright:
            try:
                browser, context, page = launch_browser(playwright)

                for row, reseau_id in lignes:
                    try:
                        anciennete_r1, anciennete_r2, etat_resil = (
                            traiter_reseau(page, reseau_id)
                        )

                        ws.Cells(row, col_r1).Value = anciennete_r1
                        ws.Cells(row, col_r2).Value = anciennete_r2
                        ws.Cells(row, col_etat).Value = etat_resil

                        print(
                            f"[EXCEL] Ligne {row} | {reseau_id} | "
                            f"R1={anciennete_r1} | "
                            f"R2={anciennete_r2} | "
                            f"Etat={etat_resil}"
                        )

                    except Exception as error:
                        ws.Cells(row, col_etat).Value = "ERREUR"

                        print(
                            f"[EXCEL][ERREUR] Ligne {row} | "
                            f"{reseau_id} | "
                            f"{type(error).__name__}: {error}"
                        )

                    finally:
                        # Sauvegarde après chaque SU.
                        wb.Save()

            finally:
                if browser:
                    close_browser(browser)

        print("[EXCEL] Traitement terminé.")

    finally:
        if wb is not None:
            try:
                wb.Save()
            finally:
                wb.Close(SaveChanges=True)

        if excel is not None:
            excel.Quit()

        pythoncom.CoUninitialize()


# ============================================================
# TEST MANUEL
# ============================================================

def run_check_wasac():
    """Test manuel du SU défini dans RESEAU_ID."""

    browser = None

    try:
        with sync_playwright() as playwright:
            browser, context, page = launch_browser(playwright)

            try:
                anciennete_r1, anciennete_r2, etat = traiter_reseau(
                    page, RESEAU_ID
                )

                print(
                    f"[WASAC] Résultat final : "
                    f"R1={anciennete_r1}, "
                    f"R2={anciennete_r2}, "
                    f"Etat Résil={etat}"
                )

                page.wait_for_timeout(10000)

            finally:
                if browser:
                    close_browser(browser)

    except Exception as error:
        print(
            f"[WASAC][ERREUR] "
            f"{type(error).__name__}: {error}"
        )


if __name__ == "__main__":
    run_check_wasac_excel()