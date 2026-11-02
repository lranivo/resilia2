"""
Socle commun des modules Salto (gel / dégel).

- lecture du fichier Excel (colonnes DPS/ACT et Statut) ;
- ouverture du portail + Salto avec nouvelles tentatives ;
- recherche d'un DPS et clic sur « Autres activités » ;
- diagnostics en cas d'échec ;
- boucle générique « une ligne Excel = un DPS » (executer_sur_excel).

Chaque module métier (gel_salto, degel_salto) ne contient que ses
propres étapes PEGA.
"""

import re
import time
from datetime import datetime
from pathlib import Path

import pythoncom
import win32com.client

from playwright.sync_api import sync_playwright

from browser import launch_browser, close_browser
from portal import open_portal, open_salto
from isolation_client import chrono, sauvegarder_diagnostic
from salto_utils import trouver_dans_frames, sans_accents
from etat_resil import (
    annoncer_ignorees,
    doit_traiter,
    normaliser_etat,
)


EXCEL_FILE = Path(__file__).resolve().parent / "resil.xlsx"
DIAGNOSTIC_DIR = Path(__file__).resolve().parent / "diagnostics"

RE_DPS = re.compile(r"^(DPS|ACT)[-\s]?\d+", re.IGNORECASE)
RE_AUTRES_ACTIVITES = re.compile(r"Autres\s+activit[eé]s", re.IGNORECASE)


def ouvrir_salto_avec_reessai(context, portal_page, tentatives=3):
    """
    Ouvre le portail puis Salto. En cas d'échec (par exemple
    net::ERR_UNEXPECTED sur le portail), ferme les onglets créés par la
    tentative et recommence, comme open_wasac dans check_wasac.
    """

    derniere_erreur = None

    for tentative in range(1, tentatives + 1):
        pages_avant = set(context.pages)

        try:
            print(
                f"[SALTO] Ouverture du portail et de Salto, "
                f"tentative {tentative}/{tentatives}..."
            )

            open_portal(portal_page)
            return open_salto(context, portal_page)

        except Exception as error:
            derniere_erreur = error
            resume = str(error).splitlines()[0] if str(error) else ""

            print(
                f"[SALTO] Tentative {tentative}/{tentatives} échouée : "
                f"{type(error).__name__}: {resume}"
            )

            for page in list(context.pages):
                if page not in pages_avant and page != portal_page:
                    try:
                        page.close()
                    except Exception:
                        pass

            if tentative < tentatives:
                time.sleep(3)

    raise RuntimeError(
        f"Ouverture de Salto impossible après {tentatives} tentatives : "
        f"{derniere_erreur}"
    )


def search_dps(page, dps_id):
    print(f"[SALTO] Recherche du champ pour {dps_id}...")

    # Le placeholder reste stable même si l'ID HTML change.
    search_field = page.get_by_placeholder(
        "Recherche...",
        exact=True
    )

    search_field.wait_for(
        state="visible",
        timeout=90000
    )

    search_field.click()
    search_field.fill(dps_id)
    search_field.press("Enter")

    print(f"[SALTO] Recherche envoyée : {dps_id}")


def ouvrir_autres_activites(page):
    """Étape 4 : clic sur « Autres activités » une fois l'écran chargé."""

    _, element = trouver_dans_frames(
        page,
        [
            lambda f: f.locator("span.menu-item-title").filter(
                has_text=RE_AUTRES_ACTIVITES
            ),
            lambda f: f.get_by_text(RE_AUTRES_ACTIVITES),
        ],
        timeout_ms=90000,
        label="Autres activités",
    )

    element.click(timeout=10000)
    print("[SALTO] « Autres activités » cliqué.")


JS_ELEMENTS_VISIBLES = """() => Array.from(document.querySelectorAll(
    "button, input[type=button], input[type=submit], a[onclick], " +
    "[role=button], [role=menuitem]"
)).filter(e => e.offsetParent !== null).slice(0, 60).map(e =>
    e.tagName + " | " +
    (e.innerText || e.value || e.title || "").trim().slice(0, 80) +
    " | name=" + (e.getAttribute("name") || "") +
    " | onclick=" + (e.getAttribute("onclick") || "").slice(0, 120)
)"""


def sauvegarder_diagnostic_complet(page, dps_id):
    """
    Diagnostic étendu : tous les onglets du navigateur, leurs frames et
    la liste des boutons / liens visibles, avec une capture par onglet.
    """

    try:
        dossier = DIAGNOSTIC_DIR
        dossier.mkdir(parents=True, exist_ok=True)
        horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = dossier / f"{dps_id}_complet_{horodatage}"

        lignes = []

        for index, onglet in enumerate(list(page.context.pages)):
            if onglet.is_closed():
                continue

            try:
                titre = onglet.title()
            except Exception:
                titre = "?"

            lignes.append(f"ONGLET {index} : {onglet.url} | {titre}")

            try:
                onglet.screenshot(path=f"{base}_onglet{index}.png")
            except Exception as error:
                lignes.append(f"  (capture impossible : {error})")

            for frame in list(onglet.frames):
                lignes.append(f"  FRAME : {frame.url}")

                try:
                    for element in frame.evaluate(JS_ELEMENTS_VISIBLES):
                        lignes.append(f"      {element}")
                except Exception:
                    continue

        Path(f"{base}.txt").write_text(
            "\n".join(lignes), encoding="utf-8"
        )

        print(f"[DIAG] Diagnostic complet : {base}.txt")

    except Exception as error:
        print(f"[DIAG] Échec du diagnostic complet : {error}")


def trouver_colonnes(ws, avec_etat=False):
    entetes = {}
    derniere_colonne = ws.UsedRange.Columns.Count

    for col in range(1, derniere_colonne + 1):
        valeur = ws.Cells(1, col).Value

        if valeur is not None:
            entetes[str(valeur).strip().casefold()] = col

    colonnes = {}

    cles = ["dps/act", "statut"]

    if avec_etat:
        cles.append("etat résil")

    for cle in cles:
        if cle not in entetes:
            raise RuntimeError(f"Colonne Excel introuvable : {cle}")

        colonnes[cle] = entetes[cle]

    return colonnes


def executer_sur_excel(nom, action, traiter_dps, statut_ok,
                       filtre_etat=None):
    """
    Boucle générique : une ligne Excel = un DPS.

    nom          : préfixe des logs (ex. "GEL_SALTO")
    action       : verbe pour les logs (ex. "geler")
    traiter_dps  : fonction (page_salto, dps) -> None, qui lève une
                   exception en cas d'échec
    statut_ok    : valeur écrite dans « Statut » une fois la ligne
                   traitée (les lignes déjà à cette valeur sont ignorées)
    filtre_etat  : None (toutes les lignes) ou un mode de etat_resil
                   (MODE_OK / MODE_A_ISOLER) : seules les lignes dont la
                   colonne « Etat Résil » correspond sont traitées.
    """

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

        colonnes = trouver_colonnes(ws, avec_etat=filtre_etat is not None)
        col_dps = colonnes["dps/act"]
        col_statut = colonnes["statut"]
        col_etat = colonnes.get("etat résil")

        derniere_ligne = ws.Cells(
            ws.Rows.Count,
            col_dps,
        ).End(-4162).Row

        lignes = []
        ignorees = []

        for row in range(2, derniere_ligne + 1):
            valeur = ws.Cells(row, col_dps).Value

            if valeur is None:
                continue

            dps = str(valeur).strip().upper()

            if not RE_DPS.match(dps):
                print(f"[EXCEL] Ligne {row} ignorée (format) : {dps!r}")
                continue

            if filtre_etat is not None:
                etat = ws.Cells(row, col_etat).Value

                if not doit_traiter(etat, filtre_etat):
                    ignorees.append((row, dps, normaliser_etat(etat)))
                    continue

            statut = ws.Cells(row, col_statut).Value

            if (
                statut is not None
                and sans_accents(str(statut).strip())
                == sans_accents(statut_ok)
            ):
                print(
                    f"[EXCEL] Ligne {row} | {dps} déjà « {statut_ok} », "
                    "ignoré."
                )
                continue

            lignes.append((row, dps))

        # Avant l'annonce du total : la progression ne compte que les
        # lignes réellement traitées.
        annoncer_ignorees(ignorees, filtre_etat)

        print(f"[EXCEL] {len(lignes)} DPS à {action}.")

        if not lignes:
            return

        with sync_playwright() as playwright:
            try:
                browser, context, portal_page = launch_browser(playwright)

                for row, dps in lignes:
                    salto_page = None
                    debut = time.monotonic()

                    print(f"\n[{nom}] Début du DPS {dps}")

                    try:
                        # Onglet Salto neuf à chaque DPS : évite de cliquer
                        # sur le menu de la fiche précédente.
                        with chrono(f"{dps} | ouverture portail + Salto"):
                            salto_page = ouvrir_salto_avec_reessai(
                                context, portal_page
                            )

                        traiter_dps(salto_page, dps)

                        ws.Cells(row, col_statut).Value = statut_ok
                        print(f"[EXCEL] Ligne {row} | {dps} | {statut_ok}")

                    except Exception as error:
                        print(
                            f"[{nom}][ERREUR] Ligne {row} | {dps} | "
                            f"{type(error).__name__}: {error}"
                        )

                        if salto_page is not None:
                            sauvegarder_diagnostic(salto_page, dps)
                            sauvegarder_diagnostic_complet(salto_page, dps)

                    finally:
                        wb.Save()

                        if salto_page is not None:
                            try:
                                salto_page.close()
                            except Exception:
                                pass

                        print(
                            f"[TEMPS] {dps} | TOTAL : "
                            f"{time.monotonic() - debut:.1f}s"
                        )

            finally:
                if browser:
                    close_browser(browser)

        print(f"[{nom}] Traitement terminé.")

    finally:
        if wb is not None:
            try:
                wb.Save()
            finally:
                wb.Close(SaveChanges=True)

        if excel is not None:
            excel.Quit()

        pythoncom.CoUninitialize()
