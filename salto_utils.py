"""
Outils communs pour les modules Salto (PEGA) : attente de stabilité,
recherche d'éléments dans le DOM, clic « humain » et écouteurs de
diagnostic (dialogues JavaScript, erreurs JS, nouveaux onglets).
"""

import time
import unicodedata
from contextlib import contextmanager


# Réglages de rapidité du clic fiable (chemin optimiste rapide ; les
# vérifications et les nouvelles tentatives rattrapent les rares échecs).
STABILITE_AVANT_CLIC_MS = 500
PAUSE_SURVOL_MS = 150


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


def attendre_page_stable(page, stabilite_ms=800, timeout_ms=8000):
    """
    Attend que la page cesse de bouger (frames et textes inchangés
    pendant `stabilite_ms`), pour ne pas cliquer pendant un rafraîchissement
    PEGA.
    """

    deadline = time.monotonic() + timeout_ms / 1000
    derniere = None
    depuis = None

    while time.monotonic() < deadline:
        signature = signature_pages(page)
        maintenant = time.monotonic()

        if signature != derniere:
            derniere = signature
            depuis = maintenant

        elif maintenant - depuis >= stabilite_ms / 1000:
            return

        page.wait_for_timeout(150)


def trouver_attache(page, builders, timeout_ms, label):
    """
    Comme trouver_dans_frames, mais n'exige pas que l'élément soit
    visible : il suffit qu'il soit présent dans le DOM (utile pour un
    menu qui ne reste affiché que quelques millisecondes).
    Retourne (frame, locator).
    """

    deadline = time.monotonic() + timeout_ms / 1000

    while time.monotonic() < deadline:
        for frame in list(page.frames):
            for build in builders:
                try:
                    locator = build(frame).first

                    if locator.count() > 0:
                        return frame, locator

                except Exception:
                    continue

        page.wait_for_timeout(100)

    raise RuntimeError(
        f"Élément introuvable ({timeout_ms / 1000:.0f}s) : {label}"
    )


def clic_lent(page, element):
    """
    Clic plus proche d'un humain : survol, courte pause, puis appui
    de 150 ms sur le début de l'élément.
    """

    element.hover(timeout=10000)
    page.wait_for_timeout(300)
    element.click(position={"x": 8, "y": 6}, delay=150, timeout=10000)


@contextmanager
def ecouteurs_diagnostic(page):
    """
    Journalise les boîtes de dialogue JavaScript, les erreurs JS et les
    nouveaux onglets. Fournit un état partagé : etat["erreur_js"] passe
    à True quand une erreur JS survient.
    """

    compteur = {"console": 0}
    etat = {"erreur_js": False}

    def sur_dialogue(dialogue):
        print(
            f"[SALTO][DIALOGUE] {dialogue.type} : {dialogue.message!r}"
        )
        # Comportement par défaut de Playwright : on refuse.
        dialogue.dismiss()

    def sur_console(message):
        if message.type == "error" and compteur["console"] < 15:
            compteur["console"] += 1
            print(f"[SALTO][CONSOLE] {message.text[:200]}")

    def sur_erreur_js(erreur):
        etat["erreur_js"] = True
        print(f"[SALTO][ERREUR JS] {str(erreur)[:200]}")

    def sur_nouvel_onglet(onglet):
        print(f"[SALTO][NOUVEL ONGLET] {onglet.url!r}")

    page.on("dialog", sur_dialogue)
    page.on("console", sur_console)
    page.on("pageerror", sur_erreur_js)
    page.context.on("page", sur_nouvel_onglet)

    try:
        yield etat

    finally:
        for evenement, fonction in (
            ("dialog", sur_dialogue),
            ("console", sur_console),
            ("pageerror", sur_erreur_js),
        ):
            try:
                page.remove_listener(evenement, fonction)
            except Exception:
                pass

        try:
            page.context.remove_listener("page", sur_nouvel_onglet)
        except Exception:
            pass


def sans_accents(texte):
    return "".join(
        c for c in unicodedata.normalize("NFD", texte)
        if unicodedata.category(c) != "Mn"
    ).casefold()


def trouver_dans_frames(page, builders, timeout_ms, label):
    """
    Attend qu'un élément soit visible dans la page ou l'un de ses frames.
    `builders` : liste de fonctions frame -> locator.
    Retourne (frame, locator).
    """

    deadline = time.monotonic() + timeout_ms / 1000

    while time.monotonic() < deadline:
        for frame in list(page.frames):
            for build in builders:
                try:
                    locator = build(frame).first

                    if locator.count() > 0 and locator.is_visible():
                        return frame, locator

                except Exception:
                    continue

        page.wait_for_timeout(200)

    raise RuntimeError(
        f"Élément introuvable ({timeout_ms / 1000:.0f}s) : {label}"
    )


def attendre_dans_frame(page, frame, builders, timeout_ms, label):
    """Même chose, mais limité à un frame donné (le formulaire)."""

    deadline = time.monotonic() + timeout_ms / 1000

    while time.monotonic() < deadline:
        for build in builders:
            try:
                locator = build(frame).first

                if locator.count() > 0 and locator.is_visible():
                    return locator

            except Exception:
                continue

        page.wait_for_timeout(200)

    raise RuntimeError(
        f"Élément introuvable ({timeout_ms / 1000:.0f}s) : {label}"
    )


def trouver_partout(page, builders, timeout_ms, label):
    """
    Comme trouver_dans_frames, mais cherche aussi dans les autres onglets
    du navigateur (au cas où PEGA ouvrirait le traitement dans un nouvel
    onglet ou une nouvelle fenêtre).
    """

    deadline = time.monotonic() + timeout_ms / 1000

    while time.monotonic() < deadline:
        for onglet in list(page.context.pages):
            if onglet.is_closed():
                continue

            for frame in list(onglet.frames):
                for build in builders:
                    try:
                        locator = build(frame).first

                        if locator.count() > 0 and locator.is_visible():
                            return frame, locator

                    except Exception:
                        continue

        page.wait_for_timeout(200)

    raise RuntimeError(
        f"Élément introuvable ({timeout_ms / 1000:.0f}s) : {label}"
    )


def attendre_absence(page, builders, timeout_ms, label):
    """
    Attend qu'aucun élément correspondant ne soit visible (dans aucun
    onglet ni frame). Lève RuntimeError si l'élément reste affiché.
    """

    deadline = time.monotonic() + timeout_ms / 1000

    while time.monotonic() < deadline:
        visible = False

        for onglet in list(page.context.pages):
            if onglet.is_closed():
                continue

            for frame in list(onglet.frames):
                for build in builders:
                    try:
                        locator = build(frame).first

                        if locator.count() > 0 and locator.is_visible():
                            visible = True
                            break

                    except Exception:
                        continue

                if visible:
                    break

            if visible:
                break

        if not visible:
            return True

        page.wait_for_timeout(150)

    raise RuntimeError(
        f"Élément toujours affiché ({timeout_ms / 1000:.0f}s) : {label}"
    )


def cliquer_avec_verification(
    page,
    localiser,
    verifier,
    label,
    timeout_localisation_ms=60000,
    timeout_verification_ms=10000,
):
    """
    Clic fiable sur un élément piloté par le JavaScript PEGA.

    Un clic trop précoce peut partir sans rien déclencher (erreur JS,
    ligne seulement sélectionnée, menu qui se referme). On vérifie donc
    que l'action a abouti ; sinon on réessaie, puis on change de méthode.

    localiser(timeout_ms)  -> Locator de l'élément à cliquer (RuntimeError
                              s'il est introuvable)
    verifier(timeout_ms)   -> résultat quelconque si l'action a abouti
                              (ex. l'étape suivante est affichée),
                              RuntimeError sinon
    Retourne le résultat de verifier().

    Avant chaque clic : page stable + survol. Méthodes : clic, clic lent,
    clic JavaScript ; jusqu'à 3 essais par méthode si une erreur JS
    survient pendant le clic (page pas prête).
    """

    strategies = [
        ("clic", lambda el: el.click(timeout=10000)),
        ("clic lent", lambda el: clic_lent(page, el)),
        ("clic JavaScript", lambda el: el.evaluate("el => el.click()")),
    ]

    deja_clique = False

    with ecouteurs_diagnostic(page) as etat:
        for numero, (nom, action) in enumerate(strategies, start=1):
            derniere = numero == len(strategies)

            for essai in range(1, 4):
                try:
                    element = localiser(
                        10000 if deja_clique else timeout_localisation_ms
                    )

                except RuntimeError:
                    if deja_clique:
                        # L'élément a disparu après un clic : l'action
                        # est peut-être en cours, la vérification tranche.
                        return verifier(30000)

                    raise

                attendre_page_stable(page, STABILITE_AVANT_CLIC_MS, 5000)

                try:
                    element.scroll_into_view_if_needed()
                    element.hover(timeout=5000)
                    page.wait_for_timeout(PAUSE_SURVOL_MS)
                except Exception:
                    pass

                etat["erreur_js"] = False
                action(element)
                deja_clique = True
                page.wait_for_timeout(200)

                suffixe = f" (essai {essai}/3)" if essai > 1 else ""
                print(f"[SALTO] « {label} » : {nom}{suffixe}.")

                if etat["erreur_js"]:
                    delai = 3000
                else:
                    delai = 30000 if derniere else timeout_verification_ms

                try:
                    return verifier(delai)

                except RuntimeError:
                    pass

                if etat["erreur_js"]:
                    print(
                        "[SALTO] Erreur JS pendant le clic : page pas "
                        "prête, attente puis nouvel essai..."
                    )
                    attendre_page_stable(page, 1500, 8000)
                    continue

                print(
                    f"[SALTO] « {label} » sans effet après le {nom}, "
                    "méthode suivante..."
                )
                break

    raise RuntimeError(
        f"« {label} » cliqué, mais l'action n'a pas abouti."
    )