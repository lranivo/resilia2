#Ouverture du portail, attente de « Mes favoris », puis saisie de SALTO au clavier et récupération du nouvel onglet.
import time

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
SALTO_DOMAIN = "salto.sso.infra.ftgroup"

PORTAL_URL = (
    "https://mytools-portal.sso.infra.ftgroup/"
    "binbeemytools/Home/Portfolio"
)

SALTO_URL = (
    "https://salto.sso.infra.ftgroup/prweb/SSOServlet/"
    "app/Salto/kmt0ufT612G4n4NNnObiyscy7-juNVt5*/!STANDARD"
)


# Réglages : le portail doit rester inchangé STABILITE_PORTAIL_MS avant
# la saisie clavier (au plus STABILITE_PORTAIL_MAX_MS, comme l'ancienne
# pause fixe de 3 s). Augmenter si "SALTO" est parfois mal saisi.
STABILITE_PORTAIL_MS = 1200
STABILITE_PORTAIL_MAX_MS = 3000


def _signature_portail(page):
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


def attendre_portail_stable(page):
    """Remplace la pause fixe de 3 s : attend que le portail ne bouge plus."""

    deadline = time.monotonic() + STABILITE_PORTAIL_MAX_MS / 1000
    derniere = None
    depuis = None

    while time.monotonic() < deadline:
        signature = _signature_portail(page)
        maintenant = time.monotonic()

        if signature != derniere:
            derniere = signature
            depuis = maintenant

        elif maintenant - depuis >= STABILITE_PORTAIL_MS / 1000:
            return

        page.wait_for_timeout(150)


def open_portal(page):
    print("[PORTAL] Ouverture du portail...")
    page.goto(
        PORTAL_URL,
        wait_until="domcontentloaded",
        timeout=120000
    )

    page.get_by_text(
        "Mes favoris", exact=False
    ).first.wait_for(
        state="visible",
        timeout=120000
    )
    attendre_portail_stable(page)

    print("[PORTAL] Portail détecté.")


def open_salto(context, portal_page):
    print("[PORTAL] Préparation de l'ouverture de Salto...")

    existing_pages = set(context.pages)
    portal_page.bring_to_front()

    try:
        with context.expect_page(timeout=60000) as page_info:
            portal_page.keyboard.type("SALTO", delay=100)
            portal_page.keyboard.press("Enter")
            # La sortie du bloc attend la création de l'onglet.

        # L'onglet a été créé, mais son URL peut être vide au départ.
        salto_page = page_info.value
        print("[SALTO] Nouvel onglet détecté, attente de navigation...")

    except PlaywrightTimeoutError:
        new_pages = [
            p for p in context.pages
            if p not in existing_pages and not p.is_closed()
        ]

        if not new_pages:
            raise RuntimeError(
                "Aucun nouvel onglet créé par le portail."
            )

        salto_page = new_pages[-1]

    # Attendre que la page quitte son URL initiale vide.
    try:
        salto_page.wait_for_url(
            "**/*",
            wait_until="commit",
            timeout=60000
        )
    except PlaywrightTimeoutError:
        pass

    # Laisser le temps à la navigation de se stabiliser.
    try:
        salto_page.wait_for_load_state(
            "domcontentloaded",
            timeout=60000
        )
    except PlaywrightTimeoutError:
        print("[SALTO] DOM toujours en chargement.")

    print(f"[SALTO] URL actuelle : {salto_page.url!r}")

    if not salto_page.url:
        # Dernière attente conditionnelle : ne pas utiliser
        # une URL directe qui contournerait le SSO.
        try:
            salto_page.wait_for_function(
                "() => window.location.href !== 'about:blank'",
                timeout=30000
            )
        except PlaywrightTimeoutError:
            raise RuntimeError(
                "L'onglet Salto est resté sur about:blank. "
                "La navigation initiée par le portail n'a pas abouti."
            )

    print(f"[SALTO] URL finale : {salto_page.url}")

    if SALTO_DOMAIN not in salto_page.url:
        raise RuntimeError(
            f"L'onglet créé par le portail n'est pas sur Salto : "
            f"{salto_page.url}"
        )

    salto_page.bring_to_front()
    print("[SALTO] Navigation du portail confirmée.")

    return salto_page