import os

from playwright.sync_api import sync_playwright

from app_paths import charger_config


# Taille de page utilisée quand Chrome est invisible (équivalent d'une
# fenêtre maximisée sur un écran Full HD).
LARGEUR = 1920
HAUTEUR = 1080


def navigateur_visible():
    """
    Chrome est lancé SANS fenêtre par défaut. Pour le voir (diagnostic),
    mettre "navigateur_visible": true dans
    %LOCALAPPDATA%\\Resilia2\\config.json, ou définir la variable
    d'environnement RESILIA_VISIBLE=1.
    """

    if os.environ.get("RESILIA_VISIBLE") == "1":
        return True

    try:
        return bool(charger_config().get("navigateur_visible", False))
    except Exception:
        return False


def launch_browser(playwright):
    visible = navigateur_visible()

    if visible:
        browser = playwright.chromium.launch(
            channel="chrome",
            headless=False,
            args=["--start-maximized"],
        )

        # no_viewport permet d'utiliser la taille réelle de la fenêtre.
        context = browser.new_context(no_viewport=True)

    else:
        browser = playwright.chromium.launch(
            channel="chrome",
            headless=True,
            args=[f"--window-size={LARGEUR},{HAUTEUR}"],
        )

        # Sans fenêtre, il faut fixer la taille de la page.
        context = browser.new_context(
            viewport={"width": LARGEUR, "height": HAUTEUR}
        )

    page = context.new_page()

    return browser, context, page


def close_browser(browser):
    if browser:
        browser.close()
