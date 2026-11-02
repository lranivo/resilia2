# Resilia2 — construction et utilisation

## 1. Contenu du dossier

| Fichier | Rôle |
|---|---|
| `resilia2.py` | Point d'entrée unique (lanceur, service, worker, performance) |
| `launcher.py` | Fenêtre « Resilia Launcher » (tkinter) |
| `service.py` | Serveur web local + exécution des scénarios |
| `worker.py` | Exécute un scénario dans son propre processus |
| `performance.py` | Performance Mode (veille / activité Teams, ou VLC) |
| `scenarios.py` | Catalogue des scénarios + suivi de progression |
| `app_paths.py` | Chemins, configuration (`config.json`) |
| `web/index.html` | Interface « Centre de Résiliation » |
| `build.bat`, `requirements.txt`, `resilia.ico` | Construction de l'exécutable |
| `check_wasac.py`, `isolation_client.py`, `portal.py`, `salto_*.py`, `gel_salto.py`, `degel_salto.py`, `validation_salto.py`, `browser.py` | Modules métier (inchangés, sauf deux chemins dans `salto_common.py`) |

`browser.py` et `main.py` viennent de votre projet : copiez-les dans ce dossier.
`main.py` n'est plus nécessaire dans l'exécutable.

## 2. Construire `resilia2.exe` (sur un poste Windows)

Prérequis : Python 3.10+ (commande `py`), accès à pypi.org.

1. Placer tous les fichiers ci-dessus dans un même dossier (avec `browser.py`).
2. Double-cliquer sur `build.bat` (ou `set MODE=onedir` avant, voir ci-dessous).
3. Résultat : `dist\resilia2.exe`.

Si l'antivirus ralentit le démarrage de la version « un seul fichier », construire
en dossier : `set MODE=onedir` puis `build.bat` → `dist\resilia2\` (copier le dossier entier).

## 3. Prérequis sur les postes

- Windows 10/11, **Google Chrome** installé (utilisé par Playwright, aucun téléchargement).
- **Microsoft Excel** installé : les modules pilotent Excel par COM (`win32com`).
- Aucun droit administrateur nécessaire.

## 4. Utilisation

1. Lancer `resilia2.exe` : la fenêtre **Resilia Launcher** s'ouvre.
2. **Resilia Service → Start** : démarre le serveur et ouvre `http://localhost:8765`
   (le port réel est lu automatiquement ; il passe au suivant s'il est occupé).
3. Dans l'interface : déposer `Resil_aammjj.xlsx`, cocher les scénarios, **Lancer**.
   Plusieurs scénarios s'exécutent à la suite, dans l'ordre de la grille.
4. À la fin : **Télécharger le fichier traité**.
5. **Performance Mode → Start** avant un long traitement (voir §6).

Boutons de l'interface :
- **Arrêter** : interrompt le scénario (tue le processus et Chrome).
- **Fermer les processus Excel** : ferme uniquement les Excel *sans fenêtre visible*
  (ceux laissés par un scénario interrompu). Vos classeurs ouverts ne sont pas touchés.
- **Télécharger le modèle Excel** : classeur avec les colonnes attendues.

Menu `▼` du lanceur : ouvrir l'interface, ouvrir le dossier des données, quitter.
Icône ☀/☾ : thème clair / sombre.

## 5. Scénarios

| Interface | Module | Colonnes requises |
|---|---|---|
| Checking | `run_check_wasac_excel` | SU, Etat Résil, Injoignabilité R1, Injoignabilité R2 |
| Isolation | `run_isolation_client` | SU, Statut |
| Validation PF | `run_validation_salto` | DPS/ACT, Statut |
| Gel de gamme | `run_gel_salto` | DPS/ACT, Statut |
| Dégel | `run_degel_salto` | DPS/ACT, Statut |
| Suppression Feuillet, Checking DPS NOK | *pas encore développés* | — |

Les colonnes sont vérifiées au chargement : un scénario dont une colonne manque est grisé.
Plusieurs scénarios écrivent dans la même colonne **Statut** : la valeur finale est celle du dernier.

## 6. Performance Mode

Mode `natif` (défaut) : empêche la mise en veille et l'extinction de l'écran
(`SetThreadExecutionState`, sans droits administrateur) et, si activé, appuie
périodiquement sur la touche **F15** (sans effet) pour que Windows et Teams voient
une activité.

Mode `vlc` : lance VLC sans interface sur un fichier silencieux en boucle
(empêche la veille ; ne maintient **pas** le statut Teams). Si VLC est absent,
bascule automatiquement en mode natif.

⚠ L'activité simulée réinitialise aussi le verrouillage automatique de la session.
Ne laissez pas un poste sans surveillance dans un lieu non sécurisé et vérifiez que
cet usage est compatible avec la politique informatique de la société.
Pour ne bloquer que la veille : `"simuler_activite_teams": false`.

## 7. Configuration et données

Dossier : `%LOCALAPPDATA%\Resilia2\` (ouvrable depuis le menu `▼` du lanceur).

- `config.json` : port, thème, ouverture auto du navigateur, réglages Performance Mode
  (`mode`, `simuler_activite_teams`, `intervalle_secondes`, `chemin_vlc`).
- `travail\` : classeurs chargés (et traités).
- `logs\` : un journal par scénario, `service.log`, `performance.log`, `crash.log`.
- `diagnostics\` : captures et listes d'éléments en cas d'échec d'un clic.

## 8. Dépannage

| Symptôme | Piste |
|---|---|
| Le lanceur ne s'ouvre pas | Voir `logs\crash.log` ; le lanceur est une instance unique |
| Service bloqué sur STARTING | Voir `logs\service.log` (pare-feu / port) |
| Un scénario échoue au démarrage | Voir `logs\<date>_<scénario>.log` et le journal de l'interface |
| Excel « fantôme » après un arrêt | Bouton **Fermer les processus Excel** |
| Blocage sur un écran Salto | Dossier `diagnostics\` : envoyer le `.txt` et la capture |
| Barre de progression à 0 % | Le format des lignes de log a changé : adapter les expressions de `scenarios.py` |

## 9. Sécurité du serveur local

Le serveur n'écoute que sur `127.0.0.1`, vérifie l'en-tête `Host` (anti DNS-rebinding)
et exige un en-tête personnalisé sur les requêtes POST (anti CSRF) : une page web
quelconque ne peut pas lancer un scénario ni fermer Excel à votre insu.
