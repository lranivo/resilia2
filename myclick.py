
import time
import pyautogui

DELAI = 600  # 10 minutes

print("Clic automatique programmé.")
print("Ne déplacez pas la souris.")
print("Appuyez sur CTRL+C pour annuler.")

try:
    time.sleep(DELAI)

    pyautogui.click()
    time.sleep(1)
    pyautogui.click()
    time.sleep(1)
    pyautogui.click()

    print("Clic effectué !")

except KeyboardInterrupt:
    print("Opération annulée.")
