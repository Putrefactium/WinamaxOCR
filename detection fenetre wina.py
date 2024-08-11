import psutil
import time
import threading
from PIL import Image
import mss
import pytesseract
import win32api
import os

# Couleur spécifique du bandeau (en RGB)
TARGET_COLOR = (28, 29, 33)  # Correspond à #1c1d21

# Facteur de réduction de résolution
RESIZE_FACTOR = 0.25

# Chemin d'installation de Tesseract OCR (à adapter selon votre installation)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Répertoire pour sauvegarder les captures d'écran
OUTPUT_DIR = "captures"

def current_timestamp():
    """Renvoie un timestamp formaté pour les logs."""
    return time.strftime("%Y-%m-%d %H:%M:%S")

def print_with_timestamp(message):
    """Affiche un message avec un timestamp."""
    print(f"[{current_timestamp()}] {message}")

def is_winamax_running():
    """Vérifie si le processus Winamax est en cours d'exécution."""
    for proc in psutil.process_iter(['pid', 'name', 'exe']):
        try:
            name = proc.info['name']
            exe = proc.info['exe']
            if name and name.lower() == "winamax.exe":
                return True
            if exe and "winamax.exe" in exe.lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False

def get_active_monitors():
    """Récupère les informations sur les écrans actifs."""
    monitors = []
    for monitor in win32api.EnumDisplayMonitors():
        monitor_info = monitor[2]  # monitor[2] contient les informations du rectangle de l'écran
        monitors.append(monitor_info)
    return monitors

def find_window_position(img, monitor_index, monitor_rect):
    """Recherche la position de la fenêtre avec le bandeau de couleur spécifique."""
    width, height = img.size
    top_y = None
    top_x = None

    # Début de la recherche de la couleur du bandeau supérieur
    start_time = time.time()

    for y in range(height):
        for x in range(width):
            if img.getpixel((x, y)) == TARGET_COLOR:
                top_y = y / RESIZE_FACTOR
                top_x = x / RESIZE_FACTOR
                print_with_timestamp(f"Pixel gris trouvé en x = {top_x} et y = {top_y} sur l'écran {monitor_index + 1}")
                break
        if top_y is not None:
            break

    # Fin de la recherche de la couleur
    color_search_time = time.time() - start_time
    print_with_timestamp(f"Recherche du bandeau sur l'écran {monitor_index + 1} durée : {color_search_time:.2f} secondes")

    if top_y is not None and top_x is not None:
        # Adapter les coordonnées pour l'écran spécifique
        capture_rect = {
            'left': int(top_x) + monitor_rect['left'],
            'top': int(top_y) + monitor_rect['top'],
            'width': 1000,
            'height': 500,
        }

        print_with_timestamp(f"Zone de capture définie : {capture_rect}")

        # Capture d'écran de la zone détectée
        with mss.mss() as sct:
            screenshot = sct.grab(capture_rect)
            img_full = Image.frombytes('RGB', screenshot.size, screenshot.bgra, 'raw', 'BGRX')

            # Sauvegarde de la capture en JPG
            if not os.path.exists(OUTPUT_DIR):
                os.makedirs(OUTPUT_DIR)
            
            file_path = os.path.join(OUTPUT_DIR, f"capture_{monitor_index + 1}.jpg")
            img_full.save(file_path, "JPEG")
            print_with_timestamp(f"Capture sauvegardée sous : {file_path}")

            # Analyse avec Tesseract
            text = pytesseract.image_to_string(img_full, lang='fra')  # Utiliser 'fra' pour le français
            # print_with_timestamp(f"Texte détecté : {text}")

            if "Statistiques" in text:
                print_with_timestamp("Le mot 'Statistiques' a été trouvé.")
            else:
                print_with_timestamp("Le mot 'Statistiques' n'a pas été trouvé.")
    else:
        print_with_timestamp(f"Fenêtre non détectée sur l'écran {monitor_index + 1}.")

def capture_and_detect_window():
    """Capture les écrans actifs et détecte les fenêtres par couleur de bandeau."""
    if not is_winamax_running():
        print_with_timestamp("Winamax n'est pas en cours d'exécution.")
        return None

    active_monitors = get_active_monitors()

    if not active_monitors:
        print_with_timestamp("Aucun écran actif trouvé.")
        return None

    with mss.mss() as sct:
        threads = []

        for index, monitor in enumerate(active_monitors):
            monitor_rect = {
                'left': monitor[0],
                'top': monitor[1],
                'width': monitor[2] - monitor[0],
                'height': monitor[3] - monitor[1]
            }

            # Capture d'écran
            start_time = time.time()
            screenshot = sct.grab(monitor_rect)
            capture_time = time.time() - start_time
            print_with_timestamp(f"Capture d'écran {index + 1} durée : {capture_time:.2f} secondes")

            img = Image.frombytes('RGB', screenshot.size, screenshot.bgra, 'raw', 'BGRX')

            # Réduction de la résolution de l'image
            resized_img = img.resize((int(img.width * RESIZE_FACTOR), int(img.height * RESIZE_FACTOR)))

            # Créer et lancer un thread pour traiter chaque écran
            thread = threading.Thread(
                target=find_window_position,
                args=(resized_img, index, monitor_rect)
            )
            threads.append(thread)
            thread.start()

        # Attendre la fin de tous les threads
        for thread in threads:
            thread.join()

def main():
    """Fonction principale qui capture les écrans, vérifie la présence de la fenêtre et affiche les résultats."""
    while True:
        start_time = time.time()
        capture_and_detect_window()
        processing_time = time.time() - start_time
        print_with_timestamp(f"Durée totale du traitement : {processing_time:.2f} secondes")

        # Ajouter une ligne de séparation pour la lisibilité des logs
        print_with_timestamp("Fin de l'itération")
        print("=" * 50)

        time.sleep(1)  # Attendre avant la prochaine vérification

if __name__ == "__main__":
    main()
