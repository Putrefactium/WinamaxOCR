import psutil
import time
import pytesseract
from PIL import Image, ImageEnhance, ImageOps, ImageFilter
import mss
import win32api
import threading

# Chemin vers tesseract.exe
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

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

def preprocess_image(image):
    """Applique des techniques de prétraitement à l'image pour améliorer la détection OCR."""
    # Simple grayscale conversion and contrast enhancement
    image = ImageOps.grayscale(image)
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(1.5)
    return image

def process_monitor(monitor_rect, index, occurrences_per_screen):
    with mss.mss() as sct:
        # Capture d'écran
        screenshot = sct.grab(monitor_rect)
        img = Image.frombytes('RGB', screenshot.size, screenshot.bgra, 'raw', 'BGRX')
        
        # Prétraitement de l'image
        img = preprocess_image(img)
        
        # Analyse OCR
        start_time = time.time()
        custom_config = r'--oem 3 --psm 6'  # OCR Engine Mode et Page Segmentation Mode
        text = pytesseract.image_to_string(img, config=custom_config, lang='fra')
        ocr_time = time.time() - start_time
        print_with_timestamp(f"Analyse OCR {index + 1} durée : {ocr_time:.2f} secondes")

        occurrences = text.lower().count("résultat net :")
        
        if occurrences > 0:
            occurrences_per_screen[index + 1] = occurrences
            print_with_timestamp(f"Texte 'résultat net :' trouvé {occurrences} fois sur l'écran numéro {index + 1}")
        else:
            print_with_timestamp(f"Texte 'résultat net :' non trouvé sur l'écran numéro {index + 1}")

        img.close()  # Libérer les ressources de l'image après analyse

def capture_and_check_keyword():
    """Capture des écrans actifs et vérifie la présence du texte spécifique sur tous les écrans."""
    if not is_winamax_running():
        print_with_timestamp("Winamax n'est pas en cours d'exécution.")
        return None

    active_monitors = get_active_monitors()
    
    if not active_monitors:
        print_with_timestamp("Aucun écran actif trouvé.")
        return None

    occurrences_per_screen = {}
    threads = []

    for index, monitor in enumerate(active_monitors):
        monitor_rect = {
            'left': monitor[0],
            'top': monitor[1],
            'width': monitor[2] - monitor[0],
            'height': monitor[3] - monitor[1]
        }
        
        thread = threading.Thread(target=process_monitor, args=(monitor_rect, index, occurrences_per_screen))
        threads.append(thread)
        thread.start()

    for thread in threads:
        thread.join()

    if occurrences_per_screen:
        return occurrences_per_screen  # Retourner le dictionnaire des occurrences par écran
    else:
        return None

def main():
    """Fonction principale qui capture les écrans, vérifie la présence du texte et affiche les résultats."""
    while True:
        start_time = time.time()

        # Convertir les secondes en un tuple struct_time
        current_local_time = time.localtime(start_time)

        # Extraire les fractions de seconde (millisecondes)
        milliseconds = int((start_time - int(start_time)) * 1000)

        # Formater l'heure actuelle avec les millisecondes
        formatted_time = time.strftime("%Y-%m-%d %H:%M:%S", current_local_time) + f".{milliseconds:03d}"

        print_with_timestamp(formatted_time)

        result = capture_and_check_keyword()
        processing_time = time.time() - start_time
        print_with_timestamp(f"Durée totale du traitement : {processing_time:.2f} secondes")
        
        if result:
            for screen, count in result.items():
                print_with_timestamp(f"Écran {screen} : {count} occurrence(s) de 'résultat net :'")
        else:
            print_with_timestamp("Texte 'résultat net :' non trouvé sur aucun écran.")
        
        # Ajouter une ligne de séparation pour la lisibilité des logs
        print_with_timestamp("Fin de l'itération")
        print("=" * 50)
        
        time.sleep(1)  # Attendre avant la prochaine vérification

if __name__ == "__main__":
    main()
