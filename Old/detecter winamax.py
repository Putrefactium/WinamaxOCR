import psutil
import time
import pytesseract
from PIL import Image, ImageEnhance, ImageOps
import mss
import win32api

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

def preprocess_image(img):
    """Prétraitement de l'image pour améliorer la reconnaissance OCR."""
    # Conversion en niveaux de gris
    gray_img = img.convert('L')

    # Augmentation du contraste
    enhancer = ImageEnhance.Contrast(gray_img)
    enhanced_img = enhancer.enhance(2)  # Ajuster le facteur d'amélioration du contraste si nécessaire

    # Inversion des couleurs
    inverted_img = ImageOps.invert(enhanced_img)

    return inverted_img

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

    with mss.mss() as sct:
        for index, monitor in enumerate(active_monitors):
            monitor_rect = {
                'left': monitor[0],
                'top': monitor[1],
                'width': monitor[2] - monitor[0],
                'height': monitor[3] - monitor[1]
            }

            # Mesure du temps de capture d'écran
            start_time = time.time()
            screenshot = sct.grab(monitor_rect)
            capture_time = time.time() - start_time
            print_with_timestamp(f"Capture d'écran {index + 1} durée : {capture_time:.2f} secondes")

            img = Image.frombytes('RGB', screenshot.size, screenshot.bgra, 'raw', 'BGRX')

            # Prétraitement de l'image
            preprocessed_img = preprocess_image(img)

            # Mesure du temps de traitement OCR
            start_time = time.time()
            text = pytesseract.image_to_string(preprocessed_img, lang='fra')
            ocr_time = time.time() - start_time
            print_with_timestamp(f"Analyse OCR {index + 1} durée : {ocr_time:.2f} secondes")

            occurrences = text.lower().count("résultat net :")
            
            if occurrences > 0:
                occurrences_per_screen[index + 1] = occurrences
                print_with_timestamp(f"Texte 'résultat net :' trouvé {occurrences} fois sur l'écran numéro {index + 1}")
            else:
                print_with_timestamp(f"Texte 'résultat net :' non trouvé sur l'écran numéro {index + 1}")

            img.close()  # Libérer les ressources de l'image après analyse

    if occurrences_per_screen:
        return occurrences_per_screen  # Retourner le dictionnaire des occurrences par écran
    else:
        return None

def main():
    """Fonction principale qui capture les écrans, vérifie la présence du texte et affiche les résultats."""
    while True:
        start_time = time.time()
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
