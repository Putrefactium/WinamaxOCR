import psutil
import time
import easyocr
from PIL import Image
import numpy as np
import mss
import win32api
from concurrent.futures import ThreadPoolExecutor, as_completed
import torch

# Initialiser EasyOCR avec GPU (vérifier la disponibilité du GPU)
reader = easyocr.Reader(['fr'], gpu=torch.cuda.is_available())

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

def process_monitor(monitor_rect, index):
    """Capture une image de l'écran, effectue l'OCR, et compte les occurrences du texte."""
    occurrences = 0
    with mss.mss() as sct:
        # Capture d'écran
        screenshot = sct.grab(monitor_rect)
        img = Image.frombytes('RGB', screenshot.size, screenshot.bgra, 'raw', 'BGRX')
        
        # Convertir l'image PIL en tableau numpy
        img_np = np.array(img)

        # Optionnel : Sauvegarder l'image pour vérification
        img.save(f"screenshot_{index}.png")
        
        # Analyse OCR avec EasyOCR
        start_time = time.time()
        result = reader.readtext(img_np, detail=1)
        ocr_time = time.time() - start_time
        print_with_timestamp(f"Analyse OCR {index + 1} durée : {ocr_time:.2f} secondes")

        # Vérifier le contenu des résultats OCR
        text_results = [text[1] for text in result]
        combined_text = ' '.join(text_results).lower()

        # Log pour débogage
        print_with_timestamp(f"Texte extrait (écran {index + 1}) : {combined_text}")

        # Compter les occurrences du texte recherché
        occurrences = combined_text.count("résultat net :")

        if occurrences > 0:
            print_with_timestamp(f"Texte 'résultat net :' trouvé {occurrences} fois sur l'écran numéro {index + 1}")
        else:
            print_with_timestamp(f"Texte 'résultat net :' non trouvé sur l'écran numéro {index + 1}")
        
        img.close()  # Libérer les ressources de l'image après analyse
    
    return index + 1, occurrences

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
    
    with ThreadPoolExecutor(max_workers=len(active_monitors)) as executor:
        futures = []
        for index, monitor in enumerate(active_monitors):
            monitor_rect = {
                'left': monitor[0],
                'top': monitor[1],
                'width': monitor[2] - monitor[0],
                'height': monitor[3] - monitor[1]
            }
            futures.append(executor.submit(process_monitor, monitor_rect, index))

        for future in as_completed(futures):
            screen, occurrences = future.result()
            occurrences_per_screen[screen] = occurrences
    
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
