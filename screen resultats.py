import win32gui
import win32con
import win32api
from PIL import Image, ImageWin, ImageGrab
import pytesseract
import time

# Chemin vers l'exécutable Tesseract-OCR (à adapter selon votre installation)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def verifier_fenetre(titre_fenetre):
    try:
        def enum_callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd) and titre_fenetre.lower() in win32gui.GetWindowText(hwnd).lower():
                windows.append(hwnd)

        hwnds = []
        win32gui.EnumWindows(enum_callback, hwnds)
        if hwnds:
            return hwnds[0]  # Retourne le premier handle trouvé
        return None
    except Exception as e:
        print(f"Erreur lors de la vérification de la fenêtre : {e}")
        return None

def capturer_fenetre(hwnd):
    try:
        # Obtenir les dimensions de la fenêtre
        rect = win32gui.GetWindowRect(hwnd)
        left, top, right, bottom = rect

        # Capturer l'écran de la fenêtre
        screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
        return screenshot
    except Exception as e:
        print(f"Erreur lors de la capture de la fenêtre : {e}")
        return None

def detecter_texte(image, texte_recherche):
    try:
        # Utiliser pytesseract pour extraire le texte de l'image
        texte_extrait = pytesseract.image_to_string(image, lang='fra')  # Utiliser le langage français
        return texte_recherche in texte_extrait
    except Exception as e:
        print(f"Erreur lors de la détection du texte : {e}")
        return False

def dessiner_image_sur_fenetre(hwnd, image_path, position_relative):
    try:
        # Charger l'image PNG à superposer
        overlay_image = Image.open(image_path).convert("RGBA")
        
        # Obtenir le contexte de l'appareil pour la fenêtre
        hdc_window = win32gui.GetWindowDC(hwnd)
        if not hdc_window:
            print("Erreur lors de l'obtention du contexte de périphérique de la fenêtre.")
            return
        
        # Créer un contexte de périphérique compatible pour la fenêtre
        hdc_mem = win32gui.CreateCompatibleDC(hdc_window)
        if not hdc_mem:
            print("Erreur lors de la création du contexte de périphérique compatible.")
            win32gui.ReleaseDC(hwnd, hdc_window)
            return
        
        # Obtenir la taille de l'image PNG
        width, height = overlay_image.size
        
        # Créer un bitmap compatible avec la fenêtre
        hbitmap = win32gui.CreateCompatibleBitmap(hdc_window, width, height)
        if not hbitmap:
            print("Erreur lors de la création du bitmap compatible.")
            win32gui.DeleteDC(hdc_mem)
            win32gui.ReleaseDC(hwnd, hdc_window)
            return
        
        # Sélectionner le bitmap dans le contexte de mémoire
        hbm_old = win32gui.SelectObject(hdc_mem, hbitmap)
        
        # Convertir l'image PNG en un format compatible avec GDI
        bitmap = ImageWin.Dib(overlay_image)
        
        # Dessiner l'image PNG sur le contexte de périphérique compatible
        bitmap.draw(hdc_mem, (0, 0, width, height))
        
        # Dessiner le bitmap sur la fenêtre
        x_dest, y_dest = position_relative
        result = win32gui.BitBlt(hdc_window, x_dest, y_dest, width, height, hdc_mem, 0, 0, win32con.SRCCOPY)
        
        if result == 0:
            error_code = win32api.GetLastError()
            print(f"Erreur lors de la copie de l'image sur la fenêtre. Code d'erreur: {error_code} ({win32api.FormatMessage(error_code).strip()})")
        else:
            print("Image copiée avec succès sur la fenêtre.")
        
        # Nettoyer
        win32gui.SelectObject(hdc_mem, hbm_old)
        win32gui.DeleteObject(hbitmap)
        win32gui.DeleteDC(hdc_mem)
        win32gui.ReleaseDC(hwnd, hdc_window)
        
    except Exception as e:
        print(f"Erreur lors de la superposition de l'image sur la fenêtre : {e}")



def surveiller_fenetre(titre_fenetre, texte_recherche, image_path, position_relative):
    try:
        hwnd = verifier_fenetre(titre_fenetre)
        
        if not hwnd:
            print(f"Aucune fenêtre avec le titre '{titre_fenetre}' n'a été détectée.")
            return

        print(f"Surveillance de la fenêtre '{titre_fenetre}' pour le texte '{texte_recherche}'.")

        while True:
            try:
                # Capturer la fenêtre
                screenshot = capturer_fenetre(hwnd)
                
                if screenshot:
                    # Vérifier la présence du texte
                    if detecter_texte(screenshot, texte_recherche):
                        print(f"Le texte '{texte_recherche}' est actuellement détecté dans la fenêtre '{titre_fenetre}'.")
                        # Dessiner l'image PNG sur la fenêtre
                        dessiner_image_sur_fenetre(hwnd, image_path, position_relative)
                    else:
                        print(f"Le texte '{texte_recherche}' n'est pas détecté. Re-vérification dans 2 secondes.")
                
                # Attendre avant de vérifier à nouveau
                time.sleep(2)
            except Exception as e:
                print(f"Erreur lors de la capture ou de l'analyse de la fenêtre : {e}")
                time.sleep(2)  # Attendre avant de réessayer
    except Exception as e:
        print(f"Erreur générale dans la surveillance de la fenêtre : {e}")

# Nom de la fenêtre à vérifier
titre_fenetre = "Winamax"

# Texte à détecter
texte_recherche = "Résultat net : "

# Chemin vers l'image PNG à superposer
image_path = r"C:\Users\Putré\Documents\code test\python winamax\creative_tech_school_logo.png"

# Position relative de l'image superposée (x, y) par rapport au coin supérieur gauche de la fenêtre
position_relative = (100, 200)  # Adaptez cette valeur à vos besoins

# Lancer la surveillance
surveiller_fenetre(titre_fenetre, texte_recherche, image_path, position_relative)
