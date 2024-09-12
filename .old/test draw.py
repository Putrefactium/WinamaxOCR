import win32gui
import win32con
import win32api
from PIL import Image, ImageWin
import time

def obtenir_fenetre_active():
    hwnd = win32gui.GetForegroundWindow()
    return hwnd

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

def surveiller_fenetre(titre_fenetre, image_path, position_relative):
    while True:
        hwnd = obtenir_fenetre_active()
        window_title = win32gui.GetWindowText(hwnd)
        
        if titre_fenetre.lower() in window_title.lower():
            print(f"Fenêtre '{titre_fenetre}' trouvée.")
            dessiner_image_sur_fenetre(hwnd, image_path, position_relative)
        else:
            print(f"Aucune fenêtre '{titre_fenetre}' active trouvée. Réessai dans 2 secondes.")
        
        # Attendre avant de réessayer
        time.sleep(2)

# Nom de la fenêtre à vérifier
titre_fenetre = "Winamax"

# Chemin vers l'image PNG à superposer
image_path = r"C:\Users\Putré\Documents\code test\python winamax\creative_tech_school_logo.png"

# Position relative de l'image superposée (x, y) par rapport au coin supérieur gauche de la fenêtre
position_relative = (100, 200)  # Adaptez cette valeur à vos besoins

# Lancer la surveillance
surveiller_fenetre(titre_fenetre, image_path, position_relative)
