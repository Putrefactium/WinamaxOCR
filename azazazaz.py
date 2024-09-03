# -*- coding: utf-8 -*-

import psutil
import win32gui
import win32process
import time
import locale
from datetime import datetime
from PIL import Image, ImageEnhance
import pytesseract
import os
import mss
import pygetwindow as gw
from PyQt5.QtWidgets import QApplication, QLabel, QPushButton, QWidget
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt, QRect
import ctypes
from ctypes import wintypes
import threading
import app
import queue
import keyboard

# Chemin d'accès à l'exécutable Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Dossier pour sauvegarder les captures d'écran
capture_folder = "captures"
if not os.path.exists(capture_folder):
    os.makedirs(capture_folder)

# Chemin de l'image du bouton
button_image_path = "DLBTN.png"

# Variables globales
button_coords = None
button_window = None
exe_name = 'winamax.exe'
window_title = "Winamax"
search_text = "Sta"

# Variables pour gérer la temporisation
last_search_time = 0
search_interval_PID = 1 # 1 seconde
start_pid_timestamp = 0 
search_interval_OCR = 1/2 # 0.5 secondes
start_ocr_timestamp = 0 

# Queue pour partager les résultats entre les threads et le reste du script
result_queue = queue.Queue()

# Variables pour gérer les threads
pids = []
handles = []
showbutton = False
pid_thread = None
handle_thread = None
found_window_thread = None
threads_done = threading.Event()

# Variable pour contrôler l'arrêt du script
stop_script = threading.Event()

def fetch_handles():
    handles = get_hwnds_by_pids(pids)
    result_queue.put(('handles', handles))  # Mettre les HWND dans la queue
    threads_done.set()  # Indiquer que la récupération des HWND est terminée

def fetch_pids(exe_name):
    pids = get_processes_by_exe(exe_name)
    result_queue.put(('pids', pids))  # Mettre les PID dans la queue
    threads_done.set()  # Indiquer que la récupération des PID est terminée

def start_pid_handle_threads(exe_name):
    global pid_thread, handle_thread, threads_done, found_window_thread, start_pid_timestamp
    start_pid_timestamp = time.time()
    threads_done.clear()  # Réinitialiser l'état des threads
    pid_thread = threading.Thread(target=fetch_pids, args=(exe_name,))
    handle_thread = threading.Thread(target=fetch_handles)
    found_window_thread = threading.Thread(target=found_window, args=(handles, window_title,))
    
    pid_thread.start()
    handle_thread.start()
    found_window_thread.start()

def start_OCR_thread():
    global ocr_thread, start_ocr_timestamp
    start_ocr_timestamp = time.time()
    threads_done.clear()  # Réinitialiser l'état des threads
    ocr_thread = threading.Thread(target=search_sta_ocr, args=(handles,))

    ocr_thread.start()

def monitor_escape_key():
    while not stop_script.is_set():
        if keyboard.is_pressed('esc'):
            # print("Touche Échap pressée. Arrêt du script.")
            stop_script.set()  # Définir l'événement pour arrêter le script
        time.sleep(0.1)  # Vérifier toutes les 100 ms 

def current_month():
    # Définit la locale en français pour afficher le nom du mois en français
    locale.setlocale(locale.LC_TIME, 'fr_FR')
    
    # Obtient le nom du mois en cours
    mois = time.strftime("%B")
    
    return mois

def print_duration(start_time, step_name):
    """Affiche la durée écoulée pour une étape donnée."""
    elapsed_time = (time.time() - start_time) * 1000  # Convertir en millisecondes
    # print(f"{step_name} - Durée : {elapsed_time:.2f} ms")

def is_window_visible(hwnd):
    """ Vérifie si une fenêtre est visible """
    return win32gui.IsWindowVisible(hwnd) != 0

def get_hwnds_by_pids(pids):
    """
    Récupère une liste de handle de fenêtre (HWND) et leurs titres pour les processus spécifiés par leurs PID.

    :param pids: Liste des PID (Process IDs) des processus pour lesquels récupérer les handles de fenêtres
    :return: Liste de tuples contenant le handle de fenêtre (HWND) et le titre de la fenêtre
    """
    start_time = time.time()
    hwnd_list = []

    def enum_window_callback(hwnd, lParam):
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if pid in pids and is_window_visible(hwnd):
            hwnd_list.append((hwnd, win32gui.GetWindowText(hwnd)))

    win32gui.EnumWindows(enum_window_callback, None)
    # print_duration(start_time, "Récupération des HWNDs")
    return hwnd_list

def get_apps_pids():
    """ Obtient les PID des applications visibles dans "Apps" """
    hwnds = []

    def enum_windows_proc(hwnd, lParam):
        if is_window_visible(hwnd):
            hwnds.append(hwnd)
        return True

    user32 = ctypes.WinDLL('user32')
    user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, ctypes.c_void_p)(enum_windows_proc), 0)

    pids = set()
    for hwnd in hwnds:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if pid:
            pids.add(pid)
    
    return pids

def get_processes_by_exe(exe_name):
    """
    Récupère une liste des PID de tous les processus correspondants au nom d'un exécutable donné, 
    et les filtre pour ceux qui sont dans la section "Apps" du Task Manager.

    :param exe_name: Nom de l'exécutable à rechercher
    :return: Liste des PID (Process IDs) des processus trouvés
    """
    start_time = time.time()
    pid_list = set()
    exe_name_lower = exe_name.lower()

    # Obtenir les PID des processus visibles dans "Apps"
    apps_pids = get_apps_pids()

    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.info['pid'] in apps_pids and proc.info['name'].lower() == exe_name_lower:
                pid_list.add(proc.info['pid'])
                # Ajouter les enfants
                pid_list.update(child.pid for child in proc.children(recursive=True))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    
    # print_duration(start_time, "Récupération des PID")
    return list(pid_list)

def screen_result(hwnd):
    """
    Capture une partie spécifique de la fenêtre identifiée par `hwnd` et retourne l'image capturée.
    La dimension correspond au rectangle composant la partie " Statistiques " de Winamax, avec les résultats de la session

    :param hwnd: Handle de la fenêtre (HWND) à capturer
    :return: Image PIL de la partie capturée de la fenêtre, ou None en cas d'erreur
    """
    start_time = time.time()
    try:
        win = gw.Window(hwnd)
        rect = (win.left, win.top, win.right, win.bottom)

        height_rectifier_up = 134  # Offset début hauteur de la partie à capturer (coté haut rectangle)
        height_rectifier_low = 178  # Offset fin hauteur de la partie à capturer (coté bas rectangle)
        width_rectifier = 900  # Offset de la largeur de la partie à capturer (cotés gauche et droit rectangle)

        window_width = rect[2] - rect[0]  # Largeur de la fenêtre
        max_offset = 300  # Décalage maximum autorisé
        threshold_width = 1414  # Largeur seuil en pixel pour commencer le décalage (les bandes noires apparaissent à ce moment)

        # Calcul du décalage progressif basé sur la largeur de la fenêtre
        if window_width > threshold_width:
            extra_width = window_width - threshold_width
            additional_offset = int(min(max_offset, extra_width / 2)) # Divisé par deux car les bandes noires sont sur les deux cotés de la fenêtre
        else:
            additional_offset = 0

        capture_rect = (
            rect[0] + additional_offset,  # Ajustement progressif du côté gauche
            rect[1] + height_rectifier_up,  # Hauteur de fenêtre moins le rectifieur de hauteur
            rect[0] + width_rectifier + additional_offset,  # Ajustement progressif de la largeur
            rect[1] + height_rectifier_low  # Hauteur de fenêtre moins le rectifieur bas
        )

        if capture_rect[0] >= capture_rect[2] or capture_rect[1] >= capture_rect[3]:
            # print(f"Les coordonnées ajustées du rectangle Résultat sont invalides : {capture_rect}")
            return None

        with mss.mss() as sct:
            img = sct.grab(capture_rect)
            img_pil = Image.frombytes('RGB', img.size, img.rgb)
        # print_duration(start_time, "Capture de la partie résultat")
        return img_pil
    except Exception as e:
        # print(f"Erreur lors de la capture de la fenêtre avec HWND {hwnd}: {e}")
        return None

def capture_word_part(hwnd, height_rectifier_up=134, width_rectifier=980, max_width=520):
    """
    Capture une partie spécifique de la fenêtre identifiée par `hwnd` et retourne l'image capturée.
    La dimension correspond au rectangle à la position relative du mot " Sta ", afin de réduire la durée du traitement OCR

    :param hwnd: Handle de la fenêtre (HWND) à capturer
    :param height_rectifier_up: Décalage vertical depuis le haut de la fenêtre pour le début de la capture
    :param width_rectifier: Décalage horizontal depuis la droite de la fenêtre pour ajuster la largeur
    :param max_width: Décalage vertical depuis le bas de la fenêtre pour ajuster la hauteur
    :return: Image PIL de la partie capturée de la fenêtre, ou None en cas d'erreur
    """
    start_time = time.time()
    try:
        win = gw.Window(hwnd)
        rect = (win.left, win.top, win.right, win.bottom)

        capture_rect = (
            rect[0],
            rect[1] + height_rectifier_up,
            rect[2] - width_rectifier,
            rect[3] - max_width
        )

        if capture_rect[0] >= capture_rect[2] or capture_rect[1] >= capture_rect[3]:
            # print(f"Les coordonnées ajustées du rectangle Capture sont invalides : {capture_rect}")
            return None

        with mss.mss() as sct:
            img = sct.grab(capture_rect)
            img_pil = Image.frombytes('RGB', img.size, img.rgb)
        # print_duration(start_time, "Capture du rectangle avec le mot à chercher")
        return img_pil
    except Exception as e:
        # print(f"Erreur lors de la capture de la fenêtre avec HWND {hwnd}: {e}")
        return None

def preprocess_image(img):
    """
    Pré-traite une image en la convertissant en niveaux de gris, en ajustant le contraste,
    et en appliquant un seuillage pour binariser l'image.

    :param img: Image PIL à pré-traiter
    :return: Image PIL pré-traitée (binarisée)
    """
    start_time = time.time()
    img = img.convert('L')
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(1)
    threshold = 64
    img = img.point(lambda p: p > threshold and 255)
    # print_duration(start_time, "Pré-traitement de l'image")
    return img

def search_text_in_image(img, search_text):
    """
    Recherche un texte spécifique dans une image en utilisant l'OCR (Reconnaissance Optique de Caractères).

    :param img: Image PIL dans laquelle rechercher le texte
    :param search_text: Texte à rechercher dans l'image
    :return: Tuple (found, elapsed_time) où :
             - `found` est un booléen indiquant si le texte a été trouvé
             - `elapsed_time` est le temps écoulé pour la recherche en secondes
    """
    text = pytesseract.image_to_string(img, lang='eng')
    found = search_text in text
    return found

class ButtonWindow(QWidget):
    """
        Initialise une fenêtre sans bordure avec un bouton superposé à une image.

        :param coords: Tuple (x, y) représentant les coordonnées de la fenêtre
        :param hwnd: Handle de la fenêtre pour la référence spécifique (non utilisé directement ici)
        :param parent: Widget parent (par défaut None)
        """
    def __init__(self, coords, hwnd, parent=None):
        super(ButtonWindow, self).__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setGeometry(QRect(coords[0], coords[1], 100, 100))
        self.hwnd = hwnd  # Stocker le HWND spécifique pour cette instance

        self.label = QLabel(self)
        self.label.setPixmap(QPixmap(button_image_path))
        self.label.setGeometry(0, 0, 100, 100)

        self.button = QPushButton(self)
        self.button.setGeometry(0, 0, 100, 100)
        self.button.setFlat(True)
        self.button.setStyleSheet("background: transparent;")
        self.button.clicked.connect(self.on_button_click)

    def on_button_click(self):
        """
        Gère les clics sur le bouton. Cette méthode est appelée lorsque le bouton est cliqué.

        - Affiche un message dans la console indiquant que le bouton a été cliqué.
        - Vérifie si le handle de la fenêtre (HWND) est défini.
        - Si oui, tente de capturer l'image de la fenêtre spécifiée par l'HWND et de l'enregistrer.
        - Sinon, affiche un message d'erreur indiquant que l'HWND n'est pas défini.
    """
        # print("Bouton cliqué")
        if self.hwnd:  # Utiliser le HWND stocké
            # print(f"Tentative de capture pour la fenêtre avec HWND : {self.hwnd}")
            capture_image_and_save(self.hwnd)
        else:
            print("Erreur : HWND non défini pour cette fenêtre.")

def capture_image_and_save(hwnd):
    """
    Capture une image de la fenêtre spécifiée par `hwnd`, enregistre l'image dans un fichier JPG avec un nom basé sur l'horodatage,
    et affiche un message de confirmation ou d'erreur.

    :param hwnd: Handle de la fenêtre dont l'image doit être capturée
    """
    result_jpg = screen_result(hwnd)
    if result_jpg:
        new_month_folder = current_month()
        # S'assurer que le nom du dossier est correctement encodé
        new_month_folder = new_month_folder.encode('utf-8').decode('utf-8')
        full_path = os.path.join(capture_folder, new_month_folder)
        # print(f"chemin : {full_path}")

        if not os.path.exists(full_path):
            os.makedirs(full_path)

        timestamp = datetime.now().strftime("%d_%m_%Y")
        file_path = os.path.join(full_path, f"{timestamp}.jpg")
        result_jpg.save(file_path, "JPEG")
        # print(f"Image sauvegardée à : {file_path}")
    else:
         print("Erreur lors de la capture de l'image.")
 
def show_button(coords, hwnd):
    """
    Affiche une fenêtre avec un bouton à la position spécifiée. Si la fenêtre est déjà affichée, met à jour sa position et son HWND.

    :param coords: Tuple (x, y) représentant les coordonnées où la fenêtre doit être affichée
    :param hwnd: Handle de la fenêtre dont l'image sera capturée lorsque le bouton est cliqué
    """
    global button_window
    if not button_window:
        button_window = ButtonWindow(coords, hwnd)
        button_window.show()
    else:
        x, y = int(coords[0]), int(coords[1])
        button_window.setGeometry(QRect(x, y, 100, 100))
        button_window.raise_()
        button_window.activateWindow()
        button_window.hwnd = hwnd  # Mettre à jour le HWND si nécessaire

def hide_button():
    """
    Cache la fenêtre du bouton si elle est actuellement affichée.
    Ferme la fenêtre et libère la référence globale à cette instance de ButtonWindow.
    """
    global button_window
    if button_window:
        button_window.close()
        button_window = None

def calculate_percentage_and_position(x, y, width, reference_width=1414, max_percentage=95, min_percentage=82):
    """
    Calcule le pourcentage de la largeur pour la position du bouton et la position en pixels.
    :param x: Coordonnée X du coin supérieur gauche de la fenêtre
    :param y: Coordonnée Y du coin supérieur gauche de la fenêtre
    :param width: Largeur de la fenêtre
    :param reference_width: Largeur de référence pour commencer à réduire le pourcentage
    :param max_percentage: Pourcentage maximum pour une petite fenêtre
    :param min_percentage: Pourcentage minimum pour une grande fenêtre
    :return: (button_x, button_y) coordonnées du bouton
    """
    if width <= reference_width:
        percentage = max_percentage
    else:
        scale = (min_percentage - max_percentage) / (2000 - reference_width)
        percentage = max_percentage + scale * (width - reference_width)

    button_x = int(x + ((percentage * width) / 100))
    button_y = int(y + 108)

    return button_x, button_y

def found_window(handles, window_title):

    foundwindow = False
        
    for hwnd, title in handles:
        if title == window_title:
            foundwindow = True
            if win32gui.IsWindowVisible(hwnd):
                # Si la fenêtre existe et n'est pas réduite on récupère sa position et dimension
                rect = win32gui.GetWindowRect(hwnd)
                x, y, x1, y1 = rect
                width = x1 - x
                height = y1 - y
                # Si elle est en -32k -32k elle est réduite
                if x == -32000 and y == -32000:
                    # print(f"[{current_time}] HWND : {hwnd}, Titre : {title}, Visible : Non (Réduite)")
                    hide_button()
                else:
                    print(f"[{current_time}] HWND : {hwnd}, Titre : {title}, Visible: Oui, Coordonnées : ({x}, {y}, {x1}, {y1}), Taille : ({width}x{height})")
                    
    result_queue.put(('found_window', foundwindow))      

def search_sta_ocr(handles):

    global showbutton
    global search_text
    global window_title

    for hwnd, title in handles:
        if title == window_title:

            img = capture_word_part(hwnd)  
                                
            if img:
                found = search_text_in_image(img, search_text)
                if found:
                    print("Texte Sta trouvé")
                else:
                     print(f"[{current_time}] Texte '{search_text}' non trouvé dans la fenêtre.")
                showbutton = found
                
app = QApplication([]) # Initialise l'environnement Qt

start_pid_handle_threads(exe_name) # Initialiser les threads pour HWND et PID pour commencer le traitement en arrière-plan
start_OCR_thread() # Initialiser le thread de reconnaissance OCR pour commencer le traitement en arrière-plan

# Démarrer le thread pour surveiller la touche Échap
escape_thread = threading.Thread(target=monitor_escape_key)
escape_thread.start()

while not stop_script.is_set():
    foundwindow = True
    current_timestamp = time.time()
    current_time = datetime.now().strftime("%H:%M:%S")

    # Vérifier l'état des threads
    if threads_done.is_set():
        while not result_queue.empty():
            data_type, data = result_queue.get()
            if data_type == 'pids':
                pids = data
                # print(f"PIDS: {pids}")
            elif data_type == 'handles':
                handles = data
                # # print(f"HWND : {handles}")   
            elif data_type == 'found_window':
                foundwindow = data
                # print(f"found_window = {foundwindow}")
                       
    if foundwindow:

        for hwnd, title in handles:
            if title == window_title:

                rect = win32gui.GetWindowRect(hwnd)

                x, y, x1, y1 = rect
                width = x1 - x
                height = y1 - y
                # Si elle est en -32k -32k elle est réduite
                if x == -32000 and y == -32000:
                    # print(f"[{current_time}] HWND : {hwnd}, Visible : Non (Réduite)")
                    hide_button()
                    # print("bouton caché")
                else:
                    # print(f"[{current_time}] HWND : {hwnd}, Visible: Oui, Coordonnées : ({x}, {y}, {x1}, {y1}), Taille : ({width}x{height})")
                    button_x, button_y = calculate_percentage_and_position(x,y,width)
                    
                    if showbutton:
                        show_button((button_x, button_y), hwnd)
                    else:
                        hide_button()
                                        
    else:
        hide_button()
        # print("bouton caché 404")

# Redémarrer les threads pour obtenir des informations mises à jour pour le PID et HWND
    if current_timestamp - start_pid_timestamp >= (search_interval_PID):
        start_pid_handle_threads(exe_name)
        print("start PID HWND thread")

    if current_timestamp - start_ocr_timestamp >= (search_interval_OCR):
        start_OCR_thread()
        print("start OCR thread")

    # Réinitialiser l'état des threads
    threads_done.clear()

    # print("===" * 30)
    # print_duration(iteration_start_time, "Temps total Boucle principale")
    app.processEvents()
    # time.sleep(1/165)

# Nettoyer les threads avant de quitter
pid_thread.join()
handle_thread.join()
ocr_thread.join()
escape_thread.join()
