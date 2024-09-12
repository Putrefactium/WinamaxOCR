import psutil
import win32gui
import win32process
import time
import locale
from datetime import datetime
from PIL import Image
import pytesseract 
import os
import mss
import pygetwindow as gw
from PyQt5.QtWidgets import QApplication, QLabel, QPushButton, QWidget
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt, QRect
import threading
import queue
import keyboard
import logging

# Global variables #

winamax_proc_name = "Winamax.exe"
winamax_window_name = "Winamax"
winamax_string_searched = "Stat"
button_image_path = r"Assets\DLBTN.png"
x_coord_window = 0
y_coord_window = 0
string_found = None
button_window_var = None
start_ocr_timestamp = 0
search_interval_OCR = 1/2 # Interval in seconds to start the OCR thread

threads_done = threading.Event()
result_queue = queue.Queue() # Queue to share results between threads and the rest of the script

# Enable verbose logging
VERBOSE_LOGGING = False

# Configure logging
logging.basicConfig(
    level=logging.DEBUG if VERBOSE_LOGGING else logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%d-%m-%Y | %H:%M:%S'
)

# Path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Path to the folder where the captures will be saved
capture_folder = "captures"
if not os.path.exists(capture_folder):
    os.makedirs(capture_folder)

def check_wmx_proc_alive_():
    """
    Check if the "winamax.exe" process is alive.
    Returns a set of unique process names found.
    """

    found_wmx_proc = set() # Set to store unique process names
    found_wmx_proc.clear() # Clear the set before checking again

    processes = psutil.process_iter() # Get a list of all running processes

    # Check if any process has the name "winamax.exe"
    for process in processes:
        if process.name() == winamax_proc_name:
            found_wmx_proc.add(process.name())

    # Log the unique process names
    for process_name in found_wmx_proc:
        logging.debug(f"Winamax process found : {process_name}")

    if not found_wmx_proc:
        logging.debug("Winamax process not found.")

    return found_wmx_proc

def get_wmx_pids_():
    """
    Retrieves the PIDs of all running instances of "winamax.exe" process.
    Returns:
        list: A list of integers representing the PIDs of the "winamax.exe" processes.
    """

    processes = psutil.process_iter() # Get a list of all running processes
    wmx_pids = [] # List to store the PIDs of "winamax.exe" processes

    # Check if any process has the name "winamax.exe" and if it's the case, add the PID to the list
    for process in processes:
        if process.name() == winamax_proc_name:
            wmx_pids.append(process.pid)

    return wmx_pids

def get_wmx_hwnd_and_title_(pids):
    """
    Retrieves a list of window handles (hwnd) and their corresponding window titles
    for the specified pids.
    Parameters:
    - pids (list): A list of process IDs (int) for which window handles need to be retrieved.
    Returns:
    - hwnd_list (list): A list of tuples containing the window handle (hwnd) and the window title (str)
      for each window associated with the specified process IDs.
    Note:
    - This function uses the win32gui and win32process modules from the pywin32 library.
    - The window handle (hwnd) is a unique identifier for a window in the Windows operating system.
    - The window title is the text displayed in the title bar of a window.
    """

    hwnd_list = []

    def enum_window_callback(hwnd, _):
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if pid in pids:
            hwnd_list.append((hwnd, win32gui.GetWindowText(hwnd))) # Append the tuple (hwnd, title) to the list if the PID matches the specified PIDs in the list
    
    win32gui.EnumWindows(enum_window_callback, None)

    return hwnd_list

def filter_hwnd_list_(hwnd_list, winamax_window_name):
    """
    Filter the hwnd_list to only include hwnd with title equal to winamax_window_name.
    Parameters:
    - hwnd_list (list): A list of tuples containing the window handle (hwnd) and the window title (str).
    - winamax_window_name (str): The window title to filter for.
    Returns:
    - hwnd_list_filtered (list): A filtered list of tuples containing the window handle (hwnd) and the window title (str).
    """

    try:
        hwnd_list_filtered = [(hwnd, title) for hwnd, title in hwnd_list if title == winamax_window_name] # Filter the list based on the window title matching the specified title
    except Exception as e:
        logging.error(f"Error occurred during filtering: {e}")
        hwnd_list_filtered = []

    return hwnd_list_filtered

def get_window_position_and_dimensions_(hwnd):
    """
    Retrieves the position and dimensions of a window specified by its hwnd.
    Parameters:
    - hwnd (int): The window handle (hwnd) of the window.
    Returns:
    - tuple: A tuple containing the x and y coordinates of the top-left corner of the window,
                as well as the width and height of the window.
    """

    rect = win32gui.GetWindowRect(hwnd) # (left, top, right, bottom)
    x, y, x1, y1 = rect
    width = x1 - x
    height = y1 - y

    return x, y, width, height

def capture_window_region_(x, y):
    """
    Captures the region of the window specified by its hwnd and coordinates.
    Parameters:
    - x (int): The x-coordinate of the top-left corner of the region.
    - y (int): The y-coordinate of the top-left corner of the region.
    - We use hard coded offsets as we know the exact text position we are looking for to form the region.
    Returns:
    - image: The captured image of the region.
    """

    region = (x, y + 134, x + 400, y + 184)

    # Capture the region of the window
    with mss.mss() as sct:
            img = sct.grab(region)
            img_pil = Image.frombytes('RGB', img.size, img.rgb)
    
    return img_pil

def OCR_string_search_(img, search_text):
    """
    Search for a specific text in an image using OCR (Optical Character Recognition).

    :param img: PIL image to search the text in
    :param search_text: Text to search in the image
    :return: True if the text is found in the image, False otherwise
    :return: string_found value to the global variable
    """

    global string_found 

    logging.debug(f"Searching for text '{search_text}' in the image.")

    try:
        text = pytesseract.image_to_string(img, lang='eng')
        found = search_text in text
        logging.debug(f"Text found: {found}")
        string_found = found # Update the global variable for the main loop

        threads_done.set()

    except Exception as e:
        logging.debug(f"Error occurred while searching for text in the image: {e}")
        found = False
        string_found = found # Update the global variable for the main loop
        threads_done.set()
        return found

def start_OCR_thread_():
    global ocr_thread, start_ocr_timestamp, img, search_text, x_coord_window, y_coord_window

    start_ocr_timestamp = time.time()

    threads_done.clear()  # Reset the state of the threads

    img = capture_window_region_(x_coord_window, y_coord_window)  # Define the 'img' variable
    search_text = winamax_string_searched  # Define the 'search_text' variable

    ocr_thread = threading.Thread(target=OCR_string_search_, args=(img, search_text,))
    logging.debug("Starting OCR thread.")

    ocr_thread.start()

def show_button_(coords, hwnd):
    """
    Displays a button window at the specified coordinates and associates it with the given window handle (hwnd).
    If the button window is already displayed, it updates its position and brings it to the front.

    Parameters:
    - coords (tuple): A tuple containing the x and y coordinates where the button should be displayed.
    - hwnd (int): The window handle (hwnd) of the window to associate with the button.

    Note:
    - This function uses a global variable `button_window_var` to keep track of the button window instance.
    - If the button window is not already displayed, it creates a new instance of `ButtonWindow` and shows it.
    - If the button window is already displayed, it updates its position, brings it to the front, and updates the associated hwnd.
    """

    global button_window_var

    if not button_window_var:
        button_window_var = ButtonWindow(coords, hwnd)
        button_window_var.show()
    else:
        button_window_var.setGeometry(QRect(coords[0], coords[1], 100, 100))
        button_window_var.raise_()
        button_window_var.activateWindow()
        button_window_var.hwnd = hwnd # Update the associated hwnd

def hide_button_():
    """
    Hides the button window if it is currently displayed.
    Closes the window and releases the global reference to this instance of ButtonWindow.
    """

    global button_window_var

    if button_window_var:
        button_window_var.close()
        button_window_var = None
        logging.debug("Button window closed.")

def capture_image_and_save_(hwnd):
    """
    Capture an image of the window specified by `hwnd` via the screen_session_result_() function, save the image in a JPG file with a name based on the timestamp,
    and display a confirmation or error message.

    :param hwnd: Handle of the window whose image needs to be captured
    """

    result_jpg = screen_session_result_(hwnd)

    if result_jpg:
        new_month_folder = current_month_()
        # Ensure the folder name is correctly encoded
        new_month_folder = new_month_folder.encode('utf-8').decode('utf-8')
        full_path = os.path.join(capture_folder, new_month_folder)
        logging.info(f"Saving the image in the folder: {full_path}")

        if not os.path.exists(full_path):
            logging.info(f"Creating the folder: {full_path}")
            os.makedirs(full_path)

        timestamp = datetime.now().strftime("%d_%m_%Y")
        file_path = os.path.join(full_path, f"{timestamp}.jpg")
        result_jpg.save(file_path, "JPEG")
        logging.info(f"Image successfully saved: {file_path}")
    else:
         logging.info("Error capturing the image.")

def current_month_():
    """
    Get the current month's name in French.

    This function sets the locale to French to display the month name in French,
    retrieves the current month's name, and returns it.

    Returns:
        str: The name of the current month in French.
    """

    # Set the locale to French to display the month name in French
    locale.setlocale(locale.LC_TIME, 'fr_FR')
    
    # Get the current month's name
    month = time.strftime("%B")
    
    return month

def screen_session_result_(hwnd):
    """
    Capture a specific part of the window identified by `hwnd` and return the captured image.
    The dimension corresponds to the rectangle composing the "Statistics" part of Winamax, with the session results.

    :param hwnd: Handle of the window (HWND) to capture
    :return: PIL image of the captured part of the window, or None in case of error
    """

    logging.debug(f"Attempting to capture for the window with HWND: {hwnd}")

    try:
        win = gw.Window(hwnd)
        rect = (win.left, win.top, win.right, win.bottom)

        height_rectifier_up = 134  # Offset start height of the part to capture (top side of the rectangle)
        height_rectifier_low = 178  # Offset end height of the part to capture (bottom side of the rectangle)
        width_rectifier = 900  # Offset of the width of the part to capture (left and right sides of the rectangle)

        window_width = rect[2] - rect[0]  # Width of the window
        max_offset = 300  # Maximum allowed offset
        threshold_width = 1414  # Threshold width in pixels to start the offset (black bars appear at this point)

        # Calculate the progressive offset based on the width of the window
        if window_width > threshold_width:
            extra_width = window_width - threshold_width
            additional_offset = int(min(max_offset, extra_width / 2))  # Divided by two because the black bars are on both sides of the window
        else:
            additional_offset = 0

        capture_rect = (
            rect[0] + additional_offset,  # Progressive adjustment of the left side
            rect[1] + height_rectifier_up,  # Window height minus the height rectifier
            rect[0] + width_rectifier + additional_offset,  # Progressive adjustment of the width
            rect[1] + height_rectifier_low  # Window height minus the bottom rectifier
        )

        if capture_rect[0] >= capture_rect[2] or capture_rect[1] >= capture_rect[3]:
            logging.debug(f"The adjusted coordinates of the capture rectangle are invalid: {capture_rect}")
            return None

        with mss.mss() as sct:
            img = sct.grab(capture_rect)
            img_pil = Image.frombytes('RGB', img.size, img.rgb)
            
            logging.debug(f"Image capture successful for the window with HWND: {hwnd}")
        return img_pil
    
    except Exception as e:
        logging.debug(f"Error capturing the window with HWND {hwnd}: {e}")
        return None

def calculate_percentage_and_position_(x, y, width):
    """
    Calculate the percentage of the width for the button position and the position in pixels.
    :param x: X coordinate of the top-left corner of the window
    :param y: Y coordinate of the top-left corner of the window
    :param width: Width of the window
    :return: (x, y) button coordinates
    """

    if width <= 1414: # Small window before the trigger point
        percentage = 95 # Max percentage for small windows
    else:
        scale = (82 - 95) / (2000 - 1414) # Scale factor for the percentage
        percentage = 95 + scale * (width - 1414) # Calculate the percentage based on the width

    x = int(x + ((percentage * width) / 100)) # Calculate the X coordinate of the button based on width
    y = int(y + 108) # Calculate the Y coordinate of the button, fixed since it's on the same height at all times

    logging.debug(f"Width: {width}, Width percentage: {percentage}%, Button coordinates: ({x}, {y})")

    return x, y

class ButtonWindow(QWidget):
    def __init__(self, coords, hwnd, parent=None):
        super(ButtonWindow, self).__init__(parent)
        self.hwnd = hwnd
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setGeometry(QRect(coords[0], coords[1], 100, 100))

        self.label = QLabel(self)
        self.label.setPixmap(QPixmap(button_image_path))
        self.label.setGeometry(0, 0, 100, 100)

        self.button = QPushButton(self)
        self.button.setGeometry(0, 0, 100, 100)
        self.button.setFlat(True)
        self.button.setStyleSheet("background: transparent;")
        self.button.clicked.connect(self.on_button_click)

    def on_button_click(self):
        logging.info("Button clicked.")
        capture_image_and_save_(self.hwnd)

def main():

    global x_coord_window, y_coord_window, start_ocr_timestamp, string_found

    app = QApplication([]) # Create a QApplication instance

    while True:      
        # Get the current timestamp
        current_timestamp = time.time() 

        # Call the function to check for the "winamax.exe" process
        found_wmx_proc = check_wmx_proc_alive_()

        # Check if any "winamax.exe" process is found
        if found_wmx_proc:
            # Get the PIDs of "winamax.exe" processes
            wmx_pids = get_wmx_pids_()
            logging.debug(f"Winamax PIDs: {wmx_pids}")

            # Get the HWNDs of "winamax.exe" processes
            wmx_hwnd_list = get_wmx_hwnd_and_title_(wmx_pids)
            logging.debug(f"Winamax HWNDs: {wmx_hwnd_list}")

            # Get the HWNDs with the specified window title
            wmx_hwnd_list_filtered = filter_hwnd_list_(wmx_hwnd_list, winamax_window_name)

            # Check if the filtered HWNDs is found
            if wmx_hwnd_list_filtered:
                # Get the first HWND and title from the filtered list
                hwnd, title = wmx_hwnd_list_filtered[0]
                logging.debug(f"Winamax window found: {title} (HWND: {hwnd})")

                # Get the position and dimensions of the window
                x, y, width, height = get_window_position_and_dimensions_(hwnd)
                logging.debug(f"Window position: ({x}, {y}), dimensions: {width}x{height}")

                # Check if the window is minimized
                if x == -32000 and y == -32000:
                    string_found = False
                    hide_button_()
                    logging.debug("Window is minimized")
                else:
                    # Store the coordinates of the window
                    x_coord_window = x
                    y_coord_window = y
                    
                    # Draw a button on the screen if string_found is True, given by OCR thread
                    if string_found:
                        logging.debug("String found, drawing button on screen.")
                        button_pos_x, button_pos_y = calculate_percentage_and_position_(x, y, width)
                        logging.debug(f"Button position: ({button_pos_x}, {button_pos_y}), hwnd: {hwnd}")
                        show_button_((button_pos_x, button_pos_y), hwnd)
                        logging.debug(f"Affichage du bouton à la position : ({button_pos_x}, {button_pos_y}")
                    else:
                        logging.debug("String not found.")
                        hide_button_()

                    # Réinitialiser l'état des threads
                    threads_done.clear()     
                    
        if current_timestamp - start_ocr_timestamp >= (search_interval_OCR):
            logging.debug("Boucle OCR thread relancé")
            start_OCR_thread_()
        
        # Check if the "escape" key is pressed
        if keyboard.is_pressed('escape'):
            logging.debug("Script terminated by user.")
            break

        app.processEvents()

        # Wait for 1 second before checking again
        # time.sleep(0.01)

if __name__ == "__main__":
    main()