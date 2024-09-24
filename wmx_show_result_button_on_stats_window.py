import psutil
import win32gui
import win32con
import win32api
import win32process
import win32clipboard
import time
import locale
from datetime import datetime
from PIL import Image
import pytesseract 
import os
import io
import mss
import pygetwindow as gw
from PyQt5.QtWidgets import QApplication, QLabel, QPushButton, QWidget
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt, QRect, QTimer
import threading
import queue
import keyboard
import logging
import re

# Global variables #

winamax_proc_name = "Winamax.exe"
winamax_window_name = "Winamax"
playground_window_name = "Playground"
winamax_string_searched = "Stat"
button_image_path = r"Assets\DLBTN.png"
x_coord_window = 0
y_coord_window = 0
string_found = None
button_stat_window_var = None
button_save_table_window_var = {} 
button_cpy_table_window_var = {}
button_instances = {}
start_ocr_timestamp = time.time()
search_interval_OCR = 1/2 # Interval in seconds to start the OCR thread
last_pixel_check_timestamp = {} # Initialize a dictionary to store the last pixel check timestamp for each hwnd
search_interval_pixel_color = 1/2 # Interval in seconds to check the pixel color
result_frame_hex_color = "#232323" # Hex color code of the result frame in Winamax

threads_done = threading.Event()
input_queue = queue.Queue() # Queue for storing inputs

# Enable verbose logging
VERBOSE_LOGGING = True

# Configure logging
logging.basicConfig(
    level=logging.DEBUG if VERBOSE_LOGGING else logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%d-%m-%Y | %H:%M:%S'
)

# Path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'

# Path to the folder where the Sessions captures will be saved
stat_folder = "Statistiques Sessions"
if not os.path.exists(stat_folder):
    os.makedirs(stat_folder)

# Path to the folder where the Table captures will be saved
tables_folder = "Résultat Tables"
if not os.path.exists(tables_folder):
    os.makedirs(tables_folder)

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

def get_explorer_pid():
    """
    Get the PID of explorer.exe.
    Returns:
    - int: The PID of explorer.exe, or None if not found.
    """
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'].lower() == 'explorer.exe':
            return proc.info['pid']
    return None

def filter_hwnd_list_main_winamax_window_(hwnd_list, window_name):
    """
    Filter the hwnd_list to only include hwnd with title equal to winamax_window_name.
    Parameters:
    - hwnd_list (list): A list of tuples containing the window handle (hwnd) and the window title (str).
    - winamax_window_name (str): The window title to filter for.
    Returns:
    - hwnd_list_filtered (list): A filtered list of tuples containing the window handle (hwnd) and the window title (str).
    """

    try:
        hwnd_list_filtered = [(hwnd, title) for hwnd, title in hwnd_list if title == window_name] # Filter the list based on the window title matching the specified title
    except Exception as e:
        logging.error(f"Error occurred during filtering: {e}")
        hwnd_list_filtered = []

    return hwnd_list_filtered

def filter_hwnd_list_winamax_tables_(hwnd_list, window_name):
    """
    Filter the hwnd_list to only include hwnd with title starting with window_name,
    excluding the window named exactly "winamax".
    This func is used to list all the tables opened in Winamax, if PlayGround isn't used.
    Parameters:
    - hwnd_list (list): A list of tuples containing the window handle (hwnd) and the window title (str).
    - window_name (str): The window title prefix to filter for.
    Returns:
    - hwnd_list_filtered (list): A filtered list of tuples containing the window handle (hwnd) and the window title (str).
    """
    try:
        hwnd_list_filtered = [
            (hwnd, title) for hwnd, title in hwnd_list 
            if title.lower().startswith(window_name.lower()) and title.lower() != window_name.lower()
        ]
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

def get_center_rectangle(window_width, window_height):
    """
    Calculate the coordinates of the rectangle centered within a table window.
    The size of the rectangle depends on the width of the main window and scales gradually until a maximum size is reached.
    
    Parameters:
    - window_width (int): The width of the main window.
    - window_height (int): The height of the main window.
    
    Returns:
    - tuple: The coordinates of the rectangle (left, top, right, bottom).
    """

    # Define the minimum and maximum sizes of the rectangle
    min_rect_width, min_rect_height = 380, 190
    max_rect_width, max_rect_height = 590, 280
    
    # Determine the size of the rectangle based on the width of the main window
    if window_width <= 800:
        rect_width, rect_height = min_rect_width, min_rect_height
    elif window_width <= 1220:
        # Scale the rectangle size gradually
        scale_factor = (window_width - 800) / (1220 - 800)
        rect_width = min_rect_width + int(scale_factor * (max_rect_width - min_rect_width))
        rect_height = min_rect_height + int(scale_factor * (max_rect_height - min_rect_height))
    else:
        rect_width, rect_height = max_rect_width, max_rect_height
    
    # Calculate the coordinates of the rectangle to center it within the main window
    left = (window_width - rect_width) // 2
    top = (window_height - rect_height) // 2
    right = left + rect_width
    bottom = top + rect_height 
    
    return (left, top, right, bottom)

def is_full_screen(hwnd):
    """
    Check if the window is in full screen mode.
    Parameters:
    - hwnd (int): The handle of the window.
    Returns:
    - bool: True if the window is in full screen mode, False otherwise.
    """
    screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    rect = win32gui.GetWindowRect(hwnd)
    return rect[0] == 0 and rect[1] == 0 and rect[2] == screen_width and rect[3] == screen_height

def is_window_visible_(hwnd):
    """
    Check if a window is visible on the screen and not obscured by another window.
    Parameters:
    - hwnd (int): The handle of the window.
    Returns:
    - bool: True if the window is visible on the screen and not obscured, False otherwise.
    """

    global button_save_table_window_var

    # Get the PID of explorer.exe
    explorer_pid = get_explorer_pid()

    # Get the window title
    title = win32gui.GetWindowText(hwnd)
    logging.debug(f"Checking visibility for window: {hwnd}, Title: {title}")
    
    if not win32gui.IsWindowVisible(hwnd):
        logging.debug(f"Window {title} is not visible.")
        return False
    
    if win32gui.IsIconic(hwnd):
        logging.debug(f"Window {title} is minimized.")
        return False
    
    rect = win32gui.GetWindowRect(hwnd) # (left, top, right, bottom)
    logging.debug(f"Window {title} rect: {rect}")

    if rect[0] == -32000 or rect[1] == -32000: # Check if the window is minimized
        logging.debug(f"Window {title} is minimized.")
        return False
    
    # Check if the window is in full screen mode
    full_screen = is_full_screen(hwnd)
    if full_screen:
        logging.debug(f"Window {title} is in full screen mode.")

    # Check if the window is obscured by another window
    top_hwnd = win32gui.GetTopWindow(None)
    while top_hwnd:
        top_title = win32gui.GetWindowText(top_hwnd)
        if top_hwnd == hwnd:
            logging.debug(f"Window {title} is the top window.") # The window is the top window
            return True
        elif top_title == "python": # Check if the window overlapping is a python frame (button is) ### FIXME: Find a better way to ignore python frames
            logging.debug(f"Ignoring invisible python window: {top_hwnd}, Title: {top_title}")
        elif full_screen and not win32gui.IsWindowVisible(top_hwnd): # Ignore invisible windows when in full screen
            logging.debug(f"Ignoring invisible window: {top_hwnd}, Title: {top_title}")
        elif win32gui.IsWindowVisible(top_hwnd): # Check if the window is visible
            top_rect = win32gui.GetWindowRect(top_hwnd)
            if (rect[0] < top_rect[2] and rect[2] > top_rect[0] and
                rect[1] < top_rect[3] and rect[3] > top_rect[1]):
                # Get the PID of the overlapping window
                _, pid = win32process.GetWindowThreadProcessId(top_hwnd)
                # Ignore explorer.exe taskbar windows (it's the taskbar overlapping our table)
                if pid == explorer_pid: # Check if the window overlapping is an explorer.exe window (taskbar is if no title)
                    if top_title != "": # Check if the window overlapping is not the taskbar
                        logging.info(f"Explorer window: {top_hwnd}, PID: {pid}, Title: {top_title}")
                        hide_table_buttons_(hwnd) # Hide the button window if the taskbar is overlapping the table
                        return False
                    else:
                        logging.info(f"Window {hwnd} (Title: {title}) is obscured by the Taskbar, so window is visible.")
                        return True
                else:
                    logging.info(f"Window {hwnd} (Title: {title}) is obscured by window {top_hwnd} (Title: {top_title}), PID: {pid}.")
                    return False
                
        top_hwnd = win32gui.GetWindow(top_hwnd, win32con.GW_HWNDNEXT)
    
    logging.debug(f"Window {title} is visible and not obscured.")
    return True

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

def show_stat_button_(coords, hwnd):
    """
    Displays a button window at the specified coordinates and associates it with the given window handle (hwnd).
    If the button window is already displayed, it updates its position and brings it to the front.

    Parameters:
    - coords (tuple): A tuple containing the x and y coordinates where the button should be displayed.
    - hwnd (int): The window handle (hwnd) of the window to associate with the button.

    Note:
    - This function uses a global variable `button_stat_window_var` to keep track of the button window instance.
    - If the button window is not already displayed, it creates a new instance of `Button_save_result` and shows it.
    - If the button window is already displayed, it updates its position, brings it to the front, and updates the associated hwnd.
    """

    global button_stat_window_var

    if not button_stat_window_var:
        button_stat_window_var = Button_save_result(coords, hwnd)
        button_stat_window_var.show()
    else:
        button_stat_window_var.setGeometry(QRect(coords[0], coords[1], 100, 100))
        button_stat_window_var.hwnd = hwnd # Update the associated hwnd

def show_table_buttons_(coords, hwnd):
    """
    Displays two button windows at the specified coordinates and associates them with the given window handle (hwnd).
    If the button windows are already displayed, it updates their positions.

    Parameters:
    - coords (tuple): A tuple containing the x and y coordinates where the button should be displayed.
    - hwnd (int): The window handle (hwnd) of the window to associate with the button.

    Note:
    - This function uses global variables `button_save_table_window_var` and `button_cpy_table_window_var` to keep track of the button window instances.
    - If the button windows are not already displayed, it creates new instances of `Button_save_table` and `Button_cpy_table` and shows them.
    - If the button windows are already displayed, it updates their positions, brings them to the front, and updates the associated hwnd.
    """

    global button_save_table_window_var, button_cpy_table_window_var

    if hwnd not in button_save_table_window_var:
        button_save_table_window_var[hwnd] = Button_save_table(coords, hwnd)
        logging.debug(f"Save table button created for HWND: {hwnd}")
        button_save_table_window_var[hwnd].show()

    else:
        button = button_save_table_window_var[hwnd]
        button.setGeometry(QRect(coords[0], coords[1], 100, 100))
        button.hwnd = hwnd # Update the associated hwnd

    if hwnd not in button_cpy_table_window_var:
        button_cpy_table_window_var[hwnd] = Button_cpy_table((coords[0], coords[1]-100), hwnd)
        logging.debug(f"Copy button created for HWND: {hwnd}")
        button_cpy_table_window_var[hwnd].show()

    else:
        button = button_cpy_table_window_var[hwnd]
        button.setGeometry(QRect(coords[0], coords[1]-100, 100, 100))
        button.hwnd = hwnd # Update the associated hwnd

def hide_stat_button_():
    """
    Hides the button window if it is currently displayed.
    Closes the window and releases the global reference to this instance of Button_save_result.
    """

    global button_stat_window_var

    if button_stat_window_var:
        button_stat_window_var.close()
        button_stat_window_var = None

        logging.debug("Button window closed.")

def hide_table_buttons_(hwnd):
    """
    Hides the save and copy button windows associated with the specified window handle (hwnd).
    Closes the windows and removes their references from the global dictionaries.
    
    Parameters:
    - hwnd (int): The window handle (hwnd) of the window whose buttons should be hidden.
    """

    global button_save_table_window_var, button_cpy_table_window_var

    if hwnd in button_save_table_window_var or hwnd in button_cpy_table_window_var:
        button_save_table_window_var[hwnd].close()
        button_cpy_table_window_var[hwnd].close()
        del button_save_table_window_var[hwnd]
        del button_cpy_table_window_var[hwnd]

        logging.debug(f"Save button closed on window with HWND: {hwnd}")
        logging.debug(f"Copy button closed on window with HWND: {hwnd}")

def save_result_screenshot_(hwnd):
    """
    Capture an image of the window specified by `hwnd` via the screen_session_result_() function, save the image in a JPG file, 
    in the Result Folder, with a name based on the timestamp, and display a confirmation or error message.

    :param hwnd: Handle of the window whose image needs to be captured
    """

    result_jpg = screen_session_result_(hwnd)

    if result_jpg:
        new_month_folder = current_month_()
        # Ensure the folder name is correctly encoded
        new_month_folder = new_month_folder.encode('utf-8').decode('utf-8')
        full_path = os.path.join(stat_folder, new_month_folder)
        logging.debug(f"Saving the image in the folder: {full_path}")

        if not os.path.exists(full_path):
            logging.debug(f"Creating the folder: {full_path}")
            os.makedirs(full_path)

        timestamp = datetime.now().strftime("%d_%m_%Y")
        file_path = os.path.join(full_path, f"{timestamp}.jpg")
        result_jpg.save(file_path, "JPEG")
        logging.info(f"Image successfully saved: {file_path}")
    else:
         logging.debug("Error capturing the image.")

def save_table_screenshot_(hwnd):
    """
    Capture an image of the table specified by `hwnd` via the screen_table_result_() function, save the image in a JPG file, 
    in the Table Result Folder, with a name based on the timestamp, and display a confirmation or error message.

    :param hwnd: Handle of the window whose image needs to be captured
    """

    # FIXIT : test without this # global button_stat_window_var

    hide_table_buttons_(hwnd) # Hide the button window before capturing the table result, will reopen once the main loop is done

    result_jpg = screen_table_result_(hwnd)

    if result_jpg:
        new_month_folder = current_month_()
        # Ensure the folder name is correctly encoded
        new_month_folder = new_month_folder.encode('utf-8').decode('utf-8')
        full_path = os.path.join(tables_folder, new_month_folder)
        logging.debug(f"Saving the image in the folder: {full_path}")

        if not os.path.exists(full_path):
            logging.debug(f"Creating the folder: {full_path}")
            os.makedirs(full_path)

        # Get the window title
        window_title = win32gui.GetWindowText(hwnd)
        # Remove "winamax" from the window title
        window_title = window_title.replace("Winamax", "").strip()
        # Remove everything between parentheses, including the parentheses themselves
        window_title = re.sub(r'\(.*?\)', '', window_title).strip()
        logging.debug(f"Window title: {window_title}")
        # Sanitize the window title to be used in the file name
        sanitized_title = "".join(c for c in window_title if c.isalnum() or c in (' ', '_')).rstrip()

        timestamp = datetime.now().strftime("%d_%m_%Y")
        file_path = os.path.join(full_path, f"{timestamp}_{sanitized_title}.jpg")
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

def screen_table_result_(hwnd):
    """
    Capture a specific part of the table identified by `hwnd` and return the captured image.
    The dimension corresponds to the rectangle composing the Result part of the Winamax table.

    :param hwnd: Handle of the window (HWND) to capture
    :return: PIL image of the captured part of the window, or None in case of error
    """

    logging.debug(f"Attempting to capture for the Result of the table with HWND: {hwnd}")

    try:
        x, y, width, height = get_window_position_and_dimensions_(hwnd) # Get the position and dimensions of the table window
        rectangle_coord = get_center_rectangle(width, height) # Get the coordinates of the result rectangle
        logging.debug(f"Table position: ({x}, {y}), dimensions: {width}x{height}")
        logging.debug(f"Result rectangle: ({rectangle_coord[0]}, {rectangle_coord[1]}, dimensions: {rectangle_coord[2] - rectangle_coord[0]}x{rectangle_coord[3] - rectangle_coord[1]}")
        

        capture_window = (
            x + rectangle_coord[0],  # Left side of the window + Left side of the rectangle
            y + rectangle_coord[1],  # Top side of the window + Top side of the rectangle
            x + rectangle_coord[2],  # Right side of the window + Right side of the rectangle
            y + rectangle_coord[3], # Bottom side of the window + Bottom side of the rectangle
        )

        logging.debug(f"Capture rectangle: ({capture_window[0]}, {capture_window[1]}, {capture_window[2]}, {capture_window[3]})")

        if capture_window[0] >= capture_window[2] or capture_window[1] >= capture_window[3]:
            logging.debug(f"The adjusted coordinates of the capture rectangle are invalid: {capture_window}")
            return None

        with mss.mss() as sct:
            img = sct.grab(capture_window)
            img_pil = Image.frombytes('RGB', img.size, img.rgb)        
            logging.debug(f"Image capture successful for the result of the table with HWND: {hwnd}")

        return img_pil
    
    except Exception as e:
        logging.debug(f"Error capturing the window with HWND {hwnd}: {e}")
        return None

def calculate_stat_btn_pos_(x, y, width):
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

def calculate_table_btn_pos_(x, y, hwnd):
    
    x, y, width, height = get_window_position_and_dimensions_(hwnd) # Get the position and dimensions of the table window
    result_rect = get_center_rectangle(width, height) # Get the coordinates of the result rectangle

    x = x - result_rect[0] + (0.95*width) # Offset to ensure table button is in the top left result rectangle
    y = y + result_rect[1]  # Offset to ensure table button is in the top left result rectangle

    logging.debug(f"Table:{hwnd}, Button coordinates: ({x}, {y})")

    return int(x), int(y)

def check_table_pixel_color_(x, y):
    """
    Check the pixel color at the specified coordinates (x, y) and return True if the color is #232323, False otherwise.
    This function is used to detect the presence of the result frame in a table window.
    Hex color code: result_frame_hex_color var : #232323 is the color of the table result window in Winamax.
    :param x: X coordinate of the pixel
    :param y: Y coordinate of the pixel
    :return: bool indicating if the pixel color is == to result_frame_hex_color var
    """

    global result_frame_hex_color

    with mss.mss() as sct:
        img = sct.grab({"left": x, "top": y, "width": 1, "height": 1})
        r, g, b = img.pixel(0, 0)

        logging.debug(f"Pixel color at ({x}, {y}): R={r}, G={g}, B={b}")

    hex_color = f"#{r:02x}{g:02x}{b:02x}"
    logging.debug(f"Hex color at ({x}, {y}): {hex_color}")

    return hex_color == result_frame_hex_color

def cpy_table_screenshot_(hwnd):

    hide_table_buttons_(hwnd) # Hide the button window before capturing the table result, will reopen once the main loop is done

    result_jpg = screen_table_result_(hwnd)

    if result_jpg:
        output = io.BytesIO()
        result_jpg.save(output, format="JPEG")
        data = output.getvalue()
        output.close()

        # Copy the image to the clipboard
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_DIB, data)
        win32clipboard.CloseClipboard()

        logging.info("Image from Hwnd : {hwnd} successfully copied to clipboard.")
    return 

class Button_save_result(QWidget):
    def __init__(self, coords, hwnd, parent=None):
        super(Button_save_result, self).__init__(parent)
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
        logging.debug("Button clicked.")
        save_result_screenshot_(self.hwnd)

class Button_save_table(QWidget):
    def __init__(self, coords, hwnd, parent=None):
        super(Button_save_table, self).__init__(parent)
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
        logging.debug("Button clicked.")
        save_table_screenshot_(self.hwnd)

class Button_cpy_table(QWidget):
    def __init__(self, coords, hwnd, parent=None):
        super(Button_cpy_table, self).__init__(parent)
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
        logging.debug("Button clicked.")
        cpy_table_screenshot_(self.hwnd)

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

            ### PART 1: Main Winamax Stats window : Drawing button and Screenshot of Results ###

            # Get the HWND of main launcher window title
            wmx_hwnd_list_filtered = filter_hwnd_list_main_winamax_window_(wmx_hwnd_list, winamax_window_name)

            # Check if the filtered Window HWND is found
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
                    hide_stat_button_()
                    logging.debug("Window is minimized")
                else:
                    # Store the coordinates of the window
                    x_coord_window, y_coord_window = x, y
                    
                    # Draw a button on the screen if string_found is True, given by OCR thread
                    if string_found:
                        logging.debug("String found, drawing button on screen.")
                        button_pos_x, button_pos_y = calculate_stat_btn_pos_(x, y, width)
                        logging.debug(f"Button position: ({button_pos_x}, {button_pos_y}), hwnd: {hwnd}")
                        show_stat_button_((button_pos_x, button_pos_y), hwnd)
                        logging.debug(f"Affichage du bouton à la position : ({button_pos_x}, {button_pos_y}")
                    else:
                        logging.debug("String not found.")
                        hide_stat_button_()

                    # Réinitialiser l'état des threads
                    threads_done.clear()   

                ### END OF PART 1 ###

                ### PART 2: Winamax Tables detection ###

            # Get the HWNDs of all Winamax tables
            wmx_hwnd_table_list = filter_hwnd_list_winamax_tables_(wmx_hwnd_list, winamax_window_name)
            logging.debug(f"Tables HWNDs and Title: {wmx_hwnd_table_list}")

            # If there's a list of tables
            if wmx_hwnd_table_list:
                # Filter the list to only include visible tables
                visible_windows = [(hwnd, title) for hwnd, title in wmx_hwnd_table_list if is_window_visible_(hwnd)]
                logging.debug(f"Visible tables: {len(visible_windows)}")

                # Get the position and dimensions of each visible table
                for hwnd, title in visible_windows:
                    x, y, width, height = get_window_position_and_dimensions_(hwnd)
                    logging.debug(f"Table {title} position: ({x}, {y}), dimensions: {width}x{height}, HWND: {hwnd}")
                    # We load the coordinates of the theorical rectangle within the table window 
                    rectangle_coord = get_center_rectangle(width, height) 

                    x_pixel_check_coord = x + rectangle_coord[0] + 20 # Offset of 20 pixels on the left side of the rectangle to ensure we're in the rectangle
                    y_pixel_check_coord = y + rectangle_coord[1] + 20 # Offset of 20 pixels on the top side of the rectangle to ensure we're in the rectangle

                     # Check the last timestamp for this hwnd
                    last_check_time = last_pixel_check_timestamp.get(hwnd, 0)  # Default to 0 if no previous check
                    
                    # If the time since the last check is greater than the specified interval, proceed
                    if current_timestamp - last_check_time >= search_interval_pixel_color: 
                        logging.debug(f"Checking pixel color for table {title}")              
                        # Check the pixel color at the specified coordinates
                        table_result_displayed = check_table_pixel_color_(x_pixel_check_coord, y_pixel_check_coord)
                        if table_result_displayed:
                            logging.debug(f"Result frame on screen : {table_result_displayed} / on table {title}")
                        else:
                            logging.debug(f"Result frame not displayed on table {title}")
                        
                        last_pixel_check_timestamp[hwnd] = current_timestamp

                    else:
                        logging.debug(f"Skipping pixel color check for table {title}")
                        
                    # TEST : Desactivate this part

                    # # If the result frame is displayed, draw a button on the screen
                    # if table_result_displayed:
                    #         button_pos_x, button_pos_y = calculate_table_btn_pos_(x, y, hwnd)
                    #         show_table_buttons_((button_pos_x, button_pos_y), hwnd)
                    #         logging.debug(f"Draw Button at position: ({button_pos_x}, {button_pos_y}), hwnd: {hwnd}")
                    # else:
                    #     logging.debug("Result frame not displayed.")
                    #     hide_table_buttons_(hwnd)

                    button_pos_x, button_pos_y = calculate_table_btn_pos_(x, y, hwnd)
                    show_table_buttons_((button_pos_x, button_pos_y), hwnd)
                    logging.debug(f"Draw Buttons at position: ({button_pos_x}, {button_pos_y}), hwnd: {hwnd}")

                    ### END OF PART 2 ###

                    ### PART 3: Thread Management for Main Winamax Detection, Results of Tables Detection and Exit main loop ###
                    
                    # Check if the "escape" key is pressed

        if keyboard.is_pressed('escape'):
            logging.debug("Script terminated by user.")
            QApplication.quit()
            return

        # Process other periodic tasks
        if current_timestamp - start_ocr_timestamp >= search_interval_OCR:
            logging.debug("Boucle OCR thread relancé")
            start_OCR_thread_()
            start_ocr_timestamp = current_timestamp  # Reset the timestamp

        app.processEvents()

        # Wait for 1 second before checking again
        time.sleep(0.5)

if __name__ == "__main__":
    main()