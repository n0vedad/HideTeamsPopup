import os
import ctypes
import tkinter as tk
import threading
import time
import logging
import tempfile
import sys
from ctypes import wintypes
from pystray import Icon as icon, Menu as menu, MenuItem as item
from PIL import Image

# Initialise exit logic and set logging
is_exiting = False
temp_dir = tempfile.gettempdir()
log_file_path = os.path.join(temp_dir, 'HideTeamsPopup.log')
logging.basicConfig(filename=log_file_path, level=logging.ERROR,format='%(asctime)s:%(levelname)s:%(message)s')

# Define the prototype for the callback function expected by the 'EnumWindows' API call.
EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

# Define a Class for some Constants that are needed
class Constants:
    # Windows message constants
    GWL_STYLE = -16
    WS_THICKFRAME = 0x00040000
    WS_MAXIMIZEBOX = 0x00010000
    STRING_LENGTH = 100
    CLASS_NAME = b"Chrome_WidgetWin_1"
    CHILD_TITLE = b"Chrome Legacy Window"
    SW_HIDE = 0
    GWL_EXSTYLE = -20

    # Icon and UI constants
    ICONS_DIR = "icons"
    ICON_FILENAME = "icon.ico"
    APP_NAME = "HideTeamsPopup"
    EXIT_LABEL = "Exit"

# Define the main class
class TeamsPopupHandler:
    def __init__(self):
        # Initializing Windows API functions using ctypes
        self.stop = False
        self.EnumWindows = ctypes.windll.user32.EnumWindows
        self.GetClassNameA = ctypes.windll.user32.GetClassNameA
        self.GetWindowLongA = ctypes.windll.user32.GetWindowLongA
        self.GetWindowTextA = ctypes.windll.user32.GetWindowTextA
        self.ShowWindow = ctypes.windll.user32.ShowWindow
        self.IsWindow = ctypes.windll.user32.IsWindow
        self.IsWindowVisible = ctypes.windll.user32.IsWindowVisible
        self.GetWindowRect = ctypes.windll.user32.GetWindowRect

    # Callback method to find and process Teams popup windows
    def find_teams_popup(self, hwnd, lParam):
        try:
            if self.IsWindow(hwnd) and self.IsWindowVisible(hwnd):
                if self.is_window_teams_popup(hwnd):
                    ctypes.cast(lParam, ctypes.POINTER(ctypes.c_void_p))[0] = hwnd
                    return False
            return True
        except Exception as e:
            logging.exception(f"Exception occurred while processing window with handle {hwnd}: {e}")
            return True

    # Method to check if a window is a Teams popup
    # This method includes checks to immediately exclude certain cases by returning False
    def is_window_teams_popup(self, hWindow):
            try:
                # Verify the window exists and is visible
                if not (self.IsWindow(hWindow) and self.IsWindowVisible(hWindow)):
                    return False

                # Get the window title
                windowTitle = (ctypes.c_char * Constants.STRING_LENGTH)()
                self.GetWindowTextA(hWindow, windowTitle, Constants.STRING_LENGTH)
                if not windowTitle.value:
                    return False

                # Get the class name of the window
                className = (ctypes.c_char * Constants.STRING_LENGTH)()
                self.GetClassNameA(hWindow, className, Constants.STRING_LENGTH)

                # Get window styles and extended styles
                styles = self.GetWindowLongA(hWindow, Constants.GWL_STYLE)
                exStyles = self.GetWindowLongA(hWindow, Constants.GWL_EXSTYLE)
                if styles == 0 or exStyles == 0:
                    return False

                # Get the window's position and size
                rect = ctypes.wintypes.RECT()
                self.GetWindowRect(hWindow, ctypes.byref(rect))

                # Determine if it is a minimized window
                isMinimized = rect.left <= -32000 and rect.top <= -32000

                # Check if it is a popup
                isPopup = (exStyles == 0x108) and not isMinimized and (rect.left >= 0 and rect.top >= 0)

                return isPopup

            except Exception as e:
                logging.exception(f"Exception in is_window_teams_popup: {e}")
                return False

    # Loop function to continuously check for Teams popups and hide them.
    # This function will be invoked for each open window, allowing us to inspect or perform actions on them.
    def check_teams_popup_loop(self):
        while not self.stop:
            try:
                hPopup = ctypes.c_void_p()
                self.EnumWindows(EnumWindowsProc(self.find_teams_popup), ctypes.byref(hPopup))
                if hPopup.value:
                    self.hide_teams_popup(hPopup.value)
            except Exception as e:
                logging.exception(f"Exception occurred in the background thread: {e}")
            time.sleep(0.5)
    
    # Method to hide a given Teams popup window
    def hide_teams_popup(self, hPopup):
        if hPopup and not self.ShowWindow(hPopup, Constants.SW_HIDE):
            logging.error(f"Failed to hide window with handle {hPopup}.")

# Function to setup paths and the system tray icon
def setup_paths_and_icon():
    root_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(root_dir, Constants.ICONS_DIR, Constants.ICON_FILENAME)

    with Image.open(icon_path) as image:
        return icon(Constants.APP_NAME, image.copy(), Constants.APP_NAME)

# Function to handle the exit action from the system tray icon 
# and stopping the background thread and the window process
def exit_action(icon, handler, popup_thread, root):
    global is_exiting
    handler.stop = True
    if popup_thread.is_alive():
        popup_thread.join()

    if icon is not None:
        icon.visible = False
        icon.stop()

    if root:
        root.quit()
        root.destroy()

    is_exiting = True

# Main
if __name__ == "__main__":
    try:
        # Initialize the Tkinter application and hide the main window as it's not needed
        root = tk.Tk()
        root.withdraw()

        # Set up the TeamsPopupHandler for managing Teams popups
        handler = TeamsPopupHandler()

        # Start the background thread that checks for and manages Teams popups
        popup_thread = threading.Thread(target=handler.check_teams_popup_loop)
        popup_thread.start()

        # Set up and display the system tray icon for the application
        icon_item = setup_paths_and_icon()
        exit_action_callback = lambda: exit_action(icon_item, handler, popup_thread, root)
        icon_item.menu = menu(item(Constants.EXIT_LABEL, exit_action_callback))
        icon_item.run()
    except Exception as e:
        logging.exception(f"An unexpected error occurred: {e}")
    finally:
        # Clean up the background thread if it's still running
        if popup_thread and popup_thread.is_alive():
            handler.stop = True
            popup_thread.join()

        # Clean exit
        if is_exiting:
            sys.exit(0)
