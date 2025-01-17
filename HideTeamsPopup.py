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
# This ensures we can track any issues and exit cleanly
is_exiting = False
temp_dir = tempfile.gettempdir()
log_file_path = os.path.join(temp_dir, 'HideTeamsPopup.log')
logging.basicConfig(filename=log_file_path, level=logging.ERROR,
                   format='%(asctime)s:%(levelname)s:%(message)s')

# Define the prototype for the callback function expected by the 'EnumWindows' API call.
# This is required for proper callback handling in the Windows API
EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

# Define a Class for some Constants that are needed throughout the application
class Constants:
    # Windows message constants for window manipulation
    GWL_STYLE = -16
    WS_THICKFRAME = 0x00040000
    WS_MAXIMIZEBOX = 0x00010000
    STRING_LENGTH = 100
    CLASS_NAME = b"Chrome_WidgetWin_1"
    CHILD_TITLE = b"Chrome Legacy Window"
    SW_HIDE = 0
    GWL_EXSTYLE = -20
    POPUP_STYLE = 0x108  # Expected style for Teams popup windows

    # Icon and UI constants for the system tray application
    ICONS_DIR = "icons"
    ICON_FILENAME = "icon.ico"
    APP_NAME = "HideTeamsPopup"
    EXIT_LABEL = "Exit"

# Define the main class for handling Teams popup windows
class TeamsPopupHandler:
    def __init__(self):
        # Cache Windows API functions for better performance
        # This reduces the overhead of repeated function lookups
        self.user32 = ctypes.WinDLL('user32', use_last_error=True)
        
        # Direct function assignments with type optimization
        # This improves performance by avoiding repeated lookups and ensuring type safety
        self.EnumWindows = self.user32.EnumWindows
        self.IsWindowVisible = self.user32.IsWindowVisible
        self.GetWindowLongA = self.user32.GetWindowLongA
        self.ShowWindow = self.user32.ShowWindow
        self.GetWindowRect = self.user32.GetWindowRect
        
        # Set function arguments and return types for better type safety
        self.IsWindowVisible.argtypes = [wintypes.HWND]
        self.IsWindowVisible.restype = wintypes.BOOL
        self.GetWindowLongA.argtypes = [wintypes.HWND, ctypes.c_int]
        self.GetWindowLongA.restype = ctypes.c_long
        
        self.stop = False

    # Callback method to find and process Teams popup windows
    # This function is called for each window during enumeration
    def find_teams_popup(self, hwnd, lParam):
        try:
            # Quick check without API calls to improve performance
            if not hwnd:
                return True
                
            # Check visibility as first filter
            # This quickly eliminates hidden windows
            if not self.IsWindowVisible(hwnd):
                return True
                
            # Check style as second quick filter
            # This identifies potential popup windows by their style
            if self.GetWindowLongA(hwnd, Constants.GWL_EXSTYLE) != Constants.POPUP_STYLE:
                return True
                
            # If we got here, it should be a Teams popup
            # Store the window handle and stop enumeration
            ctypes.cast(lParam, ctypes.POINTER(ctypes.c_void_p))[0] = hwnd
            return False
                
        except Exception as e:
            logging.exception(f"Error in find_teams_popup: {e}")
            return True

    # Method to check if a window is a Teams popup
    # This method includes checks to immediately exclude certain cases
    def is_window_teams_popup(self, hWindow):
        try:
            # Quick base checks first for performance
            if not self.IsWindowVisible(hWindow):
                return False
                
            # Direct style check before more expensive operations
            exStyles = self.GetWindowLongA(hWindow, Constants.GWL_EXSTYLE)
            if exStyles != Constants.POPUP_STYLE:
                return False
                
            # Check window position only if needed
            # This helps identify if the window is minimized
            rect = ctypes.wintypes.RECT()
            self.GetWindowRect(hWindow, ctypes.byref(rect))
            if rect.left <= -32000 or rect.top <= -32000:  # Minimized window check
                return False
                
            return True  # If all criteria are met
                
        except Exception as e:
            logging.exception(f"Exception in is_window_teams_popup: {e}")
            return False

    # Loop function to continuously check for Teams popups and hide them
    # This function runs in a separate thread to avoid blocking the main UI
    def check_teams_popup_loop(self):
        consecutive_empty = 0
        while not self.stop:
            try:
                # Create a pointer to store the found window handle
                hPopup = ctypes.c_void_p()
                self.EnumWindows(EnumWindowsProc(self.find_teams_popup), ctypes.byref(hPopup))
                
                if hPopup.value:
                    consecutive_empty = 0
                    self.hide_teams_popup(hPopup.value)
                    time.sleep(0.1)  # Short pause after finding and hiding a popup
                else:
                    consecutive_empty += 1
                    # Dynamic sleep time based on activity
                    # This reduces CPU usage when no popups are being found
                    sleep_time = min(0.5, 0.1 * consecutive_empty)
                    time.sleep(sleep_time)
                    
            except Exception as e:
                logging.exception(f"Exception in check_teams_popup_loop: {e}")
                time.sleep(0.5)

    # Method to hide a given Teams popup window
    # Returns False if the hiding operation fails
    def hide_teams_popup(self, hPopup):
        if hPopup and not self.ShowWindow(hPopup, Constants.SW_HIDE):
            logging.error(f"Failed to hide window with handle {hPopup}.")

# Function to setup paths and the system tray icon
# Returns the configured system tray icon object
def setup_paths_and_icon():
    root_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(root_dir, Constants.ICONS_DIR, Constants.ICON_FILENAME)

    with Image.open(icon_path) as image:
        return icon(Constants.APP_NAME, image.copy(), Constants.APP_NAME)

# Function to handle the exit action from the system tray icon
# This ensures clean shutdown of all components
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

# Main application entry point
if __name__ == "__main__":
    try:
        # Initialize Tkinter application and hide the main window
        # The window is not needed as we use the system tray
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