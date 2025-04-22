"""
Standalone test for Paint automation functions
This script includes simplified versions of key functions 
to test Paint automation without complex imports
"""

import subprocess
import time
import sys
import os
import logging
import pyautogui
import win32gui
import win32con

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("standalone_test.log"),
        logging.StreamHandler(sys.stdout)
    ]
)

# Basic Paint automation functions
def open_paint():
    """Opens Microsoft Paint application"""
    try:
        logging.info("Opening MS Paint...")
        process = subprocess.Popen(["mspaint.exe"])
        time.sleep(2)  # Give Paint time to open
        
        # Check if process is running
        if process.poll() is None:
            logging.info("Paint process started successfully")
            return True
        else:
            logging.error(f"Paint process exited with code: {process.returncode}")
            return False
    except Exception as e:
        logging.error(f"Error opening Paint: {str(e)}")
        return False

def close_paint():
    """Closes all MS Paint instances"""
    try:
        logging.info("Closing MS Paint...")
        os.system("taskkill /f /im mspaint.exe")
        time.sleep(1)
        logging.info("Paint closed")
        return True
    except Exception as e:
        logging.error(f"Error closing Paint: {str(e)}")
        return False

def find_paint_window():
    """Find the Paint window handle"""
    try:
        # Try different window titles
        for title in ["Untitled - Paint", "Paint", "*Paint"]:
            hwnd = win32gui.FindWindow(None, title)
            if hwnd and win32gui.IsWindowVisible(hwnd):
                logging.info(f"Found Paint window with title '{title}', hwnd: {hwnd}")
                return hwnd
        
        # If that fails, search for windows containing "Paint"
        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if "Paint" in title:
                    results.append(hwnd)
        
        paint_windows = []
        win32gui.EnumWindows(enum_callback, paint_windows)
        
        if paint_windows:
            hwnd = paint_windows[0]
            title = win32gui.GetWindowText(hwnd)
            logging.info(f"Found Paint window with title containing 'Paint': '{title}', hwnd: {hwnd}")
            return hwnd
        
        logging.error("Could not find Paint window")
        return None
    except Exception as e:
        logging.error(f"Error finding Paint window: {str(e)}")
        return None

def focus_paint_window():
    """Focus and maximize the Paint window"""
    try:
        hwnd = find_paint_window()
        if not hwnd:
            logging.error("No Paint window found to focus")
            return False
        
        # Restore if minimized
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.5)
        
        # Set as foreground window
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.5)
        
        # Maximize
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        time.sleep(0.5)
        
        # Verify focus
        if win32gui.GetForegroundWindow() == hwnd:
            logging.info("Successfully focused Paint window")
            return True
        else:
            logging.warning("Failed to focus Paint window")
            return False
    except Exception as e:
        logging.error(f"Error focusing Paint window: {str(e)}")
        return False

def get_canvas_info():
    """Get the bounds of the Paint canvas area"""
    try:
        hwnd = find_paint_window()
        if not hwnd:
            logging.error("No Paint window found")
            return None
        
        # Get window dimensions
        win_rect = win32gui.GetWindowRect(hwnd)
        win_left, win_top, win_right, win_bottom = win_rect
        win_width = win_right - win_left
        win_height = win_bottom - win_top
        
        logging.info(f"Paint window dimensions: {win_width}x{win_height}")
        
        # Estimate canvas area (adjust these values based on your Paint version)
        # These are rough estimates and may need to be adjusted
        title_height = 30
        ribbon_height = 130
        status_height = 30
        left_margin = 50
        right_margin = 30
        
        canvas_left = win_left + left_margin
        canvas_top = win_top + title_height + ribbon_height
        canvas_right = win_right - right_margin
        canvas_bottom = win_bottom - status_height
        
        canvas_info = {
            'left': canvas_left,
            'top': canvas_top,
            'right': canvas_right,
            'bottom': canvas_bottom
        }
        
        logging.info(f"Estimated canvas area: {canvas_info}")
        return canvas_info
    except Exception as e:
        logging.error(f"Error getting canvas info: {str(e)}")
        return None

def draw_rectangle(start_x, start_y, end_x, end_y):
    """Draw a rectangle from start to end coordinates"""
    try:
        logging.info(f"Drawing rectangle from ({start_x}, {start_y}) to ({end_x}, {end_y})")
        
        # Move to start position
        pyautogui.moveTo(start_x, start_y, duration=0.5)
        time.sleep(0.5)
        
        # Press mouse button
        pyautogui.mouseDown()
        time.sleep(0.5)
        
        # Drag to end position
        pyautogui.moveTo(end_x, end_y, duration=1.0)
        time.sleep(0.5)
        
        # Release mouse button
        pyautogui.mouseUp()
        time.sleep(0.5)
        
        logging.info("Rectangle drawing completed")
        return True
    except Exception as e:
        logging.error(f"Error drawing rectangle: {str(e)}")
        return False

def run_standalone_test():
    """Run a standalone test of Paint automation"""
    try:
        logging.info("Starting standalone Paint automation test")
        
        # Step 1: Open Paint
        if not open_paint():
            logging.error("Failed to open Paint. Test failed.")
            return False
        
        time.sleep(2)
        
        # Step 2: Focus Paint window
        if not focus_paint_window():
            logging.error("Failed to focus Paint window. Test failed.")
            close_paint()
            return False
        
        time.sleep(1)
        
        # Step 3: Get canvas info
        canvas_info = get_canvas_info()
        if not canvas_info:
            logging.error("Failed to get canvas info. Test failed.")
            close_paint()
            return False
        
        # Step 4: Draw a rectangle in the center of the canvas
        canvas_width = canvas_info['right'] - canvas_info['left']
        canvas_height = canvas_info['bottom'] - canvas_info['top']
        
        rect_size = min(200, canvas_width // 3, canvas_height // 3)
        
        center_x = canvas_info['left'] + (canvas_width // 2)
        center_y = canvas_info['top'] + (canvas_height // 2)
        
        rect_left = center_x - (rect_size // 2)
        rect_top = center_y - (rect_size // 2)
        rect_right = center_x + (rect_size // 2)
        rect_bottom = center_y + (rect_size // 2)
        
        if not draw_rectangle(rect_left, rect_top, rect_right, rect_bottom):
            logging.error("Failed to draw rectangle. Test failed.")
            close_paint()
            return False
        
        time.sleep(3)  # Pause to see the result
        
        # Step 5: Close Paint
        close_paint()
        
        logging.info("Standalone test completed successfully")
        return True
    except Exception as e:
        logging.error(f"Error in standalone test: {str(e)}")
        try:
            close_paint()
        except:
            pass
        return False

if __name__ == "__main__":
    success = run_standalone_test()
    print(f"\nTest {'PASSED' if success else 'FAILED'}")
    print("Check the log file for details: standalone_test.log") 