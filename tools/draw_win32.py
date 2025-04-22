"""
Direct Windows automation for MS Paint using win32com and SendKeys
This script uses more direct Windows automation techniques
"""

import os
import sys
import time
import logging
import subprocess
from PIL import ImageGrab
import win32gui
import win32con
import win32com.client
import pyautogui  # Only for screenshot comparison

def setup_logging():
    """Set up logging"""
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    log_file = os.path.join(log_dir, "win32_drawing.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return log_file

def find_paint_window():
    """Find the Paint window handle"""
    paint_window = win32gui.FindWindow(None, "Untitled - Paint")
    if not paint_window:
        paint_window = win32gui.FindWindow(None, "Paint")
    
    if paint_window:
        logging.info(f"Found Paint window: {paint_window}")
        rect = win32gui.GetWindowRect(paint_window)
        logging.info(f"Window rect: {rect}")
        return paint_window
    else:
        logging.error("Could not find Paint window")
        return None

def open_paint():
    """Open MS Paint using direct subprocess call"""
    try:
        # Check if Paint is already open
        paint_window = find_paint_window()
        if paint_window:
            logging.info("Paint is already open")
            return paint_window
        
        # Use subprocess for better control
        logging.info("Opening Paint...")
        subprocess.Popen(["mspaint.exe"])
        
        # Wait for Paint to open
        for i in range(10):
            time.sleep(1)
            paint_window = find_paint_window()
            if paint_window:
                logging.info(f"Paint opened successfully after {i+1} seconds")
                return paint_window
        
        logging.error("Failed to open Paint after 10 seconds")
        return None
    
    except Exception as e:
        logging.error(f"Error opening Paint: {str(e)}")
        return None

def focus_paint_window(paint_window):
    """Focus the Paint window using multiple methods"""
    try:
        logging.info(f"Focusing Paint window: {paint_window}")
        
        # Create shell object for more direct Windows control
        shell = win32com.client.Dispatch("WScript.Shell")
        
        # Try to restore window if minimized
        placement = win32gui.GetWindowPlacement(paint_window)
        if placement[1] == win32con.SW_SHOWMINIMIZED:
            logging.info("Window is minimized, restoring...")
            win32gui.ShowWindow(paint_window, win32con.SW_RESTORE)
            time.sleep(1)
        
        # Set as foreground window
        win32gui.SetForegroundWindow(paint_window)
        time.sleep(1)
        
        # Maximize window
        win32gui.ShowWindow(paint_window, win32con.SW_MAXIMIZE)
        time.sleep(1)
        
        # Alternative method using shell
        shell.AppActivate(paint_window)
        time.sleep(1)
        
        # Check if successful
        active_window = win32gui.GetForegroundWindow()
        if active_window == paint_window:
            logging.info("Successfully focused Paint window")
            return True
        else:
            logging.warning("Could not focus Paint window, but continuing anyway")
            return True
    
    except Exception as e:
        logging.error(f"Error focusing Paint window: {str(e)}")
        # Continue anyway
        return True

def draw_rectangle(x1, y1, x2, y2):
    """Draw a rectangle in Paint using keyboard shortcuts and SendKeys"""
    try:
        # Open Paint
        paint_window = open_paint()
        if not paint_window:
            return False
        
        # Focus window
        if not focus_paint_window(paint_window):
            logging.warning("Could not focus Paint window, but continuing anyway")
        
        # Get window dimensions
        rect = win32gui.GetWindowRect(paint_window)
        win_x, win_y, win_right, win_bottom = rect
        win_width = win_right - win_x
        win_height = win_bottom - win_y
        logging.info(f"Window size: {win_width}x{win_height}")
        
        # Take before screenshot
        screenshot_before = ImageGrab.grab()
        
        # Create shell object for SendKeys
        shell = win32com.client.Dispatch("WScript.Shell")
        
        # Clear any current selection with Escape
        shell.SendKeys("{ESC}")
        time.sleep(1)
        
        # Select rectangle tool using keyboard shortcuts
        # Alt+H to access Home tab
        logging.info("Selecting Home tab")
        shell.SendKeys("%h")  # % is Alt
        time.sleep(1)
        
        # Select Shapes
        logging.info("Selecting Shapes")
        shell.SendKeys("sh")  # Shapes
        time.sleep(1)
        
        # Select Rectangle (typically first shape in first row)
        logging.info("Selecting Rectangle shape")
        shell.SendKeys("{DOWN}")  # Navigate to shapes
        time.sleep(0.5)
        shell.SendKeys("{ENTER}")  # Select first shape (rectangle)
        time.sleep(1)
        
        # Calculate canvas coordinates
        # Paint canvas typically starts around 170px from top of window
        canvas_top = win_y + 170
        canvas_left = win_x + 5
        
        # Convert to screen coordinates
        screen_x1 = canvas_left + x1
        screen_y1 = canvas_top + y1
        screen_x2 = canvas_left + x2
        screen_y2 = canvas_top + y2
        
        logging.info(f"Drawing from ({screen_x1},{screen_y1}) to ({screen_x2},{screen_y2})")
        
        # Use mouse_event for more reliable mouse control
        # Create mouse input sequence 
        # 1. Move to start position
        win32gui.SetCursorPos((screen_x1, screen_y1))
        time.sleep(1)
        
        # 2. Press mouse down
        logging.info("Pressing mouse down")
        win32gui.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(1)
        
        # 3. Move to end position
        win32gui.SetCursorPos((screen_x2, screen_y2))
        time.sleep(1)
        
        # 4. Release mouse
        logging.info("Releasing mouse")
        win32gui.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(1)
        
        # Move cursor away to see result
        win32gui.SetCursorPos((win_x + 50, win_y + 50))
        
        # Take after screenshot
        time.sleep(1)
        screenshot_after = ImageGrab.grab()
        
        # Save debug screenshots
        debug_dir = "debug_screenshots"
        if not os.path.exists(debug_dir):
            os.makedirs(debug_dir)
        
        timestamp = int(time.time())
        before_path = os.path.join(debug_dir, f"win32_before_{timestamp}.png")
        after_path = os.path.join(debug_dir, f"win32_after_{timestamp}.png")
        screenshot_before.save(before_path)
        screenshot_after.save(after_path)
        logging.info(f"Saved debug screenshots: {before_path}, {after_path}")
        
        # Compare screenshots
        pixels_changed = 0
        for x in range(min(screen_x1, screen_x2), max(screen_x1, screen_x2), 20):
            for y in range(min(screen_y1, screen_y2), max(screen_y1, screen_y2), 20):
                try:
                    before_pixel = screenshot_before.getpixel((x, y))
                    after_pixel = screenshot_after.getpixel((x, y))
                    if before_pixel != after_pixel:
                        pixels_changed += 1
                        logging.info(f"Pixel changed at ({x}, {y}): {before_pixel} -> {after_pixel}")
                except Exception as e:
                    pass
        
        logging.info(f"Detected {pixels_changed} changed pixels")
        
        if pixels_changed > 5:
            logging.info("Rectangle appears to be drawn successfully")
            return True
        else:
            logging.error("Rectangle does not appear to be drawn (no pixel changes detected)")
            return False
    
    except Exception as e:
        logging.error(f"Error drawing rectangle: {str(e)}")
        return False

def main():
    """Main function to test drawing a rectangle"""
    setup_logging()
    logging.info("=== Starting Paint Drawing Test with Win32 ===")
    
    try:
        # Draw rectangle at 100,100 to 300,300
        result = draw_rectangle(100, 100, 300, 300)
        
        if result:
            logging.info("Successfully drew rectangle")
            print("SUCCESS: Rectangle drawn")
        else:
            logging.error("Failed to draw rectangle")
            print("ERROR: Failed to draw rectangle")
    
    except Exception as e:
        logging.error(f"Error in main: {str(e)}")
        print(f"ERROR: {str(e)}")
    
    logging.info("=== Test completed ===")

if __name__ == "__main__":
    main() 