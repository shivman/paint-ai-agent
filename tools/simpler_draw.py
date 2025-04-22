"""
Direct mouse movement approach for MS Paint drawing
This script uses direct mouse events and focuses on simplicity
"""

import os
import sys
import time
import logging
import subprocess
import pyautogui
import ctypes
import win32gui
import win32con

# Ensure fail-safe is off
pyautogui.FAILSAFE = False

def setup_logging():
    """Set up logging"""
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    log_file = os.path.join(log_dir, "simple_drawing.log")
    
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
        # Try other possible titles
        for title in ["Paint", "*.bmp - Paint", "*.png - Paint"]:
            paint_window = win32gui.FindWindow(None, title)
            if paint_window:
                break
    
    if paint_window:
        logging.info(f"Found Paint window: {paint_window}")
        rect = win32gui.GetWindowRect(paint_window)
        logging.info(f"Window rect: {rect}")
        return paint_window
    else:
        logging.error("Could not find Paint window")
        return None

def open_paint():
    """Open MS Paint using os.system"""
    try:
        # Check if Paint is already open
        paint_window = find_paint_window()
        if paint_window:
            logging.info("Paint is already open")
            return paint_window
        
        # Launch Paint
        logging.info("Launching MS Paint...")
        os.system("start mspaint")
        
        # Allow time for Paint to start
        for i in range(10):
            time.sleep(1)
            paint_window = find_paint_window()
            if paint_window:
                logging.info(f"Paint launched after {i+1} seconds")
                return paint_window
        
        logging.error("Failed to launch Paint after 10 seconds")
        return None
        
    except Exception as e:
        logging.error(f"Error opening Paint: {str(e)}")
        return None

def focus_window(hwnd):
    """Focus a window using multiple methods"""
    try:
        # Restore if minimized
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(1)
        
        # Try to bring to foreground
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(1)
        
        # Maximize
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        time.sleep(1)
        
        # Check if focused
        foreground = win32gui.GetForegroundWindow()
        if foreground == hwnd:
            logging.info("Successfully focused window")
            return True
        else:
            # Alternative approach: click on top of the window
            rect = win32gui.GetWindowRect(hwnd)
            center_x = (rect[0] + rect[2]) // 2
            top_y = rect[1] + 10
            pyautogui.click(center_x, top_y)
            time.sleep(1)
            
            # Check again
            foreground = win32gui.GetForegroundWindow()
            if foreground == hwnd:
                logging.info("Successfully focused window with click")
                return True
            else:
                logging.warning("Could not focus window, but continuing anyway")
                return True
                
    except Exception as e:
        logging.error(f"Error focusing window: {str(e)}")
        # Continue anyway
        return True

def direct_mouse_click(x, y, button='left'):
    """Perform a direct mouse click at specified coordinates"""
    pyautogui.moveTo(x, y, duration=0.5)
    time.sleep(0.5)
    pyautogui.click(button=button)
    time.sleep(0.5)

def draw_rectangle_direct(x1, y1, x2, y2):
    """
    Draw a rectangle in MS Paint using direct mouse events
    Coordinates are relative to the canvas
    """
    try:
        # Open Paint
        paint_hwnd = open_paint()
        if not paint_hwnd:
            logging.error("Failed to open Paint")
            return False
        
        # Ensure Paint is in focus
        focus_window(paint_hwnd)
        
        # Get window dimensions
        window_rect = win32gui.GetWindowRect(paint_hwnd)
        win_x, win_y, win_right, win_bottom = window_rect
        win_width = win_right - win_x
        win_height = win_bottom - win_y
        logging.info(f"Paint window: {win_width}x{win_height} at ({win_x},{win_y})")
        
        # Take screenshot before for comparison
        screenshot_before = pyautogui.screenshot()
        
        # Press Escape to ensure no tool is active
        pyautogui.press('esc')
        time.sleep(0.5)
        
        # SIMPLIFIED: Use a single direct approach that is most likely to work
        # 1. Select rectangle tool
        logging.info("Selecting rectangle tool")
        
        # First try to click on Shapes in the toolbar (usually around 35% from left)
        shapes_x = win_x + int(win_width * 0.35)
        shapes_y = win_y + 75  # Approximately 75px from top edge
        direct_mouse_click(shapes_x, shapes_y)
        
        # Now click on Rectangle (usually first shape in dropdown)
        rect_x = shapes_x
        rect_y = shapes_y + 40
        direct_mouse_click(rect_x, rect_y)
        
        # Calculate canvas area (Paint canvas typically starts ~170px from top)
        canvas_top = win_y + 170
        canvas_left = win_x + 10
        
        # Convert to screen coordinates
        screen_x1 = canvas_left + x1
        screen_y1 = canvas_top + y1
        screen_x2 = canvas_left + x2
        screen_y2 = canvas_top + y2
        
        logging.info(f"Drawing rectangle from ({screen_x1},{screen_y1}) to ({screen_x2},{screen_y2})")
        
        # First, move cursor to a "safe" position to ensure it's visible
        pyautogui.moveTo(win_x + 100, win_y + 100, duration=0.5)
        time.sleep(0.5)
        
        # Move to start position with small circle for visibility
        for _ in range(2):
            offset = 5
            pyautogui.moveTo(screen_x1 + offset, screen_y1, duration=0.1)
            pyautogui.moveTo(screen_x1, screen_y1 + offset, duration=0.1)
            pyautogui.moveTo(screen_x1 - offset, screen_y1, duration=0.1)
            pyautogui.moveTo(screen_x1, screen_y1 - offset, duration=0.1)
        
        # Position for drawing
        pyautogui.moveTo(screen_x1, screen_y1, duration=0.5)
        time.sleep(0.5)
        
        # Press mouse button with visual feedback
        logging.info("Pressing mouse DOWN")
        pyautogui.mouseDown()
        # Do it twice for reliability
        pyautogui.mouseDown()
        time.sleep(0.5)
        
        # Move to end position (do it in steps for visibility)
        # First move horizontally
        midpoint_x = screen_x2
        midpoint_y = screen_y1
        logging.info(f"Moving horizontally to ({midpoint_x}, {midpoint_y})")
        pyautogui.moveTo(midpoint_x, midpoint_y, duration=0.7)
        time.sleep(0.3)
        
        # Then move vertically to complete rectangle
        logging.info(f"Moving vertically to ({screen_x2}, {screen_y2})")
        pyautogui.moveTo(screen_x2, screen_y2, duration=0.7)
        time.sleep(0.5)
        
        # Release mouse button with visual feedback
        logging.info("Releasing mouse UP")
        pyautogui.mouseUp()
        # Do it twice for reliability
        pyautogui.mouseUp()
        time.sleep(1)
        
        # Move cursor away to see result
        pyautogui.moveTo(win_x + 50, win_y + 50, duration=0.5)
        time.sleep(1)
        
        # Take screenshot after for comparison
        screenshot_after = pyautogui.screenshot()
        
        # Save debug screenshots
        debug_dir = "debug_screenshots"
        if not os.path.exists(debug_dir):
            os.makedirs(debug_dir)
            
        timestamp = int(time.time())
        before_path = os.path.join(debug_dir, f"simple_before_{timestamp}.png")
        after_path = os.path.join(debug_dir, f"simple_after_{timestamp}.png")
        screenshot_before.save(before_path)
        screenshot_after.save(after_path)
        logging.info(f"Saved debug screenshots: {before_path}, {after_path}")
        
        # Check if pixels have changed in drawing area
        changed_pixels = 0
        for x in range(min(screen_x1, screen_x2), max(screen_x1, screen_x2), 20):
            for y in range(min(screen_y1, screen_y2), max(screen_y1, screen_y2), 20):
                try:
                    before_rgb = screenshot_before.getpixel((x, y))
                    after_rgb = screenshot_after.getpixel((x, y))
                    if before_rgb != after_rgb:
                        changed_pixels += 1
                        if changed_pixels <= 5:  # Log a few samples
                            logging.info(f"Pixel changed at ({x},{y}): {before_rgb} -> {after_rgb}")
                except Exception as e:
                    pass
        
        logging.info(f"Detected {changed_pixels} changed pixels")
        if changed_pixels > 5:
            logging.info("Rectangle appears to have been drawn successfully")
            return True
        else:
            logging.error("Rectangle may not have been drawn (no significant pixel changes)")
            return False
            
    except Exception as e:
        logging.error(f"Error drawing rectangle: {str(e)}", exc_info=True)
        return False

def main():
    """Main test function"""
    setup_logging()
    logging.info("=== Starting Simple Rectangle Drawing Test ===")
    print("Starting Simple Rectangle Drawing Test...")
    
    try:
        # Draw rectangle from 100,100 to 300,300
        logging.info("Drawing rectangle from (100,100) to (300,300)")
        result = draw_rectangle_direct(100, 100, 300, 300)
        
        if result:
            logging.info("Test PASSED: Rectangle drawn successfully")
            print("SUCCESS: Rectangle drawn")
        else:
            logging.error("Test FAILED: Could not draw rectangle")
            print("ERROR: Failed to draw rectangle")
            
    except Exception as e:
        logging.error(f"Error in test: {str(e)}", exc_info=True)
        print(f"ERROR: {str(e)}")
        
    logging.info("=== Test Completed ===")
    print("Test Completed. Check logs for details.")

if __name__ == "__main__":
    main() 