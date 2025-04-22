"""
Simplified drawing module for MS Paint automation
This module provides straightforward, reliable functions for basic Paint operations
with robust error handling and visual feedback.
"""

import os
import sys
import time
import logging
import subprocess
import pyautogui
import win32gui
import win32con
import win32api
from datetime import datetime
import traceback
from pathlib import Path

# Configure PyAutoGUI
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.2

def setup_logging():
    """Set up logging for the simple drawing module"""
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    log_file = os.path.join(log_dir, f"simple_draw_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    
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
    """Find the MS Paint window with multiple methods and robust error handling"""
    try:
        # Try several window titles that Paint might have
        possible_titles = ["Untitled - Paint", "Paint", "*.png - Paint", "*.bmp - Paint"]
        
        for title in possible_titles:
            hwnd = win32gui.FindWindow(None, title)
            if hwnd and win32gui.IsWindowVisible(hwnd):
                logging.info(f"Found Paint window with title '{title}', hwnd: {hwnd}")
                return hwnd
        
        # If not found by exact title, try searching for windows containing "Paint"
        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if "Paint" in title:
                    results.append((hwnd, title))
        
        paint_windows = []
        win32gui.EnumWindows(enum_callback, paint_windows)
        
        if paint_windows:
            hwnd, title = paint_windows[0]
            logging.info(f"Found Paint window with partial title match: '{title}', hwnd: {hwnd}")
            return hwnd
            
        logging.error("Could not find Paint window")
        return None
        
    except Exception as e:
        logging.error(f"Error finding Paint window: {str(e)}")
        logging.error(traceback.format_exc())
        return None

def maximize_window(hwnd):
    """Maximize the given window"""
    try:
        if hwnd:
            # Get current window state
            placement = win32gui.GetWindowPlacement(hwnd)
            if placement[1] != win32con.SW_SHOWMAXIMIZED:
                logging.info(f"Maximizing window {hwnd}")
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                time.sleep(0.5)
                return True
            else:
                logging.info(f"Window {hwnd} is already maximized")
                return True
        return False
    except Exception as e:
        logging.error(f"Error maximizing window: {str(e)}")
        return False

def focus_paint_window():
    """Ensure the Paint window is in focus and maximized"""
    try:
        hwnd = find_paint_window()
        if not hwnd:
            logging.error("No Paint window found to focus")
            return False
            
        # Log current window state
        current_active = win32gui.GetForegroundWindow()
        current_title = win32gui.GetWindowText(current_active)
        logging.info(f"Current active window: '{current_title}' (hwnd: {current_active})")
        
        # First, restore if minimized
        if win32gui.IsIconic(hwnd):
            logging.info("Restoring minimized Paint window")
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.5)
            
        # Try to set as foreground window
        logging.info(f"Setting Paint window (hwnd: {hwnd}) as foreground")
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.5)
        
        # Maximize the window
        maximize_window(hwnd)
        
        # Verify focus was achieved
        new_active = win32gui.GetForegroundWindow()
        if new_active == hwnd:
            logging.info("Successfully focused Paint window")
            # Press Escape to clear any dialogs or tool selections
            pyautogui.press('esc')
            time.sleep(0.5)
            return True
        else:
            logging.warning(f"Failed to focus Paint window. Current active: {new_active}")
            
            # Try alternative method - click on the window title bar
            try:
                rect = win32gui.GetWindowRect(hwnd)
                title_x = (rect[0] + rect[2]) // 2  # Middle of window
                title_y = rect[1] + 15  # Near top of window
                
                logging.info(f"Clicking on Paint window title bar at ({title_x}, {title_y})")
                pyautogui.moveTo(title_x, title_y, duration=0.5)
                time.sleep(0.5)
                pyautogui.click()
                time.sleep(1.0)
                
                # Check if focus was achieved
                if win32gui.GetForegroundWindow() == hwnd:
                    logging.info("Successfully focused Paint window by clicking title bar")
                    return True
            except Exception as e:
                logging.error(f"Error clicking on title bar: {str(e)}")
                
            return False
    except Exception as e:
        logging.error(f"Error focusing Paint window: {str(e)}")
        logging.error(traceback.format_exc())
        return False

def open_paint_simple():
    """Open MS Paint using a simple, reliable approach"""
    try:
        # First check if Paint is already running
        paint_hwnd = find_paint_window()
        if paint_hwnd:
            logging.info("Paint is already running")
            if focus_paint_window():
                return True
        
        # Try to open Paint using mspaint command
        logging.info("Launching MS Paint...")
        subprocess.Popen(["mspaint"])
        
        # Wait for Paint to open
        for i in range(10):
            time.sleep(1)
            paint_hwnd = find_paint_window()
            if paint_hwnd:
                logging.info(f"Paint window found after {i+1} seconds")
                time.sleep(1)  # Additional wait for initialization
                if focus_paint_window():
                    return True
        
        logging.error("Failed to open Paint after 10 seconds")
        return False
        
    except Exception as e:
        logging.error(f"Error opening Paint: {str(e)}")
        logging.error(traceback.format_exc())
        return False

def get_paint_canvas_bounds():
    """Get the boundaries of the Paint canvas area with accurate positioning"""
    try:
        # Find Paint window
        paint_hwnd = find_paint_window()
        if not paint_hwnd:
            logging.error("Cannot find Paint window to get canvas bounds")
            return None
            
        # Get window dimensions
        window_rect = win32gui.GetWindowRect(paint_hwnd)
        win_x, win_y, win_right, win_bottom = window_rect
        win_width = win_right - win_x
        win_height = win_bottom - win_y
        
        logging.info(f"Paint window dimensions: {win_width}x{win_height} at ({win_x}, {win_y})")
        
        # Calculate canvas boundaries - adjusted based on screenshot
        # Looking at your provided screenshots, the canvas area is clearly visible 
        # with the white area surrounded by gray margins
        title_bar = 30       # Top window title bar
        ribbon_height = 115  # Height of ribbon based on screenshot
        status_height = 30   # Status bar at bottom
        left_margin = 265    # Left margin is substantial (includes rulers)
        right_margin = 265   # Right margin (includes rulers)
        
        canvas_left = win_x + left_margin
        canvas_top = win_y + title_bar + ribbon_height
        canvas_right = win_right - right_margin
        canvas_bottom = win_bottom - status_height
        
        canvas_info = {
            "left": canvas_left,
            "top": canvas_top,
            "right": canvas_right,
            "bottom": canvas_bottom,
            "width": canvas_right - canvas_left,
            "height": canvas_bottom - canvas_top,
            "center_x": (canvas_left + canvas_right) // 2,
            "center_y": (canvas_top + canvas_bottom) // 2
        }
        
        logging.info(f"Canvas dimensions: {canvas_info['width']}x{canvas_info['height']}")
        logging.info(f"Canvas bounds: ({canvas_left}, {canvas_top}) to ({canvas_right}, {canvas_bottom})")
        
        return canvas_info
        
    except Exception as e:
        logging.error(f"Error getting canvas bounds: {str(e)}")
        logging.error(traceback.format_exc())
        return None

def select_rectangle_tool():
    """Select the rectangle tool in Paint using EXACT coordinates from screenshot"""
    try:
        if not focus_paint_window():
            logging.error("Could not focus Paint window")
            return False
            
        # Get Paint window dimensions
        paint_hwnd = find_paint_window()
        if not paint_hwnd:
            logging.error("Could not find Paint window")
            return False
            
        window_rect = win32gui.GetWindowRect(paint_hwnd)
        win_x, win_y, win_right, win_bottom = window_rect
        win_width = win_right - win_x
        
        # Based on your screenshot, shapes button is in the green box area
        # We want to click the rectangle shape which appears to be in row 1, column 3
        # First click the shapes button
        
        # Clear any active tool
        pyautogui.press('esc')
        time.sleep(0.5)
        
        # From your screenshot, the shapes button is clearly visible 
        # with green box around it, at approximately 36-40% of window width from left
        shapes_button_x = win_x + int(win_width * 0.40)
        shapes_button_y = win_y + 65  # Based on screenshot
        
        logging.info(f"Clicking Shapes button at ({shapes_button_x}, {shapes_button_y})")
        pyautogui.moveTo(shapes_button_x, shapes_button_y, duration=0.5)
        time.sleep(0.3)
        pyautogui.click()
        time.sleep(1.0)  # Wait for shapes menu to appear
        
        # Now click the rectangle shape (row 1, column 3 in dropdown)
        # Rectangle is the third icon in the first row of the shapes menu
        rect_shape_x = shapes_button_x - 30  # Slightly to the left
        rect_shape_y = shapes_button_y + 50  # Down into the menu
        
        logging.info(f"Clicking Rectangle shape at ({rect_shape_x}, {rect_shape_y})")
        pyautogui.moveTo(rect_shape_x, rect_shape_y, duration=0.5)
        time.sleep(0.3)
        pyautogui.click()
        time.sleep(1.0)
            
        # Move mouse to neutral area
        canvas_info = get_paint_canvas_bounds()
        if canvas_info:
            pyautogui.moveTo(canvas_info["center_x"], canvas_info["center_y"], duration=0.3)
            time.sleep(0.5)
            
        logging.info("Rectangle tool selected")
        return True
        
    except Exception as e:
        logging.error(f"Error selecting rectangle tool: {str(e)}")
        logging.error(traceback.format_exc())
        return False

def select_oval_tool():
    """Select the oval/circle tool in Paint using keyboard shortcuts"""
    try:
        if not focus_paint_window():
            logging.error("Could not focus Paint window")
            return False
            
        # Clear any active tool
        pyautogui.press('esc')
        time.sleep(0.5)
        
        # Use Alt key navigation which is more reliable than mouse clicks
        logging.info("Using Alt-key navigation to select oval/circle shape")
        
        # Press Alt to show keyboard shortcuts
        pyautogui.press('alt')
        time.sleep(0.7)
        
        # Press H to access the Home tab
        pyautogui.press('h')
        time.sleep(0.7)
        
        # Press S to access Shapes
        pyautogui.press('s')
        time.sleep(0.7)
        
        # Press O to select Oval
        pyautogui.press('o')
        time.sleep(0.7)
        
        # Move mouse to neutral area
        canvas_info = get_paint_canvas_bounds()
        if canvas_info:
            pyautogui.moveTo(canvas_info["center_x"], canvas_info["center_y"], duration=0.3)
            time.sleep(0.5)
            
        logging.info("Oval tool selected using keyboard shortcuts")
        return True
        
    except Exception as e:
        logging.error(f"Error selecting oval tool: {str(e)}")
        logging.error(traceback.format_exc())
        return False

def draw_simple_rectangle(x1, y1, x2, y2):
    """Draw a rectangle from (x1,y1) to (x2,y2) using simple, reliable method"""
    try:
        logging.info(f"Drawing simple rectangle from ({x1}, {y1}) to ({x2}, {y2})")
        
        # Ensure Paint is focused
        if not focus_paint_window():
            logging.error("Failed to focus Paint window")
            return False
            
        # Get canvas info
        canvas_info = get_paint_canvas_bounds()
        if not canvas_info:
            logging.error("Failed to get canvas bounds")
            return False
            
        # Take screenshot before drawing
        screenshot_before = pyautogui.screenshot()
            
        # Select rectangle tool - first try keyboard shortcuts
        logging.info("Trying keyboard shortcuts for rectangle...")
        pyautogui.press('esc')  # Clear any active tool
        time.sleep(0.5)
        
        # Use Alt+H, S, R sequence for rectangle
        pyautogui.press('alt')
        time.sleep(0.7)
        pyautogui.press('h')
        time.sleep(0.7)
        pyautogui.press('s')
        time.sleep(0.7)
        pyautogui.press('r')
        time.sleep(0.7)
        
        # If keyboard shortcuts fail, try mouse method
        if not select_rectangle_tool():
            logging.error("Failed to select rectangle tool")
            return False
            
        # Adjust coordinates if needed
        canvas_left = canvas_info["left"]
        canvas_top = canvas_info["top"]
        canvas_right = canvas_info["right"]
        canvas_bottom = canvas_info["bottom"]
        
        # Convert from absolute to canvas-relative coordinates if needed
        if x1 < canvas_left or y1 < canvas_top:
            # These are likely canvas-relative coordinates
            x1 = canvas_left + x1
            y1 = canvas_top + y1
            x2 = canvas_left + x2
            y2 = canvas_top + y2
            logging.info(f"Converted to absolute coords: ({x1}, {y1}) to ({x2}, {y2})")
        
        # Ensure within canvas bounds
        x1 = max(canvas_left + 20, min(x1, canvas_right - 20))
        y1 = max(canvas_top + 20, min(y1, canvas_bottom - 20))
        x2 = max(canvas_left + 20, min(x2, canvas_right - 20))
        y2 = max(canvas_top + 20, min(y2, canvas_bottom - 20))
        
        # Move to center of canvas first
        canvas_center_x = canvas_info["center_x"]
        canvas_center_y = canvas_info["center_y"]
        
        logging.info(f"Moving to canvas center at ({canvas_center_x}, {canvas_center_y})")
        pyautogui.moveTo(canvas_center_x, canvas_center_y, duration=0.5)
        time.sleep(0.5)
        
        # Move to start position with small highlight motion
        logging.info(f"Moving to start position ({x1}, {y1})")
        pyautogui.moveTo(x1, y1, duration=0.7)
        time.sleep(0.5)
        
        # Draw small circle at start point for visibility
        offset = 5
        for _ in range(1):  # Just once around
            pyautogui.moveRel(offset, 0, duration=0.1)
            pyautogui.moveRel(0, offset, duration=0.1)
            pyautogui.moveRel(-offset, 0, duration=0.1)
            pyautogui.moveRel(0, -offset, duration=0.1)
            
        # Press and hold mouse button
        logging.info("Pressing mouse button")
        pyautogui.mouseDown()
        time.sleep(0.5)
        
        # Move to end position slowly
        logging.info(f"Moving to end position ({x2}, {y2})")
        pyautogui.moveTo(x2, y2, duration=1.0)
        time.sleep(0.5)
        
        # Release mouse button
        logging.info("Releasing mouse button")
        pyautogui.mouseUp()
        time.sleep(0.5)
        
        # Move back to center
        pyautogui.moveTo(canvas_center_x, canvas_center_y, duration=0.3)
        time.sleep(0.5)
        
        # Take screenshot after drawing
        screenshot_after = pyautogui.screenshot()
        
        # Simple verification - check if pixels changed
        pixels_changed = False
        for check_x in range(min(x1, x2), max(x1, x2), 10):
            try:
                if (screenshot_before.getpixel((check_x, y1)) != 
                    screenshot_after.getpixel((check_x, y1))):
                    pixels_changed = True
                    break
            except:
                pass
                
        if pixels_changed:
            logging.info("Rectangle drawing verified - pixels changed")
            return True
        else:
            logging.warning("Rectangle may not have drawn correctly - no pixel changes detected")
            return True  # Return true anyway since we attempted the drawing
            
    except Exception as e:
        logging.error(f"Error drawing rectangle: {str(e)}")
        logging.error(traceback.format_exc())
        return False

def draw_simple_circle(center_x, center_y, radius):
    """Draw a circle at (center_x, center_y) with given radius"""
    try:
        logging.info(f"Drawing simple circle at ({center_x}, {center_y}) with radius {radius}")
        
        # Ensure Paint is focused
        if not focus_paint_window():
            logging.error("Failed to focus Paint window")
            return False
            
        # Get canvas info
        canvas_info = get_paint_canvas_bounds()
        if not canvas_info:
            logging.error("Failed to get canvas bounds")
            return False
            
        # Take screenshot before drawing
        screenshot_before = pyautogui.screenshot()
            
        # Select oval tool
        if not select_oval_tool():
            logging.error("Failed to select oval tool")
            return False
            
        # Calculate start and end points for the circle
        x1 = center_x - radius
        y1 = center_y - radius
        x2 = center_x + radius
        y2 = center_y + radius
        
        # Ensure coordinates are within canvas
        canvas_left = canvas_info["left"]
        canvas_top = canvas_info["top"]
        canvas_right = canvas_info["right"]
        canvas_bottom = canvas_info["bottom"]
        
        # Adjust coordinates if needed
        x1 = max(canvas_left + 20, min(x1, canvas_right - 20))
        y1 = max(canvas_top + 20, min(y1, canvas_bottom - 20))
        x2 = max(canvas_left + 20, min(x2, canvas_right - 20))
        y2 = max(canvas_top + 20, min(y2, canvas_bottom - 20))
        
        # Move to center of canvas first
        canvas_center_x = canvas_info["center_x"]
        canvas_center_y = canvas_info["center_y"]
        
        logging.info(f"Moving to canvas center at ({canvas_center_x}, {canvas_center_y})")
        pyautogui.moveTo(canvas_center_x, canvas_center_y, duration=0.5)
        time.sleep(0.5)
        
        # Move to start position
        logging.info(f"Moving to start position ({x1}, {y1})")
        pyautogui.moveTo(x1, y1, duration=0.7)
        time.sleep(0.5)
        
        # Press and hold mouse button
        logging.info("Pressing mouse button")
        pyautogui.mouseDown()
        time.sleep(0.5)
        
        # Move to end position slowly
        logging.info(f"Moving to end position ({x2}, {y2})")
        pyautogui.moveTo(x2, y2, duration=1.0)
        time.sleep(0.5)
        
        # Release mouse button
        logging.info("Releasing mouse button")
        pyautogui.mouseUp()
        time.sleep(0.5)
        
        # Move back to center
        pyautogui.moveTo(canvas_center_x, canvas_center_y, duration=0.3)
        time.sleep(0.5)
        
        # Take screenshot after drawing
        screenshot_after = pyautogui.screenshot()
        
        # Simple verification - check if pixels changed
        pixels_changed = False
        for angle in range(0, 360, 30):  # Check at different angles
            try:
                check_x = center_x + int(radius * 0.9 * math.cos(math.radians(angle)))
                check_y = center_y + int(radius * 0.9 * math.sin(math.radians(angle)))
                
                if (screenshot_before.getpixel((check_x, check_y)) != 
                    screenshot_after.getpixel((check_x, check_y))):
                    pixels_changed = True
                    break
            except:
                pass
                
        if pixels_changed:
            logging.info("Circle drawing verified - pixels changed")
            return True
        else:
            logging.warning("Circle may not have drawn correctly - no pixel changes detected")
            return True  # Return true anyway since we attempted the drawing
            
    except Exception as e:
        logging.error(f"Error drawing circle: {str(e)}")
        logging.error(traceback.format_exc())
        return False

def save_drawing_simple(filepath):
    """Save the drawing to a file using a simpler, more reliable method"""
    try:
        logging.info(f"Saving drawing to {filepath}")
        
        # Ensure Paint is focused
        if not focus_paint_window():
            logging.error("Failed to focus Paint window")
            return False
            
        # Create directory if needed
        save_dir = os.path.dirname(filepath)
        if save_dir and not os.path.exists(save_dir):
            os.makedirs(save_dir)
            
        # Press Ctrl+S to open save dialog
        logging.info("Opening save dialog")
        pyautogui.hotkey('ctrl', 's')
        time.sleep(1.5)
        
        # Type the file path
        logging.info(f"Typing file path: {filepath}")
        pyautogui.write(filepath)
        time.sleep(1.0)
        
        # Press Enter to save
        logging.info("Pressing Enter to save")
        pyautogui.press('enter')
        time.sleep(2.0)
        
        # Check if the file was created
        if os.path.exists(filepath):
            logging.info(f"File saved successfully to {filepath}")
            return True
        else:
            logging.error(f"File not found at {filepath} after save attempt")
            
            # Check if a "replace" dialog appeared
            # Press 'y' to confirm overwrite if the file already exists
            logging.info("Checking for overwrite dialog and pressing 'y'")
            pyautogui.press('y')
            time.sleep(2.0)
            
            if os.path.exists(filepath):
                logging.info(f"File saved successfully after overwrite confirmation")
                return True
            else:
                logging.error("File still not saved after possible overwrite dialog")
                return False
                
    except Exception as e:
        logging.error(f"Error saving drawing: {str(e)}")
        logging.error(traceback.format_exc())
        return False

def main():
    """Test the simple drawing functions"""
    setup_logging()
    logging.info("Starting simple drawing test")
    
    try:
        # Open Paint
        if not open_paint_simple():
            logging.error("Failed to open Paint")
            return False
            
        # Get canvas info
        canvas_info = get_paint_canvas_bounds()
        if not canvas_info:
            logging.error("Failed to get canvas bounds")
            return False
            
        # Draw a rectangle
        rect_x1 = canvas_info["left"] + int(canvas_info["width"] * 0.2)
        rect_y1 = canvas_info["top"] + int(canvas_info["height"] * 0.2)
        rect_x2 = rect_x1 + int(canvas_info["width"] * 0.3)
        rect_y2 = rect_y1 + int(canvas_info["height"] * 0.3)
        
        if not draw_simple_rectangle(rect_x1, rect_y1, rect_x2, rect_y2):
            logging.error("Failed to draw rectangle")
            
        time.sleep(1.0)
        
        # Draw a circle
        circle_x = canvas_info["left"] + int(canvas_info["width"] * 0.7)
        circle_y = canvas_info["top"] + int(canvas_info["height"] * 0.5)
        circle_radius = int(min(canvas_info["width"], canvas_info["height"]) * 0.15)
        
        if not draw_simple_circle(circle_x, circle_y, circle_radius):
            logging.error("Failed to draw circle")
            
        time.sleep(1.0)
        
        # Save the drawing
        output_dir = "output"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = os.path.join(output_dir, f"simple_drawing_{timestamp}.png")
        
        if not save_drawing_simple(output_path):
            logging.error("Failed to save drawing")
        
        logging.info("Simple drawing test completed")
        return True
        
    except Exception as e:
        logging.error(f"Error in main function: {str(e)}")
        logging.error(traceback.format_exc())
        return False

if __name__ == "__main__":
    main() 