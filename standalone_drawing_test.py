"""
Enhanced standalone test for Paint automation functions
This script draws a simple picture with multiple shapes
"""

import subprocess
import time
import sys
import os
import logging
import math
import random
import pyautogui
import win32gui
import win32con

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("enhanced_test.log"),
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

def draw_shape(start_x, start_y, end_x, end_y):
    """Draw a shape from start to end coordinates"""
    try:
        logging.info(f"Drawing shape from ({start_x}, {start_y}) to ({end_x}, {end_y})")
        
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
        
        logging.info("Shape drawing completed")
        return True
    except Exception as e:
        logging.error(f"Error drawing shape: {str(e)}")
        return False

def draw_line(start_x, start_y, end_x, end_y):
    """Draw a line from start to end coordinates"""
    return draw_shape(start_x, start_y, end_x, end_y)

def click_color_palette(color_x, color_y):
    """Click on a specific position in the color palette"""
    try:
        logging.info(f"Clicking color at ({color_x}, {color_y})")
        pyautogui.click(color_x, color_y)
        time.sleep(0.5)
        return True
    except Exception as e:
        logging.error(f"Error selecting color: {str(e)}")
        return False

def draw_house_picture():
    """Draw a simple house picture"""
    try:
        logging.info("Starting to draw a house picture")
        
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
        
        # Calculate canvas dimensions
        canvas_width = canvas_info['right'] - canvas_info['left']
        canvas_height = canvas_info['bottom'] - canvas_info['top']
        center_x = canvas_info['left'] + (canvas_width // 2)
        center_y = canvas_info['top'] + (canvas_height // 2)
        
        # Step 4: Select a blue color for the sky (click in the color palette)
        # This is an approximation - colors are in the top ribbon
        window_width = canvas_info['right'] - canvas_info['left'] + 80
        blue_color_x = canvas_info['left'] + (window_width * 0.2)  # Approximate position
        blue_color_y = canvas_info['top'] - 80  # Above the canvas in the ribbon
        
        click_color_palette(blue_color_x, blue_color_y)
        
        # Step 5: Draw house base (rectangle)
        house_width = min(300, canvas_width // 3)
        house_height = min(200, canvas_height // 3)
        
        house_left = center_x - (house_width // 2)
        house_right = center_x + (house_width // 2)
        house_top = center_y
        house_bottom = house_top + house_height
        
        draw_shape(house_left, house_top, house_right, house_bottom)
        time.sleep(1)
        
        # Step 6: Select a red color for the roof
        red_color_x = canvas_info['left'] + (window_width * 0.15)  # Approximate position
        red_color_y = canvas_info['top'] - 80  # Above the canvas in the ribbon
        
        click_color_palette(red_color_x, red_color_y)
        
        # Step 7: Draw roof (triangle) - we'll simulate it with lines
        roof_left = house_left - 20
        roof_right = house_right + 20
        roof_top = house_top - 100
        
        # First line of roof
        draw_line(roof_left, house_top, center_x, roof_top)
        time.sleep(0.5)
        
        # Second line of roof
        draw_line(center_x, roof_top, roof_right, house_top)
        time.sleep(0.5)
        
        # Step 8: Select a brown color for the door
        brown_color_x = canvas_info['left'] + (window_width * 0.25)  # Approximate position
        brown_color_y = canvas_info['top'] - 80  # Above the canvas in the ribbon
        
        click_color_palette(brown_color_x, brown_color_y)
        
        # Step 9: Draw door (rectangle)
        door_width = house_width // 5
        door_left = center_x - (door_width // 2)
        door_right = door_left + door_width
        door_top = house_bottom - (house_height // 2)
        
        draw_shape(door_left, door_top, door_right, house_bottom)
        time.sleep(1)
        
        # Step 10: Select a yellow color for windows
        yellow_color_x = canvas_info['left'] + (window_width * 0.18)  # Approximate position
        yellow_color_y = canvas_info['top'] - 80  # Above the canvas in the ribbon
        
        click_color_palette(yellow_color_x, yellow_color_y)
        
        # Step 11: Draw left window (square)
        window_size = house_width // 8
        left_window_left = house_left + (house_width // 5) - (window_size // 2)
        left_window_right = left_window_left + window_size
        window_top = house_top + (house_height // 5)
        window_bottom = window_top + window_size
        
        draw_shape(left_window_left, window_top, left_window_right, window_bottom)
        time.sleep(1)
        
        # Step 12: Draw right window (square)
        right_window_left = house_right - (house_width // 5) - (window_size // 2)
        right_window_right = right_window_left + window_size
        
        draw_shape(right_window_left, window_top, right_window_right, window_bottom)
        time.sleep(1)
        
        # Step 13: Draw sun
        # Select yellow again for the sun
        click_color_palette(yellow_color_x, yellow_color_y)
        
        sun_radius = min(80, house_width // 4)
        sun_x = house_right + sun_radius + 50
        sun_y = roof_top
        
        # Draw the sun (circle)
        sun_left = sun_x - sun_radius
        sun_top = sun_y - sun_radius
        sun_right = sun_x + sun_radius
        sun_bottom = sun_y + sun_radius
        
        draw_shape(sun_left, sun_top, sun_right, sun_bottom)
        time.sleep(1)
        
        # Step 14: Draw sun rays
        for i in range(8):
            angle = i * 45
            ray_length = sun_radius * 0.8
            
            ray_end_x = sun_x + ray_length * math.cos(math.radians(angle))
            ray_end_y = sun_y + ray_length * math.sin(math.radians(angle))
            
            draw_line(sun_x, sun_y, ray_end_x, ray_end_y)
            time.sleep(0.5)
        
        # Step 15: Draw grass at bottom
        # Select green color
        green_color_x = canvas_info['left'] + (window_width * 0.22)  # Approximate position
        green_color_y = canvas_info['top'] - 80  # Above the canvas in the ribbon
        
        click_color_palette(green_color_x, green_color_y)
        
        grass_top = house_bottom + 20
        grass_bottom = grass_top + 50
        grass_left = canvas_info['left'] + 100
        grass_right = canvas_info['right'] - 100
        
        draw_shape(grass_left, grass_top, grass_right, grass_bottom)
        
        # Wait to view the result
        time.sleep(5)
        
        # Step 16: Close Paint
        close_paint()
        
        logging.info("House drawing completed successfully")
        return True
    except Exception as e:
        logging.error(f"Error drawing house: {str(e)}")
        try:
            close_paint()
        except:
            pass
        return False

if __name__ == "__main__":
    print("Starting enhanced drawing test...")
    success = draw_house_picture()
    print(f"\nTest {'PASSED' if success else 'FAILED'}")
    print("Check the log file for details: enhanced_test.log") 