"""
Final Paint Drawing Agent Test Suite

This script provides a comprehensive test suite for MS Paint automation,
including both simple and advanced drawing tests.
"""

import subprocess
import time
import sys
import os
import logging
import math
import random
from datetime import datetime
import pyautogui
import win32gui
import win32con

# Configure logging with timestamps
log_file = f"paint_test_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
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

# Test functions
def run_quick_test():
    """Run a quick test of basic Paint functionality"""
    try:
        logging.info("----- Starting Quick Test -----")
        
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
        
        # Step 4: Draw a simple rectangle in the center
        canvas_width = canvas_info['right'] - canvas_info['left']
        canvas_height = canvas_info['bottom'] - canvas_info['top']
        
        center_x = canvas_info['left'] + (canvas_width // 2)
        center_y = canvas_info['top'] + (canvas_height // 2)
        
        rect_size = min(200, canvas_width // 3, canvas_height // 3)
        
        rect_left = center_x - (rect_size // 2)
        rect_top = center_y - (rect_size // 2)
        rect_right = center_x + (rect_size // 2)
        rect_bottom = center_y + (rect_size // 2)
        
        if not draw_shape(rect_left, rect_top, rect_right, rect_bottom):
            logging.error("Failed to draw rectangle. Test failed.")
            close_paint()
            return False
        
        logging.info("Quick test: Drew rectangle successfully")
        time.sleep(3)
        
        # Step 5: Close Paint
        close_paint()
        
        logging.info("Quick test completed successfully")
        return True
    except Exception as e:
        logging.error(f"Error in quick test: {str(e)}")
        try:
            close_paint()
        except:
            pass
        return False

def draw_shape_grid():
    """Draw a grid of different shapes"""
    try:
        logging.info("----- Starting Shape Grid Test -----")
        
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
        
        # Define grid
        rows, cols = 3, 3
        margin = 50
        
        cell_width = (canvas_width - 2 * margin) / cols
        cell_height = (canvas_height - 2 * margin) / rows
        
        # Draw grid of shapes
        for row in range(rows):
            for col in range(cols):
                # Calculate cell position
                cell_left = canvas_info['left'] + margin + (col * cell_width)
                cell_top = canvas_info['top'] + margin + (row * cell_height)
                cell_right = cell_left + cell_width - 20
                cell_bottom = cell_top + cell_height - 20
                
                # Select a random shape type (represented by drawing differently)
                shape_type = random.randint(0, 2)
                
                # Select a random color
                color_offset_x = random.randint(15, 25) * 10
                color_offset_y = random.randint(7, 10) * 10
                color_x = canvas_info['left'] + color_offset_x
                color_y = canvas_info['top'] - color_offset_y
                
                click_color_palette(color_x, color_y)
                
                if shape_type == 0:
                    # Rectangle
                    draw_shape(cell_left, cell_top, cell_right, cell_bottom)
                elif shape_type == 1:
                    # Circle
                    center_x = (cell_left + cell_right) / 2
                    center_y = (cell_top + cell_bottom) / 2
                    radius = min(cell_width, cell_height) / 3
                    
                    circle_left = center_x - radius
                    circle_top = center_y - radius
                    circle_right = center_x + radius
                    circle_bottom = center_y + radius
                    
                    draw_shape(circle_left, circle_top, circle_right, circle_bottom)
                else:
                    # Triangle (simulate with lines)
                    # First line
                    draw_line(cell_left, cell_bottom, (cell_left + cell_right) / 2, cell_top)
                    time.sleep(0.5)
                    # Second line
                    draw_line((cell_left + cell_right) / 2, cell_top, cell_right, cell_bottom)
                    time.sleep(0.5)
                    # Third line
                    draw_line(cell_right, cell_bottom, cell_left, cell_bottom)
                
                time.sleep(1)
        
        # Wait to view the result
        time.sleep(5)
        
        # Close Paint
        close_paint()
        
        logging.info("Shape grid test completed successfully")
        return True
    except Exception as e:
        logging.error(f"Error in shape grid test: {str(e)}")
        try:
            close_paint()
        except:
            pass
        return False

def draw_complex_scene():
    """Draw a more complex scene with multiple elements"""
    try:
        logging.info("----- Starting Complex Scene Test -----")
        
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
        
        # Draw sky (light blue rectangle)
        sky_color_x = canvas_info['left'] + (canvas_width * 0.2)
        sky_color_y = canvas_info['top'] - 80
        click_color_palette(sky_color_x, sky_color_y)
        
        draw_shape(canvas_info['left'], canvas_info['top'], 
                   canvas_info['right'], canvas_info['top'] + (canvas_height * 0.7))
        time.sleep(1)
        
        # Draw ground (green rectangle)
        ground_color_x = canvas_info['left'] + (canvas_width * 0.22)
        ground_color_y = canvas_info['top'] - 80
        click_color_palette(ground_color_x, ground_color_y)
        
        draw_shape(canvas_info['left'], canvas_info['top'] + (canvas_height * 0.7),
                   canvas_info['right'], canvas_info['bottom'])
        time.sleep(1)
        
        # Draw sun
        sun_color_x = canvas_info['left'] + (canvas_width * 0.18)
        sun_color_y = canvas_info['top'] - 80
        click_color_palette(sun_color_x, sun_color_y)
        
        sun_radius = min(80, canvas_width // 10)
        sun_x = canvas_info['left'] + (canvas_width * 0.8)
        sun_y = canvas_info['top'] + (canvas_height * 0.2)
        
        sun_left = sun_x - sun_radius
        sun_top = sun_y - sun_radius
        sun_right = sun_x + sun_radius
        sun_bottom = sun_y + sun_radius
        
        draw_shape(sun_left, sun_top, sun_right, sun_bottom)
        time.sleep(1)
        
        # Draw mountains
        mountain_color_x = canvas_info['left'] + (canvas_width * 0.25)
        mountain_color_y = canvas_info['top'] - 80
        click_color_palette(mountain_color_x, mountain_color_y)
        
        # Mountain 1
        m1_left = canvas_info['left'] + (canvas_width * 0.1)
        m1_bottom = canvas_info['top'] + (canvas_height * 0.7)
        m1_peak_x = m1_left + (canvas_width * 0.15)
        m1_peak_y = canvas_info['top'] + (canvas_height * 0.3)
        m1_right = m1_left + (canvas_width * 0.3)
        
        draw_line(m1_left, m1_bottom, m1_peak_x, m1_peak_y)
        time.sleep(0.5)
        draw_line(m1_peak_x, m1_peak_y, m1_right, m1_bottom)
        time.sleep(1)
        
        # Mountain 2
        m2_left = canvas_info['left'] + (canvas_width * 0.25)
        m2_bottom = canvas_info['top'] + (canvas_height * 0.7)
        m2_peak_x = m2_left + (canvas_width * 0.2)
        m2_peak_y = canvas_info['top'] + (canvas_height * 0.25)
        m2_right = m2_left + (canvas_width * 0.4)
        
        draw_line(m2_left, m2_bottom, m2_peak_x, m2_peak_y)
        time.sleep(0.5)
        draw_line(m2_peak_x, m2_peak_y, m2_right, m2_bottom)
        time.sleep(1)
        
        # Mountain 3
        m3_left = canvas_info['left'] + (canvas_width * 0.5)
        m3_bottom = canvas_info['top'] + (canvas_height * 0.7)
        m3_peak_x = m3_left + (canvas_width * 0.15)
        m3_peak_y = canvas_info['top'] + (canvas_height * 0.35)
        m3_right = m3_left + (canvas_width * 0.3)
        
        draw_line(m3_left, m3_bottom, m3_peak_x, m3_peak_y)
        time.sleep(0.5)
        draw_line(m3_peak_x, m3_peak_y, m3_right, m3_bottom)
        time.sleep(1)
        
        # Draw a simple house
        house_left = center_x - (canvas_width * 0.1)
        house_right = center_x + (canvas_width * 0.1)
        house_bottom = canvas_info['top'] + (canvas_height * 0.8)
        house_top = house_bottom - (canvas_height * 0.2)
        
        # House base (brown rectangle)
        house_color_x = canvas_info['left'] + (canvas_width * 0.3)
        house_color_y = canvas_info['top'] - 80
        click_color_palette(house_color_x, house_color_y)
        
        draw_shape(house_left, house_top, house_right, house_bottom)
        time.sleep(1)
        
        # House roof (red triangle)
        roof_color_x = canvas_info['left'] + (canvas_width * 0.15)
        roof_color_y = canvas_info['top'] - 80
        click_color_palette(roof_color_x, roof_color_y)
        
        roof_peak_y = house_top - (canvas_height * 0.1)
        
        draw_line(house_left - (canvas_width * 0.02), house_top, center_x, roof_peak_y)
        time.sleep(0.5)
        draw_line(center_x, roof_peak_y, house_right + (canvas_width * 0.02), house_top)
        time.sleep(1)
        
        # Wait to view the result
        time.sleep(5)
        
        # Close Paint
        close_paint()
        
        logging.info("Complex scene test completed successfully")
        return True
    except Exception as e:
        logging.error(f"Error in complex scene test: {str(e)}")
        try:
            close_paint()
        except:
            pass
        return False

# Main function to run tests
def run_all_tests():
    """Run all Paint automation tests"""
    tests = [
        {"name": "Quick Test", "function": run_quick_test},
        {"name": "Shape Grid Test", "function": draw_shape_grid},
        {"name": "Complex Scene Test", "function": draw_complex_scene}
    ]
    
    results = {}
    
    print(f"\n===== Paint Drawing Agent Test Suite =====")
    print(f"Log file: {log_file}")
    print(f"\nRunning {len(tests)} tests...\n")
    
    for i, test in enumerate(tests):
        test_name = test["name"]
        test_function = test["function"]
        
        print(f"Test {i+1}/{len(tests)}: {test_name}...")
        
        try:
            start_time = time.time()
            success = test_function()
            end_time = time.time()
            
            duration = end_time - start_time
            status = "PASSED" if success else "FAILED"
            
            results[test_name] = {
                "success": success,
                "duration": duration
            }
            
            print(f"  {status} ({duration:.2f} seconds)\n")
            
        except Exception as e:
            logging.error(f"Error executing test '{test_name}': {str(e)}")
            results[test_name] = {
                "success": False,
                "duration": 0
            }
            print(f"  FAILED (Exception occurred)\n")
    
    # Print summary
    print("\n===== Test Results Summary =====")
    passed = sum(1 for r in results.values() if r["success"])
    failed = len(results) - passed
    
    print(f"Total tests: {len(results)}")
    print(f"Passed: {passed}")
    print(f"Failed: {failed}")
    print("\nDetailed results:")
    
    for test_name, result in results.items():
        status = "PASSED" if result["success"] else "FAILED"
        print(f"  {test_name}: {status} ({result['duration']:.2f} seconds)")
    
    print(f"\nLog file: {log_file}")
    
    return passed == len(results)

if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1) 