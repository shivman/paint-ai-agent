"""
Simple MS Paint drawing script with minimal dependencies.
This script directly uses pyautogui and win32gui to interact with MS Paint.
"""

import os
import time
import sys
import logging
import pyautogui
import win32gui
import win32con

def setup_logging():
    """Set up logging"""
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    
    log_file = os.path.join(log_dir, "draw_simple.log")
    
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
        return paint_window
    else:
        logging.error("Could not find Paint window")
        return None

def open_paint():
    """Open MS Paint"""
    logging.info("Opening MS Paint")
    try:
        # Try to find existing Paint window
        paint_window = find_paint_window()
        
        if paint_window:
            logging.info("Paint is already open")
            return paint_window
            
        # Open Paint if not already open
        os.system("start mspaint")
        time.sleep(2)
        
        # Find the Paint window
        paint_window = find_paint_window()
        if paint_window:
            logging.info("Successfully opened Paint")
            return paint_window
        else:
            logging.error("Failed to open Paint")
            return None
            
    except Exception as e:
        logging.error(f"Error opening Paint: {str(e)}")
        return None

def focus_paint_window(paint_window):
    """Focus and maximize the Paint window"""
    try:
        logging.info(f"Attempting to focus Paint window: {paint_window}")
        
        # First try the standard way
        # Check if window is minimized
        placement = win32gui.GetWindowPlacement(paint_window)
        if placement[1] == win32con.SW_SHOWMINIMIZED:
            logging.info("Paint window was minimized, restoring...")
            win32gui.ShowWindow(paint_window, win32con.SW_RESTORE)
            time.sleep(1)
            
        # Try multiple methods to bring window to front
        for attempt in range(3):
            logging.info(f"Focus attempt {attempt+1}/3")
            
            # Bring window to front
            win32gui.SetForegroundWindow(paint_window)
            time.sleep(1)
            
            # Maximize window
            win32gui.ShowWindow(paint_window, win32con.SW_MAXIMIZE)
            time.sleep(1)
            
            # Alternative method with Alt+Tab simulation
            pyautogui.keyDown('alt')
            time.sleep(0.2)
            pyautogui.press('tab')
            time.sleep(0.2)
            pyautogui.keyUp('alt')
            time.sleep(1)
            
            # Check if successfully focused
            active_window = win32gui.GetForegroundWindow()
            if active_window == paint_window:
                logging.info("Successfully focused Paint window")
                return True
                
            # If not successful, try one more approach
            pyautogui.click(100, 100)  # Click in a likely Paint window area
            time.sleep(0.5)
        
        # Last resort: just proceed anyway
        logging.warning("Could not focus Paint window, continuing anyway")
        return True
            
    except Exception as e:
        logging.error(f"Error focusing Paint window: {str(e)}")
        # Proceed anyway as a last resort
        return True

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
        focus_paint_window(paint_hwnd)
        
        # Get window dimensions
        window_rect = win32gui.GetWindowRect(paint_hwnd)
        win_x, win_y, win_right, win_bottom = window_rect
        win_width = win_right - win_x
        win_height = win_bottom - win_y
        logging.info(f"Paint window: {win_width}x{win_height} at ({win_x},{win_y})")
        
        # Take screenshot before for comparison
        screenshot_before = pyautogui.screenshot()
        
        # Clear any active tool with Escape
        pyautogui.press('esc')
        time.sleep(1)
        
        # IMPORTANT: Correctly select rectangle shape tool
        logging.info("Selecting rectangle shape tool (step-by-step)")
        
        # Step 1: First click the HOME tab to make sure we're in the right tab
        home_tab_x = win_x + 50  # Home tab is typically on the left side
        home_tab_y = win_y + 30  # Tabs are at the top
        logging.info(f"Clicking HOME tab at ({home_tab_x}, {home_tab_y})")
        pyautogui.moveTo(home_tab_x, home_tab_y, duration=0.5)
        time.sleep(0.5)
        pyautogui.click()
        time.sleep(1)
        
        # Step 2: Click on "Shapes" button in ribbon (typically ~35% from left)
        shapes_button_x = win_x + int(win_width * 0.35)
        shapes_button_y = win_y + 75
        logging.info(f"Clicking SHAPES button at ({shapes_button_x}, {shapes_button_y})")
        pyautogui.moveTo(shapes_button_x, shapes_button_y, duration=0.5)
        time.sleep(0.5)
        pyautogui.click()
        time.sleep(1.5)  # Longer wait for shapes menu to appear
        
        # Step 3: Click on Rectangle (first shape in dropdown)
        # Rectangle is typically the first option in the dropdown
        rect_x = shapes_button_x
        rect_y = shapes_button_y + 40
        logging.info(f"Clicking RECTANGLE shape at ({rect_x}, {rect_y})")
        pyautogui.moveTo(rect_x, rect_y, duration=0.5)
        time.sleep(0.5)
        pyautogui.click()
        time.sleep(1.5)
        
        # Verify tool selection by clicking in a neutral area
        neutral_x = win_x + win_width//2
        neutral_y = win_y + 120
        pyautogui.moveTo(neutral_x, neutral_y, duration=0.3)
        time.sleep(0.5)
        
        # Calculate canvas area (typical Paint canvas starts ~170px from top)
        canvas_top = win_y + 170
        canvas_left = win_x + 10
        
        # Convert coordinates
        screen_x1 = canvas_left + x1
        screen_y1 = canvas_top + y1
        screen_x2 = canvas_left + x2
        screen_y2 = canvas_top + y2
        
        logging.info(f"Drawing rectangle from ({screen_x1},{screen_y1}) to ({screen_x2},{screen_y2})")
        
        # Move to start position with visual feedback
        logging.info(f"Moving to start position ({screen_x1}, {screen_y1})")
        
        # First, move to a neutral position then to start
        pyautogui.moveTo(canvas_left + 50, canvas_top + 50, duration=0.5)
        time.sleep(0.5)
        
        # Make a small circle at start point for visibility
        for _ in range(3):  # More circles for better visibility
            offset = 5
            pyautogui.moveTo(screen_x1 + offset, screen_y1, duration=0.1)
            pyautogui.moveTo(screen_x1, screen_y1 + offset, duration=0.1)
            pyautogui.moveTo(screen_x1 - offset, screen_y1, duration=0.1)
            pyautogui.moveTo(screen_x1, screen_y1 - offset, duration=0.1)
        
        # Position at exact start
        pyautogui.moveTo(screen_x1, screen_y1, duration=0.5)
        time.sleep(1)
        
        # Press mouse button with EXTRA emphasis
        logging.info("PRESSING MOUSE DOWN")
        pyautogui.mouseDown(button='left')
        time.sleep(1)  # Hold longer
        
        # Move to end position very slowly (first horizontal, then vertical for clear rectangle)
        logging.info(f"Moving to end position ({screen_x2}, {screen_y2})")
        pyautogui.moveTo(screen_x2, screen_y1, duration=1.0)  # First go horizontally
        time.sleep(0.5)  
        pyautogui.moveTo(screen_x2, screen_y2, duration=1.0)  # Then go vertically
        time.sleep(1)  # Hold at final position
        
        # Release mouse button
        logging.info("RELEASING MOUSE UP")
        pyautogui.mouseUp(button='left')
        time.sleep(1)
        
        # Move away to see result clearly
        pyautogui.moveTo(canvas_left + 50, canvas_top + 50, duration=0.5)
        time.sleep(1)
        
        # Take screenshot after for comparison
        screenshot_after = pyautogui.screenshot()
        
        # Compare screenshots
        changed_pixels = 0
        sample_points = []
        
        # Check border pixels specifically (more likely to change with rectangle)
        # Top border
        for x in range(min(screen_x1, screen_x2), max(screen_x1, screen_x2), 10):
            try:
                if screenshot_before.getpixel((x, screen_y1)) != screenshot_after.getpixel((x, screen_y1)):
                    changed_pixels += 1
                    if len(sample_points) < 5:
                        sample_points.append((x, screen_y1))
            except Exception:
                pass
                
        # Bottom border
        for x in range(min(screen_x1, screen_x2), max(screen_x1, screen_x2), 10):
            try:
                if screenshot_before.getpixel((x, screen_y2)) != screenshot_after.getpixel((x, screen_y2)):
                    changed_pixels += 1
                    if len(sample_points) < 5:
                        sample_points.append((x, screen_y2))
            except Exception:
                pass
                
        # Left border
        for y in range(min(screen_y1, screen_y2), max(screen_y1, screen_y2), 10):
            try:
                if screenshot_before.getpixel((screen_x1, y)) != screenshot_after.getpixel((screen_x1, y)):
                    changed_pixels += 1
                    if len(sample_points) < 5:
                        sample_points.append((screen_x1, y))
            except Exception:
                pass
                
        # Right border
        for y in range(min(screen_y1, screen_y2), max(screen_y1, screen_y2), 10):
            try:
                if screenshot_before.getpixel((screen_x2, y)) != screenshot_after.getpixel((screen_x2, y)):
                    changed_pixels += 1
                    if len(sample_points) < 5:
                        sample_points.append((screen_x2, y))
            except Exception:
                pass
        
        # Save debug screenshots regardless of result
        debug_dir = "debug_screenshots"
        if not os.path.exists(debug_dir):
            os.makedirs(debug_dir)
            
        timestamp = int(time.time())
        before_path = os.path.join(debug_dir, f"before_{timestamp}.png")
        after_path = os.path.join(debug_dir, f"after_{timestamp}.png")
        screenshot_before.save(before_path)
        screenshot_after.save(after_path)
        logging.info(f"Saved debug screenshots: {before_path}, {after_path}")
                
        logging.info(f"Detected {changed_pixels} changed pixels on rectangle borders")
        for point in sample_points:
            try:
                before_color = screenshot_before.getpixel(point)
                after_color = screenshot_after.getpixel(point)
                logging.info(f"Pixel changed at {point}: {before_color} -> {after_color}")
            except Exception:
                pass
        
        if changed_pixels > 5:
            logging.info("Rectangle appears to have been drawn successfully")
            return True
        else:
            logging.error("No significant pixel changes detected, rectangle likely not drawn")
            return False
            
    except Exception as e:
        logging.error(f"Error drawing rectangle: {str(e)}", exc_info=True)
        return False

def main():
    """Main function to draw a rectangle"""
    setup_logging()
    logging.info("Starting simple MS Paint drawing test")
    
    try:
        # Draw a rectangle from (100,100) to (300,300)
        logging.info("Drawing rectangle from (100,100) to (300,300)")
        result = draw_rectangle_direct(100, 100, 300, 300)
        
        if result:
            logging.info("Successfully drew rectangle")
            print("SUCCESS: Rectangle drawn")
        else:
            logging.error("Failed to draw rectangle")
            print("ERROR: Failed to draw rectangle")
            
    except Exception as e:
        logging.error(f"Error in main: {str(e)}")
        print(f"ERROR: {str(e)}")
        
    logging.info("Test completed")

if __name__ == "__main__":
    main() 