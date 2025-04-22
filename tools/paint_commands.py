import os
import time
import win32api
import win32con
import pyautogui
from pywinauto.keyboard import send_keys
from pywinauto import Application, timings
import win32gui
import subprocess
import math
import logging
from datetime import datetime
import sys
from pathlib import Path

# Configure logging
log_dir = "logs"
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

log_file = os.path.join(log_dir, f"paint_operations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler()
    ]
)

# Configure PyAutoGUI
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5  # Add delay between actions

def log_window_info():
    """Log information about the Paint window and screen."""
    try:
        paint_window = win32gui.FindWindow(None, "Untitled - Paint")
        if paint_window:
            window_rect = win32gui.GetWindowRect(paint_window)
            screen_width, screen_height = pyautogui.size()
            
            logging.info(f"Screen resolution: {screen_width}x{screen_height}")
            logging.info(f"Paint window rectangle: {window_rect}")
            logging.info(f"Paint window dimensions: {window_rect[2]-window_rect[0]}x{window_rect[3]-window_rect[1]}")
            
            # Get window state
            placement = win32gui.GetWindowPlacement(paint_window)
            state = "Normal"
            if placement[1] == win32con.SW_SHOWMAXIMIZED:
                state = "Maximized"
            elif placement[1] == win32con.SW_SHOWMINIMIZED:
                state = "Minimized"
            logging.info(f"Paint window state: {state}")
            
            return True
    except Exception as e:
        logging.error(f"Error getting window info: {str(e)}")
        return False

def verify_mouse_position(expected_x, expected_y):
    """Verify if mouse is at the expected position."""
    actual_x, actual_y = pyautogui.position()
    logging.info(f"Mouse position - Expected: ({expected_x}, {expected_y}), Actual: ({actual_x}, {actual_y})")
    return abs(actual_x - expected_x) < 5 and abs(actual_y - expected_y) < 5

def highlight_click_location(x, y, duration=0.5):
    """Visually highlight where the mouse will click"""
    current_x, current_y = pyautogui.position()
    # Move to location with visual feedback
    pyautogui.moveTo(x, y, duration=duration)
    # Small circular motion to highlight the spot
    pyautogui.moveRel(5, 0, duration=0.1)
    pyautogui.moveRel(0, 5, duration=0.1)
    pyautogui.moveRel(-5, 0, duration=0.1)
    pyautogui.moveRel(0, -5, duration=0.1)
    return True

def safe_click(x, y, description="", highlight=True):
    """Safely perform a click with visual feedback"""
    try:
        if highlight:
            highlight_click_location(x, y)
        current_x, current_y = pyautogui.position()
        logging.info(f"Clicking {description} at ({x}, {y}). Current mouse at ({current_x}, {current_y})")
        pyautogui.click(x, y)
        time.sleep(0.5)  # Wait after click
        return True
    except Exception as e:
        logging.error(f"Failed to click at ({x}, {y}): {str(e)}")
        return False

def find_paint_window():
    """Find classic Paint window using multiple methods"""
    try:
        # Try different window class and title combinations
        window_specs = [
            ("MSPaintApp", None),  # Classic Paint class name
            (None, "Untitled - Paint"),  # Default title
            (None, "Paint"),  # Alternative title
        ]
        
        for class_name, title in window_specs:
            hwnd = win32gui.FindWindow(class_name, title)
            if hwnd and win32gui.IsWindowVisible(hwnd):
                # Verify this is classic Paint by checking window class
                window_class = win32gui.GetClassName(hwnd)
                logging.info(f"Found Paint window with class: {window_class}")
                if window_class == "MSPaintApp":
                    logging.info("Confirmed classic MSPaint application")
                return hwnd
                
        # Last resort: search by partial title and class
        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                window_class = win32gui.GetClassName(hwnd)
                if "Paint" in title and window_class == "MSPaintApp":
                    logging.info(f"Found Paint window via enumeration: {title}, class: {window_class}")
                    results.append(hwnd)
                    
        paint_windows = []
        win32gui.EnumWindows(enum_callback, paint_windows)
        
        if paint_windows:
            return paint_windows[0]
            
        return None
    except Exception as e:
        logging.error(f"Error finding Paint window: {str(e)}")
        return None

def ensure_paint_focused():
    """Ensure Paint window is focused and maximized"""
    try:
        # Find Paint window
        paint_window = find_paint_window()
        if not paint_window:
            logging.error("Paint window not found")
            return False
            
        # Get window info
        window_rect = win32gui.GetWindowRect(paint_window)
        logging.info(f"Found Paint window at {window_rect}")
        
        # Multiple attempts to focus window
        for attempt in range(3):
            try:
                # Restore window if minimized
                win32gui.ShowWindow(paint_window, win32con.SW_RESTORE)
                time.sleep(0.5)
                
                # Set foreground window
                win32gui.SetForegroundWindow(paint_window)
                time.sleep(0.5)
                
                # Maximize window
                win32gui.ShowWindow(paint_window, win32con.SW_MAXIMIZE)
                time.sleep(0.5)
                
                # Verify window is active
                active_window = win32gui.GetForegroundWindow()
                if active_window == paint_window:
                    # Additional verification - try to get window text
                    try:
                        win32gui.GetWindowText(active_window)
                        logging.info("Paint window successfully focused")
                        return True
                    except:
                        continue
                        
            except Exception as e:
                logging.warning(f"Focus attempt {attempt + 1} failed: {str(e)}")
                time.sleep(1)
                
        logging.error("Failed to focus Paint window after multiple attempts")
        return False
        
    except Exception as e:
        logging.error(f"Error in ensure_paint_focused: {str(e)}")
        return False

def open_paint():
    """Opens classic Microsoft Paint application."""
    print("Attempting to open classic Paint...")
    logging.info("Attempting to open classic Paint application")
    
    # First check if Paint is already running
    if ensure_paint_focused():
        print("Found existing Paint window and focused it")
        return True
    
    # Try opening Paint using Windows start command (more reliable)
    try:
        print("Opening classic Paint using Windows start command...")
        logging.info("Launching Paint using Windows start command")
        
        # Use Windows start command to launch mspaint.exe
        os.system('start mspaint')
        time.sleep(2)
        
        # Try to focus the window
        for attempt in range(5):  # Try multiple times with longer delays
            if ensure_paint_focused():
                # Verify this is classic Paint
                paint_window = find_paint_window()
                if paint_window:
                    window_class = win32gui.GetClassName(paint_window)
                    if window_class == "MSPaintApp":
                        print("Classic Paint opened and focused successfully")
                        logging.info(f"Classic Paint opened with window class: {window_class}")
                        return True
                    else:
                        print(f"Warning: Opened Paint with class {window_class}, not classic MSPaintApp")
                        # Continue anyway
                        return True
            
            # Wait longer between attempts
            time.sleep(1 + attempt)
            logging.info(f"Paint focus attempt {attempt+1} failed, retrying...")
            
        # If we couldn't focus the window, check if we can find it at all
        paint_windows = []
        
        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if "Paint" in title:
                    try:
                        window_class = win32gui.GetClassName(hwnd)
                        results.append((hwnd, title, window_class))
                    except:
                        results.append((hwnd, title, "Unknown"))
                        
        win32gui.EnumWindows(enum_callback, paint_windows)
        
        if paint_windows:
            logging.info(f"Found Paint windows but couldn't focus: {paint_windows}")
            print(f"Paint window found but couldn't be focused")
            return True
                
    except Exception as e:
        print(f"Failed to open classic Paint: {str(e)}")
        logging.error(f"Failed to open classic Paint: {str(e)}")
    
    print("Failed to open Paint")
    return False

def select_shape_tool(shape_type='rectangle'):
    """
    Select a shape tool in MS Paint based on the Windows 11 Paint UI.
    Supported shapes: rectangle, circle (ellipse), line, triangle
    """
    try:
        logging.info(f"Selecting shape tool: {shape_type}")
        
        # Get Paint window handle
        paint_window = find_paint_window()
        if not paint_window:
            logging.error("Failed to find Paint window")
            return False
            
        # Ensure Paint is in focus
        if not ensure_paint_focused():
            logging.error("Failed to focus Paint window")
            return False
        
        # Get window dimensions
        window_rect = win32gui.GetWindowRect(paint_window)
        win_x, win_y, win_right, win_bottom = window_rect
        win_width = win_right - win_x
        win_height = win_bottom - win_y
        logging.info(f"Paint window: {win_width}x{win_height} at ({win_x},{win_y})")
        
        # Clear any active tool selection
        pyautogui.press('esc')
        time.sleep(1)
        
        # Based on the Windows 11 Paint UI screenshot:
        # Shapes button is visible in the toolbar, around 1/3 from the left
        shapes_button_rel_x = 0.35  # Relative position from left edge (35%)
        shapes_button_y = win_y + 55  # Fixed Y position from top
        
        shapes_button_x = win_x + int(win_width * shapes_button_rel_x)
        
        logging.info(f"Clicking Shapes button at ({shapes_button_x}, {shapes_button_y})")
        # Move to shapes button with visual feedback
        pyautogui.moveTo(shapes_button_x, shapes_button_y, duration=0.5)
        time.sleep(0.5)
        pyautogui.click()
        time.sleep(1.5)  # Longer wait for shapes menu to appear
        
        # Map shape names to relative positions in the dropdown grid
        # These positions are based on Windows 11 Paint UI
        shape_positions = {
            'line': (0, 0),         # First item, first row
            'curve': (1, 0),        # Second item, first row
            'rectangle': (0, 1),    # First item, second row
            'square': (1, 1),       # Second item, second row
            'oval': (2, 1),         # Third item, second row
            'circle': (3, 1),       # Fourth item, second row
            'triangle': (0, 2),     # First item, third row
        }
        
        # Check if requested shape is supported
        if shape_type.lower() not in shape_positions:
            logging.warning(f"Shape type '{shape_type}' not explicitly supported, defaulting to rectangle")
            shape_type = 'rectangle'
        
        # Get grid position for the selected shape
        grid_x, grid_y = shape_positions[shape_type.lower()]
        
        # Calculate pixel offset from the shapes button
        # Each grid item is approximately 40px wide and 35px tall
        item_width = 40
        item_height = 35
        
        shape_x = shapes_button_x + (grid_x * item_width)
        shape_y = shapes_button_y + 40 + (grid_y * item_height)  # 40px initial dropdown offset
        
        logging.info(f"Clicking {shape_type} shape at ({shape_x}, {shape_y})")
        pyautogui.moveTo(shape_x, shape_y, duration=0.5)
        time.sleep(0.5)
        pyautogui.click()
        time.sleep(1.5)
        
        # Move to neutral area to complete selection
        neutral_x = win_x + win_width//2
        neutral_y = win_y + 160
        pyautogui.moveTo(neutral_x, neutral_y, duration=0.3)
        time.sleep(0.5)
        
        logging.info(f"Successfully selected {shape_type} shape tool")
        return True
        
    except Exception as e:
        logging.error(f"Error selecting shape tool: {str(e)}", exc_info=True)
        return False

def select_brush_tool():
    """
    Select the brush tool in MS Paint based on the Windows UI layout.
    """
    try:
        logging.info("Selecting brush tool")
        
        # Get Paint window handle
        paint_window = find_paint_window()
        if not paint_window:
            logging.error("Failed to find Paint window")
            return False
            
        # Ensure Paint is in focus
        if not ensure_paint_focused():
            logging.error("Failed to focus Paint window")
            return False
        
        # Get window dimensions
        window_rect = win32gui.GetWindowRect(paint_window)
        win_x, win_y, win_right, win_bottom = window_rect
        win_width = win_right - win_x
        win_height = win_bottom - win_y
        logging.info(f"Paint window: {win_width}x{win_height} at ({win_x},{win_y})")
        
        # Clear any active tool selection
        pyautogui.press('esc')
        time.sleep(0.5)
        
        # Calculate brush tool position (typically in the Brushes section)
        # In typical Paint layout, brushes are located around 15% from the left in the toolbar
        brush_button_rel_x = 0.15  # Relative position from left edge (15%)
        brush_button_y = win_y + 55  # Fixed Y position from top
        
        brush_button_x = win_x + int(win_width * brush_button_rel_x)
        
        logging.info(f"Clicking Brush button at ({brush_button_x}, {brush_button_y})")
        
        # Move to brush button with visual feedback
        highlight_click_location(brush_button_x, brush_button_y, duration=0.3)
        time.sleep(0.2)
        pyautogui.click()
        time.sleep(1.0)  # Wait for brush dropdown to appear
        
        # Select the first brush in the dropdown (usually the standard brush)
        # Calculate position for the first brush option
        first_brush_x = brush_button_x
        first_brush_y = brush_button_y + 40  # 40px down from button
        
        logging.info(f"Selecting first brush at ({first_brush_x}, {first_brush_y})")
        pyautogui.moveTo(first_brush_x, first_brush_y, duration=0.3)
        time.sleep(0.2)
        pyautogui.click()
        time.sleep(0.5)
        
        # Move to neutral area to complete selection
        neutral_x = win_x + win_width//2
        neutral_y = win_y + 160
        pyautogui.moveTo(neutral_x, neutral_y, duration=0.2)
        time.sleep(0.2)
        
        # Additional verification - try to click in the canvas area
        canvas_info = get_paint_canvas_info()
        if canvas_info:
            # Handle both tuple and dictionary formats
            if isinstance(canvas_info, tuple):
                canvas_x, canvas_y, canvas_width, canvas_height = canvas_info
                canvas_center_x = canvas_x + (canvas_width // 2)
                canvas_center_y = canvas_y + (canvas_height // 2)
            else:
                canvas_center_x = canvas_info['left'] + (canvas_info['width'] // 2)
                canvas_center_y = canvas_info['top'] + (canvas_info['height'] // 2)
                
            pyautogui.moveTo(canvas_center_x, canvas_center_y, duration=0.2)
            time.sleep(0.2)
        
        logging.info("Successfully selected brush tool")
        return True
        
    except Exception as e:
        logging.error(f"Error selecting brush tool: {str(e)}", exc_info=True)
        return False

def get_paint_canvas_info():
    """
    Get Paint canvas information with accurate boundaries based on Windows UI layout.
    
    Returns:
        For backwards compatibility, returns a tuple (x, y, width, height) representing 
        the canvas position and dimensions. The tuple object will also have keys and 
        values accessible via dict-like access for existing code that expects a dictionary.
    """
    try:
        # Find Paint window using robust method
        paint_window = find_paint_window()
        if not paint_window:
            logging.error("Paint window not found")
            return None
            
        # Get window dimensions
        window_rect = win32gui.GetWindowRect(paint_window)
        win_x, win_y, win_right, win_bottom = window_rect
        win_width = win_right - win_x
        win_height = win_bottom - win_y
        
        logging.info(f"Paint window dimensions: {win_width}x{win_height} at ({win_x}, {win_y})")
        
        # Calculate canvas area based on Windows Paint UI
        title_height = 32       # Top window title bar
        ribbon_height = 93      # Height of ribbon and toolbars
        status_height = 26      # Bottom status bar
        left_margin = 10        # Left margin (minimal in Win11 Paint)
        right_margin = int(win_width * 0.25)  # Right margin (approx 25% of window width based on screenshot)
        
        canvas_left = win_x + left_margin
        canvas_top = win_y + title_height + ribbon_height
        canvas_right = win_right - right_margin
        canvas_bottom = win_bottom - status_height
        
        canvas_width = canvas_right - canvas_left
        canvas_height = canvas_bottom - canvas_top
        
        # Visualize canvas boundaries with subtle highlights
        highlight_click_location(canvas_left, canvas_top, duration=0.2)
        highlight_click_location(canvas_right, canvas_top, duration=0.2)
        highlight_click_location(canvas_right, canvas_bottom, duration=0.2)
        highlight_click_location(canvas_left, canvas_bottom, duration=0.2)
        
        # Create a special tuple that can also be accessed like a dictionary
        # This provides backward compatibility with existing code
        canvas_tuple = (canvas_left, canvas_top, canvas_width, canvas_height)
        
        # Add dictionary attributes to the tuple for backward compatibility
        canvas_dict = {
            'left': canvas_left,
            'top': canvas_top,
            'right': canvas_right,
            'bottom': canvas_bottom,
            'width': canvas_width,
            'height': canvas_height
        }
        
        # Create a class with both tuple and dict access
        class CanvasInfo(tuple):
            def __getitem__(self, key):
                if isinstance(key, str):
                    return canvas_dict[key]
                return tuple.__getitem__(self, key)
                
            def __contains__(self, key):
                if isinstance(key, str):
                    return key in canvas_dict
                return tuple.__contains__(self, key)
                
            def get(self, key, default=None):
                if isinstance(key, str):
                    return canvas_dict.get(key, default)
                return tuple.__getitem__(self, key) if key < len(self) else default
                
        # Create combined object
        canvas_info = CanvasInfo(canvas_tuple)
        
        logging.info(f"Canvas boundaries: {canvas_dict}")
        logging.info(f"Returning as tuple: {canvas_tuple}")
        
        return canvas_info
        
    except Exception as e:
        logging.error(f"Error getting canvas info: {str(e)}")
        return None

def select_color(color_name='black'):
    """Select a color in Paint using keyboard shortcuts and mouse"""
    try:
        if not ensure_paint_focused():
            logging.error("Could not focus Paint window")
            return False
            
        # Get Paint window dimensions
        paint_window = find_paint_window()
        if not paint_window:
            logging.error("Could not find Paint window")
            return False
            
        window_rect = win32gui.GetWindowRect(paint_window)
        
        # Updated color palette coordinates for Windows 11 Paint
        # Colors are in the toolbar, with black at around 60% from left
        color_positions = {
            'black': (0.60, 75),   # Black is in the first row
            'white': (0.63, 75),   # White is next to black
            'red': (0.66, 75),     # Red follows
            'blue': (0.69, 75),    # Then blue
            'green': (0.72, 75)    # And green
        }
        
        if color_name not in color_positions:
            color_name = 'black'  # Default to black if color not found
            
        # Calculate absolute coordinates
        rel_x, fixed_y = color_positions[color_name]
        window_width = window_rect[2] - window_rect[0]
        
        color_x = window_rect[0] + int(window_width * rel_x)
        color_y = window_rect[1] + fixed_y
        
        # Try to select color
        for attempt in range(3):
            try:
                # Ensure we're in the right mode
                pyautogui.press('esc')  # Clear any active tool
                time.sleep(0.5)
                
                # Log the attempt
                logging.info(f"Attempting to select {color_name} color at ({color_x}, {color_y})")
                
                # Move to and click color with visual feedback
                pyautogui.moveTo(color_x, color_y, duration=0.5)
                time.sleep(0.5)
                pyautogui.click()
                time.sleep(0.5)
                
                # Verify color selection by checking if we can click
                pyautogui.moveTo(color_x + 5, color_y, duration=0.2)
                time.sleep(0.2)
                pyautogui.moveTo(color_x, color_y, duration=0.2)
                
                logging.info(f"Selected {color_name} color at ({color_x}, {color_y})")
                return True
                
            except Exception as e:
                logging.warning(f"Color selection attempt {attempt + 1} failed: {str(e)}")
                time.sleep(1)
                
        logging.error(f"Failed to select {color_name} color after multiple attempts")
        return False
        
    except Exception as e:
        logging.error(f"Error in select_color: {str(e)}")
        return False

def draw_shape(start_x, start_y, end_x, end_y, shape_type='rectangle'):
    """
    Draw a shape in MS Paint with better error handling and visual feedback.
    
    Args:
        start_x, start_y: Starting coordinates for the shape
        end_x, end_y: Ending coordinates for the shape
        shape_type: Type of shape ('rectangle', 'circle', 'triangle')
    
    Returns:
        True if successful, False otherwise
    """
    logging.info(f"Drawing {shape_type} from ({start_x}, {start_y}) to ({end_x}, {end_y})")
    
    try:
        # Step 1: Ensure Paint is focused
        if not ensure_paint_focused():
            logging.error("Failed to focus Paint window - cannot draw shape")
            return False
            
        # Step 2: Get canvas information
        canvas_info = get_paint_canvas_info()
        if not canvas_info:
            logging.error("Failed to get canvas information - cannot draw shape")
            return False
            
        # Step 3: Validate coordinates are within canvas
        if not validate_coordinates(start_x, start_y, end_x, end_y, canvas_info):
            logging.error("Invalid coordinates for drawing - must be within canvas boundaries")
            
            # Convert tuple to dictionary if needed
            if isinstance(canvas_info, tuple):
                canvas_left, canvas_top, canvas_width, canvas_height = canvas_info
                canvas_right = canvas_left + canvas_width
                canvas_bottom = canvas_top + canvas_height
            else:
                canvas_left = canvas_info['left']
                canvas_top = canvas_info['top']
                canvas_right = canvas_info['right']
                canvas_bottom = canvas_info['bottom']
                
            # Adjust coordinates to be within canvas if needed
            start_x = max(canvas_left, min(start_x, canvas_right))
            start_y = max(canvas_top, min(start_y, canvas_bottom))
            end_x = max(canvas_left, min(end_x, canvas_right))
            end_y = max(canvas_top, min(end_y, canvas_bottom))
            logging.info(f"Adjusted coordinates to: ({start_x}, {start_y}) to ({end_x}, {end_y})")
        
        # Step 4: Select the appropriate shape tool
        if not select_shape_tool(shape_type):
            logging.error(f"Failed to select {shape_type} tool - cannot draw shape")
            return False
        
        # Step 5: Move to a neutral position first, then to start position
        # Get canvas center
        if isinstance(canvas_info, tuple):
            canvas_left, canvas_top, canvas_width, canvas_height = canvas_info
            neutral_x = canvas_left + (canvas_width // 2)
            neutral_y = canvas_top + (canvas_height // 2)
        else:
            neutral_x = (canvas_info['left'] + canvas_info['right']) // 2
            neutral_y = (canvas_info['top'] + canvas_info['bottom']) // 2
        
        # Move to neutral position first
        pyautogui.moveTo(neutral_x, neutral_y, duration=0.5)
        time.sleep(0.5)
        
        # Log starting position
        logging.info(f"Moving to start position ({start_x}, {start_y})")
        
        # Step 6: Move to start position with visual feedback
        highlight_click_location(start_x, start_y, duration=0.3)
        pyautogui.moveTo(start_x, start_y, duration=0.7)
        time.sleep(0.5)
        
        # Step 7: Press and hold mouse button
        pyautogui.mouseDown()
        time.sleep(0.5)  # Hold briefly
        
        # Log ending position
        logging.info(f"Moving to end position ({end_x}, {end_y})")
        
        # Step 8: Move to end position with smooth motion
        pyautogui.moveTo(end_x, end_y, duration=1.0)
        time.sleep(0.5)  # Hold briefly
        
        # Step 9: Release mouse button
        pyautogui.mouseUp()
        time.sleep(0.5)
        
        # Step 10: Move away from the shape to avoid unintended interactions
        if isinstance(canvas_info, tuple):
            canvas_left, canvas_top, canvas_width, canvas_height = canvas_info
            away_x = canvas_left + (canvas_width // 2)
            away_y = canvas_top - 50  # Move above the canvas
        else:
            away_x = canvas_info['left'] + (canvas_info['width'] // 2)
            away_y = canvas_info['top'] - 50  # Move above the canvas
            
        pyautogui.moveTo(away_x, away_y, duration=0.5)
        
        logging.info(f"Successfully drew {shape_type}")
        return True
        
    except Exception as e:
        logging.error(f"Error drawing shape: {str(e)}", exc_info=True)
        # Emergency release of mouse button in case of error
        pyautogui.mouseUp()
        return False

def validate_coordinates(start_x, start_y, end_x, end_y, canvas_info):
    """
    Validate that coordinates are within canvas boundaries
    
    Args:
        start_x, start_y: Starting coordinates for the shape
        end_x, end_y: Ending coordinates for the shape
        canvas_info: Tuple (x, y, width, height) or dict with canvas info
        
    Returns:
        True if coordinates are valid, False otherwise
    """
    try:
        # Handle both dictionary and tuple formats
        if isinstance(canvas_info, tuple):
            # Tuple format: (x, y, width, height)
            canvas_left, canvas_top, canvas_width, canvas_height = canvas_info
            canvas_right = canvas_left + canvas_width
            canvas_bottom = canvas_top + canvas_height
        else:
            # Dictionary format with 'left', 'top', 'right', 'bottom' keys
            canvas_left = canvas_info['left']
            canvas_top = canvas_info['top']
            canvas_right = canvas_info['right'] 
            canvas_bottom = canvas_info['bottom']
        
        # Check if all coordinates are within canvas boundaries
        is_valid = (
            start_x >= canvas_left and start_x <= canvas_right and
            start_y >= canvas_top and start_y <= canvas_bottom and
            end_x >= canvas_left and end_x <= canvas_right and
            end_y >= canvas_top and end_y <= canvas_bottom
        )
        
        if not is_valid:
            logging.warning(
                f"Coordinates ({start_x}, {start_y}) to ({end_x}, {end_y}) " +
                f"are outside canvas boundaries: left={canvas_left}, " +
                f"top={canvas_top}, right={canvas_right}, " +
                f"bottom={canvas_bottom}"
            )
            
        return is_valid
    except Exception as e:
        logging.error(f"Error validating coordinates: {str(e)}")
        return False

def draw_rectangle(x1, y1, x2, y2):
    """
    Draw a rectangle in Windows 11 Paint from (x1,y1) to (x2,y2)
    Coordinates are relative to the canvas
    """
    try:
        logging.info(f"Drawing rectangle from ({x1}, {y1}) to ({x2}, {y2})")
        
        # Step 1: Ensure Paint is focused
        if not ensure_paint_focused():
            logging.error("Failed to focus Paint window")
            return False
            
        # Step 2: Get canvas information
        canvas_info = get_paint_canvas_info()
        if not canvas_info:
            logging.error("Failed to get canvas information")
            return False
            
        # Step 3: Calculate the screen coordinates based on the canvas info
        # If coordinates are small numbers, assume they're relative to canvas
        # Otherwise assume they're already screen coordinates
        if max(x1, y1, x2, y2) < 1000:  # Likely canvas-relative
            screen_x1 = canvas_info['left'] + x1
            screen_y1 = canvas_info['top'] + y1
            screen_x2 = canvas_info['left'] + x2
            screen_y2 = canvas_info['top'] + y2
            logging.info(f"Converting canvas coordinates: ({x1},{y1}) -> ({screen_x1},{screen_y1}), ({x2},{y2}) -> ({screen_x2},{screen_y2})")
        else:
            screen_x1 = x1
            screen_y1 = y1
            screen_x2 = x2
            screen_y2 = y2
            logging.info(f"Using screen coordinates: ({screen_x1},{screen_y1}) to ({screen_x2},{screen_y2})")
            
        # Step 4: Validate that coordinates are within canvas
        if not (canvas_info['left'] <= screen_x1 <= canvas_info['right'] and
                canvas_info['top'] <= screen_y1 <= canvas_info['bottom'] and
                canvas_info['left'] <= screen_x2 <= canvas_info['right'] and
                canvas_info['top'] <= screen_y2 <= canvas_info['bottom']):
            logging.warning("Drawing coordinates are outside canvas boundaries, adjusting...")
            screen_x1 = max(canvas_info['left'] + 10, min(screen_x1, canvas_info['right'] - 10))
            screen_y1 = max(canvas_info['top'] + 10, min(screen_y1, canvas_info['bottom'] - 10))
            screen_x2 = max(canvas_info['left'] + 10, min(screen_x2, canvas_info['right'] - 10))
            screen_y2 = max(canvas_info['top'] + 10, min(screen_y2, canvas_info['bottom'] - 10))
            logging.info(f"Adjusted to: ({screen_x1},{screen_y1}) to ({screen_x2},{screen_y2})")
            
        # Step 5: Take screenshot before for comparison
        screenshot_before = pyautogui.screenshot()
        
        # Step 6: Explicitly select the rectangle tool
        if not select_shape_tool('rectangle'):
            logging.error("Failed to select rectangle tool")
            return False
            
        # Wait to ensure tool is selected
        time.sleep(1)
            
        # Step 7: Move to center of canvas first for better visibility
        canvas_center_x = (canvas_info['left'] + canvas_info['right']) // 2
        canvas_center_y = (canvas_info['top'] + canvas_info['bottom']) // 2
        pyautogui.moveTo(canvas_center_x, canvas_center_y, duration=0.5)
        time.sleep(0.5)
        
        # Step 8: Move to start position with clear visual feedback
        logging.info(f"Moving to start position ({screen_x1}, {screen_y1})")
        pyautogui.moveTo(screen_x1, screen_y1, duration=0.8)
        
        # Small circle around start point for visual feedback
        offset = 5
        pyautogui.moveRel(offset, 0, duration=0.1)
        pyautogui.moveRel(0, offset, duration=0.1)
        pyautogui.moveRel(-offset, 0, duration=0.1)
        pyautogui.moveRel(0, -offset, duration=0.1)
        pyautogui.moveTo(screen_x1, screen_y1, duration=0.1)
        time.sleep(0.5)
        
        # Step 9: Press and hold mouse button
        logging.info("Pressing mouse button down")
        pyautogui.mouseDown()
        time.sleep(0.5)
        
        # Step 10: Move to end position in two stages for better control
        # First move horizontally to create visible feedback
        middle_x = screen_x2
        middle_y = screen_y1
        logging.info(f"Moving horizontally to ({middle_x}, {middle_y})")
        pyautogui.moveTo(middle_x, middle_y, duration=0.5)
        time.sleep(0.2)
        
        # Then move vertically to complete the rectangle
        logging.info(f"Moving vertically to final position ({screen_x2}, {screen_y2})")
        pyautogui.moveTo(screen_x2, screen_y2, duration=0.5)
        time.sleep(0.5)
        
        # Step 11: Release mouse button
        logging.info("Releasing mouse button")
        pyautogui.mouseUp()
        time.sleep(1)
        
        # Step 12: Move away from the rectangle for clarity
        pyautogui.moveTo(canvas_center_x, canvas_center_y - 50, duration=0.3)
        time.sleep(0.5)
        
        # Step 13: Take screenshot after to verify changes
        screenshot_after = pyautogui.screenshot()
        
        # Step 14: Verify drawing by comparing screenshots
        # Check if pixels changed in the rectangle area
        pixels_changed = 0
        sample_points = []
        
        # Check the perimeter of the rectangle for changes
        # Top and bottom edges
        for x in range(min(screen_x1, screen_x2), max(screen_x1, screen_x2), 10):
            # Top edge
            try:
                if screenshot_before.getpixel((x, min(screen_y1, screen_y2))) != screenshot_after.getpixel((x, min(screen_y1, screen_y2))):
                    pixels_changed += 1
                    if len(sample_points) < 5:
                        sample_points.append((x, min(screen_y1, screen_y2)))
            except Exception as e:
                logging.warning(f"Error checking pixel at ({x}, {min(screen_y1, screen_y2)}): {e}")
                
            # Bottom edge
            try:
                if screenshot_before.getpixel((x, max(screen_y1, screen_y2))) != screenshot_after.getpixel((x, max(screen_y1, screen_y2))):
                    pixels_changed += 1
                    if len(sample_points) < 5:
                        sample_points.append((x, max(screen_y1, screen_y2)))
            except Exception as e:
                logging.warning(f"Error checking pixel at ({x}, {max(screen_y1, screen_y2)}): {e}")
                
        # Left and right edges
        for y in range(min(screen_y1, screen_y2), max(screen_y1, screen_y2), 10):
            # Left edge
            try:
                if screenshot_before.getpixel((min(screen_x1, screen_x2), y)) != screenshot_after.getpixel((min(screen_x1, screen_x2), y)):
                    pixels_changed += 1
                    if len(sample_points) < 5:
                        sample_points.append((min(screen_x1, screen_x2), y))
            except Exception as e:
                logging.warning(f"Error checking pixel at ({min(screen_x1, screen_x2)}, {y}): {e}")
                
            # Right edge
            try:
                if screenshot_before.getpixel((max(screen_x1, screen_x2), y)) != screenshot_after.getpixel((max(screen_x1, screen_x2), y)):
                    pixels_changed += 1
                    if len(sample_points) < 5:
                        sample_points.append((max(screen_x1, screen_x2), y))
            except Exception as e:
                logging.warning(f"Error checking pixel at ({max(screen_x1, screen_x2)}, {y}): {e}")
                
        logging.info(f"Detected {pixels_changed} changed pixels around rectangle boundaries")
        
        # Save screenshots for debug purposes
        debug_dir = os.path.join(log_dir, "debug_screenshots")
        os.makedirs(debug_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        before_path = os.path.join(debug_dir, f"rectangle_before_{timestamp}.png")
        after_path = os.path.join(debug_dir, f"rectangle_after_{timestamp}.png")
        screenshot_before.save(before_path)
        screenshot_after.save(after_path)
        logging.info(f"Saved debug screenshots to {before_path} and {after_path}")
        
        if pixels_changed > 5:
            logging.info("Rectangle drawing verified successful - detected pixel changes")
            return True
        else:
            logging.warning("Few or no pixel changes detected, rectangle may not have been drawn correctly")
            return False
            
    except Exception as e:
        logging.error(f"Error drawing rectangle: {str(e)}", exc_info=True)
        # Emergency release of mouse button in case of error
        try:
            pyautogui.mouseUp()
        except:
            pass
        return False

def draw_circle(x, y, radius):
    """
    Draw a circle in MS Paint at the specified center and radius.
    Takes screenshots before and after to verify drawing was successful.
    """
    try:
        logging.info(f"Drawing circle at ({x}, {y}) with radius {radius}")
        
        # Calculate start and end points for drawing the circle
        start_x = x - radius
        start_y = y - radius
        end_x = x + radius
        end_y = y + radius
        
        # Ensure Paint is focused
        if not ensure_paint_focused():
            logging.error("Failed to focus Paint window")
            return False
            
        # Screenshot before drawing
        screenshot_before = pyautogui.screenshot()
        time.sleep(0.5)
        
        # Select circle/ellipse tool
        if not select_shape_tool(shape_type='circle'):
            logging.error("Failed to select circle tool")
            return False
            
        # Get canvas information
        canvas_info = get_paint_canvas_info()
        if not canvas_info:
            logging.error("Failed to get canvas information")
            return False
            
        # Validate that coordinates are within canvas
        if (start_x < canvas_info['left'] or start_y < canvas_info['top'] or
            end_x > canvas_info['left'] + canvas_info['width'] or
            end_y > canvas_info['top'] + canvas_info['height']):
            logging.error(f"Circle coordinates exceed canvas boundaries: ({start_x},{start_y}) to ({end_x},{end_y})")
            logging.error(f"Canvas boundaries: ({canvas_info['left']},{canvas_info['top']}) to "
                          f"({canvas_info['left'] + canvas_info['width']},{canvas_info['top'] + canvas_info['height']})")
            return False
            
        # Move to center of canvas first for better visual feedback
        canvas_center_x = canvas_info['left'] + (canvas_info['width'] // 2)
        canvas_center_y = canvas_info['top'] + (canvas_info['height'] // 2)
        pyautogui.moveTo(canvas_center_x, canvas_center_y, duration=0.5)
        time.sleep(0.5)
        
        # Move to start position
        logging.info(f"Moving to start position ({start_x}, {start_y})")
        pyautogui.moveTo(start_x, start_y, duration=0.5)
        time.sleep(0.5)
        
        # Click and drag to end position
        logging.info(f"Dragging to end position ({end_x}, {end_y})")
        pyautogui.mouseDown()
        time.sleep(0.3)
        pyautogui.moveTo(end_x, end_y, duration=1.0)
        time.sleep(0.3)
        pyautogui.mouseUp()
        time.sleep(1.0)
        
        # Move away from the drawn circle
        pyautogui.moveTo(canvas_center_x, canvas_center_y, duration=0.5)
        time.sleep(0.5)
        
        # Screenshot after drawing
        screenshot_after = pyautogui.screenshot()
        time.sleep(0.5)
        
        # Compare before and after to verify changes
        # Focus on the circle boundary pixels
        pixels_changed = False
        for check_x in range(start_x, end_x, 10):
            check_y_top = start_y
            check_y_bottom = end_y
            
            # Check if pixels at the boundary changed
            if (screenshot_before.getpixel((check_x, check_y_top)) != 
                screenshot_after.getpixel((check_x, check_y_top)) or
                screenshot_before.getpixel((check_x, check_y_bottom)) != 
                screenshot_after.getpixel((check_x, check_y_bottom))):
                pixels_changed = True
                break
                
        if pixels_changed:
            logging.info("Successfully drew circle - pixel changes detected")
            return True
        else:
            logging.error("Failed to draw circle - no pixel changes detected")
            
            # Save debug images
            debug_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug")
            os.makedirs(debug_dir, exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            before_path = os.path.join(debug_dir, f"circle_before_{timestamp}.png")
            after_path = os.path.join(debug_dir, f"circle_after_{timestamp}.png")
            
            screenshot_before.save(before_path)
            screenshot_after.save(after_path)
            logging.info(f"Saved debug screenshots to {before_path} and {after_path}")
            
            return False
            
    except Exception as e:
        logging.error(f"Error drawing circle: {str(e)}", exc_info=True)
        return False

def add_text_in_paint(text, x, y):
    """Adds text at specified coordinates in Paint."""
    try:
        if not ensure_paint_focused():
            logging.error("Could not focus Paint window")
            return False
            
        # Select black color first
        if not select_color('black'):
            logging.error("Failed to select black color")
            return False
            
        # Select text tool using keyboard navigation
        pyautogui.press('esc')  # Clear any previous selection
        time.sleep(0.2)
        pyautogui.hotkey('alt')
        time.sleep(0.2)
        pyautogui.press('h')  # Home tab
        time.sleep(0.2)
        pyautogui.press('a')  # Text tool
        time.sleep(0.5)
        
        # Get Paint canvas information
        canvas_info = get_paint_canvas_info()
        if not canvas_info:
            print("Could not get Paint canvas information")
            return False
            
        # Validate and adjust coordinates to fit within canvas
        x = max(0, min(x, canvas_info['width']))
        y = max(0, min(y, canvas_info['height']))
        
        # Convert to screen coordinates
        screen_x = canvas_info['left'] + x
        screen_y = canvas_info['top'] + y
        
        print(f"Adding text at ({x}, {y})")
        print(f"Screen coordinates: ({screen_x}, {screen_y})")
        
        # Click at position and type text
        pyautogui.click(screen_x, screen_y)
        time.sleep(0.5)
        pyautogui.write(text)
        time.sleep(0.5)
        
        # Click outside to finish text input
        pyautogui.click(screen_x + 100, screen_y + 100)
        time.sleep(0.2)
        
        return True
    except Exception as e:
        print(f"Error adding text: {str(e)}")
        return False

def save_image(path):
    """Saves the Paint image to specified path."""
    try:
        if not ensure_paint_focused():
            logging.error("Could not focus Paint window")
            return False
            
        # Convert path to Windows format and ensure it's absolute
        path = os.path.abspath(path).replace('/', '\\')
        logging.info(f"Saving image with absolute path: {path}")
        
        # Create directory if it doesn't exist
        directory = os.path.dirname(path)
        if not os.path.exists(directory):
            os.makedirs(directory)
            logging.info(f"Created directory: {directory}")
            
        # Press Ctrl+S for save dialog
        pyautogui.hotkey('ctrl', 's')
        time.sleep(2)  # Longer delay for save dialog
        
        # Clear any existing text and wait
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.5)
        pyautogui.press('delete')
        time.sleep(0.5)
        
        # Type the path directly
        pyautogui.write(path)
        time.sleep(1)
        
        # Ensure PNG format is selected (Alt+T to move to file type, then type PNG)
        pyautogui.hotkey('alt', 't')
        time.sleep(0.5)
        pyautogui.write('PNG')
        time.sleep(0.5)
        
        # Press Enter to save
        pyautogui.press('enter')
        time.sleep(3)  # Longer delay for save operation
        
        # Handle potential overwrite dialog
        try:
            overwrite_dialog = win32gui.FindWindow("#32770", "Paint")
            if overwrite_dialog:
                logging.info("Overwrite dialog detected")
                pyautogui.press('left')  # Move to Yes
                time.sleep(0.5)
                pyautogui.press('enter')
                time.sleep(2)
        except Exception:
            pass
            
        # Verify file exists with retry
        for attempt in range(5):
            if os.path.exists(path):
                file_size = os.path.getsize(path)
                if file_size > 0:
                    logging.info(f"File saved successfully at {path} (Size: {file_size} bytes)")
                    return True
            if attempt < 4:
                logging.warning(f"File not found or empty, waiting... (attempt {attempt + 1}/5)")
                time.sleep(2 ** attempt)  # Exponential backoff
                
        logging.error(f"Failed to verify file at {path}")
        return False
        
    except Exception as e:
        logging.error(f"Error saving image: {str(e)}", exc_info=True)
        return False

def draw_text(x, y, text):
    """Draw text in Paint at specified coordinates"""
    try:
        if not ensure_paint_focused():
            logging.error("Could not focus Paint window")
            return False
            
        # Get canvas info
        canvas_info = get_paint_canvas_info()
        if not canvas_info:
            logging.error("Could not get canvas info")
            return False
            
        # Validate coordinates are within canvas
        canvas_left, canvas_top, canvas_right, canvas_bottom = canvas_info
        if not (canvas_left <= x <= canvas_right and canvas_top <= y <= canvas_bottom):
            logging.error(f"Text coordinates ({x}, {y}) outside canvas bounds")
            return False
            
        # Select text tool - click at 20% from left, 75px from top
        window_width = pyautogui.size().width
        text_tool_x = int(window_width * 0.20)
        text_tool_y = 75
        
        logging.info(f"Selecting text tool at ({text_tool_x}, {text_tool_y})")
        pyautogui.moveTo(text_tool_x, text_tool_y, duration=0.5)
        time.sleep(0.5)
        pyautogui.click()
        time.sleep(0.5)
        
        # Move to text position with visual feedback
        logging.info(f"Moving to text position ({x}, {y})")
        pyautogui.moveTo(x, y, duration=0.5)
        time.sleep(0.5)
        
        # Click to set text cursor
        pyautogui.click()
        time.sleep(0.5)
        
        # Type text slowly for reliability
        logging.info(f"Typing text: {text}")
        pyautogui.write(text, interval=0.1)
        time.sleep(0.5)
        
        # Click away to finish text entry
        away_x = x + 50
        away_y = y + 50
        if away_x > canvas_right:
            away_x = x - 50
        if away_y > canvas_bottom:
            away_y = y - 50
            
        pyautogui.moveTo(away_x, away_y, duration=0.3)
        pyautogui.click()
        time.sleep(0.5)
        
        logging.info("Successfully drew text")
        return True
        
    except Exception as e:
        logging.error(f"Error in draw_text: {str(e)}")
        return False

def draw_direct_rectangle(x1, y1, x2, y2):
    """Simple direct approach to draw a rectangle in Paint"""
    try:
        logging.info(f"Drawing rectangle directly from ({x1}, {y1}) to ({x2}, {y2})")
        
        # Ensure Paint is focused and maximized
        if not ensure_paint_focused():
            logging.error("Failed to focus Paint window")
            return False
        
        # Get the Paint window handle for screenshots and verification
        paint_window = find_paint_window()
        if not paint_window:
            logging.error("Failed to find Paint window")
            return False
        
        # Take screenshot before actions to detect changes
        screenshot_before = pyautogui.screenshot()
        
        # Get window dimensions with direct Win32 call for more accuracy
        window_rect = win32gui.GetWindowRect(paint_window)
        win_x, win_y, win_right, win_bottom = window_rect
        window_width = win_right - win_x
        window_height = win_bottom - win_y
        logging.info(f"Paint window dimensions: {window_width}x{window_height} at ({win_x},{win_y})")
        
        # Press Escape to clear any current tool or selection
        pyautogui.press('esc')
        time.sleep(1)
        
        # Try a more reliable method to select the rectangle tool
        # First try using keyboard shortcut for shapes (Alt+H, SH)
        logging.info("Using keyboard shortcuts to select shape tools")
        pyautogui.hotkey('alt', 'h')
        time.sleep(1)
        pyautogui.press('s')
        time.sleep(1)
        pyautogui.press('h')
        time.sleep(1)
        
        # Now select rectangle (first option in shapes)
        logging.info("Selecting rectangle shape")
        pyautogui.press('r')  # First letter of rectangle
        time.sleep(1)
        
        # Alternative: Try clicking directly on specific toolbar positions
        # These coordinates are for the rectangle tool in the Home tab
        shape_x = win_x + int(window_width * 0.33)  # 33% from left
        shape_y = win_y + 75                       # 75px from top
        
        logging.info(f"Clicking directly on shapes tool at ({shape_x}, {shape_y})")
        pyautogui.moveTo(shape_x, shape_y, duration=1.0)
        time.sleep(1)
        pyautogui.click()
        time.sleep(2)
        
        # Then click on Rectangle (first shape in menu)
        rect_x = shape_x
        rect_y = shape_y + 40
        logging.info(f"Clicking rectangle shape at ({rect_x}, {rect_y})")
        pyautogui.moveTo(rect_x, rect_y, duration=1.0)
        time.sleep(1)
        pyautogui.click()
        time.sleep(2)
        
        # Get canvas info for drawing area
        canvas_info = get_paint_canvas_info()
        if not canvas_info:
            logging.error("Failed to get canvas info")
            return False
        
        # Log canvas info in greater detail
        logging.info(f"Canvas details: Left={canvas_info['left']}, Top={canvas_info['top']}, "
                     f"Right={canvas_info['right']}, Bottom={canvas_info['bottom']}, "
                     f"Width={canvas_info['width']}, Height={canvas_info['height']}")
        
        # Convert coordinates to screen coordinates
        if x1 < 1000 and y1 < 1000:  # Likely canvas-relative
            screen_x1 = canvas_info['left'] + x1
            screen_y1 = canvas_info['top'] + y1
            screen_x2 = canvas_info['left'] + x2
            screen_y2 = canvas_info['top'] + y2
            logging.info(f"Converting canvas coordinates: ({x1},{y1}) -> ({screen_x1},{screen_y1}), ({x2},{y2}) -> ({screen_x2},{screen_y2})")
        else:
            screen_x1 = x1
            screen_y1 = y1
            screen_x2 = x2
            screen_y2 = y2
            logging.info(f"Using as screen coordinates: ({screen_x1},{screen_y1}) to ({screen_x2},{screen_y2})")
            
        # Validate drawing coordinates are within canvas
        if not (canvas_info['left'] <= screen_x1 <= canvas_info['right'] and
                canvas_info['top'] <= screen_y1 <= canvas_info['bottom'] and
                canvas_info['left'] <= screen_x2 <= canvas_info['right'] and
                canvas_info['top'] <= screen_y2 <= canvas_info['bottom']):
            logging.warning("Drawing coordinates are outside canvas boundaries, adjusting...")
            screen_x1 = max(canvas_info['left'] + 10, min(screen_x1, canvas_info['right'] - 10))
            screen_y1 = max(canvas_info['top'] + 10, min(screen_y1, canvas_info['bottom'] - 10))
            screen_x2 = max(canvas_info['left'] + 10, min(screen_x2, canvas_info['right'] - 10))
            screen_y2 = max(canvas_info['top'] + 10, min(screen_y2, canvas_info['bottom'] - 10))
            logging.info(f"Adjusted to: ({screen_x1},{screen_y1}) to ({screen_x2},{screen_y2})")
        
        # Move to center of canvas first to ensure tool selection
        canvas_center_x = (canvas_info['left'] + canvas_info['right']) // 2
        canvas_center_y = (canvas_info['top'] + canvas_info['bottom']) // 2
        pyautogui.moveTo(canvas_center_x, canvas_center_y, duration=0.5)
        time.sleep(1)
        
        # Draw the rectangle with exaggerated movements for visibility
        logging.info(f"Moving to start position ({screen_x1}, {screen_y1}) - VISIBLE FEEDBACK")
        for i in range(5):  # Add a small circular motion for visibility
            offset = 3
            pyautogui.moveTo(screen_x1 + offset, screen_y1, duration=0.1)
            pyautogui.moveTo(screen_x1, screen_y1 + offset, duration=0.1)
            pyautogui.moveTo(screen_x1 - offset, screen_y1, duration=0.1)
            pyautogui.moveTo(screen_x1, screen_y1 - offset, duration=0.1)
        
        pyautogui.moveTo(screen_x1, screen_y1, duration=1.0)
        time.sleep(1)
        
        # Press and hold with visible feedback
        logging.info("PRESSING MOUSE DOWN")
        pyautogui.mouseDown(button='left')
        time.sleep(1)
        
        # Move to end position very slowly for better control
        logging.info(f"DRAGGING to end position ({screen_x2}, {screen_y2})")
        pyautogui.moveTo(screen_x2, screen_y1, duration=1.0)  # First go horizontally
        time.sleep(0.5)
        pyautogui.moveTo(screen_x2, screen_y2, duration=1.0)  # Then vertically
        time.sleep(1)
        
        # Release mouse button
        logging.info("RELEASING MOUSE BUTTON")
        pyautogui.mouseUp(button='left')
        time.sleep(1)
        
        # Take screenshot after to verify drawing
        screenshot_after = pyautogui.screenshot()
        
        # Compare screenshots to see if something changed
        pixels_changed = 0
        sample_points = []
        for x in range(min(screen_x1, screen_x2), max(screen_x1, screen_x2), 10):
            for y in range(min(screen_y1, screen_y2), max(screen_y1, screen_y2), 10):
                if screenshot_before.getpixel((x, y)) != screenshot_after.getpixel((x, y)):
                    pixels_changed += 1
                    if len(sample_points) < 5:
                        sample_points.append((x, y))
        
        logging.info(f"Detected {pixels_changed} changed pixels in drawing area")
        for point in sample_points:
            logging.info(f"Pixel changed at {point}: {screenshot_before.getpixel(point)} -> {screenshot_after.getpixel(point)}")
        
        if pixels_changed > 10:
            logging.info("Rectangle appears to be successfully drawn (detected changes)")
            return True
        else:
            logging.error("No pixel changes detected, rectangle might not have been drawn")
            return False
        
    except Exception as e:
        logging.error(f"Error drawing rectangle: {str(e)}", exc_info=True)
        return False

def close_paint():
    """Closes all instances of Microsoft Paint, handling save dialogs."""
    try:
        logging.info("Attempting to close Paint...")
        # Try to find Paint window
        paint_window = find_paint_window()
        if paint_window:
            logging.info("Found Paint window, closing it")
            # Try to close window gracefully
            win32gui.PostMessage(paint_window, win32con.WM_CLOSE, 0, 0)
            time.sleep(1.5)
            
            # Check for save dialog using window enumeration
            def find_save_dialog():
                dialog_found = [False]
                dialog_hwnd = [None]
                
                def enum_callback(hwnd, results):
                    if win32gui.IsWindowVisible(hwnd):
                        title = win32gui.GetWindowText(hwnd)
                        # Check for Paint or "Save" in the title
                        if "Paint" in title and ("save" in title.lower() or "?" in title):
                            logging.info(f"Found potential save dialog: '{title}'")
                            dialog_found[0] = True
                            dialog_hwnd[0] = hwnd
                            
                win32gui.EnumWindows(enum_callback, None)
                return dialog_found[0], dialog_hwnd[0]
            
            # Check for save dialog and handle it
            dialog_found, dialog_hwnd = find_save_dialog()
            if dialog_found:
                logging.info("Save dialog detected, clicking Don't Save")
                
                # First try to use keyboard shortcut for "Don't Save"
                pyautogui.press('n')  # 'n' for "Don't Save" on English Windows
                time.sleep(1)
                
                # If dialog still exists, try clicking the "Don't Save" button
                dialog_still_exists, _ = find_save_dialog()
                if dialog_still_exists:
                    # Find dialog position to click the "Don't Save" button
                    try:
                        # Get dialog position
                        dialog_rect = win32gui.GetWindowRect(dialog_hwnd)
                        d_left, d_top, d_right, d_bottom = dialog_rect
                        dialog_width = d_right - d_left
                        dialog_height = d_bottom - d_top
                        
                        # Calculate position for Don't Save button (middle button)
                        # Based on the screenshot, it's the middle button in the dialog
                        dont_save_x = d_left + (dialog_width // 2)
                        dont_save_y = d_bottom - 20  # 20px from bottom
                        
                        logging.info(f"Clicking 'Don't Save' button at ({dont_save_x}, {dont_save_y})")
                        pyautogui.moveTo(dont_save_x, dont_save_y, duration=0.5)
                        pyautogui.click()
                        time.sleep(1)
                    except Exception as e:
                        logging.error(f"Error handling save dialog: {str(e)}")
                        # Fall back to Alt+N for "Don't Save"
                        pyautogui.hotkey('alt', 'n')
                        time.sleep(1)
            
            # Check if Paint is still open
            time.sleep(1)
            if find_paint_window():
                logging.warning("Paint didn't close gracefully, trying to kill process")
                os.system("taskkill /f /im mspaint.exe")
                
        else:
            # Try to kill all Paint processes
            logging.info("No Paint window found, trying to kill any Paint processes")
            os.system("taskkill /f /im mspaint.exe")
            
        # Verify Paint is closed
        time.sleep(1.5)
        if not find_paint_window():
            logging.info("Paint closed successfully")
            return True
        else:
            logging.error("Failed to close Paint")
            return False
            
    except Exception as e:
        logging.error(f"Error closing Paint: {str(e)}")
        return False 