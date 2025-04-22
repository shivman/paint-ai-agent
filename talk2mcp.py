import os
import json
import time
from datetime import datetime
from dotenv import load_dotenv
import google.generativeai as genai
import win32gui
from tools.paint_commands import (
    open_paint,
    ensure_paint_focused,
    draw_shape,
    draw_rectangle,
    draw_circle,
    select_color,
    add_text_in_paint,
    save_image,
    close_paint,
    draw_text
)
import pyautogui
import logging

# Configure logging
logging.basicConfig(level=logging.INFO,
                   format='%(asctime)s - %(levelname)s - %(message)s')

# Load environment variables
load_dotenv()
GOOGLE_API_KEY = os.getenv('GOOGLE_API_KEY')
if not GOOGLE_API_KEY:
    raise ValueError("GOOGLE_API_KEY not found in environment variables!")

# Configure the Gemini API
logging.info("Configuring Gemini API...")
genai.configure(api_key=GOOGLE_API_KEY)

# Initialize the model
logging.info("Initializing Gemini model...")
model = genai.GenerativeModel('gemini-2.0-flash')

# Tool descriptions for function calling
TOOLS = {
    "open_paint": {
        "description": "Opens Microsoft Paint application. Must be called before any drawing commands.",
        "parameters": {}
    },
    "ensure_paint_focused": {
        "description": "Makes sure the Paint window is in focus and maximized. Use this if drawing commands are not working properly.",
        "parameters": {}
    },
    "draw_shape": {
        "description": "Draws a shape in Paint. Paint must be opened first.",
        "parameters": {
            "shape_type": "str: Type of shape ('rectangle', 'oval', 'triangle', 'line', 'pentagon', 'hexagon', 'star')",
            "start_x": "int: Starting x coordinate (100-1000)",
            "start_y": "int: Starting y coordinate (100-800)",
            "end_x": "int: Ending x coordinate (100-1000)", 
            "end_y": "int: Ending y coordinate (100-800)"
        }
    },
    "select_color": {
        "description": "Selects a color for drawing in Paint. Paint must be opened first.",
        "parameters": {
            "color_name": "str: Color name (e.g., 'red', 'blue', 'green', 'black', 'white', 'yellow', 'purple', 'orange', 'dark red', 'lime', 'indigo', 'brown')"
        }
    },
    "add_text_in_paint": {
        "description": "Adds text at specified coordinates in Paint. Paint must be opened first.",
        "parameters": {
            "text": "str: Text to add",
            "x": "int: X coordinate (100-1000)",
            "y": "int: Y coordinate (100-800)"
        }
    },
    "save_image": {
        "description": "Saves the Paint image to specified path",
        "parameters": {
            "path": "str: Full path where to save the image"
        }
    },
    "close_paint": {
        "description": "Closes Microsoft Paint application.",
        "parameters": {}
    }
}

# Global variables for canvas dimensions
CANVAS_X = 0
CANVAS_Y = 0
CANVAS_WIDTH = 0
CANVAS_HEIGHT = 0
WINDOW_FOUND = False

def update_canvas_dimensions():
    """Update the canvas dimensions based on the Paint window size"""
    global CANVAS_X, CANVAS_Y, CANVAS_WIDTH, CANVAS_HEIGHT, WINDOW_FOUND
    WINDOW_FOUND = False
    
    def callback(hwnd, extra):
        global CANVAS_X, CANVAS_Y, CANVAS_WIDTH, CANVAS_HEIGHT, WINDOW_FOUND
        if win32gui.GetWindowText(hwnd) == 'Untitled - Paint' or 'Paint' in win32gui.GetWindowText(hwnd):
            rect = win32gui.GetWindowRect(hwnd)
            logging.info(f"Found Paint window at {rect}")
            
            # Calculate canvas area (excluding toolbars and margins)
            # Toolbar height is approximately 150 pixels from top
            # Left margin is approximately 5 pixels
            # Right margin is approximately 25 pixels
            # Bottom margin is approximately 50 pixels
            CANVAS_X = max(rect[0] + 5, 0)  # Ensure positive value
            CANVAS_Y = max(rect[1] + 150, 0)  # Ensure positive value
            CANVAS_WIDTH = max(rect[2] - rect[0] - 30, 100)  # Minimum width 100
            CANVAS_HEIGHT = max(rect[3] - rect[1] - 200, 100)  # Minimum height 100
            
            logging.info(f"Canvas dimensions: x={CANVAS_X}, y={CANVAS_Y}, width={CANVAS_WIDTH}, height={CANVAS_HEIGHT}")
            WINDOW_FOUND = True
    
    try:
        win32gui.EnumWindows(callback, None)
        return WINDOW_FOUND
    except Exception as e:
        logging.error(f"Error in update_canvas_dimensions: {e}")
        return False

# Load calibrated positions from paint_positions.txt
def load_paint_positions():
    """Load positions from paint_positions.txt"""
    positions = {}
    try:
        with open('../pjt2 paint test/paint_positions.txt', 'r') as f:
            for line in f:
                if ':' in line:
                    name, pos = line.strip().split(':')
                    name = name.strip()
                    # Convert string tuple to actual coordinates
                    pos = eval(pos.strip())
                    positions[name] = pos
        logging.info(f"Loaded {len(positions)} positions from paint_positions.txt")
        return positions
    except Exception as e:
        logging.error(f"Error loading paint_positions.txt: {e}")
        return None

# Load calibrated positions
CALIBRATED_POSITIONS = load_paint_positions()

def log_llm_response(prompt, response):
    """Logs the LLM interaction to a JSON file."""
    os.makedirs('LLM_LOGS', exist_ok=True)
    log_entry = {
        'timestamp': datetime.now().isoformat(),
        'prompt': prompt,
        'response': response
    }
    
    log_file = 'LLM_LOGS/session_log.json'
    try:
        with open(log_file, 'r') as f:
            logs = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        logs = []
    
    logs.append(log_entry)
    
    with open(log_file, 'w') as f:
        json.dump(logs, f, indent=2)

def execute_tool(tool_name, params):
    """Executes the specified tool with given parameters."""
    print(f"Executing tool: {tool_name} with params: {params}")
    try:
        if isinstance(tool_name, dict):
            print("Error: Tool name was passed as a dictionary. Skipping execution.")
            return False
            
        tool_name = tool_name.strip()  # Remove any whitespace
        
        if tool_name == 'open_paint':
            print("Opening Paint...")
            success = open_paint()
            if success:
                # Update canvas dimensions after opening Paint
                time.sleep(2)  # Wait longer for window to settle
                if update_canvas_dimensions():
                    print("Paint opened and canvas dimensions updated successfully")
                    return True
                else:
                    print("Paint opened but failed to get canvas dimensions")
                    return False
            return success
            
        elif tool_name == 'ensure_paint_focused':
            print("Focusing Paint window...")
            success = ensure_paint_focused()
            if success:
                # Update canvas dimensions after focusing
                time.sleep(1)
                if update_canvas_dimensions():
                    print("Paint focused and canvas dimensions updated successfully")
                    return True
                else:
                    print("Paint focused but failed to get canvas dimensions")
                    return False
            return success
            
        elif tool_name == 'draw_shape':
            if not WINDOW_FOUND:
                if not update_canvas_dimensions():
                    print("Cannot draw shape: Paint window not found")
                    return False
            
            shape_type = params.get('shape_type', 'rectangle').lower()
            
            # Get relative coordinates (0-1000 range)
            rel_start_x = params.get('start_x', params.get('x1', 300))
            rel_start_y = params.get('start_y', params.get('y1', 200))
            rel_end_x = params.get('end_x', params.get('x2', 500))
            rel_end_y = params.get('end_y', params.get('y2', 400))
            
            # Special handling for circles to ensure equal dimensions
            if shape_type in ['circle']:
                # Calculate center point and radius
                center_x = (rel_start_x + rel_end_x) // 2
                center_y = (rel_start_y + rel_end_y) // 2
                # Use the smaller of width or height to ensure it's a circle
                radius = min(abs(rel_end_x - rel_start_x), abs(rel_end_y - rel_start_y)) // 2
                # Recalculate start and end points to make a perfect circle
                rel_start_x = center_x - radius
                rel_start_y = center_y - radius
                rel_end_x = center_x + radius
                rel_end_y = center_y + radius
            
            # Convert relative coordinates to actual screen coordinates
            start_x = int(CANVAS_X + (rel_start_x * CANVAS_WIDTH / 1000))
            start_y = int(CANVAS_Y + (rel_start_y * CANVAS_HEIGHT / 1000))
            end_x = int(CANVAS_X + (rel_end_x * CANVAS_WIDTH / 1000))
            end_y = int(CANVAS_Y + (rel_end_y * CANVAS_HEIGHT / 1000))
            
            # Ensure coordinates are within safe bounds
            start_x = max(CANVAS_X + 10, min(start_x, CANVAS_X + CANVAS_WIDTH - 10))
            start_y = max(CANVAS_Y + 10, min(start_y, CANVAS_Y + CANVAS_HEIGHT - 10))
            end_x = max(CANVAS_X + 10, min(end_x, CANVAS_X + CANVAS_WIDTH - 10))
            end_y = max(CANVAS_Y + 10, min(end_y, CANVAS_Y + CANVAS_HEIGHT - 10))
            
            # For circles, ensure the aspect ratio remains 1:1 after bounds checking
            if shape_type in ['circle']:
                width = end_x - start_x
                height = end_y - start_y
                size = min(width, height)
                center_x = (start_x + end_x) // 2
                center_y = (start_y + end_y) // 2
                start_x = center_x - size // 2
                start_y = center_y - size // 2
                end_x = center_x + size // 2
                end_y = center_y + size // 2
            
            print(f"Drawing {shape_type} from ({start_x}, {start_y}) to ({end_x}, {end_y})")
            
            # Map shape types to their calibrated names
            shape_map = {
                'rectangle': 'rectangle',
                'circle': 'oval',  # We use the oval tool but maintain 1:1 ratio for circles
                'oval': 'oval',
                'triangle': 'triangle',
                'line': 'line',
                'pentagon': 'pentagon',
                'hexagon': 'hexagon',
                'star': 'five point star'
            }
            
            calibrated_shape = shape_map.get(shape_type, shape_type)
            
            if calibrated_shape not in CALIBRATED_POSITIONS:
                print(f"Shape type '{shape_type}' not found in calibration data")
                return False
            
            # Click the shape tool
            shape_x, shape_y = CALIBRATED_POSITIONS[calibrated_shape]
            pyautogui.moveTo(shape_x, shape_y, duration=0.2)
            pyautogui.click()
            time.sleep(0.5)
            
            # For circles, hold Shift key while drawing to maintain aspect ratio
            if shape_type in ['circle']:
                pyautogui.keyDown('shift')
            
            # Move to center of canvas first
            canvas_center_x = CANVAS_X + (CANVAS_WIDTH // 2)
            canvas_center_y = CANVAS_Y + (CANVAS_HEIGHT // 2)
            pyautogui.moveTo(canvas_center_x, canvas_center_y, duration=0.2)
            pyautogui.click()
            time.sleep(0.5)
            
            # Draw the shape with smoother movements
            pyautogui.moveTo(start_x, start_y, duration=0.2)
            time.sleep(0.2)
            pyautogui.mouseDown()
            time.sleep(0.2)
            pyautogui.moveTo(end_x, end_y, duration=0.5)
            time.sleep(0.2)
            pyautogui.mouseUp()
            
            # Release Shift key if we were drawing a circle
            if shape_type in ['circle']:
                pyautogui.keyUp('shift')
            
            time.sleep(0.5)
            
            # Deselect the shape by:
            # 1. Click the select tool
            if 'select' in CALIBRATED_POSITIONS:
                select_x, select_y = CALIBRATED_POSITIONS['select']
                pyautogui.moveTo(select_x, select_y, duration=0.2)
                pyautogui.click()
                time.sleep(0.3)
                
                # 2. Click in an empty area of the canvas (top-left corner with offset)
                empty_x = CANVAS_X + 20
                empty_y = CANVAS_Y + 20
                pyautogui.moveTo(empty_x, empty_y, duration=0.2)
                pyautogui.click()
                
                # 3. Move mouse to a safe position
                pyautogui.moveTo(CANVAS_X + 50, CANVAS_Y + 50, duration=0.2)
            
            print(f"Drew {shape_type} shape")
            return True
            
        elif tool_name == 'select_color':
            color_name = params['color_name'].lower()
            print(f"Selecting color: {color_name}")
            
            # Map common color names to calibrated names from paint_positions.txt
            color_map = {
                'blue': 'indigo',
                'light blue': 'torquoise',
                'dark blue': 'blue grey',
                'orange': 'orange',
                'yellow': 'yellow',
                'light yellow': 'light yellow',
                'green': 'green',
                'light green': 'lime',
                'purple': 'purple',
                'violet': 'lavender',
                'brown': 'brown',
                'pink': 'rose',
                'red': 'red',
                'dark red': 'dark red',
                'black': 'black color',
                'white': 'white',
                'gray': 'gray-50%',
                'grey': 'gray-50%',
                'light gray': 'grey-25%',
                'light grey': 'grey-25%'
            }
            
            calibrated_color = color_map.get(color_name, color_name)
            
            if calibrated_color not in CALIBRATED_POSITIONS:
                print(f"Color '{color_name}' not found in calibration data")
                return False
                
            color_x, color_y = CALIBRATED_POSITIONS[calibrated_color]
            pyautogui.moveTo(color_x, color_y, duration=0.2)
            pyautogui.click()
            time.sleep(0.5)
            
            print(f"Selected color: {color_name}")
            return True
            
        elif tool_name == 'add_text_in_paint':
            # Convert coordinates to actual screen positions
            rel_x = params.get('x', 400)
            rel_y = params.get('y', 100)
            
            # Calculate actual screen coordinates
            screen_x = int(CANVAS_X + (rel_x * CANVAS_WIDTH / 1000))
            screen_y = int(CANVAS_Y + (rel_y * CANVAS_HEIGHT / 1000))
            
            # Ensure coordinates are within safe bounds
            screen_x = max(CANVAS_X + 10, min(screen_x, CANVAS_X + CANVAS_WIDTH - 10))
            screen_y = max(CANVAS_Y + 10, min(screen_y, CANVAS_Y + CANVAS_HEIGHT - 10))
            
            # Check for text tool using the correct calibrated name
            if 'text tool' not in CALIBRATED_POSITIONS:
                print("Text tool position not found in calibration data")
                return False
                
            # Select text tool
            text_x, text_y = CALIBRATED_POSITIONS['text tool']
            pyautogui.moveTo(text_x, text_y, duration=0.2)
            pyautogui.click()
            time.sleep(0.5)
            
            # Click where to add text
            pyautogui.moveTo(screen_x, screen_y, duration=0.2)
            pyautogui.click()
            time.sleep(0.5)
            
            # Type the text
            pyautogui.write(params['text'])
            time.sleep(0.5)
            
            # Click away to finish text entry (use select tool to deselect)
            if 'select' in CALIBRATED_POSITIONS:
                select_x, select_y = CALIBRATED_POSITIONS['select']
                pyautogui.moveTo(select_x, select_y, duration=0.2)
                pyautogui.click()
                time.sleep(0.3)
                
                # Click in an empty area
                empty_x = CANVAS_X + 20
                empty_y = CANVAS_Y + 20
                pyautogui.moveTo(empty_x, empty_y, duration=0.2)
                pyautogui.click()
                
                # Move to safe position
                pyautogui.moveTo(CANVAS_X + 50, CANVAS_Y + 50, duration=0.2)
            
            print(f"Added text: {params['text']} at ({screen_x}, {screen_y})")
            return True
            
        elif tool_name == 'save_image':
            print(f"Saving image to {params['path']}")
            success = save_image(params['path'])
            return success
            
        elif tool_name == 'close_paint':
            print("Closing Paint...")
            success = close_paint()
            print("Paint closed successfully" if success else "Failed to close Paint")
            return success
            
        else:
            print(f"Unknown tool: {tool_name}")
            return False
            
    except KeyError as e:
        print(f"Missing parameter for {tool_name}: {e}")
        return False
    except Exception as e:
        print(f"Error executing {tool_name}: {str(e)}")
        return False

def process_llm_response(response_text):
    """Processes the LLM response to extract and execute tool calls."""
    print(f"\nProcessing LLM response...")
    lines = response_text.split('\n')
    
    total_commands = sum(1 for line in lines if line.strip().startswith('TOOL:'))
    executed_commands = 0
    successful_commands = 0
    
    print(f"Found {total_commands} commands to execute")
    
    for line in lines:
        line = line.strip()
        if line.startswith('TOOL:'):
            # Remove 'TOOL:' prefix and split by pipe
            executed_commands += 1
            command_text = line[5:].strip()  # Remove 'TOOL:' and any leading/trailing whitespace
            parts = command_text.split('|')
            
            if len(parts) >= 1:
                # Extract tool name
                tool_name = parts[0].strip()
                params = {}
                
                # Parse parameters if they exist
                if len(parts) > 1:
                    try:
                        params = json.loads(parts[1].strip())
                    except json.JSONDecodeError as e:
                        print(f"Error parsing parameters: {str(e)}")
                        continue
                
                # Execute the tool with parsed name and parameters
                print(f"\nExecuting command {executed_commands}/{total_commands}: {tool_name}")
                success = execute_tool(tool_name, params)
                if success:
                    successful_commands += 1
    
    print(f"\nExecution summary: {successful_commands}/{total_commands} commands executed successfully")
    return successful_commands == total_commands

def select_tool(tool_name):
    """Select a tool in Paint using calibrated positions"""
    if not CALIBRATED_POSITIONS:
        print("No calibration data available")
        return False
    
    if tool_name.lower() not in CALIBRATED_POSITIONS:
        print(f"Tool {tool_name} not found in calibration data")
        return False
    
    try:
        # Click on the tool
        x, y = CALIBRATED_POSITIONS[tool_name.lower()]
        pyautogui.click(x, y)
        print(f"Selected tool: {tool_name}")
        time.sleep(0.5)  # Short delay after selecting tool
        return True
    except Exception as e:
        print(f"Error selecting tool: {e}")
        return False

def select_color(color_name):
    """Select a color in Paint using calibrated positions"""
    if not CALIBRATED_POSITIONS:
        print("No calibration data available")
        return False
    
    if color_name.lower() not in CALIBRATED_POSITIONS:
        print(f"Color {color_name} not found in calibration data")
        return False
    
    try:
        # Click on the color
        x, y = CALIBRATED_POSITIONS[color_name.lower()]
        pyautogui.click(x, y)
        print(f"Selected color: {color_name}")
        time.sleep(0.5)  # Short delay after selecting color
        return True
    except Exception as e:
        print(f"Error selecting color: {e}")
        return False

def main():
    print("\n===== Paint Drawing Agent =====")
    print("Enter your drawing instructions (or 'quit' to exit):")
    print("\nNOTE: For any drawing command, Paint will be opened automatically if it's not already open.")
    print("TIP: Be specific about colors, shapes, and coordinates in your requests.")
    
    # Verify calibration data is loaded
    if not CALIBRATED_POSITIONS:
        print("ERROR: Calibration data could not be loaded!")
        return
    
    while True:
        try:
            user_prompt = input("\nWhat would you like to draw? > ")
            if user_prompt.lower() in ['quit', 'exit', 'q']:
                break
                
            logging.info(f"Processing user prompt: {user_prompt}")
            
            # Prepare prompt for Gemini
            system_prompt = f"""You are a Paint Drawing Agent that controls Microsoft Paint through various tools.
Available tools: {json.dumps(TOOLS, indent=2)}

IMPORTANT RULES:
1. ALWAYS start with 'open_paint' command before any drawing commands
2. Each command must be on a new line
3. Use exact format: TOOL: tool_name | {{"param1": value1, "param2": value2}}
4. For shapes and text, use coordinates that make sense for a typical screen (e.g., 100-1000 range)
5. For colors, use standard color names: red, blue, green, black, white, yellow, purple, orange, etc.
6. When the user requests a complex drawing, break it down into multiple shape commands
7. Always select a color before drawing shapes or text

Example response for drawing a house:
TOOL: open_paint
TOOL: ensure_paint_focused
TOOL: select_color | {{"color_name": "blue"}}
TOOL: draw_rectangle | {{"x1": 200, "y1": 300, "x2": 400, "y2": 500}}
TOOL: select_color | {{"color_name": "red"}}
TOOL: draw_shape | {{"start_x": 150, "start_y": 300, "end_x": 450, "end_y": 200, "shape_type": "triangle"}}

Parse the user's request and respond with appropriate tool calls to achieve the desired drawing.
Remember to ALWAYS start with opening Paint and select a color before each new shape!
"""
            
            try:
                print("\nSending request to AI...")
                logging.info("Sending request to Gemini API...")
                # Get response from Gemini
                response = model.generate_content(
                    system_prompt + "\n\nUser request: " + user_prompt
                )
                
                # Log the interaction
                log_llm_response(user_prompt, response.text)
                
                print("\nAI response received. Here's what I'll do:")
                # Extract and print the plain text part for the user to see
                for line in response.text.split('\n'):
                    if line.strip().startswith('TOOL:'):
                        print(f"- {line.strip()[5:]}")
                
                # Process and execute the tools
                execution_success = process_llm_response(response.text)
                
                # Provide feedback based on execution
                if execution_success:
                    print("\nAll drawing commands executed successfully!")
                else:
                    print("\nSome drawing commands failed. You might need to try again with different instructions.")
                
            except Exception as e:
                logging.error(f"Error during API call or execution: {str(e)}")
                print(f"Error: {str(e)}")
                print("Please try again or type 'quit' to exit.")
                
        except KeyboardInterrupt:
            print("\nOperation cancelled by user. Type 'quit' to exit properly.")
            continue
        except Exception as e:
            logging.error(f"Unexpected error: {str(e)}")
            print(f"An unexpected error occurred: {str(e)}")
            print("Please try again or type 'quit' to exit.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logging.error(f"Fatal error: {str(e)}")
        print(f"Fatal error occurred: {str(e)}")
    finally:
        print("\nThank you for using Paint Drawing Agent!") 