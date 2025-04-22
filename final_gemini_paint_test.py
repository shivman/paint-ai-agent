import os
import json
import logging
from datetime import datetime
import google.generativeai as genai
from dotenv import load_dotenv
from tools.paint_commands import (
    open_paint,
    ensure_paint_focused,
    get_paint_canvas_info,
    draw_shape,
    select_shape_tool,
    draw_text,
    select_color
)

# Load environment variables
load_dotenv()
GOOGLE_API_KEY = os.getenv('GOOGLE_API_KEY')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(f'gemini_paint_test_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'),
        logging.StreamHandler()
    ]
)

class GeminiPaintTest:
    def __init__(self):
        # Configure Gemini
        genai.configure(api_key=GOOGLE_API_KEY)
        self.model = genai.GenerativeModel('gemini-pro')
        
        # Load calibration profile
        self.load_calibration()
        
        # Test scenarios
        self.test_scenarios = [
            {
                "name": "Basic Shape Drawing",
                "prompt": "Draw a red rectangle in the center of the canvas"
            },
            {
                "name": "Complex Shape Composition",
                "prompt": "Draw a house using a triangle for the roof and a square for the base"
            },
            {
                "name": "Text Integration",
                "prompt": "Write 'Hello World' in blue color at the top of the canvas"
            }
        ]
    
    def load_calibration(self):
        """Load calibration profile from file"""
        try:
            with open('calibration_profiles/default.json', 'r') as f:
                self.positions = json.load(f)
            logging.info("Loaded calibration profile successfully")
        except Exception as e:
            logging.error(f"Failed to load calibration profile: {e}")
            raise
    
    async def process_gemini_response(self, prompt):
        """Process the drawing instructions from Gemini"""
        try:
            response = await self.model.generate_content_async(
                f"You are a Paint drawing assistant. Given the following instruction, "
                f"provide specific drawing commands in JSON format with coordinates and colors: {prompt}"
            )
            
            # Parse the response and extract drawing commands
            try:
                commands = json.loads(response.text)
                logging.info(f"Received commands from Gemini: {commands}")
                return commands
            except json.JSONDecodeError:
                logging.error("Failed to parse Gemini response as JSON")
                return None
                
        except Exception as e:
            logging.error(f"Error getting response from Gemini: {e}")
            return None
    
    def execute_drawing_commands(self, commands):
        """Execute the drawing commands in Paint"""
        try:
            # Ensure Paint is open and focused
            if not ensure_paint_focused():
                if not open_paint():
                    raise Exception("Failed to open or focus Paint")
            
            # Get canvas information
            canvas_info = get_paint_canvas_info()
            if not canvas_info:
                raise Exception("Failed to get canvas information")
            
            # Execute each drawing command
            for cmd in commands:
                if cmd['type'] == 'shape':
                    # Select color
                    select_color(cmd['color'])
                    
                    # Select shape tool
                    select_shape_tool(cmd['shape'])
                    
                    # Draw shape
                    draw_shape(
                        cmd['start_x'],
                        cmd['start_y'],
                        cmd['end_x'],
                        cmd['end_y'],
                        cmd['shape']
                    )
                
                elif cmd['type'] == 'text':
                    # Write text
                    draw_text(
                        cmd['x'],
                        cmd['y'],
                        cmd['text'],
                        cmd['color']
                    )
                
                logging.info(f"Executed command: {cmd}")
                
        except Exception as e:
            logging.error(f"Error executing drawing commands: {e}")
            return False
        
        return True
    
    async def run_tests(self):
        """Run all test scenarios"""
        results = []
        
        for scenario in self.test_scenarios:
            logging.info(f"\nRunning test scenario: {scenario['name']}")
            
            try:
                # Get drawing commands from Gemini
                commands = await self.process_gemini_response(scenario['prompt'])
                if not commands:
                    raise Exception("Failed to get valid commands from Gemini")
                
                # Execute the commands
                success = self.execute_drawing_commands(commands)
                
                results.append({
                    "scenario": scenario['name'],
                    "prompt": scenario['prompt'],
                    "success": success
                })
                
                logging.info(f"Test scenario completed: {scenario['name']} - {'Success' if success else 'Failed'}")
                
            except Exception as e:
                logging.error(f"Error in test scenario {scenario['name']}: {e}")
                results.append({
                    "scenario": scenario['name'],
                    "prompt": scenario['prompt'],
                    "success": False,
                    "error": str(e)
                })
        
        return results

def main():
    try:
        # Initialize test suite
        test_suite = GeminiPaintTest()
        
        # Run tests
        import asyncio
        results = asyncio.run(test_suite.run_tests())
        
        # Log results
        logging.info("\nTest Results Summary:")
        for result in results:
            status = "✓ Passed" if result['success'] else "✗ Failed"
            logging.info(f"{status} - {result['scenario']}")
            if not result['success'] and 'error' in result:
                logging.info(f"  Error: {result['error']}")
        
    except Exception as e:
        logging.error(f"Test suite failed: {e}")

if __name__ == "__main__":
    main() 