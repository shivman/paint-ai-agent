import os
import json
import time
import logging
import keyboard
import pyautogui
import subprocess
from datetime import datetime

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(f'calibration_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'),
        logging.StreamHandler()
    ]
)

class PaintCalibrator:
    def __init__(self):
        self.positions = {}
        self.profile_path = "calibration_profiles"
        self.current_profile = "default"
        
        # Create profiles directory if it doesn't exist
        if not os.path.exists(self.profile_path):
            os.makedirs(self.profile_path)
    
    def open_paint(self):
        """Open Microsoft Paint application"""
        try:
            # Close any existing Paint windows
            subprocess.run(['taskkill', '/F', '/IM', 'mspaint.exe'], 
                         stdout=subprocess.DEVNULL, 
                         stderr=subprocess.DEVNULL)
            time.sleep(1)
            
            # Start Paint using the Windows Run command
            subprocess.run(['cmd', '/c', 'start', 'mspaint'], shell=True)
            logging.info("Opening Paint... Please wait")
            time.sleep(3)  # Wait for Paint to open
            
            # Move mouse to center of screen to ensure Paint window is active
            screen_width, screen_height = pyautogui.size()
            pyautogui.moveTo(screen_width // 2, screen_height // 2)
            pyautogui.click()
            time.sleep(0.5)
            
            return True
        except Exception as e:
            logging.error(f"Error opening Paint: {e}")
            return False

    def get_mouse_position(self):
        """Interactive calibration process"""
        required_items = [
            "rectangle", "circle", "triangle", "line",
            "text", "black", "blue", "red", "green"
        ]
        
        logging.info("\nCalibration Instructions:")
        logging.info("1. Move your mouse to each tool/shape you want to calibrate")
        logging.info("2. Type the name of the tool/color")
        logging.info("3. Press SPACE to record the position")
        logging.info("4. Press 'q' when you're done")
        
        logging.info("\nRequired items to calibrate:")
        for item in required_items:
            logging.info(f"- {item}")
        
        while True:
            if keyboard.is_pressed('q'):
                break
            
            if keyboard.is_pressed('space'):
                x, y = pyautogui.position()
                name = input("\nEnter the name of this tool/color: ").lower().strip()
                
                if self.validate_position(x, y):
                    self.positions[name] = (x, y)
                    logging.info(f"Recorded {name} at position: ({x}, {y})")
                    
                    # Visual feedback
                    original_pos = pyautogui.position()
                    pyautogui.moveTo(x, y)
                    time.sleep(0.2)
                    pyautogui.moveTo(original_pos)
                else:
                    logging.warning(f"Invalid position recorded for {name}: ({x}, {y})")
                
                time.sleep(0.5)  # Prevent multiple readings
        
        # Check if all required items are calibrated
        missing_items = [item for item in required_items if item not in self.positions]
        if missing_items:
            logging.warning(f"Missing calibration for: {', '.join(missing_items)}")
    
    def validate_position(self, x, y):
        """Validate if the recorded position is within the screen bounds"""
        screen_width, screen_height = pyautogui.size()
        return 0 <= x <= screen_width and 0 <= y <= screen_height
    
    def save_profile(self, profile_name=None):
        """Save calibration profile"""
        if not profile_name:
            profile_name = self.current_profile
        
        filepath = os.path.join(self.profile_path, f"{profile_name}.json")
        try:
            with open(filepath, 'w') as f:
                json.dump(self.positions, f, indent=4)
            logging.info(f"Calibration profile saved to: {filepath}")
            return True
        except Exception as e:
            logging.error(f"Error saving calibration profile: {e}")
            return False

def main():
    try:
        # Initialize calibrator
        calibrator = PaintCalibrator()
        
        logging.info("Starting Paint and calibration process...")
        logging.info("1. Opening Paint")
        
        if not calibrator.open_paint():
            logging.error("Failed to open Paint. Please try again.")
            return
        
        logging.info("2. Please position the Paint window as desired")
        logging.info("3. Calibration will begin in 5 seconds...")
        time.sleep(5)
        
        calibrator.get_mouse_position()
        
        # Save the calibration profile
        profile_name = input("\nEnter a name for this calibration profile (or press Enter for 'default'): ").strip()
        if not profile_name:
            profile_name = "default"
        
        if calibrator.save_profile(profile_name):
            logging.info("Calibration completed successfully!")
        else:
            logging.error("Failed to save calibration profile.")
        
    except Exception as e:
        logging.error(f"Calibration process failed: {e}")

if __name__ == "__main__":
    main() 