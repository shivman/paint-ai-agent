# Paint Drawing Agent

A Python-based automation tool that uses Google's Gemini AI to control Microsoft Paint through natural language commands. This project enables users to draw shapes, write text, and manipulate colors in MS Paint using simple English instructions.

## Features

- Natural language control of MS Paint
- Automated drawing of shapes (circles, rectangles, lines)
- Text insertion with position control
- Color selection and management
- Window management and canvas positioning
- Detailed logging and error handling
- Position calibration system

## Prerequisites

- Windows 10 or later
- Python 3.8+
- Microsoft Paint (mspaint.exe)
- Google Cloud API key for Gemini AI

## Installation

1. Clone the repository:
```bash
git clone [repository-url]
cd paint-drawing-agent
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

3. Create a `.env` file in the project root and add your Google API key:
```
GOOGLE_API_KEY=your_api_key_here
```

## Project Structure

```
paint-drawing-agent/
├── talk2mcp.py           # Main application file
├── tools/                # Core automation tools
│   ├── __init__.py
│   └── paint_commands.py # Paint automation commands
├── calibration_profiles/ # Stored calibration data
├── LLM_LOGS/            # AI interaction logs
├── logs/                # Application logs
└── requirements.txt     # Project dependencies
```

## Usage

1. Start the Paint Drawing Agent:
```bash
python talk2mcp.py
```

2. Enter natural language commands, for example:
- "Draw a red circle in the center"
- "Write 'Hello World' in black at the top"
- "Draw a blue rectangle on the right side"

3. To exit the program, type 'quit'

## Command Examples

```
> Draw a red circle in the center
> Write 'Hello World' in black at the top
> Draw a blue rectangle at coordinates 400,300
> Draw a green line from top to bottom
```

## Calibration

The system uses a calibration system to accurately locate Paint interface elements. To recalibrate:

1. Run the calibration script:
```bash
python enhanced_calibrate.py
```

2. Follow the on-screen instructions to calibrate tool positions

## Logging

The system maintains several types of logs:

- Application logs: `/logs/paint_agent_[timestamp].log`
- LLM interaction logs: `/LLM_LOGS/session_log.json`
- Calibration logs: `/calibration_profiles/`

## Error Handling

The system includes comprehensive error handling for:
- Window management issues
- Drawing operation failures
- AI communication errors
- Position calibration problems

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## Troubleshooting

Common issues and solutions:

1. **Paint Window Not Found**
   - Ensure MS Paint is installed
   - Try running Paint manually first

2. **Drawing Position Issues**
   - Run calibration again
   - Check screen resolution matches calibration

3. **AI Communication Errors**
   - Verify API key in .env file
   - Check internet connection

## License

[Your License Here]

## Acknowledgments

- Google Gemini AI for natural language processing
- PyAutoGUI for GUI automation
- Win32GUI for Windows interaction