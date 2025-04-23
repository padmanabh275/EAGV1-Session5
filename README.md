# ASCII Value Calculator with PowerPoint and Email Integration

## Overview
This project is an interactive system that calculates the sum of exponentials of ASCII values for a given string and visualizes the result in Microsoft PowerPoint. It then sends an email with the results and the PowerPoint presentation as an attachment. The system uses a carefully structured prompt design that follows best practices for LLM reasoning and tool usage.

## Features
- Convert text to ASCII values
- Calculate exponential sums of ASCII values
- Create automated PowerPoint presentations
- Send emails with results and attachments using Gmail API
- Structured prompt design for reliable LLM reasoning
- Multi-turn conversation support
- Robust error handling and validation

## Project Structure
```
Session4/
├── talk2mcp-2.py         # Main client script with structured prompt
├── example2-3.py          # Server with PowerPoint and email functionality
├── gmail_server.py        # Dedicated Gmail API server
└── README.md              # This file
```

## Prerequisites
- Python 3.x
- Microsoft PowerPoint
- Gmail account with API access
- Required Python packages:
  - python-dotenv
  - mcp
  - google-genai
  - python-pptx
  - pywinauto
  - google-auth-oauthlib
  - google-api-python-client

## Setup

### 1. Install Required Packages
```bash
pip install python-dotenv mcp google-genai python-pptx pywinauto google-auth-oauthlib google-api-python-client
```

### 2. Set Up Environment Variables
Create a `.env` file in the project root with the following variables:
```
GEMINI_API_KEY=your_gemini_api_key
```

### 3. Set Up Gmail API
1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the Gmail API for your project
4. Create OAuth 2.0 credentials (Desktop application)
5. Download the credentials JSON file
6. Save the credentials file as `credentials.json` in your project directory

## Usage

### Running with PowerPoint and Email (example2-3.py)
```bash
python talk2mcp-2.py
```

This will:
1. Convert "INDIA" to ASCII values: [73, 78, 68, 73, 65]
2. Calculate the sum of exponentials of these values
3. Create a PowerPoint presentation with the result in a rectangle
4. Send an email with the results and the PowerPoint presentation attached

### Running with Gmail API (gmail_server.py)
```bash
python gmail_server.py --creds-file-path credentials.json --token-path token.json --dev
```

This will:
1. Start the Gmail API server with OAuth 2.0 authentication
2. Open a browser window for you to authorize the application
3. Save the token for future use

## Prompt Design and Reasoning Framework

The system uses a structured prompt design that follows these key principles:

### 1. Explicit Reasoning Instructions
- Step-by-step thinking process
- Clear validation requirements
- Pre-action verification

### 2. Structured Output Format
- Consistent response patterns
- Clear function call syntax
- Predictable result formats

### 3. Tool Usage Guidelines
- Clear separation of reasoning and tool usage
- Specific parameters for each operation
- Validation before tool execution

### 4. Multi-turn Support
- Context maintenance between iterations
- Clear continuation prompts
- Result tracking and validation

### 5. Error Handling
- Comprehensive error detection
- Alternative approaches
- Input validation
- Tool availability checks

## Process Flow
1. **Text Conversion**: The input text "INDIA" is converted to ASCII values
2. **Mathematical Calculation**: The exponential sum of these ASCII values is calculated
3. **PowerPoint Visualization**: 
   - A PowerPoint presentation is created
   - A rectangle is drawn on the first slide
   - The final result is added inside the rectangle
4. **Email Sending**:
   - An email is sent to padmanabhbosamia@gmail.com
   - The email includes the input text, ASCII values, and final result
   - The PowerPoint presentation is attached to the email

## Functions

### Main Tools
- `strings_to_chars_to_int`: Converts a string to a list of ASCII values
- `int_list_to_exponential_sum`: Calculates the sum of exponentials of a list of integers

### PowerPoint Operations
- `open_powerpoint`: Opens a new PowerPoint presentation
- `draw_rectangle`: Draws a rectangle on the first slide
- `add_text_in_powerpoint`: Adds text to the first slide
- `close_powerpoint`: Closes PowerPoint

### Email Operations
- `send_gmail`: Sends an email using Gmail API with optional attachment

## Error Handling
The code includes robust error handling for:
- Parameter validation
- PowerPoint operations
- Email sending
- File operations
- LLM response validation
- Tool execution verification

## Notes
- The rectangle in PowerPoint is drawn at coordinates (2,2) to (7,5)
- The program will automatically close PowerPoint after operations
- The email is sent to erav3group@gmail.com
- The PowerPoint file is saved as 'presentation.pptx'
- The prompt enforces strict response formats and validation

## Gmail API Authentication
The Gmail API implementation uses OAuth 2.0 for authentication:
1. The first time you run the server, it will open a browser window
2. You'll need to log in to your Google account and authorize the application
3. The authorization token will be saved for future use
4. The token will be refreshed automatically when it expires

## Troubleshooting
- If PowerPoint doesn't open, ensure it's installed and accessible
- If email sending fails, check your Gmail API credentials and authorization
- If the rectangle doesn't appear, try adjusting the coordinates
- If the text doesn't appear, check if PowerPoint is in the correct state
- If LLM responses are unexpected, verify the prompt structure and tool descriptions

## License
This project is provided as-is for educational purposes. 
