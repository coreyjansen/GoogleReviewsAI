# GoogleReviewsAI

A Python application for Windows that automates the process of responding to Google Business reviews using AI-generated responses. The application uses OpenAI's API with a fine-tuned model to generate contextually appropriate responses based on your business policies.

## Features

- Extracts Google Business reviews from Outscraper-generated Excel files
- Generates AI responses using a fine-tuned OpenAI model
- Provides a GUI interface for review management
- Allows manual review and editing of AI-generated responses
- Automatically posts responses to Google Business reviews
- Supports background processing for seamless operation

## Prerequisites

- Python 3.8+
- Google Chrome browser
- Windows OS
- OpenAI API access with a fine-tuned model
- Chrome profile with access to your Google Business account
- Outscraper.com account

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/GoogleReviewsAI.git
cd GoogleReviewsAI
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

3. Create a `.env` file in the project root with the following structure:
```env
# OpenAI API key for your fine-tuned model
OPENAI_API_KEY=your-api-key-here

# Chrome user profile path for Selenium
profile_path=C:\\Users\\YourUsername\\AppData\\Local\\Google\\Chrome\\User Data

# Chrome profile directory
profile_directory=Default

# Excel file download path (usually your Downloads folder)
Ruta_Archivo_Excel=C:\\Users\\YourUsername\\Downloads
```

## Outscraper Configuration

1. Create an account at [Outscraper.com](https://outscraper.com)
2. Set up a new scraping task for your Google Business reviews
3. Important: Select ALL available columns when configuring the export
4. Required columns for the script:
   - review_text
   - owner_answer
   - author_title
   - review_rating
   - review_datetime_utc
   - review_link
   - reviews_link
   - name
   - reviews
   - rating

5. Download the Excel file to the path specified in your `.env` file

## Configuration

1. Setup your Chrome profile:
   - Make sure you're logged into your Google Business account
   - The profile path in `.env` should point to your Chrome user data directory

2. Fine-tune your OpenAI model:
   - Create a fine-tuned model with your business policies and response style
   - Update the model name in the code (currently set as "ft:gpt-3.5-turbo-1106:your-model-name")

## Usage

1. Download your latest reviews from Outscraper.com
2. Run the application:
```bash
python main.py
```

3. The GUI will display:
   - Business information from the Outscraper export
   - Recent reviews
   - AI-generated responses
   - Options to edit responses
   - Buttons to navigate through reviews

4. For each review:
   - Review the AI-generated response
   - Edit if necessary
   - Click "Responder" to post the response
   - The response will be posted in the background

[Rest of the README remains the same...]

## File Structure

```
GoogleReviewsAI/
├── main.py            # Main application file
├── requirements.txt   # Python dependencies
├── .env              # Environment variables
└── README.md         # Documentation
```

## Dependencies

- wxPython: GUI framework
- pandas: Data manipulation
- selenium: Web automation
- openai: AI response generation
- python-dotenv: Environment variable management
- Additional dependencies listed in requirements.txt

## Important Notes

- The application requires a logged-in Chrome profile with access to Google Business
- Responses are generated based on your fine-tuned model's training
- Background processing allows for continuous GUI interaction
- Excel files with reviews should be in the specified download directory

## Security Considerations

- Keep your `.env` file secure and never commit it to version control
- Regularly update your API keys
- Use secure Chrome profiles
- Monitor API usage and costs

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request
