import wx
import pandas as pd
import glob
import webbrowser
import os
import openai
from openai import OpenAI
import re
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeDriverService
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import traceback
import time
import datetime
import threading
from dotenv import load_dotenv
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='review_responder.log'
)
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()

# Initialize OpenAI client with API key from environment
client = OpenAI(
    api_key=os.getenv('OPENAI_API_KEY')
)

# Define the base prompt template for review responses
RESPONSE_PROMPT = """
As the owner of a business, you need to respond to a customer review (in the same language as the review).
Consider the following business policies in your response:
- Entry fee based service
- Special discounts for specific conditions
- Service through waiters or app
- Unlimited orders with penalties for waste
- Mandatory beverage order

Please respond directly to the review without repeating any part of this prompt.
Here's the review: {}
"""

def get_latest_file(path):
    """
    Get the most recently downloaded Excel file from the specified path.
    
    Args:
        path (str): Path to search for Excel files
        
    Returns:
        str: Path to the most recent Excel file
    """
    try:
        list_of_files = glob.glob(os.path.join(path, '*.xlsx'))
        if not list_of_files:
            raise FileNotFoundError("No Excel files found in the specified directory")
        return max(list_of_files, key=os.path.getctime)
    except Exception as e:
        logger.error(f"Error finding latest file: {e}")
        raise

def generate_ai_response(review_text, max_retries=3):
    """
    Generate an AI response for a review using OpenAI's API.
    
    Args:
        review_text (str): The review text to respond to
        max_retries (int): Maximum number of retry attempts
        
    Returns:
        str: Generated response text
    """
    for attempt in range(max_retries):
        try:
            # Truncate review text to prevent token limit issues
            truncated_review = review_text[:700]
            
            if truncated_review.lower() == "nan":
                truncated_review = "Thank you for visiting our establishment."
                
            messages = [
                {
                    "role": "system", 
                    "content": "You are a business owner responding to customer reviews."
                },
                {
                    "role": "user", 
                    "content": RESPONSE_PROMPT.format(truncated_review)
                }
            ]
            
            response = client.chat.completions.create(
                model=os.getenv('OPENAI_MODEL_NAME', "ft:gpt-3.5-turbo-1106:default"),
                messages=messages,
                max_tokens=200,
                temperature=0.40
            )
            
            return response.choices[0].message.content.strip()
            
        except Exception as e:
            logger.error(f"Error generating AI response (attempt {attempt + 1}): {e}")
            if attempt == max_retries - 1:
                return "Response unavailable at this time."
            time.sleep(2 ** attempt)  # Exponential backoff

[Rest of the code with similar improvements...]
