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
from selenium.common.exceptions import (
    StaleElementReferenceException,
    TimeoutException,
    NoSuchElementException
)
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
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

# Base prompt template for the AI model
RESPONSE_PROMPT = """
As a business owner, respond to this customer review (in the same language as the review).
The business is called Shorty's Plumbing and Heating. They are a plumbing and heating company in Winnipeg Manitoba Canada.

Please respond directly:
Review: {}
"""

def get_latest_file(path):
    """Get the most recently downloaded Excel file from the specified path."""
    try:
        list_of_files = glob.glob(os.path.join(path, '*.xlsx'))
        if not list_of_files:
            raise FileNotFoundError("No Excel files found in specified directory")
        return max(list_of_files, key=os.path.getctime)
    except Exception as e:
        logger.error(f"Error finding latest file: {e}")
        raise

def load_review_data():
    """Load and prepare review data from the Excel file."""
    try:
        excel_path = os.getenv('Ruta_Archivo_Excel')
        if not excel_path:
            raise ValueError("Excel file path not found in environment variables")
        
        df = pd.read_excel(
            get_latest_file(excel_path),
            usecols=['review_text', 'owner_answer']
        )
        
        # Filter for unanswered reviews
        df = df[df['owner_answer'].isna()]
        df['review_text'] = df['review_text'].astype(str)
        
        return df['review_text'].tolist()
    except Exception as e:
        logger.error(f"Error loading review data: {e}")
        raise

def generate_ai_response(review_text, max_retries=3):
    """Generate AI response for a review using OpenAI's API."""
    for attempt in range(max_retries):
        try:
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
                model=os.getenv('OPENAI_MODEL_NAME', "gpt4o"),
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

class ReviewFrame(wx.Panel):
    """Panel for displaying and managing individual reviews."""
    def __init__(self, parent):
        super().__init__(parent)
        self.data = None
        self.init_controls()
        self.create_layout()

    def init_controls(self):
        """Initialize all UI controls."""
        # Author section
        self.author_title_label = wx.StaticText(self, label="Author:")
        self.author_title_text = wx.StaticText(self)
        
        # Review section
        self.review_text_label = wx.StaticText(self, label="Review:")
        self.review_text_text = wx.TextCtrl(self, style=wx.TE_READONLY | wx.TE_MULTILINE)
        
        # AI Response section
        self.airesponse_label = wx.StaticText(self, label="Generated Response:")
        self.airesponse_text = wx.TextCtrl(self, style=wx.TE_MULTILINE)
        
        # Rating and Date
        self.review_rating_label = wx.StaticText(self, label="Stars:")
        self.review_rating_text = wx.StaticText(self)
        self.review_datetime_utc_label = wx.StaticText(self, label="Date:")
        self.review_datetime_utc_text = wx.StaticText(self)
        
        # Review link
        self.review_link_label = wx.StaticText(self, label="Go to Review", style=wx.CURSOR_HAND)
        self.review_link_label.SetForegroundColour(wx.BLUE)
        self.review_link_label.Bind(wx.EVT_LEFT_DOWN, self.on_hyperlink)
        
        # Response status
        self.responded_label = wx.StaticText(self, label="")
        
        # Response button
        self.respond_button = wx.Button(self, label="Respond")
        self.respond_button.Bind(wx.EVT_BUTTON, self.start_responding)
        self.create_layout()

    def create_layout(self):
        sizer = wx.BoxSizer(wx.VERTICAL)
        grid_sizer = wx.FlexGridSizer(2, 5, 5)
        grid_sizer.AddMany([(self.author_title_label, 0, wx.EXPAND),
                            (self.author_title_text, 0, wx.EXPAND),
                            (self.review_text_label, 0, wx.EXPAND),
                            (self.review_text_text, 1, wx.EXPAND),
                            (self.airesponse_label, 0, wx.EXPAND),
                            (self.airesponse_text, 1, wx.EXPAND),
                            (wx.StaticText(self), 0), # Espacio en blanco para alinear el botón con airesponse_text
                            (self.respond_button, 0, wx.EXPAND),
                            #espacio para confirmar respuesta
                            (self.responded_label, 0, wx.EXPAND),
                            (self.review_rating_label, 0, wx.EXPAND),
                            (self.review_rating_text, 0, wx.EXPAND),
                            (self.review_datetime_utc_label, 0, wx.EXPAND),
                            (self.review_datetime_utc_text, 0, wx.EXPAND),
                            (wx.StaticText(self), 0),
                            (self.review_link_label, 0, wx.EXPAND)])
        grid_sizer.AddGrowableCol(1, 1)
        grid_sizer.AddGrowableRow(1, 1)
        sizer.Add(grid_sizer, 1, wx.ALL | wx.EXPAND, 5)
        self.SetSizer(sizer)

    def on_hyperlink(self, event):
        if self.data is not None:
            webbrowser.open(self.data['review_link'])

    def update_data(self, data):
        self.data = data
        self.author_title_text.SetLabel(str(data.get('author_title', 'N/A')))
        self.review_text_text.SetValue(str(data.get('review_text', 'N/A')))
        self.review_rating_text.SetLabel(str(data.get('review_rating', 'N/A')))
        self.review_datetime_utc_text.SetLabel(str(data.get('review_datetime_utc', 'N/A')))
        if 'review_link' in data:
            self.review_link_label.SetLabel("Go to the review")
            self.review_link_label.SetToolTip(str(data['review_link']))
        else:
            self.review_link_label.SetLabel("No link available")
            self.review_link_label.SetToolTip("")


    #actualizar el control self.airesponse_text para mostrar cada respuesta de la API
    def update_ai_response_text(self, response_text):
        self.airesponse_text.SetValue(response_text)
        print(f"response_text: {response_text}")

        try:
         self.airesponse_text.SetValue(response_text) 
        except Exception as e:
         print("Error al actualizar respuesta:", e)
   
     #create function to respond in google
    def on_respond_in_google(self, event=None):
        try:
            profile_path = r"C:\Users\Corey\AppData\Local\Google\Chrome\User Data"
            profile_directory = "Default"

            #chrome_options = webdriver.ChromeOptions()
            #chrome_options.add_argument(f"user-data-dir={profile_path}")
            #chrome_options.add_argument(f"--profile-directory={profile_directory}")
            #chrome_options.add_argument("--headless")
            #chrome_options.add_argument('window-size=1920x1080')

            options = Options()
            options.add_argument(f"user-data-dir={profile_path}")
            options.add_argument(f"--profile-directory={profile_directory}")
            #options.add_argument("--headless")
            options.add_argument("window-size=1920x1080")

            #chrome_driver_path = r"D:\OneDrive\ReseñasRespondedor\chromedriver-win64\chromedriver.exe"
            #chrome_service = ChromeDriverService(chrome_driver_path)
            #capabilities = DesiredCapabilities.CHROME.copy()
            #capabilities['loggingPrefs'] = {'browser': 'ALL'}
            #driver = webdriver.Chrome(service=chrome_service, options=chrome_options, desired_capabilities=capabilities)
            #driver.implicitly_wait(10)

            # Inicializa el servicio de ChromeDriver con ChromeDriverManager
            service = ChromeDriverService(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
            driver.implicitly_wait(10)


            review_link = self.data.get('reviews_link')
            if review_link:
                driver.get(review_link)

            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-sort-id="newestFirst"]')))
            time.sleep(5)
            recent_button = driver.find_element(By.CSS_SELECTOR, 'div[data-sort-id="newestFirst"]')
            recent_button.click()

            driver.switch_to.window(driver.window_handles[-1])
            time.sleep(10)

            WebDriverWait(driver, 30).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'a[href*="google.com/maps/contrib/"]')))
            try:
                WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, 'div.gws-localreviews__general-reviews-block')))
            except TimeoutException:
                print("Error: Las reseñas no se cargaron después de hacer clic en 'Más recientes'")
                return

            current_author = self.data['author_title']
            scroll_window = driver.find_element(By.CLASS_NAME, "review-dialog-list")
            for i in range(10):
                try:
                    author_element = driver.find_element(By.XPATH, f'//a[contains(text(), "{current_author}")]')
                    print(f"Checking author: {author_element.text}")  # Registro para ver qué autor está siendo verificado
                except NoSuchElementException:
                    driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", scroll_window)
                    time.sleep(3)

            first_review = driver.find_element(By.CSS_SELECTOR, 'a[href*="google.com/maps/contrib/"]')
            driver.execute_script("arguments[0].scrollIntoView();", first_review)
            actions = ActionChains(driver)
            actions.move_to_element(first_review).perform()

            all_authors = driver.find_elements(By.CSS_SELECTOR, 'a[href*="google.com/maps/contrib/"]')
            for author in all_authors:
                print(f"Checking author: {author.text}")  # Registro para ver qué autor está siendo verificado
                if current_author == author.text:
                    print("Author matched!")  # Registro cuando se encuentra una coincidencia
                    review_element = author.find_element(By.XPATH, './../../../..')
                    respond_button = review_element.find_element(By.XPATH, ".//*[contains(text(), 'Responder')]")
                    respond_button.click()

                    responseFrame = WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, '/local/business')]")))
                    driver.switch_to.frame(responseFrame)
                    response_textbox = driver.find_element(By.CSS_SELECTOR, 'textarea[aria-label="Tu respuesta pública"]')
                    response_textbox.send_keys(self.airesponse_text.GetValue())

                    submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.nCP5yc.AjY5Oe.DuMIQc.LQeN7.FwaX8')))
                    submit_button.click()
                    break

            for entry in driver.get_log('browser'):
                print(entry)

            driver.quit()
            wx.CallAfter(self.show_responded_message)

        except Exception as e:
            print(f"Error: {e}")

   
    def show_responded_message(self):
        self.responded_label.SetLabel("Respondida!")
        self.responded_label.SetForegroundColour(wx.Colour(0, 255, 0))  # Color verde
        self.Layout()  # Refrescar el layout para mostrar los cambios


    def start_responding(self, event):
        thread = threading.Thread(target=self.on_respond_in_google)
        thread.start() 



            

    

class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        super().__init__(parent, title=title)
        self.panel = wx.ScrolledWindow(self)
        self.panel.SetScrollbars(1, 1, 600, 400)
        self.current_page = 0
        self.data = pd.read_excel(self.get_latest_file())
        self.name_label = wx.StaticText(self.panel, label="Nombre:")
        self.name_text = wx.StaticText(self.panel)
        self.reviews_link_label = wx.StaticText(self.panel, label="Ver todas las reseñas", style=wx.CURSOR_HAND)
        self.reviews_link_label.SetForegroundColour(wx.BLUE)
        self.reviews_link_label.Bind(wx.EVT_LEFT_DOWN, self.on_hyperlink)
        self.reviews_label = wx.StaticText(self.panel, label="Número total de reseñas:")
        self.reviews_text = wx.StaticText(self.panel)
        self.rating_label = wx.StaticText(self.panel, label="Rating:")
        self.rating_text = wx.StaticText(self.panel)
        self.name_text.SetLabel(str(self.data['name'][0]))
        self.reviews_text.SetLabel(str(self.data['reviews'][0]))
        self.rating_text.SetLabel(str(self.data['rating'][0]))
        self.review_frames = [ReviewFrame(self.panel) for _ in range(3)]
        self.previous_button = wx.Button(self.panel, label="Anterior")
        self.next_button = wx.Button(self.panel, label="Siguiente")
        self.previous_button.Bind(wx.EVT_BUTTON, self.on_previous)
        self.next_button.Bind(wx.EVT_BUTTON, self.on_next)
        self.create_layout()
        self.update_reviews()

    def create_layout(self):
        sizer = wx.BoxSizer(wx.VERTICAL)
        grid_sizer = wx.FlexGridSizer(2, 5, 5)
        grid_sizer.AddMany([(self.name_label, 0, wx.EXPAND),
                            (self.name_text, 0, wx.EXPAND),
                            (self.reviews_link_label, 0, wx.EXPAND),
                            (wx.StaticText(self.panel), 0),
                            (self.reviews_label, 0, wx.EXPAND),
                            (self.reviews_text, 0, wx.EXPAND),
                            (self.rating_label, 0, wx.EXPAND),
                            (self.rating_text, 0, wx.EXPAND)])
        grid_sizer.AddGrowableCol(1, 1)
        sizer.Add(grid_sizer, 0, wx.ALL | wx.EXPAND, 5)
        for review_frame in self.review_frames:
            sizer.Add(review_frame, 1, wx.ALL | wx.EXPAND, 5)
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        button_sizer.Add(self.previous_button, 0, wx.ALL, 5)
        button_sizer.Add(self.next_button, 0, wx.ALL, 5)
        sizer.Add(button_sizer, 0, wx.ALIGN_CENTER)
        self.panel.SetSizer(sizer)
        self.Fit()

    def get_latest_file(self):
        #list_of_files = glob.glob('D:/Descargas/*.xlsx')
        list_of_files = glob.glob('C:/Users/Enriq/Downloads/*.xlsx')
        latest_file = max(list_of_files, key=os.path.getmtime)
        return latest_file

    def on_previous(self, event):
        if self.current_page > 0:
            self.current_page -= 1
            self.update_reviews()

    def on_next(self, event):
        if (self.current_page + 1) * 3 < len(self.data):
            self.current_page += 1
            self.update_reviews()

    def on_hyperlink(self, event):
        webbrowser.open(self.data['reviews_link'][0])

    def update_reviews(self):
        start_index = self.current_page * 3
        for i in range(3):
        # Si hay datos disponibles, llenar el control con esos datos
            if start_index + i < len(self.data):
                self.review_frames[i].update_data(self.data.iloc[start_index + i])
                response_index = start_index + i
                response_text = responses[response_index]
                self.review_frames[i].update_ai_response_text(response_text) 
            else:
            # Si no hay datos, limpie el control
                self.review_frames[i].update_data({})
                self.review_frames[i].update_ai_response_text("")
            
         
            
        


if __name__ == "__main__":
    app = wx.App()
    frame = MyFrame(None, "Respondedor Reseñas")
    frame.Show()
    app.MainLoop()

