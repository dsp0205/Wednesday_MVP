# Import necessary libraries
import json
import openai
import os
import requests
import re
import pyaudio
import numpy as np
from io import BytesIO
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import speech_recognition as sr
import io
import wave
import subprocess   
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeDriverService
import pytesseract
import cv2
import pyautogui
from PIL import Image
import win32com.client as win32
import pywinauto
from pywinauto.application import Application


# Load OpenAI API key and ChromeDriver path from config file
with open('config.json') as f:
    config = json.load(f)

openai.api_key = config['openai_api_key']

chromedriver_path = config['path_chromedriver']
chrome_options = webdriver.ChromeOptions()

# Configure ChromeDriver to run in incognito mode
chrome_options.add_argument("--incognito")
chromedriver_service = ChromeDriverService(executable_path=chromedriver_path)

# Initialize the webdriver
driver = webdriver.Chrome(service=chromedriver_service, options=chrome_options)


# Function to record audio using PyAudio
def record_audio(duration, sample_rate):
    pa = pyaudio.PyAudio()
    stream = pa.open(format=pyaudio.paInt16,
                     channels=1,
                     rate=sample_rate,
                     input=True,
                     frames_per_buffer=1024)

    frames = []
    for _ in range(0, int(sample_rate / 1024 * duration)):
        data = stream.read(1024)
        frames.append(data)

    stream.stop_stream()
    stream.close()
    pa.terminate()

    wav_buffer = io.BytesIO()
    with wave.open(wav_buffer, 'wb') as wav_file:
        wav_file.setnchannels(1)
        wav_file.setsampwidth(pa.get_sample_size(pyaudio.paInt16))
        wav_file.setframerate(sample_rate)
        wav_file.writeframes(b''.join(frames))

    wav_buffer.seek(0)
    return wav_buffer


# Function to transcribe audio using OpenAI's Whisper API
def transcribe_audio(audio_data):
    openai.api_key = config['openai_api_key']

    
    with open("temp_audio.wav", "wb") as f:
        f.write(audio_data.getbuffer())

    
    try:
        with open("temp_audio.wav", "rb") as audio_file:
            response = openai.Audio.transcribe("whisper-1", audio_file)
        
        print("API response:", response)  

        transcription = response['text']  
        return transcription
    except Exception as e:
        print("Error transcribing audio:", e)
        return None
    
    os.remove("temp_audio.wav")


# Function to get text from a screenshot using Tesseract OCR
def get_screenshot_text():
    screenshot = pyautogui.screenshot()
    screenshot.save("screenshot.png")

    img = cv2.imread("screenshot.png")
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    text = pytesseract.image_to_string(gray)

    os.remove("screenshot.png")
    return text


# Function to click on a specific text on screen using PyAutoGUI
def click_on_screen(text_to_find):
    try:
        location = pyautogui.locateOnScreen(text_to_find)
        if location:
            center = pyautogui.center(location)
            pyautogui.click(center)
        else:
            print(f"'{text_to_find}' not found on the screen.")
    except Exception as e:
        print("Error clicking on screen:", e)


# Function to search for images on Google Images
def search_for_images(search_term):
    driver.get(f"https://www.google.com/search?q={search_term}&tbm=isch")


# Function to open a website
def open_website(url):
    driver.get(url)


# Function to log in to LinkedIn
def login_linkedin(email, password):
    driver.get("https://www.linkedin.com/login")

    try:
        email_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'username'))
        )
    except NoSuchElementException:
        print("Error: Unable to locate the email field.")
        return

    print("Entering email...")
    email_field.send_keys(email)

    try:
        password_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'password'))
        )
    except NoSuchElementException:
        print("Error: Unable to locate the password field.")
        return

    print("Entering password...")
    password_field.send_keys(password)

    try:
        login_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'button[type="submit"]'))
        )
    except NoSuchElementException:
        print("Error: Unable to locate the login button.")
        return

    print("Clicking login button...")
    login_button.click()
    time.sleep(5)


# Function to search for jobs on LinkedIn
def search_linkedin_jobs(job_title):
    driver.get("https://www.linkedin.com/jobs")

    try:
        search_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[aria-label="Search by title, skill or company"]'))
        )
    except NoSuchElementException:
        print("Error: Unable to locate the search field.")
        return

    print("Entering search query...")
    search_field.send_keys(job_title)
    search_field.send_keys(Keys.RETURN)
    time.sleep(5)


# Function to log in to Twitter
def login_twitter(username, password):
    driver.get("https://twitter.com/i/flow/login")
    
    try:
        username_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="text"]'))
        )
    except NoSuchElementException:
        print("Error: Unable to locate the username field.")
        return

    print("Entering username...")
    username_field.send_keys(username)
    username_field.send_keys(Keys.RETURN)
    time.sleep(2)

    try:
        password_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="password"]'))
        )
    except NoSuchElementException:
        print("Error: Unable to locate the password field.")
        return

    print("Entering password...")
    password_field.send_keys(password)
    password_field.send_keys(Keys.RETURN)
    time.sleep(5)

    try:
        tweet_box = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//div[@role="textbox"]'))
        )
        tweet_box.send_keys(Keys.RETURN)
    except NoSuchElementException as e:
        print("Error: Unable to locate the tweet box or the tweet button.")
        print(e)



# Function to post a tweet on Twitter
def tweet(content):
    try:
        tweet_box = driver.find_element(By.XPATH, '//div[@aria-label="Tweet text"]')
        tweet_box.send_keys(content)
        time.sleep(1)

        tweet_button = driver.find_element(By.XPATH, '//div[@data-testid="tweetButtonInline"]')
        tweet_button.click()
        time.sleep(2)
    except NoSuchElementException as e:
        print("Error: Unable to locate the tweet box or the tweet button.")
        print(e)


# Function to generate a tweet using OpenAI's GPT-3
def generate_tweet(topic):
    prompt = f"Write an informative and engaging tweet about {topic}. The tweet should raise awareness and provide insights about the issue:"
    
    response = openai.Completion.create(
        engine="davinci",
        prompt=prompt,
        max_tokens=100,
        n=1,
        stop=None,
        temperature=0.7,
    )

    generated_tweet = response.choices[0].text.strip()
    return generated_tweet



# Function to generate a text using OpenAI's GPT-3
def generate_text(topic):
    prompt = f"Write a well-structured and informative paragraph about {topic}:"

    response = openai.Completion.create(
        engine='text-davinci-002',
        prompt=prompt,
        max_tokens=1024,
        n=1,
        stop=None,
        temperature=0.7
    )

    return response.choices[0].text.strip()


# Function to write the generated text in a Word document
def write_about_in_word(topic, generated_text):
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = True
    doc = word.Documents.Add()
    range = doc.Range()
    range.InsertAfter(f"Here's some information about {topic}:")
    range.InsertParagraphAfter()
    range.InsertAfter(generated_text)
    doc.SaveAs("output.docx")
  
    app = Application().connect(path=word.Path)
    app.top_window().maximize()
    


# Function to generate a command using OpenAI's GPT-3
def generate_command(prompt):
    open_website_pattern = re.compile(r'open\s+(\w+)', re.IGNORECASE)
    match = open_website_pattern.match(prompt)
    if match:
        website_name = match.group(1)
        url = f"https://www.{website_name}.com"
        return "start", url
    
    if "get screenshot text" in prompt.lower():
        return "get_screenshot_text"

    if "click on screen" in prompt.lower():
        text_to_find = prompt.lower().replace("click on screen", "", 1).strip()
        return "click_on_screen", text_to_find
    
    tweet_about_pattern = re.compile(r'tweet\s+about\s+(.+)', re.IGNORECASE)
    match = tweet_about_pattern.match(prompt)
    if match:
        topic = match.group(1).strip(" .,")  
        generated_tweet = generate_tweet(topic)
        return "tweet_about", generated_tweet

    search_images_pattern = re.compile(r'search\s+for\s+(.+\b)\s+images', re.IGNORECASE)
    match = search_images_pattern.match(prompt)
    if match:
        search_term = match.group(1)
        return "search_for_images", search_term

    search_linkedin_pattern = re.compile(r'search\s+for\s+(.+\b)\s+jobs\s+on\s+linkedin', re.IGNORECASE)
    match = search_linkedin_pattern.match(prompt)
    if match:
        job_title = match.group(1)
        return "search_linkedin_jobs", job_title

    search_pattern = re.compile(r'search\s+(.+)', re.IGNORECASE)
    match = search_pattern.match(prompt)

    if match:
        search_query = match.group(1)
        url = f'https://www.google.com/search?q={search_query}'
        return ('start', url)

    if "login to twitter" in prompt.lower() and "tweet about" in prompt.lower():
        tweet_content = prompt.lower().replace("login to twitter and tweet about", "", 1).strip()
        return "tweet_about", tweet_content

    write_about_pattern = re.compile(r'write\s+about\s+(.+)', re.IGNORECASE)
    match = write_about_pattern.match(prompt)
    if match:
        topic = match.group(1).strip(" .,")
        return "write_about_in_word", topic

    response = openai.Completion.create(
        engine="davinci",
        prompt=f"{prompt}",
        max_tokens=50,
        n=1,
        stop=None,
        temperature=0.5,
    )

    command = response.choices[0].text.strip()
    return command


# Function to execute the generated command
def execute_command(command):
    if isinstance(command, tuple):
        if command[0] == "start":
            open_website(command[1])
        elif command[0] == "search_for_images":
            search_for_images(command[1])
        elif command[0] == "tweet_about":
            tweet_text = command[1]
            tweet(tweet_text)
        elif command[0] == "write_about_in_word":
            topic = command[1]
            generated_text = generate_text(topic)
            write_about_in_word(topic, generated_text)
    else:
        try:
            result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True, text=True)
            print(result.stdout)
        except Exception as e:
            print("Error executing command:", e)

    if command == "get_screenshot_text":
        text = get_screenshot_text()
        print("Screenshot text:", text)
    elif command[0] == "click_on_screen":
        text_to_find = command[1]
        click_on_screen(text_to_find)

    if command == "login_to_twitter":
        login_twitter(config['twitter_username'], config['twitter_password'])
    elif command[0] == "tweet_about":
        tweet(command[1])
    elif command[0] == "search_linkedin_jobs":
        job_title = command[1]
        login_linkedin(config['linkedin_email'], config['linkedin_password'])
        search_linkedin_jobs(job_title)



# Main function to listen for commands and execute them
def main():
    print("Listening for commands...")
    sample_rate = 16000
    listen = True

    while True:
        if listen:
            try:
                print("Recording audio...")
                duration = 5
                audio_data = record_audio(duration, sample_rate)
                transcription = transcribe_audio(audio_data)
                print("You said:", transcription)

                if transcription is not None:
                    if "stop listening" in transcription.lower():
                        print("Stopping.")
                        break

                    if "wednesday" in transcription.lower():
                        listen = True
                        continue

                    command = generate_command(transcription)
                    print("Executing command:", command)
                    if command == "login_to_twitter":
                        login_twitter(config['twitter_username'], config['twitter_password'])
                    elif command[0] == "tweet_about":
                        tweet_text = command[1]
                        login_twitter(config['twitter_username'], config['twitter_password'])
                        tweet(tweet_text)

                    elif command[0] == "search_linkedin_jobs":
                        job_title = command[1]
                        login_linkedin(config['linkedin_email'], config['linkedin_password'])
                        search_linkedin_jobs(job_title)
                        
                    else:
                        execute_command(command)

                 
                    listen = False
                else:
                    print("No transcription found. Listening again...")

            except Exception as e:
                print("Error:", e)
        else:
            print("Waiting for the 'Hello' keyword to resume listening...")
            duration = 5
            audio_data = record_audio(duration, sample_rate)
            transcription = transcribe_audio(audio_data)
            print("You said:", transcription)

            if transcription is not None and "hello" in transcription.lower():
                listen = True

if __name__ == "__main__":
    try:
        main()
    finally:
        driver.quit()

