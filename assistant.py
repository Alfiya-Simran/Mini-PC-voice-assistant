import speech_recognition as sr
import pyttsx3
import pyautogui
import subprocess
import webbrowser
import keyboard
import tkinter as tk
import threading
import time
import screen_brightness_control as sbc
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import urllib.parse 
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import datetime
import requests
import os
import random
import pygame
import pywhatkit as kit
import spotipy
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import pyperclip
from spotipy.oauth2 import SpotifyOAuth
recognizer = sr.Recognizer()
engine = pyttsx3.init()
listening = True 
# For volume control
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
current_site = ""
last_search_term = ""
pygame.mixer.init()
sp = spotipy.Spotify(auth_manager=SpotifyOAuth(client_id="d215475aac53450c864cc8ba3c3422f5",
                                               client_secret="d22d3d7c8e544a7bb03719975a015b43",
                                               redirect_uri="http://127.0.0.1:8888/callback",
                                               scope=["user-library-read", "user-modify-playback-state", "user-read-playback-state", "playlist-modify-public"]))
def speak(text):
    engine.say(text)
    engine.runAndWait()

def handle_command(command):
    global current_site
    if "hello" in command:
        speak("Hello, Hope you are doing great, How can i help you today!")
    elif control_chrome(command):
        return
    elif "what is the time" in command:
        now = datetime.datetime.now().strftime("%I:%M %p")
        speak(f"The current time is {now}")
    elif "what is the weather in" in command:
        city = command.split("weather in", 1)[1].strip()
        if city:
            speak(f"Getting weather in {city}")
            get_weather(city)
        else:
            speak("Please tell me a city after saying 'weather in'.")
    elif "scroll up" in command:
        pyautogui.scroll(500)
        speak("Scrolling up")

    elif "scroll down" in command:
        pyautogui.scroll(-500)
        speak("Scrolling down")

    elif "switch window" in command or "next window" in command:
        keyboard.press_and_release("alt+tab")
        speak("Switching window")
    elif "minimise window" in command or "minimize window" in command:
        pyautogui.hotkey("win", "down")
        speak("Minimizing window")
    elif "maximise window" in command or "maximize window" in command:
        pyautogui.hotkey("win", "up")
        speak("Maximizing window")
    elif "start dictation" in command:
        start_dictation()
    elif "erase" in command or "delete the word" in command or "remove the word" in command:
        word = command.split("word", 1)[1].strip()
        if word:
            erase_specific_word(word)
        else:
            speak("Please specify which word to erase.")
    elif "clear everything" in command or "remove all text" in command or "delete everything" in command:
        pyautogui.hotkey("ctrl", "a")      # Select all
        pyautogui.press("backspace")       # Delete
        speak("Cleared all the text.")
    elif "delete last line" in command or "remove previous line" in command:
        pyautogui.hotkey("shift", "home")   # Select from cursor to start of line
        pyautogui.press("backspace")        # Delete line
        speak("Deleted the last line.")


    elif "open notepad" in command:
        subprocess.Popen("notepad.exe")
        speak("Opening Notepad")
    elif "open doc" in command:
        webbrowser.open("https://docs.google.com/document/u/0/")
        speak("opening Google Doc")
    elif "new doc file" in command or "create new document" in command:
        speak("Creating a new document in Google Docs")
        webbrowser.open("https://docs.new")
    elif "save document" in command:
        speak("saving document")
        pyautogui.hotkey("ctrl","s")
    elif "open slides" in command:
        webbrowser.open("https://docs.google.com/presentation/u/0/")
        speak("opening Google Slides")
    elif "start presentation" in command or "start presenting" in command:
        speak("Starting presentation")
        time.sleep(1)
        pyautogui.hotkey("ctrl", "f5")
    elif "next slide" in command:
        pyautogui.hotkey("right")
    elif "previous slide" in command:
        pyautogui.hotkey("left")
    elif "slide" in command and any(word in command for word in ["go to", "number", "open", "jump to"]):
        match = re.search(r'\d+', command)
        if match:
            slide_num = int(match.group())
            pyautogui.write(str(slide_num))
            pyautogui.press("enter")
            speak(f"Jumping to slide {slide_num}")
        elif "first slide" in command:
            pyautogui.write("1")
            pyautogui.press("enter")
            speak("Opening the first slide")
        elif "last slide" in command:
            total_slides = get_total_slides_interactive()
            if total_slides:
                pyautogui.write(str(total_slides))
                pyautogui.press("enter")
                speak(f"Opening slide {total_slides}")
            else:
                speak("I couldn't understand the slide number.")
    elif "stop presenting" in command:
        pyautogui.hotkey("esc")
        speak("Ending Presentation")
    elif "new file" in command:
        subprocess.Popen("notepad.exe")
        time.sleep(0)  # Give it a moment to open :)
        pyautogui.hotkey("ctrl", "n")
        speak("New file created")
    elif command.startswith("type") or command.startswith("write"):
        text_to_type = command.replace("type", "").replace("write", "").strip()
        if text_to_type:
            pyautogui.write(text_to_type, interval=0.05)
            speak(f"Typing: {text_to_type}")
        else:
            speak("What should I type?")
    
    elif "save this file" in command or "save" in command:
        pyautogui.hotkey("ctrl","s")
        speak("file saved")
    elif "select all" in command:
        pyautogui.hotkey("ctrl","a")
        speak("Selecting")
    elif "copy" in command:
        pyautogui.hotkey("ctrl", "c")
        speak("Copying")
    elif "paste" in command:
        pyautogui.hotkey("ctrl", "v")
        speak("Pasting")
    elif "undo" in command:
        pyautogui.hotkey("ctrl", "z")
        speak("Undoing last action")
    elif "redo" in command:
        pyautogui.hotkey("ctrl", "y")
        speak("Redoing last action")
    elif "open calculator" in command:
        subprocess.Popen("calc.exe")
        speak("Opening Calculator")
    elif "open gmail" in command or "check my email" in command or "go to inbox" in command:
        speak("opening gmail")
        webbrowser.open("https://mail.google.com")
    elif "compose email" in command:
        speak("Opening Gmail to compose a new email")
        webbrowser.open("https://mail.google.com/mail/u/0/#inbox?compose=new")
    
    elif "open youtube" in command:
        webbrowser.open("https://youtube.com")
        current_site = "youtube"
        speak("Opening YouTube")
    elif "search for" in command:
        if "in chrome" in command:
            try:
                search_term = command.split("search for", 1)[1].split("in chrome")[0].strip()
                if search_term:
                    speak(f"Searching for {search_term} in Chrome")
                    os.system("start chrome")
                    time.sleep(3)
                    pyautogui.hotkey("ctrl", "l")
                    time.sleep(0.5)
                    pyautogui.write(f"https://www.google.com/search?q={search_term}", interval=0.05)
                    pyautogui.press("enter")
                else:
                    speak("Please specify a search term.")
            except Exception as e:
                print(e)
                speak("Something went wrong while searching in Chrome.")
        else:
            search_term = command.split("search for", 1)[1].strip()
            global last_search_term
            last_search_term = search_term
            if current_site == "youtube":
                speak(f"Searching for {search_term} on YouTube")
                time.sleep(5)
                pyautogui.press("tab")
                pyautogui.press("tab")
                pyautogui.press("tab")
                pyautogui.press("tab")
                pyautogui.write(search_term, interval=0.05)
                pyautogui.press("enter")
            elif current_site == "google":
                speak(f"Searching for {search_term} on Google")
                time.sleep(5)
                pyautogui.write(search_term, interval=0.05)
                pyautogui.press("enter")
            else:
                speak("Please open a site first like YouTube or Google, or say 'search for ... in Chrome'")
    elif "click on" in command:
        click_target = command.split("click on", 1)[1].strip()
        if "in chrome" in click_target:
            click_target = click_target.replace("in chrome", "").strip()
            speak(f"Trying to click on {click_target} in Chrome")
            time.sleep(2)
            pyautogui.hotkey("ctrl", "f")  # Open Find bar
            time.sleep(1)
            pyautogui.write(click_target, interval=0.05)
            time.sleep(1)
            pyautogui.press("esc")  # Exit Find bar
            pyautogui.press("tab")  # Move to matched result
            pyautogui.press("enter")  # Click it
        elif current_site == "youtube" and click_target:
            speak(f"Trying to click on {click_target} in YouTube search results")
            time.sleep(2)
            pyautogui.hotkey("ctrl", "f")
            time.sleep(1)
            pyautogui.write(click_target, interval=0.05)
            time.sleep(1)
            pyautogui.press("esc")
            pyautogui.press("tab")
            pyautogui.press("enter")
        elif current_site == "google" and click_target:
            speak(f"Trying to click on {click_target} in Google search results")
            time.sleep(2)
            pyautogui.hotkey("ctrl", "f")
            time.sleep(1)
            pyautogui.write(click_target, interval=0.05)
            time.sleep(1)
            pyautogui.press("esc")
            pyautogui.press("tab")
            pyautogui.press("enter")
        else:
            speak("Please open or search in Google or YouTube or say 'click on ... in Chrome'")

    elif "get mouse position" in command or "what is my mouse position" in command:
        x, y = pyautogui.position()
        speak(f"Your mouse is at {x} and {y}")
        print(f"Mouse position: ({x}, {y})")
    elif "close tab" in command:
        pyautogui.hotkey('ctrl', 'w')  # Close the current tab
        speak("Closing the current tab")
    elif "take screenshot" in command:
        screenshot = pyautogui.screenshot()
        screenshot.save("screenshot.png")
        speak("Screenshot taken and saved")
    elif "increase brightness" in command:
        try:
            sbc.set_brightness("+10")
            speak("Increased brightness")
        except:
            speak("Could not adjust brightness")

    elif "decrease brightness" in command:
        try:
            sbc.set_brightness("-10")
            speak("Decreased brightness")
        except:
            speak("Could not adjust brightness")

    elif "increase volume" in command:
        current = volume.GetMasterVolumeLevelScalar()
        volume.SetMasterVolumeLevelScalar(min(current + 0.1, 1.0), None)
        speak("Increased volume")

    elif "decrease volume" in command:
        current = volume.GetMasterVolumeLevelScalar()
        volume.SetMasterVolumeLevelScalar(max(current - 0.1, 0.0), None)
        speak("Decreased volume")

    elif "mute the sound" in command:
        volume.SetMute(1, None)
        speak("Muted")
#not working
    elif "unmute the sound" in command:
        volume.SetMute(0, None)
        speak("Unmuted")
    elif "define" in command or "what does" in command:
        word = command.split("define", 1)[1].strip() if "define" in command else command.split("what does", 1)[1].strip()
        get_definition(word)
    elif "play" in command:
        song_query = command.split("play", 1)[1].strip()
        play_music_on_youtube(song_query)
    elif "play on spotify" in command:
        song_query = command.split("play on spotify", 1)[1].strip()
        play_song_on_spotify(song_query)

    elif "close tab" in command:
        pyautogui.hotkey('ctrl', 'w')  # Close the current tab
        speak("Closing the current tab")
    elif "close window" in command:
        pyautogui.hotkey('alt','f4')
        speak("Closing the current window")
    elif "latest news" in command or "give me the news" in command:
        speak("Fetching the latest news for you.")
        speak_news()  # Call the function to fetch and speak news
    elif "exit" in command or "Bye" in command:
        speak("Have a great day!")
        root.quit()

    else:
        speak("Command not recognized")
def get_total_slides_interactive():
    speak("Please paste the Google Slides URL into the terminal.")
    slides_url = input("Enter Google Slides URL: ")

    try:
        service = Service("chromedriver.exe")
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        driver = webdriver.Chrome(service=service, options=options)

        driver.get(slides_url)
        time.sleep(5)

        # Count slide thumbnails
        slide_thumbnails = driver.find_elements(By.CSS_SELECTOR, "div[role='listitem']")
        total = len(slide_thumbnails)
        driver.quit()
        speak(f"This presentation has {total} slides")
        return total

    except Exception as e:
        print(f"Error detecting slides: {e}")
        speak("Failed to get the number of slides")
        return None
def erase_specific_word(word):
    try:
        pyautogui.hotkey("ctrl", "a")     # Select all
        pyautogui.hotkey("ctrl", "c")     # Copy selected text
        time.sleep(0.1)                   # Small delay for clipboard
        content = pyperclip.paste()

        if word in content:
            new_content = content.replace(word, "")
            pyautogui.press("backspace")  # Clear existing
            pyautogui.write(new_content, interval=0.03)
            speak(f"Removed the word {word}")
        else:
            speak(f"The word {word} was not found.")
    except Exception as e:
        print("Error:", e)
        speak("Sorry, I couldn't erase that word.")
def process_spoken_text(text):
    replacements = {
        " comma": ",",
        " period": ".",
        " question mark": "?",
        " exclamation mark": "!",
        " new line": "\n",
        " next line": "\n",
        " colon": ":",
        " semicolon": ";",
        " dash": "-",
        " open bracket": "(",
        " close bracket": ")",
        " slash": "/",
        " backslash": "\\"
    }
    for word, symbol in replacements.items():
        text = text.replace(word, symbol)
    return text
def start_dictation():
    speak("Dictation started. I'm listening... Say 'stop dictation' to end.")
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)
        while True:
            try:
                print("Listening for dictation...")
                audio = recognizer.listen(source, timeout=6)
                command = recognizer.recognize_google(audio).lower()

                if "stop dictation" in command:
                    speak("Dictation ended.")
                    break

                typed_text = process_spoken_text(command)
                pyautogui.write(typed_text + " ", interval=0.04)

            except sr.WaitTimeoutError:
                continue
            except sr.UnknownValueError:
                continue
            except sr.RequestError:
                speak("API error during dictation")
                break

def search_and_click(query, click_text):
    try:
        service = Service("chromedriver.exe")  # or path to your chromedriver
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        driver = webdriver.Chrome(service=service, options=options)
        driver.get(f"https://www.google.com/search?q={urllib.parse.quote(query)}")
        speak(f"Searching Google for {query}")
        time.sleep(3)  # wait for page to load

        # Try clicking a result that contains the spoken phrase
        links = driver.find_elements(By.XPATH, "//h3")
        for link in links:
            if click_text.lower() in link.text.lower():
                link.click()
                speak(f"Opening {click_text}")
                return

        speak("Sorry, I couldn't find that result.")
    except Exception as e:
        print(e)
        speak("Something went wrong while searching.")
def get_weather(city):
    api_key = "acf2bec6364b3ef36062702281c438c9"  # <- Replace this with your actual API key
    base_url = "http://api.openweathermap.org/data/2.5/weather?"
    complete_url = f"{base_url}q={city}&appid={api_key}&units=metric"

    try:
        response = requests.get(complete_url)
        data = response.json()

        # Show full response in console for debugging
        print(data)

        if data.get("cod") != 200:
            error_msg = data.get("message", "Unable to retrieve weather info.")
            speak(f"Couldn't get weather for {city}. Error: {error_msg}")
            return

        weather = data["weather"][0]["description"]
        temp = data["main"]["temp"]
        humidity = data["main"]["humidity"]
        pressure = data["main"]["pressure"]

        weather_report = (
            f"The weather in {city} is {weather}. "
            f"Temperature is {temp} degrees Celsius. "
            f"Humidity is {humidity} percent and pressure is {pressure} hPa."
        )
        speak(weather_report)

    except Exception as e:
        print("Exception occurred:", e)
        speak("Sorry, something went wrong while getting the weather.")
def get_definition(word):
    url = f"https://api.dictionaryapi.dev/api/v2/entries/en/{word}"
    try:
        response = requests.get(url)
        data = response.json()

        if "title" in data and data["title"] == "No Definitions Found":
            speak("Sorry, I couldn't find a definition for that word.")
            return

        # Extracting the definition
        definition = data[0]["meanings"][0]["definitions"][0]["definition"]
        speak(f"The definition of {word} is: {definition}")

    except requests.exceptions.RequestException as e:
        speak("Sorry, there was an error with the dictionary service.")
        print(e)
def get_news():
    api_key = "c5482c0721d0430f94a284efa3a571d1"  # Replace with your NewsAPI key
    url = f'https://newsapi.org/v2/top-headlines?country=us&apiKey={api_key}'

    try:
        response = requests.get(url)
        data = response.json()

        if data["status"] == "ok":
            articles = data["articles"]

            # Get top 5 headlines
            headlines = []
            for i, article in enumerate(articles[:10]):  # Get top 5 headlines
                headline = article["title"]
                headlines.append(f"Headline {i+1}: {headline}")

            return headlines
        else:
            return ["Sorry, I couldn't fetch the news at the moment."]
    
    except Exception as e:
        return [f"An error occurred: {e}"]
def speak_news():
    headlines = get_news()

    # Initialize the speech engine
    engine = pyttsx3.init()

    # Speak each headline
    for headline in headlines:
        print(headline)  # Optional: Print headlines in console
        engine.say(headline)
        engine.runAndWait()
def play_music_on_youtube(song_query):
    kit.playonyt(song_query)
    speak(f"Playing {song_query} on YouTube.")
def play_song_on_spotify(song_name):
    try:
        # Search for the song on Spotify
        results = sp.search(q=song_name, limit=1, type='track')
        if results['tracks']['items']:
            track_name = results['tracks']['items'][0]['name']
            track_uri = results['tracks']['items'][0]['uri']

           
            web_url = f"https://open.spotify.com/track/{track_uri.split(':')[2]}"  # Get the track ID from the URI
            webbrowser.open(web_url)  # Open the song in Spotify's web player
            speak(f"Opening {track_name} on Spotify.")
        else:
            speak(f"Sorry, I couldn't find {song_name} on Spotify.")
    except Exception as e:
        speak("Sorry, there was an error while searching for the song on Spotify.")
        print(f"Error: {e}")
def control_chrome(command):
    if "open chrome" in command:
        os.system("start chrome")
        speak("Opening Chrome")
        return True
    elif "open new tab" in command:
        pyautogui.hotkey("ctrl", "t")
        speak("Opening new tab")
        return True
    elif "close tab" in command:
        pyautogui.hotkey("ctrl", "w")
        speak("Closing tab")
        return True
    elif "previous tab" in command:
        pyautogui.hotkey("ctrl", "shift", "tab")
        speak("Switching to previous tab")
        return True
    elif "open incognito" in command:
        pyautogui.hotkey("ctrl", "shift", "n")
        speak("Opening incognito mode")
        return True
    elif "reload" in command:
        pyautogui.hotkey("ctrl", "r")
        speak("Refreshing the page")
        return True
    elif "close chrome" in command:
        pyautogui.hotkey("alt", "f4")
        speak("Closing Chrome")
        return True
    return False

def continuous_listen():
    global listening
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)
        speak("Hello, i am your voice assisstant Ana")

        while True:
            try:
                audio = recognizer.listen(source, timeout=5)
                command = recognizer.recognize_google(audio).lower()
                print("You said:", command)
                handle_command(command)
            except sr.WaitTimeoutError:
                continue  # No speech detected, continue listening
            except sr.UnknownValueError:
                continue  # Unrecognized speech, ignore
            except sr.RequestError:
                speak("API error")
                break
# Floating mic button
root = tk.Tk()
root.title("Mic")
root.geometry("80x80")
root.attributes("-topmost", True)
root.overrideredirect(True)
mic_button = tk.Button(root, text="🎤", font=("Arial", 20), command=continuous_listen, bg="black", fg="white")
mic_button.pack(expand=True, fill="both")

# Bottom-right corner
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = screen_width - 100
y = screen_height - 150
root.geometry(f"+{x}+{y}")
threading.Thread(target=continuous_listen, daemon=True).start()
root.mainloop()
