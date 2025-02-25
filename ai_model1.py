import pyttsx3
import pywhatkit
import datetime
import wikipedia
from datetime import date
import webbrowser
import os
import subprocess
import psutil
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
import requests
from bs4 import BeautifulSoup
import noisereduce as nr
import pyautogui
import time
import pytz
from timezonefinder import TimezoneFinder
from geopy.geocoders import Nominatim
import sys
from fuzzywuzzy import process
import webbrowser
import openai
from dotenv import load_dotenv
import speech_recognition as sr

# Load environment variables
load_dotenv()
api_key = os.getenv("GITHUB_TOKEN") # get your github token
if not api_key:
    print("Error: API key not found. Please check your .env file.")
else:
    print("API key loaded successfully.")

# Initialize OpenAI client
client = openai.OpenAI(
    base_url="https://models.inference.ai.azure.com",
    api_key=api_key,
)

# Initialize text-to-speech engine
engine = pyttsx3.init(driverName="sapi5")
engine.setProperty("rate", 190)
engine.setProperty("volume", 1.0)

current_mode = "text"  # Default to text mode

# dictionary for fuzzy matching
commands_dict = {
    "chrome": ["chrome", "chrom", "crome", "browsr", "google"],
    "notepad": ["notepad", "note", "notepd", "editor"],
    "calculator": ["calculator", "calc", "calcu"],
    "youtube": ["youtube", "yt", "utub", "you tube"],
    "screenshot": ["screenshot", "ss", "screen shot"],
    "shutdown": ["shutdown", "shut down", "exit", "quit"],
}

def get_best_match(user_input):
    """get  the best matching command using fuzzy matching"""
    best_match, score = process.extractOne(user_input, commands_dict.keys())
    if score > 70: 
        return best_match
    return None 

def talk(text):
    """Output text (and speech in voice mode)"""
    print("Assistant:", text)
    if current_mode == "voice":
        engine.say(text)
        engine.runAndWait()

def toggle_mode():
    """Toggle between text and voice modes"""
    global current_mode
    if current_mode == "text":
        current_mode = "voice"
        print("Switched to voice mode. Say 'switch to text' to change back.")
        talk("Voice mode activated. Say 'switch to text' to change back.")
    else:
        current_mode = "text"
        print("Switched to text mode. Type 'switch to voice' to change back.")
        talk("Text mode activated. Type 'switch to voice' to change back.")

def get_voice_input():
    """Get voice input from the user"""
    recognizer = sr.Recognizer()
    try:
        with sr.Microphone() as mic:
            print("Listening...")
            recognizer.adjust_for_ambient_noise(mic, duration=0.2)
            audio = recognizer.listen(mic, timeout=5, phrase_time_limit=5)
            
            text = recognizer.recognize_google(audio)
            text = text.lower()
            print(f"You (voice): {text}")
            return text
    except sr.UnknownValueError:
        print("Sorry, I did not understand that.")
        talk("Sorry, I did not understand that.")
        return ""
    except sr.RequestError as e:
        print(f"Error with the speech service: {e}")
        talk("Sorry, there was an error with the speech service.")
        return ""
    except Exception as e:
        print(f"Error: {e}")
        talk("Sorry, something went wrong while listening.")
        return ""

def get_user_input():
    """Get user input based on current mode"""
    if current_mode == "text":
        return input("You: ").lower().strip()
    else:
        voice_input = get_voice_input()
        # If voice recognition failed, fall back to text input
        if not voice_input:
            talk("Voice recognition failed. Please type your command.")
            return input("You: ").lower().strip()
        return voice_input

def search_google(query):
    """Search Google for a query"""
    url = f"https://www.google.com/search?q={query}"
    webbrowser.open(url)
    talk(f"Here are the results for {query}")

def get_greeting():
    """Get appropriate greeting based on time of day"""
    hour = int(time.strftime("%H"))
    if hour < 12:
        greeting = "Good Morning Sir!"
    elif hour < 15:
        greeting = "It's Noon Sir! You should rest"
    elif hour < 17:
        greeting = "Good Afternoon Sir!"
    elif hour < 19:
        greeting = "Good Evening Sir!"
    elif hour > 21:
        greeting = "It's late night sir! You should sleep"
    else:
        greeting = "It's Night sir!"
    
    talk(greeting)
    return greeting

def get_time_in_location(location="Bangladesh"):
    """Get the current time in a specific location"""
    geolocator = Nominatim(user_agent="geoapi")
    try:
        location_data = geolocator.geocode(location)
        if not location_data:
            talk("Sorry, I couldn't find that location.")
            return
        lat, lon = location_data.latitude, location_data.longitude
        tf = TimezoneFinder()
        timezone_str = tf.timezone_at(lng=lon, lat=lat)
        if timezone_str is None:
            talk("Sorry, I couldn't determine the timezone.")
            return
        timezone = pytz.timezone(timezone_str)
        local_time = datetime.datetime.now(timezone).strftime("%I:%M %p")
        talk(f"The current time in {location} is {local_time}")
    except Exception as e:
        print(f"Error: {e}")
        talk("An error occurred while fetching the time.")

def change_volume(increase=True):
    """Increase or decrease system volume"""
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = interface.QueryInterface(IAudioEndpointVolume)
    current_volume = volume.GetMasterVolumeLevelScalar()
    step = 0.1
    if increase:
        new_volume = min(current_volume + step, 1.0)
    else:
        new_volume = max(current_volume - step, 0.0)
    volume.SetMasterVolumeLevelScalar(new_volume, None)
    talk(f"Volume set to {int(new_volume * 100)} percent")

def open_vs_code_new_tab():
    """Open a new tab in VS Code"""
    vscode_path = r"Enter your vs code path"
    subprocess.Popen([vscode_path, "--new-window"])
    talk("Opening a new tab in VS Code")

def control_pc(command):
    """Control PC functions like shutdown, restart"""
    if "shutdown" in command:
        talk("Shutting down your PC.")
        os.system("shutdown /s /t 5")
    elif "restart" in command:
        talk("Restarting your PC.")
        os.system("shutdown /r /t 5")
    elif "open notepad" in command:
        talk("Opening Notepad")
        os.system("notepad")
    else:
        return False 
    return True  

def take_screenshot():
    """Take a screenshot and save it"""
    screenshot_folder = os.path.join(os.getcwd(), "Screenshots")
    os.makedirs(screenshot_folder, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    save_path = os.path.join(screenshot_folder, f"screenshot_{timestamp}.png")
    try:
        screenshot = pyautogui.screenshot()
        screenshot.save(save_path)
        talk("Screenshot taken and saved successfully")
    except Exception as e:
        print(f"Error taking screenshot: {e}")
        talk("Sorry, I couldn't take the screenshot")

def search_google_results(query):
    """Search Google and try to extract an answer"""
    try:
        url = f"https://www.google.com/search?q={query}"
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(url, headers=headers)
        soup = BeautifulSoup(response.text, "html.parser")
        answer = soup.find("div", class_="BNeawe").text  
        if not answer:
            answer = soup.find("span", class_="hgKElc").text  
        talk(answer)
    except Exception as e:
        print("Couldn't fetch search results.")
        talk("I couldn't find the answer.")

def close_application(app_name):
    """Close a running application by name"""
    for process in psutil.process_iter(attrs=['pid', 'name']):
        if app_name.lower() in process.info['name'].lower():
            talk(f"Closing {process.info['name']}")
            os.kill(process.info['pid'], 9)
            return
    talk(f"No running application found with name {app_name}")

def show_help():
    talk("Here’s a list of commands I can understand. You can use these in text or voice mode:")
    help_text = [
        "1. 'switch to voice' - Switch to voice input mode.",
        "2. 'switch to text' - Switch to text input mode.",
        "3. 'quit' - Exit the assistant.",
        "4. 'time' - Show current time in Bangladesh (or 'time in [location]').",
        "5. 'google [query]' - Search Google for something.",
        "6. 'play [song name]' - Play a song on YouTube.",
        "7. 'screenshot' - Take and save a screenshot.",
        "8. 'volume up' - Increase system volume.",
        "9. 'volume down' - Decrease system volume.",
        "10. 'explain [topic]' - Get a detailed explanation from Wikipedia.",
        "11. 'chat [topic]' - Talk to me about anything!",
        "12. 'open chrome' - Open Google Chrome.",
        "13. 'shutdown' - Shut down your PC.",
        "14. 'help' - Show this help menu."  
    ]
    for line in help_text:
        talk(line)
    talk("If I don’t recognize a command, I’ll try to answer using my AI knowledge!")

def chat_with_gpt(prompt):
    """Chat with GPT model"""
    try:
        response = client.chat.completions.create(
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt},
            ],
            model="gpt-4o-mini",
            temperature=0.8,
            max_tokens=200,
            top_p=0.5
        )
        answer = response.choices[0].message.content.strip()
        return answer
    except Exception as e:
        return f"Sorry, I couldn't process that request: {e}"

def run_assistant():
    """Main function to run the assistant"""
    print("Welcome to the AI Assistant!")
    print("You can type 'switch to voice' to use voice commands or 'switch to text' to use text commands.")
    print("Type 'quit' or say 'quit' to exit.")
    print("Type 'help' to see the avaliable commands")
    
    get_greeting()
    
    while True:
        command = get_user_input()
        
        if not command:
            continue
            
        # Mode switching commands
        if command == "switch to voice":
            toggle_mode()
            continue
        elif command == "switch to text":
            toggle_mode()
            continue
        elif command == "quit":
            talk("Shutting down. Have a nice day, sir!")
            sys.exit(0)
            
        if control_pc(command):
            continue
            
        # Try to get best match from command dictionary
        best_match = get_best_match(command)
        if best_match:
            command = best_match
            
        print(f"Debug - Recognized Command: {command}")
        
        # Handle different commands
        if "chat" in command or "talk to jarvis" in command:
            user_prompt = command.replace("chat", "").replace("talk to jarvis", "").strip()
            if not user_prompt:
                talk("What would you like to talk about?")
                user_prompt = get_user_input()
            answer = chat_with_gpt(user_prompt)
            print("GPT response:", answer)
            talk(answer)
           
        elif "play" in command:
            song = command.replace("play", "").strip()
            talk("Playing " + song)
            pywhatkit.playonyt(song)
        # some of the pre recommended song link give them a try or u can remove them 
        elif "hit the song" in command:
            talk("Playing the song")
            pywhatkit.playonyt("https://youtu.be/UTPFrdVJ2Ho?si=JheD_XsFLBB2YUNp")

        elif "hit the funny" in command or "funny one" in command:
            talk("Playing the song")
            pywhatkit.playonyt("https://youtu.be/Jyeracn7S9I?si=fSggMt9uN83HcqXN")

        elif "hit the hindi" in command or "hindi song" in command:
            talk("Playing the song")
            pywhatkit.playonyt("https://youtu.be/eK5gPcFjQps?si=HXpSxqGU7sA5Hsk4")

        elif "hit the phonk" in command:
            talk("Playing the song")
            pywhatkit.playonyt("https://youtu.be/U539QFg4QFg?si=njJcBDWI8GekPXc1")

        elif "time" in command:
            if "time in" in command:
                location = command.replace("time in", "").strip()
                get_time_in_location(location)
            else:
                get_time_in_location("Bangladesh")

        elif "screenshot" in command or "take screenshot" in command:
            take_screenshot()

        elif "new tab" in command or "another code" in command:
            open_vs_code_new_tab()

        elif "date" in command:
            current_date = date.today()
            talk(str(current_date))

        elif "close" in command or "terminate" in command:
            app = command.replace("close", "").replace("terminate", "").strip()
            close_application(app)

        elif "who" in command:
            anything = command.replace("how", "").replace("who", "").replace("what", "")
            try:
                info = wikipedia.summary(anything, sentences=3)
                talk(info)
            except Exception as e:
                talk("Sorry, I couldn't fetch that information.")
        
        elif "hello" in command or "hi" in command or "how are you" in command:
            talk("I am fine! How can I help you sir?")
        
        elif "explain" in command:
            topic = command.replace("explain", "").strip()
            try:
                info = wikipedia.summary(topic, sentences=5)
                talk(info)
            except Exception as e:
                talk("Sorry, I couldn't fetch that information.")
        
        elif "increase volume" in command or "volume up" in command or "increase the volume" in command:
            change_volume(increase=True)
        
        elif "decrease volume" in command or "volume down" in command:
            change_volume(increase=False)
        
        elif "google" in command or "search" in command:
            quest = command.replace("google", "").replace("search", "").strip()
            talk(f"Searching {quest}")
            url = f"https://www.google.com/search?q={quest}"
            webbrowser.open(url)
        # pre recommended application u can add more or remove them 
        # just add any application path 
        elif "open voicemod" in command or "run voice" in command or "voicemod" in command:
            talk("Opening Voicemod")
            subprocess.Popen(r"Enter path")
        
        elif "open file explorer" in command or "open explorer" in command or "open this pc" in command:
            talk("Opening File Explorer")
            os.startfile("explorer")
        
        elif "telegram" in command:
            talk("Opening Telegram")
            subprocess.Popen(r"Enter path")
        
        elif "chrome" in command:
            talk("Opening Chrome")
            subprocess.Popen(r"Enter path")
        
        elif "word" in command or "wordpad" in command:
            talk("Opening WordPad")
            os.startfile("write")
        # pre recommended settings command add more or remove them
        elif "open settings" in command or "pc settings" in command:
            talk("Opening Windows Settings")
            subprocess.Popen("start ms-settings:", shell=True)
        
        elif "dp settings" in command:
            talk("Opening Display Settings")
            subprocess.Popen("start ms-settings:display", shell=True)
        
        elif "cmd" in command:
            talk("Opening CMD")
            subprocess.Popen("cmd", shell=True)
        
        elif "open cap" in command or "run cap" in command or "launch cap" in command:
            talk("Opening Capcut")
            subprocess.Popen(r"Enter path")
        
        elif "store" in command or "microsoft store" in command or "ms store" in command:
            talk("Opening Microsoft Store")
            subprocess.Popen("start ms-windows-store:", shell=True)
        
        elif "open steam" in command or "steam" in command or "run steam" in command:
            talk("Opening Steam")
            subprocess.Popen(r"Enter path")
        
        elif "calculator" in command:
            talk("Opening Calculator")
            subprocess.Popen("calc", shell=True)
        
        elif command == "help":
         show_help()
         continue

        else:
            # use GPT for a response
            try:
                response = chat_with_gpt(command)
                talk(response)
            except:
                talk("I don't understand, please say that again")

if __name__ == "__main__":
    run_assistant()

# i have tried my best with my narrow knowledge give it a try 
# also if u have any suggestion to make it more advanced or more userfriendly please modify it
# don't forget to share your thought