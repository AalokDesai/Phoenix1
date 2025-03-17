import tkinter as tk
import webbrowser
import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser as wb
import os
import subprocess
import winshell
import pyautogui
import pywhatkit
import time
from GoogleNews import GoogleNews
import pandas as pd
import imdb
from threading import Thread

# Initialize text-to-speech engine
engine = pyttsx3.init()

def recognize_speech(status_text):
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)
        status_text.set("Listening...")
        audio = recognizer.listen(source)

        try:
            status_text.set("Recognizing...")
            query = recognizer.recognize_google(audio).lower()
            status_text.set(f"You said: {query}")
            return query
        except sr.UnknownValueError:
            status_text.set("Sorry, I couldn't understand your voice. Please try again.")
        except sr.RequestError:
            status_text.set("Unable to access Google Speech Recognition service. Please try again.")

def speak(audio):
    engine.say(audio)
    engine.runAndWait()
    update_conversation(f"Phoenix: {audio}")

def update_conversation(text):
    conversation_text.config(state=tk.NORMAL)
    conversation_text.insert(tk.END, f"{text}\n\n")
    conversation_text.config(state=tk.DISABLED)
    conversation_text.see(tk.END)


def movie():
    moviesdb = imdb.IMDb()
    speak("please tell me the movie name sir")
    text = recognize_speech(status_text)
    movies = moviesdb.search_movie(text)
    speak("Searching for" + text)
    if len(movies) ==0:
        speak("No result found")
    else:
        speak("I found these: ")
        for movie in movies:
            title = movie["title"]
            year = movie["year"]
            speak(f"{title}-{year}")
            info = movie.getID()
            movie = moviesdb.get_movie(info)
            rating = movie["rating"]
            plot = movie["plot outline"]
            if year<int(datetime.datetime.now().strftime("%Y")):
                speak(f"{title} was released in {year} has IMDB rating of {rating}. The plot summary of movie is {plot}")
                print(f"{title} was released in {year} has IMDB rating of {rating}. The plot summary of movie is {plot}")
                break
            else:
                speak(f"{title} will release in {year} has IMDB rating of {rating}. The plot summary of movie is {plot}")
                print(f"{title} was released in {year} has IMDB rating of {rating}. The plot summary of movie is {plot}")
                break

def OpenApp():
    speak("which app would you like to open?")
    print("which app would you like to open?")
    command = recognize_speech(status_text).lower()
    if ('calculator' in command):
        speak('Opening calculator')
        os.startfile('C:\\Windows\\System32\\calc.exe')
    elif ('paint' in command):
        speak('Opening msPaint')
        pyautogui.hotkey("win", "r")
        pyautogui.typewrite("mspaint")
        pyautogui.hotkey("enter")
    elif ('notepad' in command):
        speak('Opening notepad')
        os.startfile('C:\\Windows\\System32\\notepad.exe')
    elif ('editor' in command):
        speak('Opening your Visual studio code')
        os.startfile('C:\\Users\\Malay\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Visual Studio Code.exe')

    else:
        speak("No results found")

def CloseApp():
    speak("which app would you like to close?")
    print("which app would you like to close?")
    command = recognize_speech(status_text).lower()
    if ('calculator' in command):
        calculator_window = pyautogui.getWindowsWithTitle("Calculator")[0]

        pyautogui.hotkey("alt", "f4")
        speak("Calculator closed")
    elif ('paint' in command):
        paint_window = pyautogui.getWindowsWithTitle("paint")[0]

        pyautogui.hotkey("alt", "f4")
        speak("okay boss closing paint")
    elif ('notepad' in command):
        note_window = pyautogui.getWindowsWithTitle("notepad")[0]

        pyautogui.hotkey("alt", "f4")
        speak("okay boss closing notepad")
    elif ('editor' in command):
        editor_window = pyautogui.getWindowsWithTitle("editor")[0]

        pyautogui.hotkey("alt", "f4")
        speak("okay boss closing editor")
    
    else:
        speak("No results found")

def search_youtube(query):
    search_url = f"https://www.youtube.com/results?search_query={query}"
    webbrowser.open(search_url)
    update_conversation(f"Searching for '{query}' on YouTube...")


def process_command(command):
    if "time" in command:
        now = datetime.datetime.now().strftime("%I:%M %p")
        speak(f"The current time is {now}")
    elif "date" in command:
        now = datetime.datetime.now().strftime("%A, %B %d, %Y")
        speak(f"Today is {now}")
    elif "who are you" in command or "hu r u" in command:
        speak("I'm phoenix, a desktop voice assistant created by Malay and Aalok.")
    elif "how are you" in command or "how r u" in command:
        speak("I'm doing great! What about you?")
    elif "fine" in command or "good" in command:
        speak("Glad to hear that!")
    elif "movies" in command:
        movie()
    elif "wikipedia" in command:
        try:
            speak("Ok wait, I'm searching...")
            command = command.replace("wikipedia", "")
            result = wikipedia.summary(command, sentences=2)
            speak(result)
        except Exception as e:
            speak("Can't find this page, please ask something else")
    elif "open youtube" in command:
        wb.open("https://www.youtube.com")
        update_conversation("Opening YouTube...")
    elif "whatsapp" in command:
        send_whatsapp_message()
    elif "open google" in command:
        wb.open("https://www.google.com")
        update_conversation("Opening Google...")
    elif "open stack overflow" in command:
        wb.open("https://stackoverflow.com")
        update_conversation("Opening Stack Overflow...")
    elif "minimise the window" in command:
        pyautogui.hotkey('win', 'd')
        update_conversation("Window minimized successfully")
    elif "volume up" in command:
        pyautogui.hotkey('volumeup')
        speak("volume up succesfull")
    elif "volume down" in command:
        pyautogui.hotkey('volumedown')
        speak("Volume down successfully")
    elif "shutdown" in command or "turn off" in command:
        subprocess.call(["shutdown", "/s"])
        speak("System is on its way to shutdown")
    elif "restart" in command:
        subprocess.call(["shutdown", "/r"])
        speak("Restarting system...")
    elif "switch window" in command:
        pyautogui.keyDown("alt")
        pyautogui.press("tab")
        pyautogui.keyUp("alt")
        update_conversation("Switching window...")
    elif "news" in command:
        news()
    elif "empty recycle bin" in command:
        winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
        update_conversation("Recycle Bin emptied successfully")

    elif "open notepad" in command:
        subprocess.Popen(["notepad.exe"])
        speak("Opening Notepad...")
    elif "open calculator" in command:
        subprocess.Popen(["calc.exe"])
        update_conversation("Opening Calculator...")
    elif "minimise the window" in command:
        pyautogui.hotkey('win', 'd')
        speak("window minimise successfully")

    elif "open app" in command:
        OpenApp()
    elif "close app" in command:
        CloseApp()

    elif "where is" in command:
        location = command.replace("where is", "").strip()
        speak("Location...")
        speak(location)
        webbrowser.open(f"https://www.google.com/maps/place/{location}")


    elif "remember that" in command:
        speak("What should I remember")
        data = recognize_speech(status_text)
        speak("You said me to remember that" + data)
        print("You said me to remember that " + str(data))
        remember = open("data.txt", "w")
        remember.write(data)
        remember.close()

    elif "do you remember anything" in command:
        remember = open("data.txt", "r")
        speak("You told me to remember that " + remember.read())
        print("You told me to remember that " + str(remember))

    elif "click my photo" in command:
        pyautogui.press("super")
        pyautogui.typewrite("camera")
        pyautogui.press("enter")
        pyautogui.sleep(2)
        speak("SMILE")
        pyautogui.press("enter")
        time.sleep(2)
        pyautogui.hotkey('alt', 'f4')


    elif 'open gmail' in command:
        speak('opening your google gmail')
        webbrowser.open('https://mail.google.com/mail/')

    elif 'open maps' in command:
        speak('opening google maps')
        webbrowser.open('https://www.google.co.in/maps/')

    elif 'oprn flipkart' in command:
        speak('Opening flipkart online shopping website')
        webbrowser.open("https://www.flipkart.com/")

    elif 'open amazon' in command:
        speak('Opening amazon online shopping website')
        webbrowser.open("https://www.amazon.in/")

    elif 'close gmail' in command:
        gmail_window = pyautogui.getWindowsWithTitle("gmail")[0]
        gmail_window.activate()
        pyautogui.hotkey("ctrl", "w")
        speak("gmail closed")

    elif 'close maps' in command:
        maps_window = pyautogui.getWindowsWithTitle("maps")[0]
        maps_window.activate()
        pyautogui.hotkey("ctrl", "w")
        speak("maps closed")

    elif 'close flipkart' in command:
        flipkart_window = pyautogui.getWindowsWithTitle("flipkart")[0]
        flipkart_window.activate()
        pyautogui.hotkey("ctrl", "w")
        speak("flipkart closed")

    elif 'close amazon' in command:
        amazon_window = pyautogui.getWindowsWithTitle("amazon")[0]
        amazon_window.activate()
        pyautogui.hotkey("ctrl", "w")
        speak("amazon closed")

    elif "search" in command and "on youtube" in command:
        # Extract the query from the command
        start_index = command.find("search") + len("search")
        end_index = command.find("on youtube")
        query = command[start_index:end_index].strip()
        if query:
            search_youtube(query)
        else:
            speak("Please specify something to search for.")
    elif "play song" in command or "play music" in command:
        yt()

    elif "screenshot" in command or "take screenshot" in command:
        save_folder = r"C:\Users\Malay\Desktop\phoenix\screenshot"
        take_screenshot(save_folder)
        take_screenshot(save_folder)
    elif "quit" in command:
        quit()

    else:
        update_conversation("Sorry, I didn't get that.")

def take_screenshot(save_folder):
    try:
        # Create the folder if it doesn't exist
        if not os.path.exists(save_folder):
            os.makedirs(save_folder)

        # Capture a screenshot using pyautogui
        screenshot = pyautogui.screenshot()

        # Specify the file path to save the screenshot
        screenshot_path = os.path.join(save_folder, "screenshot.png")

        # Save the screenshot to the specified path
        screenshot.save(screenshot_path)

        print(f"Screenshot saved at: {screenshot_path}")
    except Exception as e:
        print(f"Error occurred while taking screenshot: {e}")

def send_whatsapp_message():
    # Open WhatsApp Web in a new browser tab
    webbrowser.open('https://web.whatsapp.com/', new=2)
    speak("Please log in to WhatsApp Web and wait for a few seconds...")
    time.sleep(15)

    while True:
        # Click on the search box to start a new chat
        pyautogui.click(250, 218)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        time.sleep(1)

        # Ask for recipient's name using voice input
        speak("Please say the recipient's name or say 'quit' to cancel.")
        recipient_name = recognize_speech(status_text)

        if recipient_name == 'quit':
            speak("Message sending canceled.")
            pyautogui.hotkey("alt", "w")  # Close the WhatsApp Web tab
            return  # Exit the function if message sending is canceled

        # Type the recipient's name
        pyautogui.typewrite(recipient_name)
        time.sleep(1)

        # Click on the topmost search result
        pyautogui.click(295, 371)
        time.sleep(1)

        # Confirm recipient name with the user
        speak(f"Are you sure you want to message '{recipient_name}'? Say 'yes' or 'no'.")
        confirmation = recognize_speech(status_text)

        if confirmation == 'yes':
            # Proceed to message sending
            speak("Please say your message.")
            message = recognize_speech(status_text)

            if message:
                # Click on the message box and type the message
                pyautogui.click(835, 956)
                pyautogui.typewrite(message)
                pyautogui.press('enter')  # Press Enter after typing the message

                speak("Message sent successfully.")
                print("Message sent successfully.")
                time.sleep(5)
                pyautogui.hotkey('alt', 'f4')
                break  # Exit the loop after sending the message
            else:
                speak("No message detected. Please try again.")
        elif confirmation == 'no':
            speak("Please say the recipient's name again.")
        else:
            speak("Invalid response. Please try again.")


def yt():
    speak("Boss, can you please say the name of the song?")
    print("Boss, can you please say the name of the song?")
    song = recognize_speech(status_text)
    song = song.replace("play", "").strip()

    if song:
        speak(f"Playing {song} on YouTube")
        print(f"Playing {song} on YouTube")
        pywhatkit.playonyt(song)
    else:
        speak("Sorry, I couldn't understand the song name.")

def news():
    news = GoogleNews(period="id")
    news.search("India")
    result = news.result()
    data = pd.DataFrame.from_dict(result)
    data = data.drop(columns=["img"])
    data.head()

    for i in result:
        print(i["title"])
        speak(i["title"])


def start_assistant():
    speak("Welcome back sir! Phoenix at your service , please tell me how may I help you.")

    listening = True
    while listening:
        command = recognize_speech(status_text)
        if command:
            if "stop" in command or "exit" in command:
                speak("Goodbye!")
                listening = False
            else:
                process_command(command)

def start_listening():
    global listening
    listening = True
    Thread(target=start_assistant).start()

def exit_application():
    root.destroy()

# Create the main window using tkinter
root = tk.Tk()
root.title("Phoenix: An Voice based Desktop Application")
root.geometry("800x800")
root.configure(bg="lightblue")

status_text = tk.StringVar()
status_text.set("")

status_label = tk.Label(root, textvariable=status_text, fg="blue")
status_label.pack(pady=10)


# Create a frame for conversation display
conversation_frame = tk.Frame(root, bg="white")
conversation_frame.pack(pady=20, padx=10, fill=tk.BOTH, expand=True)

# Create a text widget to display the conversation
conversation_text = tk.Text(conversation_frame, bg="white", wrap=tk.WORD, font=("Helvetica", 12), height=15, width=50)
conversation_text.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
conversation_text.config(state=tk.DISABLED)  # Set initial state to disabled

# Create a button to start the assistant
start_button = tk.Button(root, text="Start Assistant", command=start_listening, bg="blue", fg="white",
                         font=("Helvetica", 12))
start_button.pack(pady=20)

exit_button = tk.Button(root, text="Exit", command=exit_application, bg="red", fg="white", font=("Helvetica", 12))
exit_button.pack(pady=10)

# Function to run the tkinter main loop
def run_gui():
    root.mainloop()

# Run the GUI application
run_gui()
