import speech_recognition as sr
import datetime
import wikipedia
import pyaudio
import time
import webbrowser
import pyautogui # pip install pyautogui
from PIL import Image, ImageGrab # pip install pillow

hour = int(datetime.datetime.now().hour)
def speak(str):
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)

def screenShot():
    image = ImageGrab.grab()      
    image.show()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour <= 12:
        speak("Good Morning Sir")
        print("Good Morning Sir")
    elif hour >= 12 and hour <= 18:
        speak("Good Afternoon Sir")
        print("Good Afternoon Sir")
    else:
        speak("Good Evevning Sir")
        print("Good Evevning Sir")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening....")
        r.pause_threshold = 1
        audio = r.listen(source)
        

    try:
        print("Recognizing....")
        query = r.recognize_google(audio, language="en-us")
        print(f"User said : {query}\n")
    
    except Exception as e:
        print("Please say that again....")
        speak("Sorry, Please say that again....")
        return "None"
    return query   

if __name__ == "__main__":
    speak("Starting Jarvis ")
    wishMe()
    while True:

        query = takeCommand().lower()
         
        if 'wikipedia' in query:
            speak("Searching wikipedia")
            print("Searching wikipedia")
            query = query.replace("wikipedia", "")            
            results = wikipedia.summary(query, sentences=2)
            speak("According to wikipedia")
            print("According to wikipedia")
            print(results)
            speak(results)

        elif 'nikita' in query:
            print("Hello Ma'am, If I'm not wrong then you're sir's sister Nikita ?")
            speak("Hello Ma'am, If I'm not wrong then you're sir's sister Nikita ?")                    
        elif 'google' in query:
            speak('opening google')
            webbrowser.open('google.com')
        elif 'youtube' in query:
            speak('opening youtube')
            webbrowser.open('youtube.com')
        elif 'whatsapp' in query:
            speak('opening whatsapp')
            webbrowser.open('whatsapp.com')
        elif 'jarvis' in query:
            speak("Yes Sir , How may I help you ?")
            print("Yes Sir , How may I help you ?")
        elif 'how are you ?' in query:
            speak("I am fine , tell me how are you ? ")
            print("I am fine , tell me how are you ? ")
        elif 'I am fine ' in query:
            speak("Oh ! I just wanted to hear that ")
            print("Oh ! I just wanted to hear that ")
        elif 'good'in query:
            speak("Thanks bhai , meri taarif karne ke liye tum bhi bahut achhe ho")
        elif 'screenshot' in query:
            speak("Taking screenshot of the current window")
            screenShot()
            speak("Taken the screenshot and now opening it")
        elif 'bye' in query:
            speak("Bye sir , have a nice day,")
            print("Bye Sir, have a nice day,")
            if hour > 19:
                print("Bye sir and have a nice sleep")
                speak("Bye sir and have a nice sleep")
            break

