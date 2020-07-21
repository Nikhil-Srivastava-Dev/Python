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


def wishMeSir():
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
        
def wishMeMaam():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour <= 12:
        speak("Good Morning Ma'am")
        print("Good Morning Ma'am")
    elif hour >= 12 and hour <= 18:
        speak("Good Afternoon Ma'am")
        print("Good Afternoon Ma'am")
    else:
        speak("Good Evevning Ma'am")
        print("Good Evevning Ma'am")       


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
    
    speak("This programme is under development enter Yes to continue No to exit")
    print("This programme is under development enter Yes to continue No to exit")
    yesno = input("Enter Yes to continue, No to exit :::  ")
    lo = yesno.lower()
    if lo == "yes":
        query = takeCommand().lower()

        speak("Please tell whether you are male or female ")
        print("Please tell whether you are male or female ")
        mf = input("Enter Male if you are male, or enter Female if you are female")
        lmf = mf.lower()
        speak("Ok thanks for telling me ")
        speak("Starting Jarvis ")
        if lmf == "male":
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
                    speak("Thankyou for praising me :")
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
        elif lmf == "female":
            wishMeMaam()
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

#                 elif 'nikita' in query:
#                     print("Hello Ma'am, If I'm not wrong then you're sir's sister Nikita ?")
#                     speak("Hello Ma'am, If I'm not wrong then you're sir's sister Nikita ?")                    
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
                    speak("Yes Ma'am , How may I help you ?")
                    print("Yes Ma'am , How may I help you ?")
                elif 'how are you ?' in query:
                    speak("I am fine , tell me how are you ? ")
                    print("I am fine , tell me how are you ? ")
                elif 'I am fine ' in query:
                    speak("Oh ! I just wanted to hear that ")
                    print("Oh ! I just wanted to hear that ")
                elif 'good'in query:
                    speak("Thankyou for praising me :)")
                elif 'screenshot' in query:
                    speak("Taking screenshot of the current window")
                    screenShot()
                    speak("Taken the screenshot and now opening it")
                elif 'bye' in query:
                    speak("Bye Ma'am , have a nice day,")
                    print("Bye Ma'am, have a nice day,")
                    if hour > 19:
                        print("Bye Ma'am and have a nice sleep")
                        speak("Bye Ma'am and have a nice sleep")
                    break
    elif lo == "no":
        speak("Exiting the programme")
        print("Invalid input !!!!!")            
#         break
    else:
        
        speak("Invalid input !!!!!")
        print("Invalid input !!!!!")            
