import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import pyautogui

## for powerpoint
import win32com.client
import time

# for cpu and battery functionality
import psutil

# for jokes
import pyjokes

#### variable  ####
VScodePath = "C:\\Users\\Brijesh\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"






engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice',voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour<12:
        speak("Good Morning ! ")
    elif hour >= 12  and hour<18:
        speak("Good Afternoon")
    else:
        speak("Good Evening")
    speak("Please Tell me How Can I help you ?")

def screenshot():
    img = pyautogui.screenshot()
    img.save('C:\\Users\\Brijesh\\Download/screenshot.png')

def cpu():
    usage = str(psutil.cpu_percent())
    speak("CPU is at"+usage)

    battery = psutil.sensors_battery()
    speak("battery is at")
    speak(battery.percent)

def joke():
    for i in range(5):
        speak(pyjokes.get_jokes()[i])


def ppt():
    app = win32com.client.Dispatch("PowerPoint.Application")
    presentation = app.Presentations.Open(FileName=u'G:\HIS\SoSe 2020\\New SCS (1).pptx', ReadOnly=1)
    codePath = "C:\\Users\\Brijesh\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
    presentation.SlideShowSettings.Run()
    while True:
        #take a voice input of inner presentation
        presentationQuery = takeCommand().lower()
        if 'next' in presentationQuery:
            presentation.SlideShowWindow.View.Next()
            presentationQuery = ""
        elif 'previous' in presentationQuery:
            presentation.SlideShowWindow.View.Previous()
            presentationQuery = ""
        elif 'stop' in presentationQuery:
            presentation.SlideShowWindow.View.Exit()
        elif 'quit' in presentationQuery:
            presentation.SlideShowWindow.View.Exit()
            app.Quit()        

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening")
        #r.phrase_threshold = 1
       # var = input("Please enter something: ")
        audio = r.listen(source)
      
    try:
        print("Recognizing......")
        query = r.recognize_google(audio,language='en-in')
        print(f"User Said : {query}\n")

    except Exception as e:
        #print(e)
        print("say that again Please")
        return "None"
    return query

if __name__ == "__main__":
    wishMe()
    while True:
        query = takeCommand().lower()

        if 'wikipedia' in query:
            speak('Searching wikipedia...')
            query = query.replace("wikipedia","")
            results = wikipedia.summary(query,sentences=2)
            speak("According to wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'play music' in query:
            music_dir = 'E:\\songs\\songs\\bR!je$#  favorite'
            songs = os.listdir(music_dir)
            #print(songs)
            os.startfile(os.path.join(music_dir,songs[0]))
            ## Can add next song, Last song

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S:")
            speak(f"The time is : {strTime}")
            
        elif 'thank you' in query:
            speak("Welcome ")

        elif "how are you" in query:
            speak("I am fine and you ?")
        
        elif 'open code' in query:
            
            os.startfile(VScodePath)

        elif 'open presentation' in query:
            ppt()
            # or command like next slide, previous slide
        elif 'screenshot' in query:
            speak("taking screenshot")
            screenshot()
        elif 'cpu' in query:
            cpu()

        elif 'joke' in query:
            joke()
        