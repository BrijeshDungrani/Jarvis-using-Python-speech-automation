import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import pyautogui
from sys import platform
import sys
import psutil
from youtube import youtube
# from diction import translate
from news import speak_news, getNewsUrl
from loc import weather
## for powerpoint
import win32com.client
import time
# for cpu and battery functionality
import psutil

# for jokes
import pyjokes

#### variable  ####
VScodePath = "C:\\Users\\Brijesh\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
music_dir = 'E:\\songs\\songs\\bR!je$#  favorite'





engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice',voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def takeCommand():
    query = input("Please enter something: ")
    # r = sr.Recognizer()
    # with sr.Microphone() as source:
    #     print("Listening")
    #     #r.phrase_threshold = 1
    #    # var = input("Please enter something: ")
    #     audio = r.listen(source)
      
    # try:
    #     print("Recognizing......")
    #     query = r.recognize_google(audio,language='en-in')
    #     print(f"User Said : {query}\n")

    # except Exception as e:
    #     #print(e)
    #     print("say that again Please")
    #     return "None"
    return query



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
    speak('taking screenshot')
    img = pyautogui.screenshot()
    img.save('C:\\Users\\Brijesh\\Downloads\\screenshot.png')

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

def chromeAsDefault():
    if platform == "linux" or platform == "linux2":
        chrome_path = '/usr/bin/google-chrome'

    elif platform == "darwin":
        chrome_path = 'open -a /Applications/Google\ Chrome.app'

    elif platform == "win32":
        chrome_path = 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
    else:
        print('Unsupported OS')
        exit(1)

    webbrowser.register(
        'chrome', None, webbrowser.BackgroundBrowser(chrome_path))





if __name__ == "__main__":
    chromeAsDefault()
    wishMe()
    weather()
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
            speak("Opening youtube")
            webbrowser.get('chrome').open_new_tab('https://youtube.com')
        elif 'open google' in query:
            speak("Opening Google")
            webbrowser.get('chrome').open_new_tab('https://google.com')
        elif 'open stackoverflow' in query:
            speak("Opening Stackoverflow")
            webbrowser.get('chrome').open_new_tab('https://stackoverflow.com')
        elif 'open moodle' in query:
            speak("Opening Moodle")
            webbrowser.get('chrome').open_new_tab('https://moodle.frankfurt-university.de/')
        elif 'german to english' in query:
            speak("sure")
            webbrowser.get('chrome').open_new_tab('https://translate.google.de/?hl=en&ui=tob&sl=de&tl=en&op=translate')
        elif 'english to german' in query :
            speak("sure")
            webbrowser.get('chrome').open_new_tab('https://translate.google.de/?hl=en&ui=tob&sl=en&tl=de&op=translate')
            
        elif 'play music' in query:
            speak("Playing music for you")
           
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
        elif 'shutdown' in query:
            if platform == "win32":
                os.system('shutdown /p /f')
            elif platform == "linux" or platform == "linux2" or "darwin":
                os.system('poweroff')

        elif 'sleep' in query:
            sys.exit()

        elif 'search youtube' in query:
            speak('What you want to search on Youtube?')
            youtube(takeCommand())
        

        elif 'search' in query:
            speak('What do you want to search for?')
            search = takeCommand()
            url = 'https://google.com/search?q=' + search
            webbrowser.get('chrome').open_new_tab(
                url)
            speak('Here is What I found for' + search)

        elif 'location' in query:
            speak('What is the location?')
            location = takeCommand()
            url = 'https://google.nl/maps/place/' + location + '/&amp;'
            webbrowser.get('chrome').open_new_tab(url)
            speak('Here is the location ' + location)

        # elif 'dictionary' in query:
        #     speak('What you want to search in your intelligent dictionary?')
        #     translate(takeCommand())

        elif 'news' in query:
            speak('Ofcourse sir..')
            speak_news()
            speak('Do you want to read the full news...')
            test = takeCommand()
            if 'yes' in test:
                speak('Ok Sir, Opening browser...')
                webbrowser.open(getNewsUrl())
                speak('You can now read the full news from this website.')
            else:
                speak('No Problem Sir')