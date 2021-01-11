from index import *
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