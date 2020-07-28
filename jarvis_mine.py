import os
import speech_recognition as sr 
import pyttsx3
import datetime
import gtts as gtts
import webbrowser
import smtplib
import psutil
import requests

engine = pyttsx3.init() 
engine = pyttsx3.init('sapi5') 
engine.setProperty('rate', 120) 
voices = engine.getProperty('voices')
engine.setProperty('voice','voices[4].id')
for voice in voices:
    print("Voice:")
    print(" - ID: %s" % voice.id)
    print(" - Name: %s" % voice.name)
    print(" - Languages: %s" % voice.languages)
    print(" - Gender: %s" % voice.gender)
    print(" - Age: %s" % voice.age)
# engine.say(',  ,  jay swaminarayan ,   Hello sir,i am your assistant , JARVIS ,  , here i am for your virtually help  , how may I help you sir ??') 

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
           print("Listening...")
           r.pause_threshold = 1
           audio = r.listen(source)
    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)    
        print("Say that again please...")  
        return "None"
    return query       
     
def speech_BY_machine(audio):
    engine.say(audio)
    engine.runAndWait()

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('your email', 'your_password')
    server.sendmail('email', to, content)
    server.close()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        # speak("Good Morning!")
        speech_BY_machine("Good Morning!")

    elif hour>=12 and hour<18:
        speech_BY_machine("Good Afternoon!")   

    else:
        speech_BY_machine("Good Evening!")  
    speech_BY_machine(f"Sir, i am jarvis your personal artificial intlegent assistant,  Please tell me how may I help you")  
    # engine.say("I am Jarvis Sir. Please tell me how may I help you")   

def NewsFromBBC(): 
      
    # BBC news api 
    speech_BY_machine("here top news for you")
    rl = " https://newsapi.org/v1/articles?source=bbc-news&sortBy=top&apiKey=4dbc17e007ab436fb66416009dfb59a8"
  
    # fetching data in json format 
    bbc = requests.get(url).json() 
  
    # getting all articles in a string article 
    article = bbc["articles"] 
  
    # empty list which will  
    # contain all trending news 
    results = [] 
      
    for ar in article: 
        results.append(ar["title"]) 
          
    for i in range(len(results)): 
          
        # printing all trending news 
        print(i + 1, results[i]) 
  
    #to read the news out loud for us 
    from win32com.client import Dispatch 
    speak = Dispatch("SAPI.Spvoice") 
    speak.Speak(results)                  
  

if __name__ == "__main__":
    wishMe()
    
    while(1):

        query = takeCommand().lower()
           
        if 'play music' in query:
             dir = 'D:\\Aaaa\\src'
             songs = os.listdir(music_dir)
             print(songs)    
             os.startfile(os.path.join(dir, songs[0]))

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'what is time now' in query:  
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speech_BY_machine(f"Sir, the time is {strTime}")       

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'send email' in query:
            try:
                speech_BY_machine("What should I say?")
                content = takeCommand()
                to = "email"    
                sendEmail(to, content)
                speech_BY_machine("Email has been sent!")
            except Exception as e:
                print(e)
                speech_BY_machine("sorry SIR !. I am not able to send this email") 
        elif 'what is your operating system' in query:
            speech_BY_machine(os.name)
            print(os.name)

        elif "shutdown system" in query:
            os.system("shutdown /s /t 1"); 

        elif "show me battery percentage" in query:
            battery = psutil.sensors_battery()
            plugged = battery.power_plugged
            percent = str(battery.percent)
            if plugged==False: plugged="Not Plugged In"
            else: plugged="Plugged In"
            speech_BY_machine(percent)
            speech_BY_machine(plugged)
            print(percent+'% | '+plugged)    

        elif "tell me news" in query:
              NewsFromBBC() 

        elif 'stop jarvis' in query: 
            speak("Thanks for giving me your time") 
            exit()       
                                
        else:
            speech_BY_machine("i cant understand what you are saying , it is not programed yet ")

