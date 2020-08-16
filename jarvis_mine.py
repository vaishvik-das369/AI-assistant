import os
import speech_recognition as sr 
import pyttsx3
import datetime
import gtts as gtts
import webbrowser
import smtplib
import psutil
import requests
import numpy as np
import pandas as pd
import matplotlib as plt
from pywinauto import application
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
        speech_BY_machine("Say that again please...")  
        return "None"
    return query       
     
def speech_BY_machine(audio):
    engine.say(audio)
    engine.runAndWait()

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('iamvaishvik@gmail.com', 'python@ml@ai')
    server.sendmail('vaishvikpatel001@gmail.com', to, content)
    server.close()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        # speak("Good Morning!")
        speech_BY_machine("Good Morning!, jay swaminarayan")

    elif hour>=12 and hour<18:
        speech_BY_machine("Good Afternoon!, jay swaminarayan")   

    else:
        speech_BY_machine("Good Evening!, jay swaminarayan")  
    speech_BY_machine(f"   hello parth Sir,    i am jarvis your personal artificial intelligence assistant,  Please tell me how may I help you")  
    engine.say("I am Jarvis Sir. Please tell me how may I help you")   

def NewsFromBBC(): 
      
    speech_BY_machine("here top news for you")
    main_url = " https://newsapi.org/v1/articles?source=bbc-news&sortBy=top&apiKey=4dbc17e007ab436fb66416009dfb59a8"
    open_bbc_page = requests.get(main_url).json() 
    article = open_bbc_page["articles"] 
    results = [] 
      
    for ar in article: 
        results.append(ar["title"]) 
          
    for i in range(len(results)): 
        print(i + 1, results[i]) 
    from win32com.client import Dispatch 
    speak = Dispatch("SAPI.Spvoice") 
    speak.Speak(results)                  
  
if __name__ == "__main__":
    wishMe()
    
    while(1):

        query = takeCommand().lower()

        if ('goodbye') in query:                          
            rand = ['Goodbye Sir', 'Jarvis powering off in 3, 2, 1, 0']
            speech_BY_machine(rand)
            break       

        elif 'play music' in query:
             music_dir = 'D:\\Aaaa\\src'
             songs = os.listdir(music_dir)
             print(songs)    
             os.startfile(os.path.join(music_dir, songs[0]))

        elif 'open google and search' in query:
            from selenium import webdriver 
            
            from selenium import webdriver 
            
            # create webdriver object 
            driver = webdriver.Firefox()
            speech_BY_machine("what do you looking for , sir ?") 
            google = takeCommand()
            keyword = google
            
          
            driver.get("https://google.co.in / search?q ="+keyword) 
                
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
                to = "iamvaishvik@gmail.com"    
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
            speech_BY_machine("Thanks for giving me your time") 
            exit()      

        elif 'where we are' in query:
                query = query.split(" ")
                location = query[2]
                speech_BY_machine("Hold on vaishvik sir, I will show you where " + location + " is.")
                os.system("chromium-browser https://www.google.nl/maps/place/" + location + "/&amp;")   

        elif 'find regressor' in query:
            speech_BY_machine('i am also perform machine learning , sir')
            dataset = pd.read_csv('Salary_Data.csv')
            x = dataset.iloc[:,0:-1].values
            speech_BY_machine(x)
            y = dataset.iloc[:,-1].values
            speech_BY_machine(y)

        elif 'open my facebook account' in query:
             webbrowser.open("https://www.facebook.com/aksharbhram.das")

        elif 'i will call you' in query:
            speech_BY_machine("ok , sir")
            
            def call_you_later():
                hi = takeCommand()
            
                if 'hello jarvis' in hi:
                    return hi;

        elif 'open notepad' in query:
            app =  application.Application()
            app.start("Notepad.exe")            

        elif 'birthday' in query:
            dataset = pd.read_csv('Yuva database.csv')
            # print(dataset)
            from datetime import date 
            today = datetime.datetime.now().strftime("%d-%m")
            print("Today date is: ", today) 
            bday = dataset['Birth Date']

            for item in dataset.iterrows():
                bday = item['Birth Date'].strftime("%d-%m")
                print(bday)
                if(today == bday):
                    print(item['Name'])
        else:
             speech_BY_machine("i cant understand what you are saying , it is not programed yet ")
             
