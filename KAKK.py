import speech_recognition as sr
import pyttsx3
import datetime
import wikipedia
import webbrowser
import os
import time
import subprocess

import wolframalpha
import json
import requests
import pyaudio
import pyjokes
import feedparser
import smtplib
import ctypes

import requests
import shutil
from twilio.rest import Client
from clint.textui import progress

from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import cv2
import sys

print('Loading your AI personal assistant - Jarvis')

engine=pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
engine.setProperty('voice','voices[0].id')


def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour=datetime.datetime.now().hour
    if hour>=0 and hour<12:
        speak("Hello,Good Morning")
        print("Hello,Good Morning")
    elif hour>=12 and hour<18:
        speak("Hello,Good Afternoon")
        print("Hello,Good Afternoon")
    else:
        speak("Hello,Good Evening")
        print("Hello,Good Evening")

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()

    # Enable low security in gmail
    server.login('email', 'password')
    server.sendmail('semail', to, content)
    server.close()

def takeCommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio=r.listen(source)

        try:
            statement=r.recognize_google(audio,language='en-in')
            print(f"user said:{statement}\n")

        except Exception as e:
            speak("Pardon me, please say that again")
            return "None"
        return statement

speak("Loading your AI personal assistant Jarvis")
wishMe()


if __name__=='__main__':


    while True:
        speak("Tell me how can I help you now?")
        statement = takeCommand().lower()
        if statement==0:
            continue

        if "good bye" in statement or "ok bye" in statement:
            speak('your personal assistant Jarvis is shutting down,Good bye')
            print('your personal assistant Jarvis is shutting down,Good bye')
            break

        if 'wikipedia' in statement:
            speak('Searching Wikipedia...')
            statement =statement.replace("wikipedia", "")
            results = wikipedia.summary(statement, sentences=3)
            speak("According to Wikipedia")
            print(results)
            speak(results)
            
        elif 'all commands' in statement:
            pyautogui.alert("wikipedia\nopen youtube\njoke\nopen google\nopen gmail\nemail from satrajit\nweather in\nhow are you\nfine/good\nwrite a note\nopen note\ntime\nwho are you\nwho made you\nwhat can you do\nwho created/discovered you\nplay music/song\nopen stackoverflow\nnews\nsearch\ntake picture\nquestion\nlock windows\nshutdown system\nempty recycle bin\ndo not listen/stop listening\nwhere is\nrestart\npower off/sign out")

        elif 'open youtube' in statement:
            webbrowser.open_new_tab("https://www.youtube.com")
            speak("youtube is open now")
            time.sleep(5)


        elif 'joke' in statement:
            speak(pyjokes.get_joke())

        elif 'open google' in statement:
            webbrowser.open_new_tab("https://www.google.com")
            speak("Google chrome is open now")
            time.sleep(5)

        elif 'open gmail' in statement:
            webbrowser.open_new_tab("https://mail.google.com/mail/u/0/?pli=1#inbox")
            speak("Google Mail open now")
            time.sleep(5)

        elif 'email from satrajit' in statement:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "Receiver email address"
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif "weather in" in statement:
                listening = True
                api_key = "api key"
                weather_url = "http://api.openweathermap.org/data/2.5/weather?"
                statement = statement.split(" ")
                location = str(statement[2])
                url = weather_url + "appid=" + api_key + "&q=" + location
                js = requests.get(url).json()
                if js["cod"] != "404":
                    weather = js["main"]
                    temp = weather["temp"]
                    hum = weather["humidity"]
                    desc = js["weather"][0]["description"]
                    resp_string = " The temperature in Kelvin is " + str(temp) + " The humidity is " + str(
                        hum) + " and The weather description is " + str(desc)
                    speak(resp_string)
                else:
                    speak("City Not Found")







        elif 'how are you' in statement:
            speak("I am fine, Thank you")
            speak("How are you, Sir")
            time.sleep(2)
        elif 'fine' in statement or "good" in statement:
            speak("It's good to know that your fine")

        elif "write a note" in statement:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('jarvis.txt', 'w')

            file.write(note)

        elif "open note" in statement:
            speak("Showing Notes")
            file = open("jarvis.txt", "r")
            print(file.read())
            speak(file.read(6))


        elif 'time' in statement:
            strTime=datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"the time is {strTime}")

        elif 'who are you' in statement or 'what can you do' in statement:
            speak('I am Jarvis version 1 point O your persoanl assistant. I am programmed to minor tasks like'
                  'opening youtube,google chrome,gmail and stackoverflow ,predict time,take a photo,search wikipedia,predict weather' 
                  'in different cities , get top headline news from times of india and you can ask me computational or geographical questions too!')

        



        elif "who made you" in statement or "who created you" in statement or "who discovered you" in statement:
            speak("I was built by Satrajit")
            print("I was built by Satrajit")
        elif 'play music' in statement or "play song" in statement:
            speak("Here you go with music")
            # music_dir = "G:\\Song"
            music_dir = "C:\\Users\\satra\\Music"
            songs = os.listdir(music_dir)
            print(songs)
            random = os.startfile(os.path.join(music_dir, songs[1]))
            time.sleep(500)
        elif "open stack overflow" in statement:
            webbrowser.open_new_tab("https://stackoverflow.com/login")
            speak("Here is stackoverflow")

        elif 'news' in statement:
            news = webbrowser.open_new_tab("https://timesofindia.indiatimes.com/home/headlines")
            speak('Here are some headlines from the Times of India,Happy reading')
            time.sleep(6)


        elif 'search'  in statement:
            statement = statement.replace("search", "")
            webbrowser.open_new_tab(statement)
            time.sleep(5)

        elif 'take picture' in statement:

            cap = cv2.VideoCapture(0)  # video capture source camera (Here webcam of laptop)
            ret, frame = cap.read()  # return a single frame in variable `frame`

            while (True):
                cv2.imshow('img1', frame)  # display the captured image
                if cv2.waitKey(1) & 0xFF == ord('y'):  # save on pressing 'y'
                    cv2.imwrite('images/c1.png', frame)
                    cv2.destroyAllWindows()
                    break

            cap.release()





        elif 'question' in statement:
            speak('I can answer to computational and geographical questions and what question do you want to ask now')
            question=takeCommand()
            app_id="app id"
            client = wolframalpha.Client('client id')
            res = client.query(question)
            answer = next(res.results).text
            speak(answer)
            print(answer)
        elif 'lock windows' in statement:
            speak("Locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in statement:
            speak('Hold on a sec!Your system will shut down soon')
            subprocess.call('shutdown/p/f')

        elif 'empty recycle bin' in statement:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak('Recycle Bin has been emptied')

        elif "do not listen" in statement or "stop listening" in statement:
            speak("Ok Boss I am not listening")

            time.sleep(100)

        elif "where is" in statement:
            listening = True
            statement = statement.split(" ")
            location_url = "https://www.google.com/maps/place/" + str(statement[2])
            speak("Hold on, I will show you where " + statement[2] + " is.")
            webbrowser.open_new_tab(location_url)

        


        elif "restart" in statement:
            subprocess.call(["shutdown", "/r"])

        elif "log off" in statement or "sign out" in statement:
            speak("Ok , your pc will log off in 10 sec make sure you exit from all applications")
            subprocess.call(["shutdown", "/l"])

time.sleep(3)
