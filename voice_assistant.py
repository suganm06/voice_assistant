from logging import exception
import subprocess
import googletrans
import pyttsx3
import speech_recognition as sr
import datetime
import webbrowser
import os
import time
import winshell
import pyaudio
from googletrans import Translator
from urllib.request import urlopen
from gtts import gTTS
from termcolor import colored

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        say("Good Morning boss !")
    elif hour >= 12 and hour < 18:
        say("Good Afternoon boss !")
    else:
        say("Good Evening boss !")
    say("I am Jarvis")
    say("your Assistant")
    say('What do you want me to do')

    
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.adjust_for_ambient_noise(source,duration=1)
        audio = r.listen(source)
        try:
            print("Recognizing...")
            call = r.recognize_google(audio, language='en-in')
            print(f"User said: {call}\n")
        except Exception as e:
            print(e)
            print(colored("Unable to Recognize your voice!!","red"))
            return "None"
        return call

def say(audio):
    engine.say(audio)
    engine.runAndWait()

def get_key(val):
    for key, value in googletrans.LANGUAGES.items():
        if val == value:
            return key

    return

if __name__ == '__main__':
    clear = lambda: os.system('cls')
    wishMe()

while True:
    call= takeCommand().lower()

    if 'time' in call:
        strTime=datetime.datetime.now().strftime("%H:%M:%S")
        say(f"the time is {strTime} boss")
            
    elif 'shutdown' in call:
        say("Your system is going to shut down, good bye boss")
        subprocess.call(["shutdown","-f","-s","-t","1"])
                 
    elif 'empty recycle bin' in call:
        winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
        say("emptied recycle bin boss")
                
    elif "restart" in call:
        say("restarting your system boss")
        subprocess.call(["shutdown","-f","-r","-t","1"]) 
        
    elif 'translate' in call:
            print('Say the source language')
            say('Say the source language')
            from_language = takeCommand().lower()
            print('Say the destination language')
            say('Say the destination language')
            to_language = takeCommand().lower()
            say('Say the sentence')
            translator = Translator()
            try:
                get_sentence = takeCommand()
                print("Phrase to be translated: " + get_sentence)
                text_to_translate = translator.translate(text=get_sentence, src=get_key(from_language),dest=get_key(to_language))
                text = text_to_translate.text
                print(text)
                speak = gTTS(text=text, lang=get_key(to_language), slow=False)
                speak.save("captured_voice.mp3")
                os.system("start captured_voice.mp3")
            except sr.UnknownValueError:
                print("Unable to recognize input")
            except sr.RequestError as e:
                print("unable to provide output".format(e))
            continue
        
    elif 'open google' in call:
        say('What do you want me to search boss')
        lib = takeCommand()
        webbrowser.open("https://www.google.co.in/search?q=" +(str(lib))+ "&oq="+(str(lib))+"&gs_l=serp.12..0i71l8.0.0.0.6391.0.0.0.0.0.0.0.0..0.0....0...1c..64.serp..0.0.0.UiQhpfaBsuU")   
    elif "close google" in call:
        subprocess.call(["taskkill","/F","/IM","chrome.exe"])  
            
    elif 'open youtube' in call:
        say('What do you want me to search boss')
        youtube_search = takeCommand()
        webbrowser.open('https://www.youtube.com/results?search_query=' + youtube_search)  
    elif "close youtube" in call:
        subprocess.call(["taskkill","/F","/IM","chrome.exe"])  
        
    elif 'open powerpoint presentation' in call:
            say("opening Power Point presentation boss")
            power = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
            os.startfile(power)  
    elif "close powerpoint presentation" in call:
        subprocess.call(["taskkill","/F","/IM","POWERPNT.EXE"])  
                
    elif  'open microsoft word' in call:
            say("opening microsoft word boss")
            word = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
            os.startfile(word)               
    elif "close microsoft word" in call:
        subprocess.call(["taskkill","/F","/IM","WINWORD.EXE"])  
        
    elif 'open excel' in call:
            say("opening excel boss")
            excel=r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
            os.startfile(excel)
    elif "close excel" in call:
        subprocess.call(["taskkill","/F","/IM","EXCEL.EXE"]) 
        
    elif 'open ms teams' in call:
            say("opening ms teams boss")
            teams=r"C:\Users\123.DESKTOP-74G99SN\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams"
            os.startfile(teams)
    elif "close ms teams" in call:
        subprocess.call(["taskkill","/F","/IM","Teams.exe"]) 
        
    elif "open vs code" in call:
        say("opening vs code boss")
        vs_code = r"C:\Users\123.DESKTOP-74G99SN\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code\Visual Studio Code"   
        os.startfile(vs_code)  
    elif "close vs code" in call:
        subprocess.call(["taskkill","/F","/IM","Code.exe"]) 
        
    elif "open discord" in call:
        say("opening discord boss")
        discord= r"C:\Users\123.DESKTOP-74G99SN\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Discord Inc\Discord"   
        os.startfile(discord)
    elif "close discord" in call:
        subprocess.call(["taskkill","/F","/IM","Discord.exe"]) 
            
    elif "open files" in call:
        say("opening files boss")
        files=r"C:\Windows\explorer.exe"
        os.startfile(files)  
        
    elif "don't listen" in call or "stop listening" in call:
            say("for how much time you want to stop jarvis from listening commands")
            say("for 20 seconds")
            say("or for 30 seconds")
            a=takeCommand()
            if "20" in a:
                say("not listening for 20 seconds")
                time.sleep(20)
                say("not listened for 20 seconds.")
            elif "30" in a:
                say("not listening for 30 seconds")
                time.sleep(30)             
                say("not listened for 30 seconds")
                       
    elif "goodbye" in call or "exit" in call:
        say("good bye boss")
        break    
            