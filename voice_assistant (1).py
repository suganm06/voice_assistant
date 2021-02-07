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
from googletrans import Translator
from urllib.request import urlopen
from gtts import gTTS

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        say("Good Morning Master !")
    elif hour >= 12 and hour < 18:
        say("Good Afternoon Master !")
    else:
        say("Good Evening Master !")
    say("I am Jarvis your Assistant")
    say('What do you want me to do')


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.adjust_for_ambient_noise(source, duration=1)
        audio = r.listen(source)
        try:
            print("Recognizing...")
            call = r.recognize_google(audio, language='en-in')
            print(f"User said: {call}\n")
        except Exception as e:
            print(e)
            say("Unable to Recognize your voice.")
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
    call = takeCommand().lower()

    if 'time' in call:
        strTime = datetime.datetime.now().strftime("%H:%M:%S")
        say(f"the time is {strTime}")

    elif 'shutdown' in call:
        say("Your system is going to shut down")
        subprocess.call(["shutdown", "-f", "-s", "-t", "1"])

    elif 'empty recycle bin' in call:
        winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
        say("emptied recycle bin")

    elif "restart" in call:
        say("restarting your system")
        subprocess.call(["shutdown", "-f", "-r", "-t", "1"])

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
            text_to_translate = translator.translate(text=get_sentence, src=get_key(from_language),
                                                     dest=get_key(to_language))
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

    elif 'google' in call:
        say('What do you like to search')
        google_search = takeCommand()
        webbrowser.open('https://www.google.com/?#q=' + google_search)
        say("opening google ")

    elif 'youtube' in call:
        say('What do you like to search')
        youtube_search = takeCommand()
        webbrowser.open('https://www.youtube.com/results?search_query=' + youtube_search)
        say(("opening youtube"))

    elif 'powerpoint presentation' in call:
        say("opening Power Point presentation")
        power = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
        os.startfile(power)

    elif 'microsoft word' in call:
        say("opening microsoft word")
        word = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
        os.startfile(word)

    elif 'excel' in call:
        say("opening excel")
        excel = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
        os.startfile(excel)

    elif 'ms teams' in call:
        say("opening ms teams")
        teams = r"C:\Users\123.DESKTOP-74G99SN\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams"
        os.startfile(teams)

    elif "vs code" in call:
        say("opening vs code")
        vs_code = r"C:\Users\123.DESKTOP-74G99SN\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code\Visual Studio Code"
        os.startfile(vs_code)

    elif "don't listen" in call or "stop listening" in call:
        say("for how much time you want to stop jarvis from listening commands")
        say("for 20 seconds")
        say("or for 30 seconds")
        a = takeCommand()
        if "20" or "twenty" in a:
            say("not listening for 20")
            time.sleep(20)
            print(20)
        elif "30" or "thirty" in a:
            time.sleep(30)

    elif "goodbye" in call or 'bye' in call:
        say("good bye, see you later")
        break