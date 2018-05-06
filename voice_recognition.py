# -*- coding: utf-8 -*-
import cv2
#from gtts import gTTS
import speech_recognition as sr
from win32com.client import constants, Dispatch
import os
import webbrowser
#import winsound
from pygame import mixer

speaker = Dispatch("SAPI.SpVoice")
chrome_path = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
webbrowser.register('chrome', None,webbrowser.BackgroundBrowser(chrome_path),1)
mixer.init()


def talkToMe(audio):
    print(audio)
    Msg = audio
    
    global speaker
    speaker.Speak(Msg)
    #os.system('mpg123 audio.mp3')
    #os.system("say '"+audio+"'")
    #tts = gTTS(text=audio,lang='en')
    #tts.save('audio.mp3')
    
def myCommand():

    r = sr.Recognizer()

    with sr.Microphone() as source:
        print('I am ready for your next command')
        #r.pause_threshold = 1
        #r.adjust_for_ambient_noise(source,duration=1)
        audio = r.listen(source)

    try:
        command = r.recognize_google(audio)
        print('You said: ' + command + '\n')
        if 'goodbye' in command or 'bye' in command:
            talkToMe('Good bye Krisna, see you later')
            return 'exit'
        
        assistant(command)
        #talkToMe('You said ' + command)
    
    except sr.UnknownValueError:
        #assistant(myCommand())
        return
    
    return command

def assistant(command):
    
    if 'hai Jarvis' in command or 'hey Jarvis' in command or 'hi' in command:
        talkToMe("hai Krisna")       
    elif 'how are you' in command or 'how are you today' in command or 'are you today' in command:
        #print("I'm fine thank you")
        talkToMe("I'm fine thank you")    
    elif "how the weather today" in command or "how's the weather today" in command:
        talkToMe("The weather is good, do you want to go somewhere ?")        
    elif 'thank you' in command or 'thanks' in command:
        talkToMe('no problem')    
    elif 'open YouTube' in command or 'open a YouTube' in command or 'open the YouTube' in command:
        talkToMe('Wait a moment')
        url = 'https://www.youtube.com'
        webbrowser.get('chrome').open_new_tab(url)
    elif 'close the YouTube' in command or 'close YouTube' in command:
        talkToMe('Wait a moment')
        os.system("taskkill /im chrome.exe /f")
    elif 'play a music' in command or 'play music' in command:
        talkToMe('Wait a moment')
        mixer.music.load("corazon-santana.mp3")
        mixer.music.play()
    elif 'stop the music' in command or 'stop music' in command:
        talkToMe('Wait a moment')
        mixer.music.stop()
    else:
        talkToMe('Please say again ?')

os.system('jarvis2.gif')
talkToMe("Hi I'm Jarvis, May I help you ?")
#cam = cv2.VideoCapture(0)


while True:

    ret = myCommand()
    
    if ret=='exit':
        #cam.release()
        os.system("taskkill /im Microsoft.Photos.exe /f")
        break
    
del speaker    
#cv2.destroyAllWindows()