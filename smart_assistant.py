import subprocess
import pyttsx3
import random2
import os
import winshell
import pyjokes
import feedparser
import smtplib
import datetime 
import requests
from bs4 import BeautifulSoup
import win32com.client as wincl
import pyautogui
import speech_recognition as sr
import gui_automation



voiceEngine = pyttsx3.init('sapi5')
voices = voiceEngine.getProperty('voices')
voiceEngine.setProperty('voice', voices[1].id)

def speak(text):
    voiceEngine.say(text)
    voiceEngine.runAndWait()

class gui_control:
        
    def mute_unmute(self):
        print("Pressing Mute/Unmute Key")
        pyautogui.typewrite(['volumemute'])

      
    def volume_up(self):
        pyautogui.typewrite(['volumeup'])
        

    def volume_down(self):
        pyautogui.typewrite(['volumedown']) 

    
    def appopen(self):
        pyautogui.hotkey("ctrl","esc")    
        pyautogui.typewrite(fw)
        pyautogui.press("enter")
        
    def ipadd(self):
        pyautogui.hotkey("ctrl","esc")
        pyautogui.typewrite("cmd")
        pyautogui.press("enter")
        pyautogui.typewrite("ipconfig")
        pyautogui.press("enter")
        
    def new_folder(self):
        pyautogui.hotkey("ctrl","esc")
        pyautogui.typewrite("cmd")
        pyautogui.press("enter")
        pyautogui.typewrite("cd documents")
        pyautogui.press("enter")
        pyautogui.typewrite("mkdir new")
        pyautogui.press("enter")
    
    def sysinfo(self):
        pyautogui.hotkey("ctrl","esc")
        pyautogui.typewrite("cmd")
        pyautogui.press("enter")
        pyautogui.typewrite("systeminfo")
        pyautogui.press("enter")
             
    def close(self):
        pyautogui.hotkey("alt","f4")
def getWeather(city_name):
    print('Displaying Weather report for: ' + city)
    url = 'https://wttr.in/{}'.format(city)
    res = requests.get(url)
    print(res.text)

    
    
    
gui = gui_control()
rec = sr.Recognizer()
print("Wishing.")
time = int(datetime.datetime.now().hour)
global uname,asname
if time>= 0 and time<12:
    speak("Good Morning!")
elif time<18:
    speak("Good Afternoon!")

else:
    speak("Good Evening!")


speak("I am your Voice Assistant,How can I help you?")


def takeCommand():
    recog = sr.Recognizer()
    
    with sr.Microphone() as source:
        print("Listening to the user")
        recog.pause_threshold = 1
        userInput = recog.listen(source)

    try:
        print("Recognizing the command")
        command = recog.recognize_google(userInput, language ='en-in')
        print(f"Command is: {command}\n")

    except Exception as e:
        print(e)
        print("Unable to Recognize the voice.")
        return "None"

    return command

with sr.Microphone() as mic:
    while True:
        try:
            print('you can start talking now')
            audio = rec.listen(mic)
            print("Recognizing the command")
            a = rec.recognize_google(audio)
            a = rec.recognize_google(audio).lower()

            
        except Exception as ex:
            print("Sorry. Could not understand.\n\n")
            continue
        print(a) 
        
        
        
        if (a == "quit program") or (a == "exit program"):
            break
    
        elif a == "mute" or a=="unmute":
            gui.mute_unmute()
    
        elif a == "volume up" or a=="sound up":
            gui.volume_up()
    
        elif a == "volume down" or a=="sound down":
            gui.volume_down()
    
        elif a== "ip address of the system":
            gui.ipadd()
        
        elif a== "create a new folder":
            gui.new_folder
        
        elif a=="system info" or a=="system information":
            gui.sysinfo()
            
        elif "weather" in a:
            speak(" Please tell your city name ")
            print("City name : ")
            city = takeCommand()
            getWeather(city) 

        elif 'time' in a:
            strTime = datetime.datetime.now()
            curTime=str(strTime.hour)+"hours"+str(strTime.minute)+"minutes"+str(strTime.second)+"seconds"
            speak(f" the time is {curTime}")
            print(curTime)
            
            
        elif 'joke' in a:
            speak(pyjokes.get_joke())
            
    
        elif "open" in a:
            n=2
            all=a.split(" ")
            fw = all[n-1]
            print(fw)
            gui.appopen()
            speak(fw)
        
        elif a =="close":
            gui.close()
        
        elif 'shutdown system' in a:
                speak("Hold On a Sec ! Your system is on its way to shut down")
                subprocess.call('shutdown / p /f')
                
        elif "restart" in a:
            subprocess.call(["shutdown", "/r"])
            
        elif "sleep" in a:
            speak("Setting in sleep mode")
            subprocess.call("shutdown / h")
            
        elif 'exit' in a:
            speak("Thanks for giving me your time")
            break
            
        
    