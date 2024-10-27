from win32com.client import Dispatch
import requests
from bs4 import BeautifulSoup
import datetime
from covid_india import states
import pytz
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import pyjokes

def speak(audio):
    speaking = Dispatch('SAPI.Spvoice')
    speaking.speak(audio)
def open_cmd():
    os.system('start cmd')
def open_calculator():
    os.system('start calculator')
def open_notepad():
    os.system('start notepad')
def covid():
    while True:
        speak('Which state information do you want?')
        print('Which state information do you want?')
        places = takeCommand()
        if 'Total corona cases' in places:
            updates = states.getdata('Total')
            print(updates)
        elif 'cases in Telangana' in places:
            updates = states.getdata('Telangana')
            print(updates)
        elif 'cases in Andhra pradesh' in places:
            updates = states.getdata('Andhra Pradesh')
            print(updates)
        elif 'cases in Arunachal pradesh' in places:
            updates = states.getdata('Arunachal Pradesh')
            print(updates)
        elif 'cases in Assam' in places:
            updates = states.getdata('Assam')
            print(updates)
        elif 'cases in Bihar' in places:
            updates = states.getdata('Bihar')
            print(updates)
        elif 'cases in Chandigarh' in places:
            updates = states.getdata('Chandigarh')
            print(updates)
        elif 'cases in Chhattisgarh' in places:
            updates = states.getdata('Chhattisgarh')
            print(updates)
        elif 'cases in Delhi' in places:
            updates = states.getdata('Delhi')
            print(updates)
        elif 'cases in Goa' in places:
            updates = states.getdata('Goa')
            print(updates)
        elif 'cases in Gujarat' in places:
            updates = states.getdata('Gujarat')
            print(updates)
        elif 'cases in Haryana' in places:
            updates = states.getdata('Haryana')
            print(updates)
        elif 'cases in Himachal Pradesh' in places:
            updates = states.getdata('Himachal Pradesh')
            print(updates)
        elif 'cases in Jammu and Kashmir' in places:
            updates = states.getdata('Jammu and Kashmir')
            print(updates)
        elif 'cases in Jharkhand' in places:
            updates = states.getdata('Jharkhand')
            print(updates)
        elif 'cases in Karnataka' in places:
            updates = states.getdata('Karnataka')
            print(updates)
        elif 'cases in Kerala*' in places:
            updates = states.getdata('Kerala*')
            print(updates)
        elif 'cases in Ladakh' in places:
            updates = states.getdata('Ladakh')
            print(updates)
        elif 'cases in Lakshadweep' in places:
            updates = states.getdata('Lakshadweep')
            print(updates)
        elif 'cases in Madhya Pradesh' in places:
            updates = states.getdata('Madhya Pradesh')
            print(updates)
        elif 'cases in Maharashtra' in places:
            updates = states.getdata('Maharashtra')
            print(updates)
        elif 'cases in Manipur' in places:
            updates = states.getdata('Manipur')
            print(updates)
        elif 'cases in Meghalaya' in places:
            updates = states.getdata('Meghalaya')
            print(updates)
        elif 'cases in Mizoram' in places:
            updates = states.getdata('Mizoram')
            print(updates)
        elif 'cases in Nagaland' in places:
            updates = states.getdata('Nagaland')
            print(updates)
        elif 'cases in Odisha' in places:
            updates = states.getdata('Odisha')
            print(updates)
        elif 'cases in Puducherry' in places:
            updates = states.getdata('Puducherry')
            print(updates)
        elif 'cases in Punjab' in places:
            updates = states.getdata('Punjab')
            print(updates)
        elif 'cases in Rajasthan' in places:
            updates = states.getdata('Rajasthan')
            print(updates)
        elif 'cases in Sikkim' in places:
            updates = states.getdata('Sikkim')
            print(updates)
        elif 'cases in Tamil Nadu' in places:
            updates = states.getdata('Tamil Nadu')
            print(updates)
        elif 'cases in Tripura' in places:
            updates = states.getdata('Tripura')
            print(updates)
        elif 'cases in Uttarakhand' in places:
            updates = states.getdata('Uttarakhand')
            print(updates)
        elif 'cases in Uttar Pradesh' in places:
            updates = states.getdata('Uttar Pradesh')
            print(updates)
        elif 'cases in West Bengal' in places:
            updates = states.getdata('West Bengal')
            print(updates)
        elif 'exit' in places:
            speak('i hope you got the information, call me aytime mam! have a great day')
            quit()
def date():
    todayDate=datetime.datetime.now(tz=pytz.timezone('Asia/Kolkata'))
    print(todayDate.strftime('%d %B, %Y'))
date()
def greet():
    t_hour = datetime.datetime.now().hour
    if 24> t_hour <4:
        speak("Pleasant Night mam!, Jarvis at Your Command")
    elif 12> t_hour >4:
        speak("Good Morning mam, Jarvis at Your Command")
    elif 18> t_hour >12:
        speak("Good Afternoon mam!, Jarvis at Your Command")
    else:
        speak("Good Evening mam!, Jarvis at Your Command")

greet()



def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Jarvis At Your Service!")
        r.pause_threshold = 2
        command = r.listen(source)
    try:
        print("Recognizing...")
        recognized = r.recognize_google(command, language='en-in')
        print(recognized)
    except Exception as e:
        print(e)
        statement="Pardon mam.., I Couldn't Recognize Your Voice, If its nothing to command, i'll take a leave"
        print(statement)
        speak(statement)
        return None
    return recognized

def jokes():
    my_jokes = pyjokes.get_joke('en', category='all')
    print(my_jokes)
    speak(my_jokes)
def Temp():
    search = "Temperature in Hyderabad"
    url = f"https://www.google.com/search?q={search}"
    r = requests.get(url)
    data = BeautifulSoup(r.text, "html.parser")
    temperature = data.find("div",class_ ="BNeawe" ).text
    speak(f"The Temperature in Hyderabad is{temperature} celcius")
    print(f"The Temperature in Hyderabad is{temperature} celcius")


if _name_ == '_main_':
    while True:
        query = takeCommand().lower()
        if 'what is the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"madam, the time is {strTime}")
            print(f"madam, the time is {strTime}")
        elif "what is your name" in query:
            speak("iam jarvis ai voice assistant")
            print("iam jarvis ai voice assistant")
        elif 'open notepad' in query:
            open_notepad()
        elif ' today date' in query:
            date()
        elif 'open command prompt' in query or 'open cmd' in query:
            open_cmd()
        elif 'open calculator' in query:
            open_calculator()
        elif 'wikipedia' in query:  #if wikipedia found in the query then this block will be executed
            speak('Searching Wikipedia...')
            print('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print("According to Wikipedia")
            print(results)
            speak(results)
        elif 'how are you' in query:
            speak("I am fine, Thank you")
            print("I am fine, Thank you")
            speak("How are you, mam")
            print("How are you, mam")
        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")
            print("It's good to know that your fine")
        elif "who are you" in query:
            speak("I am your jarvis voice assistant ")
            print("I am your jarvis voice assistant ")
        elif "open google" in query:
            webbrowser.open_new_tab("google.com")
        elif "open youtube" in query:
            webbrowser.open_new_tab("youtube.com")
        elif "open facebook" in query:
            webbrowser.open_new_tab("facebook.com")
        elif 'corona cases' in query:
            covid()
        elif "open instagram" in query:
            webbrowser.open_new_tab("instagram.com")
        elif "open chrome" in query:
            speak("what should i search madam?")
            print("what should i search madam?")
            search = takeCommand()
            chromepath = 'C://Program Files//Google//Chrome//Application//chrome.exe %s'
            webbrowser.get(chromepath).open_new_tab(search+'.com')
        elif "make a note" in query:
            speak("What Should i write down sir?")
            print("What Should i write down sir?")
            note = takeCommand()
            remember = open('pytext.txt', 'w')
            remember.write(note)
            remember.close()
            speak("content added successfully in pytext.txt" + note)
            print("content added successfully in pytext.txt" + note)
        elif "tell me the temperature" in query:
            Temp()
        elif "tell me a joke" in query:
            jokes()
        elif 'exit' in query:
            speak('mam . Call Me Anytime, at your service')
            print('mam . Call Me Anytime, at your service')
            quit()