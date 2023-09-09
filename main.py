import os
import wikipedia
import speech_recognition as sr
import webbrowser
import pyjokes
import pywhatkit
import datetime
import win32com.client
speaker = win32com.client.Dispatch("SAPI.SpVoice")


def say(text):
    speaker.Speak(text)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as input:
        r.pause_threshold = 1
        audio = r.listen(input)
        try:
            query = r.recognize_google(audio, language="en-in")
            return query
        except Exception as e:
            return "Some Error Occurred"

if __name__ == '__main__':
    print('PyCharm')
    say("Hello I am your personal AI")

    while True:
        print("Listening ....")
        command = takeCommand()
        print("User Said : ",command)
        print("Recognized the user command")

        if 'hello' in command.lower():
            say('hello, how can i help you ??')

        elif 'who are you' in command:
            say('I am an mini ai  virtual assistant to help you')

        elif 'can you do' in command:
            say('''i can play songs on youtube , tell you a joke, search on wikipedia,
             tell date and time,find your 
            location, locate area on map, open different websites like instagram, youtube,gmail, git hub, 
            stack overflow and searches on google. How may i help you ??''')


        elif 'play' in command:
            song = command.replace('play', '')
            say('Playing' + song)
            pywhatkit.playonyt(song)

        elif 'tell me about' in command:
            name = command.replace('tell me about', '')
            info = wikipedia.summary(name, 1)
            say(info)

        elif "open site" in command.lower():
            site = command[10:]
            site = site.strip()
            if f"Open site {site}".lower() in command.lower():
                say(f"Opening {site}...")
                webbrowser.open(f"https://www.{site}.com")

        elif "music" in command.lower():
            MusicPath = "C:\\Users\\dhira\\PycharmProjects\\AIvoiceAssistant\\Dracula.mp4"
            from os import startfile
            startfile(MusicPath)

        elif "time" in command.lower():
            time_now = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"Time now is : {time_now}")

        elif 'date' in command:
            today = datetime.date.today()
            d2 = today.strftime("%B %d, %Y")
            say(f'The todays date is {d2}')

        elif 'joke' in command:
            _joke = pyjokes.get_joke()
            print(_joke)
            say(_joke)

        elif 'search' in command:
            search = 'https://www.google.com/search?q=' + command
            say('searching... ')
            webbrowser.open(search)

        elif "thank you" in command.lower():
            say("Welcome! I am very happy by helping you sir.")

        elif "close".lower() in command.lower():
            say("Good bye! Have a Nice day sir!")
            break

        elif 'stop' in command:
            say("Good bye! Have a Nice day sir!")
            break

        elif 'tata' in command:
            say("Good bye! Have a Nice day sir!")
            break

        elif 'bye' in command:
            say("Good bye! Have a Nice day sir!")
            break
