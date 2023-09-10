import os
import random
import openai
from apiConfig import apikey
import wikipedia
import speech_recognition as sr
import webbrowser
import pyjokes
import pywhatkit
import datetime
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")

chatStr = ""

def chat(conversation):
    global chatStr
    print(chatStr)
    openai.api_key = apikey

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=conversation,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    chatStr += f"User : {conversation}\n\nPersonal AI Answer : "
    try:
        if "write".lower() in conversation.lower():
            aiCall(conversation)
            return ""
        say(response["choices"][0]["text"])
        chatStr += f"{response['choices'][0]['text']}\n"
        return response["choices"][0]["text"]
    except Exception as e:
        return "Some Error has Occurred"



def aiCall(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"

    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    try:
        text += response["choices"][0]["text"]
        if not os.path.exists("Openai"):
            os.mkdir("Openai")
        # with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip() }.txt", "w") as f:
        with open(f"Openai/FileNo. - {random.randint(1,100)}.txt", "w") as f:
            f.write(text)
    except Exception as e:
        return "Sorry! some error Occurs."

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

    print('Hello I am your personal AI')
    say("Hello I am your personal AI")

    while True:
        print("Listening ....")
        command = takeCommand()
        print("User Said : ", command)
        print("Recognized the user command")

        if 'hello' in command.lower():
            say('hello, how can i help you ??')

        elif 'who are you' in command:
            say('I am an mini ai  virtual assistant to help you')

        elif 'can you do' in command:
            say('''i can play songs on youtube , can chat with you, can write email as per your demand on subject, 
            tell you a joke, search on wikipedia, tell date and time,find your location, locate area on map, 
            open different websites like instagram, youtube,gmail, git hub, stack overflow and searches on google. 
            How may i help you ??''')


        elif 'play'.lower() in command.lower():
            song = command.replace('play', '')
            say('Playing' + song)
            pywhatkit.playonyt(song)

        elif 'tell me about'.lower() in command.lower():
            name = command.replace('tell me about', '')
            info = wikipedia.summary(name, 1)
            say(info)

        elif "open site".lower() in command.lower():
            site = command[10:]
            site = site.strip()
            if f"Open site {site}".lower() in command.lower():
                say(f"Opening {site}...")
                webbrowser.open(f"https://www.{site}.com")

        elif "music".lower() in command.lower():
            MusicPath = "C:\\Users\\dhira\\PycharmProjects\\AIvoiceAssistant\\Dracula.mp4"
            from os import startfile

            startfile(MusicPath)

        elif "time".lower() in command.lower():
            time_now = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"Time now is : {time_now}")

        elif 'date'.lower() in command:
            today = datetime.date.today()
            d2 = today.strftime("%B %d, %Y")
            say(f'The todays date is {d2}')

        elif 'joke'.lower() in command:
            _joke = pyjokes.get_joke()
            print(_joke)
            say(_joke)

        elif 'search'.lower() in command:
            search = 'https://www.google.com/search?q=' + command
            say('searching... ')
            webbrowser.open(search)

        elif "thank you".lower() in command.lower():
            say("Welcome! I am very happy by helping you sir.")

        elif "close".lower() in command.lower():
            say("Good bye! Have a Nice day sir!")
            break

        elif 'stop' in command.lower():
            say("Good bye! Have a Nice day sir!")
            break

        elif 'tata' in command.lower():
            say("Good bye! Have a Nice day sir!")
            break

        elif 'bye' in command.lower():
            say("Good bye! Have a Nice day sir!")
            break

        elif "Using Open AI".lower() in command.lower():
            aiCall(prompt=command)

        else:
            print("Now Chatting...")
            chat(command)
