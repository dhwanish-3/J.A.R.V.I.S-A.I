import speech_recognition as sr
import os
import win32com.client
import webbrowser
import openai
import datetime
from config import apikey

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def say(text):
    # os.system(f"say {text}") # macOS
    speaker.Speak(text) # windows

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said : {query}")
            return query
        except Exception as e:
            return "Some Error Occured : Sorry from JARVIS"

chatStr = ""

def chat(query):
    global chatStr
    print(chatStr)
    openai.api_key = apikey
    chatStr += f"Dhwany: {query}\n Jarvis: "
    try:
        response = openai.Completion.create(
            model = "text-davinci-003",
            prompt  = chatStr,
            temperature = 0.7,
            max_tokens = 256,
            top_p = 1,
            frequency_penalty = 0,
            presence_penalty = 0
        )
        say(response["choices"][0]["text"])
        chatStr += f"{response['choices'][0]['text']}\n"
        return response["choices"][0]["text"]
    except Exception as e:
        say("Something went wrong...")
        chatStr += "No response\n"
        return "Something went wrong..."

def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for Prompt: {prompt} \n************************** \n\n"
    try:
        response = openai.Completion.create(
            model = "text-davinci-003",
            prompt = prompt,
            temperature = 0.7,
            max_tokens = 256,
            top_p = 1,
            frequency_penalty =  0,
            presense_penalty = 0
        )
        text += response["choices"][0]["text"]
    except Exception as e:
        text += "No response"
    if not os.path.exists("Openai"):
        os.mkdir("Openai")
    with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip()}.txt", "w") as f:
        f.write(text)

# todo: Add more sites
sites = [
        ['Google', 'https://www.google.com'],
        ['Youtube', 'https://www.youtube.com'],
        ['Spotify', 'https://open.spotify.com'],
        ['Whatsapp', 'https://web.whatsapp.com'],
        ['GitHub', 'https://www.github.com'],
        ['Wikipedia', 'https://www.wikipedia.com'],
        ]

# todo: Add more apps
apps = [
        ['Android Studio', r"C:\Users\dhwan\Desktop\Android Studio.lnk"],
        ['Blender', r"C:\Users\dhwan\Desktop\Blender.lnk"],
        ['Spotify', r"C:\Users\dhwan\Desktop\Spotify.lnk"],
        ['Settings', r"C:\Users\dhwan\Desktop\Settings.lnk"],
        ['Ubuntu', r"C:\Users\dhwan\Desktop\Ubuntu.lnk"],
        ['V S Code', r"C:\Users\dhwan\Desktop\Visual Studio Code.lnk"],
        ['Video Player', r"C:\Users\Public\Desktop\VLC media player.lnk"],
        ]


musicPath = r"C:\Users\dhwan\Music\download\Post Malone - Chemical (Official Lyric Video).mp3"

if __name__ == '__main__':

    say("Hello, I am Jarvis A I  ... How can I help you...")

    while True:
        print("Listening...")
        command = takeCommand()

        for site in sites:
            if f"Open {site[0]}".lower() in command.lower():
                say(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])
        
        for app in apps:
            if f"Open {app[0]}".lower() in command.lower():
                say(f"Opening {app[0]} sir...")
                os.startfile(app[1])

        if "Play music".lower() in command:
            say("Launching the music app sir...")
            os.startfile(musicPath)

        elif "The time".lower() in command:
            hour = datetime.datetime.now().strftime("%H")
            min = datetime.datetime.now().strftime("%M")
            say(f"Sir time is {hour} {min}")
        
        elif "Use artificial intelligence".lower() in command.lower():
            ai(prompt=command)

        elif "Thank you".lower() in command.lower():
            say("I am happy to serve you...")

        elif "reset chat".lower() in command.lower():
            chatStr = ""
        
        elif "Quit Jarvis".lower() in command.lower():
            exit()
        
        else:
            print("Chatting...")
            chat(command)
        # say(command)