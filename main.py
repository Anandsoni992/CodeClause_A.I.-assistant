import speech_recognition as sr
import wikipedia
import win32com.client
import webbrowser
import os
import datetime
import wolframalpha

speaker = win32com.client.Dispatch("SAPI.SpVoice")

appID = '8XAKPH-PVTKT4LU7Y'
wolframClient = wolframalpha.Client(appID)


def say(text):
    speaker.Speak(text)


def search_wikipedia(query=""):
    searchResults = wikipedia.search(query)
    if not searchResults:
        print("no wikipedia result ")
        return "no result received"
    try:
        wikiPage = wikipedia.page(searchResults[0])
    except wikipedia.DisambiguationError as error:
        wikiPage = wikipedia.page(error.options[0])
    print(wikiPage.title)
    wikiSummary = str(wikiPage.summary)
    return wikiSummary


def listOrdict(var):
    if isinstance(var, list):
        return var[0]['plaintext']
    else:
        return var['plaintext']


def search_wolframalpha(query =""):
    response = wolframClient.query(query)

    if response['@success'] == 'false':
        return "sorry, unable to compute"
    else:
        result = ""

        pod0 = response['pod'][0]
        pod1 = response['pod'][1]

        if (('result') in pod1['@title'].lower()) or (pod1.get('@primary)', 'false') == 'true') or (
                'definition' in pod1['@title'].lower()):
            result = listOrdict(pod1['subpod'])
            return result.split('(')[0]
        else:
            question = listOrdict(pod0['subpod'])
            return question.split('(')[0]


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.6
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said : {query}")
            return query
        except Exception as e:
            say("i did not quite catch that")


if __name__ == '__main__':
    say("hello my liege, how may i help you")
    while True:
        print("listening...")
        query = takeCommand()
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "htps://www.wikipedia.com"],
                 ["google", "https://www.google.com"]]
        for site in sites:
            if f"open {site[0]}".lower() in query.lower():
                say(f"opening {site[0]} my liege....")
                webbrowser.open(site[1])
        if "play some music" in query:
            say("playing L's theme , my liege")
            musicpath = "D:\Personnel AI\L's theme A.mp3"
            os.startfile(musicpath)

        if "what is the time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"My liege, the time is {strfTime} ")

        if 'Wikipedia' in query:
            query = " ".join(query.split()[1:])
            say("querying the universal data-blank....")
            say(search_wikipedia(query))

        if 'calculate' in query:
            query = " ".join(query.split()[1:])
            say("computing....")
            try:
                result = search_wolframalpha(query)
                say(result)
            except:
                say("sorry, unable to compute....")
        if 'log' in query:
            say("ready to record your note..")
            newNote = takeCommand().lower()
            now = datetime.now().strftime("%H:%M:%S")
            with open('note_%s.txt' % now, 'w') as newfile:
                newfile.write(newNote)
            say("note written")
        if "who are you" in query:
            say("I am a personal assistant created by anand")
        elif "what is your name" in query:
            say("My name is igris")
        elif "how are you" in query:
            say("I am fine, thank you")
        elif "what is your purpose" in query:
            say("I am created to help you")
        elif "who created you" in query:
            say("I was created by anand")
        elif "what is your age" in query:
            say("I am a computer program, I am not born yet")
        elif "what is your gender" in query:
            say("I am a computer program, I am not born yet")
        elif "what is your profession" in query:
            say("I am a computer program, I am not born yet")
        elif "what is your hobby" in query:
            say("I am a computer program, I am not born yet")
        elif "what is your favorite color" in query:
            say("I am a computer program, I am not born yet")
        if "exit" in query:
            say("Gooddbye ,  my liege")
            break
