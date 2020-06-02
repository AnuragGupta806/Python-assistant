import pyttsx3
import win32com
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import smtplib

engine = pyttsx3.init('sapi5')  # sapi5 is a microsoft Api for speech to reply
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[0].id)

chromedir= 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
"""Set web browser location to google chrome"""

def speak(audio):
    engine.say(audio)
    engine.runAndWait()



def wishme():
    hour = int(datetime.datetime.now().hour)

    if (hour>=0 and hour<12) :
        speak("Good Morning!")

    elif (hour>= 12 and hour<18):
        speak("Good Afternoon!")

    else:
        speak("Good Evening!")

    speak("I m your assistant sir! Please tell how may I help you")



def takeCommand():
    """It takes microphone input from the user and returns string output"""
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening.....")
        r.pause_threshold = 1

        # r.energy_threshold = 100    
        """ this is used to give how much intensity the voice should be so that other noise cannot be heard"""
        audio = r.listen(source)

    try:
        print("Recognizing....")
        query = r.recognize_google(audio, language ='en-in' )
        print(f"User said : {query}\n")

    except Exception as e:
        print(e)            # print error for debugging code
        print("Say that again Please....")
        return "None"
    return query


def sendEmail(to, content):
    """[Function to send email]

    Arguments:
        to {[string]} -- [email address of the user]
        content {[type]} -- [content of the email]
    Requirement:
        MUST ENABLE LESS SECURED APPS TO SEND EMAIL TO YOUR GMAIL ACCOUNT
    """
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('youremail@gmail.com', 'your-password-here')
    server.sendmail('youremail@gmail.com', to, content)
    server.close()


if __name__ == "__main__":
    # speak("Hello how are you!!")
    wishme()
    while True:
        query = takeCommand().lower()
        
        # logic for executing task based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia.....')
            query = query.replace('wikipedia', "")
            result = wikipedia.summary(query, sentences =2)
            speak("According to Wikipedia:")
            print(result)
            speak(result)

        elif 'open youtube' in query:
            webbrowser.get(chromedir).open("Youtube.com")

        elif 'open google' in query:
            webbrowser.get(chromedir).open("Google.com")
            
        elif 'open stackoverflow' in query:
            webbrowser.get(chromedir).open_new_tab("stackoverflow.com")
        
        elif 'play music' in query:
            music = 'C:\\Users\\Asus\\Music'
            songs = os.listdir(music)
            # print(songs)
            os.startfile(os.path.join(music, songs[0]))
        
        elif 'the time' in query:
            strtime =datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strtime}")

        elif 'open code' in query:
            codePath = "C:\\Users\\Asus\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)

        elif 'email to anurag' in query:
            try:
                speak("What should i Say?")
                content = takeCommand()
                to = "recievermail@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry not able to send the email")
