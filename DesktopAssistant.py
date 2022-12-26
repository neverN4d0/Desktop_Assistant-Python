import tkinter as tk
import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import smtplib
import PyPDF2
import win32com.client
from tkinter import *




names = []
emails = []
passwords = []



def speaktext(text):

    friend = pyttsx3.init()
    voices = friend.getProperty('voices')
    friend.setProperty('voice',voices[1].id)
    friend.say(text)
    friend.runAndWait()



engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)



def speak(audio):


    engine.say(audio)
    engine.runAndWait()



def wishMe():


    hour = int(datetime.datetime.now().hour)

    if hour>=0 and hour<12:

        speak("Good Morning!")

    elif hour>=12 and hour<18:

        speak("Good Afternoon!")   

    else:
        
        speak("Good Evening!")  

    speaktext("Hello, I am the RD bot. How may I help you")  



def audioinput():

    
    r = sr.Recognizer()
    with sr.Microphone() as source:

        r.pause_threshold = 1
        audio = r.listen(source)

    try: 

        query = r.recognize_google(audio, language='en-in')

    except Exception as e:

        print(e)
        return "None"

    return query  



def sendEmail(to, content):


    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('desktopassistantesting@gmail.com', 'thisisadummyaccount')
    server.sendmail('desktopassistant@gmail.com', to, content)
    server.close()



def mainpage():


    if __name__ == "__main__":

        wishMe()
        while True:
            
            speaktext("Listening...")
            query = audioinput().lower()



            if 'wikipedia' in query:

                speaktext('Searching Wikipedia...')
                query = query.replace("wikipiidea", "")
                results = wikipedia.summary(query, sentences=2)
                speaktext("According to Wikipedia")
                speaktext(results)
                print('According to Wikipedia')
                print(results)
                


            elif 'open youtube' in query:

                speaktext('What would you like to search for?')
                speaktext('Listening...')
                query_youtube = audioinput().lower()
                webbrowser.open(f"https://www.youtube.com/search?search_query="+query_youtube)



            elif 'play motivational music' in query:

                speaktext("Opening music...")
                music_dir = 'C:\\Users\\REHAA\\OneDrive\\Documents\\Songs'
                songs = os.listdir(music_dir)
                os.startfile(os.path.join(music_dir, songs[0]))


            
            elif 'open google' in query:
                
                speaktext('What would you like to search for?')
                speaktext('Listening...')
                query_google = audioinput().lower()
                webbrowser.open(f"https://www.google.com/search?q="+query_google)
                


            elif 'open stackoverflow' in query:
                
                speaktext('What would you like to search for?')
                speaktext("Listening...")
                query_stackoverflow = audioinput().lower()
                webbrowser.open(f"https://stackoverflow.com/search?q="+query_stackoverflow)



            elif 'open github' in query:

                speaktext('What would you like to search for?')
                speaktext("Listening...")
                query_github = audioinput().lower()
                webbrowser.open(f"https://github.com/search?q="+query_github)
            


            elif 'open word' in query:
                
                speaktext('What would you like to type out?')
                query_word = audioinput().lower()
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = True
                document = word.Documents.Add()
                selection = word.Selection
                selection.TypeText(query_word)
                document.SaveAs("C:\\Users\\REHAA\\OneDrive\\Documents\\document.docx")
            

            elif 'email to rehaan' in query:
                try:
                    speak("What should I say?")
                    content = audioinput()
                    to = "rehaan.dev@gmail.com"
                    sendEmail(to, content)
                    speaktext("Email has been sent to Rehaan")
                except Exception as e:
                    print(e)
                    speaktext("Sorry my friend, I am unable to send this email")



            elif 'read othello' or 'read of hello' in query:

                book = open('UNANNOTATED_OTHELLO_TEXT-PDF.pdf', 'rb')
                reader = PyPDF2.PdfFileReader(book)
                pages = reader.numPages
                print(pages)

                speaker = pyttsx3.init()    
                speaktext('What page number would you like to read? Please enter it below')
                page_choice = int(input("What page number would you like to read from Othello?: "))
                page = reader.getPage(page_choice-1)
                text = page.extractText()
                speaker.say(text)
                speaker.runAndWait()



            elif 'the time' in query:

                strTime = datetime.datetime.now().strftime("%H:%M:%S")    
                speaktext(f"Sir, the time is {strTime}")  



            elif query == "":

                speaktext('Please repeat that')
                return 



            elif 'stop' in query:
                speaktext('Desktop Assistant closed')
                break
                


            elif 'close' in query:
                speaktext('Desktop Assistant closed')
                break
                
            
     
def aboutme():


    aboutwindow = tk.Tk()
    aboutwindow.title('About Me')
    aboutwindow.geometry("1920x1080")
    aboutwindow.configure(bg="#000033")

    form_frame_four = tk.Frame(aboutwindow)
    generalinfo = tk.Label(form_frame_four, text="This is a basic desktop assistant that performs basic functions such as opening google. Once you hear the word 'listening...' wait for one second and speak your command")
    moreinfo = tk.Label(form_frame_four, text="To perform a command, for example opening word, simply say 'open word'. This works for other commands such as opening google.")

    form_frame_four.pack()
    generalinfo.pack(side='top')
    moreinfo.pack(side='bottom')

    aboutwindow.mainloop

    

def helppage():


    helpwindow = tk.Tk()
    helpwindow.title("Help Page")
    helpwindow.geometry("1920x1080")
    helpwindow.configure(bg="#000033")

    form_frame_three = tk.Frame(helpwindow)
    aboutmebutton = tk.Button(form_frame_three, text='About Me', command=aboutme)

    aboutmebutton.pack(side='bottom')
    form_frame_three.pack()

    helpwindow.mainloop()



def submit():


    name = name_entry.get()
    email = email_entry.get()
    password = password_entry.get()

    names.append(name)
    emails.append(email)
    passwords.append(password)

    name_entry.delete(0, 'end')
    email_entry.delete(0, 'end')
    password_entry.delete(0, 'end')

    window.destroy()

    new_window = tk.Tk()
    new_window.geometry("1920x1080")
    new_window.configure(bg="#000033")
    new_window.title("RD Technologies")

    form_frame_two = tk.Frame(new_window)
    open_button = tk.Button(form_frame_two, text='Open Desktop Assistant?', command=mainpage)
    helpbutton = tk.Button(form_frame_two, text="Help Page", command=helppage)

    helpbutton.pack(side='bottom')
    open_button.pack()
    form_frame_two.pack()


    new_window.mainloop()



window = tk.Tk()
window.title("Sign Up")
window.geometry("1920x1080")

window.grid_columnconfigure(0, weight=1)
window.grid_rowconfigure(0, weight=1)

form_frame = tk.Frame(window)

name_label = tk.Label(form_frame, text="Name:")
name_entry = tk.Entry(form_frame)
email_label = tk.Label(form_frame, text="Email:")
email_entry = tk.Entry(form_frame)
password_label = tk.Label(form_frame, text="Password:")
password_entry = tk.Entry(form_frame, show="*")

submit_button = tk.Button(form_frame, text="Submit", command=submit)

name_label.pack()
name_entry.pack()
email_label.pack()
email_entry.pack()
password_label.pack()
password_entry.pack()
submit_button.pack()

form_frame.pack()

window.mainloop()

