import requests
import json
import time
import speech_recognition as sr
from win32com.client import Dispatch

def speak(text):
    speak = Dispatch("SAPI.SpVoice")
    voices = speak.GetVoices()
    speak.Voice = voices[1]  # Change voice
    speak.Rate = -2  # Slow down speech
    speak.Speak(text)

def take_command():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        recognizer.pause_threshold = 1
        audio = recognizer.listen(source)
    try:
        print("Recognizing...")
        query = recognizer.recognize_google(audio, language='en-in')
        print(f"User said: {query}")
        return query
    except:
        print("Sorry, I couldn't understand. Please try again.")
        return "None"

def summarize_text(text, words=10):
    return " ".join(text.split()[:words]) + "..."

if __name__ == '__main__':
    speak("Would you like to type or speak the topic?")
    print("1. Type\n2. Speak")
    choice = input("Choose an option: ")

    if choice == "2":
        speak("Listening for news topic...")
        topic = take_command()
    else:
        print("\n--Topics--")
        print("bussiness,entertainment,health,science,sports,technology","\n")
        topic = input("Enter a topic: ")
    
    url = f"https://newsapi.org/v2/everything?q={topic}&sortBy=publishedAt&apiKey=YOUR_API_KEY"
    news_dict = requests.get(url).text
    news_dict = json.loads(news_dict)
    arts = news_dict.get('articles', [])

    num_articles = int(input("How many news articles do you want to hear? "))
    
    for i, article in enumerate(arts[:num_articles]):
        title = summarize_text(article['title'])
        speak(title)
        print(f"{i+1}. {title}")
        print("Source:", article['source']['name'])
        print("Read more:", article['url'])

        speak("Moving on to the next news...")
        time.sleep(2)
        print("\n")
    
