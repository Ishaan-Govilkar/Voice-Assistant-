import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import ctypes
import time
from time import gmtime, strftime
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from AppOpener import open
#--------------------------------------------------------------------------
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
engine.setProperty('rate', 150)
#--------------------------------------------------------------------------
def speak(audio):
	engine.say(audio)
	engine.runAndWait()
#--------------------------------------------------------------------------
def wishMe():
	hour = int(datetime.datetime.now().hour)
	if hour>= 0 and hour<12:
		speak("Good Morning Sir !")

	elif hour>= 12 and hour<18:
		speak("Good Afternoon Sir !")

	else:
		speak("Good Evening Sir !")
#--------------------------------------------------------------------------
	assisname = "Zendroid"
	speak("I am your Assistant")
	#speak(assisname)
	#speak("I am your personal assistant that is ready to help you do anything, browse the web, show you photos etcetra. I am happy to help the Genz ee, hence my name Zendroid. You may also call me Zen.")
	
#--------------------------------------------------------------------------
def uname():
	uname = "ishaaan"
	print("Welcome Mr.", uname)
	speak("Welcome mister ishaaan")
	speak("How can i Help you, Sir")
#--------------------------------------------------------------------------
def takeCommand():
	
	r = sr.Recognizer()
	
	with sr.Microphone() as source:
		
		print("Listening...")
		r.pause_threshold = 1
		audio = r.listen(source)

	try:
		query = r.recognize_google(audio, language ='en-in')
		print(f"User said: {query}\n")

	except Exception as e:
		print(e)
		print("Unable to pick up any voice.")
		return "None"
	
	return query

if __name__ == '__main__':
	clear = lambda: os.system('cls')
	
	# This Function will clean any
	# command before execution of this python file
	clear()
	wishMe()
	uname()
	
	while True:
		
		query = takeCommand().lower()
		
		# All the commands said by user will be
		# stored here in 'query' and will be
		# converted to lower case for easily
		# recognition of command

		if 'wikipedia' in query:
			speak('searching wikipedia...')
			query = query.replace('wikipedia', '')
			result = wikipedia.summary(query, sentences=3)
			print(result)
			speak(result)

		elif 'open youtube' in query:
			speak('opening youtube')
			webbrowser.open("https://www.youtube.com")

		elif 'open phone link' in query:
			speak('opening phone link')
			open("Phone Link")

		elif 'open google' in query:
			speak('opening google')
			webbrowser.open('https://www.google.co.uk')

		elif'open microsoft edge' in query:
			speak('opening microsoft edge')
			open('Microsoft Edge')

		elif 'play music' in query or "play song" in query:
			speak("Here you go with music")
			open('Spotify')
			
		elif 'open discord' in query:
			speak("Here you go to Discord")
			open('Discord')
			
		elif 'open photos' in query:
			speak("Here you go to your memories")
			open('Photos')

		

		elif'open google chat' in query:
			speak('opening google chat')
			open('Google Chat')

		elif'open wondershare filmora' in query:
			speak('opening wondershare filmora')
			open('Wondershare Filmora 12')

		elif'open microsoft ofice' in query:
			speak('opening microsoft office')
			open('Microsoft 365(Office)')

		elif'open microsoft store' in query:
			speak('opening microsoft store')
			open('Microsoft Store')

		elif'settings' in query:
			speak('opening settings')
			open('Settings')

		elif'open clipchamp' in query:
			speak('opening clipchamp video editor')
			open("Clipchamp - Video Editor")

		elif'open access' in query:
			speak('opening access')
			open('Access')

		elif'open cortana' in query:
			speak('bro! come on, i am better')

		elif'what do you think about google assistant' in query:
			speak('Of course she is better, but i am becoming better everyday!')
			print('Of course she is better, but i am becoming better everyday!')

		elif'what do you think about siri'in query:
			speak('Of course she is better, but i am becoming better everyday!')
			print('Of course she is better, but i am becoming better everyday!')

		elif 'open calculator' in query:
			speak("opening calculator")
			open ('Calculator')

		elif 'open camera' in query:
			speak("opening camera")
			open ('camera')

		elif 'open map' in query:
			speak("opening map")
			open ('Maps')

		elif 'open' and 'excel' in query:
			speak("opening excel")
			open ('Excel')

		elif'open one note' in query:
			speak('opening one note')
			open('OneNote for Windows 10')

		elif "weather" in query:
			speak('Here you go to seeing the weather')
			open('Weather')


		elif "stop listening" in query or "bye"in query:
			speak("Thanks for giving me your time")
			exit()

		elif 'the time' in query:   
			speak(strftime("%H:%M:%S"))

		elif 'how are you' in query:
			speak("I am fine, Thank you")
			speak("How are you, Sir")

		elif 'fine' in query or "good" in query:
			speak("It's good to know that your fine")

		elif "change my name to" in query:
			query = query.replace("change my name to", "")
			assisname = query

		elif "change name" in query:
			speak("What would you like to call me, Sir ")
			assisname = takeCommand()
			speak("Thanks for naming me")

		elif "what's your name" in query or "What is your name" in query:
			speak("My friends call me")
			speak(assisname)
			print("My friends call me", assisname)

		elif 'exit' in query:
			speak("Thanks for giving me your time")
			exit()

		elif "who made you" in query or "who created you" in query:
			speak("I have been created by Ishaaan.")

		elif "what can you do?" in query:
			speak("I am your personal assistant that is ready to help you do anything, browse the web, show you photos etcetra.")
			
		elif 'joke' in query:
			speak(pyjokes.get_joke())

		elif "calculate" in query:
			
			app_id = "ishaan.govilkar@gmail.com"
			client = wolframalpha.Client(app_id)
			indx = query.lower().split().index('calculate')
			query = query.split()[indx + 1:]
			res = client.query(query)
			answer = next(res.results).text
			print("The answer is " + answer)
			speak("The answer is " + answer)

		elif 'news' in query:
			speak('Should i give you news from Times Of India or Sakaal Times or open world news?')
			if "times of india" in query:
				webbrowser.open("https://timesofindia.indiatimes.com/india")

			elif "sakaal" in query:
				webbrowser.open("https://www.esakal.com")

			else:
				webbrowser.open("https://www.bbc.com/news/world")

		

		elif 'lock window' in query:
				speak("locking the device")
				ctypes.windll.user32.LockWorkStation()

		elif 'shutdown system' in query:
				speak("Hold On a Sec ! Your system is on its way to shut down")
				subprocess.call('shutdown / p /f')
				
		elif 'empty recycle bin' in query:
			winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
			speak("Recycle Bin Recycled")

		elif "where is" in query:
			query = query.replace("where is", "")
			location = query
			speak("User asked to Locate")
			speak(location)
			webbrowser.open("https://www.google.com/maps/place/"+location)

		elif "camera" in query or "take a photo" in query:
			ec.capture(0, "Zen Camera ", "img.jpg")

		elif "restart" in query:
			subprocess.call(["shutdown", "/r"])
			
		elif "hibernate" in query or "sleep" in query:
			speak("Hibernating")
			subprocess.call("shutdown / h")

		elif "log off" in query or "sign out" in query:
			speak("Make sure all the application are closed before sign-out")
			time.sleep(5)
			subprocess.call(["shutdown", "/l"])

		elif "write a note" in query:
			speak("What should i write, sir")
			note = takeCommand()
			file = open('zen_note.txt', 'w')
			speak("Sir, Should i include date and time")
			snfm = takeCommand()
			if 'yes' in snfm or 'sure' in snfm:
				strTime = datetime.datetime.now().strftime("% H:% M:% S")
				file.write(strTime)
				file.write(" :- ")
				file.write(note)
			else:
				file.write(note)
		
		elif "show note" in query:
			speak("Showing Notes")
			file = open("jarvis.txt", "r")
			print(file.read())
			speak(file.read(6))

		elif "zendorid" in query or "hey" in query or "hello" in query:
			speak("Hello! What can i do for you?")
			
		elif "open wikipedia" in query:
			webbrowser.open("wikipedia.com")

		elif "Good Morning" in query:
			speak("A warm" +query)
			speak("How are you Mister")
			speak(assisname)

		elif "what is" in query or "who is" in query:
			
			# Use the same API key
			# that we have generated earlier
			client = wolframalpha.Client("ishaan.govilkar@gmail.com")
			res = client.query(query)
			
			try:
				print (next(res.results).text)
				speak (next(res.results).text)
			except StopIteration:
				print ("No results")
#--------------------------------------------------------------------------
