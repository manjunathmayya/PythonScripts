'''
This script runs an infinite loop, with a delay of 5 seconds for between iterations, waiting for our voice. 
If required phrases are found within these 5 seconds, corresponding actions are done.
I found 5 seconds as the optimum value, which can be changed ofcourse, based on your needs.

Currently, the script does 2 things. 
1. Open selected applications 
2. Search your entire computer for required files.
    For this to work, we need 'Everything' tool from https://www.voidtools.com/.
    Using its command line exe (EX.exe), search results are displayed in a command prompt.
    
Many more features can be added to this.    

'''


import speech_recognition as sr
import os

# obtain audio from the microphone
r = sr.Recognizer()

while True:

    with sr.Microphone() as source:
        print("\n ---  Say something! --- \n")
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source, phrase_time_limit = 5)    
    
    
    try:
        # for testing purposes, we're just using the default API key
        # to use another API key, use `r.recognize_google(audio, key="GOOGLE_SPEECH_RECOGNITION_API_KEY")`
        # instead of `r.recognize_google(audio)`
        
        text = r.recognize_google(audio).lower()
        
        if 'open' in text or 'start' in text:
            if 'notepad' in text:
                if 'plus' in text:
                    os.system("start notepad++.exe")
                else:
                    os.system("start notepad.exe")
            if 'visual studio' in text:
                os.system("start devenv.exe")
            if 'virtual machine' in text:
                os.system("start vmplayer.exe")  
            if 'spider' in text:
                os.system("start spyder") 
            if 'word' in text:
                os.system("start winword") 
            if 'excel' in text:
                os.system("start excel") 
            if 'everything' in text:
                path = "E:\Softwares\Everything-1.2.1.371"
                os.chdir(path)
                os.system("start Everything-1.2.1.371.exe") 
            if 'command' in text:
                os.system("start cmd.exe") 
            if 'chrome' in text:
                os.system("start chrome.exe") 
            
        elif 'search' in text:
            text = text.replace('search', '')
            path = "E:\Softwares\Everything-1.2.1.371\ES-1.1.0.10\es.exe"
            output = os.system(path + ' ' + text) 
                
    except sr.UnknownValueError:
        print("Google Speech Recognition could not understand audio")
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; {0}".format(e))        
