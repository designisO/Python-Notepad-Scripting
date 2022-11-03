# Python Scripting Project: Speed Reader



import time
import win32com.client

###Shell Connection
shell = win32com.client.Dispatch("WScript.Shell")
shell.Run("notepad")
time.sleep(1)
shell.AppActivate("Notepad")

###Message
msg="""Introduction to Orion3000 


DesignisO@proton.me
10:48 PM : 5 minutes ago.
to me from "Speed Read Script"
I am a self taught developer and designer who enjoys innovation and tinkering. I am currently studying software development, IOT, and mobile app development while implementing great UI. As a hobby, I enjoy the world of coding so research is done often to keep with trends and updates. While doing the task at hand a great cup of coffee always keeps the "fuel of code". With that being said, feel free to donate a cup to let as the blog continues to grow with great content.  

From Orion3000.xyz
by Orion3000
"""  

###For Loop for sending characters in sequence with a delay

delay=0.03
for i in msg:
    time.sleep(delay)
    shell.SendKeys(i, 0)


