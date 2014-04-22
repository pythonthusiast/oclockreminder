__author__ = 'Eko Wibowo'

import systray
import win32com.client
import datetime
import threading
import sys
import time

sys.coinit_flags = 0
import pythoncom
from pythoncom import CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED

#class MsAgent(object):
'''
Construct an MS Agent object, display it and let it says the time exactly every hour
'''
#    def __init__(self):
#agent=win32com.client.Dispatch('Agent.Control')
#charId = 'James'
#agent.Connected=1
#agent.Characters.Load(charId)
#print agent
#thread = None

def say_the_time_hourly():
    say_the_time()

def say_the_time(sysTrayIcon):
    '''
    Speak up the time!
    '''
    CoInitializeEx(COINIT_MULTITHREADED)
    agent = win32com.client.Dispatch('Agent.Control')
    charId = 'James'
    agent.Connected=1
    try:
        agent.Characters.Load(charId)
    except Exception as ex:
        print ex

    now = datetime.datetime.now()
    str_now = '%s:%s:%s' % (now.hour, now.minute, now.second)
    agent.Characters(charId).Show()
    print 'Speak up!'
    speak = agent.Characters(charId).Speak('The time is %s' % str_now)
    time.sleep(3)
    hide = agent.Characters(charId).Hide()
    time.sleep(5)
    agent.Characters.Unload(charId)

def bye(sysTrayIcon):
    '''
    Unload msagent object from memory
    '''

    #agent.Characters.Unload(charId)
    #thread.cancel()

def wakeup_next_hour(sysTrayIcon):
    '''
    Run a thread that will wake up exactly n-o'clock
    '''
    now = datetime.datetime.now()
    next_hour = now + datetime.timedelta(seconds = 3)
    next_hour_oclock = datetime.datetime(next_hour.year, next_hour.month, next_hour.day, next_hour.hour, next_hour.minute, next_hour.second )#0, 0)
    seconds = next_hour_oclock - now
    thread = threading.Timer(seconds.total_seconds(), say_the_time_hourly, [sysTrayIcon])
    thread.start()

if __name__ == '__main__':
    import itertools, glob

    icons = itertools.cycle(glob.glob('*.ico'))
    hover_text = "What can I do for you Sir?"

    #agent = MsAgent()
    menu_options = (('Say the time', icons.next(), say_the_time),)

    trayApp = systray.SysTrayIcon(icons.next(), hover_text, menu_options, on_quit=bye, default_menu_index=1)
    wakeup_next_hour(trayApp)
    trayApp.pumpMessage()
