__author__ = 'Eko Wibowo'

import systray
import win32com.client
import datetime
import threading
import time

class MsAgent(object):
    '''
    Construct an MS Agent object, display it and let it says the time exactly every hour
    '''
    def __init__(self):
        self.agent=win32com.client.Dispatch('Agent.Control')
        self.charId = 'James'
        self.agent.Connected=1
        self.agent.Characters.Load(self.charId)

    def say_the_time(self, sysTrayIcon):
        '''
        Speak up the time!
        '''
        
        now = datetime.datetime.now()
        str_now = '%s:%s:%s' % (now.hour, now.minute, now.second)
        self.agent.Characters(self.charId).Show()
        self.agent.Characters(self.charId).Speak('The time is %s' % str_now)
        time.sleep(3)
        self.agent.Characters(self.charId).Hide()
        time.sleep(5)
        self.agent.Characters.Unload(self.charId)

    def wakeup_next_hour(self):
        '''
        Run a thread that will wake up exactly n-o'clock
        '''
        now = datetime.datetime.now()
        next_hour = now + datetime.timedelta(seconds = 3)
        next_hour_oclock = datetime.datetime(next_hour.year, next_hour.month, next_hour.day, next_hour.hour, next_hour.minute, next_hour.second )#0, 0)
        seconds = next_hour_oclock - now
        thread = threading.Timer(seconds.total_seconds(), self.say_the_time_hourly)
        thread.start()

    def say_the_time_hourly(self):
        self.say_the_time(None)
        self.wakeup_next_hour()

    def bye(self, sysTrayIcon):
        '''
        Unload msagent object from memory
        '''
        self.agent.Characters.Unload(self.charId)
        self.thread.cancel()

if __name__ == '__main__':
    import itertools, glob

    icons = itertools.cycle(glob.glob('*.ico'))
    hover_text = "What can I do for you Sir?"

    agent = MsAgent()
    menu_options = (('Say the time', icons.next(), agent.say_the_time),)

    agent.wakeup_next_hour()
    systray.SysTrayIcon(icons.next(), hover_text, menu_options, on_quit=agent.bye, default_menu_index=1)