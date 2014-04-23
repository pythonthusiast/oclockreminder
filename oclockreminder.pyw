__author__ = 'Eko Wibowo'

import systray
import win32com.client
import datetime
import threading
import time
import pythoncom

class MsAgent(object):
    '''
    Construct an MS Agent object, display it and let it says the time exactly every hour
    '''
    def __init__(self):
        self.charId = 'James'

    def say_the_time(self, sysTrayIcon):
        '''
        Speak up the time!
        '''
        pythoncom.CoInitialize()
        agent = win32com.client.Dispatch('Agent.Control')
        charId = 'James'
        agent.Connected=1
        try:
            agent.Characters.Load(charId)
        except Exception as ex:
            print ex

        now = datetime.datetime.now()

        if now.hour == 0 and now.minute == 0:
            str_now = ' exactly %s o''clock' % now.hour
        else:
            str_now = '%s:%s:%s' % (now.hour, now.minute, now.second)
        agent.Characters(charId).Show()

        agent.Characters(charId).Speak('The time is %s' % str_now)
        time.sleep(3)
        agent.Characters(charId).Hide()
        time.sleep(5)
        agent.Characters.Unload(charId)
        pythoncom.CoUninitialize()

    def wakeup_next_hour(self):
        '''
        Run a thread that will wake up exactly n-o'clock
        '''
        now = datetime.datetime.now()
        next_hour = now + datetime.timedelta(hours = 1)
        next_hour_oclock = datetime.datetime(next_hour.year, next_hour.month, next_hour.day, next_hour.hour, 0, 0)
        seconds = next_hour_oclock - now
        self.thread = threading.Timer(seconds.total_seconds(), self.say_the_time_hourly)
        self.thread.start()

    def say_the_time_hourly(self):
        '''
        say the time and then schedule for another hour at exactly n-th o'clock
        '''
        self.say_the_time(None)
        self.wakeup_next_hour()

    def bye(self, sysTrayIcon):
        '''
        Stop any running thread, if any
        '''
        if hasattr(self,'thread'):
            self.thread.cancel()

if __name__ == '__main__':
    import itertools, glob

    icons = itertools.cycle(glob.glob('*.ico'))
    hover_text = "What can I do for you Sir?"

    agent = MsAgent()
    menu_options = (('Say the time', icons.next(), agent.say_the_time),)

    agent.wakeup_next_hour()
    systray.SysTrayIcon(icons.next(), hover_text, menu_options, on_quit=agent.bye, default_menu_index=1)