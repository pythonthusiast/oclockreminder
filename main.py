__author__ = 'Eko Wibowo'

import systray
import win32com.client
import datetime
import threading

class MsAgent(object):
    def __init__(self):
        self.ag=win32com.client.Dispatch('Agent.Control')
        self.ag.Connected=1
        self.ag.Characters.Load('James')

    def say_the_time(self, sysTrayIcon):
        try:
            self.ag.Connected=1
            self.ag.Characters.Load('James')
        except Exception as ex:
            pass
        now = datetime.datetime.now()
        str_now = '%s:%s:%s' % (now.hour, now.minute, now.second)
        self.ag.Characters('James').Show()
        self.ag.Characters('James').Speak('The time is %s' % str_now)
        self.ag.Characters('James').Hide()

    def bye(self, sysTrayIcon):
        self.ag.Characters.Unload('James')

    def wakeup_next_hour(self):
        now = datetime.datetime.now()
        next_hour = now + datetime.timedelta(hours = 1)
        next_hour_oclock = datetime.datetime(next_hour.year, next_hour.month, next_hour.day, next_hour.hour, 0, 0)
        seconds = next_hour_oclock - now
        t = threading.Timer(seconds.total_seconds(), self.say_the_time)
        t.start()

if __name__ == '__main__':
    import itertools, glob

    icons = itertools.cycle(glob.glob('*.ico'))
    hover_text = "What can I do for you Sir?"

    agent = MsAgent()
    menu_options = (('Say the time', icons.next(), agent.say_the_time),
                   )

    agent.wakeup_next_hour()

    systray.SysTrayIcon(icons.next(), hover_text, menu_options, on_quit=agent.bye, default_menu_index=1)