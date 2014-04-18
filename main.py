__author__ = 'Eko Wibowo'

import systray
import win32com.client
import datetime
import threading

ag=win32com.client.Dispatch('Agent.Control')
ag.Connected=1
ag.Characters.Load('James')

def say_the_time(sysTrayIcon):
    try:
        ag.Connected=1
        ag.Characters.Load('James')
    except Exception as ex:
        pass
    now = datetime.datetime.now()
    str_now = '%s:%s:%s' % (now.hour, now.minute, now.second)
    ag.Characters('James').Show()
    ag.Characters('James').Speak('The time is %s' % str_now)
    ag.Characters('James').Hide()

def bye(sysTrayIcon):
    ag.Characters.Unload('James')

def wakeup_next_hour():
    now = datetime.datetime.now()
    next_hour = now + datetime.timedelta(hours = 1)
    next_hour_oclock = datetime.datetime(next_hour.year, next_hour.month, next_hour.day, next_hour.hour, 0, 0)
    seconds = next_hour_oclock - now
    t = threading.Timer(seconds.total_seconds(), say_the_time)
    t.start()

if __name__ == '__main__':
    import itertools, glob

    icons = itertools.cycle(glob.glob('*.ico'))
    hover_text = "What can I do for you Sir?"

    menu_options = (('Say the time', icons.next(), say_the_time),
                   )

    wakeup_next_hour()

    systray.SysTrayIcon(icons.next(), hover_text, menu_options, on_quit=bye, default_menu_index=1)