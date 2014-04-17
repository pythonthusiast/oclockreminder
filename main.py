__author__ = 'Eko Wibowo'

import systray
if __name__ == '__main__':
    import itertools, glob
    import win32com.client

    ag=win32com.client.Dispatch('Agent.Control')
    ag.Connected=1
    ag.Characters.Load('James')

    icons = itertools.cycle(glob.glob('*.ico'))
    hover_text = "What can I do for you Sir?"
    def say_the_time(sysTrayIcon):
        print "Hello World."
    def simon(sysTrayIcon):
        print "Hello Simon."
    def switch_icon(sysTrayIcon):
        sysTrayIcon.icon = icons.next()
        sysTrayIcon.refresh_icon()
    menu_options = (('Say the time', icons.next(), say_the_time),
                   )
    def bye(sysTrayIcon): print 'Bye, then.'

    systray.SysTrayIcon(icons.next(), hover_text, menu_options, on_quit=bye, default_menu_index=1)