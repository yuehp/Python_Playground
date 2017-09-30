# install Active.Tck, in Python 3.x first 't' in tkinter is in lowercase
from tkinter import Tk
from time import sleep
# tkMessageBox in Python2.x becomes  tkinter.messagebox in Python3.x
from tkinter.messagebox import showwarning
# install PythonWin
import win32com.client as win32

warn = lambda app: showwarning(app, 'Exit?')
RANGE = range(3, 8)

def ppoint():
    app = 'PowerPoint'
    ppoint = win32.gencache.EnsureDispatch('%s.Application' % app)
    pres = ppoint.Presentations.Add()
    ppoint.Visible = True

    s1 = pres.Slides.Add(1, win32.constants.ppLayoutText)
    sleep(1)
    # TypeError: 'Shapes' object does not support indexing
    # API chagned. From 'Shapes[0]' to 'Shapes(1)'
    # the index used to start from 0, now it starts from 1
    s1a = s1.Shapes(1).TextFrame.TextRange
    s1a.Text = 'Python-to-%s Demo' % app
    sleep(1)
    s1b = s1.Shapes(2).TextFrame.TextRange
    for i in RANGE:
        s1b.InsertAfter("Line %d\r\n" %i)
        sleep(1)
    s1b.InsertAfter("\r\nTh-th-th-that's all folks!")

    s2 = pres.Slides.Add(2, win32.constants.ppLayoutText)
    s2a = s2.Shapes(1).TextFrame.TextRange
    s2a.Text = 'Hello'

    s1a.Text = 'Changed'

    warn(app)
    pres.Close()
    ppoint.Quit()

if __name__ == '__main__':
    Tk().withdraw()
    ppoint()
