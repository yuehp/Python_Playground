#

from tkinter import Tk, Label, Entry, Button
from time import sleep
import win32com.client as win32

INDENT = '    '
DEMO = '''
PRESENTATION TITLE
    optional subtitle

slide 1 title
    slide 1 bullet 1
    slide 1 bullet 2

slide 2 title
    slide 2 bullet 1
    slide 2 bullet 2
        slide 2 bullet 2a
        slide 2 bullet 2b
'''

def txt2ppt(lines):
    ppoint = win32.gencache.EnsureDispatch(
        'PowerPoint.Application')
    pres = ppoint.Presentations.Add()
    ppoint.Visible = True
    sleep(2)
    nslide = 1
    for line in lines:
        if not line:
            continue
        linedata = line.split(INDENT)
        if len(linedata) == 1:
            title = (line == line.upper())
            if title:
                stype = win32.constants.ppLayoutTitle
            else:
                stype = win32.constants.ppLayoutText

            s = pres.Slides.Add(nslide, stype)
            ppoint.ActiveWindow.View.GotoSlide(nslide)
            s.Shapes(1).TextFrame.TextRange.Text = line.title()
            body = s.Shapes(2).TextFrame.TextRange
            nline = 1
            nslide += 1
            sleep((nslide<4) and 0.5 or 0.01)
        else:
            line = '%s\r\n' % line.lstrip()
            body.InsertAfter(line)
            para = body.Paragraphs(nline)
            para.IndentLevel = len(linedata) - 1
            nline += 1
            sleep((nslide<4) and 0.25 or 0.01)

    s = pres.Slides.Add(nslide, win32.constants.ppLayoutTitle)
    ppoint.ActiveWindow.View.GotoSlide(nslide)
    s.Shapes(1).TextFrame.TextRange.Text = "It's time for a slide-show".upper()
    sleep(1.)
    for i in range(3, 0, -1):
        s.Shapes(2).TextFrame.TextRange.Text = str(i)
        sleep(1.)

    pres.SlideShowSettings.ShowType = 1 # The constant does not exist. Use number instead. 'win32.constants.ShowType-Speaker'    # change 'ppShowType-Speaker' to 'ShowType-Speaker'
    ss = pres.SlideShowSettings.Run()
    pres.ApplyTemplate(r'd:\image.potx')       # replace with your own template here
    s.Shapes(1).TextFrame.TextRange.Text = 'FINIS'
    s.Shapes(2).TextFrame.TextRange.Text = ''

def _start(ev=None):
    fn = en.get().strip()   # what does 'en' mean here ?
    try:
        f = open(fn, 'U')
    except IOError:         # SyntaxError: invalid syntax
        import io #from cStringIO import StringIO     # use 'io.StringIO' instead in Python 3.x
        f = io.StringIO(DEMO)
        en.delete(0, 'end')
        if fn.lower() == 'demo':
            en.insert(0, fn)
        else:
            import os
            en.inset(0,
                r"DEMO (can't open %s: %s" % (
                os.path.join(os.getcwd(), fn), str(e)))
        en.update_idletasks()
    txt2ppt(line.rstrip() for line in f)
    f.close()

if __name__ == '__main__':
    tk = Tk()
    lb = Label(tk, text='Enter file [or "DEMO"]:')
    lb.pack()
    en = Entry(tk)
    en.bind('<Return>', _start)
    en.pack()
    en.focus_set()
    quit = Button(tk, text='QUIT',
        command=tk.quit, fg='white', bg='red')
    quit.pack(fill='x', expand=True)
    tk.mainloop()

# About 'cStringIO'
# The StringIO and cStringIO modules are gone. Instead, 
# import the io module and use io.StringIO or io.
# BytesIO for text and data respectively.