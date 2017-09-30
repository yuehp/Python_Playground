# install Active.Tck, in Python 3.x first 't' in tkinter is in lowercase
from tkinter import Tk
from time import sleep
# tkMessageBox in Python2.x becomes  tkinter.messagebox in Python3.x
from tkinter.messagebox import showwarning
# install PythonWin
import win32com.client as win32

warn = lambda app: showwarning(app, 'Exit?')
RANGE = range(3, 8)

def excel():
    app = 'Excel'
    # 启动 Excel
    xl = win32.gencache.EnsureDispatch('%s.Application' % app)
    # 添加一个工作簿
    ss = xl.Workbooks.Add()
    # 取得活动工作表
    sh = ss.ActiveSheet
    # 使工作簿可见
    xl.Visible = True
    sleep(1)

    # 在第一个单元格写入标题
    sh.Cells(1,1).Value = 'Python-to-%s Demo' % app
    sleep(1)
    # 在第1列的3~7行，写入内容
    for i in RANGE:
        sh.Cells(i,1).Value = 'Line %d' % i
        sleep(1)
    sh.Cells(i+2,1).Value = "Th-th-th-that's all floks!"

    warn(app)
    # 工作簿退出，不保存
    ss.Close(False)
    xl.Application.Quit()

if __name__ == '__main__':
    Tk().withdraw()
    excel()