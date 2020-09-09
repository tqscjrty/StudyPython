import win32api
import win32con
import win32gui
import win32ui
import win32clipboard as w
import os

'''
hWnd控件句柄
filename保存图片的文件路径及文件名
'''

def snapJpg(hWnd,filename):
    #获取句柄窗口的大小信息
    left, top, right, bot = win32gui.GetWindowRect(hWnd)
    width = right - left
    height = bot - top
    #返回句柄窗口的设备环境，覆盖整个窗口，包括非客户区，标题栏，菜单，边框
    hWndDC = win32gui.GetWindowDC(hWnd)
    #创建设备描述表
    mfcDC = win32ui.CreateDCFromHandle(hWndDC)
    #创建内存设备描述表
    saveDC = mfcDC.CreateCompatibleDC()
    #创建位图对象准备保存图片
    saveBitMap = win32ui.CreateBitmap()
    #为bitmap开辟存储空间
    saveBitMap.CreateCompatibleBitmap(mfcDC,width,height)
    #将截图保存到saveBitMap中
    saveDC.SelectObject(saveBitMap)
    #保存bitmap到内存设备描述表
    saveDC.BitBlt((0,0), (width,height), mfcDC, (0, 0), win32con.SRCCOPY)
    #如果要截图到打印设备：
    ###最后一个int参数：0-保存整个窗口，1-只保存客户区。如果PrintWindow成功函数返回值为1
    #result = windll.user32.PrintWindow(hWnd,saveDC.GetSafeHdc(),0)
    #print(result) #PrintWindow成功则输出1
    #保存图像
    bmpinfo = saveBitMap.GetInfo()
    print("bmpinfo=",bmpinfo)
    bmpstr = saveBitMap.GetBitmapBits(True)
    ###生成图像
    saveBitMap.SaveBitmapFile(saveDC,filename)
    #内存释放
    win32gui.DeleteObject(saveBitMap.GetHandle())
    saveDC.DeleteDC()
    mfcDC.DeleteDC()
    win32gui.ReleaseDC(hWnd,hWndDC)  


'''
取得控件的文本
hwnd为控件的父句柄
ID为控件在父控件中序号ID

'''

def GetTextData(hwnd,ID):
    #传入父句柄，控件ID
    hwnd=win32gui.GetDlgItem(hwnd,ID)
    buf_size = win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0) + 1  # 要加上截尾的字节  
    str_buffer = win32gui.PyMakeBuffer(buf_size)  # 生成buffer对象  
    win32api.SendMessage(hwnd, win32con.WM_GETTEXT, buf_size, str_buffer)  # 获取buffer  
    address, length = win32gui.PyGetBufferAddressAndLen(str_buffer) 
    text = win32gui.PyGetString(address, length-1)
    return text

'''
设置控件的文本
hwnd为控件的父句柄
ID为控件在父控件中序号ID
text为要设置的文本
'''

def SetTextData(hwnd,ID,text):
    #传入父句柄，控件ID,text为字符串
    handle=win32gui.GetDlgItem(hwnd,ID)
    win32api.SendMessage(handle, win32con.WM_SETTEXT, 0, text)

def _MyCallback( hwnd, extra ):
    windows = extra
    temp=[]
    if win32gui.GetClassName(hwnd)=="#32770" and win32gui.GetWindowText(hwnd)=="" :
        temp.append(hwnd)
        temp.append(win32gui.GetClassName(hwnd))
        temp.append(win32gui.GetWindowText(hwnd))
        windows[hwnd] = temp
def start(hwnd):
    #处理选择选择题
    hwnd1=win32gui.FindWindowEx(hwnd,0,u"AfxOleControl42s",None)
    hwnd2=win32gui.FindWindowEx(hwnd1,0,u"AfxWnd42s",None)
    hedit=win32gui.FindWindowEx(hwnd2,0,u"RichEdit20A",None)
    print("hwnd=%x",hwnd,"hwnd1=%x"%hwnd1,"hwnd2=%x"%hwnd2,"hedit=%x"%hedit)

    #美术单选题
    for i in range(0,17) :
        win32api.SendMessage(hwnd, 0x464, i, 0)
        win32api.Sleep(50)
        filename="D:\EXAM\美术单选题"+str(i)+'.bmp'
        snapJpg(hwnd1,filename)

    #美术多选题
    for i in range(0,6) :
        win32api.SendMessage(hwnd, 0x464, i, 1)
        win32api.Sleep(50)
        filename="D:\EXAM\美术多选题"+str(i)+'.bmp'
        snapJpg(hwnd1,filename)
        
    #美术判断题    
    for i in range(0,15) :
        win32api.SendMessage(hwnd, 0x464, i, 2)
        win32api.Sleep(50)
        filename="D:\EXAM\美术判断题"+str(i)+'.bmp'
        snapJpg(hwnd1,filename)

    #音乐单选题
    for i in range(0,10) :
        win32api.SendMessage(hwnd, 0x464, i, 3)
        win32api.Sleep(50)
        filename="D:\EXAM\音乐单选题"+str(i)+'.bmp'
        snapJpg(hwnd1,filename)

    #音乐多选题
    for i in range(0,5) :
        win32api.SendMessage(hwnd, 0x464, i, 4)
        win32api.Sleep(50)
        filename="D:\EXAM\音乐多选题"+str(i)+'.bmp'
        snapJpg(hwnd1,filename)
        
    #音乐判断题    
    for i in range(0,15) :
        win32api.SendMessage(hwnd, 0x464, i, 5)
        win32api.Sleep(50)
        filename="D:\EXAM\音乐判断题"+str(i)+'.bmp'
        snapJpg(hwnd1,filename)
  
def TestEnumWindows():
    windows = {}
    win32gui.EnumWindows(_MyCallback, windows)
    for item in windows :
        hwnd=windows[item][0]
        try:
            hdlg=win32gui.GetDlgItem(hwnd,0x516)
            if hdlg:                
                hwnd=windows[item][0]
                print  (hex(windows[item][0]))
                start(hwnd)
        except:
             print('')
             
win32api.Sleep(3000)
TestEnumWindows()

