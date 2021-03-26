import win32gui,win32con,os,time,win32ui
import win32clipboard as clipboard
from PIL import Image
from moni import Mouse_And_Key
from io import BytesIO
mn=Mouse_And_Key()#创建模拟类
def pic_ctrl_c(img):
    #将图片置入剪辑版
    output = BytesIO()#如是StringIO分引起TypeError: string argument expected, got 'bytes'
    img.convert("RGB").save(output, "BMP")# 以BMP格式保存流
    data = output.getvalue()[14:]#bmp文件头14个字节丢弃
    output.close()
    clipboard.OpenClipboard() #打开剪贴板
    clipboard.EmptyClipboard() #先清空剪贴板
    clipboard.SetClipboardData(win32con.CF_DIB, data)  #将图片放入剪贴板
    clipboard.CloseClipboard()
def getMessage(name):
    print(name+'有消息，正在控制微信通知伙伴！')
    #弹出微信
    os.startfile(r'C:\Users\Public\Desktop\微信.lnk')
    time.sleep(1)
    #找到微信窗口
    hwnd_wx = win32gui.FindWindow('WeChatMainWndForPC', '微信')

    # 得到目标窗口的矩阵信息
    rect = win32gui.GetWindowRect(hwnd_wx )
    x, y, w, h = rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]
    print(x,y)
    # 点击搜索框
    time.sleep(0.2)
    mn.mouse_left_click(135+x ,37+y, 2)#双击
    #输入文字
    time.sleep(0.2)
    mn.set_text(qunName)#设置剪辑版文本
    mn.set_window_foregGroun(hwnd_wx)
    mn.key_even_zuhe(86,17)#粘贴文本
    time.sleep(1)
    mn.key_even(13)#回车
    time.sleep(0.2)
    #粘贴图片消息
    pic_ctrl_c(get_img_message(hwnd))
    time.sleep(0.2)
    mn.set_window_foregGroun(hwnd_wx)
    mn.key_even_zuhe(86, 17)  # 粘贴图片
    time.sleep(0.5)
    mn.set_text(name)  # 设置剪辑版文本
    time.sleep(0.2)
    mn.set_window_foregGroun(hwnd_wx)
    mn.key_even_zuhe(86, 17)  # 粘贴文本
    time.sleep(0.2)
    #回车发送
    mn.key_even(13)
    print(name + '发送成功！')
def get_img_message(hWnd,):
    # 获取句柄窗口的大小信息
    left, top, right, bot = win32gui.GetWindowRect(hWnd)
    width = right - left
    height = bot - top
    # 返回句柄窗口的设备环境，覆盖整个窗口，包括非客户区，标题栏，菜单，边框
    hWndDC = win32gui.GetWindowDC(hWnd)
    # 创建设备描述表
    mfcDC = win32ui.CreateDCFromHandle(hWndDC)
    # 创建内存设备描述表
    saveDC = mfcDC.CreateCompatibleDC()
    # 创建位图对象准备保存图片
    saveBitMap = win32ui.CreateBitmap()
    # 为bitmap开辟存储空间
    saveBitMap.CreateCompatibleBitmap(mfcDC, width, height)
    # 将截图保存到saveBitMap中
    saveDC.SelectObject(saveBitMap)
    # 保存bitmap到内存设备描述表
    saveDC.BitBlt((0, 0), (width, height), mfcDC, (0, 0), win32con.SRCCOPY)
    #PIL保存
    ###获取位图信息
    bmpinfo = saveBitMap.GetInfo()
    bmpstr = saveBitMap.GetBitmapBits(True)
    ###生成图像
    im_PIL = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'BGRX', 0, 1)
    return im_PIL
title='MsgStrengthenRemindDlg'
type_='MsgRemindDlg'
haveToReply=False#判断是否需要回复的标识
#获取群名字和自己的名字
with open('setting.txt','r',encoding='utf-8')as f:
    text=f.read()
    text=text.split('\n')
    qunName=text[0].split('=')[-1]
    name = text[1].split('=')[-1]
    print(qunName,name)


while True:
    time.sleep(1)
    #寻找消息窗口，并且返回句柄'
    hwnd=win32gui.FindWindow(type_,title)
    if hwnd==0:#如果没找到说明不需要判断
        pass
    else:

        #判断消息窗口是否弹出
        isVisible=win32gui.IsWindowVisible(hwnd)
        if isVisible == 0:
            haveToReply = False
        if haveToReply == False:#如果没有发送消息
            if isVisible != 0:
                #将消息窗口内容截图
                pic_ctrl_c(get_img_message(hwnd))
                getMessage(name)
                haveToReply = True









