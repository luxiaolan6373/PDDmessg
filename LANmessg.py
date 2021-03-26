from socket import *#
from PIL import Image#图片处理库
from playsound import playsound#播放音乐库
from Orc_baidu import Baiduorc
import win32gui,win32con,os,time,win32ui,io,winsound
import ctypes
def get_img_message(hWnd):
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
    #转成字节集
    imgByteArr = io.BytesIO()
    im_PIL.save(imgByteArr, format='png')
    imgByteArr = imgByteArr.getvalue()
    return imgByteArr
def udp_send(udp_socket,messg,ip):
    try:
        sen_messg=''
        for item in messg['words_result']:
            sen_messg=sen_messg+item['words']+'\n'
        # 接收用户输入端口号
        port = 8888
        # 发送消息 内容进行编码
        udp_socket.sendto(sen_messg.encode("gbk"), (ip, port))
        print('消息发送成功！',sen_messg)
    except Exception as err:
        print(err)
        # 接收用户输入端口号
        port = 8888
        # 发送消息 内容进行编码
        sen_messg = '获取消息内容失败！'
        udp_socket.sendto(sen_messg.encode("gbk"), (ip, port))
        print('消息发送成功！但是内容失败！')
def udp_recvfrom(udp_socket):
    # 接收消息 最多4096个字节
    get_mes, get_ip = udp_socket.recvfrom(4096)
    print("收到来自%s的消息:%s" % (gethostbyaddr(get_ip[0])[0]+str(get_ip), get_mes.decode("gbk")))
    playsound('messg.mp3')
    time.sleep(0.5)
def main():
    # 创建套接字
    udp_socket = socket(AF_INET, SOCK_DGRAM)
    # 设置固定端口
    udp_socket.bind(("", 8888))
    print("*" * 50)
    print("----------拼多多值班消息共享器（建议将伙伴的电脑名修改为自己能认识的名字）纯模拟无入侵操作----------")
    print("1.自动转发消息（下班休息的伙伴，选择这个）")
    print("2.我是值班人员（值班熬夜的兄弟，选择这个）")
    print("*" * 50)
    user = input("请输入要执行的操作序号:")
    if user == "1":
        ip=input('请输入值班人员的ip地址:')
        # 解决控制台卡顿的问题，关闭快速编辑模式
        kernel32 = ctypes.windll.kernel32
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), 128)
        #初始化百度识字对象
        bd = Baiduorc(AK="Weui83nXBM1ox6ozFzPF9bng", SK="4AEmIGiInMv7gzlATntAs3pjHZrlCrsK")
        title = 'MsgStrengthenRemindDlg'
        type_ = 'MsgRemindDlg'
        haveToReply = False  # 判断这条消息是否已经是提醒过了的
        while True:
            time.sleep(1)
            # 寻找消息窗口，并且返回句柄'
            hwnd = win32gui.FindWindow(type_, title)
            if hwnd == 0:  # 如果没找到说明不需要判断
                pass
            else:
                # 判断消息窗口是否弹出
                isVisible = win32gui.IsWindowVisible(hwnd)
                if isVisible == 0:#消息窗口不见了，说明没消息
                    haveToReply = False
                if haveToReply == False:  # 如果没有准备发送消息
                    if isVisible != 0:#消息来了
                        # 等待5秒，PDD平台自带的机器人有些问题可以自动回，所以等5秒先,如果窗口还在说明还没回复
                        time.sleep(5)
                        isVisible = win32gui.IsWindowVisible(hwnd)
                        if isVisible != 0:
                            # 将消息窗口内容截图，并且调用识字，去识别内容
                            # 有些识别不了也没关系，咱们程序的目的是提醒值班人员有客户来了
                            messg=bd.get_text(get_img_message(hwnd))
                            #给值班人员的电脑发送局域网消息
                            udp_send(udp_socket,messg,ip)
                            #标识这条消息已经提醒过了，就不要重复的提醒了，等窗口不见了再来重置这个开关
                            haveToReply = True
    elif user == "2":
        # 解决控制台卡顿的问题，关闭快速编辑模式
        kernel32 = ctypes.windll.kernel32
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), 128)
        print('您的ip地址为：', gethostbyname(gethostname()))
        print('正在监听其他伙伴ip的消息...')
        while True:
            udp_recvfrom(udp_socket)
    else:
        input("输入有误!")
    # 关闭套接字
    udp_socket.close()
if __name__ == "__main__":
    main()
