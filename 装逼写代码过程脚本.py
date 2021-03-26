import time,win32api,win32gui,win32con

if __name__ == '__main__':
    text=mn.get_text()#读取剪辑版文字
    text = [ord(c) for c in text]  # 将字符串转换为整数型列表
    input('已经自动读取了剪辑版的文本（需要打的字都自己先复制好）\n请在回车后，5秒内用鼠标选择好输入的目标位置')
    time.sleep(5)
    point = win32api.GetCursorPos()
    hWndEdit=win32gui.WindowFromPoint(point)
    print(hWndEdit)
    for x in text:  # 依次发送字节串中的每个字节
        win32gui.SendMessage(hWndEdit, win32con.WM_CHAR, x, 0)
        time.sleep(0.01)
    print('完毕！')


