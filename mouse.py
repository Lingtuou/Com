from pynput.mouse import Listener
import xlwt
import time

xls1 = xlwt.Workbook()
sht1 = xls1.add_sheet("Sheet1")
xls2 = xlwt.Workbook()
sht2 = xls2.add_sheet("Sheet1")
xls3 = xlwt.Workbook()
sht3 = xls3.add_sheet("Sheet1")

a = 0
b = 0
c = 0

def on_move(x, y):
    global a
    # 监听鼠标移动
    print('Pointer moved to {0}'.format((x, y)))
    now = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    print (now)
    sht1.write(a,0,now)
    sht1.write(a,1,'Pointer moved to {0}'.format((x, y)))
    a += 1
    xls1.save(r'D:\Data\Move.xls')

def on_click(x, y, button, pressed):
    # 监听鼠标点击
    global b
    print('{0} at {1}'.format('Pressed' if pressed else 'Released', (x, y)))
    now = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    print (now)
    sht2.write(b,0,now)
    sht2.write(b,1,'{0} at {1}'.format('Pressed' if pressed else 'Released', (x, y)))
    b += 1
    xls2.save(r'D:\Data\Clicks.xls')


def on_scroll(x, y, dx, dy):
    # 监听鼠标滚轮
    global c
    print('Scrolled {0}'.format((x, y)))
    now = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    print (now)
    sht3.write(c,0,now)
    sht3.write(c,1,'Scrolled {0}'.format((x, y)))
    c += 1
    xls3.save(r'D:\Data\Scrolled.xls')


# 连接事件以及释放
with Listener(on_move=on_move, on_click=on_click, on_scroll=on_scroll) as listener:
    listener.join()

