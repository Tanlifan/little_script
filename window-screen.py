# -*- coding: utf-8 -*-

"""
须知0:
本程序需要安装以下模块,如果没有安装,请使用pip install进行安装
numpy
pywin32
PyQt5 这个包截图比较快,如果安装了opencv,可以使用opencv进行截图,但是速度会慢一些
paddleocr
paddle 参考须知1安装

须知1:
paddle需要运行以下命令进行安装,如果下载速度慢,可以下载对应的whl文件,然后使用pip install进行安装
python -m pip install paddlepaddle-gpu==2.2.2.post111 -f https://www.paddlepaddle.org.cn/whl/windows/mkl/avx/stable.html
须知2:
如果提示: If this call came from a _pb2.py file, your generated code is out of date and must be regenerated with protoc >= 3.19.0
可以运行下面的命令降级
pip install --upgrade protobuf==3.20.1
"""

import sys   # 需要安装sys模块, 用于退出程序
import time   # 需要安装time模块, 用于获取时间戳
import numpy    # 需要安装numpy模块, 用于图像转换
import winsound # 需要安装winsound模块, 用于播放声音
import win32gui # 需要安装pywin32模块, 用于获取窗口句柄
from PyQt5.QtWidgets import QApplication # 需要安装PyQt5模块, 用于截图
from PyQt5.QtGui import qRed, qGreen, qBlue # 需要安装PyQt5模块, 用于截图
from paddleocr import PaddleOCR # 需要安装paddleocr模块, 用于识别

# 获取窗口句柄
def get_all_hwnd(hwnd, mouse):

    if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
        hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})

# 图像转换
def image_to_cv(img):
    
    tmp = img
    
    #使用numpy创建空的图象
    cv_image = numpy.zeros((tmp.height(), tmp.width(), 3), dtype=numpy.uint8)
    
    # 范围
    x1 = 0
    y1 = 0
    x2 = tmp.width()
    y2 = tmp.height()
    for row in range(y1, y2):
        for col in range(x1, x2):
            r = qRed(tmp.pixel(col, row))
            g = qGreen(tmp.pixel(col, row))
            b = qBlue(tmp.pixel(col, row))
            cv_image[row,col,0] = b
            cv_image[row,col,1] = g
            cv_image[row,col,2] = r
    
    return cv_image

if __name__ == '__main__':

    # OCR识别
    # use_gpu = True, 表示使用GPU，如果没有GPU，可以设置为False，使用CPU
    # show_log = False, 表示不显示预测过程中的log信息
    # lang = 'ch', 表示使用中文模型, 默认为英文模型
    engine = PaddleOCR(use_gpu=True, show_log=False, lang='ch')

    hwnd_title = dict()

    win32gui.EnumWindows(get_all_hwnd, 0)
    for item in hwnd_title.items():
        # 打印句柄信息
        print("{} - {}".format(item[0], item[1]))

    hwnd = None
    for h, t in hwnd_title.items():
        if t != "" and t == '王庆波':
            hwnd = h
            break


    if hwnd == None:
        print('未找到窗口')
        sys.exit(0)

    app = QApplication(sys.argv)
    screen = QApplication.primaryScreen()

    results = ['','','','','','','','','','']
    index = 0

    while True:
        # 间隔0.3秒截图一次
        time.sleep(0.3)
        img = screen.grabWindow(hwnd).toImage()
        # 保存图片
        img.save('test.png')
        img = image_to_cv(img)
        # start time
        start = time.time()
        result = engine(img)
        # end time
        end = time.time()
        # print('耗时：', end - start)

        infos = result[1]
        for info in infos:
            text = info[0]
            if ('好' in text) and text not in results:
                print(text)
                winsound.Beep(600,500)
                results[index] = text
                index = index + 1
                if index == 10:
                    index = 0
        
        # 五分钟打印一次， 时间戳
        if int(time.time()) % 300 == 0:
            print('运行中...' + time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

