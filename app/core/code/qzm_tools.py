# coding: utf-8
import aircv as ac
import traceback
import os
import time


import cv2
import numpy as np
import win32api, win32con, win32gui, win32ui
from win32api import GetSystemMetrics

SCREENIMG = 'img'


class FindTimeOutError(Exception):
    pass


class FindImgError(Exception):
    pass


class OpenImgError(Exception):
    pass


class ImgFinder(object):
    """识别图像位置
    	:params imsrc: 原始图片名, 默认保存在img文件夹中
    	:params imsch: 目标图片名 
        1.方法window_capture,截全屏并且保存到文件夹中,默认是img文件夹
        2.方法get_position, 返回目标在原图的中位置(只识别唯一的一个)
	"""
    def __init__(self, imsch, imsrc='screen.png'):
        try:
            self.element_name = imsch
            self.imsrc = ac.imread(os.path.join(SCREENIMG, imsrc))
            self.imsch = ac.imread(os.path.join(SCREENIMG, imsch))
        except:
            traceback.print_exc()
            raise OpenImgError('打开图片失败')

    @staticmethod
    def window_capture():
        hwnd = 0 # 窗口的编号，0号表示当前活跃窗口
        # 根据窗口句柄获取窗口的设备上下文DC（Divice Context）
        hwndDC = win32gui.GetWindowDC(hwnd)
        # 根据窗口的DC获取mfcDC
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        # mfcDC创建可兼容的DC
        saveDC = mfcDC.CreateCompatibleDC()
        # 创建bigmap准备保存图片
        saveBitMap = win32ui.CreateBitmap()
        # 获取监控器信息
        MoniterDev = win32api.EnumDisplayMonitors(None, None)
        w = MoniterDev[0][2][2]
        h = MoniterDev[0][2][3]
        # print w,h　　　#图片大小
        # 为bitmap开辟空间
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
        # 高度saveDC，将截图保存到saveBitmap中
        saveDC.SelectObject(saveBitMap)
        # 截取从左上角（0，0）长宽为（w，h）的图片
        saveDC.BitBlt((0, 0), (w, h), mfcDC, (0, 0), win32con.SRCCOPY)
        saveBitMap.SaveBitmapFile(saveDC, os.path.join(SCREENIMG, 'screen.png'))
        # 销毁占用的内存
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(0, hwndDC)
        time.sleep(0.2)

    def get_all_element(self, _confidence=0.85):
        """
        返回所有识别准确的图片,的横坐标和纵坐标
        :return: [(int(x1), int(y1)), (int(x2), int(y2))...]
        """
        real_result = ac.find_all_template(self.imsrc, self.imsch, threshold=_confidence)
        print(self.element_name)
        print(real_result)
        if real_result:
            # 结构化数据，返回结果集[(int(x1), int(y1)), (int(x2), int(y2))...
            element_positions = []
            for x in real_result:
                element_positions.append(tuple(map(int, x['result'])))
            return element_positions
        traceback.print_exc()
        raise FindImgError('无法识别图片')

    def get_element(self, dirction=1, index=0, _confidence=0.85):
        """
        返回符合条件的从左到右指定图片坐标
        :param dirction: 以水平方向找，还是垂直方向。0表水平，1表垂直
        :param index: 从左到右第几个, index可以是负数
        :return: (int(x), int(y))
        """
        all_element = self.get_all_element(_confidence)
        sorted_result = sorted(all_element, key=lambda x: x[dirction])
        return sorted_result[index]

def until_get_element(element, dirction=1, index=0, _confidence=0.85, fun=None, attempts=50, f_args=None):
    for i in range(attempts):
        ImgFinder.window_capture()
        btnfinder = ImgFinder(element)
        try:
            result = btnfinder.get_element(dirction, index, _confidence=_confidence)
        except FindImgError:
            if fun:
                if f_args and isinstance(f_args, list):
                    fun(*f_args)
                else:
                    fun()
            else:
                pass
        else:
            return result
    else:
        return None


def until_dis_element(element, fun=None, attempts=50, _confidence=0.85):
    for i in range(attempts):
        ImgFinder.window_capture()
        btnfinder = ImgFinder(element)
        try:
            result = btnfinder.get_element(_confidence=_confidence)
        except FindImgError:
            return True
        else:
            if fun:
                fun()
            else:
                pass
    else:
        return False


def until_move_element(element, before_position, attempts=50):
    for i in range(attempts):
        ImgFinder.window_capture()
        btnfinder = ImgFinder(element)
        try:
            result = btnfinder.get_element()
        except FindImgError:
            return True
        else:
            if before_position != result:
                return True
    else:
        return False



def get_last_screen_gray():
    screen = ac.imread(os.path.join(SCREENIMG, 'screen.png'))
    im_gray = cv2.cvtColor(screen, cv2.COLOR_BGR2GRAY)
    im_binary = cv2.threshold(im_gray, 180, 255, cv2.THRESH_BINARY)
    height, weigh = im_gray.shape
    sum = 0
    for i in im_binary[1]:
        sum += np.sum(i)
    return sum / (height * weigh)


def until_stable_window(before_average_gray, attempts=99999):
    for i in range(attempts):
        ImgFinder.window_capture()
        now_average_gray = get_last_screen_gray()
        print(before_average_gray, now_average_gray)
        if abs(now_average_gray - before_average_gray) < 1.2:
            ImgFinder.window_capture()
            return True
    else:
        return False


def until_change_window(before_average_gray, attempts=20):
    from collections import deque
    screen_gray_history = deque(maxlen=3)
    for i in range(attempts):
        ImgFinder.window_capture()
        now_average_gray = get_last_screen_gray()
        screen_gray_history.append(now_average_gray)
        print(before_average_gray, now_average_gray)
        if abs(now_average_gray - before_average_gray) > 1.2:
            return screen_gray_history
    else:
        return None
