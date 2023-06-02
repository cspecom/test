from __future__ import print_function
import os
import random
import sys
import time
from random import randint
import subprocess
from subprocess import check_output
import numpy as np
import pyautogui
import pywinauto
import pywinauto.mouse as mouse
import requests
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
import re
from io import StringIO
import contextlib

dirname = os.path.dirname(__file__)
file_dir = os.path.join(dirname, 'assets\\')


# win32_py = os.path.join(dirname, 'venv\\lib\\site-packages\\tzlocal\\win32.py')
# print("Repair file: " + win32_py)

def get_pid(name):
    return check_output(["pidof", name])


class ScreenRes(object):
    @classmethod
    def set(cls, width: object = None, height: object = None, depth: object = 32) -> object:
        '''
        Set the primary display to the specified mode
        '''
        if width and height:
            print('Setting resolution to {}x{}'.format(width, height, depth))
        else:
            print('Setting resolution to defaults')

        if sys.platform == 'win32':
            cls._win32_set(width, height, depth)
        elif sys.platform.startswith('linux'):
            cls._linux_set(width, height, depth)
        elif sys.platform.startswith('darwin'):
            cls._osx_set(width, height, depth)

    @classmethod
    def get(cls):
        if sys.platform == 'win32':
            return cls._win32_get()
        elif sys.platform.startswith('linux'):
            return cls._linux_get()
        elif sys.platform.startswith('darwin'):
            return cls._osx_get()

    @classmethod
    def get_modes(cls):
        if sys.platform == 'win32':
            return cls._win32_get_modes()
        elif sys.platform.startswith('linux'):
            return cls._linux_get_modes()
        elif sys.platform.startswith('darwin'):
            return cls._osx_get_modes()

    @staticmethod
    def _win32_get_modes():
        '''
        Get the primary windows display width and height
        '''
        import win32api
        from pywintypes import error
        modes = []
        i = 0
        try:
            while True:
                mode = win32api.EnumDisplaySettings(None, i)
                modes.append((
                    int(mode.PelsWidth),
                    int(mode.PelsHeight),
                    int(mode.BitsPerPel),
                ))
                i += 1
        except error:
            pass

        return modes

    @staticmethod
    def _win32_get():
        '''
        Get the primary windows display width and height
        '''
        import ctypes
        user32 = ctypes.windll.user32
        screensize = (
            user32.GetSystemMetrics(0),
            user32.GetSystemMetrics(1),
        )
        return screensize

    @staticmethod
    def _win32_set(width=None, height=None, depth=32):
        '''
        Set the primary windows display to the specified mode
        '''
        # Gave up on ctypes, the struct is really complicated
        # user32.ChangeDisplaySettingsW(None, 0)
        import win32api
        if width and height:
            if not depth:
                depth = 32
            mode = win32api.EnumDisplaySettings()
            mode.PelsWidth = width
            mode.PelsHeight = height
            mode.BitsPerPel = depth

            win32api.ChangeDisplaySettings(mode, 0)
        else:
            win32api.ChangeDisplaySettings(None, 0)

    @staticmethod
    def _win32_set_default():
        '''
        Reset the primary windows display to the default mode
        '''
        # Interesting since it doesn't depend on pywin32
        import ctypes
        user32 = ctypes.windll.user32
        # set screen size
        user32.ChangeDisplaySettingsW(None, 0)

    @staticmethod
    def _linux_set(width=None, height=None, depth=32):
        raise NotImplementedError()

    @staticmethod
    def _linux_get():
        raise NotImplementedError()

    @staticmethod
    def _linux_get_modes():
        raise NotImplementedError()

    @staticmethod
    def _osx_set(width=None, height=None, depth=32):
        raise NotImplementedError()

    @staticmethod
    def _osx_get():
        raise NotImplementedError()

    @staticmethod
    def _osx_get_modes():
        raise NotImplementedError()


def type(string):
    for x in string:
        time.sleep(random.uniform(0.09, 0.2))
        pyautogui.keyDown(x)
        time.sleep(random.uniform(0.09, 0.2))
        pyautogui.keyUp(x)


class person:
    _rd_first_name = ''
    _rd_last_name = ''
    _rd_user_name = ''
    _rd_password = ''
    _rd_mail_recover = ''
    _rd_day = ''
    _rd_month = ''
    _rd_year = ''
    _rd_gender = 0

    def __init__(self, _rd_first_name, _rd_last_name, _rd_user_name, _rd_password, _rd_mail_recover, _rd_day, _rd_month,
                 _rd_year, _rd_gender):
        self._rd_first_name = _rd_first_name
        self._rd_last_name = _rd_last_name
        self._rd_user_name = _rd_user_name
        self._rd_password = _rd_password
        self._rd_mail_recover = _rd_mail_recover
        self._rd_day = _rd_day
        self._rd_month = _rd_month
        self._rd_year = _rd_year
        self._rd_gender = _rd_gender


_arr_gmail_infor = []


def gen_infor():
    f_male_first_name = None
    f_female_first_name = None
    f_middle_name = None
    f_last_name = None
    f_adjective = None
    f_animals = None
    f_nationa = None
    f_nouns = None
    f_profesional = None

    ran_mail = randint(3, 7)
    # 2000
    for x in range(5):
        _rd_first_name = ''
        _rd_male_first_name = ''
        _rd_female_first_name = ''
        _rd_middle_name = ''
        _rd_last_name = ''
        _rd_user_name = ''
        _rd_password = ''
        _rd_mail_recover = ''
        _rd_day = str(randint(2, 27))
        _rd_month = str(randint(2, 11))
        _rd_year = str(randint(1970, 2001))
        _rd_gender = randint(1, 2)  # 1 male - 2 female
        _arr_str = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't',
                    'u', 'v', 'w', 'x', 'y', 'z']
        # _middle_name = random.shuffle(_arr_str)

        _rd_adjective = ''
        _rd_animals = ''
        _rd_nationa = ''
        _rd_nouns = ''
        _rd_profesional = ''
        _rd_extm = ''

        _extmrc = ''
        _arr_mail_reco = [1, 1, 1, 1, 2, 2, 2, 2, 2, 2]

        # get male first name
        _ran_index_of_male_first_name = randint(1, 3023)
        _index_get_male_first_name = 0
        f_male_first_name = open(file_dir + "en_male_first_name.txt", "r")
        for x in f_male_first_name:
            _index_get_male_first_name = _index_get_male_first_name + 1
            # last_name = last_name + " " + x.strip('\n').split(" ")[ii]
            if _index_get_male_first_name == _ran_index_of_male_first_name:
                _rd_male_first_name = x.strip('\n')
                break
        # f_male_first_name.close()

        # get female first name
        _ran_index_of_female_first_name = randint(1, 5045)
        _index_get_female_first_name = 0
        f_female_first_name = open(file_dir + "en_female_first_name.txt", "r")
        for x in f_female_first_name:
            _index_get_female_first_name = _index_get_female_first_name + 1
            if _index_get_female_first_name == _ran_index_of_female_first_name:
                _rd_female_first_name = x.strip('\n')
                break
        # f_female_first_name.close()

        # get middle_name
        _ran_index_of_middle_name = randint(1, 88798)
        _index_get_middle_name = 0
        f_middle_name = open(file_dir + "en_middle_name.txt", "r")
        for x in f_middle_name:
            _index_get_middle_name = _index_get_middle_name + 1
            if _index_get_middle_name == _ran_index_of_middle_name:
                _rd_middle_name = x.strip('\n')
                break
        # f_middle_name.close()

        # get last name
        _ran_index_of_lastname = randint(1, 88700)
        _index_get_latname = 0
        f_last_name = open(file_dir + "en_last_name.txt", "r")
        for x in f_last_name:
            _index_get_latname = _index_get_latname + 1
            if _index_get_latname == _ran_index_of_lastname:
                _rd_last_name = x.strip('\n')
                break
        # f_last_name.close()

        # get adjective
        _ran_index_of_adjective = randint(1, 4840)
        _index_get_adjective = 0
        f_adjective = open(file_dir + "en_adjective.txt", "r")
        for x in f_adjective:
            _index_get_adjective = _index_get_adjective + 1
            if _index_get_adjective == _ran_index_of_adjective:
                _rd_adjective = x.strip('\n')
                break
        # f_adjective.close()

        # get animals
        _ran_index_of_animals = randint(1, 1158)
        _index_get_animals = 0
        f_animals = open(file_dir + "en_animals.txt", "r")
        for x in f_animals:
            _index_get_animals = _index_get_animals + 1
            if _index_get_animals == _ran_index_of_animals:
                _rd_animals = x.strip('\n')
                break
        # f_animals.close()

        # get nationa
        _ran_index_of_nationa = randint(1, 185)
        _index_get_nationa = 0
        f_nationa = open(file_dir + "en_nationa.txt", "r")
        for x in f_nationa:
            _index_get_nationa = _index_get_nationa + 1
            if _index_get_nationa == _ran_index_of_nationa:
                _rd_nationa = x.strip('\n')
                break
        # f_nationa.close()

        # get nouns
        _ran_index_of_nouns = randint(1, 5915)
        _index_get_nouns = 0
        f_nouns = open(file_dir + "en_nouns.txt", "r")
        for x in f_nouns:
            _index_get_nouns = _index_get_nouns + 1
            if _index_get_nouns == _ran_index_of_nouns:
                _rd_nouns = x.strip('\n')
                break
        # f_nouns.close()

        # get profesional
        _ran_index_of_profesional = randint(1, 4840)
        _index_get_profesional = 0
        f_profesional = open(file_dir + "en_profesional.txt", "r")
        for x in f_profesional:
            _index_get_profesional = _index_get_profesional + 1
            if _index_get_profesional == _ran_index_of_profesional:
                _rd_profesional = x.strip('\n')
                break
        # f_profesional.close()

        # get extm
        _ran_index_of_extm = randint(1, 10375)
        _index_get_extm = 0
        f_extm = open(file_dir + "extm.txt", "r")
        for x in f_extm:
            _index_get_extm = _index_get_extm + 1
            if _index_get_extm == _ran_index_of_extm:
                _rd_extm = x.strip('\n')
                break
        # f_profesional.close()

        # style male_first_name
        _rd_choose_male_first_name_style = randint(0, 3)
        if _rd_choose_male_first_name_style == 0:
            try:
                _rd_male_first_name = _rd_male_first_name[0:randint(3, 7)]
                _rd_female_first_name = _rd_female_first_name[0:randint(3, 7)]
            except:
                print('_rd_male_first_name < 7')

        # style middle_name
        _rd_choose_middle_name_style = randint(0, 3)
        if _rd_choose_middle_name_style == 0:
            try:
                _rd_middle_name = _rd_middle_name[0:randint(3, 7)]
            except:
                print('_rd_last_name < 7')

        # style last_name
        _rd_choose_last_name_style = randint(0, 3)
        if _rd_choose_last_name_style == 0:
            try:
                _rd_last_name = _rd_last_name[0:randint(3, 7)]
            except:
                print('_rd_last_name < 7')

        _arr_random_compornent_user_name = []
        if _rd_gender == 1:  # male
            _rd_first_name = _rd_male_first_name
            _arr_random_compornent_user_name = [
                _rd_male_first_name,
                _rd_middle_name,
                _rd_last_name,
                _rd_adjective,
                _rd_animals,
                _rd_nationa,
                _rd_nouns,
                _rd_profesional
            ]

        if _rd_gender == 2:  # female
            _rd_first_name = _rd_female_first_name
            _arr_random_compornent_user_name = [
                _rd_female_first_name,
                _rd_middle_name,
                _rd_last_name,
                _rd_adjective,
                _rd_animals,
                _rd_nationa,
                _rd_nouns,
                _rd_profesional
            ]

        _arr_birth_day = [
            _rd_day + _rd_month,
            _rd_year
        ]

        _rd_arr_chdb = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
        random.shuffle(_rd_arr_chdb)

        _is_not_number_first = False

        while len(_rd_user_name) < 17 or len(_rd_password) < 12:
            _is_include_number = randint(1, 5)
            if _is_include_number == 1 and len(_rd_user_name) > 4:
                random.shuffle(_arr_birth_day)
                random.shuffle(_rd_arr_chdb)

                _rd_user_name = _rd_user_name + _arr_birth_day[0]
                _rd_mail_recover = _rd_mail_recover + _arr_birth_day[1]
                _rd_password = _rd_password + _arr_birth_day[randint(0, 1)]
                _rd_password = _rd_password + _rd_arr_chdb[0]
            random.shuffle(_arr_random_compornent_user_name)
            _rd_sub_of_cpn_0 = randint(1, 3)
            _rd_length_of_sub = randint(2, 6)
            _rd_cpn_user_name = ''
            if _rd_sub_of_cpn_0 == 1:
                try:
                    _rd_cpn_user_name = _arr_random_compornent_user_name[0]
                    _rd_cpn_user_name = _rd_cpn_user_name[0, _rd_length_of_sub]
                except:
                    _rd_cpn_user_name = _arr_random_compornent_user_name[0]
            _rd_user_name = _rd_user_name + _rd_cpn_user_name
            _rd_mail_recover = _rd_mail_recover + _arr_random_compornent_user_name[1]
            _cpn_pass = _arr_random_compornent_user_name[2]
            try:
                _cpn_pass = _cpn_pass[0:3]
            except:
                _cpn_pass = _rd_user_name[0:2]
            _rd_password = _rd_password + _cpn_pass
            if len(_rd_password) < randint(2, 5):
                _rd_password = _rd_password + _rd_arr_chdb[0]
            _is_not_number_first = True

        if len(_rd_user_name) > 22:
            _rd_user_name = _rd_user_name[0:randint(17, 19)]

        if len(_rd_mail_recover) > 19:
            _rd_mail_recover = _rd_mail_recover[0:randint(13, 15)]

        if len(_rd_password) > 16:
            _rd_password = _rd_password[0:randint(12, 13)]

        while len(_rd_password) < 12:
            _rd_password = _rd_password + _rd_arr_chdb[0] + _rd_day

        _rd_is_custom_password = randint(1, 2)
        if _rd_is_custom_password == 1:
            _rd_password = _rd_password.lower()
            _rd_password = _rd_password[0].upper() + _rd_password[1:len(_rd_password)]

        _rd_user_name = _rd_user_name.lower()
        _rd_mail_recover = _rd_mail_recover.lower() + '@' + _rd_extm

        _arr_gmail_infor.append(
            person(_rd_first_name, _rd_last_name, _rd_user_name, _rd_password, _rd_mail_recover, _rd_day, _rd_month,
                   _rd_year, _rd_gender))

    f_male_first_name.close()
    f_female_first_name.close()
    f_middle_name.close()
    f_last_name.close()
    f_adjective.close()
    f_animals.close()
    f_nationa.close()
    f_nouns.close()
    f_profesional.close()


# generate infors
gen_infor()


def random_choice():
    options = [2, 3, 4, 5, 6]
    weights = [0.1, 0.2, 0.3, 0.3, 0.1]
    choice = random.choices(options, weights=weights)[0]
    index = options.index(choice) + 1
    print("Chọn số thứ", index, ": ", choice)

def append_err(err):
    with open(file_dir + "err.txt", "a") as myfile:
        myfile.write(err)
        myfile.write('\n')
        myfile.write('\n')


def append_mail_to_file(_rd_first_name, _rd_last_name, _rd_user_name, _rd_password, _rd_mail_recover, _rd_birthday):
    with open(file_dir + "data.txt", "a") as myfile:
        myfile.write(_rd_first_name)
        myfile.write('|')
        myfile.write(_rd_last_name)
        myfile.write('|')
        myfile.write(_rd_user_name)
        myfile.write('|')
        myfile.write(_rd_password)
        myfile.write('|')
        myfile.write(_rd_mail_recover)
        myfile.write('|')
        myfile.write(_rd_birthday)
        myfile.write('\n')


def bezier_curve(xa, ya, xb, yb, xc, yc):
    # Tính toán các hệ số của đường thẳng nối hai điểm
    def line(x1, y1, x2, y2):
        a = (y1 - y2) / (x1 - x2)
        b = y1 - a * x1
        return (a, b)

    # Khởi tạo danh sách các điểm trên đường cong
    points = []

    # Lặp qua các giá trị của x từ xa đến xc
    for i in range(0, 10):
        # Tính toán các giá trị của y tương ứng trên hai đường thẳng
        (a1, b1) = line(xa, ya, xb, yb)
        (a2, b2) = line(xb, yb, xc, yc)
        x1 = i * (xb - xa) / 10 + xa
        y1 = a1 * x1 + b1
        x2 = i * (xc - xb) / 10 + xb
        y2 = a2 * x2 + b2

        # Tính toán các hệ số của đường thẳng nối hai điểm trên hai đường thẳng
        (a3, b3) = line(x1, y1, x2, y2)

        # Tính toán giá trị của x và y trên đường thẳng này
        x = i * (x2 - x1) / 10 + x1
        y = a3 * x + b3

        # Thêm điểm này vào danh sách các điểm trên đường cong
        points.append((x, y))
    return points


def move_mouse(points):
    # Di chuyển chuột theo các điểm trên đường cong Bezier
    for point in points:
        mouse.move(coords=(round(point[0]), round(point[1])))
        time.sleep(0.001)


def move_to(x, y):
    # Lấy tọa độ hiện tại của chuột
    current_pos = pyautogui.position()
    a = current_pos[0]
    b = current_pos[1]

    # Sinh ngẫu nhiên tọa độ của điểm trung gian
    c = random.randint(0, 10)
    d = random.randint(0, 10)

    # Tính toán các điểm trên đường cong Bezier từ (a,b) đến (x,y) qua (c,d)
    points = bezier_curve(a, b, c, d, x, y)

    # Di chuyển chuột theo các điểm này
    move_mouse(points)
    # # Click vào điểm đến
    # pywinauto.mouse.click('left', x, y)


def move_to_and_click(x, y):
    move_to(x, y)
    # Click vào điểm đến
    pywinauto.mouse.click('left', x, y)


def mouse_move_to_rad():
    # Lấy kích thước màn hình
    width, height = pyautogui.size()
    # Tạo một số ngẫu nhiên trong phạm vi từ 1/5 đến 4/5 kích thước khung hình
    x = random.randint(int(round(width // 6)), int(round((4 * width) // 5)))
    y = random.randint(int(round(height // 6)), int(round((4 * height) // 5)))
    # Di chuyển chuột đến (x,y)
    move_to(x, y)


# Cuộn chuột xuống x lần, mỗi lần 1 khoảng cách random
# def scroll_down(num_of_scrolls, speed=random.randint(1, 3)):
#     for i in range(num_of_scrolls):
#         pyautogui.scroll(speed * random.randint(-400, -150))
#         time.sleep(random.uniform(1, 6))


def click_random_items_all_type():
    # Lưu title hiện tại của cửa sổ
    current_title = app.top_window().texts()[0]
    attempts = 0
    while attempts < 5:
        listing_options = app.top_window().child_window(title="Listing options selector. List View selected.",
                                                        control_type="Button")
        width, height = pyautogui.size()
        if listing_options.exists():
            rd_x = np.random.randint(int(round(width // 4)), int(round(width // 2)))
            rd_y = np.random.randint(int(round(height // 4)), int(round((4 * height) // 5)))
        else:
            rd_x = np.random.randint(int(round(width // 4)), int(round((4 * width) // 4)))
            rd_y = np.random.randint(int(round(height // 4)), int(round((4 * height) // 4)))

        # Di chuyển chuột đến vị trí random và click vào vị trí đó
        move_to(rd_x, rd_y)
        time.sleep(random.uniform(1, 2))
        pywinauto.mouse.click(button='left', coords=(rd_x, rd_y))
        time.sleep(2)

        # Kiểm tra nếu title đã thay đổi thì đã click vào items
        if app.top_window().texts()[0] != current_title:
            print("Click random items thành công!")
            time.sleep(2)
            break
        else:
            # Increment the number of attempts
            attempts += 1
            # If we've failed to click on an item 10 times, scroll down and try again
            if attempts == 5:
                scroll_down(1)
                time.sleep(2)
                attempts = 0
            else:
                print(f"Click random items không thành công, di chuyển chuột đến vị trí ngẫu nhiên mới và click tiếp")


def click_element_by_auto_id(auto_id, controltype, x_offset=2, y_offset=2):
    # Tìm phần tử có tên là auto_id trong cửa sổ hiện tại của ứng dụng Windows, tìm bằng title.
    element = app.top_window().child_window(auto_id=auto_id, control_type=controltype)

    # Kiểm tra xem phần tử đó có đang hiển thị hay không.
    if element.is_visible():
        # Lấy kích thước và vị trí của phần tử trên màn hình
        element_rect = element.rectangle()
        element_x_left = element_rect.left
        element_x_right = element_rect.right
        element_y_top = element_rect.top
        element_y_bottom = element_rect.bottom

        # Tạo ra một cặp tọa độ ngẫu nhiên trên phần tử và click vào vị trí đó.
        width, height = pyautogui.size()
        button_x = random.randint(element_x_left + x_offset, element_x_right - x_offset)
        button_y = random.randint(element_y_top + y_offset, element_y_bottom - y_offset)
        while button_y > height - 40:
            button_y = random.randint(element_y_top + y_offset, element_y_bottom - y_offset)
        move_to(button_x, button_y)
        time.sleep(random.uniform(0.5, 2))
        pywinauto.mouse.click(button='left', coords=(button_x, button_y))
    else:
        print(f"Phần tử {auto_id} không hiển thị trên màn hình")

    # Dừng chương trình một thời gian ngẫu nhiên từ 1 đến 3 giây để đợi cho phần tử được load hoàn tất.
    time.sleep(random.uniform(1, 3))


def click_element_by_title(element_name, controltype, x_offset=2, y_offset=2):
    # Tìm phần tử có tên là element_name trong cửa sổ hiện tại của ứng dụng Windows, tìm bằng title.
    element = app.top_window().child_window(title=element_name, control_type=controltype)

    # Kiểm tra xem phần tử đó có đang hiển thị hay không.
    if element.is_visible():
        # Lấy kích thước và vị trí của phần tử trên màn hình
        element_rect = element.rectangle()
        element_x_left = element_rect.left
        element_x_right = element_rect.right
        element_y_top = element_rect.top
        element_y_bottom = element_rect.bottom

        # Tạo ra một cặp tọa độ ngẫu nhiên trên phần tử và click vào vị trí đó.
        width, height = pyautogui.size()
        button_x = random.randint(element_x_left + x_offset, element_x_right - x_offset)
        button_y = random.randint(element_y_top + y_offset, element_y_bottom - y_offset)
        while button_y > height - 40:
            button_y = random.randint(element_y_top + y_offset, element_y_bottom - y_offset)
        move_to(button_x, button_y)
        time.sleep(random.uniform(0.5, 2))
        pywinauto.mouse.click(button='left', coords=(button_x, button_y))

    else:
        print(f"Phần tử {element_name} không hiển thị trên màn hình")

    # Dừng chương trình một thời gian ngẫu nhiên từ 1 đến 3 giây để đợi cho phần tử được load hoàn tất.
    time.sleep(random.uniform(1, 3))


def click_element_by_title_re(element_name, controltype, x_offset=2, y_offset=2):
    # Tìm phần tử có tên là element_name trong cửa sổ hiện tại của ứng dụng Windows, tìm bằng title.
    element = app.top_window().child_window(title_re=element_name, control_type=controltype)

    # Kiểm tra xem phần tử đó có đang hiển thị hay không.
    if element.is_visible():
        # Lấy kích thước và vị trí của phần tử trên màn hình
        element_rect = element.rectangle()
        element_x_left = element_rect.left
        element_x_right = element_rect.right
        element_y_top = element_rect.top
        element_y_bottom = element_rect.bottom

        # Tạo ra một cặp tọa độ ngẫu nhiên trên phần tử và click vào vị trí đó.
        width, height = pyautogui.size()
        button_x = random.randint(element_x_left + x_offset, element_x_right - x_offset)
        button_y = random.randint(element_y_top + y_offset, element_y_bottom - y_offset)
        while button_y > height - 40:
            button_y = random.randint(element_y_top + y_offset, element_y_bottom - y_offset)
        move_to(button_x, button_y)
        time.sleep(random.uniform(0.5, 2))
        pywinauto.mouse.click(button='left', coords=(button_x, button_y))

    else:
        print(f"Phần tử {element_name} không hiển thị trên màn hình")

    # Dừng chương trình một thời gian ngẫu nhiên từ 1 đến 3 giây để đợi cho phần tử được load hoàn tất.
    time.sleep(random.uniform(1, 3))


def click_element_by_classname(element_name, controltype, x_offset=2, y_offset=2):
    # Tìm phần tử có tên là element_name trong cửa sổ hiện tại của ứng dụng Windows, tìm bằng title.
    element = app.top_window().child_window(class_name=element_name, control_type=controltype)

    # Kiểm tra xem phần tử đó có đang hiển thị hay không.
    if element.is_visible():
        # Lấy kích thước và vị trí của phần tử trên màn hình
        element_rect = element.rectangle()
        element_x_left = element_rect.left
        element_x_right = element_rect.right
        element_y_top = element_rect.top
        element_y_bottom = element_rect.bottom

        # Tạo ra một cặp tọa độ ngẫu nhiên trên phần tử và click vào vị trí đó.
        width, height = pyautogui.size()
        button_x = random.randint(element_x_left + x_offset, element_x_right - x_offset)
        button_y = random.randint(element_y_top + y_offset, element_y_bottom - y_offset)
        while button_y > height - 40:
            button_y = random.randint(element_y_top + y_offset, element_y_bottom - y_offset)
        move_to(button_x, button_y)
        time.sleep(random.uniform(0.5, 2))
        pywinauto.mouse.click(button='left', coords=(button_x, button_y))
    else:
        print(f"Phần tử {element_name} không hiển thị trên màn hình")

    # Dừng chương trình một thời gian ngẫu nhiên từ 1 đến 3 giây để đợi cho phần tử được load hoàn tất.
    time.sleep(random.uniform(2, 4))


def click_element_by_classname_re(element_name, controltype, x_offset=2, y_offset=2):
    # Tìm phần tử có tên là element_name trong cửa sổ hiện tại của ứng dụng Windows, tìm bằng title.
    element = app.top_window().child_window(class_name_re=element_name, control_type=controltype)

    # Kiểm tra xem phần tử đó có đang hiển thị hay không.
    if element.is_visible():
        # Lấy kích thước và vị trí của phần tử trên màn hình
        element_rect = element.rectangle()
        element_x_left = element_rect.left
        element_x_right = element_rect.right
        element_y_top = element_rect.top
        element_y_bottom = element_rect.bottom

        # Tạo ra một cặp tọa độ ngẫu nhiên trên phần tử và click vào vị trí đó.
        width, height = pyautogui.size()
        button_x = random.randint(element_x_left + x_offset, element_x_right - x_offset)
        button_y = random.randint(element_y_top + y_offset, element_y_bottom - y_offset)
        while button_y > height - 40:
            button_y = random.randint(element_y_top + y_offset, element_y_bottom - y_offset)
        move_to(button_x, button_y)
        time.sleep(random.uniform(0.5, 2))
        pywinauto.mouse.click(button='left', coords=(button_x, button_y))
    else:
        print(f"Phần tử {element_name} không hiển thị trên màn hình")

    # Dừng chương trình một thời gian ngẫu nhiên từ 1 đến 3 giây để đợi cho phần tử được load hoàn tất.
    time.sleep(random.uniform(2, 4))


# Cuộn chuột xuống x lần, mỗi lần 1 khoảng cách random
def scroll_down(num_of_scrolls, speed=random.randint(1, 3)):
    for i in range(num_of_scrolls):
        pyautogui.scroll(speed * random.randint(-400, -150))
        time.sleep(random.uniform(1, 6))


# Cuộn chuột lên x lần, mỗi lần 1 khoảng cách random
def scroll_up(num_of_scrolls, speed=random.randint(1, 3)):
    for i in range(num_of_scrolls):
        pyautogui.scroll(speed * random.randint(150, 400))
        time.sleep(random.uniform(1, 6))


# Cuộn chuột xuống cho đến khi thấy element và đọc 10-25 giây
def scroll_down_to_an_element_by_title(element_name, controltype, speed=random.randint(1, 3)):
    is_element_visible = False
    while (is_element_visible == False):
        time.sleep(random.uniform(0.1, 2))
        try:
            target_element = app.top_window().child_window(title=element_name, control_type=controltype)
            if target_element.is_visible():
                is_element_visible = True
                pyautogui.scroll(random.randint(-130, -100))
                print(f"Tìm thấy {element_name} thành công")
                time.sleep(random.uniform(1, 2))
            break
        except:
            pyautogui.scroll(speed * random.randint(-150, -100))
            print(f"Try to get {element_name}")


# Cuộn chuột xuống cho đến khi thấy element và đọc 10-25 giây (dùng title_re)
def scroll_down_to_an_element_by_title_re(element_name, controltype, speed=random.randint(1, 3)):
    is_element_visible = False
    scroll_down_times = 0
    while (is_element_visible == False):
        time.sleep(random.uniform(0.1, 2))
        try:
            target_element = app.top_window().child_window(title_re=element_name, control_type=controltype)
            if target_element.is_visible():
                is_element_visible = True
                pyautogui.scroll(random.randint(100, 150))
                time.sleep(random.uniform(1, 2))
            break
        except:
            scroll_down_times = scroll_down_times + 1
            pyautogui.scroll(speed * random.randint(-150, -100))
            print(f"Try to get {element_name}")


# Cuộn chuột lên cho đến khi thấy element và đọc 10-25 giây
def scroll_up_to_an_element_by_title(element_name, controltype, speed=random.randint(1, 3)):
    is_element_visible = False
    while (is_element_visible == False):
        time.sleep(random.uniform(0.1, 2))
        try:
            target_element = app.top_window().child_window(title=element_name, control_type=controltype)
            if target_element.is_visible():
                is_element_visible = True
                pyautogui.scroll(random.randint(100, 150))
                time.sleep(random.uniform(10, 25))
            break
        except:
            pyautogui.scroll(speed * random.randint(100, 150))
            print(f"Try to get {element_name}")


# Cuộn chuột lên cho đến khi thấy element và đọc (dạng title_re)
def scroll_up_to_an_element_by_title_re(element_name, controltype, speed=random.randint(1, 3)):
    is_element_visible = False
    while (is_element_visible == False):
        time.sleep(random.uniform(0.1, 2))
        try:
            target_element = app.top_window().child_window(title_re=element_name, control_type=controltype)
            if target_element.is_visible():
                is_element_visible = True
                pyautogui.scroll(random.randint(100, 150))
                time.sleep(random.uniform(1, 3))
            break
        except:
            pyautogui.scroll(speed * random.randint(100, 150))
            print(f"Try to get {element_name}")


# Cuộn chuột xuống cho đến khi thấy element và click vào, tìm bằng title
def scroll_down_to_element_and_click_by_title(element_name, controltype, speed=randint(1, 3), x_offset=2, y_offset=2):
    is_element_visible = False
    while (is_element_visible == False):
        time.sleep(random.uniform(0.1, 2))
        try:
            target_element = app.top_window().child_window(title=element_name, control_type=controltype)
            if target_element.is_visible():
                is_element_visible = True
                mouse_move_to_rad()
                pyautogui.scroll(random.randint(-150, -100))
                time.sleep(random.uniform(1, 5))
                click_element_by_title(element_name, controltype, x_offset, y_offset)
            break
        except:
            pyautogui.scroll(speed * random.randint(-150, -100))
            print(f"Try to get {element_name}")


# Cuộn chuột xuống cho đến khi thấy element và click vào, tìm bằng title
def scroll_down_to_element_and_click_by_class_name(element_name, controltype, speed=randint(1, 3), x_offset=2,
                                                   y_offset=2):
    is_element_visible = False
    while (is_element_visible == False):
        time.sleep(random.uniform(0.1, 2))
        try:
            target_element = app.top_window().child_window(class_name=element_name, control_type=controltype)
            if target_element.is_visible():
                is_element_visible = True
                mouse_move_to_rad()
                pyautogui.scroll(random.randint(-150, -100))
                time.sleep(random.uniform(1, 5))
                click_element_by_classname(element_name, controltype, x_offset, y_offset)
                print(f"Click vào {element_name} thành công!")
                time.sleep(1)
            break
        except:
            pyautogui.scroll(speed * random.randint(-150, -100))
            print(f"Try to get {element_name}")


# Cuộn chuột xuống cho đến khi thấy element và click vào, tìm bằng auto_id
def scroll_down_to_element_and_click_by_auto_id(auto_id, controltype, speed=randint(1, 3), x_offset=2, y_offset=2):
    is_element_visible = False
    while (is_element_visible == False):
        time.sleep(random.uniform(0.1, 2))
        try:
            target_element = app.top_window().child_window(title=auto_id, control_type=controltype)
            if target_element.is_visible():
                mouse_move_to_rad()
                pyautogui.scroll(random.randint(-150, -100))
                is_element_visible = True
                time.sleep(random.uniform(1, 5))
                click_element_by_auto_id(auto_id, controltype, x_offset, y_offset)
                print(f"Click vào {auto_id} thành công!")
            break
        except:
            pyautogui.scroll(speed * random.randint(-150, -100))
            print(f"Try to get {auto_id}")


# Cuộn chuột lên cho đến khi thấy element và click vào bằng Title
def scroll_up_to_element_and_click_by_title(element_name, controltype, speed=random.randint(2, 3), x_offset=2,
                                            y_offset=2):
    is_element_visible = False
    while (is_element_visible == False):
        time.sleep(random.uniform(0.1, 1))
        try:
            target_element = app.top_window().child_window(title=element_name, control_type=controltype)
            if target_element.is_visible():
                is_element_visible = True
                pyautogui.scroll(random.randint(100, 150))
                mouse_move_to_rad()
                time.sleep(random.uniform(1, 5))
                click_element_by_title(element_name, controltype, x_offset, y_offset)
                print(f"Click vào {element_name} thành công!")
            break
        except:
            pyautogui.scroll(speed * random.randint(100, 150))
            print(f"Try to get {element_name}")


# Cuộn chuột lên cho đến khi thấy element và click vào bằng Classname
def scroll_up_to_element_and_click_by_Class_name(element_name, controltype, speed=random.randint(2, 3), x_offset=2,
                                                 y_offset=2):
    is_element_visible = False
    while (is_element_visible == False):
        time.sleep(random.uniform(0.1, 1))
        try:
            target_element = app.top_window().child_window(class_name=element_name, control_type=controltype)
            if target_element.is_visible():
                is_element_visible = True
                pyautogui.scroll(random.randint(100, 150))
                mouse_move_to_rad()
                time.sleep(random.uniform(1, 5))
                click_element_by_classname(element_name, controltype, x_offset, y_offset)
                print(f"Click vào {element_name} thành công!")
            break
        except:
            pyautogui.scroll(speed * random.randint(100, 150))
            print(f"Try to get {element_name}")


# Cuộn chuột lên cho đến khi thấy element và click vào bằng auto_id
def scroll_up_to_element_and_click_by_auto_id(auto_id, controltype, speed=random.randint(2, 3), x_offset=2, y_offset=2):
    is_element_visible = False
    while (is_element_visible == False):
        time.sleep(random.uniform(0.1, 1))
        try:
            target_element = app.top_window().child_window(auto_id=auto_id, control_type=controltype)
            if target_element.is_visible():
                is_element_visible = True
                pyautogui.scroll(random.randint(100, 150))
                mouse_move_to_rad()
                time.sleep(random.uniform(1, 2))
                click_element_by_auto_id(auto_id, controltype, x_offset, y_offset)
                print(f"Click vào {auto_id} thành công!")
            break
        except:
            pyautogui.scroll(speed * random.randint(100, 150))
            print(f"Try to get {auto_id}")


# Chạy Nekoray
def nekoray_VPN_start():
    # Khởi chạy Nekoray
    nekoray_path = os.path.expandvars(r"%USERPROFILE%\Desktop\nekoray\nekoray.exe")
    f = Application(backend="uia").start(nekoray_path)
    time.sleep(5)
    # time.sleep(5)
    nekoray = Application(backend="uia").connect(title_re='.*NekoRay*.')
    nekoray_main_window = nekoray.window(title_re='.*NekoRay*.')
    maximize_button = nekoray_main_window.child_window(title="Maximize", control_type="Button")
    mimimize_button = nekoray_main_window.child_window(title="Minimize", control_type="Button")
    time.sleep(3)
    maximize_button.click()
    print('maximize thanh cong')
    time.sleep(2)
    nekoray = Application(backend="uia").connect(title_re='.*NekoRay*.')
    VPN_check_box = nekoray_main_window.child_window(title="VPN Mode", control_type="CheckBox")
    VPN_check_box.wait('visible')
    # VPN_check_box.print_control_identifiers()
    # Lấy giá trị của checkbox
    Check_box_value = VPN_check_box.get_toggle_state()
    print('Giá trị của checkbox 1-True, 0- False:', Check_box_value)
    if Check_box_value == 1:
        print('Chế độ VPN đã được bật, kiểm tra lại Proxy')
    else:
        VPN_check_box.click_input()
        print('Bật chế độ VPN thành công')
    # time.sleep(20)
    mimimize_button.click()
    time.sleep(20)


def nekoray_exit():
    #    Khởi chạy Nekoray
    nekoray_path = os.path.expandvars(r"%USERPROFILE%\Desktop\nekoray\nekoray.exe")
    Application(backend="uia").start(nekoray_path)
    time.sleep(5)
    nekoray = Application(backend="uia").connect(title_re='.*NekoRay*.')
    nekoray_main_window = nekoray.window(title_re='.*NekoRay*.')
    maximize_button = nekoray_main_window.child_window(title="Maximize", control_type="Button")
    mimimize_button = nekoray_main_window.child_window(title="Minimize", control_type="Button")
    time.sleep(2)
    maximize_button.click()
    print('maximize thanh cong')
    time.sleep(1)
    # nekoray = Application(backend="uia").connect(title_re='.*NekoRay*.')
    program_button = nekoray_main_window.child_window(title="Program", control_type="Button")
    program_button.wait('visible')
    time.sleep(1)
    program_button.click()
    time.sleep(2)
    move_to(116, 351)
    pywinauto.mouse.click(button='left', coords=(116, 351))
    time.sleep(5)


# Về trang chủ ở bất kỳ trang nào ở ebay
def go_to_ebay_home_page():
    mouse_move_to_rad()
    time.sleep(random.uniform(1, 3))
    scroll_up_to_element_and_click_by_title("eBay Logo", "Hyperlink")


# Đọc ebay Help&Contact
def read_ebay_selling_policy():
    click_element_by_title("Help & Contact", "Hyperlink")
    mouse_move_to_rad()
    scroll_down_to_element_and_click_by_title("Selling", "Hyperlink")
    time.sleep(random.uniform(3, 6))
    scroll_down(randint(6, 10))


# lựa chọn màu sắc, kích thước của item nếu có:
def choose_random_item_option(auto_ID="x-msku__select-box-1000", control_type="ComboBox", list_height=20):
    combo_box = app.top_window().child_window(auto_id=auto_ID, control_type=control_type)
    if combo_box.exists():
        # Click vào combobox
        click_element_by_auto_id(auto_ID, control_type)
        print(f"Click ComboBox thành công")
        # combo_box.print_control_identifiers()
        # time.sleep(random.uniform(1, 2))
        with contextlib.redirect_stdout(StringIO()) as f:
            combo_box.print_control_identifiers()
            identifiers_str = f.getvalue()

        # print(identifiers_str)
        pattern = r'title="([^"]*)"[\s\S]*?auto_id="([^"]*)"[\s\S]*?control_type="([^"]*)"'
        matches = re.findall(pattern, identifiers_str)
        # result = [[match[0], match[1], match[2]] for match in matches]
        result = [match[0] for match in matches]
        selected_element = random.choice(result)
        # Tìm phần tử thích hợp
        index = None
        while index is None:
            selected_element = random.choice(result)
            for i, element in enumerate(result):
                if element == selected_element and i > 1 and "Out Of Stock" not in element:
                    index = i
                    break

        # In ra số thứ tự của phần tử thích hợp
        print(f"Giá trị được chọn là: {selected_element}")
        print(f"Số thứ tự của giá trị được chọn là: {index}")

        total_index = len(result)
        print(f"Tổng số lựa chọn là: {total_index}")

        # Xác định tọa độ của lựa chọn:
        # List_height mặc định là 20pixel -> Khoảng rộng của ô List
        # Lấy tọa độ của combobox:
        width, height = pyautogui.size()
        combo_box_rect = combo_box.rectangle()
        combo_box_x_left = combo_box_rect.left
        combo_box_x_right = combo_box_rect.right
        combo_box_y_top = combo_box_rect.top
        combo_box_y_bottom = combo_box_rect.bottom
        # Tọa độ của List_Item được chọn
        list_item_x = random.randint(combo_box_x_left + 2, combo_box_x_right - 2)
        if combo_box_y_bottom + 5 + (total_index - 1) * list_height <= height - 40:
            list_item_y = random.randint(combo_box_y_bottom + 5 + (index - 1) * list_height + 2,
                                         combo_box_y_bottom + 5 + index * list_height - 2)
        else:
            list_item_y = random.randint(combo_box_y_top - 5 - (total_index - index) * list_height + 2,
                                         combo_box_y_top - 5 - (total_index - index - 1) * list_height - 2)
        print(f"Tọa độ click là: {list_item_x}, {list_item_y}")
        move_to(list_item_x, list_item_y)
        pywinauto.mouse.click(button='left', coords=(list_item_x, list_item_y))
        print(f"Chọn giá trị random cho item thành công")
    else:
        print("Không có option")


# Sort item search by random:
def sort_selector():
    sort_selector= app.top_window().child_window(title_re="Sort selector.*", control_type ="Button")
    if sort_selector.exists():
        # Click vào sort_selector
        click_element_by_title_re("Sort selector.*", "Button")
        print(f"Click sort_selector thành công")
        sort_selector.print_control_identifiers()
        # time.sleep(random.uniform(1, 2))
        with contextlib.redirect_stdout(StringIO()) as f:
            sort_selector.print_control_identifiers()
            identifiers_str = f.getvalue()

        print(identifiers_str)
        pattern = r'title="([^"]*)"[\s\S]*?auto_id="([^"]*)"[\s\S]*?control_type="([^"]*)"'
        matches = re.findall(pattern, identifiers_str)
        # result = [[match[0], match[1], match[2]] for match in matches]
        result = [match[0] for match in matches]
        selected_element = random.choice(result)
        # Tìm phần tử thích hợp
        index = None
        while index is None:
            selected_element = random.choice(result)
            for i, element in enumerate(result):
                if element == selected_element and i > 1 and "Out Of Stock" not in element:
                    index = i
                    break

        # In ra số thứ tự của phần tử thích hợp
        print(f"Giá trị được chọn là: {selected_element}")
        print(f"Số thứ tự của giá trị được chọn là: {index}")

        total_index = len(result)
        print(f"Tổng số lựa chọn là: {total_index}")

        # Xác định tọa độ của lựa chọn:
        # List_height mặc định là 20pixel -> Khoảng rộng của ô List
        # Lấy tọa độ của combobox:
        width, height = pyautogui.size()
        combo_box_rect = combo_box.rectangle()
        combo_box_x_left = combo_box_rect.left
        combo_box_x_right = combo_box_rect.right
        combo_box_y_top = combo_box_rect.top
        combo_box_y_bottom = combo_box_rect.bottom
        # Tọa độ của List_Item được chọn
        list_item_x = random.randint(combo_box_x_left + 2, combo_box_x_right - 2)
        if combo_box_y_bottom + 5 + (total_index - 1) * list_height <= height - 40:
            list_item_y = random.randint(combo_box_y_bottom + 5 + (index - 1) * list_height + 2,
                                         combo_box_y_bottom + 5 + index * list_height - 2)
        else:
            list_item_y = random.randint(combo_box_y_top - 5 - (total_index - index) * list_height + 2,
                                         combo_box_y_top - 5 - (total_index - index - 1) * list_height - 2)
        print(f"Tọa độ click là: {list_item_x}, {list_item_y}")
        move_to(list_item_x, list_item_y)
        pywinauto.mouse.click(button='left', coords=(list_item_x, list_item_y))
        print(f"Chọn giá trị random cho item thành công")
    else:
        print("Không có option")


# Thao tác trong trang items:
def auto_actions_on_the_detailed_item_page():
    # Check if have close button
    if app.top_window().child_window(title_re=".*close.*", control_type="Button").exists():
        click_element_by_title_re(".*close.*", "Button")
    # 1.Click vào hình ảnh để phóng to:
    click_element_by_classname_re(".*ux-image-carousel.*", "Button", 30, 5)
    print("Click vào Image thành công")
    time.sleep(3)

    is_next_image_button = False
    while not is_next_image_button:
        try:
            next_image_button = app.top_window().child_window(title="Next image - Item images",
                                                              control_type="Button")
            if next_image_button.exists():
                is_next_image_button = True
                # Click vào nút next image, một số lần ngẫu nhiên
                click_element_by_title("Next image - Item images", "Button", 7, 7)
                time.sleep(random.uniform(1, 3))
                num_clicks = random.randint(1, 6)
                for i in range(num_clicks):
                    pyautogui.click()
                    time.sleep(random.uniform(1, 4))
                if random.random() < 0.5:
                    # Click vào nút previous image, một số lần ngẫu nhiên
                    click_element_by_title("Previous image - Item images", "Button", 7, 7)
                    time.sleep(random.uniform(1, 2))
                    num_clicks = random.randint(1, 6)
                    for i in range(num_clicks):
                        pyautogui.click()
                        time.sleep(random.uniform(0.5, 2))
            else:
                break
        except:
            break
    time.sleep(random.uniform(0.1, 1))
    click_element_by_title("Close image gallery dialog", "Button", 5, 5)
    time.sleep(random.uniform(0.1, 1))
    mouse_move_to_rad()
    time.sleep(random.uniform(0.1, 1))
    scroll_down_to_an_element_by_title("About this item", "Button")
    time.sleep(2)
    random_percent = random.random()
    print(random_percent)
    # if random.percent < 0.03:
    #     scroll_up(0.5)
    #     mouse_move_to_rad()
    #     click_element_by_title("Shipping, returns & payments", "Button")
    # print("Không bấm vào Shipping detail, tiếp tục")
    # scroll_down(2)fire stic
    # time.sleep(10)
    if random_percent <= 0.5:
        scroll_up_to_an_element_by_title("Add to watchlist", "Button")
        scroll_up(1)
        choose_random_item_option("x-msku__select-box-1000")
        choose_random_item_option("x-msku__select-box-1001")
        choose_random_item_option("x-msku__select-box-1002")
        choose_random_item_option("x-msku__select-box-1003")
        time.sleep(2)
        scroll_down_to_element_and_click_by_title("Add to watchlist", "Button", 1)

        # scroll_up_to_an_element_by_title("Add to cart", "HyperLink",2)
        time.sleep(3)
        click_element_by_title("Add to cart", "Hyperlink")
        time.sleep(3)
        actions = [
            lambda: click_element_by_title_re(".*Close button.*", "Button"),  # Close Dialog
            lambda: click_element_by_title_re(".*Checkout.*", "Hyperlink"),  # Go to Checkout Page
            lambda: click_element_by_title("Go to cart", "Hyperlink"),  # Go to Cart
            lambda: click_element_by_title("See all", "Hyperlink")  # See other items like this
        ]
        weights = [0.2, 0.3, 0.4, 0.1]
        random_action = random.choices(actions, weights=weights)[0]
        random_action()

    else:
        if random.random() <= 0.8:
            scroll_up_to_an_element_by_title_re(".*Similar.*", "Text")
            scroll_down(1)
            click_random_items_all_type()
            auto_actions_on_the_detailed_item_page()
        else:
            print("Không bấm chọn items khác, tiếp tục")


# Kiểm tra IP
is_session_ok = False
while not is_session_ok:
    try:
        country_begin = ''
        ip_request = requests.get(url='http://lumtest.com/myip.json', params='PARAMS', timeout=8)
        while country_begin == '':
            try:
                ip_data_begin = ip_request.json()
                ip_begin = ip_data_begin['ip']
                country_begin = ip_data_begin['country']
                print('country_begin: ' + country_begin)
                time.sleep(3)
            except:
                time.sleep(5)
                ip_begin = ''
                print('try get ip_begin')

        ip = requests.get(url='http://lumtest.com/myip.json', params='PARAMS', timeout=8)
        ip_data = ip.json()
        country = ip_data['country']
        if country != 'US' and country != '':
            nekoray_VPN_start()
            continue

        tz_timezone = ip_data['geo']['tz']
        city_name = ip_data['geo']['city']
        region_name = ip_data['geo']['region_name']
        postal_code = ip_data['geo']['postal_code']
        if postal_code == '':
            postal_code = 'non'
        print('IP OK')
        is_session_ok = True
    except Exception as e:
        print('Error occurred:', e)

# Thiết lập kích thước màn hình:
is_session_ok = True
try:
    if is_session_ok == True:
        ran_solution = random.randint(1, 1)
        if ran_solution == 1:
            ScreenRes.set(1920, 1080)
        if ran_solution == 2:
            ScreenRes.set(1366, 768)
        if ran_solution == 3:
            ScreenRes.set(1600, 900)
        time.sleep(3)
except:
    print('Set solution false!')
# Clear DNS cache
subprocess.run(['ipconfig', '/flushdns'], check=True)
print("All ARE OKAY, LET'S GO")
time.sleep(5)
# Khởi động trình duyệt Microsoft Edge và đợi cho tới khi cửa sổ "New tab - Profile 1" được mở ra
Application(backend="uia").start(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
time.sleep(20)

for gmail in _arr_gmail_infor:
    try:
        try:
            app = Application(backend="uia").connect(title_re=".*Edge.*")
            app.top_window().wait('ready', timeout=60)
            app.top_window().maximize()
            time.sleep(10)
            address_bar = app.window(title_re='.*Profile.*').child_window(auto_id="view_1020",
                                                                                      control_type="Edit")
            address_bar.wait('ready', timeout=60)
        except Exception as e:
            print(f"Lỗi khi kết nối đến trình duyệt Microsoft Edge: {e}")
            exit()

        # Nhập địa chỉ trang web và chờ cho đến khi nó được tải hoàn tất
        try:
            address_bar.set_text("ebay.com")
            time.sleep(2)
            send_keys("{ENTER}")
            app.top_window().wait('ready', timeout=60)
        except Exception as e:
            print(f"Lỗi khi truy cập trang web: {e}")
            exit()

        # Tìm lời mời tải ứng dụng ebay, nếu có click vào nút close để đóng!
        app_count = 0
        app_button_visible = False
        while not app_button_visible:
            try:
                # Tìm phần tử qua title:
                app_button = app.top_window().child_window(title="Dismiss this dialog", control_type="Button")
                time.sleep(1)
                if app_button.exists():
                    app_button_visible = True
                    click_element_by_title("Dismiss this dialog", "Button")
                    print("Đóng lời mời tải ứng dụng ebay thành công!")
            except:
                print(f"Kiểm tra lại xuất hiện của lời mời tải ứng dụng Ebay lần thứ {app_count + 1}")
                time.sleep(1)
            app_count = app_count + 1
            if app_count == 3:
                print(f"Không có lời mời tải ứng dụng Ebay sau {app_count} lần kiểm tra")
                break

        # Đọc Ebay Seller Policy, xác suất là 20%
        time.sleep(random.uniform(2, 3))
        mouse_move_to_rad()

        if random.random() <= 0.1:
            read_ebay_selling_policy()
            # Về trang chủ của ebay
            go_to_ebay_home_page()
            mouse_move_to_rad()

        # time.sleep(random.uniform(1, 3))
        # Cuộn lên và tìm click vào Ô search Items, nếu không tìm thấy, sẽ chờ mãi ở đây:
        scroll_up_to_element_and_click_by_auto_id("gh-ac", "ComboBox")
        time.sleep(random.uniform(1, 2))

        # Nhập từ khóa vào ô tìm kiếm
        type(gmail._rd_first_name)
        print(f"Nhập từ khóa {gmail._rd_first_name} để tìm kiếm thành công ")

        time.sleep(random.uniform(1, 2))
        send_keys('{ENTER}')
        time.sleep(random.uniform(1, 2))
        mouse_move_to_rad()
        time.sleep(random.uniform(3, 5))

        # Tùy chọn cho tìm kiếm:
        # Type of Listing:
        Type_of_listing = app.top_window().child_window(title="Auction", control_type="Hyperlink")
        Type_of_listing.wait('visible', timeout=30)
        sort_selector = app.top_window().child_window(title_re="Sort selector.*", control_type="Button")
        if sort_selector.exists():
            # Click vào sort_selector
            click_element_by_title_re("Sort selector.*", "Button")
            print(f"Click sort_selector thành công")
            # random_choice:
            options = [2, 3, 4, 5, 6]
            weights = [0.1, 0.1, 0.7, 0.05, 0.05]
            choice = random.choices(options, weights=weights)[0]
            index = choice
            # In ra số thứ tự của phần tử thích hợp
            print("Chọn số thứ", index, ": ", choice)
            print(f"Giá trị được chọn là: {choice}")
            print(f"Số thứ tự của giá trị được chọn là: {index}")
            # Xác định tọa độ của lựa chọn:
            # Lấy tọa độ của combobox:
            width, height = pyautogui.size()
            sort_selector_rect = sort_selector.rectangle()
            sort_selector_x_left = sort_selector_rect.left
            sort_selector_x_right = sort_selector_rect.right
            sort_selector_y_top = sort_selector_rect.top
            sort_selector_y_bottom = sort_selector_rect.bottom
            # Tọa độ của List_Item được chọn
            selected_x = random.randint(sort_selector_x_left - (sort_selector_x_right-sort_selector_x_left)+10, sort_selector_x_right - 5)
            selected_y = random.randint(sort_selector_y_bottom + 15 + (index - 1) * 35 + 4,
                                             sort_selector_y_bottom + 15 + index * 35 - 4)

            print(f"Tọa độ click là: {selected_x}, {selected_y}")
            move_to(selected_x, selected_y)
            pywinauto.mouse.click(button='left', coords=(selected_x, selected_y))
            print(f"Chọn giá trị cho sort_selector thành công")
        else:
            print("Không có option")

        if Type_of_listing.exists():
            time.sleep(3)
            options = ["Auction", "Buy It Now"]
            for option in options:
                if random.random() < 0.3:
                    click_element_by_title(option, "Hyperlink")
                    time.sleep(3)
        # Sort by:
        if random.random() < 0.3:
            click_element_by_title_re("Listing options selector.*", "Button")
            Listing_options = app.top_window().child_window(title_re="Listing options selector.*", control_type="Button")

            width, height = pyautogui.size()
            Listing_options_rect = Listing_options.rectangle()
            Listing_options_x_left = Listing_options_rect.left
            Listing_options_x_right = Listing_options_rect.right
            Listing_options_y_top = Listing_options_rect.top
            Listing_options_y_bottom = Listing_options_rect.bottom
            # Tọa độ của List_Item được chọn
            selected_x = random.randint(Listing_options_x_left - 30 + 5 , Listing_options_x_right - 5)
            selected_y = random.randint(Listing_options_y_bottom + 5 + 5, Listing_options_y_bottom + 5 + 35 - 5)
            move_to(selected_x, selected_y)
            pywinauto.mouse.click(button='left', coords=(selected_x, selected_y))
            print(f"Chọn Listing_options thành công")
            time.sleep(10)

        # Cuon chuot
        scroll_actions = [(random.randint(2, 4), "down"), (random.randint(1, 2), "up"), (random.randint(3, 5), "down"),
                          (random.randint(1, 2), "up")]
        for scroll_amount, scroll_direction in scroll_actions:
            if scroll_direction == "down":
                scroll_down(scroll_amount)
            else:
                scroll_up(scroll_amount)
            # Randomly move the mouse with a probability of 50%
            if random.random() < 0.7:
                mouse_move_to_rad()

        # Click vào 1 items random trong danh sách tìm kiếm
        click_random_items_all_type()
        time.sleep(3)
        mouse_move_to_rad()
        # Thao tác trong trang items chi tiết:
        auto_actions_on_the_detailed_item_page()

    except:
        print("It's OK, no problem, bro")
        time.sleep(10)

