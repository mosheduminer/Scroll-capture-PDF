import PySimpleGUI as sg
from PIL import Image, ImageGrab, ImageFile
from img2pdf import convert
from os import remove
from os.path import join
from tempfile import mkdtemp
from shutil import rmtree
from time import sleep
from win32gui import GetCursorPos
from win32api import GetAsyncKeyState, mouse_event, keybd_event
from win32con import VK_LBUTTON, MOUSEEVENTF_WHEEL, VK_NEXT, KEYEVENTF_EXTENDEDKEY
import ctypes


def crop(original, top, bottom, left, right, mode):
    try:
        im = original.Crop((left, top, im_height-bottom, im_width-right))
    except:
        if mode == "top":
            top -= 1
        elif mode == "bottom":
            bottom -= 1
        elif mode == "left":
            left -= 1
        elif mode == "right":
            right -= 1
        elif mode == "new":
            top, bottom, left, right = 0
    return im, top, bottom, left, right


if __name__ == '__main__':
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
    # set up gui
    layout1 = [[sg.Text("Target folder")],
               [sg.In(), sg.FolderBrowse()],
               [sg.In("Target name")],
               [sg.Text("Number of pages")],
               [sg.In()],
               [sg.Checkbox("Use number of pages for scrolling", default=True, enable_events=True)],
               [sg.Text("Max scroll")],
               [sg.In(disabled=True)],
               [sg.Combo(["Mouse", "Pagedown"], size=(10, 1))],
               [sg.RButton("Select area"), sg.CloseButton("Cancel")],
               [sg.RButton("Begin capture")]]
    win1 = sg.Window("Otzar capture").Layout(layout1)

    win2_active = False

    while True:
        ev1, values = win1.Read()
        print("event: ", ev1)
        print("values: ", values)

        if ev1 is None or ev1 == "Cancel":
            exit(0)

        # check value of checkbox when changed
        elif ev1 == 3:
            if values[3] is False:
                layout1[7][0].Update(disabled=False)
            elif values[3] is True:
                layout1[7][0].Update(value="", disabled=True)

        elif ev1 == "Select area":
            sleep(0.1)

            # get target area
            # takes position of pointer when mouse button is released
            # first get top left corner
            while True:
                state = GetAsyncKeyState(VK_LBUTTON)
                if state <= 0:
                    while True:
                        next_state = GetAsyncKeyState(VK_LBUTTON)
                        if next_state != state:
                            break
                    s_x, s_y = GetCursorPos()
                    break
            sleep(0.2)

            # get bottom right corner
            while True:
                state = GetAsyncKeyState(VK_LBUTTON)
                if state < 0:
                    while True:
                        next_state = GetAsyncKeyState(VK_LBUTTON)
                        if next_state != state:
                            break
                    e_x, e_y = GetCursorPos()
                    break
            if s_x > e_x:
                s_x, e_x = e_x, s_x
            if s_y > e_y:
                s_y, e_y = e_y, s_y

        elif ev1 == "Begin capture":
            if values[0] == "":
                sg.PopupOK("Please select a directory.")
            elif values[1] == "":
                sg.PopupOK("Please enter a target file name.")
            elif values[2] == "":
                sg.PopupOk("Please enter the number of pages.")
            elif ("e_x" not in globals()) | ("s_x" not in globals()):
                sg.PopupOK("you haven't selected a capture area!")
            else:
                pages = int(values[2])
                # set scroll number
                if values[3] is True:
                    scroll_num = pages
                elif values[3] is False:
                    scroll_num = int(values[4])

                target_path = join(values[0], values[1])
                x = int((e_x - s_x) / 2 + e_x)
                y = int((e_y - s_y) / 2 + e_y)
                if values[5] == "Mouse":
                    def scroll():
                        mouse_event(MOUSEEVENTF_WHEEL, x, y, -1, 0)
                elif values[5] == "Pagedown":
                    def scroll():
                        keybd_event(VK_NEXT, KEYEVENTF_EXTENDEDKEY)

                temp = mkdtemp()
                sleep(2)
                for i in range(1, scroll_num+1):
                    screen = ImageGrab.grab(bbox=(s_x, s_y, e_x, e_y))
                    screen.save(join(temp, f"temp{i}.png"))
                    if i != scroll_num:
                        scroll()
                        sleep(0.3)

                with open(f"{target_path}.pdf", "wb") as f:
                    f.write(convert([join(temp, f"temp{i}.png") for i in range(1, pages+1)]))
                rmtree(temp)
                win1.UnHide()
