import autoit
from win32api import GetWindowLong
import win32con
import time
import pyautogui
import sys


c_path = "C:\\Program Files (x86)\\Cisco\\Cisco AnyConnect Secure Mobility Client\\vpnui.exe"
c_title1 = "Cisco AnyConnect Secure Mobility Client"

c_connect_btn = "Connect"
c_disconnect_btn = "Disconnect"
c_host_style2 = 0x08000000
c_btn1_style2 = 0x08000000

c_host_control = '[CLASS:Edit; INSTANCE:1]'
c_btn1_control = '[CLASS:Button; INSTANCE:1]'
c_user_control = '[CLASS:Edit; INSTANCE:1]'
c_password_control = '[CLASS:Edit; INSTANCE:2]'
c_ok_control = '[CLASS:Button; INSTANCE:1]'
c_cancel_control = '[CLASS:Button; INSTANCE:2]'


def _open_wnd():
    print("_open_wnd; path="+c_path)
    print("_open_wnd; title=" + c_title1)
    autoit.run(c_path)
    while True:
        if autoit.win_exists(c_title1):
            break
        time.sleep(1)
    autoit.win_activate(c_title1)
    return autoit.win_get_handle(c_title1)


def _get_control(l_win, control):
    handle = autoit.control_get_handle(l_win, control)
    style = GetWindowLong(handle, win32con.GWL_STYLE)
    text = autoit.control_get_text_by_handle(l_win, handle)
    return handle, text, style


def _set_control_text(hwnd, hcontrol, text):
    if text == autoit.control_get_text_by_handle(hwnd, hcontrol):
        return True
    autoit.control_click_by_handle(hwnd, hcontrol)
    pyautogui.hotkey('Ctrl', 'a')
    pyautogui.typewrite(text, 0.02)
    return True


def _set_host(l_win):
    host_ctl = _get_control(l_win, c_host_control)
    btn1_ctl = _get_control(l_win, c_btn1_control)
    print('_set_host; host control', host_ctl[0], host_ctl[1], hex(host_ctl[2]))
    print('_set_host; btn1 control', btn1_ctl[0], btn1_ctl[1], hex(btn1_ctl[2]))
    if host_ctl[1] == c_host:
        return True

    if (host_ctl[1] != c_host) and (btn1_ctl[1] == c_disconnect_btn) and not(btn1_ctl[2] & c_btn1_style2):
        autoit.control_click_by_handle(l_win, btn1_ctl[0])
        while True:
            time.sleep(1)
            host_ctl = _get_control(l_win, c_host_control)
            if not(host_ctl[2] & c_host_style2):
                break

    if not(host_ctl[2] & c_host_style2):
        return _set_control_text(l_win, host_ctl[0], c_host)
    else:
        return False


def _main_proc():
    print("host=" + c_host)
    print("user_name=" + c_user_name)
    if c_password != "":
        print("password=***************")
    l_win1 = _open_wnd()
    print("window handle=", l_win1)
    while True:
        if _set_host(l_win1):
            print("host set ok")
            l_btn1 = _get_control(l_win1, c_btn1_control)
            print('button1: ', l_btn1[0], l_btn1[1], hex(l_btn1[2]))
            if (l_btn1[1] == c_connect_btn) and not(l_btn1[2] & c_btn1_style2):
                autoit.control_click_by_handle(l_win1, l_btn1[0])
                print("button1 clicked, waiting for login prompt")
            elif (l_btn1[1] == c_connect_btn) and (l_btn1[2] & c_btn1_style2):
                print("window activating: " + c_title2)
            elif l_btn1[1] == c_disconnect_btn:
                print("already connected, exiting")
                autoit.win_close_by_handle(l_win1)
                return

            try:
                l_win2 = autoit.win_get_handle(c_title2)
                print("activating login prompt, handle ", l_win2)
                autoit.win_activate_by_handle(l_win2)
                break
            except:
                l_win2 = None
        else:
            print("host is not set")
            return

        time.sleep(1)

    print("setting user name")
    _set_control_text(
        l_win2,
        autoit.control_get_handle(
            l_win2,
            c_user_control),
        c_user_name)

    if c_password != "":
        print("setting password")
        _set_control_text(
            l_win2,
            autoit.control_get_handle(
                l_win2,
                c_password_control),
            c_password)

        autoit.control_click_by_handle(
            l_win2,
            autoit.control_get_handle(
                l_win2,
                c_ok_control))
    else:
        autoit.control_click_by_handle(
            l_win2,
            autoit.control_get_handle(
                l_win2,
                c_password_control))



if __name__ == "__main__":
    c_password = ""
    c_host = ""
    c_user_name =""
    for a in sys.argv[1:]:
        name, value = a.split('=', 1)
        if value[0] == '"':
            value = value[1:-1]
        if name == 'host':
            c_host = value
        elif name == 'user_name':
            c_user_name = value
        elif name == 'password':
            c_password = value

    if (c_user_name == "") | (c_host == ""):
        print("usage: " + sys.argv[0] + " host=<host> user_name=<user_name> [password=<password>]")
        print("example: " + sys.argv[0] + " host=vpn.mydomain.com user_name=JShmidt password=\"myPassPhrase\"")
        input("press Enter key")
    else:
        c_title2 = "Cisco AnyConnect | " + c_host
        try:
            _main_proc()
        except :
            print(sys.exc_info()[0])
    print("exiting in 3 seconds...")
    time.sleep(3)