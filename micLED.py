#!/usr/bin/env python3

import subprocess
import sys
import threading

import pynput
import requests
import win32con
import win32gui_struct

try:
    import winxpgui as win32gui
except ImportError:
    import win32gui


class SysTrayApp():
    FIRST_ID = 1023


    def __init__(self, icon_filename, hover_text, menu_options, double_click_action=None):
        self.double_click_action = double_click_action
        self.hover_text = hover_text
        self.icon_init = False
        self.menu_actions = {}
        self.next_action_id = self.FIRST_ID

        # Menu options list will be used to create menu (with sub-menus)
        # Add default QUIT action to the menu options list
        self.menu_options = self.create_menu_options(menu_options + [['Quit', self.destroy]])
        # Menu actions list will be used to map action ID (click in the menu) to the function
        self.create_menu_actions(self.menu_options)

        # self.next_action_id is used only for assigning IDs to menu options
        del self.next_action_id

        # Register the Window class
        window_class = win32gui.WNDCLASS()
        window_class.lpszClassName = "SysTrayApp"
        # Retrieves a module handle for the specified module. If this parameter is NULL,
        # GetModuleHandle returns a handle to the file used to create the calling process (.exe file)
        window_class.hInstance = win32gui.GetModuleHandle(None)
        # Redraws the entire window if a movement or size adjustment changes the height/width of the client area
        window_class.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW
        # Arrow style
        window_class.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        # Background
        window_class.hbrBackground = win32con.COLOR_WINDOW
        # Map window procedures
        window_class.lpfnWndProc = {
            # win32gui.RegisterWindowMessage("TaskbarCreated"): self.restart,
            win32con.WM_DESTROY: self.destroy,
            win32con.WM_COMMAND: self.command,
            win32con.WM_USER+20: self.notify
        }

        self.hwnd = win32gui.CreateWindow(
            # Registers a window class for subsequent use in calls to the CreateWindow
            win32gui.RegisterClass(window_class),
            "SysTrayApp",
            win32con.WS_OVERLAPPED | win32con.WS_SYSMENU,
            0,
            0,
            win32con.CW_USEDEFAULT,
            win32con.CW_USEDEFAULT,
            0,
            0,
            win32gui.GetModuleHandle(None),
            None
        )

        self.menu = win32gui.CreatePopupMenu()

        self.create_menu(self.menu, self.menu_options)

        self.change_icon(icon_filename)

        win32gui.UpdateWindow(self.hwnd)


    def start_sys_tray_app(self):
        win32gui.PumpMessages()


    def create_menu_options(self, menu_options):
        menu_options_with_ids = []

        for menu_option in menu_options:
            self.next_action_id += 1

            if callable(menu_option[1]):
                menu_options_with_ids.append([self.next_action_id] + menu_option)
            else:
                menu_options_with_ids.append(
                    [self.next_action_id, menu_option[0]] + [self.create_menu_options(menu_option[1])]
                )

        return menu_options_with_ids


    def create_menu_actions(self, menu_options):
        for menu_option_data in menu_options:
            if callable(menu_option_data[2]):
                self.menu_actions[menu_option_data[0]] = menu_option_data[2]
            else:
                self.create_menu_actions(menu_option_data[2])


    def create_menu(self, menu, menu_options):
        for menu_option_data in menu_options[::-1]:
            if callable(menu_option_data[2]):
                item, _ = win32gui_struct.PackMENUITEMINFO(text=menu_option_data[1], wID=menu_option_data[0])
                win32gui.InsertMenuItem(menu, 0, 1, item)
            else:
                submenu = win32gui.CreatePopupMenu()
                self.create_menu(submenu, menu_option_data[2])
                item, _ = win32gui_struct.PackMENUITEMINFO(text=menu_option_data[1], hSubMenu=submenu)
                win32gui.InsertMenuItem(menu, 0, 1, item)


    def change_icon(self, icon_filename):
        notify_id = (
            self.hwnd,
            0,
            win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
            win32con.WM_USER + 20,
            win32gui.LoadImage(
                win32gui.GetModuleHandle(None),
                icon_filename,
                win32con.IMAGE_ICON,
                0,
                0,
                win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            ),
            self.hover_text
        )

        if self.icon_init:
            win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, notify_id)
        else:
            win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, notify_id)
            self.icon_init = True


    def destroy(self, hwnd, msg, wparam, lparam):
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (self.hwnd, 0))
        win32gui.PostQuitMessage(0)

        # TODO: Add proper exit handler
        sys.exit()


    def command(self, hwnd, msg, wparam, lparam):
        menu_option_id = win32gui.LOWORD(wparam)
        self.execute_menu_option(menu_option_id)


    def notify(self, hwnd, msg, wparam, lparam):
        if lparam == win32con.WM_LBUTTONDBLCLK:
            if callable(self.double_click_action):
                self.double_click_action()
        elif lparam == win32con.WM_RBUTTONUP:
            # TODO: Menu -> logs, quit
            self.show_menu()
        elif lparam == win32con.WM_LBUTTONUP:
            pass

        return True


    def show_menu(self):
        pos = win32gui.GetCursorPos()
        win32gui.SetForegroundWindow(self.hwnd)
        win32gui.TrackPopupMenu(self.menu, win32con.TPM_LEFTALIGN, pos[0], pos[1], 0, self.hwnd, None)
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)


    def execute_menu_option(self, menu_option_id):
        menu_option_action = self.menu_actions[menu_option_id]

        if menu_option_action == self.destroy:
            win32gui.DestroyWindow(self.hwnd)
        else:
            menu_option_action(self)


class MicHandler():
    def __init__(self):
        self.mute_icon = "mute.ico"
        self.unmute_icon = "unmute.ico"

        self.sys_tray_app_change_icon = None
        self.mic_mute = False
        self.mute_mic()

    def change_mic_state(self):
        if self.mic_mute:
            self.unmute_mic()
        else:
            self.mute_mic()
        # self.mic_mute = not self.mic_mute

    def mute_mic(self):
        subprocess.call(["nircmd.exe", "mutesysvolume", "1", "default_record"])

        requests.get("http://IP")

        if callable(self.sys_tray_app_change_icon):
            self.sys_tray_app_change_icon(self.mute_icon)

        self.mic_mute = True

    def unmute_mic(self):
        subprocess.call(["nircmd.exe", "mutesysvolume", "0", "default_record"])

        requests.get("http://IP")

        if callable(self.sys_tray_app_change_icon):
            self.sys_tray_app_change_icon(self.unmute_icon)

        self.mic_mute = False


def micled_start(mic_handler, mute_icon, hover_text, menu_options):
    sys_tray_app = SysTrayApp(mute_icon, hover_text, menu_options, mic_handler.change_mic_state)

    mic_handler.sys_tray_app_change_icon = sys_tray_app.change_icon

    sys_tray_app.start_sys_tray_app()


if __name__ == '__main__':
    MUTE_ICON = "mute.ico"
    HOVER_TEXT = "micLED"
    MENU_OPTIONS = []
    # Example of menu options list
    #
    # def first_option(SysTrayApp):
    #     print("first_option")
    # def second_option(SysTrayApp):
    #     print("second_option")
    #
    # MENU_OPTIONS = [
    #     ['First option', first_option],
    #     ['A sub-menu', [
    #         ['Second option', second_option]
    #     ]]
    # ]

    # Create microphone handler
    mic_handler = MicHandler()

    # Start system tray application
    threading.Thread(target=micled_start, args=(mic_handler, MUTE_ICON, HOVER_TEXT, MENU_OPTIONS)).start()

    # WIN + Z
    COMBINATIONS = [
        {
            pynput.keyboard.Key.cmd,
            pynput.keyboard.KeyCode(char='z')
        },
        {
            pynput.keyboard.Key.cmd,
            pynput.keyboard.KeyCode(char='Z')
        },
    ]

    current = set()

    def on_press(key):
        if any([key in combo for combo in COMBINATIONS]):
            current.add(key)
            if any(all(_key in current for _key in combo) for combo in COMBINATIONS):
                mic_handler.change_mic_state()

    def on_release(key):
        try:
            if any([key in combo for combo in COMBINATIONS]):
                current.remove(key)
        except KeyError:
            pass


    with pynput.keyboard.Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()
