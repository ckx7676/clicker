import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk, ImageDraw
import threading
import win32gui, win32api, win32con
import time
import random
import heapq
import ctypes
import sys
import keyboard
import winsound
import os
import configparser
import subprocess

DEFAULT_START_HOT_KEY = 'alt+e'
DEFAULT_STOP_HOT_KEY = 'alt+q'
START_HOT_KEY = DEFAULT_START_HOT_KEY
STOP_HOT_KEY = DEFAULT_STOP_HOT_KEY
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.ini")
ICO_FILE = "app.ico"


def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def load_config(config_file):
    global START_HOT_KEY, STOP_HOT_KEY
    if os.path.exists(config_file):
        try:
            config = configparser.ConfigParser()
            config.read(config_file, encoding='utf-8')
            if 'Hotkeys' in config:
                start_hotkey = config['Hotkeys'].get('start', DEFAULT_START_HOT_KEY).strip()
                stop_hotkey = config['Hotkeys'].get('stop', DEFAULT_STOP_HOT_KEY).strip()
                START_HOT_KEY = start_hotkey
                STOP_HOT_KEY = stop_hotkey
        except Exception as e:
            messagebox.showwarning("警告", f"加载配置失败，使用默认值：{str(e)}")
    else:
        try:
            config = configparser.ConfigParser()
            config['Hotkeys'] = {'start': START_HOT_KEY, 'stop': STOP_HOT_KEY}
            with open(config_file, 'w', encoding='utf-8') as f:
                config.write(f)
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}")


load_config(resource_path(CONFIG_FILE))


def sleep(ms, random_delay=False):
    if random_delay:
        ms += (random.randint(3, 10) / 1000)
    x, y = divmod(ms, 1000)
    for _ in range(int(x)): time.sleep(1)
    time.sleep(y / 1000)


class Keyboard:
    def __init__(self, hwnd=None):
        self.hwnd = hwnd

    def kPress(self, key):
        if self.hwnd:
            scancode = win32api.MapVirtualKey(key, 0)
            lParam_down = 1 | (scancode << 16)
            lParam_up = 1 | (scancode << 16) | (1 << 30) | (1 << 31)
            win32api.PostMessage(self.hwnd, win32con.WM_KEYDOWN, key, lParam_down)
            sleep(5, True)
            win32api.PostMessage(self.hwnd, win32con.WM_KEYUP, key, lParam_up)
        else:
            win32api.keybd_event(key, 0, 0, 0)
            sleep(5, True)
            win32api.keybd_event(key, 0, win32con.KEYEVENTF_KEYUP, 0)


class BindWindowButton(tk.Button):
    def __init__(self, master=None, window_label=None, **kwargs):
        super().__init__(master, **kwargs)
        self.window_label = window_label
        self.hwnd = None
        self.last_hwnd_check_time = 0
        self.hwnd_cache_duration = 0.5
        self.hwnd_exists_cache = True
        self.master_hwnd = self.get_top_level_hwnd()
        self.crosshair_img = ImageTk.PhotoImage(self.create_crosshair_image(30, 30))
        self.crosshair_selected_img = ImageTk.PhotoImage(self.create_crosshair_image(30, 30, True))
        self.drag_img = ImageTk.PhotoImage(Image.new("RGBA", (30, 30), (0, 0, 0, 0)))
        self.config(image=self.crosshair_img)
        self.bind("<ButtonPress-1>", self.start_drag)

    def create_crosshair_image(self, width, height, fill=False):
        image = Image.new("RGBA", (width, height), (255, 255, 255, 0))
        draw = ImageDraw.Draw(image)
        if fill:
            draw.chord((0, 0, width - 1, height - 1), 0, 360, fill="skyblue", width=1)
        draw.arc((0, 0, width - 1, height - 1), 0, 360, fill="black", width=1)
        draw.line((width // 2, 0, width // 2, height), fill="black", width=1)
        draw.line((0, height // 2, width, height // 2), fill="black", width=1)
        return image

    def get_top_level_hwnd(self):
        hwnd = self.winfo_id()
        while True:
            parent_hwnd = ctypes.windll.user32.GetParent(hwnd)
            if not parent_hwnd:
                break
            hwnd = parent_hwnd
        return hwnd

    def is_child_window(self, hwnd, parent_hwnd):
        if hwnd == parent_hwnd:
            return True
        while hwnd:
            hwnd = win32gui.GetParent(hwnd)
            if hwnd == parent_hwnd:
                return True
        return False

    def start_drag(self, event):
        self.hwnd = None
        if self.window_label:
            self.window_label.config(text='step1.拖动准心到目标窗口进行绑定')
        self.config(cursor="crosshair")
        self.config(image=self.drag_img)
        self.bind("<Motion>", self.on_drag)
        self.bind("<ButtonRelease-1>", self.end_drag)

    def on_drag(self, event):
        x, y = win32api.GetCursorPos()
        hwnd = win32gui.WindowFromPoint((x, y))
        self.hwnd = hwnd

    def end_drag(self, event):
        self.config(cursor="")
        self.unbind("<Motion>")
        self.unbind("<ButtonRelease-1>")
        if self.is_child_window(self.hwnd, self.master_hwnd):
            self.hwnd = None
        if self.window_label and self.hwnd:
            title = win32gui.GetWindowText(self.hwnd)
            self.window_label.config(text=f'step1.绑定窗口:{title}')
            self.config(image=self.crosshair_selected_img)
        else:
            self.config(image=self.crosshair_img)

    def check_hwnd_exist(self, use_cache=False):
        current_time = time.time()
        if use_cache and self.hwnd_exists_cache:
            if current_time - self.last_hwnd_check_time < self.hwnd_cache_duration:
                return self.hwnd_exists_cache
        exist = ctypes.windll.user32.IsWindow(self.hwnd)
        self.hwnd_exists_cache = exist
        self.last_hwnd_check_time = current_time
        if not exist:
            if self.window_label:
                self.window_label.config(text='step1.拖动准心到目标窗口进行绑定')
            self.config(image=self.crosshair_img)
            self.hwnd = None
        return exist

    def set_state(self, state):
        self.config(state=state)
        if state == "disabled":
            self.unbind("<ButtonPress-1>")
        else:
            self.bind("<ButtonPress-1>", self.start_drag)


class WindowFinder(tk.Frame):
    def __init__(self, master=None, cnf={}, **kwargs):
        super().__init__(master, cnf, **kwargs)
        self.window_label = tk.Label(self, text='step1.拖动准心到目标窗口进行绑定', font=("Arial", 12), width=30,
                                     anchor='w')
        self.window_label.pack(pady=10)

        self.bind_button = BindWindowButton(self, self.window_label)
        self.bind_button.pack()

    def set_state(self, state):
        self.bind_button.set_state(state)


class KeyButton(tk.Button):
    def __init__(self, master=None, root=None, cnf={}, width=20, height=1, **kwargs):
        super().__init__(master, cnf, **kwargs)
        self.key = None
        self.key_code = None
        self.root = root
        self.config(text='点击开始按键', width=width, height=height)
        self.config(command=lambda: self.start_listening(self))

    def start_listening(self, button):
        def on_key_press(event):
            button.key = event.keysym
            button.key_code = event.keycode
            button.config(text=f'按键: {button.key}', bg='skyblue')
            self.root.unbind('<Key>')

        def on_click(event):
            if event.widget is not button:
                if button.key is None:
                    button.config(text='点击开始按键', bg='SystemButtonFace')
                self.root.unbind('<Key>')
                self.root.unbind('<Button-1>')

        self.root.bind('<Key>', on_key_press)
        self.root.bind('<Button-1>', on_click)
        button.config(text='等待按键...', bg='SystemButtonFace')
        button.key = None
        button.key_code = None


class DelayEntry(tk.Entry):
    def __init__(self, master=None, root=None, cnf={}, **kwargs):
        super().__init__(master, cnf, **kwargs)
        self.root = root
        self.bind("<FocusIn>", self.start_listening)

    def start_listening(self, event):
        def on_click(event):
            if event.widget is not self:
                self.root.focus()
                self.root.unbind('<Button-1>')

        self.root.bind('<Button-1>', on_click)


class KeyDelayFrame(tk.Frame):
    def __init__(self, master=None, root=None, cnf={}, **kwargs):
        super().__init__(master, cnf, **kwargs)
        root = root or master
        self.key_button = KeyButton(self, root)
        self.key_button.pack(side=tk.LEFT, padx=5)

        self.vcmd = self.register(self.validate_non_negative_integers_input)
        self.delay_entry = DelayEntry(self, root, width=8, justify='center', validate="key",
                                      validatecommand=(self.vcmd, '%P'))
        self.delay_entry.insert(0, "100")
        self.delay_entry.bind('<Return>', lambda event: event.widget.master.focus())
        self.delay_entry.bind('<Escape>', lambda event: event.widget.master.focus())
        self.delay_entry.pack(side=tk.LEFT, padx=5)
        self.pack(anchor='w')

    def validate_non_negative_integers_input(self, text):
        if not text or (text.isdigit() and 0 <= int(text) < 1e7):
            return True
        return False

    def set_state(self, state):
        self.key_button.config(state=state)
        self.delay_entry.config(state=state)


class KeyListener(tk.Frame):
    def __init__(self, master=None, root=None, key_num=10, cnf={}, **kwargs):
        super().__init__(master, cnf, **kwargs)
        root = root or master

        tk.Label(self, text="step2.绑定键盘按键", font=("Arial", 12), width=30, anchor='w').pack(pady=10)
        title_frame = tk.Frame(self)
        title_frame.pack(anchor='w')
        key_label = tk.Label(title_frame, text="按键", width=20)
        key_label.pack(side=tk.LEFT, padx=5)
        delay_label = tk.Label(title_frame, text="间隔(毫秒)", width=8)
        delay_label.pack(side=tk.LEFT, padx=5)

        self.key_list = list()
        for _ in range(key_num):
            key_frame = KeyDelayFrame(self, root)
            self.key_list.append(key_frame)

    def set_state(self, state):
        for key_frame in self.key_list:
            if state == 'disabled':
                if key_frame.key_button.key:
                    key_frame.set_state(state)
            else:
                key_frame.set_state(state)


class PlayPauseButton(tk.Canvas):
    def __init__(self, master=None, window_finder=None, key_listener=None, width=45, height=45, cursor='hand2',
                 **kwargs):
        super().__init__(master, width=width, height=height, cursor=cursor, **kwargs)
        self.window_finder = window_finder
        self.key_listener = key_listener
        self.runing = 0
        self.create_play_icon()
        self.bind("<Button-1>", self.toggle_state)
        keyboard.add_hotkey(START_HOT_KEY, lambda: self.toggle_start())
        keyboard.add_hotkey(STOP_HOT_KEY, lambda: self.toggle_stop())

    def create_play_icon(self):
        self.delete("all")
        self.create_polygon(10, 10, 10, 38, 38, 24, fill="green", outline="white")

    def create_pause_icon(self):
        self.delete("all")
        self.create_rectangle(10, 10, 36, 38, fill="red", outline="white")

    def toggle_stop(self):
        if self.runing:
            self.create_play_icon()
            self.window_finder.set_state("normal")
            self.key_listener.set_state("normal")
            winsound.PlaySound(resource_path("stop.wav"), winsound.SND_FILENAME | winsound.SND_ASYNC)
            self.runing = 0

    def toggle_start(self):
        if not self.runing:
            ready, hwnd, loop_keys = self.check_run_ready()
            if not ready:
                return
            self.create_pause_icon()
            self.window_finder.set_state("disabled")
            self.key_listener.set_state("disabled")
            winsound.PlaySound(resource_path("start.wav"), winsound.SND_FILENAME | winsound.SND_ASYNC)
            self.runing = 1
            threading.Thread(target=self.run_loop, args=(hwnd, loop_keys), daemon=True).start()

    def toggle_state(self, event):
        if not self.runing:
            self.toggle_start()
        else:
            self.toggle_stop()

    def check_run_ready(self):
        ready, hwnd, loop_keys = False, None, list()
        bind_button = self.window_finder.bind_button
        bind_button.check_hwnd_exist()
        hwnd = bind_button.hwnd
        if not hwnd:
            return ready, hwnd, loop_keys
        key_list = self.key_listener.key_list
        for key_frame in key_list:
            press_key = key_frame.key_button.key
            press_key_code = key_frame.key_button.key_code
            delay_ms = key_frame.delay_entry.get()
            if not press_key:
                continue
            if not delay_ms:
                key_frame.delay_entry.insert(0, "0")
                delay_ms = 0
            loop_keys.append((press_key_code, int(delay_ms)))
        if not loop_keys:
            return ready, hwnd, loop_keys
        ready = True
        return ready, hwnd, loop_keys

    def run_loop(self, hwnd, loop_keys):
        try:
            bind_button = self.window_finder.bind_button
            kb = Keyboard(hwnd)
            key_data_heap = list()
            for idx, (press_key_code, sleep_time) in enumerate(loop_keys):
                heapq.heappush(key_data_heap, (idx, press_key_code, sleep_time))
            while self.runing and hwnd == bind_button.hwnd and bind_button.check_hwnd_exist(True):
                next_press_time, press_key_code, sleep_time = heapq.heappop(key_data_heap)
                while (time.time() * 1000) < next_press_time and self.runing and bind_button.check_hwnd_exist(True):
                    sleep(5, True)
                kb.kPress(press_key_code)
                next_press_time = time.time() * 1000 + sleep_time
                heapq.heappush(key_data_heap, (next_press_time, press_key_code, sleep_time))
                sleep(5, True)
        finally:
            self.toggle_stop()


class Page(tk.Frame):
    def __init__(self, master=None, root=None, cnf={}, **kwargs):
        super().__init__(master, cnf, **kwargs)
        root = root or master

        self.window_finder = WindowFinder(master)
        self.window_finder.pack()

        self.key_listener = KeyListener(master, root, key_num=12)
        self.key_listener.pack()

        self.label = tk.Label(master, text=f"step3.点击▷/□启动({START_HOT_KEY})或暂停({STOP_HOT_KEY})",
                              font=("Arial", 12),
                              width=30, anchor='w').pack(pady=10)
        self.play_pause_button = PlayPauseButton(master, self.window_finder, self.key_listener)
        self.play_pause_button.pack()

    def remove(self):
        self.play_pause_button.toggle_stop()
        self.play_pause_button.toggle_start = lambda: 0
        self.play_pause_button.toggle_stop = lambda: 0


class SimpleHotkeySettings:
    def __init__(self, parent=None, on_save_callback=None):
        self.parent = parent
        self.on_save_callback = on_save_callback
        if self.parent and not self.parent.winfo_exists():
            self.parent = None
        self.config_file = resource_path(CONFIG_FILE)
        self.setting_window = None
        self.DEFAULT_HOTKEYS = {
            'start': DEFAULT_START_HOT_KEY,
            'stop': DEFAULT_STOP_HOT_KEY
        }
        self.hotkeys = self.DEFAULT_HOTKEYS.copy()
        self.modify_buttons = []
        self.load_config()
        self.create_window()

    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                config = configparser.ConfigParser()
                config.read(self.config_file, encoding='utf-8')
                if 'Hotkeys' in config:
                    start_hotkey = config['Hotkeys'].get('start', self.DEFAULT_HOTKEYS['start']).strip()
                    stop_hotkey = config['Hotkeys'].get('stop', self.DEFAULT_HOTKEYS['stop']).strip()
                    self.hotkeys['start'] = start_hotkey if start_hotkey else self.DEFAULT_HOTKEYS['start']
                    self.hotkeys['stop'] = stop_hotkey if stop_hotkey else self.DEFAULT_HOTKEYS['stop']
            except Exception as e:
                messagebox.showwarning("警告", f"加载配置失败，使用默认值：{str(e)}")
                self.save_config()
        else:
            self.save_config()

    def save_config(self):
        if not self.hotkeys['start'] or not self.hotkeys['stop']:
            messagebox.showwarning("警告", "快捷键不能为空，保存失败！")
            return False
        try:
            config = configparser.ConfigParser()
            config['Hotkeys'] = self.hotkeys
            with open(self.config_file, 'w', encoding='utf-8') as f:
                config.write(f)
            return True
        except Exception as e:
            messagebox.showerror("错误", f"保存失败：{str(e)}")
            return False

    def create_window(self):
        self.window = tk.Toplevel(self.parent) if self.parent else tk.Tk()
        self.window.title("快捷键设置")
        window_width = 300
        window_height = 200

        try:
            ico_path = resource_path(ICO_FILE)
            if os.path.isfile(ico_path):
                self.window.iconbitmap(ico_path)
        except Exception as e:
            pass

        self.window.geometry(f"{window_width}x{window_height}")
        self.window.resizable(False, False)

        self.window.protocol("WM_DELETE_WINDOW", self.on_window_close)

        self.window.update_idletasks()
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        tk.Label(self.window, text="快捷键设置", font=("Arial", 12, "bold")).pack(pady=10)

        start_frame = tk.Frame(self.window)
        start_frame.pack(pady=5)
        tk.Label(start_frame, text="启动:", width=6).pack(side=tk.LEFT)
        self.start_var = tk.StringVar(value=self.hotkeys['start'])
        tk.Label(start_frame, textvariable=self.start_var, width=15,
                 relief="sunken", padx=5, pady=2).pack(side=tk.LEFT, padx=5)
        start_btn = tk.Button(start_frame, text="修改", width=6,
                              command=lambda: self.set_hotkey('start'))
        start_btn.pack(side=tk.LEFT)
        self.modify_buttons.append(start_btn)

        stop_frame = tk.Frame(self.window)
        stop_frame.pack(pady=5)
        tk.Label(stop_frame, text="停止:", width=6).pack(side=tk.LEFT)
        self.stop_var = tk.StringVar(value=self.hotkeys['stop'])
        tk.Label(stop_frame, textvariable=self.stop_var, width=15,
                 relief="sunken", padx=5, pady=2).pack(side=tk.LEFT, padx=5)
        stop_btn = tk.Button(stop_frame, text="修改", width=6,
                             command=lambda: self.set_hotkey('stop'))
        stop_btn.pack(side=tk.LEFT)
        self.modify_buttons.append(stop_btn)

        btn_frame = tk.Frame(self.window)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="重置", width=8, command=self.reset_hotkeys).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="保存", width=8, command=self.save).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="关闭", width=8, command=self.on_window_close).pack(side=tk.LEFT, padx=5)

    def reset_hotkeys(self):
        if self.setting_window and self.setting_window.winfo_exists():
            self.setting_window.grab_release()
            self.setting_window.destroy()
            self.setting_window = None

        self.hotkeys = self.DEFAULT_HOTKEYS.copy()
        self.start_var.set(self.DEFAULT_HOTKEYS['start'])
        self.stop_var.set(self.DEFAULT_HOTKEYS['stop'])

        messagebox.showinfo("重置成功",
                            "所有快捷键已恢复为初始默认值！\n启动：alt+e | 停止：alt+q\n\n请点击【保存】按钮生效修改。")

    def normalize_key(self, keysym):
        key_map = {
            'return': 'enter',
            'escape': 'esc',
            'delete': 'del',
            'backspace': 'backspace',
            'tab': 'tab',
            'space': 'space',
            'kp_0': '0', 'kp_1': '1', 'kp_2': '2', 'kp_3': '3', 'kp_4': '4',
            'kp_5': '5', 'kp_6': '6', 'kp_7': '7', 'kp_8': '8', 'kp_9': '9',
            'kp_add': '+', 'kp_subtract': '-', 'kp_multiply': '*', 'kp_divide': '/',
            'kp_decimal': '.', 'kp_enter': 'enter'
        }
        if keysym.startswith('f') and keysym[1:].isdigit():
            return keysym.upper()
        if len(keysym) == 1 and keysym.isalpha():
            return keysym.lower()
        return key_map.get(keysym.lower(), keysym.lower())

    def set_hotkey(self, hotkey_type):
        if self.setting_window and self.setting_window.winfo_exists():
            self.setting_window.lift()
            self.setting_window.focus_force()
            return

        self.setting_window = tk.Toplevel(self.window)
        self.setting_window.title(f"设置{'启动' if hotkey_type == 'start' else '停止'}快捷键")
        setting_window_width = 300
        setting_window_height = 150
        self.setting_window.geometry(f"{setting_window_width}x{setting_window_height}")
        self.setting_window.resizable(False, False)

        try:
            ico_path = resource_path(ICO_FILE)
            if os.path.isfile(ico_path):
                self.setting_window.iconbitmap(ico_path)
        except Exception as e:
            pass

        self.setting_window.grab_set()
        self.setting_window.focus_force()

        self.disable_modify_buttons(True)

        self.setting_window.update_idletasks()
        x = self.window.winfo_x() + (self.window.winfo_width() - setting_window_width) // 2
        y = self.window.winfo_y() + (self.window.winfo_height() - setting_window_height) // 2
        self.setting_window.geometry(f"{setting_window_width}x{setting_window_height}+{x}+{y}")

        tip_label = tk.Label(self.setting_window,
                             text="请按下新的快捷键组合\n（按ESC取消，按Enter确认）",
                             pady=5)
        tip_label.pack()

        key_label = tk.Label(self.setting_window, text="等待按键...", font=("Arial", 10, "bold"))
        key_label.pack(pady=5)

        pressed_modifiers = set()
        pressed_key = None
        is_capturing = True

        modifier_mapping = {
            'control_l': 'ctrl', 'control_r': 'ctrl',
            'shift_l': 'shift', 'shift_r': 'shift',
            'alt_l': 'alt', 'alt_r': 'alt',
            'win_l': 'win', 'win_r': 'win'
        }

        def key_event_record(event):
            if not is_capturing or not key_label.winfo_exists():
                return

            nonlocal pressed_key
            keysym = event.keysym.lower()

            if keysym == 'escape':
                close_setting_window()
                return
            elif keysym == 'return':
                if pressed_modifiers or pressed_key:
                    finish_capture()
                return

            if keysym in modifier_mapping:
                pressed_modifiers.add(modifier_mapping[keysym])
            else:
                normalized_key = self.normalize_key(keysym)
                if (normalized_key not in ['??', 'caps_lock', 'num_lock', 'scroll_lock']
                        and not normalized_key.startswith('iso_')):
                    pressed_key = normalized_key

            update_display()

        def update_display():
            if not is_capturing or not key_label.winfo_exists():
                return

            modifier_order = ['ctrl', 'alt', 'shift', 'win']
            hotkey_parts = [mod for mod in modifier_order if mod in pressed_modifiers]

            if pressed_key:
                hotkey_parts.append(pressed_key)

            display_text = '+'.join(hotkey_parts) if hotkey_parts else "等待按键..."

            if hotkey_parts and not pressed_key:
                display_text += "（无效：需包含主键）"

            current_hotkey = '+'.join(hotkey_parts) if (hotkey_parts and pressed_key) else ""
            if current_hotkey and current_hotkey == self.hotkeys['stop' if hotkey_type == 'start' else 'start']:
                display_text += "（重复：已被占用）"

            key_label.config(text=display_text)

        def finish_capture():
            nonlocal is_capturing
            is_capturing = False

            modifier_order = ['ctrl', 'alt', 'shift', 'win']
            hotkey_parts = [mod for mod in modifier_order if mod in pressed_modifiers]
            if pressed_key:
                hotkey_parts.append(pressed_key)
            hotkey = '+'.join(hotkey_parts)

            if not hotkey:
                messagebox.showwarning("提示", "快捷键不能为空！")
                is_capturing = True
                return

            if pressed_key is None:
                messagebox.showwarning("提示", "快捷键必须包含一个非修饰键（如A/1/F5等）！")
                is_capturing = True
                return

            other_type = 'stop' if hotkey_type == 'start' else 'start'
            if hotkey == self.hotkeys[other_type]:
                messagebox.showwarning("提示", f"该快捷键已被{('停止' if other_type == 'stop' else '启动')}功能占用！")
                is_capturing = True
                return

            if hotkey_type == 'start':
                self.start_var.set(hotkey)
                self.hotkeys['start'] = hotkey
            else:
                self.stop_var.set(hotkey)
                self.hotkeys['stop'] = hotkey

            messagebox.showinfo("成功", f"快捷键已设置为: {hotkey}")
            close_setting_window()

        def close_setting_window():
            nonlocal is_capturing
            is_capturing = False

            try:
                self.setting_window.unbind('<KeyPress>')
                self.setting_window.unbind('<KeyRelease>')
                self.setting_window.grab_release()
                self.setting_window.destroy()
            except:
                pass
            self.setting_window = None

            self.disable_modify_buttons(False)
            if self.window and self.window.winfo_exists():
                self.window.focus_force()

        self.setting_window.bind('<KeyPress>', key_event_record)
        self.setting_window.bind('<KeyRelease>', key_event_record)
        self.setting_window.protocol("WM_DELETE_WINDOW", close_setting_window)

        btn_frame = tk.Frame(self.setting_window)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="确认", command=finish_capture, width=8).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="取消", command=close_setting_window, width=8).pack(side=tk.LEFT, padx=5)

    def disable_modify_buttons(self, disable=True):
        state = "disabled" if disable else "normal"
        for btn in self.modify_buttons:
            if btn.winfo_exists():
                btn.config(state=state)

    def on_window_close(self):
        try:
            if self.setting_window and self.setting_window.winfo_exists():
                self.setting_window.grab_release()
                self.setting_window.destroy()
        except:
            pass

        try:
            if self.parent:
                self.window.destroy()
            else:
                self.window.quit()
                self.window.destroy()
        except:
            pass

    def save(self):
        if self.save_config():
            messagebox.showinfo("成功", "快捷键设置已保存，应用将重启...")

            if self.on_save_callback:
                self.window.after(100, self.on_save_callback)
            else:
                self.restart_app()

    def restart_app(self):
        try:
            keyboard.unhook_all()
            if self.parent:
                self.parent.destroy()
            else:
                self.window.destroy()

            python = sys.executable
            subprocess.Popen([python] + sys.argv)

            if self.parent:
                self.parent.quit()
            else:
                self.window.quit()

        except Exception as e:
            messagebox.showerror("错误", f"重启应用失败：{str(e)}")

    def run(self):
        if not self.parent:
            self.window.mainloop()


class MultiTabs(ttk.Notebook):
    def __init__(self, master=None, tab=None, root=None, init_tab_amount=1, **kw):
        super().__init__(master, **kw)
        self.tab = tab
        self.root = root
        self.page_list = list()
        self.max_page_amount = 8

        self.add_tab_frame = tk.Frame(self)
        self.add(self.add_tab_frame, text="+")
        self.bind("<<NotebookTabChanged>>", self.tab_changed)

        for _ in range(init_tab_amount):
            self.add_new_tab()
        self.select(0)

    def tab_changed(self, event):
        current_tab_index = self.index("current")
        add_tab_index = self.index(self.add_tab_frame)
        if current_tab_index == add_tab_index:
            self.add_new_tab()

    def add_new_tab(self):
        if not self.tab:
            return
        if len(self.page_list) > self.max_page_amount:
            self.select(self.max_page_amount)
            return
        new_page = tk.Frame(self)
        new_page.pack()
        self.tab(new_page, self.root).pack()
        new_page_insert_index = self.index(self.add_tab_frame)
        self.insert(new_page_insert_index, new_page, text=f"Tab {new_page_insert_index}")
        self.add_close_button(new_page)
        self.page_list.append(new_page)
        self.select(new_page_insert_index)

    def add_close_button(self, tab):
        close_button = tk.Button(tab, text="x", command=lambda: self.close_tab(tab))
        close_button.place(relx=1, rely=0, anchor=tk.NE)

    def close_tab(self, tab):
        tab_index = self.index(tab)
        add_tab_index = self.index(self.add_tab_frame)
        if tab_index != add_tab_index:
            self.forget(tab_index)
            self.page_list.remove(tab)
            page = tab.children.get('!page')
            if page:
                page.remove()
            if tab_index > 0:
                self.select(tab_index - 1)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("点点侠")
        self.root.iconbitmap(resource_path(ICO_FILE))

        window_width = 350
        window_height = 670
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window_x = max(0, screen_width - window_width - 160)
        window_y = max(0, (screen_height - window_height) // 2 - 140)
        self.root.geometry(f'{window_width}x{window_height}+{window_x}+{window_y}')
        self.root.resizable(0, 0)

        self.control_frame = tk.Frame(root, height=40)
        self.control_frame.pack(fill=tk.X, pady=(5, 0))

        self.hotkey_manager = None
        self.setup_hotkey_button()

        self.pages = MultiTabs(root, Page, root=root)
        self.pages.pack(fill=tk.BOTH, expand=True)

        style = ttk.Style()
        style.layout("TNotebook", [])

    def setup_hotkey_button(self):
        self.settings_btn = tk.Button(
            self.control_frame,
            text=" ⚙ 快捷键设置",
            command=self.open_hotkey_settings,
            font=("Arial", 9),
            width=12
        )
        self.settings_btn.pack(side=tk.LEFT, padx=0)

    def open_hotkey_settings(self):
        if self.hotkey_manager is None:
            self.hotkey_manager = SimpleHotkeySettings(parent=self.root, on_save_callback=self.restart_app)

        if not hasattr(self.hotkey_manager, 'window') or self.hotkey_manager.window is None:
            self.hotkey_manager.create_window()

        try:
            self.hotkey_manager.window.deiconify()
            self.hotkey_manager.window.lift()
            self.hotkey_manager.window.focus_force()
        except Exception as e:
            self.hotkey_manager.create_window()
            self.hotkey_manager.window.deiconify()

    def restart_app(self):
        try:
            if self.hotkey_manager and hasattr(self.hotkey_manager, 'window'):
                try:
                    if self.hotkey_manager.window and self.hotkey_manager.window.winfo_exists():
                        self.hotkey_manager.window.destroy()
                except:
                    pass
            cmd = []
            if hasattr(sys, 'frozen') or '__compiled__' in globals():
                if 'pyinstaller' in sys.executable.lower():
                    cmd = [sys.executable] + sys.argv[1:]
                else:
                    current_exe = os.path.abspath(sys.argv[0])
                    if not os.path.exists(current_exe):
                        raise FileNotFoundError(f"Nuitka编译的EXE不存在：{current_exe}")
                    cmd = [current_exe] + sys.argv[1:]
            else:
                python = sys.executable
                script = os.path.abspath(sys.argv[0])
                cmd = [python, script] + sys.argv[1:]
            creationflags = 0
            if sys.platform == 'win32':
                creationflags = 0x00000008
            subprocess.Popen(
                cmd,
                creationflags=creationflags,
                cwd=os.getcwd()
            )
            self.root.quit()
            self.root.destroy()
        except Exception as e:
            messagebox.showerror("错误", f"重启应用失败：{str(e)}")

    def on_closing(self):
        if self.hotkey_manager and hasattr(self.hotkey_manager, 'window'):
            try:
                if (self.hotkey_manager.window and
                        hasattr(self.hotkey_manager.window, 'winfo_exists') and
                        self.hotkey_manager.window.winfo_exists()):
                    self.hotkey_manager.window.destroy()
            except:
                pass
        self.root.destroy()


def main():
    root = tk.Tk()
    app = App(root)

    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == '__main__':
    if ctypes.windll.shell32.IsUserAnAdmin():
        main()
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, None, None, 1)
        sys.exit(0)
