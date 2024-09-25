import tkinter as tk
from tkinter import ttk
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

START_HOT_KEY = 'alt+e'
STOP_HOT_KEY = 'alt+q'


def sleep(ms, random_delay=False):
    if random_delay:
        ms += (random.randint(3, 20) / 1000)
    x, y = divmod(ms, 1000)
    for _ in range(int(x)): time.sleep(1)
    time.sleep(y / 1000)


def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


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

    def check_hwnd_exist(self):
        exist = ctypes.windll.user32.IsWindow(self.hwnd)
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
            while self.runing and hwnd == bind_button.hwnd and bind_button.check_hwnd_exist():
                next_press_time, press_key_code, sleep_time = heapq.heappop(key_data_heap)
                while (time.time() * 1000) < next_press_time and self.runing and hwnd == bind_button.hwnd and \
                        bind_button.check_hwnd_exist():
                    sleep(5)
                kb.kPress(press_key_code)
                next_press_time = time.time() * 1000 + sleep_time
                heapq.heappush(key_data_heap, (next_press_time, press_key_code, sleep_time))
                sleep(5)
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
        self.root.iconbitmap(resource_path("app.ico"))
        window_width = 350
        window_height = 650
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window_x = max(0, screen_width - window_width - 160)
        window_y = max(0, (screen_height - window_height) // 2 - 140)
        self.root.geometry(f'{window_width}x{window_height}+{window_x}+{window_y}')
        self.root.resizable(0, 0)

        self.pages = MultiTabs(root, Page, root=root)
        self.pages.pack(fill=tk.BOTH, expand=True)

        style = ttk.Style()
        style.layout("TNotebook", [])


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == '__main__':
    if ctypes.windll.shell32.IsUserAnAdmin():
        main()
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
