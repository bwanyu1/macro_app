# action_panel.py
import customtkinter as ctk
import pyautogui as pag
import threading
import time
from typing import Callable

from utils import KEY_LIST, flush_modifiers, busy_wait

class ActionPanel(ctk.CTkFrame):
    """
    「操作」タブ：クリック/キー、短押し/長押し、追従/座標、検索付きキー選択
    外部から hotkey_manager.start() を呼ぶため、on_start をコールバックで受け取る。
    """
    def __init__(self, master, on_start: Callable[[], None], stop_flag_ref, **kwargs):
        super().__init__(master, **kwargs)
        self.on_start = on_start
        self.stop_flag_ref = stop_flag_ref  # lambda: bool
        self.use_follow_mouse = True
        self.selected_key = "enter"

        # UI
        ctk.CTkLabel(self, text="アクションを選んでください").pack(padx=20, pady=(20, 6))
        self.action_option = ctk.CTkOptionMenu(
            self, values=["左クリック","右クリック","ダブルクリック","キー入力"],
            command=self._on_action_change
        )
        self.action_option.set("左クリック"); self.action_option.pack(padx=20, pady=(0, 16))

        ctk.CTkLabel(self, text="押し方を選んでください").pack(padx=20, pady=(6, 6))
        self.press_option = ctk.CTkOptionMenu(
            self, values=["短押し","長押し"],
            command=self._on_press_change
        )
        self.press_option.set("短押し"); self.press_option.pack(padx=20, pady=(0, 16))

        # マウス追従
        self.follow_mouse_checkbox = ctk.CTkCheckBox(
            self, text="マウス追従モード", command=self._toggle_follow
        )
        self.follow_mouse_checkbox.select()
        self.follow_mouse_checkbox.pack(padx=20, pady=(6,10))

        # 座標指定
        self.frame_coordinates = ctk.CTkFrame(self)
        ctk.CTkLabel(self.frame_coordinates, text="座標設定").pack(padx=10, pady=(10,6))
        fr_xy = ctk.CTkFrame(self.frame_coordinates); fr_xy.pack(padx=10, pady=(6,10))
        ctk.CTkLabel(fr_xy, text="X:").grid(row=0, column=0, padx=6)
        self.entry_x = ctk.CTkEntry(fr_xy, width=90); self.entry_x.grid(row=0, column=1, padx=6); self.entry_x.insert(0, "960")
        ctk.CTkLabel(fr_xy, text="Y:").grid(row=0, column=2, padx=6)
        self.entry_y = ctk.CTkEntry(fr_xy, width=90); self.entry_y.grid(row=0, column=3, padx=6); self.entry_y.insert(0, "540")
        ctk.CTkButton(self.frame_coordinates, text="現在のマウス位置を取得", command=self._update_mouse_position).pack(padx=10, pady=(6,10))
        self.frame_coordinates.pack_forget()

        # 短押し設定（テキストボックス）
        self.frame_repeat = ctk.CTkFrame(self)
        ctk.CTkLabel(self.frame_repeat, text="何回繰り返すか（例: 1000）").pack(padx=10, pady=(10,6))
        self.entry_repeat_count = ctk.CTkEntry(self.frame_repeat); self.entry_repeat_count.insert(0, "1"); self.entry_repeat_count.pack(padx=10, pady=(0,10))
        ctk.CTkLabel(self.frame_repeat, text="間隔（秒・0.001〜OK）").pack(padx=10, pady=(8,6))
        self.entry_repeat_interval = ctk.CTkEntry(self.frame_repeat); self.entry_repeat_interval.insert(0, "0.005"); self.entry_repeat_interval.pack(padx=10, pady=(0,10))
        self.frame_repeat.pack(padx=20, pady=(10,10))

        # 長押し設定
        self.frame_long = ctk.CTkFrame(self)
        ctk.CTkLabel(self.frame_long, text="長押し秒数").pack(padx=10, pady=(10,6))
        self.entry_seconds = ctk.CTkEntry(self.frame_long); self.entry_seconds.insert(0, "2"); self.entry_seconds.pack(padx=10, pady=(0,10))
        self.frame_long.pack_forget()

        # キー選択（キー入力時のみ表示）
        self.frame_key = ctk.CTkFrame(self)
        ctk.CTkLabel(self.frame_key, text="キーを検索・選択").pack(padx=20, pady=(10,6))
        # ComboBox が使える環境
        try:
            from customtkinter import CTkComboBox
            self.key_combo = CTkComboBox(self.frame_key, values=KEY_LIST, command=self._on_key_select)
            self.key_combo.set(self.selected_key)
            self.key_combo.pack(padx=20, pady=(0,12), fill="x")
        except Exception:
            # フォールバック: Entry + 簡易
            self.key_combo = ctk.CTkEntry(self.frame_key); self.key_combo.insert(0, self.selected_key)
            self.key_combo.pack(padx=20, pady=(0,12), fill="x")
            self.key_combo.bind("<FocusOut>", lambda e: self._on_key_select(self.key_combo.get()))
        self.frame_key.pack_forget()

        # 実行ボタン
        ctk.CTkButton(self, text="Start (⌘+Shift: mac / Alt+Shift: Win)", command=self.on_start).pack(padx=20, pady=(10,10))
        ctk.CTkLabel(self, text="※実行中は ESC でキャンセルできます", text_color="gray").pack(padx=20, pady=(0,20))

    # ===== UIハンドラ =====
    def _on_action_change(self, choice):
        if choice == "キー入力":
            self.frame_key.pack(padx=20, pady=(10,20))
        else:
            self.frame_key.pack_forget()

    def _on_press_change(self, choice):
        if choice == "短押し":
            self.frame_repeat.pack(padx=20, pady=(10,10))
            self.frame_long.pack_forget()
        else:
            self.frame_repeat.pack_forget()
            self.frame_long.pack(padx=20, pady=(10,10))

    def _toggle_follow(self):
        self.use_follow_mouse = bool(self.follow_mouse_checkbox.get())
        if self.use_follow_mouse:
            self.frame_coordinates.pack_forget()
        else:
            self.frame_coordinates.pack(padx=20, pady=(10,10))

    def _update_mouse_position(self):
        x, y = pag.position()
        self.entry_x.delete(0, "end"); self.entry_x.insert(0, str(x))
        self.entry_y.delete(0, "end"); self.entry_y.insert(0, str(y))

    def _on_key_select(self, value):
        self.selected_key = (value or "enter").strip()

    # ===== 実行ロジック（ホットキーから呼び出される想定） =====
    def run_worker(self):
        """Alt+Shift(macは⌘+Shift)の起動後に実行される処理本体"""
        action = self.action_option.get()
        press  = self.press_option.get()

        # vals
        def _f(txt, default):
            try:
                v = float(txt.get()); return max(0.0, v)
            except: return default
        def _i(txt, default):
            try:
                v = int(txt.get()); return max(0, v)
            except: return default

        seconds = _f(self.entry_seconds, 1.0)
        count   = _i(self.entry_repeat_count, 1)
        interval= _f(self.entry_repeat_interval, 0.005)

        if not self.use_follow_mouse:
            try:
                fx = int(self.entry_x.get()); fy = int(self.entry_y.get())
            except:
                fx, fy = 960, 540

        # 修飾キー離れ待ち（対策）
        flush_modifiers()

        stop = self.stop_flag_ref

        # クリック
        if action in ("左クリック","右クリック","ダブルクリック"):
            if press == "短押し":
                for _ in range(max(1, count)):
                    if stop(): break
                    x, y = pag.position() if self.use_follow_mouse else (fx, fy)
                    if action == "左クリック": pag.click(x, y)
                    elif action == "右クリック": pag.rightClick(x, y)
                    else: pag.doubleClick(x, y)
                    if busy_wait(interval, stop): break
            else:
                x, y = pag.position() if self.use_follow_mouse else (fx, fy)
                pag.moveTo(x, y); pag.mouseDown()
                t0 = time.time()
                while time.time() - t0 < seconds:
                    if stop():
                        pag.mouseUp(); break
                    time.sleep(0.01)
                else:
                    pag.mouseUp()

        # キー
        else:
            key = self.selected_key or "enter"
            if press == "短押し":
                for _ in range(max(1, count)):
                    if stop(): break
                    pag.press(key)
                    if busy_wait(interval, stop): break
            else:
                pag.keyDown(key)
                t0 = time.time()
                while time.time() - t0 < seconds:
                    if stop():
                        pag.keyUp(key); break
                    time.sleep(0.01)
                else:
                    pag.keyUp(key)
        print("完了！")
