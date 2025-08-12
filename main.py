# main.py
import customtkinter as ctk
import threading

from hotkeys import HotkeyManager
from action_panel import ActionPanel
from macro_editor import MacroEditor

# --- launch logging (Finder 起動でもログが残る) ---
import os, sys, datetime, traceback

LOG_DIR = os.path.expanduser("~/Library/Logs/AuterGUI")
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, "launch.log")

def _log(msg: str):
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now().isoformat()}] {msg}\n")
    except Exception:
        # 予期せぬファイル権限などで書けない場合は無視
        pass

def excepthook(etype, evalue, etb):
    # 例外をログに残し、Finder 起動でも気づけるよう簡易ダイアログを出す
    try:
        _log("UNCAUGHT:\n" + "".join(traceback.format_exception(etype, evalue, etb)))
    finally:
        try:
            import tkinter as tk
            from tkinter import messagebox
            rt = tk.Tk(); rt.withdraw()
            messagebox.showerror("AuterGUI 起動エラー", f"{etype.__name__}: {evalue}\n\n詳しくは {LOG_FILE}")
            rt.destroy()
        except Exception:
            pass
    # 元のハンドラへ委譲
    try:
        sys.__excepthook__(etype, evalue, etb)
    except Exception:
        pass

sys.excepthook = excepthook
_log("==== LAUNCH ====")

import platform, subprocess

# ダークテーマ
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Auto GUI（Win/mac対応・マクロ＋インスペクタ）")
        self.geometry("900x720")

        # 停止フラグ（スレッド間共有）
        self._stop_flag = False

        # タブ
        self.tabs = ctk.CTkTabview(self); self.tabs.pack(fill="both", expand=True)
        self.tab_actions = self.tabs.add("操作")
        self.tab_macro   = self.tabs.add("マクロ")

        # 「操作」タブ
        self.action_panel = ActionPanel(
            self.tab_actions,
            on_start=self._on_start_hotkey,
            stop_flag_ref=lambda: self._stop_flag
        )
        self.action_panel.pack(fill="both", expand=True)

        # 「マクロ」タブ
        self.macro_editor = MacroEditor(
            self.tab_macro,
            stop_flag_ref=lambda: self._stop_flag
        )
        self.macro_editor.pack(fill="both", expand=True)

        # ホットキー管理
        self.hk = HotkeyManager(on_fire=self._fire_action, on_esc=self._on_esc)

        # mac の権限促しは GUI 構築後に遅延実行（Finder 起動時のクラッシュ回避）
        if platform.system() == "Darwin":
            try:
                self.after(800, self._mac_deferred_ax_prompt)
            except Exception as e:
                _log(f"schedule deferred ax prompt failed: {e}")

    # ---- macOS: アクセシビリティ/入力監視の促しは GUI 初期化後に安全に実行 ----
    def _mac_deferred_ax_prompt(self):
        # Finder 起動での EXC_BAD_ACCESS を避けるため、ネイティブ API 呼び出しは避け、設定アプリを開くだけにする
        try:
            if platform.system() == "Darwin":
                # すでに許可済みなら何もしない（厳密判定は避ける）
                subprocess.run(
                    ["open", "x-apple.systempreferences:com.apple.preference.security?Privacy_Accessibility"],
                    check=False
                )
                subprocess.run(
                    ["open", "x-apple.systempreferences:com.apple.preference.security?Privacy_ListenEvent"],
                    check=False
                )
                _log("Opened System Settings for Accessibility/Input Monitoring (deferred)")
        except Exception as e:
            _log(f"Deferred AX prompt failed: {e}")

        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ==== 実行フロー ====
    def _on_start_hotkey(self):
        """Start ボタン→ホットキー待機開始"""
        self._stop_flag = False
        self.hk.start()

    def _on_esc(self):
        """グローバル ESC"""
        self._stop_flag = True

    def _fire_action(self):
        """グローバル起動キー（Alt+Shift または ⌘+Shift）押下時"""
        self._stop_flag = False
        # アクション実行は別スレッド
        threading.Thread(target=self.action_panel.run_worker, daemon=True).start()
        # 起動後は必要ならホットキーを止める（誤発火防止）
        self.hk.stop()

    # ==== 終了処理 ====
    def _on_close(self):
        try:
            self.hk.stop()
        finally:
            self.destroy()

if __name__ == "__main__":
    app = App()
    app.mainloop()
