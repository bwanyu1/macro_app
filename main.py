# main.py
import customtkinter as ctk
import threading

from hotkeys import HotkeyManager
from action_panel import ActionPanel
from macro_editor import MacroEditor

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
