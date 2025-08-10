# hotkeys.py
import platform
import threading
from typing import Callable, Optional

IS_MAC = platform.system() == "Darwin"

class HotkeyManager:
    """
    起動: Win/Linux -> Alt+Shift, mac -> Cmd+Shift
    停止: ESC
    """
    def __init__(self, on_fire: Callable[[], None], on_esc: Callable[[], None]):
        self.on_fire = on_fire
        self.on_esc = on_esc
        self._listener = None
        self._mac_thread: Optional[threading.Thread] = None

    # ==== mac: Quartzポーリング ====
    if IS_MAC:
        try:
            import Quartz
        except Exception:
            Quartz = None

        def _mac_cmdshift_is_down(self) -> bool:
            if self.Quartz is None:
                return False
            flags = self.Quartz.CGEventSourceFlagsState(
                self.Quartz.kCGEventSourceStateCombinedSessionState
            )
            return bool(
                (flags & self.Quartz.kCGEventFlagMaskCommand) and
                (flags & self.Quartz.kCGEventFlagMaskShift)
            )

        def _mac_esc_is_down(self) -> bool:
            if self.Quartz is None:
                return False
            # ESC 仮想キー: 53
            return self.Quartz.CGEventSourceKeyState(
                self.Quartz.kCGEventSourceStateCombinedSessionState, 53
            )

        def _mac_loop(self):
            # ⌘+Shift 立ち上がり検出
            prev = False
            print("<< Command+Shift を押すと実行開始します。ESC で途中キャンセル >>")
            while self._mac_thread is not None:
                now = self._mac_cmdshift_is_down()
                if now and not prev:
                    self.on_fire()
                    return
                prev = now
                # ESC は worker 側でポーリング（ここでは起動に専念）
                import time; time.sleep(0.03)

        def start(self):
            if self.Quartz is None:
                print("※ mac では pyobjc が必要です: pip install pyobjc")
                return
            self.stop()
            self._mac_thread = threading.Thread(target=self._mac_loop, daemon=True)
            self._mac_thread.start()

        def stop(self):
            self._mac_thread = None

    # ==== Win/Linux: pynput ====
    else:
        try:
            from pynput import keyboard as pk
        except Exception:
            pk = None

        def start(self):
            if self.pk is None:
                print("※ Win/Linux では pynput が必要です: pip install pynput")
                return
            self.stop()
            mapping = {
                '<alt>+<shift>': self.on_fire,
                '<esc>': self.on_esc,
            }
            self._listener = self.pk.GlobalHotKeys(mapping)
            self._listener.start()
            print("<< Alt+Shift を押すと実行開始します。ESC で途中キャンセル >>")

        def stop(self):
            if self._listener:
                try:
                    self._listener.stop()
                except Exception:
                    pass
                self._listener = None
