# utils.py
import platform
import time
import pyautogui as pag

IS_WIN = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"

# mac の ESC ポーリング用
if IS_MAC:
    try:
        import Quartz
    except Exception:
        Quartz = None
else:
    Quartz = None

# Windows の修飾キー状態取得
if IS_WIN:
    import ctypes
    _user32 = ctypes.windll.user32
else:
    _user32 = None

# 共有キー一覧（検索付きUIで使用）
KEY_LIST = [
    '\t', '\n', '\r', ' ', '!', '"', '#', '$', '%', '&', "'", '(',
    ')', '*', '+', ',', '-', '.', '/', '0', '1', '2', '3', '4', '5', '6', '7',
    '8', '9', ':', ';', '<', '=', '>', '?', '@', '[', '\\', ']', '^', '_', '`',
    'a', 'b', 'c', 'd', 'e','f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o',
    'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '{', '|', '}', '~',
    'accept', 'add', 'alt', 'altleft', 'altright', 'apps', 'backspace',
    'browserback', 'browserfavorites', 'browserforward', 'browserhome',
    'browserrefresh', 'browsersearch', 'browserstop', 'capslock', 'clear',
    'convert', 'ctrl', 'ctrlleft', 'ctrlright', 'decimal', 'del', 'delete',
    'divide', 'down', 'end', 'enter', 'esc', 'escape', 'execute', 'f1', 'f10',
    'f11', 'f12', 'f13', 'f14', 'f15', 'f16', 'f17', 'f18', 'f19', 'f2', 'f20',
    'f21', 'f22', 'f23', 'f24', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9',
    'final', 'fn', 'hanguel', 'hangul', 'hanja', 'help', 'home', 'insert', 'junja',
    'kana', 'kanji', 'launchapp1', 'launchapp2', 'launchmail',
    'launchmediaselect', 'left', 'modechange', 'multiply', 'nexttrack',
    'nonconvert', 'num0', 'num1', 'num2', 'num3', 'num4', 'num5', 'num6',
    'num7', 'num8', 'num9', 'numlock', 'pagedown', 'pageup', 'pause', 'pgdn',
    'pgup', 'playpause', 'prevtrack', 'print', 'printscreen', 'prntscrn',
    'prtsc', 'prtscr', 'return', 'right', 'scrolllock', 'select', 'separator',
    'shift', 'shiftleft', 'shiftright', 'sleep', 'space', 'stop', 'subtract', 'tab',
    'up', 'volumedown', 'volumemute', 'volumeup', 'win', 'winleft', 'winright', 'yen',
    'command', 'option', 'optionleft', 'optionright'
]

def modifiers_still_down() -> bool:
    """Alt/Option/Shift が押下中かをざっくり判定"""
    if IS_MAC and Quartz:
        flags = Quartz.CGEventSourceFlagsState(Quartz.kCGEventSourceStateCombinedSessionState)
        return bool(
            (flags & Quartz.kCGEventFlagMaskShift) or
            (flags & Quartz.kCGEventFlagMaskAlternate)  # Option=Alt
        )
    if IS_WIN and _user32:
        VK_SHIFT, VK_MENU = 0x10, 0x12
        return bool((_user32.GetAsyncKeyState(VK_SHIFT) & 0x8000) or
                    (_user32.GetAsyncKeyState(VK_MENU)  & 0x8000))
    return False

def flush_modifiers(timeout=1.5):
    """修飾キーが離れるのを待ち、最後に保険で keyUp を送る（対策）"""
    t0 = time.time()
    while time.time() - t0 < timeout and modifiers_still_down():
        time.sleep(0.02)
    try:
        pag.keyUp('shift')
    except Exception:
        pass
    for alt_name in ('option', 'alt'):
        try:
            pag.keyUp(alt_name)
        except Exception:
            pass

def esc_pressed() -> bool:
    """ESCが押されているか（macのみ直接検知。その他はワーカーの stop_flag を見る）"""
    if IS_MAC and Quartz:
        return Quartz.CGEventSourceKeyState(
            Quartz.kCGEventSourceStateCombinedSessionState, 53
        )
    return False

def busy_wait(seconds: float, stop_flag_getter) -> bool:
    """
    高精度待機。途中で stop_flag_getter() or ESC(mac) で True を返したら中断。
    戻り値: True=中断/停止, False=予定どおり完了
    """
    start = time.perf_counter()
    while time.perf_counter() - start < max(0.0, seconds):
        if stop_flag_getter():
            return True
        if IS_MAC and esc_pressed():
            return True
    return False
    