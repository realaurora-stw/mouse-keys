#!/usr/bin/env python3
# WASD/Arrows continuous mouse move + Shift left-click + 1/2 vertical scrolling + 3/4 horizontal scrolling + F middle-click.
# - Uses a real low-level keyboard hook (blocks keys from other apps) except when Ctrl or Win is held
# - Tracks keydown/keyup in the hook and uses that state for movement
# - Shift + WASD/Arrows -> 75% slower (i.e. 25% speed) while Shift is held
# - Ctrl disables script actions and allows normal key behavior (Ctrl+A/C/V etc)
# - Holding either Windows key (LWin/RWin) also disables script-blocking so Win+Shift+S works
# - '[' sets system volume to 24% instantly; ']' sets system volume to 42% instantly
# - Handles Ctrl+C / graceful shutdown cleanly
# - Works on Windows only

import ctypes, ctypes.wintypes, sys, threading, time, signal, os, glob

if sys.platform != "win32":
    print("Windows only.")
    sys.exit(1)

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
winmm = ctypes.windll.winmm  # for waveOutSetVolume fallback

# pointer-sized types
PTR_SIZE = ctypes.sizeof(ctypes.c_void_p)
if PTR_SIZE == 8:
    ULONG_PTR = ctypes.c_uint64
    LRESULT = ctypes.c_int64
else:
    ULONG_PTR = ctypes.c_uint32
    LRESULT = ctypes.c_long

# Structures
class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode", ctypes.wintypes.DWORD),
        ("scanCode", ctypes.wintypes.DWORD),
        ("flags", ctypes.wintypes.DWORD),
        ("time", ctypes.wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR)
    ]

# Windows input structures for SendInput (mouse + keyboard)
class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long),
        ("dy", ctypes.c_long),
        ("mouseData", ctypes.wintypes.DWORD),
        ("dwFlags", ctypes.wintypes.DWORD),
        ("time", ctypes.wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.wintypes.WORD),
        ("wScan", ctypes.wintypes.WORD),
        ("dwFlags", ctypes.wintypes.DWORD),
        ("time", ctypes.wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]

class _INPUTunion(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT)]

class INPUT(ctypes.Structure):
    _anonymous_ = ("u",)
    _fields_ = [("type", ctypes.wintypes.DWORD), ("u", _INPUTunion)]

# Constants
WH_KEYBOARD_LL = 13
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_SYSKEYDOWN = 0x0104
WM_SYSKEYUP = 0x0105

WHEEL_DELTA = 120

# Virtual keys used
VK_W, VK_A, VK_S, VK_D = 0x57, 0x41, 0x53, 0x44
VK_LSHIFT, VK_RSHIFT = 0xA0, 0xA1
VK_LCONTROL, VK_RCONTROL, VK_CONTROL = 0xA2, 0xA3, 0x11
VK_LMENU, VK_RMENU, VK_MENU = 0xA4, 0xA5, 0x12  # Alt keys
VK_ESCAPE = 0x1B
VK_TAB = 0x09
VK_0, VK_NUMPAD0 = 0x30, 0x60
VK_1, VK_2, VK_NUMPAD1, VK_NUMPAD2 = 0x31, 0x32, 0x61, 0x62
VK_3, VK_4, VK_NUMPAD3, VK_NUMPAD4 = 0x33, 0x34, 0x63, 0x64
VK_OEM_MINUS, VK_SUBTRACT = 0xBD, 0x6D
VK_OEM_PLUS, VK_ADD = 0xBB, 0x6B
VK_CAPITAL = 0x14  # Caps Lock
VK_OEM_3 = 0xC0    # ` ~ (US keyboard)
VK_RETURN = 0x0D   # Enter
VK_F = 0x46        # F key

# Arrow keys
VK_LEFT = 0x25
VK_UP = 0x26
VK_RIGHT = 0x27
VK_DOWN = 0x28

# Windows keys
VK_LWIN = 0x5B
VK_RWIN = 0x5C

# '[' and ']' virtual keys (US layout)
VK_OEM_4 = 0xDB  # '['
VK_OEM_6 = 0xDD  # ']'

# mouse flags
MOUSEEVENTF_MOVE       = 0x0001
MOUSEEVENTF_LEFTDOWN   = 0x0002
MOUSEEVENTF_LEFTUP     = 0x0004
MOUSEEVENTF_RIGHTDOWN  = 0x0008
MOUSEEVENTF_RIGHTUP    = 0x0010
MOUSEEVENTF_MIDDLEDOWN = 0x0020
MOUSEEVENTF_MIDDLEUP   = 0x0040
MOUSEEVENTF_WHEEL      = 0x0800
MOUSEEVENTF_HWHEEL     = 0x01000
MOUSEEVENTF_ABSOLUTE   = 0x8000

# keyboard/input constants
INPUT_MOUSE = 0
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002

# Console color (Windows 10+ Virtual Terminal)
VT_ENABLED = False
CSI = "\x1b["
C_RESET = CSI + "0m"
C_BOLD = CSI + "1m"
C_DIM = CSI + "2m"
C_RED = CSI + "31m"
C_GREEN = CSI + "32m"
C_YELLOW = CSI + "33m"
C_BLUE = CSI + "34m"
C_MAGENTA = CSI + "35m"
C_CYAN = CSI + "36m"
C_WHITE = CSI + "37m"

def _c(s, color):
    return f"{color}{s}{C_RESET}" if VT_ENABLED else s

def enable_vt_console_colors():
    global VT_ENABLED
    try:
        hOut = kernel32.GetStdHandle(-11)  # STD_OUTPUT_HANDLE
        if not hOut or hOut == ctypes.c_void_p(-1).value:
            VT_ENABLED = False
            return
        mode = ctypes.wintypes.DWORD(0)
        if not kernel32.GetConsoleMode(hOut, ctypes.byref(mode)):
            VT_ENABLED = False
            return
        ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
        new_mode = ctypes.wintypes.DWORD(mode.value | ENABLE_VIRTUAL_TERMINAL_PROCESSING)
        if kernel32.SetConsoleMode(hOut, new_mode):
            VT_ENABLED = True
        else:
            VT_ENABLED = False
    except Exception:
        VT_ENABLED = False

# Hook prototype
HOOKPROC = ctypes.WINFUNCTYPE(LRESULT, ctypes.c_int, ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM)

# prototypes
user32.SetWindowsHookExW.argtypes = [ctypes.c_int, HOOKPROC, ctypes.wintypes.HINSTANCE, ctypes.wintypes.DWORD]
user32.SetWindowsHookExW.restype = ctypes.c_void_p
user32.CallNextHookEx.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM]
user32.CallNextHookEx.restype = LRESULT
user32.UnhookWindowsHookEx.argtypes = [ctypes.c_void_p]
user32.UnhookWindowsHookEx.restype = ctypes.wintypes.BOOL

user32.GetMessageW.argtypes = [ctypes.POINTER(ctypes.wintypes.MSG), ctypes.wintypes.HWND, ctypes.wintypes.UINT, ctypes.wintypes.UINT]
user32.GetMessageW.restype = ctypes.c_int

# SendInput prototype
SendInput = user32.SendInput
SendInput.argtypes = (ctypes.wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int)
SendInput.restype = ctypes.wintypes.UINT

# Globals
hook_handle = None
_hook_proc_ref = None
_loaded_module_handle = None

# Which keys we intercept and block (blocking behavior)
# ESC removed as requested (won't forcibly exit). TAB is included but will be allowed through when Alt is held.
# Arrow keys included and blocked when active (they act like WASD).
SCRIPT_KEYS = {
    VK_W, VK_A, VK_S, VK_D,
    VK_LSHIFT, VK_RSHIFT,
    VK_1, VK_2, VK_NUMPAD1, VK_NUMPAD2,
    VK_3, VK_4, VK_NUMPAD3, VK_NUMPAD4,
    VK_OEM_MINUS, VK_SUBTRACT,
    VK_OEM_PLUS, VK_ADD,
    VK_CAPITAL,   # block CapsLock physical key (we handle it ourselves)
    VK_TAB,       # listen for TAB (for dragging) but do NOT block when Alt is held
    VK_OEM_3,     # backtick/tilde (reserved for toggle)
    VK_UP, VK_DOWN, VK_LEFT, VK_RIGHT,  # arrow keys behave like WASD
    VK_F,         # F => middle click
}

# pressed keys tracked by the hook:
pressed_keys = set()
pressed_lock = threading.Lock()

# active toggle (ON by default)
active = True

# For Shift+Enter passthrough: we synthesize a Shift while Enter is pressed so the app sees Shift+Enter.
synth_shift_active = False
synth_shift_vk = None  # VK_LSHIFT or VK_RSHIFT we synthesized

# dragging state (True when we've sent a left-button down and haven't yet released)
dragging_active = False

def _format_last_error():
    err = kernel32.GetLastError()
    if not err:
        return "No error code"
    buf = ctypes.create_unicode_buffer(2048)
    kernel32.FormatMessageW(0x00001000, None, err, 0, buf, len(buf), None)
    return f"WinErr {err}: {buf.value.strip()}"

# helper mouse/keyboard functions using SendInput (so Windows treats them like real input)
def get_cursor_pos():
    pt = ctypes.wintypes.POINT()
    user32.GetCursorPos(ctypes.byref(pt))
    return pt.x, pt.y

def set_cursor_pos(x, y):
    """
    Move cursor using SendInput absolute coordinates across the virtual screen.
    This generates WM_MOUSEMOVE events that Windows UI animations (taskbar/settings) will see.
    """
    # virtual screen bounds
    left = user32.GetSystemMetrics(76)   # SM_XVIRTUALSCREEN
    top  = user32.GetSystemMetrics(77)   # SM_YVIRTUALSCREEN
    width = user32.GetSystemMetrics(78)  # SM_CXVIRTUALSCREEN
    height = user32.GetSystemMetrics(79) # SM_CYVIRTUALSCREEN

    # fallback to primary screen if something weird
    if width <= 1: width = max(1, user32.GetSystemMetrics(0))
    if height <= 1: height = max(1, user32.GetSystemMetrics(1))

    # clamp coords to virtual screen
    tx = max(left, min(left + width - 1, int(x)))
    ty = max(top,  min(top + height - 1, int(y)))

    # convert to 0..65535 range expected by SendInput for absolute movement
    abs_x = int((tx - left) * 65535 // (width - 1))
    abs_y = int((ty - top)  * 65535 // (height - 1))

    inp = INPUT()
    inp.type = INPUT_MOUSE
    # populate the MOUSEINPUT sub-structure
    inp.mi.dx = abs_x
    inp.mi.dy = abs_y
    inp.mi.mouseData = 0
    inp.mi.dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE
    inp.mi.time = 0
    inp.mi.dwExtraInfo = 0

    SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))

def send_left_click():
    # left down
    down = INPUT()
    down.type = INPUT_MOUSE
    down.mi.dx = 0; down.mi.dy = 0; down.mi.mouseData = 0
    down.mi.dwFlags = MOUSEEVENTF_LEFTDOWN
    down.mi.time = 0; down.mi.dwExtraInfo = 0
    SendInput(1, ctypes.byref(down), ctypes.sizeof(down))
    # left up
    up = INPUT()
    up.type = INPUT_MOUSE
    up.mi.dx = 0; up.mi.dy = 0; up.mi.mouseData = 0
    up.mi.dwFlags = MOUSEEVENTF_LEFTUP
    up.mi.time = 0; up.mi.dwExtraInfo = 0
    SendInput(1, ctypes.byref(up), ctypes.sizeof(up))

def send_left_down():
    down = INPUT()
    down.type = INPUT_MOUSE
    down.mi.dx = 0; down.mi.dy = 0; down.mi.mouseData = 0
    down.mi.dwFlags = MOUSEEVENTF_LEFTDOWN
    down.mi.time = 0; down.mi.dwExtraInfo = 0
    SendInput(1, ctypes.byref(down), ctypes.sizeof(down))

def send_left_up():
    up = INPUT()
    up.type = INPUT_MOUSE
    up.mi.dx = 0; up.mi.dy = 0; up.mi.mouseData = 0
    up.mi.dwFlags = MOUSEEVENTF_LEFTUP
    up.mi.time = 0; up.mi.dwExtraInfo = 0
    SendInput(1, ctypes.byref(up), ctypes.sizeof(up))

def send_right_click():
    # right down
    down = INPUT()
    down.type = INPUT_MOUSE
    down.mi.dx = 0; down.mi.dy = 0; down.mi.mouseData = 0
    down.mi.dwFlags = MOUSEEVENTF_RIGHTDOWN
    down.mi.time = 0; down.mi.dwExtraInfo = 0
    SendInput(1, ctypes.byref(down), ctypes.sizeof(down))
    # right up
    up = INPUT()
    up.type = INPUT_MOUSE
    up.mi.dx = 0; up.mi.dy = 0; up.mi.mouseData = 0
    up.mi.dwFlags = MOUSEEVENTF_RIGHTUP
    up.mi.time = 0; up.mi.dwExtraInfo = 0
    SendInput(1, ctypes.byref(up), ctypes.sizeof(up))

def send_middle_click():
    # middle down
    down = INPUT()
    down.type = INPUT_MOUSE
    down.mi.dx = 0; down.mi.dy = 0; down.mi.mouseData = 0
    down.mi.dwFlags = MOUSEEVENTF_MIDDLEDOWN
    down.mi.time = 0; down.mi.dwExtraInfo = 0
    SendInput(1, ctypes.byref(down), ctypes.sizeof(down))
    # middle up
    up = INPUT()
    up.type = INPUT_MOUSE
    up.mi.dx = 0; up.mi.dy = 0; up.mi.mouseData = 0
    up.mi.dwFlags = MOUSEEVENTF_MIDDLEUP
    up.mi.time = 0; up.mi.dwExtraInfo = 0
    SendInput(1, ctypes.byref(up), ctypes.sizeof(up))

def send_scroll(delta):
    sc = INPUT()
    sc.type = INPUT_MOUSE
    sc.mi.dx = 0; sc.mi.dy = 0
    sc.mi.mouseData = int(delta)
    sc.mi.dwFlags = MOUSEEVENTF_WHEEL
    sc.mi.time = 0; sc.mi.dwExtraInfo = 0
    SendInput(1, ctypes.byref(sc), ctypes.sizeof(sc))

def send_hscroll(delta):
    """
    Horizontal scroll. Positive delta scrolls RIGHT, negative scrolls LEFT.
    """
    sc = INPUT()
    sc.type = INPUT_MOUSE
    sc.mi.dx = 0; sc.mi.dy = 0
    sc.mi.mouseData = int(delta)
    sc.mi.dwFlags = MOUSEEVENTF_HWHEEL
    sc.mi.time = 0; sc.mi.dwExtraInfo = 0
    SendInput(1, ctypes.byref(sc), ctypes.sizeof(sc))

def send_key_vk(vk):
    """Send a synthetic keyboard press+release for virtual-key 'vk' via SendInput."""
    send_vk_down(vk)
    time.sleep(0.001)
    send_vk_up(vk)

def send_vk_down(vk):
    down = INPUT()
    down.type = INPUT_KEYBOARD
    down.ki.wVk = vk
    down.ki.wScan = 0
    down.ki.dwFlags = 0
    down.ki.time = 0
    down.ki.dwExtraInfo = 0
    SendInput(1, ctypes.byref(down), ctypes.sizeof(down))

def send_vk_up(vk):
    up = INPUT()
    up.type = INPUT_KEYBOARD
    up.ki.wVk = vk
    up.ki.wScan = 0
    up.ki.dwFlags = KEYEVENTF_KEYUP
    up.ki.time = 0
    up.ki.dwExtraInfo = 0
    SendInput(1, ctypes.byref(up), ctypes.sizeof(up))

def send_caps_double_toggle():
    """
    Send two synthetic CapsLock toggles (press+release twice).
    This results in net-zero change of CapsLock state but produces real keyboard events.
    """
    send_key_vk(VK_CAPITAL)
    time.sleep(0.01)
    send_key_vk(VK_CAPITAL)
    time.sleep(0.01)

def _always_block_key(vk):
    """
    Keys that are always blocked from reaching apps, regardless of active/Ctrl:
    - VK_OEM_3 (backtick) is RESERVED for toggling.
    """
    return vk == VK_OEM_3

def _maybe_release_synth_shift():
    global synth_shift_active, synth_shift_vk
    if synth_shift_active and synth_shift_vk is not None:
        try:
            send_vk_up(synth_shift_vk)
        except Exception:
            pass
        synth_shift_active = False
        synth_shift_vk = None

# Volume helper using waveOutSetVolume fallback (left & right channels)
def set_master_volume_percent(percent):
    """
    Set system volume to a percentage (0..100) using waveOutSetVolume as a compatible fallback.
    If percent is outside 0..100, it will be clamped.
    """
    try:
        p = float(percent)
        if p < 0.0: p = 0.0
        if p > 100.0: p = 100.0
        # waveOutSetVolume expects 0x0000..0xFFFF for each channel.
        vol = int(round((p / 100.0) * 0xFFFF)) & 0xFFFF
        dwVolume = (vol << 16) | vol  # high word = right, low = left
        # waveOutSetVolume takes (HWAVEOUT hwo, DWORD dwVolume)
        # Using hwo = 0 sets the first waveform-audio output device. This commonly maps to master.
        winmm.waveOutSetVolume(ctypes.c_uint(0), ctypes.c_uint(dwVolume))
    except Exception as e:
        # swallow errors but print for debugging
        print(_c("Volume set failed:", C_RED), e)

# Low-level keyboard hook: track keydown/up in pressed_keys and block as required.
def _low_level_keyboard_proc(nCode, wParam, lParam):
    global active, synth_shift_active, synth_shift_vk
    try:
        if nCode >= 0:
            k = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
            vk = int(k.vkCode)
            flags = int(k.flags)
            # Detect injected events: LLKHF_INJECTED == 0x10
            injected = (flags & 0x10) != 0

            # Handle keydown / keyup tracking BUT ignore synthetic/injected events
            if not injected:
                if wParam == WM_KEYDOWN or wParam == WM_SYSKEYDOWN:
                    with pressed_lock:
                        pressed_keys.add(vk)
                elif wParam == WM_KEYUP or wParam == WM_SYSKEYUP:
                    with pressed_lock:
                        pressed_keys.discard(vk)

            # Determine modifiers
            with pressed_lock:
                ctrl_held = (VK_LCONTROL in pressed_keys) or (VK_RCONTROL in pressed_keys) or (VK_CONTROL in pressed_keys)
                alt_held = (VK_LMENU in pressed_keys) or (VK_RMENU in pressed_keys) or (VK_MENU in pressed_keys)
                shift_held = (VK_LSHIFT in pressed_keys) or (VK_RSHIFT in pressed_keys)
                win_held = (VK_LWIN in pressed_keys) or (VK_RWIN in pressed_keys)

            # Special handling: allow Shift+Enter to reach the app by synthesizing Shift while Enter is pressed
            # Only needed when we would otherwise block Shift (i.e., active and not ctrl_held).
            if not injected:
                if (wParam == WM_KEYDOWN or wParam == WM_SYSKEYDOWN) and vk == VK_RETURN:
                    if active and not ctrl_held and shift_held and not synth_shift_active:
                        # Prefer left shift if held; else right; default to left as fallback
                        used_vk = VK_LSHIFT if (VK_LSHIFT in pressed_keys) else (VK_RSHIFT if (VK_RSHIFT in pressed_keys) else VK_LSHIFT)
                        try:
                            send_vk_down(used_vk)
                            synth_shift_active = True
                            synth_shift_vk = used_vk
                        except Exception:
                            pass
                elif (wParam == WM_KEYUP or wParam == WM_SYSKEYUP) and (vk == VK_LSHIFT or vk == VK_RSHIFT):
                    # On physical shift release, release our synthetic shift if we had one
                    if synth_shift_active and synth_shift_vk == vk:
                        try:
                            send_vk_up(synth_shift_vk)
                        except Exception:
                            pass
                        synth_shift_active = False
                        synth_shift_vk = None

            # Immediate '[' and ']' handling: set system volume on physical (non-injected) keydown.
            # We do NOT block these keys — they are allowed through.
            if not injected and (wParam == WM_KEYDOWN or wParam == WM_SYSKEYDOWN):
                if vk == VK_OEM_4:  # '['
                    set_master_volume_percent(24.0)
                    # print feedback
                    print(_c("Volume -> 24%", C_BLUE))
                elif vk == VK_OEM_6:  # ']'
                    set_master_volume_percent(42.0)
                    print(_c("Volume -> 42%", C_BLUE))

            # Decide whether to block the key from other apps
            # Only block physical (non-injected) keys we care about when active and neither Ctrl nor Win is held.
            if not injected:
                # Always-blocked keys (backtick)
                if _always_block_key(vk):
                    return int(1)

                # If this is TAB and Alt is held, let it through (so Alt+Tab works normally)
                if vk == VK_TAB and alt_held:
                    pass
                else:
                    # block script keys only when active and Ctrl is NOT held and Win is NOT held
                    if (vk in SCRIPT_KEYS) and active and not ctrl_held and not win_held:
                        return int(1)
    except Exception:
        # Never allow exceptions to escape the callback; forward to next hook
        pass

    # Otherwise pass to next hook
    try:
        return int(user32.CallNextHookEx(hook_handle, nCode, wParam, lParam))
    except Exception:
        return int(0)

def install_keyboard_hook():
    global hook_handle, _hook_proc_ref, _loaded_module_handle
    # keep callback alive
    hook_proc = HOOKPROC(_low_level_keyboard_proc)
    _hook_proc_ref = hook_proc

    # Try a few module handle strategies (similar to previous robust script)
    tried = []
    # 1: GetModuleHandleW(None)
    try:
        h = kernel32.GetModuleHandleW(None)
        if h:
            hinst = ctypes.wintypes.HINSTANCE(h)
            hook_handle = user32.SetWindowsHookExW(WH_KEYBOARD_LL, hook_proc, hinst, 0)
            tried.append(("GetModuleHandleW(None)", h, bool(hook_handle)))
            if hook_handle:
                return hook_handle
    except Exception:
        pass

    # 2: ctypes.pythonapi._handle
    try:
        h = getattr(ctypes.pythonapi, "_handle", None)
        if h:
            hinst = ctypes.wintypes.HINSTANCE(int(h))
            hook_handle = user32.SetWindowsHookExW(WH_KEYBOARD_LL, hook_proc, hinst, 0)
            tried.append(("ctypes.pythonapi._handle", h, bool(hook_handle)))
            if hook_handle:
                return hook_handle
    except Exception:
        pass

    # 3: GetModuleHandleW(exe basename)
    try:
        exe_basename = os.path.basename(sys.executable)
        if exe_basename:
            h = kernel32.GetModuleHandleW(exe_basename)
            if h:
                hinst = ctypes.wintypes.HINSTANCE(h)
                hook_handle = user32.SetWindowsHookExW(WH_KEYBOARD_LL, hook_proc, hinst, 0)
                tried.append((f"GetModuleHandleW({exe_basename})", h, bool(hook_handle)))
                if hook_handle:
                    return hook_handle
    except Exception:
        pass

    # 4: try LoadLibraryW on python DLLs next to exe
    try:
        exe_dir = os.path.dirname(sys.executable) or os.getcwd()
        for path in glob.glob(os.path.join(exe_dir, "python*.dll")):
            try:
                lib = kernel32.LoadLibraryW(path)
                if lib:
                    hinst = ctypes.wintypes.HINSTANCE(lib)
                    hook_handle = user32.SetWindowsHookExW(WH_KEYBOARD_LL, hook_proc, hinst, 0)
                    tried.append((f"LoadLibraryW({path})", lib, bool(hook_handle)))
                    if hook_handle:
                        _loaded_module_handle = lib
                        return hook_handle
                    else:
                        # free and continue
                        kernel32.FreeLibrary(lib)
            except Exception:
                pass
    except Exception:
        pass

    # report failure
    msg = "Failed to install keyboard hook. Attempts:\n"
    for entry in tried:
        msg += f"  {entry[0]} => handle {entry[1]} installed={entry[2]}\n"
    msg += "Last WinErr: " + _format_last_error()
    raise OSError(msg)

def uninstall_keyboard_hook():
    global hook_handle, _hook_proc_ref, _loaded_module_handle
    if hook_handle:
        try:
            user32.UnhookWindowsHookEx(hook_handle)
        except Exception:
            pass
        hook_handle = None
    if _loaded_module_handle:
        try:
            kernel32.FreeLibrary(ctypes.c_void_p(_loaded_module_handle))
        except Exception:
            pass
        _loaded_module_handle = None
    _hook_proc_ref = None

# message pump (must run in the thread that installed the hook)
def message_pump():
    msg = ctypes.wintypes.MSG()
    while True:
        bRet = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
        if bRet == 0 or bRet == -1:
            break
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))

# movement + behavior worker reads pressed_keys (NOT GetAsyncKeyState)
def movement_worker(stop_event):
    global active, dragging_active
    # STEP=4 with TICK=0.01 -> ~400 px/s default
    STEP = 4
    TICK = 0.01
    MIN_STEP, MAX_STEP = 1, 50

    SCROLL_AMOUNT = WHEEL_DELTA
    # 50% slower than before (was 0.05s => 20/s); now 0.10s => 10/s
    SCROLL_INTERVAL = 0.10

    try:
        user32.SetProcessDPIAware()
    except Exception:
        pass

    left = user32.GetSystemMetrics(76)
    top  = user32.GetSystemMetrics(77)
    width = user32.GetSystemMetrics(78)
    height = user32.GetSystemMetrics(79)
    right = left + width - 1
    bottom = top + height - 1

    def px_per_sec(step): return int(step / TICK)

    # Intro/help (concise, colored)
    print(_c("Keyboard hook installed. Controls:", C_GREEN))
    print("  - Move: WASD or Arrow keys (hold).")
    print("  - Speed: - / + to decrease/increase.")
    print("  - Clicks: Shift = Left-click; CapsLock = Right-click; F = Middle-click.")
    print("  - Scroll: 1=Up, 2=Down, 3=Left, 4=Right (hold for continuous).")
    print("  - Drag: Hold Tab + (Shift or any WASD/Arrow) to start dragging; keep holding Tab to continue dragging; release Tab to drop.")
    print("  - Toggle: ` (backtick) toggles script ON/OFF.")
    print("  - Passthrough: Hold Ctrl or either Windows key to temporarily disable script and allow normal typing / Win shortcuts.")
    print("  - Volume keys: [ -> 24% | ] -> 42%")
    print("  - Quit: Ctrl+C")
    print(_c(f"Status: ON  | Move: ~{px_per_sec(STEP)} px/s | VScroll: {int(1/SCROLL_INTERVAL)} notches/s | HScroll: {int(1/SCROLL_INTERVAL)} notches/s", C_CYAN))

    prev_shift = False
    prev_scroll_pressed_v = False
    prev_scroll_pressed_h = False
    next_scroll_time_v = 0.0
    next_scroll_time_h = 0.0
    prev_plus = prev_minus = False
    prev_caps = False
    prev_f = False
    prev_backtick = False

    try:
        while not stop_event.is_set():
            now = time.time()

            with pressed_lock:
                snapshot = set(pressed_keys)

            # Also treat Arrow keys as WASD for movement/drag logic
            effective = set(snapshot)
            if VK_UP in snapshot: effective.add(VK_W)
            if VK_LEFT in snapshot: effective.add(VK_A)
            if VK_DOWN in snapshot: effective.add(VK_S)
            if VK_RIGHT in snapshot: effective.add(VK_D)

            # detect modifier keys
            ctrl_held = (VK_LCONTROL in snapshot) or (VK_RCONTROL in snapshot) or (VK_CONTROL in snapshot)
            alt_held = (VK_LMENU in snapshot) or (VK_RMENU in snapshot) or (VK_MENU in snapshot)
            win_held = (VK_LWIN in snapshot) or (VK_RWIN in snapshot)

            # Toggle with ` (backtick)
            backtick_now = VK_OEM_3 in snapshot
            if backtick_now and not prev_backtick:
                active = not active
                print(_c(f"* Script {'ON' if active else 'OFF'}", C_GREEN if active else C_RED))
                # If we turned OFF, ensure no drag is left held
                if not active and dragging_active:
                    try:
                        send_left_up()
                    except Exception:
                        pass
                    dragging_active = False
            prev_backtick = backtick_now

            # Speed adjust
            plus_now = (VK_OEM_PLUS in snapshot) or (VK_ADD in snapshot)
            minus_now = (VK_OEM_MINUS in snapshot) or (VK_SUBTRACT in snapshot)
            if active:
                if plus_now and not prev_plus:
                    STEP = min(MAX_STEP, STEP + 1)
                    print(_c(f"* Speed: ~{px_per_sec(STEP)} px/s (STEP={STEP})", C_YELLOW))
                if minus_now and not prev_minus:
                    STEP = max(MIN_STEP, STEP - 1)
                    print(_c(f"* Speed: ~{px_per_sec(STEP)} px/s (STEP={STEP})", C_YELLOW))
            prev_plus = plus_now; prev_minus = minus_now

            # If Ctrl or Windows key held, disable script actions so shortcuts work normally
            if ctrl_held or win_held:
                # Reset edge trackers so next press triggers correctly when Ctrl/Win released
                prev_shift = False
                prev_scroll_pressed_v = False
                prev_scroll_pressed_h = False
                next_scroll_time_v = 0.0
                next_scroll_time_h = 0.0
                prev_caps = False
                prev_f = False
                # If we were dragging, ensure release so we don't leave mouse stuck
                if dragging_active:
                    try:
                        send_left_up()
                    except Exception:
                        pass
                    dragging_active = False
                time.sleep(TICK)
                continue

            # If Alt is held, make sure we aren't starting/continuing a drag caused by Tab
            if alt_held:
                if dragging_active:
                    try:
                        send_left_up()
                    except Exception:
                        pass
                    dragging_active = False

            if active:
                # Movement
                shift_now = (VK_LSHIFT in snapshot) or (VK_RSHIFT in snapshot)
                # Determine speed multiplier: if shift held while moving, reduce speed to 25% (75% slower)
                speed_multiplier = 0.25 if shift_now else 1.0
                step_effective = max(1, int(round(STEP * speed_multiplier)))

                dx = step_effective * (int(VK_D in effective) - int(VK_A in effective))
                dy = step_effective * (int(VK_S in effective) - int(VK_W in effective))
                if dx != 0 or dy != 0:
                    x, y = get_cursor_pos()
                    nx = max(left, min(right, x + dx))
                    ny = max(top,  min(bottom, y + dy))
                    set_cursor_pos(nx, ny)

                # Vertical Scrolling (1 up, 2 down)
                up_pressed = (VK_1 in snapshot) or (VK_NUMPAD1 in snapshot)
                down_pressed = (VK_2 in snapshot) or (VK_NUMPAD2 in snapshot)
                scroll_dir_v = 1 if up_pressed and not down_pressed else (-1 if down_pressed and not up_pressed else 0)
                scroll_pressed_v = scroll_dir_v != 0
                if scroll_pressed_v:
                    if not prev_scroll_pressed_v or now >= next_scroll_time_v:
                        send_scroll(scroll_dir_v * SCROLL_AMOUNT)
                        next_scroll_time_v = now + SCROLL_INTERVAL
                else:
                    next_scroll_time_v = 0.0

                # Horizontal Scrolling (3 left, 4 right)
                left_pressed = (VK_3 in snapshot) or (VK_NUMPAD3 in snapshot)
                right_pressed = (VK_4 in snapshot) or (VK_NUMPAD4 in snapshot)
                # For HWHEEL: positive is RIGHT, negative is LEFT
                scroll_dir_h = (1 if right_pressed and not left_pressed else (-1 if left_pressed and not right_pressed else 0))
                scroll_pressed_h = scroll_dir_h != 0
                if scroll_pressed_h:
                    if not prev_scroll_pressed_h or now >= next_scroll_time_h:
                        send_hscroll(scroll_dir_h * SCROLL_AMOUNT)
                        next_scroll_time_h = now + SCROLL_INTERVAL
                else:
                    next_scroll_time_h = 0.0

                # Dragging logic CHANGE:
                # - Start drag when Tab is down AND (Shift OR any WASD/Arrow) becomes true (Tab-first or keys-first both work).
                # - Once dragging started, KEEP dragging while Tab remains held (regardless of movement keys).
                # - Releasing Tab (or Alt/Ctrl/script off) stops dragging.
                tab_now = VK_TAB in snapshot
                wasd_or_arrows_held = any(k in effective for k in (VK_W, VK_A, VK_S, VK_D))
                start_drag_condition = (tab_now and (shift_now or wasd_or_arrows_held)) and (not alt_held)

                # Shift click (fire on press edge) -- but skip if this press is starting a drag or we are currently dragging
                if shift_now and not prev_shift:
                    if (not wasd_or_arrows_held) and (not start_drag_condition) and (not dragging_active):
                        send_left_click()
                prev_shift = shift_now
                prev_scroll_pressed_v = scroll_pressed_v
                prev_scroll_pressed_h = scroll_pressed_h

                # Start drag if not already dragging and start condition hit
                if (not dragging_active) and start_drag_condition:
                    try:
                        send_left_down()
                        dragging_active = True
                    except Exception:
                        pass
                # Stop drag if we are dragging but Tab released or Alt is held (or script turned off)
                elif dragging_active and (not tab_now or alt_held):
                    try:
                        send_left_up()
                        dragging_active = False
                    except Exception:
                        pass

                # CapsLock -> perform double-toggle (net no state change) and then RIGHT CLICK (on press edge)
                caps_now = VK_CAPITAL in snapshot
                if caps_now and not prev_caps:
                    try:
                        send_caps_double_toggle()
                        send_right_click()
                    except Exception:
                        pass
                prev_caps = caps_now

                # F -> Middle click (on press edge)
                f_now = VK_F in snapshot
                if f_now and not prev_f:
                    try:
                        send_middle_click()
                    except Exception:
                        pass
                prev_f = f_now
            else:
                # When OFF, reset edge trackers so next press triggers immediately
                prev_shift = False
                prev_scroll_pressed_v = False
                prev_scroll_pressed_h = False
                next_scroll_time_v = 0.0
                next_scroll_time_h = 0.0
                prev_caps = False
                prev_f = False
                if dragging_active:
                    try:
                        send_left_up()
                    except Exception:
                        pass
                    dragging_active = False

            time.sleep(TICK)
    finally:
        # Ensure that if the worker exits we don't leave the left button held down
        if dragging_active:
            try:
                send_left_up()
            except Exception:
                pass
            dragging_active = False
        # Ensure we don't leave a synthetic shift pressed
        _maybe_release_synth_shift()

# Graceful shutdown on Ctrl+C
stop_evt = threading.Event()
def _sigint_handler(signum, frame):
    # Ask message pump to quit and stop worker
    stop_evt.set()
    try:
        user32.PostQuitMessage(0)
    except Exception:
        pass

signal.signal(signal.SIGINT, _sigint_handler)

def main():
    enable_vt_console_colors()

    # Install hook in main thread (message_pump will run here)
    try:
        install_keyboard_hook()
        print(_c("Keyboard hook installed successfully.", C_GREEN))
    except Exception as e:
        print(_c("Failed to install keyboard hook:", C_RED), e)
        print("Try: run from an Administrator command prompt or use system python (not MS Store).")
        return

    # Start movement worker
    worker = threading.Thread(target=movement_worker, args=(stop_evt,), daemon=True)
    worker.start()

    # Run message pump in main thread (required)
    try:
        message_pump()
    finally:
        # ensure worker stops and hook removed
        stop_evt.set()
        try:
            worker.join(timeout=1.0)
        except Exception:
            pass
        # Make sure dragging released before uninstalling hook
        try:
            if dragging_active:
                send_left_up()
        except Exception:
            pass
        # Release any synthetic shift we may have pressed for Shift+Enter passthrough
        _maybe_release_synth_shift()
        uninstall_keyboard_hook()
        print(_c("Keyboard hook removed. Exiting.", C_MAGENTA))

if __name__ == "__main__":
    main()
