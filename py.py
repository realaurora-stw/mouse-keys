#!/usr/bin/env python3
# WASD continuous mouse move + Shift left-click + 1/2 scrolling.
# - Uses a real low-level keyboard hook (blocks keys from other apps) except when Ctrl is held
# - Tracks keydown/keyup in the hook and uses that state for movement
# - Shift + WASD -> 75% slower (i.e. 25% speed) while Shift is held
# - Ctrl disables script actions and allows normal key behavior (Ctrl+A/C/V etc)
# - Handles Ctrl+C / graceful shutdown cleanly
# - Works on Windows only
# UPDATE: Use SendInput for movement/click/scroll so OS hover animations (taskbar, Settings, etc) trigger.
# UPDATE: Caps Lock (physical key) has been repurposed: when pressed it will (1) perform two synthetic CapsLock toggles
# (so the net CapsLock state remains unchanged) and (2) perform a RIGHT CLICK using SendInput.
# The physical CapsLock key is blocked from other apps while script is active (unless Ctrl is held).

import ctypes, ctypes.wintypes, sys, threading, time, signal, os, glob

if sys.platform != "win32":
    print("Windows only.")
    sys.exit(1)

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

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
VK_ESCAPE = 0x1B
VK_0, VK_NUMPAD0 = 0x30, 0x60
VK_1, VK_2, VK_NUMPAD1, VK_NUMPAD2 = 0x31, 0x32, 0x61, 0x62
VK_OEM_MINUS, VK_SUBTRACT = 0xBD, 0x6D
VK_OEM_PLUS, VK_ADD = 0xBB, 0x6B
VK_CAPITAL = 0x14  # Caps Lock

# mouse flags
MOUSEEVENTF_MOVE      = 0x0001
MOUSEEVENTF_LEFTDOWN  = 0x0002
MOUSEEVENTF_LEFTUP    = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP   = 0x0010
MOUSEEVENTF_WHEEL     = 0x0800
MOUSEEVENTF_ABSOLUTE  = 0x8000

# keyboard/input constants
INPUT_MOUSE = 0
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002

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
SCRIPT_KEYS = {
    VK_W, VK_A, VK_S, VK_D,
    VK_LSHIFT, VK_RSHIFT,
    VK_1, VK_2, VK_NUMPAD1, VK_NUMPAD2,
    VK_OEM_MINUS, VK_SUBTRACT,
    VK_OEM_PLUS, VK_ADD,
    VK_0, VK_NUMPAD0,
    VK_ESCAPE,
    VK_CAPITAL,   # block CapsLock physical key (we handle it ourselves)
}
ALWAYS_BLOCK_KEYS = {VK_0, VK_NUMPAD0, VK_ESCAPE}

# pressed keys tracked by the hook:
pressed_keys = set()
pressed_lock = threading.Lock()

# active toggle (ON by default)
active = True

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

def send_scroll(delta):
    sc = INPUT()
    sc.type = INPUT_MOUSE
    sc.mi.dx = 0; sc.mi.dy = 0
    sc.mi.mouseData = int(delta)
    sc.mi.dwFlags = MOUSEEVENTF_WHEEL
    sc.mi.time = 0; sc.mi.dwExtraInfo = 0
    SendInput(1, ctypes.byref(sc), ctypes.sizeof(sc))

def send_key_vk(vk):
    """Send a synthetic keyboard press+release for virtual-key 'vk' via SendInput."""
    down = INPUT()
    down.type = INPUT_KEYBOARD
    down.ki.wVk = vk
    down.ki.wScan = 0
    down.ki.dwFlags = 0
    down.ki.time = 0
    down.ki.dwExtraInfo = 0
    SendInput(1, ctypes.byref(down), ctypes.sizeof(down))

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
    # small pause to ensure the OS processes them distinctly
    send_key_vk(VK_CAPITAL)
    time.sleep(0.01)
    send_key_vk(VK_CAPITAL)
    # a tiny delay to settle (not strictly required, but safer)
    time.sleep(0.01)

# Low-level keyboard hook: track keydown/up in pressed_keys and block as required.
def _low_level_keyboard_proc(nCode, wParam, lParam):
    global active
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

            # Determine if Ctrl is held (we allow normal key behavior when Ctrl is down)
            with pressed_lock:
                ctrl_held = (VK_LCONTROL in pressed_keys) or (VK_RCONTROL in pressed_keys) or (VK_CONTROL in pressed_keys)

            # Decide whether to block the key from other apps
            # Only block physical (non-injected) keys we care about when active and Ctrl is NOT held
            if (not injected) and (vk in SCRIPT_KEYS):
                # always block toggle & exit keys
                if vk in ALWAYS_BLOCK_KEYS:
                    return int(1)
                # block script keys only when active and Ctrl is NOT held
                if active and not ctrl_held:
                    return int(1)
                # If ctrl_held, fall through and let other apps receive the key (e.g., Ctrl+A)
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
    global active
    STEP = 3
    TICK = 0.01
    MIN_STEP, MAX_STEP = 1, 50

    SCROLL_AMOUNT = WHEEL_DELTA
    SCROLL_INTERVAL = 0.05

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

    print("WASD->move | Shift->click & Shift+WASD->75% slower | 1/2->scroll | 0->toggle ON/OFF | -/+->speed (when ON) | Esc->quit")
    print("CapsLock -> RIGHT CLICK (CapsLock state is toggled twice, net no change).")
    print("Script keys are blocked from other apps when active EXCEPT when Ctrl is held.")
    print(f"Status: ON | Move speed: ~{px_per_sec(STEP)} px/s | Scroll: {int(1/SCROLL_INTERVAL)} notches/s (amount={SCROLL_AMOUNT})")

    prev_shift = False
    prev_esc = False
    prev_scroll_pressed = False
    next_scroll_time = 0.0
    prev_zero = prev_plus = prev_minus = False
    prev_caps = False

    while not stop_event.is_set():
        now = time.time()

        with pressed_lock:
            snapshot = set(pressed_keys)

        # Escape: exit
        esc_now = VK_ESCAPE in snapshot
        if esc_now and not prev_esc:
            print("Exiting...")
            stop_event.set()
            try:
                user32.PostQuitMessage(0)
            except Exception:
                pass
            break
        prev_esc = esc_now

        # Toggle 0
        zero_now = (VK_0 in snapshot) or (VK_NUMPAD0 in snapshot)
        if zero_now and not prev_zero:
            active = not active
            print("Toggled:", "ON" if active else "OFF")
        prev_zero = zero_now

        # Speed adjust
        plus_now = (VK_OEM_PLUS in snapshot) or (VK_ADD in snapshot)
        minus_now = (VK_OEM_MINUS in snapshot) or (VK_SUBTRACT in snapshot)
        if active:
            if plus_now and not prev_plus:
                STEP = min(MAX_STEP, STEP + 1)
                print(f"Speed: ~{px_per_sec(STEP)} px/s (STEP={STEP})")
            if minus_now and not prev_minus:
                STEP = max(MIN_STEP, STEP - 1)
                print(f"Speed: ~{px_per_sec(STEP)} px/s (STEP={STEP})")
        prev_plus = plus_now; prev_minus = minus_now

        # If Ctrl held, disable script actions so Ctrl+keys behave normally
        ctrl_held = (VK_LCONTROL in snapshot) or (VK_RCONTROL in snapshot) or (VK_CONTROL in snapshot)
        if ctrl_held:
            # Reset edge trackers so next press triggers correctly when Ctrl released
            prev_shift = False
            prev_scroll_pressed = False
            next_scroll_time = 0.0
            prev_caps = False
            # Skip script behaviors entirely while Ctrl held (prevents interfering with combos)
            time.sleep(TICK)
            continue

        if active:
            # Movement
            shift_now = (VK_LSHIFT in snapshot) or (VK_RSHIFT in snapshot)
            # Determine speed multiplier: if shift held while moving, reduce speed to 25% (75% slower)
            speed_multiplier = 0.25 if shift_now else 1.0
            step_effective = max(1, int(round(STEP * speed_multiplier)))

            dx = step_effective * (int(VK_D in snapshot) - int(VK_A in snapshot))
            dy = step_effective * (int(VK_S in snapshot) - int(VK_W in snapshot))
            if dx != 0 or dy != 0:
                x, y = get_cursor_pos()
                nx = max(left, min(right, x + dx))
                ny = max(top,  min(bottom, y + dy))
                # use SendInput-based mover to ensure hover/WM_MOUSEMOVE events are delivered properly
                set_cursor_pos(nx, ny)

            # Scrolling
            up_pressed = (VK_1 in snapshot) or (VK_NUMPAD1 in snapshot)
            down_pressed = (VK_2 in snapshot) or (VK_NUMPAD2 in snapshot)
            scroll_dir = 1 if up_pressed and not down_pressed else (-1 if down_pressed and not up_pressed else 0)
            scroll_pressed = scroll_dir != 0
            if scroll_pressed:
                if not prev_scroll_pressed or now >= next_scroll_time:
                    send_scroll(scroll_dir * SCROLL_AMOUNT)
                    next_scroll_time = now + SCROLL_INTERVAL
            else:
                next_scroll_time = 0.0

            # Shift click (fire on press edge)
            # NEW: If shift is pressed while any WASD key is already held, DO NOT click (user requested).
            if shift_now and not prev_shift:
                wasd_held = any(k in snapshot for k in (VK_W, VK_A, VK_S, VK_D))
                if not wasd_held:
                    send_left_click()
            prev_shift = shift_now
            prev_scroll_pressed = scroll_pressed

            # CapsLock -> perform double-toggle (net no state change) and then RIGHT CLICK (on press edge)
            caps_now = VK_CAPITAL in snapshot
            if caps_now and not prev_caps:
                # perform double synthetic toggles so actual CapsLock state remains unchanged
                try:
                    send_caps_double_toggle()
                    # then perform right click
                    send_right_click()
                except Exception:
                    pass
            prev_caps = caps_now
        else:
            # When OFF, reset edge trackers so next press triggers immediately
            prev_shift = False
            prev_scroll_pressed = False
            next_scroll_time = 0.0
            prev_caps = False

        time.sleep(TICK)

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
    # Install hook in main thread (message_pump will run here)
    try:
        install_keyboard_hook()
        print("Keyboard hook installed successfully.")
    except Exception as e:
        print("Failed to install keyboard hook:", e)
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
        uninstall_keyboard_hook()
        print("Keyboard hook removed. Exiting.")

if __name__ == "__main__":
    main()

