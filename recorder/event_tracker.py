import time
from dataclasses import dataclass, field
from typing import Optional

from pynput import mouse as pmouse, keyboard as pkeyboard


@dataclass
class ActionEvent:
    timestamp: float
    action_type: str        # click | scroll | key | screenshot | note | start | stop
    description: str
    screenshot_b64: Optional[str] = None
    note: str = ""
    x: int = 0
    y: int = 0


class EventTracker:
    def __init__(self, on_event):
        self._on_event = on_event
        self._mouse_listener    = None
        self._keyboard_listener = None
        self._active = False
        self._last_click_time = 0.0
        self._min_click_interval = 0.5

    def start(self):
        self._active = True
        self._mouse_listener = pmouse.Listener(
            on_click=self._on_click,
            on_scroll=self._on_scroll,
        )
        self._keyboard_listener = pkeyboard.Listener(
            on_press=self._on_key_press,
        )
        self._mouse_listener.start()
        self._keyboard_listener.start()

    def stop(self):
        self._active = False
        if self._mouse_listener:
            self._mouse_listener.stop()
        if self._keyboard_listener:
            self._keyboard_listener.stop()

    def _on_click(self, x, y, button, pressed):
        if not self._active or not pressed:
            return
        now = time.time()
        if now - self._last_click_time < self._min_click_interval:
            return
        self._last_click_time = now
        btn_name = (
            "Links" if button == pmouse.Button.left else
            "Rechts" if button == pmouse.Button.right else "Mitte"
        )
        self._on_event("click", f"{btn_name}klick bei ({x}, {y})", x, y)

    def _on_scroll(self, x, y, dx, dy):
        if not self._active:
            return
        direction = "unten" if dy < 0 else "oben"
        self._on_event("scroll", f"Gescrollt nach {direction} bei ({x}, {y})", x, y)

    def _on_key_press(self, key):
        if not self._active:
            return
        try:
            key_str = key.char if hasattr(key, "char") and key.char else str(key).replace("Key.", "")
        except AttributeError:
            key_str = str(key).replace("Key.", "")
        special = {
            "enter", "tab", "backspace", "delete", "escape", "space",
            "f1", "f2", "f3", "f4", "f5", "f6", "f7",
            "f11", "f12", "ctrl_l", "ctrl_r", "alt_l", "alt_r",
            "shift", "cmd", "page_up", "page_down", "home", "end",
        }
        if key_str.lower() in special:
            self._on_event("key", f"Taste gedrückt: {key_str}", 0, 0)
