import threading
from io import BytesIO
from typing import Optional

import mss
import mss.tools
from PIL import Image


class ScreenCapture:
    def __init__(self):
        self._lock = threading.Lock()

    def capture(self, quality: int = 85) -> Optional[str]:
        """Vollbild-Screenshot als Base64-JPEG."""
        import base64
        try:
            with self._lock:
                with mss.mss() as sct:
                    monitor = sct.monitors[0]
                    sct_img = sct.grab(monitor)
                img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
                buf = BytesIO()
                img.save(buf, format="JPEG", quality=quality, optimize=True)
                return base64.b64encode(buf.getvalue()).decode("utf-8")
        except Exception as e:
            print(f"Screenshot-Fehler: {e}")
            return None

    def capture_thumbnail(self, quality: int = 75, max_width: int = 1280) -> Optional[str]:
        """Screenshot verkleinert auf max_width als Base64-JPEG."""
        import base64
        try:
            with self._lock:
                with mss.mss() as sct:
                    monitor = sct.monitors[0]
                    sct_img = sct.grab(monitor)
                img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
                if img.width > max_width:
                    ratio = max_width / img.width
                    img = img.resize((max_width, int(img.height * ratio)), Image.LANCZOS)
                buf = BytesIO()
                img.save(buf, format="JPEG", quality=quality, optimize=True)
                return base64.b64encode(buf.getvalue()).decode("utf-8")
        except Exception as e:
            print(f"Screenshot-Fehler: {e}")
            return None
