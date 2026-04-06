
import time
import requests
import win32gui
import win32process
import psutil
import json
import os
import ctypes
import base64
import subprocess
import socket
import uuid

def get_device_id():
    hostname = socket.gethostname()
    mac = uuid.getnode()
    return f"{hostname}-{mac}"

try:
    import win32con
    import win32ui
except Exception:
    win32con = None
    win32ui = None

try:
    from PIL import Image
except Exception:
    Image = None

# 🔐 Config & Authentication
CONFIG_FILE = "agent_config.json"

class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]

def get_idle_duration():
    """Returns the time in seconds since the last keyboard/mouse input."""
    lii = LASTINPUTINFO()
    lii.cbSize = ctypes.sizeof(lii)
    if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)):
        millis = ctypes.windll.kernel32.GetTickCount() - lii.dwTime
        return millis / 1000.0
    return 0

def load_config():
    if not os.path.exists(CONFIG_FILE):
        default_config = {
            "employee_id": "INFI0023",
            "api_token": "SECRET123",
            "backend_url": "http://localhost:8000/activity",
            "idle_threshold_seconds": 60
        }
        with open(CONFIG_FILE, "w") as f:
            json.dump(default_config, f, indent=4)
        return default_config
    with open(CONFIG_FILE, "r") as f:
        return json.load(f)

def get_active_window_info():
    try:
        window = win32gui.GetForegroundWindow()
        if not window:
            return "Desktop", "Home", None
        _, pid = win32process.GetWindowThreadProcessId(window)
        if pid <= 0:
            return "Idle", "System Activity", None
        process = psutil.Process(pid)
        return process.name(), win32gui.GetWindowText(window), process.exe()
    except Exception:
        return "Idle", "Busy", None


def extract_icon_png_base64(exe_path: str, size: int = 64) -> str | None:
    """
    Extract the app's real Windows icon from the EXE and return it as base64 PNG.
    Requires: pywin32 + pillow (PIL).
    """
    if not exe_path or not os.path.exists(exe_path):
        return None
    if win32ui is None or win32con is None or Image is None:
        return None

    try:
        large, small = win32gui.ExtractIconEx(exe_path, 0)
        hicon = (large[0] if large else (small[0] if small else None))
        if not hicon:
            return None

        hdc_screen = win32gui.GetDC(0)
        dc = win32ui.CreateDCFromHandle(hdc_screen)
        memdc = dc.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(dc, size, size)
        memdc.SelectObject(bmp)

        memdc.FillSolidRect((0, 0, size, size), 0x00000000)
        win32gui.DrawIconEx(memdc.GetHandleOutput(), 0, 0, hicon, size, size, 0, None, win32con.DI_NORMAL)

        bmpinfo = bmp.GetInfo()
        bmpstr = bmp.GetBitmapBits(True)
        img = Image.frombuffer(
            "RGBA",
            (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
            bmpstr,
            "raw",
            "BGRA",
            0,
            1,
        )
        img = img.transpose(Image.FLIP_TOP_BOTTOM)

        from io import BytesIO
        buf = BytesIO()
        img.save(buf, format="PNG")
        return base64.b64encode(buf.getvalue()).decode("utf-8")
    except Exception:
        return None
    finally:
        try:
            win32gui.ReleaseDC(0, hdc_screen)
        except Exception:
            pass

def main():
    config = load_config()
    emp_id = config["employee_id"]
    token = config["api_token"]
    # Assuming the backend is at a known domain for bootstrap
    domain = os.getenv("AGENT_BACKEND_DOMAIN", "192.168.1.14:8000")
    bootstrap_url = f"http://{domain}/agent/config/{emp_id}"
    backend_url = config["backend_url"]
    idle_limit = config.get("idle_threshold_seconds", 60)

    print(f"--- Enterprise Activity Agent Started ---")
    print(f"Tracking: {emp_id}")

    last_reported_activity = None
    last_sent_time = 0
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    icon_upload_url = backend_url.replace("/activity", "/app-icons/upload") if "/activity" in backend_url else None
    uploaded_icons = set()

    while True:
        try:
            # 1. Check for True Idle (No input)
            idle_secs = get_idle_duration()
            if idle_secs > idle_limit:
                app, title = "Idle", f"Inactive for {int(idle_secs)}s"
                exe_path = None
            else:
                app, title, exe_path = get_active_window_info()
    
            # 1.1 Get current SSID (best effort)
            ssid = None
            try:
                raw_out = subprocess.check_output('netsh wlan show interfaces', shell=True).decode('ascii', errors='ignore')
                for line in raw_out.split('\n'):
                    if " SSID" in line and " BSSID" not in line:
                        ssid = line.split(':')[-1].strip()
            except:
                pass
    
            current_activity = {"app": app, "window": title, "ssid": ssid}
            time_since_last_sent = time.time() - last_sent_time
    
            # 2. Conditional Reporting (Change OR Heartbeat every 15s)
            if current_activity != last_reported_activity or time_since_last_sent > 15:
                success = False
                # 🚀 Retry Logic (3 attempts)
                for attempt in range(3):
                    try:
                        payload = {
                            "employee_id": emp_id,
                            "app": app,
                            "window": title,
                            "ssid": ssid,
                            "device_id": get_device_id()
                        }
                        response = requests.post(backend_url, json=payload, headers=headers, timeout=5)
                        if response.status_code == 200:
                            success = True
                            last_sent_time = time.time()
                            break
                        elif response.status_code == 403:
                            print("❌ AUTH ERROR: Token rejected by server.")
                            break
                    except Exception as e:
                        print(f"⚠️ Connection glitch (attempt {attempt+1}/3): {e}")
                        time.sleep(2)
    
                if success:
                    # Upload the real app icon once per app (best-effort).
                    low_app = (app or "").strip().lower()
                    if icon_upload_url and exe_path and low_app and low_app not in uploaded_icons and low_app not in ("idle", "desktop"):
                        b64png = extract_icon_png_base64(exe_path)
                        if b64png:
                            try:
                                up = {"employee_id": emp_id, "app": low_app, "icon_base64_png": b64png}
                                r2 = requests.post(icon_upload_url, json=up, headers=headers, timeout=8)
                                if r2.status_code == 200:
                                    uploaded_icons.add(low_app)
                            except Exception:
                                pass
                    if current_activity != last_reported_activity:
                        print(f"Activity Sync: {app} - {title[:30]}...")
                    last_reported_activity = current_activity
    
            time.sleep(3)
        except Exception as e:
            print(f"💥 GLOBAL AGENT ERROR: {e}")
            time.sleep(5)

if __name__ == "__main__":
    main()
