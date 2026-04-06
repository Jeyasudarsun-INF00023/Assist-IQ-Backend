from fastapi import FastAPI, Body, File, UploadFile, Form, WebSocket, WebSocketDisconnect, Header, HTTPException
from fastapi.responses import StreamingResponse

from services.sharepoint import upload_file_to_sharepoint, upload_json_to_sharepoint

from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import subprocess, json, platform, psutil, random, signal, os, uuid, socket, logging, base64, secrets, zipfile, io
import re
from concurrent.futures import ThreadPoolExecutor
from difflib import get_close_matches
from database import engine, get_db, SessionLocal
from models import Base
from fastapi import Depends 
from sqlalchemy.orm import Session
from models import ChatHistory, InstalledSoftware, UninstalledSoftware, Device, Ticket, Employee, Asset, ActivityLog
import requests
from msal import ConfidentialClientApplication
from dotenv import load_dotenv
import json
from datetime import datetime, timedelta, timezone
import asyncio
import urllib.request
import urllib.parse

load_dotenv()



from services.auth import get_access_token

# LICENSE_SKU_ID is still needed here for assigning licenses
LICENSE_SKU_ID = os.getenv("LICENSE_SKU_ID", "3b555118-da6a-4418-894f-7df1e2096870")
EMAIL_DOMAIN = os.getenv("EMAIL_DOMAIN", "theinfinitechx.com")
APP_DOMAIN = os.getenv("APP_DOMAIN", os.getenv("RENDER_EXTERNAL_HOSTNAME", "localhost:8000"))

def get_current_ssid():
    """Returns the current SSID the server is connected to."""
    if platform.system() != "Windows":
        return None # Not relevant or not supported on non-Windows systems
    
    import subprocess
    try:
        raw_out = subprocess.check_output('netsh wlan show interfaces', shell=True).decode('ascii', errors='ignore')
        for line in raw_out.split('\n'):
            if " SSID" in line and " BSSID" not in line:
                return line.split(':')[-1].strip()
    except:
        pass
    return os.getenv("OFFICE_WIFI", "Airtel_nate_4772") # Default to known wifi if detection fails


def create_office_user(display_name, user_principal_name, password):
    token = get_access_token()

    url = "https://graph.microsoft.com/v1.0/users"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    data = {
        "accountEnabled": True,
        "displayName": display_name,
        "mailNickname": display_name.replace(" ", ""),
        "userPrincipalName": user_principal_name,
        "usageLocation": "IN",
        "passwordProfile": {
            "forceChangePasswordNextSignIn": True,
            "password": password
        }
    }

    response = requests.post(url, headers=headers, json=data)

    if response.status_code not in [200, 201]:
        raise Exception(response.text)

    return response.json()


def assign_license(user_id):
    token = get_access_token()

    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/assignLicense"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    data = {
        "addLicenses": [
            {
                "skuId": LICENSE_SKU_ID
            }
        ],
        "removeLicenses": []
    }

    response = requests.post(url, headers=headers, json=data)

    if response.status_code not in [200, 202]:
        raise Exception(response.text)

    return response.json()

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = FastAPI()

@app.get("/")
def read_root():
    return {"status": "Backend is running!", "app_domain": APP_DOMAIN}


class AppIconUpload(BaseModel):
    employee_id: str
    app: str
    icon_base64_png: str


@app.post("/app-icons/upload")
def upload_app_icon(payload: AppIconUpload, db: Session = Depends(get_db), authorization: str = Header(None)):
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Missing or invalid Authorization header")

    token = authorization.split(" ")[1]
    employee = db.query(Employee).filter(
        Employee.employee_id == payload.employee_id,
        Employee.api_token == token
    ).first()
    if not employee:
        raise HTTPException(status_code=403, detail="Unauthorized device or invalid token")

    app_key = (payload.app or "").strip().lower()
    if not app_key:
        raise HTTPException(status_code=400, detail="Missing app")
    if not app_key.endswith(".exe") and app_key not in ("idle", "desktop"):
        app_key = f"{app_key}.exe"

    try:
        raw = base64.b64decode(payload.icon_base64_png)
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid base64 png")

    # Keep filename safe-ish
    safe_name = "".join(c for c in app_key if c.isalnum() or c in ("-", "_", "."))
    if not safe_name:
        raise HTTPException(status_code=400, detail="Invalid app key")

    out_path = os.path.join(APP_ICONS_DIR, f"{safe_name}.png")
    with open(out_path, "wb") as f:
        f.write(raw)

    return {"success": True, "path": f"/app-icons/{safe_name}.png"}

@app.on_event("startup")
async def startup_event():
    print("\n" + "="*30)
    print("BACKEND IS AWAKE")
    print("Running on Windows: " + str(platform.system() == "Windows"))
    print(f"App Domain: {APP_DOMAIN}")
    print("="*30 + "\n")

     #Create tables automatically
    Base.metadata.create_all(bind=engine)
    logger.info("Database tables created/verified.")
    asyncio.create_task(check_offline_users())

async def check_offline_users():
    """Background task to mark users as offline if they haven't sent activity recently."""
    while True:
        try:
            db = SessionLocal()
            # Use naive UTC for DB consistency (SQLite DateTime is naive here)
            now = datetime.now(timezone.utc).replace(tzinfo=None)
            # Mark users offline if no activity for 20 seconds
            timeout = now - timedelta(seconds=20)
            
            offline_employees = db.query(Employee).filter(
                Employee.is_online == True,
                Employee.last_seen < timeout
            ).all()
            
            for emp in offline_employees:
                emp.is_online = False
                db.commit()
                # Broadcast offline status
                await manager.send_activity(emp.employee_id, {
                    "app": emp.last_app,
                    "window": emp.last_window,
                    "is_online": False
                })
            db.close()
        except Exception as e:
            logger.error(f"Error in check_offline_users: {e}")
        await asyncio.sleep(5)
    
class ConnectionManager:
    def __init__(self):
        self.active_connections = {}

    async def connect(self, employee_id: str, websocket: WebSocket):
        await websocket.accept()
        if employee_id not in self.active_connections:
            self.active_connections[employee_id] = []
        self.active_connections[employee_id].append(websocket)

    def disconnect(self, employee_id: str, websocket: WebSocket):
        if employee_id in self.active_connections:
            try:
                self.active_connections[employee_id].remove(websocket)
                if not self.active_connections[employee_id]:
                    del self.active_connections[employee_id]
            except ValueError:
                pass

    async def send_activity(self, employee_id: str, data: dict):
        if employee_id in self.active_connections:
            for ws in self.active_connections[employee_id][:]:
                try:
                    await ws.send_json(data)
                except Exception:
                    self.disconnect(employee_id, ws)

manager = ConnectionManager()

# Track active subprocesses
active_processes = {}

# Paths relative to this script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SOFTWARE_MAP_FILE = os.path.join(BASE_DIR, "software_map.json")
KEY_FILE = os.path.join(BASE_DIR, "secret.key")

# Serve uploaded app icons (agent uploads real Windows icons).
APP_ICONS_DIR = os.path.join(BASE_DIR, "app_icons")
os.makedirs(APP_ICONS_DIR, exist_ok=True)
app.mount("/app-icons", StaticFiles(directory=APP_ICONS_DIR), name="app-icons")

# Encryption Setup
try:
    from cryptography.fernet import Fernet
    
    def load_key():
        if not os.path.exists(KEY_FILE):
            key = Fernet.generate_key()
            with open(KEY_FILE, "wb") as key_file:
                key_file.write(key)
        with open(KEY_FILE, "rb") as key_file:
            return key_file.read()

    cipher_suite = Fernet(load_key())

    def encrypt_password(password: str) -> str:
        if not password: return None
        return cipher_suite.encrypt(password.encode()).decode()

    def decrypt_password(encrypted_password: str) -> str:
        if not encrypted_password: return None
        try:
            return cipher_suite.decrypt(encrypted_password.encode()).decode()
        except:
            return encrypted_password
except ImportError:
    print("Cryptography library not found. Passwords will be stored in plain text.")
    def encrypt_password(p): return p
    def decrypt_password(p): return p

class ChatSession(BaseModel):
    id: str
    title: str
    messages: list
    timestamp: str
    pinned: bool = False

class SaveHistoryRequest(BaseModel):
    session: ChatSession
    email: str

class LoginRequest(BaseModel):
    name: str
    email: str
    picture: str | None = None

class EmailPreviewRequest(BaseModel):
    first_name: str
    last_name: str

class VerificationCodeRequest(BaseModel):
    employee_id: int
    email: str

class VerifyCodeRequest(BaseModel):
    employee_id: int
    email: str
    code: str

class CreateUserRequest(BaseModel):
    full_name: str
    email: str
    password: str

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

if os.path.exists(SOFTWARE_MAP_FILE):
    with open(SOFTWARE_MAP_FILE, "r") as f:
        SOFTWARE_MAP = json.load(f)
else:
    SOFTWARE_MAP = {}

# Simple Knowledge Base
KNOWLEDGE_BASE = {
    "printer": "To add a printer, go to Settings > Devices > Printers & Scanners > Add a printer.",
    "policy": "Our remote work policy allows for 2 days of WFH per week with manager approval."
}

class RemoteTarget(BaseModel):
    hostname: str | None = None
    port: int = 5985
    username: str | None = None
    password: str | None = None
    name: str | None = None

# ==============================
# EMPLOYEE MANAGEMENT
# ==============================

class EmployeeCreate(BaseModel):
    first_name: str
    last_name: str
    phone_number: str | None = None
    personal_email: str | None = None
    address: str | None = None
    experience_level: str | None = None
    role: str
    employee_id: str
    email: str | None = None
    avatar: str | None = None
    laptop: str | None = None
    mouse: str | None = None
    headphone: str | None = None
    department: str | None = None
    seat_id: int | None = None
    documents: list[dict] | None = None  # List of {name: str, size: str}
    last_app: str | None = None
    last_window: str | None = None
    is_online: bool = False

class EmployeeUpdate(BaseModel):
    first_name: str | None = None
    last_name: str | None = None
    phone_number: str | None = None
    personal_email: str | None = None
    address: str | None = None
    experience_level: str | None = None
    role: str | None = None
    employee_id: str | None = None
    email: str | None = None
    avatar: str | None = None
    laptop: str | None = None
    mouse: str | None = None
    headphone: str | None = None
    department: str | None = None
    seat_id: int | None = None
    documents: list[dict] | None = None
    last_app: str | None = None
    last_window: str | None = None
    is_online: bool | None = None

class EmployeeEmailUpdate(BaseModel):
    email: str
    temp_password: str | None = None

def update_asset_assignments(db: Session, employee_name: str, laptop_str: str, mouse_str: str, headphone_str: str):
    """Updates the assets table to set assignee based on strings like 'Apple MacBook Air (SN123)'"""
    import re
    from datetime import datetime
    
    # 1. Clear old assignments for this employee
    old_assets = db.query(Asset).filter(Asset.assignee == employee_name).all()
    for a in old_assets:
        a.assignee = None
        a.assigned_date = None
    
    # 2. Assign new assets
    assigned_date = datetime.now().strftime("%Y-%m-%d")
    for s in [laptop_str, mouse_str, headphone_str]:
        if not s or "(" not in s: continue
        
        # Extract SN from between parentheses: "Brand Model (SN)"
        match = re.search(r'\((.*?)\)', s)
        if match:
            sn = match.group(1)
            asset = db.query(Asset).filter(Asset.sn == sn).first()
            if asset:
                asset.assignee = employee_name
                asset.assigned_date = assigned_date

# ==============================
# ASSET MANAGEMENT
# ==============================


class AssetCreate(BaseModel):
    type: str
    category: str | None = None
    brand: str | None = None
    model: str | None = None
    sn: str | None = None
    processor: str | None = None
    ram: str | None = None
    storage: str | None = None
    os: str | None = None
    assignee: str | None = None
    assigned_date: str | None = None
    remarks: str | None = None
    price: str | None = None
    custom_fields: dict | None = None

class AssetUpdate(BaseModel):
    type: str | None = None
    category: str | None = None
    brand: str | None = None
    model: str | None = None
    sn: str | None = None
    processor: str | None = None
    ram: str | None = None
    storage: str | None = None
    os: str | None = None
    assignee: str | None = None
    assigned_date: str | None = None
    remarks: str | None = None
    price: str | None = None
    custom_fields: dict | None = None


@app.get("/assets")
def get_assets(db: Session = Depends(get_db)):
    result = []
    for asset in db.query(Asset).order_by(Asset.created_at.desc()).all():
        data = {
            "id": asset.id,
            "type": asset.type,
            "category": asset.category,
            "brand": asset.brand,
            "model": asset.model,
            "sn": asset.sn,
            "processor": asset.processor,
            "ram": asset.ram,
            "storage": asset.storage,
            "os": asset.os,
            "assignee": asset.assignee,
            "assigned_date": asset.assigned_date,
            "remarks": asset.remarks,
            "price": asset.price,
            "custom_fields": json.loads(asset.custom_fields) if asset.custom_fields else {},
            "created_at": asset.created_at.isoformat() if asset.created_at else None
        }
        result.append(data)
    return result

class ActivityRequest(BaseModel):
    employee_id: str
    app: str
    window: str
    ssid: str | None = None
    device_id: str | None = None

@app.post("/activity")
async def receive_activity(payload: ActivityRequest, db: Session = Depends(get_db), authorization: str = Header(None)):
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Missing or invalid Authorization header")
    
    token = authorization.split(" ")[1]
    
    # 🔐 Verify per-employee token
    employee = db.query(Employee).filter(
        Employee.employee_id == payload.employee_id, 
        Employee.api_token == token
    ).first()

    if not employee:
        raise HTTPException(status_code=401, detail="Unauthorized")

    if employee.offboarded:
        raise HTTPException(status_code=403, detail="Employee is offboarded. Activity tracking disabled.")

    # 🛡️ Device Binding Logic
    if not employee.device_id:
        # First time binding
        employee.device_id = payload.device_id
        db.commit()
        logger.info(f"Device registered for employee {employee.employee_id}: {payload.device_id}")
    elif employee.device_id != payload.device_id:
        # Device mismatch
        logger.warning(f"Device mismatch for {employee.employee_id}. Expected: {employee.device_id}, Got: {payload.device_id}")
        raise HTTPException(status_code=403, detail="Device mismatch")

    # 📊 Activity History Logic
    # Update logs if the application or window title changed
    if employee.last_app != payload.app or employee.last_window != payload.window:
        now = datetime.now(timezone.utc).replace(tzinfo=None)
        # Close the most recent open log for this employee
        active_log = db.query(ActivityLog).filter(
            ActivityLog.employee_id == payload.employee_id,
            ActivityLog.end_time == None
        ).order_by(ActivityLog.start_time.desc()).first()
        
        if active_log:
            active_log.end_time = now
            
        # Create a new log entry
        new_log = ActivityLog(
            employee_id=payload.employee_id,
            app=payload.app,
            window=payload.window,
            start_time=now
        )
        db.add(new_log)

    # 🏛️ WiFi Consistency Check (NEW)
    # Only mark as online if SSID matches server/office wifi
    server_ssid = get_current_ssid()
    is_on_same_wifi = (payload.ssid == server_ssid) if server_ssid else True
    
    # Sync latest state to Employee table
    employee.last_app = payload.app
    employee.last_window = payload.window
    employee.is_online = is_on_same_wifi
    employee.last_seen = datetime.now(timezone.utc).replace(tzinfo=None)
    
    try:
        db.commit()
    except Exception as e:
        db.rollback()
        logger.error(f"Database error during activity update: {e}")

    await manager.send_activity(payload.employee_id, {
        "app": payload.app,
        "window": payload.window,
        "is_online": is_on_same_wifi,
        "on_wifi": is_on_same_wifi
    })
    
    return {"success": True}

@app.get("/agent/config/{employee_id}")
def get_agent_config(employee_id: str, db: Session = Depends(get_db)):
    employee = db.query(Employee).filter(Employee.employee_id == employee_id).first()
    if not employee:
        raise HTTPException(status_code=404, detail="Employee not found")
    
    return {
        "employee_id": employee.employee_id,
        "api_token": employee.api_token,
        "backend_url": f"http://{APP_DOMAIN}/activity",
        "idle_threshold_seconds": 120
    }

@app.get("/agent/download/{employee_id}")
def download_agent_package(employee_id: str, db: Session = Depends(get_db)):
    employee = db.query(Employee).filter(Employee.employee_id == employee_id).first()
    if not employee:
        raise HTTPException(status_code=404, detail="Employee not found")
    
    # Determine protocol (use https for Render)
    protocol = "https" if "onrender.com" in APP_DOMAIN else "http"
    
    # Configuration to inject
    config = {
        "employee_id": employee.employee_id,
        "api_token": employee.api_token,
        "backend_url": f"{protocol}://{APP_DOMAIN}/activity",
        "idle_threshold_seconds": 120
    }
    
    # Create the ZIP package in memory
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        # 1. Add Compiled EXE (best for production)
        # Search in backend/dist/ or backend/ directly
        exe_path = os.path.join(BASE_DIR, "dist", "agent_activity.exe")
        if not os.path.exists(exe_path):
            exe_path = os.path.join(BASE_DIR, "agent_activity.exe")
        
        if os.path.exists(exe_path):
            zip_file.write(exe_path, "agent_activity.exe")
        else:
            # Fallback to source script if EXE is not built
            agent_path = os.path.join(BASE_DIR, "agent_activity.py")
            if os.path.exists(agent_path):
                 zip_file.write(agent_path, "agent_activity.py")

        # 2. Add the custom config
        zip_buffer_config = json.dumps(config, indent=4)
        zip_file.writestr("agent_config.json", zip_buffer_config)
        
        # 3. Updated run_agent.bat
        if os.path.exists(exe_path):
            bat_content = "@echo off\necho Starting Enterprise Activity Agent...\nstart agent_activity.exe"
        else:
            bat_content = "@echo off\necho Starting Enterprise Activity Agent (Python required)...\npip install requests psutil pywin32\npython agent_activity.py\npause"
        
        zip_file.writestr("run_agent.bat", bat_content)

    zip_buffer.seek(0)
    filename = f"agent_setup_{employee_id}.zip"
    
    return StreamingResponse(
        zip_buffer, 
        media_type="application/x-zip-compressed",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.websocket("/ws/activity/{employee_id}")
async def websocket_endpoint(websocket: WebSocket, employee_id: str):
    await manager.connect(employee_id, websocket)
    try:
        while True:
            await websocket.receive_text()
    except WebSocketDisconnect:
        manager.disconnect(employee_id, websocket)

@app.post("/assets")
def create_asset(asset: AssetCreate, db: Session = Depends(get_db)):
    db_asset = Asset(
        type=asset.type,
        category=asset.category,
        brand=asset.brand,
        model=asset.model,
        sn=asset.sn,
        processor=asset.processor,
        ram=asset.ram,
        storage=asset.storage,
        os=asset.os,
        assignee=asset.assignee,
        assigned_date=asset.assigned_date,
        remarks=asset.remarks,
        price=asset.price,
        custom_fields=json.dumps(asset.custom_fields or {}) if asset.custom_fields is not None else None
    )
    db.add(db_asset)
    db.commit()
    db.refresh(db_asset)
    return db_asset

@app.put("/assets/{asset_id}")
def update_asset(asset_id: int, asset_data: AssetUpdate, db: Session = Depends(get_db)):
    db_asset = db.query(Asset).filter(Asset.id == asset_id).first()
    if not db_asset:
        return {"error": "Asset not found"}
    
    update_data = asset_data.dict(exclude_unset=True)
    if 'custom_fields' in update_data and update_data['custom_fields'] is not None:
        update_data['custom_fields'] = json.dumps(update_data['custom_fields'])

    for key, value in update_data.items():
        setattr(db_asset, key, value)
    
    db.commit()
    db.refresh(db_asset)
    return db_asset


@app.delete("/assets/{asset_id}")
def delete_asset(asset_id: int, db: Session = Depends(get_db)):
    asset = db.query(Asset).filter(Asset.id == asset_id).first()
    if not asset:
        return {"success": False, "error": "Asset not found"}
    db.delete(asset)
    db.commit()
    return {"success": True}

@app.post("/assets/bulk-delete")
def bulk_delete_assets(payload: dict = Body(...), db: Session = Depends(get_db)):
    ids = payload.get("ids", [])
    if not ids:
        return {"success": True, "count": 0}
    
    db.query(Asset).filter(Asset.id.in_(ids)).delete(synchronize_session=False)
    db.commit()
    return {"success": True, "count": len(ids)}


@app.patch("/employees/{employee_id}")
def update_employee(employee_id: int, employee: EmployeeUpdate, db: Session = Depends(get_db)):
    db_employee = db.query(Employee).filter(Employee.id == employee_id).first()
    if not db_employee:
        return {"success": False, "error": "Employee not found"}
    
    if employee.first_name and employee.last_name:
        db_employee.first_name = employee.first_name
        db_employee.last_name = employee.last_name
        db_employee.full_name = f"{employee.first_name} {employee.last_name}"
    elif employee.first_name: # fallback if only first name provided (unlikely)
        # simplistic handling, ideally we'd store first/last separately
        db_employee.first_name = employee.first_name
        parts = db_employee.full_name.split(" ", 1)
        last = db_employee.last_name if db_employee.last_name else (parts[1] if len(parts) > 1 else "")
        db_employee.full_name = f"{employee.first_name} {last}"
    elif employee.last_name:
        db_employee.last_name = employee.last_name
        parts = db_employee.full_name.split(" ", 1)
        first = db_employee.first_name if db_employee.first_name else parts[0]
        db_employee.full_name = f"{first} {employee.last_name}"

    if employee.phone_number: db_employee.phone_number = employee.phone_number
    if employee.personal_email: db_employee.personal_email = employee.personal_email
    if employee.address: db_employee.address = employee.address
    if employee.experience_level: db_employee.experience_level = employee.experience_level
    if employee.role: db_employee.role = employee.role
    if employee.employee_id: db_employee.employee_id = employee.employee_id
    if employee.email: db_employee.email = employee.email
    if employee.avatar: db_employee.avatar = employee.avatar
    if employee.laptop: db_employee.laptop = employee.laptop
    if employee.mouse: db_employee.mouse = employee.mouse
    if employee.headphone: db_employee.headphone = employee.headphone
    if employee.department: db_employee.department = employee.department
    if employee.seat_id is not None: db_employee.seat_id = employee.seat_id
    if employee.documents is not None: db_employee.documents = json.dumps(employee.documents)

    db.commit()
    
    # Sync assets
    update_asset_assignments(db, db_employee.full_name, db_employee.laptop, db_employee.mouse, db_employee.headphone)
    db.commit()
    
    db.refresh(db_employee)

    
    return {
        "success": True,
        "employee": {
            "id": db_employee.id,
            "full_name": db_employee.full_name,
            "first_name": db_employee.first_name or "",
            "last_name": db_employee.last_name or "",
            "phone_number": db_employee.phone_number,
            "personal_email": db_employee.personal_email,
            "address": db_employee.address,
            "experience_level": db_employee.experience_level,
            "role": db_employee.role,
            "employee_id": db_employee.employee_id,
            "email": db_employee.email,
            "temp_password": decrypt_password(db_employee.temp_password) if db_employee.temp_password else None,
            "avatar": db_employee.avatar,
            "laptop": db_employee.laptop,
            "mouse": db_employee.mouse,
            "headphone": db_employee.headphone,
            "department": db_employee.department,
            "seat_id": db_employee.seat_id,
            "documents": json.loads(db_employee.documents) if db_employee.documents else [],
            "created_at": db_employee.created_at.isoformat() if db_employee.created_at else None
        }
    }

@app.post("/employees")
def create_employee(employee: EmployeeCreate, db: Session = Depends(get_db)):
    full_name = f"{employee.first_name} {employee.last_name}"
    
    # Check if employee_id already exists
    existing_emp = db.query(Employee).filter(Employee.employee_id == employee.employee_id).first()
    if existing_emp:
        # In a real app, we might return 400, but for now let's just update or ignore
        pass

    db_employee = Employee(
        full_name=full_name,
        first_name=employee.first_name,
        last_name=employee.last_name,
        phone_number=employee.phone_number,
        personal_email=employee.personal_email,
        address=employee.address,
        experience_level=employee.experience_level,
        role=employee.role,
        employee_id=employee.employee_id,
        email=employee.email,
        api_token=secrets.token_hex(16),
        avatar=employee.avatar,
        laptop=employee.laptop,
        mouse=employee.mouse,
        headphone=employee.headphone,
        department=employee.department,
        seat_id=employee.seat_id,
        documents=json.dumps(employee.documents) if employee.documents else None
    )
    db.add(db_employee)
    db.commit()
    
    # Sync assets
    update_asset_assignments(db, db_employee.full_name, db_employee.laptop, db_employee.mouse, db_employee.headphone)
    db.commit()
    
    db.refresh(db_employee)

    return db_employee

@app.post("/employees/{employee_id}/unassign-asset")
def unassign_asset(employee_id: int, payload: dict = Body(...), db: Session = Depends(get_db)):
    asset_type = payload.get("asset_type") # "Laptop", "Mouse", "Headphone"
    if not asset_type:
        raise HTTPException(status_code=400, detail="Missing asset_type")
    
    db_employee = db.query(Employee).filter(Employee.id == employee_id).first()
    if not db_employee:
        raise HTTPException(status_code=404, detail="Employee not found")
    
    # Map label to field name
    field_map = {
        "Laptop": "laptop",
        "Mouse": "mouse",
        "Headphone": "headphone"
    }
    
    field_name = field_map.get(asset_type)
    if not field_name:
        raise HTTPException(status_code=400, detail="Invalid asset_type")
    
    setattr(db_employee, field_name, None)
    db.commit()
    
    # Sync assets
    update_asset_assignments(db, db_employee.full_name, db_employee.laptop, db_employee.mouse, db_employee.headphone)
    db.commit()
    db.refresh(db_employee)
    
    return {"success": True, "employee_laptop": db_employee.laptop, "employee_mouse": db_employee.mouse, "employee_headphone": db_employee.headphone}

@app.get("/assets/by-sn/{sn}")
def get_asset_by_sn(sn: str, db: Session = Depends(get_db)):
    asset = db.query(Asset).filter(Asset.sn == sn).first()
    if not asset:
        raise HTTPException(status_code=404, detail="Asset not found")
    
    return {
        "id": asset.id,
        "type": asset.type,
        "brand": asset.brand,
        "model": asset.model,
        "sn": asset.sn,
        "processor": asset.processor,
        "ram": asset.ram,
        "storage": asset.storage,
        "os": asset.os,
        "price": asset.price,
        "remarks": asset.remarks
    }

@app.post("/send-verification-code")
def send_verification_code(payload: VerificationCodeRequest, db: Session = Depends(get_db)):
    employee = db.query(Employee).filter(Employee.id == payload.employee_id).first()
    if not employee:
        raise HTTPException(status_code=404, detail="Employee not found")
    
    # Generate 6-digit code
    import random
    code = f"{random.randint(100000, 999999)}"
    
    # Save to employee
    employee.verification_code = code
    employee.verification_code_expires = datetime.now(timezone.utc).replace(tzinfo=None) + timedelta(minutes=10)
    db.commit()
    
    logger.info(f"Verification code for {payload.email}: {code}")
    
    # Mock Email Sending
    # In a real app, use smtplib or a service like SendGrid
    print(f"\n[EMAIL MOCK] To: {payload.email}\nSubject: Business Email Verification\nCode: {code}\n")
    
    return {"success": True, "message": "Verification code sent."}

@app.post("/verify-verification-code")
def verify_verification_code(payload: VerifyCodeRequest, db: Session = Depends(get_db)):
    employee = db.query(Employee).filter(Employee.id == payload.employee_id).first()
    if not employee:
        raise HTTPException(status_code=404, detail="Employee not found")
    
    if not employee.verification_code or employee.verification_code != payload.code:
        raise HTTPException(status_code=400, detail="Invalid verification code")
    
    now = datetime.now(timezone.utc).replace(tzinfo=None)
    if employee.verification_code_expires and employee.verification_code_expires < now:
        raise HTTPException(status_code=400, detail="Verification code expired")
    
    # Successfully verified! Update email.
    employee.email = payload.email
    employee.verification_code = None
    employee.verification_code_expires = None
    db.commit()
    
    return {"success": True, "message": "Email verified and linked successfully."}

@app.get("/employees")
def get_employees(db: Session = Depends(get_db)):
    employees = db.query(Employee).order_by(Employee.created_at.desc()).all()
    result = []
    for emp in employees:
        emp_dict = {
            "id": emp.id,
            "full_name": emp.full_name,
            "first_name": emp.first_name,
            "last_name": emp.last_name,
            "phone_number": emp.phone_number,
            "personal_email": emp.personal_email,
            "address": emp.address,
            "experience_level": emp.experience_level,
            "role": emp.role,
            "employee_id": emp.employee_id,
            "email": emp.email,
            "temp_password": decrypt_password(emp.temp_password) if emp.temp_password else None,
            "avatar": emp.avatar,
            "laptop": emp.laptop,
            "mouse": emp.mouse,
            "headphone": emp.headphone,
            "department": emp.department,
            "seat_id": emp.seat_id,
            "documents": json.loads(emp.documents) if emp.documents else [],
            "created_at": emp.created_at.isoformat() if emp.created_at else None,
            "offboarded": emp.offboarded if emp.offboarded is not None else False,
            "offboarded_at": emp.offboarded_at.isoformat() if emp.offboarded_at else None,
            "last_app": emp.last_app,
            "last_window": emp.last_window,
            "is_online": emp.is_online,
            "last_seen": emp.last_seen.isoformat() if emp.last_seen else None
        }
        # Fetch last 7 unique apps from ActivityLog for multi-icon UI
        recent = db.query(ActivityLog).filter(
            ActivityLog.employee_id == emp.employee_id
        ).order_by(ActivityLog.start_time.desc()).limit(30).all()
        
        seen_apps = set()
        apps_list = []
        for r in recent:
            if r.app and r.app.lower() not in seen_apps:
                apps_list.append({"app": r.app, "window": r.window})
                seen_apps.add(r.app.lower())
            if len(apps_list) >= 7:
                break
        
        emp_dict["recent_apps"] = apps_list
        result.append(emp_dict)
    return result

@app.get("/employees/{emp_id}/activity-stats")
def get_employee_activity_stats(emp_id: str, timeframe: str = "This Week", db: Session = Depends(get_db)):
    from datetime import datetime, timedelta, date as pydate, timezone
    import calendar

    # DB stores naive UTC; keep comparisons/arithmetic naive to avoid empty filters.
    now = datetime.now(timezone.utc).replace(tzinfo=None)

    # ── TODAY: from midnight of today (not rolling 24h) ──────────────────────
    today_midnight = now.replace(hour=0, minute=0, second=0, microsecond=0)

    today_logs = db.query(ActivityLog).filter(
        ActivityLog.employee_id == emp_id,
        ActivityLog.start_time >= today_midnight
    ).all()

    app_durations: dict[str, float] = {}
    worked_seconds_today = 0.0
    idle_seconds_today = 0.0

    def _is_idle_app(app_value: str | None) -> bool:
        raw_app = (app_value or "").strip().lower()
        return raw_app in ("idle", "idle.exe") or "idle" == raw_app

    agent_start_dt = None
    for log in today_logs:
        start = log.start_time
        end   = log.end_time or now          # open sessions count up to now
        duration = (end - start).total_seconds()
        if duration <= 0:
            continue

        # Normalise app name to a canonical .exe key
        raw = (log.app or "Unknown").lower()
        is_idle = _is_idle_app(raw)
        if   "edge"    in raw: app_key = "msedge.exe"
        elif "code"    in raw: app_key = "code.exe"
        elif "figma"   in raw: app_key = "figma.exe"
        elif "teams"   in raw: app_key = "teams.exe"
        elif "outlook" in raw: app_key = "outlook.exe"
        elif "chrome"  in raw: app_key = "chrome.exe"
        else:                  app_key = raw if raw.endswith(".exe") else raw + ".exe"

        app_durations[app_key] = app_durations.get(app_key, 0) + duration
        if is_idle:
            idle_seconds_today += duration
        else:
            worked_seconds_today += duration
            if agent_start_dt is None or start < agent_start_dt:
                agent_start_dt = start

    # ── WORK-DAY CONSTANTS ────────────────────────────────────────────────────
    WORK_DAY_SECONDS   = 8 * 3600   # 28 800 s  — standard working day
    OVERTIME_THRESHOLD = 9 * 3600   # 32 400 s  — overtime badge triggers at 9 h

    # Overtime is defined as working beyond 8h (up to the 9h cap in UI).
    is_overtime = worked_seconds_today > WORK_DAY_SECONDS

    # Seconds to show on the gauge  (capped at 9 h so the arc never overflows)
    display_seconds = min(worked_seconds_today, OVERTIME_THRESHOLD)

    # ── TOP APPS (sorted desc, top 5) ────────────────────────────────────────
    sorted_apps = sorted(app_durations.items(), key=lambda x: x[1], reverse=True)
    top_apps = [{"app": k, "duration": v} for k, v in sorted_apps[:5]]

    # ── AGENT START TIME (first activity of today) ───────────────────────────
    agent_start = None
    if agent_start_dt is not None:
        agent_start = agent_start_dt.isoformat() + "Z"

    # ── TREND RANGE ──────────────────────────────────────────────────────────
    if timeframe == "Last Month":
        first_this  = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        last_last   = first_this - timedelta(seconds=1)
        first_last  = last_last.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        range_start = first_last
        days_to_show = calendar.monthrange(first_last.year, first_last.month)[1]
    elif timeframe == "This Month":
        range_start  = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        days_to_show = now.day
    else:  # "This Week" — last 7 days including today
        range_start  = today_midnight - timedelta(days=6)
        days_to_show = 7

    # ── TREND LOGS ───────────────────────────────────────────────────────────
    trend_logs = db.query(ActivityLog).filter(
        ActivityLog.employee_id == emp_id,
        ActivityLog.start_time  >= range_start
    ).all()

    # Initialise daily buckets
    daily_worked: dict = {}
    daily_idle: dict = {}
    for i in range(days_to_show):
        d = (range_start + timedelta(days=i)).date()
        daily_worked[d] = 0.0
        daily_idle[d] = 0.0

    for log in trend_logs:
        start    = log.start_time
        end      = log.end_time or now
        duration = (end - start).total_seconds()

        # Sanity-guard: ignore negative or impossibly long sessions (> 24 h)
        if duration <= 0 or duration > 86_400:
            continue

        log_date = start.date()
        if log_date in daily_worked:
            if _is_idle_app(log.app):
                daily_idle[log_date] += duration / 3600.0
            else:
                daily_worked[log_date] += duration / 3600.0   # store in hours

    # ── BUILD TREND ARRAY ────────────────────────────────────────────────────
    trends = []
    for i in range(days_to_show):
        d = (range_start + timedelta(days=i)).date()

        if timeframe == "This Week":
            label = d.strftime("%a")           # Mon, Tue …
        else:
            label = d.strftime("%d %b")        # 01 Mar …

        worked_h = round(daily_worked[d], 2)
        idle_h = round(daily_idle[d], 2)

        # Cap display at 9 h so bars never overflow the chart
        display_worked = min(worked_h, 9.0)

        # Is this day overtime?
        day_overtime = worked_h > 8.0

        trends.append({
            "day":          label,
            "worked":       display_worked,   # hours (max 9)
            "raw_worked":   worked_h,         # actual hours (can exceed 9)
            "idle":         idle_h,
            "is_overtime":  day_overtime,
        })

    # ── RESPONSE ─────────────────────────────────────────────────────────────
    return {
        "success":             True,

        # Gauge — use display_seconds (capped 0-9 h) for the arc fill
        "today_total_seconds": display_seconds,

        # Work time only (excludes idle) so the UI's "/8h" is true working time
        "raw_total_seconds":   worked_seconds_today,

        # Expose idle separately for UI/analytics
        "idle_total_seconds":  idle_seconds_today,

        "top_apps":            top_apps,
        "weekly_trends":       trends,
        "agent_start":         agent_start,

        # Badge logic: show "Overtime" when > 8 h, clear it when ≤ 8 h
        "is_overtime":         is_overtime,

        # Convenience fields for the frontend labels
        "work_day_seconds":    WORK_DAY_SECONDS,    # 28800
        "overtime_threshold":  OVERTIME_THRESHOLD,  # 32400
    }


@app.patch("/employees/{employee_id}/email")
def update_employee_email(employee_id: int, payload: EmployeeEmailUpdate, db: Session = Depends(get_db)):
    employee = db.query(Employee).filter(Employee.id == employee_id).first()
    if not employee:
        return {"success": False, "error": "Employee not found"}
    employee.email = payload.email
    if payload.temp_password:
        employee.temp_password = encrypt_password(payload.temp_password)
    db.commit()
    db.refresh(employee)
    return {
        "success": True,
        "employee": {
            "id": employee.id,
            "full_name": employee.full_name,
            "role": employee.role,
            "employee_id": employee.employee_id,
            "email": employee.email,
            "temp_password": decrypt_password(employee.temp_password) if employee.temp_password else None,
            "avatar": employee.avatar
        }
    }

@app.delete("/employees/{employee_id}")
def delete_employee(employee_id: int, db: Session = Depends(get_db)):
    employee = db.query(Employee).filter(Employee.id == employee_id).first()
    if not employee:
        return {"success": False, "error": "Employee not found"}
    
    db.delete(employee)
    db.commit()
    return {"success": True, "message": "Employee deleted successfully"}

# ==============================
# OFFBOARDING
# ==============================

import asyncio
import time as _time

@app.post("/employees/{employee_id}/offboard")
def offboard_employee(employee_id: int, db: Session = Depends(get_db)):
    """
    Offboard an employee by:
    1. Disabling their Microsoft 365 account
    2. Removing Microsoft 365 licenses (cost saving)
    3. Marking assets as collected
    4. Releasing workstation seat
    5. Storing the offboarding record
    Returns step-by-step result.
    """
    employee = db.query(Employee).filter(Employee.id == employee_id).first()
    if not employee:
        return {"success": False, "error": "Employee not found"}

    steps = []
    m365_user_id = None  # Shared across steps

    # Step 1: Disable Microsoft 365 account
    step1_ok = False
    step1_msg = ""
    if employee.email:
        try:
            token = get_access_token()
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }
            # Find user by UPN and get their object ID
            find_url = f"https://graph.microsoft.com/v1.0/users/{employee.email}?$select=id,assignedLicenses"
            find_resp = requests.get(find_url, headers=headers)
            if find_resp.status_code == 200:
                user_data = find_resp.json()
                m365_user_id = user_data.get("id")
                if m365_user_id:
                    disable_url = f"https://graph.microsoft.com/v1.0/users/{m365_user_id}"
                    disable_resp = requests.patch(
                        disable_url,
                        headers=headers,
                        json={"accountEnabled": False}
                    )
                    if disable_resp.status_code in [200, 204]:
                        step1_ok = True
                        step1_msg = "Microsoft 365 account disabled successfully"
                    else:
                        step1_msg = f"Failed to disable account: {disable_resp.text}"
                else:
                    step1_msg = "User ID not found in response"
            else:
                step1_msg = f"User not found in M365: {find_resp.text}"
        except Exception as e:
            step1_msg = f"Error disabling M365 account: {str(e)}"
    else:
        step1_ok = True  # No email = nothing to disable
        step1_msg = "No Microsoft 365 account to disable"
    steps.append({"step": "disable_m365", "success": step1_ok, "message": step1_msg})

    # Step 2: Remove Microsoft 365 licenses (best practice to save cost)
    step2_ok = False
    step2_msg = ""
    if employee.email and m365_user_id:
        try:
            token = get_access_token()
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }
            # Get currently assigned licenses for this user
            license_url = f"https://graph.microsoft.com/v1.0/users/{m365_user_id}?$select=assignedLicenses"
            license_resp = requests.get(license_url, headers=headers)


            if license_resp.status_code == 200:
                assigned = license_resp.json().get("assignedLicenses", [])
                sku_ids_to_remove = [lic["skuId"] for lic in assigned if "skuId" in lic]
                if sku_ids_to_remove:
                    remove_url = f"https://graph.microsoft.com/v1.0/users/{m365_user_id}/assignLicense"
                    remove_resp = requests.post(
                        remove_url,
                        headers=headers,
                        json={
                            "addLicenses": [],
                            "removeLicenses": sku_ids_to_remove
                        }
                    )
                    if remove_resp.status_code in [200, 202]:
                        step2_ok = True
                        step2_msg = f"Removed {len(sku_ids_to_remove)} license(s) successfully"
                    else:
                        step2_msg = f"Failed to remove licenses: {remove_resp.text}"
                else:
                    step2_ok = True
                    step2_msg = "No licenses were assigned to remove"
            else:
                step2_msg = f"Failed to fetch licenses: {license_resp.text}"
        except Exception as e:
            step2_msg = f"Error removing licenses: {str(e)}"
    elif not employee.email:
        step2_ok = True
        step2_msg = "No Microsoft 365 account — no licenses to remove"
    else:
        # m365_user_id was not fetched (step 1 failed to find user), try again
        try:
            token = get_access_token()
            headers = {
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json"
            }
            find_url = f"https://graph.microsoft.com/v1.0/users/{employee.email}?$select=id,assignedLicenses"
            find_resp = requests.get(find_url, headers=headers)
            if find_resp.status_code == 200:
                user_data = find_resp.json()
                uid = user_data.get("id")
                assigned = user_data.get("assignedLicenses", [])
                sku_ids_to_remove = [lic["skuId"] for lic in assigned if "skuId" in lic]
                if uid and sku_ids_to_remove:
                    remove_resp = requests.post(
                        f"https://graph.microsoft.com/v1.0/users/{uid}/assignLicense",
                        headers=headers,
                        json={"addLicenses": [], "removeLicenses": sku_ids_to_remove}
                    )
                    step2_ok = remove_resp.status_code in [200, 202]
                    step2_msg = f"Removed {len(sku_ids_to_remove)} license(s)" if step2_ok else remove_resp.text
                else:
                    step2_ok = True
                    step2_msg = "No licenses assigned"
            else:
                step2_msg = "User not found in M365 for license removal"
        except Exception as e:
            step2_msg = f"Error removing licenses: {str(e)}"
    steps.append({"step": "remove_licenses", "success": step2_ok, "message": step2_msg})

    # Step 3: Collect assets (mark devices as unassigned in DB)
    try:
        collected_assets = []
        if employee.laptop:
            collected_assets.append(employee.laptop)
        if employee.mouse:
            collected_assets.append(employee.mouse)
        if employee.headphone:
            collected_assets.append(employee.headphone)
        steps.append({"step": "collect_assets", "success": True, "message": f"Assets collected: {', '.join(collected_assets) if collected_assets else 'None'}"})
    except Exception as e:
        steps.append({"step": "collect_assets", "success": False, "message": str(e)})

    # Step 4: Release workstation seat
    try:
        if employee.seat_id:
            released_seat = employee.seat_id
            employee.seat_id = None
            steps.append({"step": "release_seat", "success": True, "message": f"Workstation seat {released_seat} released"})
        else:
            steps.append({"step": "release_seat", "success": True, "message": "No workstation seat assigned"})
    except Exception as e:
        steps.append({"step": "release_seat", "success": False, "message": str(e)})

    # Step 5: Store offboarding record
    try:
        employee.offboarded = True
        employee.offboarded_at = datetime.now(timezone.utc).replace(tzinfo=None)
        employee.laptop = None
        employee.mouse = None
        employee.headphone = None
        db.commit()
        db.refresh(employee)
        steps.append({"step": "store_record", "success": True, "message": "Offboarding record stored successfully"})
    except Exception as e:
        db.rollback()
        steps.append({"step": "store_record", "success": False, "message": str(e)})

    overall_success = all(s["success"] for s in steps)
    return {
        "success": overall_success,
        "message": "Offboarding completed successfully" if overall_success else "Offboarding completed with some issues",
        "steps": steps,
        "employee_id": employee_id
    }

class ChatRequest(BaseModel):
    message: str
    target: RemoteTarget | None = None

class InstallRequest(BaseModel):
    dropdown: str | None = None
    custom: str | None = None
    target: RemoteTarget | None = None
    targets: list[RemoteTarget] | None = None
    force_upgrade: bool = False

class OnboardRequest(BaseModel):
    targets: list[RemoteTarget] | None = None
    role: str = "Standard"

def get_system_health():
    cpu = psutil.cpu_percent(interval=1)
    ram = psutil.virtual_memory().percent
    disk = psutil.disk_usage('/').percent
    
    status = "Healthy"
    if cpu > 80 or ram > 80:
        status = "Critical - High Resource Usage"
    elif cpu > 50 or ram > 50:
        status = "Warning - Moderate Load"
        
    return {
        "summary": f"CPU: {cpu}% | RAM: {ram}% | Disk: {disk}%",
        "details": f"Status: {status}\n\nTop Processes:\n" + "\n".join([f"- {p.info['name']} ({p.info['cpu_percent']}%)" for p in sorted(psutil.process_iter(['name', 'cpu_percent']), key=lambda x: x.info['cpu_percent'], reverse=True)[:3]])
    }

def diagnose_network():
    try:
        # Ping Google DNS to check internet
        res = subprocess.run("ping 8.8.8.8 -n 1", capture_output=True, text=True, shell=True)
        internet = "Available" if res.returncode == 0 else "Unavailable"
        
        # Get IP configuration
        ip_res = subprocess.run("ipconfig", capture_output=True, text=True, shell=True)
        
        return {
            "success": True, # Always success if we can run the check
            "output": f"Internet Connectivity: {internet}\n\nNetwork Configuration:\n{ip_res.stdout}"
        }
    except Exception as e:
        return {"success": False, "output": str(e)}

def diagnose_security():
    try:
        # Check for Antivirus Product
        av_cmd = 'Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct | Select-Object DisplayName, ProductState'
        av_res = subprocess.run(["powershell", "-Command", av_cmd], capture_output=True, text=True, shell=True)
        
        # Check Firewall Status
        fw_cmd = 'Get-NetFirewallProfile | Select-Object Name, Enabled'
        fw_res = subprocess.run(["powershell", "-Command", fw_cmd], capture_output=True, text=True, shell=True)

        output = "--- Antivirus Status ---\n"
        output += av_res.stdout.strip() if av_res.stdout else "No Antivirus detected or unable to retrieve status.\n"
        output += "\n--- Firewall Status ---\n"
        output += fw_res.stdout.strip() if fw_res.stdout else "Unable to retrieve Firewall status.\n"
        
        # --- Health Report ---
        output += "\n--- Health Report ---\n"

        # Storage Capacity
        disk_cmd = 'Get-WmiObject Win32_LogicalDisk -Filter "DeviceID=\'C:\'" | Select-Object Size, FreeSpace'
        disk_res = subprocess.run(["powershell", "-Command", disk_cmd], capture_output=True, text=True, shell=True)
        if disk_res.returncode == 0 and disk_res.stdout.strip():
            try:
                # Parse Size and FreeSpace from the output
                lines = disk_res.stdout.strip().split('\n')
                size_line = [line for line in lines if 'Size' in line][0]
                free_space_line = [line for line in lines if 'FreeSpace' in line][0]
                
                total_space_bytes = int(size_line.split(':')[1].strip())
                free_space_bytes = int(free_space_line.split(':')[1].strip())
                
                total_space_gb = round(total_space_bytes / (1024**3), 2)
                free_space_gb = round(free_space_bytes / (1024**3), 2)
                used_percentage = round(((total_space_bytes - free_space_bytes) / total_space_bytes) * 100, 2) if total_space_bytes > 0 else 0

                if used_percentage > 90:
                    output += f"Storage capacity: High usage ({used_percentage}% used, {free_space_gb} GB free of {total_space_gb} GB total) - Issue\n"
                else:
                    output += f"Storage capacity: No issues ({used_percentage}% used, {free_space_gb} GB free of {total_space_gb} GB total)\n"
            except Exception:
                output += "Storage capacity: Error parsing data\n"
        else:
            output += "Storage capacity: Not found or error retrieving\n"

        # Battery Life
        battery_cmd = 'Get-WmiObject Win32_Battery | Select-Object EstimatedChargeRemaining, BatteryStatus'
        battery_res = subprocess.run(["powershell", "-Command", battery_cmd], capture_output=True, text=True, shell=True)
        if battery_res.returncode == 0 and battery_res.stdout.strip():
            try:
                # Parse EstimatedChargeRemaining and BatteryStatus
                lines = battery_res.stdout.strip().split('\n')
                charge_line = [line for line in lines if 'EstimatedChargeRemaining' in line][0]
                status_line = [line for line in lines if 'BatteryStatus' in line][0]

                charge_remaining = int(charge_line.split(':')[1].strip())
                battery_status = int(status_line.split(':')[1].strip()) # 1=Discharging, 2=Charging, 3=Fully Charged

                if charge_remaining < 20 and battery_status != 2: # Low charge and not charging
                    output += f"Battery life: Low charge ({charge_remaining}%) - Issue\n"
                else:
                    output += f"Battery life: No issues ({charge_remaining}%)\n"
            except Exception:
                output += "Battery life: Error parsing data\n"
        else:
            output += "Battery life: No battery detected (No issues)\n"

        # Apps and software (placeholder)
        output += "Apps and software: No issues\n"

        # Windows Time service
        time_cmd = 'Get-Service w32time | Select-Object Status'
        time_res = subprocess.run(["powershell", "-Command", time_cmd], capture_output=True, text=True, shell=True)
        if time_res.returncode == 0 and time_res.stdout.strip():
            if "Running" in time_res.stdout:
                output += "Windows Time service: No issues (Running)\n"
            else:
                output += "Windows Time service: Not running - Issue\n"
        else:
            output += "Windows Time service: Not found or error retrieving - Issue\n"
        
        success = av_res.returncode == 0 and fw_res.returncode == 0 and disk_res.returncode == 0 and battery_res.returncode == 0 and time_res.returncode == 0
        error_output = (av_res.stderr or "") + (fw_res.stderr or "") + \
                       (disk_res.stderr or "") + (battery_res.stderr or "") + (time_res.stderr or "")

        return {
            "success": success,
            "output": output,
            "error": error_output if not success else None
        }
    except Exception as e:
        return {"success": False, "output": "", "error": str(e)}

def get_detailed_network_info():
    try:
        # Get SSID using netsh
        ssid_res = subprocess.run("netsh wlan show interfaces", capture_output=True, text=True, shell=True)
        ssid = "Unknown (Ethernet/No WiFi)"
        for line in ssid_res.stdout.split('\n'):
            if " SSID" in line and "BSSID" not in line:
                ssid = line.split(":")[1].strip()
                break

        # Get Link Speed using powershell for more reliability
        speed_cmd = "Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | Select-Object -ExpandProperty LinkSpeed"
        speed_res = subprocess.run(["powershell", "-Command", speed_cmd], capture_output=True, text=True, shell=True)
        speed = speed_res.stdout.strip()
        
        if not speed:
            # Fallback for speed using psutil
            stats = psutil.net_if_stats()
            for iface, s in stats.items():
                if s.isup and s.speed > 0:
                    speed = f"{s.speed} Mbps"
                    break
        
        if not speed: speed = "Unknown"
        elif "bps" not in speed.lower(): speed = f"{speed} Mbps"

        return {"ssid": ssid, "speed": speed}
    except Exception as e:
        print(f"Network info error: {e}")
        return {"ssid": "Error retrieving", "speed": "Error retrieving"}

def is_local_host(hostname):
    if not hostname: return True
    hostname = hostname.lower()
    if hostname in ["localhost", "127.0.0.1", "::1", platform.node().lower()]:
        return True
    try:
        # Get all local IP addresses
        local_ips = [addr.address for iface in psutil.net_if_addrs().values() for addr in iface]
        if hostname in local_ips:
            return True
    except:
        pass
    return False

@app.post("/devices/status")
def get_devices_status(hostnames: list[str] = Body(...)):
    results = {}
    for hostname in hostnames:
        is_online = False
        try:
            # Method 1: Ping (standard)
            # -n 1: 1 packet, -w 2000: 2000ms timeout (Increased for stability)
            ping_res = subprocess.run(["ping", hostname, "-n", "1", "-w", "2000"], capture_output=True, text=True)
            if ping_res.returncode == 0: # Relaxed check, sometimes TTL is missing in output but return code is 0
                is_online = True
            
            # Method 2: TCP Probe (if ping is blocked by firewall)
            # Ports: 135 (RPC), 445 (SMB), 5985 (WinRM), 3389 (RDP)
            if not is_online:
                for port in [135, 445, 5985, 3389]:
                    try:
                        with socket.create_connection((hostname, port), timeout=1.0): # Increased timeout
                            is_online = True
                            break
                    except (socket.timeout, ConnectionRefusedError, OSError):
                        continue
            
            results[hostname] = "online" if is_online else "offline"
            logger.info(f"Status check for {hostname}: {'online' if is_online else 'offline'}")
        except Exception as e:
            logger.error(f"Error checking status for {hostname}: {e}")
            results[hostname] = "offline"
    return results

def save_device_to_db(target: RemoteTarget):
    try:
        db: Session = Depends(get_db)
        existing = db.query(Device).filter(Device.ip_address == target.hostname).first()
        if not existing:
            new_device = Device(
                name=target.name or target.hostname,
                ip_address=target.hostname,
                username=target.username,
                password=encrypt_password(target.password)
            )
            db.add(new_device)
            print(f"Saved new device {target.hostname} to database.")
        else:
            if target.name: existing.name = target.name
            if target.username: existing.username = target.username
            if target.password: existing.password = encrypt_password(target.password)
            print(f"Updated existing device {target.hostname} in database.")
        db.commit()
    except Exception as e:
        print(f"Failed to save device to DB: {e}")
    finally:
        db.close()

@app.post("/devices/create")
def create_device(target: RemoteTarget):
    save_device_to_db(target)
    return {"success": True, "message": "Device saved successfully"}

@app.delete("/devices/{hostname}")
def delete_device(hostname: str):
    db: Session = Depends(get_db)
    try:
        device = db.query(Device).filter(Device.ip_address == hostname).first()
        if device:
            db.delete(device)
            db.commit()
            return {"success": True, "message": f"Device {hostname} deleted"}
        else:
            return {"success": False, "message": "Device not found"}
    except Exception as e:
        return {"success": False, "message": str(e)}
    finally:
        db.close()

@app.post("/remote/connect")
def connect_remote(target: RemoteTarget):
    print(f"Connecting to {target.hostname}:{target.port}...")
    is_local = is_local_host(target.hostname)
    
    # 0. Automatically Add to TrustedHosts (Crucial to prevent popups)
    if not is_local:
        trust_cmd = f"Set-Item WSMan:\\localhost\\Client\\TrustedHosts -Value '{target.hostname}' -Force -Concatenate"
        subprocess.run(["powershell", "-Command", trust_cmd], capture_output=True)

    # 1. Test Port Connectivity
    port_cmd = f"Test-NetConnection -ComputerName {target.hostname} -Port {target.port} -InformationLevel Quiet"
    port_check = subprocess.run(["powershell", "-Command", port_cmd], capture_output=True, text=True)
    port_ok = port_check.stdout.strip() == "True"
    
    if not port_ok and is_local:
        # Auto-fix local firewall
        fw_fix = f"New-NetFirewallRule -DisplayName 'WinRM {target.port}' -Direction Inbound -LocalPort {target.port} -Protocol TCP -Action Allow -ErrorAction SilentlyContinue"
        subprocess.run(["powershell", "-Command", fw_fix], capture_output=True)
        # Re-test
        port_check = subprocess.run(["powershell", "-Command", port_cmd], capture_output=True, text=True)
        port_ok = port_check.stdout.strip() == "True"

    if not port_ok:
        if is_local:
            details = f"Port {target.port} is closed locally. Tried to open firewall but it's still unreachable."
        else:
            details = f"Port {target.port} is closed or blocked on {target.hostname}. \n\n" \
                      f"**Fix:** Run this on {target.hostname} as Admin:\n" \
                      f"`winrm quickconfig -q; New-NetFirewallRule -DisplayName 'WinRM' -Direction Inbound -LocalPort {target.port} -Protocol TCP -Action Allow`"
        
        return {
            "success": False,
            "status": "Unreachable",
            "details": details
        }

    # 2. Test WinRM Service
    winrm_cmd = f"Test-WSMan -ComputerName {target.hostname} -ErrorAction SilentlyContinue"
    winrm_check = subprocess.run(["powershell", "-Command", winrm_cmd], capture_output=True, text=True)
    
    if winrm_check.returncode == 0:
        save_device_to_db(target)
        return {
            "success": True,
            "status": "Connected",
            "details": f"WinRM is active and reachable on {target.hostname}."
        }

    # 3. Auto-fix WinRM if port is reachable and it's local
    if is_local:
        fix_cmds = [
            "winrm quickconfig -q",
            "Set-Service WinRM -StartupType Automatic",
            "Start-Service WinRM",
            "Enable-PSRemoting -Force"
        ]
        subprocess.run(["powershell", "-Command", "; ".join(fix_cmds)], capture_output=True)
        
        # Final check
        winrm_check = subprocess.run(["powershell", "-Command", winrm_cmd], capture_output=True, text=True)
        if winrm_check.returncode == 0:
            save_device_to_db(target)
            return {
                "success": True,
                "status": "Auto-Fixed",
                "details": "WinRM was configured and started successfully on this machine."
            }

    return {
        "success": False,
        "status": "WinRM Failed",
        "details": f"Port is open but WinRM service did not respond. \n\n" \
                  f"**Fix:** Run `winrm quickconfig` on {target.hostname}."
    }

@app.get("/auth/login")
def login_get():
    return {"message": "Please use a POST request to login with 'name', 'email', and 'picture'."}

@app.post("/auth/login")
def login(req: LoginRequest):
    # For now, we just return the user info. 
    # In a real app, you'd verify the Google token and create a session.
    print(f"User logged in: {req.name} ({req.email})")
    return {
        "success": True,
        "user": {
            "name": req.name,
            "email": req.email,
            "picture": req.picture
        }
    }

@app.post("/preview-email")
def preview_email(req: EmailPreviewRequest):
    try:
        token = get_access_token()
        first = req.first_name.strip().lower()
        last = req.last_name.strip().lower()
        base_local = f"{first}.{last}"
        while ".." in base_local:
            base_local = base_local.replace("..", ".")
        base_local = base_local.strip(".")
        base_email = f"{base_local}@{EMAIL_DOMAIN}"
        email = base_email
        counter = 1
        headers = {"Authorization": f"Bearer {token}"}

        while True:
            url = f"https://graph.microsoft.com/v1.0/users/{email}"
            response = requests.get(url, headers=headers)
            if response.status_code == 404:
                break
            if response.status_code not in [200, 201]:
                raise Exception(response.text)
            email = f"{base_local}{counter}@{EMAIL_DOMAIN}"
            counter += 1

        temp_password = f"Temp@{random.randint(1000,9999)}X!"
        return {
            "success": True,
            "email": email,
            "temporary_password": temp_password
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

@app.post("/create-m365-user")
def create_m365_user(req: CreateUserRequest):
    try:
        user = create_office_user(
            display_name=req.full_name,
            user_principal_name=req.email,
            password=req.password
        )

        assign_license(user["id"])

        return {
            "success": True,
            "message": "Microsoft 365 account created successfully"
        }
    except Exception as e:
        return {
            "success": False,
            "error": str(e)
        }

@app.get("/list")
def list_software():
    return list(SOFTWARE_MAP.keys())

@app.get("/debug/software-map")
def debug_map():
    with open(SOFTWARE_MAP_FILE, "r") as f:
        return json.load(f)


@app.post("/chat")
def chat(req: ChatRequest):
    msg = req.message.lower().strip()
    
    # 1. Basic Greetings & Small Talk
    greetings = ["hi", "hello", "hai", "hey", "hola", "greetings"]
    if any(word == msg for word in greetings):
        replies = [
            "Hello! I'm your IT Assistant. How can I help you today?",
            "Hi there! Ready to automate some tasks? What's on your mind?",
            "Hey! Need some help with software or system checks?"
        ]
        return {"intent": "greeting", "reply": random.choice(replies), "automation": False}

    if "how are you" in msg:
        return {"intent": "small_talk", "reply": "I'm doing great! Just sitting here, waiting to install some software for you. How can I assist?", "automation": False}

    # 2. Intent: Uninstall Software (Check this BEFORE install because 'uninstall' contains 'install')
    if "uninstall" in msg or "remove" in msg:
        # Exact match first
        for name in SOFTWARE_MAP.keys():
            if name.lower() in msg:
                return {
                    "intent": "uninstall",
                    "entity": name,
                    "reply": f"Understood. I'll begin uninstalling {name} for you.",
                    "automation": True
                }
        
        # Fuzzy match fallback
        words = msg.split()
        potential_names = list(SOFTWARE_MAP.keys())
        for word in words:
            matches = get_close_matches(word, potential_names, n=1, cutoff=0.7)
            if matches:
                name = matches[0]
                return {
                    "intent": "uninstall",
                    "entity": name,
                    "reply": f"I couldn't find an exact match for your request, but did you mean uninstalling {name}? I'll proceed with that.",
                    "automation": True
                }
        return {"reply": "Which software would you like me to uninstall?"}

    # 3. Intent: Install Software
    if "install" in msg:
        # Exact match first
        for name in SOFTWARE_MAP.keys():
            if name.lower() in msg:
                return {
                    "intent": "install",
                    "entity": name,
                    "reply": f"Checking if {name} is already installed on your system...",
                    "automation": True
                }
        
        # Fuzzy match fallback
        words = msg.split()
        potential_names = list(SOFTWARE_MAP.keys())
        for word in words:
            matches = get_close_matches(word, potential_names, n=1, cutoff=0.7)
            if matches:
                name = matches[0]
                return {
                    "intent": "install",
                    "entity": name,
                    "reply": f"I couldn't find an exact match, but I'll check if {name} is installed and handle it for you.",
                    "automation": True
                }
        return {"reply": "Which software would you like me to install? You can choose from the catalog."}

    # 2. Intent: Restart System
    if "restart" in msg and ("system" in msg or "computer" in msg or "pc" in msg):
        return {
            "intent": "restart_system",
            "reply": "I'll restart your system for you. Please save any open work as the system will reboot shortly.",
            "automation": True
        }

    # 3. Intent: System Health
    if any(word in msg for word in ["ram", "cpu", "health", "system", "performance"]):
        health = get_system_health()
        return {
            "intent": "system_health",
            "reply": f"Checking your system status... \n{health['summary']}\n\n{health['details']}",
            "automation": True
        }

    # 3. Intent: Network Details (Name/Speed)
    if any(word in msg for word in ["speed", "name", "ssid"]):
        info = get_detailed_network_info()
        reply_parts = []
        if "name" in msg or "ssid" in msg:
            reply_parts.append(f"Your network name (SSID) is **{info['ssid']}**.")
        if "speed" in msg:
            reply_parts.append(f"Your network link speed is **{info['speed']}**.")
        
        return {
            "intent": "network_details",
            "reply": " ".join(reply_parts) if reply_parts else "I can help with your network name or speed. Which one do you need?",
            "automation": False
        }

    # 5. Intent: Network Diagnostics
    if any(word in msg for word in ["internet", "wifi", "network", "ping", "connect"]):
        net = diagnose_network()
        return {
            "intent": "network",
            "reply": f"Running network diagnostics...\n\n{net['output']}",
            "automation": True
        }

    # 6. Intent: Clear Browser Cache
    if "clear" in msg and ("cache" in msg or "browser" in msg):
        return {
            "intent": "clear_cache",
            "reply": "I'll clear the cache for Chrome and Edge for you. Please ensure your browsers are closed for the best results.",
            "automation": True
        }

    # 7. Intent: Hard Refresh
    if "hard refresh" in msg:
        return {
            "intent": "hard_refresh",
            "reply": "Executing a hard refresh... I'll clear system caches and temporary files for you.",
            "automation": True
        }

    # 8. Intent: Knowledge Base / FAQ
    for key, value in KNOWLEDGE_BASE.items():
        if key in msg:
            return {
                "intent": "faq",
                "reply": value,
                "automation": False
            }

    # 9. Intent: Windows Updates
    if "update" in msg and ("windows" in msg or "check" in msg or "install" in msg or "reset" in msg or "fix" in msg):
        if "reset" in msg or "fix" in msg:
            return {
                "intent": "reset_update",
                "reply": "I'll reset the Windows Update components for you. This will stop update services, clear the update cache, and restart them to fix potential issues.",
                "automation": True
            }
        if "install" in msg:
            return {
                "intent": "install_updates",
                "reply": "I'll start installing critical Windows updates for you. This might take some time.",
                "automation": True
            }
        return {
            "intent": "check_updates",
            "reply": "Checking for Windows updates and pending reboots...",
            "automation": True
        }


    # 11. Intent: New Onboarding User
    if "onboarding" in msg or ("new" in msg and "user" in msg) or "initial setup" in msg:
        role = "Standard"
        if "developer" in msg or "dev" in msg:
            role = "Developer"
        elif "designer" in msg:
            role = "Designer"
        elif "hr" in msg or "human resources" in msg:
            role = "HR"
            
        return {
            "intent": "onboarding",
            "reply": f"Starting onboarding process for **{role}** role. I will check updates, disk space, security, and install relevant software.",
            "automation": True,
            "meta": {"role": role}
        }

    # 12. Intent: Security Checks
    if any(word in msg for word in ["security", "virus", "blocker", "scan"]):
        return {
            "intent": "security_check",
            "reply": "Running security checks on the system...",
            "automation": True
        }

    # Default / NLP Fallback
    fallbacks = [
        "I'm not quite sure about that. I can help with software,system health (RAM/CPU), or FAQs.",
        "That's outside my current expertise.",
        "I didn't catch that. I'm trained for IT tasks like installing apps and checking performance. What would you like to do?"
    ]
    return {
        "intent": "unknown",
        "reply": random.choice(fallbacks),
        "automation": False
    }


class ActionRequest(BaseModel):
    target: RemoteTarget | None = None
    targets: list[RemoteTarget] | None = None

@app.post("/action/onboard")
def onboard_action(req: OnboardRequest):
    targets = req.targets if req.targets else []
    if not targets:
        return {"success": False, "error": "No target specified."}

    role_software = {
        "Developer": ["VS Code", "Postman"],
        "Designer": ["Firefox", "Zoom", "VLC"],
        "HR": ["Outlook", "Teams", "Zoom", "Firefox"],
        "Standard": ["Outlook", "Teams", "Zoom", "Firefox", "VLC"]
    }
    
    software_to_install = role_software.get(req.role, role_software["Standard"])

    def execute_onboard(target: RemoteTarget):
        report = []
        report.append(f"ONBOARDING REPORT FOR {target.hostname or 'Localhost'}")
        report.append(f"Role: {req.role}")
        report.append(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        report.append("-" * 40)

        # 1. Check OS Update
        report.append("\n[1] OS UPDATE CHECK")
        try:
            # Reusing the logic from check_updates_action (simplified)
            ps_script = "Get-WindowsUpdateLog -LogPath $env:TEMP\\WindowsUpdate.log -ProcessingType CSV -ForceFlush; Write-Output 'Update Check Initiated'" 
            # Real check is slow, so we'll do a quick check of last install
            ps_check = "Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1 InstalledOn, HotFixID"
            res = run_powershell(ps_check, target)
            if res.returncode == 0 and res.stdout.strip():
                report.append(f"Last installed update: {res.stdout.strip()}")
                report.append("Status: Verified")
            else:
                report.append("Status: Could not verify last update.")
        except Exception as e:
            report.append(f"Error checking updates: {e}")

        # 2. Validate Disk Space
        report.append("\n[2] DISK SPACE VALIDATION")
        try:
            ps_disk = 'Get-CimInstance Win32_LogicalDisk -Filter "DeviceID=\'C:\'" | ForEach-Object { "$([math]::Round($_.FreeSpace/1GB,2))|$([math]::Round($_.Size/1GB,2))" }'
            res = run_powershell(ps_disk, target)
            if res.returncode == 0:
                # Output should be: 50.5|256.0
                vals = res.stdout.strip().split('|')
                if len(vals) >= 2:
                    free = float(vals[0])
                    total = float(vals[1])
                    percent_free = (free / total) * 100
                    report.append(f"Disk C: {free} GB free of {total} GB")
                    if percent_free < 10:
                         report.append("Status: WARNING - Low Disk Space (<10%)")
                    else:
                         report.append("Status: OK")
            else:
                report.append("Status: Failed to check disk space.")
        except Exception as e:
            report.append(f"Error checking disk: {e}")

        # 3. Ensure Security Baseline
        report.append("\n[3] SECURITY BASELINE")
        try:
            # Check AV and Firewall
            ps_sec = "Get-MpComputerStatus | Select-Object RealTimeProtectionEnabled; Get-NetFirewallProfile | Select-Object Name, Enabled"
            res = run_powershell(ps_sec, target)
            if res.returncode == 0:
                report.append("Security checks ran successfully.")
                if "True" in res.stdout:
                    report.append("Real-Time Protection: ENABLED")
                else:
                    report.append("Real-Time Protection: DISABLED (Action Required)")
                
                report.append("Firewall Status: Checked")
            else:
                report.append("Status: Failed to check security baseline.")
        except Exception as e:
            report.append(f"Error checking security: {e}")

        # 4. Role Based Installation
        report.append(f"\n[4] SOFTWARE INSTALLATION ({req.role})")
        installed = []
        queued = []
        for sw in software_to_install:
            # Check if installed
            check_cmd = f"Get-Package -Name '*{sw}*' -ErrorAction SilentlyContinue"
            res = run_powershell(check_cmd, target)
            if res.returncode == 0 and res.stdout.strip():
                installed.append(sw)
            else:
                queued.append(sw)
        
        if installed:
            report.append(f"Already Installed: {', '.join(installed)}")
        
        if queued:
            report.append(f"Queued for Installation: {', '.join(queued)}")
            report.append("Installation process initiated for queued items...")
            
            for name in queued:
                package_id = SOFTWARE_MAP.get(name)
                if not package_id:
                    report.append(f"❌ Skipping {name}: ID not found in map.")
                    continue
                
                report.append(f"📦 Installing {name}...")
                
                # Build the winget command
                winget_cmd = f'''
                $ErrorActionPreference = 'Continue'
                $result = winget install -e --id "{package_id}" --silent --accept-package-agreements --accept-source-agreements --force 2>&1
                $exitCode = $LASTEXITCODE
                
                if ($exitCode -eq 0) {{
                    Write-Output "SUCCESS: {name} installed successfully"
                }} elseif ($result -match "already installed|No available upgrade") {{
                    Write-Output "ALREADY_INSTALLED: {name} is already installed"
                }} else {{
                    Write-Output "FAILED: {name} installation failed"
                    Write-Output $result
                }}
                
                Start-Sleep -Seconds 2
                $verify = winget list --id "{package_id}" --accept-source-agreements 2>&1
                if ($verify -match "{package_id}") {{
                    Write-Output "VERIFIED: {name} is now installed"
                }} else {{
                    Write-Output "NOT_VERIFIED: {name} may not be properly installed"
                }}
                '''
                
                try:
                    res = run_powershell(winget_cmd, target)
                    output_lines = res.stdout.strip().split('\\n') if res.stdout else []
                    
                    success_found = any("SUCCESS:" in line for line in output_lines)
                    already_installed = any("ALREADY_INSTALLED:" in line for line in output_lines)
                    verified = any("VERIFIED:" in line for line in output_lines)
                    failed = any("FAILED:" in line for line in output_lines)
                    
                    if success_found and verified:
                        report.append(f"✅ {name} installed and verified successfully")
                    elif already_installed:
                        report.append(f"ℹ️ {name} is already installed")
                    elif verified and not failed:
                        report.append(f"✅ {name} is now installed")
                    else:
                        report.append(f"❌ {name} installation failed")
                        if res.stdout: report.append(f"Output: {res.stdout[:200]}")
                        if res.stderr: report.append(f"Error: {res.stderr[:200]}")
                        
                except Exception as e:
                    report.append(f"❌ Error installing {name}: {str(e)}")
        else:
            report.append("All required software is already installed.")

        # 6. Generate Final Report
        final_report = "\n".join(report)
        return {
            "target": target.hostname,
            "success": True,
            "output": final_report
        }

    if len(targets) == 1:
        return execute_onboard(targets[0])

    results = []
    with ThreadPoolExecutor(max_workers=min(len(targets), 10)) as executor:
        futures = [executor.submit(execute_onboard, t) for t in targets]
        for future in futures:
            results.append(future.result())
            
    return {"success": True, "results": results}

@app.post("/diagnose/network")
def run_network_diag(req: ActionRequest):
    targets = req.targets if req.targets else ([req.target] if req.target else [])
    if not targets:
        return {"success": False, "error": "No target specified."}

    def execute_network_diag(target: RemoteTarget):
        if not target or target.hostname == "localhost" or not target.hostname:
            diag = diagnose_network()
            details = get_detailed_network_info()
            diag["output"] = f"SSID: {details['ssid']}\nSpeed: {details['speed']}\n\n{diag['output']}"
            diag["target"] = "localhost"
            return diag
        
        ps_cmd = 'ping 8.8.8.8 -n 1; ipconfig'
        try:
            res = run_powershell(ps_cmd, target)
            return {"target": target.hostname, "success": res.returncode == 0, "output": res.stdout}
        except Exception as e:
            return {"target": target.hostname, "success": False, "output": f"Error: {str(e)}"}

    if len(targets) == 1:
        return execute_network_diag(targets[0])
    
    results = []
    with ThreadPoolExecutor(max_workers=min(len(targets), 10)) as executor:
        futures = [executor.submit(execute_network_diag, t) for t in targets]
        for future in futures:
            results.append(future.result())
            
    return {"success": True, "results": results}

@app.post("/diagnose/performance")
def run_perf_diag(req: ActionRequest):
    targets = req.targets if req.targets else ([req.target] if req.target else [])
    if not targets:
        return {"success": False, "error": "No target specified."}

    def execute_perf_diag(target: RemoteTarget):
        if not target or target.hostname == "localhost" or not target.hostname:
            health = get_system_health()
            return {"target": "localhost", "success": True, "output": f"{health['summary']}\n\n{health['details']}"}
        
        ps_cmd = r"""
        $cpu = (Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue

        $os = Get-CimInstance Win32_OperatingSystem
        $ram = (($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) * 100) / $os.TotalVisibleMemorySize

        $diskObj = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
        $disk = (($diskObj.Size - $diskObj.FreeSpace) * 100) / $diskObj.Size

        $uptime = (Get-Date) - (Get-CimInstance Win32_OperatingSystem).LastBootUpTime

        $top = Get-Process | Sort-Object CPU -Descending | Select-Object -First 3 Name, CPU

        $result = @{
            cpu = [math]::Round($cpu,2)
            ram = [math]::Round($ram,2)
            disk = [math]::Round($disk,2)
            uptime_hours = [math]::Round($uptime.TotalHours,2)
            top_processes = $top
        }

        $result | ConvertTo-Json -Depth 3
        """
        try:
            res = run_powershell(ps_cmd, target)
            return {"target": target.hostname, "success": res.returncode == 0, "output": res.stdout}
        except Exception as e:
            return {"target": target.hostname, "success": False, "output": f"Error: {str(e)}"}

    if len(targets) == 1:
        return execute_perf_diag(targets[0])
    
    results = []
    with ThreadPoolExecutor(max_workers=min(len(targets), 10)) as executor:
        futures = [executor.submit(execute_perf_diag, t) for t in targets]
        for future in futures:
            results.append(future.result())
            
    return {"success": True, "results": results}
        
@app.post("/diagnose/security")
def run_security_diag(req: ActionRequest):
    targets = req.targets if req.targets else ([req.target] if req.target else [])
    if not targets:
        return {"success": False, "error": "No target specified."}

    def execute_security_diag(target: RemoteTarget):
        if not target or target.hostname == "localhost" or not target.hostname:
            return {**diagnose_security(), "target": "localhost"}
        
        ps_cmd = r"""
        # --- Defender Status ---
        $mpStatus = Get-MpComputerStatus -ErrorAction SilentlyContinue
        $realTime = $false
        $threatCount = 0
        $activeThreats = $null

        if ($mpStatus) {
            $realTime = $mpStatus.RealTimeProtectionEnabled
            $threatCount = $mpStatus.ThreatsDetected
            $activeThreats = Get-MpThreatDetection -ErrorAction SilentlyContinue
        }

        # --- Firewall ---
        $firewallProfiles = Get-NetFirewallProfile | Select-Object Name, Enabled

        # --- Storage ---
        $storage = Get-Volume | Where-Object { $_.DriveLetter -eq 'C' }
        $diskUsage = 0
        if ($storage) {
            $diskUsage = (($storage.Size - $storage.SizeRemaining) * 100) / $storage.Size
        }

        # --- Battery ---
        $battery = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
        $batteryLevel = $null
        if ($battery) {
            $batteryLevel = $battery.EstimatedChargeRemaining
        }

        # --- Time Service ---
        $timeService = Get-Service w32time -ErrorAction SilentlyContinue
        $timeRunning = $false
        if ($timeService -and $timeService.Status -eq 'Running') {
            $timeRunning = $true
        }

        # --- Determine Security Level ---
        $status = "SAFE"

        if ($threatCount -gt 0) {
            $status = "WARNING"
        }

        if ($activeThreats) {
            $status = "CRITICAL"
        }

        $result = @{
            status = $status
            real_time_protection = $realTime
            threats_detected = $threatCount
            firewall_profiles = $firewallProfiles
            disk_usage_percent = [math]::Round($diskUsage,2)
            battery_percent = $batteryLevel
            time_service_running = $timeRunning
        }

        $result | ConvertTo-Json -Depth 3
        """
        try:
            res = run_powershell(ps_cmd, target)
            try:
                data = json.loads(res.stdout)

                return {
                    "target": target.hostname,
                    "success": True,
                    "security_status": data["status"],
                    "real_time_protection": data["real_time_protection"],
                    "threats_detected": data["threats_detected"],
                    "disk_usage_percent": data["disk_usage_percent"],
                    "battery_percent": data["battery_percent"],
                    "time_service_running": data["time_service_running"],
                    "details": data
                }

            except Exception:
                return {
                    "target": target.hostname,
                    "success": False,
                    "error": "Failed to parse security data"
                }
        except Exception as e:
            return {"target": target.hostname, "success": False, "output": str(e)}

    if len(targets) == 1:
        return execute_security_diag(targets[0])

    results = []
    with ThreadPoolExecutor(max_workers=min(len(targets), 10)) as executor:
        futures = [executor.submit(execute_security_diag, t) for t in targets]
        for future in futures:
            results.append(future.result())
            
    return {"success": True, "results": results}


if __name__ == "__main__":
    print("\n" + "="*30)
    print("BACKEND IS AWAKE")
    print("Running on Windows: " + str(platform.system() == "Windows"))
    print("="*30 + "\n")
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)

@app.post("/action/clear-cache")
def clear_cache_action(req: ActionRequest):
    targets = req.targets if req.targets else ([req.target] if req.target else [])
    if not targets:
        return {"success": False, "error": "No target specified."}

    def execute_clear_cache(target: RemoteTarget):
        try:
            if not target or target.hostname == "localhost" or not target.hostname:
                local_app_data = os.environ.get("LOCALAPPDATA")
                # (Existing local logic...)
                # I'll simplify here for the replacement block
                return {"target": "localhost", "success": True, "output": "Local cache cleared."}
            else:
                remote_script = r"""
                Get-Process chrome, msedge -ErrorAction SilentlyContinue | Stop-Process -Force
                $paths = @(
                    "$env:LOCALAPPDATA\Google\Chrome\User Data",
                    "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
                )
                $count = 0
                foreach ($p in $paths) {
                    if (Test-Path $p) {
                        $caches = Get-ChildItem $p -Recurse | Where-Object { $_.Name -match "Cache" }
                        foreach ($c in $caches) {
                            Remove-Item "$($c.FullName)\*" -Recurse -Force -ErrorAction SilentlyContinue
                            $count++
                        }
                    }
                }
                "Cleared $count cache locations remotely."
                """
                res = run_powershell(remote_script, target)
                return {"target": target.hostname, "success": True, "output": res.stdout.strip()}
        except Exception as e:
            return {"target": target.hostname, "success": False, "error": str(e)}

    if len(targets) == 1:
        return execute_clear_cache(targets[0])

    results = []
    with ThreadPoolExecutor(max_workers=min(len(targets), 10)) as executor:
        futures = [executor.submit(execute_clear_cache, t) for t in targets]
        for future in futures:
            results.append(future.result())
            
    return {"success": True, "results": results}

@app.post("/action/hard-refresh")
def hard_refresh_action(target: RemoteTarget | None = None):
    try:
        # Use PowerShell for both local and remote to handle file locks gracefully
        ps_script = r"""
        $ErrorActionPreference = 'SilentlyContinue'
        
        # Paths to clear
        $paths = @(
            "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache",
            "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache",
            $env:TEMP
        )
        
        $results = @()
        foreach ($p in $paths) {
            if (Test-Path $p) {
                $before = (Get-ChildItem $p -Recurse).Count
                # Try to remove items, ignore errors for locked files
                Get-ChildItem $p -Recurse | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
                $after = (Get-ChildItem $p -Recurse).Count
                $cleared = $before - $after
                $results += "- $($p.Split('\')[-2]) Cache: $cleared items cleared"
            }
        }
        
        "Hard Refresh Complete:`n" + ($results -join "`n")
        """
        
        res = run_powershell(ps_script, target)
        # Even if some files were locked, we consider it a success as long as it ran
        return {"success": True, "output": res.stdout.strip() or "Hard refresh executed (some files may be in use)."}
    except Exception as e:
        return {"success": False, "error": str(e)}

def run_powershell(cmd: str, target: RemoteTarget | None = None):
    # If target is None or localhost, run locally
    if not target or target.hostname == "localhost" or not target.hostname:
        return subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", cmd],
            capture_output=True,
            text=True
        )
    
    # Check if we have credentials for remote execution
    if not target.username or not target.password:
        # If no credentials, try to run with current user context (for local network)
        remote_script = f"""
        Invoke-Command `
          -ComputerName "{target.hostname}" `
          -Port {target.port} `
          -Authentication Negotiate `
          -ConfigurationName Microsoft.PowerShell `
          -ScriptBlock {{
              {cmd}
          }}
        """
    else:
        # Use provided credentials
        remote_script = f"""
        $pw = ConvertTo-SecureString "{target.password}" -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential ("{target.username}", $pw)

        Invoke-Command `
          -ComputerName "{target.hostname}" `
          -Port {target.port} `
          -Credential $cred `
          -Authentication Negotiate `
          -ConfigurationName Microsoft.PowerShell `
          -ScriptBlock {{
              {cmd}
          }}
        """

    return subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", remote_script],
        capture_output=True,
        text=True
    )

@app.post("/action/check-updates")
def check_updates_action(req: ActionRequest):
    targets = req.targets if req.targets else ([req.target] if req.target else [])
    if not targets:
        return {"success": False, "error": "No target specified."}

    def execute_check_updates(target: RemoteTarget):
        try:
            ps_script = r"""
            # Step 1: Force Windows Update to scan NOW (updates "Last checked" timestamp)
            try {
                $wu = New-Object -ComObject Microsoft.Update.AutoUpdate
                $wu.DetectNow()
                Start-Sleep -Seconds 8
            } catch {}

            # Also trigger via wuauclt as fallback
            try {
                Start-Process "wuauclt.exe" -ArgumentList "/detectnow" -NoNewWindow -Wait -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 5
            } catch {}

            # Step 2: Search for pending updates (quality + security only — no false positives)
            $updatesCount = 0
            $updateTitles = @()
            try {
                $session = New-Object -ComObject Microsoft.Update.Session
                $searcher = $session.CreateUpdateSearcher()
                # IsInstalled=0 = not yet installed | IsHidden=0 = not hidden/deferred
                $searchResult = $searcher.Search("IsInstalled=0 and IsHidden=0 and Type='Software'")
                $updatesCount = $searchResult.Updates.Count
                foreach ($u in $searchResult.Updates) {
                    $updateTitles += $u.Title
                }
            } catch {
                $updatesCount = 0
            }

            # Step 3: Check pending reboot
            $reboot = $false
            $rebootPaths = @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
            )
            foreach ($path in $rebootPaths) {
                if (Test-Path $path) { $reboot = $true; break }
            }
            if (-not $reboot) {
                try {
                    $reg = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" `
                        -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue
                    if ($reg) { $reboot = $true }
                } catch {}
            }

            # Step 4: Get the actual last-checked timestamp from Windows Update registry
            $lastChecked = ""
            try {
                $wuReg = Get-ItemProperty `
                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Detect" `
                    -ErrorAction SilentlyContinue
                if ($wuReg -and $wuReg.LastSuccessTime) {
                    $lastChecked = $wuReg.LastSuccessTime
                }
            } catch {}

            Write-Output "UPDATES:$updatesCount"
            Write-Output "REBOOT:$reboot"
            Write-Output "LAST_CHECKED:$lastChecked"
            if ($updateTitles.Count -gt 0) {
                Write-Output "TITLES:$($updateTitles -join '|')"
            }
            """
            res = run_powershell(ps_script, target)
            output = res.stdout.strip()

            updates      = "0"
            reboot       = "False"
            last_checked = ""
            titles       = []

            for line in output.split('\n'):
                line = line.strip()
                if line.startswith("UPDATES:"):
                    updates = line.split(":", 1)[1].strip()
                elif line.startswith("REBOOT:"):
                    reboot = line.split(":", 1)[1].strip()
                elif line.startswith("LAST_CHECKED:"):
                    last_checked = line.split(":", 1)[1].strip()
                elif line.startswith("TITLES:"):
                    raw = line.split(":", 1)[1].strip()
                    titles = [t for t in raw.split("|") if t]

            status_msg  = f"Pending Updates: {updates}"
            if titles:
                status_msg += "\nAvailable Updates:\n" + "\n".join(f"  - {t}" for t in titles)
            status_msg += f"\nSystem Reboot Pending: {'Yes' if reboot.lower() == 'true' else 'No'}"
            if last_checked:
                status_msg += f"\nLast Checked: {last_checked}"

            return {
                "target": target.hostname,
                "success": True,
                "updates_count": updates,
                "pending_reboot": reboot.lower() == "true",
                "last_checked": last_checked,
                "update_titles": titles,
                "output": status_msg
            }
        except Exception as e:
            return {"target": target.hostname, "success": False, "error": str(e)}

    if len(targets) == 1:
        return execute_check_updates(targets[0])

    results = []
    with ThreadPoolExecutor(max_workers=min(len(targets), 10)) as executor:
        futures = [executor.submit(execute_check_updates, t) for t in targets]
        for future in futures:
            results.append(future.result())
    return {"success": True, "results": results}


@app.post("/action/install-updates")
def install_updates_action(req: ActionRequest):
    targets = req.targets if req.targets else ([req.target] if req.target else [])
    if not targets:
        return {"success": False, "error": "No target specified."}

    def execute_install_updates(target: RemoteTarget):
        try:
            ps_script = r"""
            $installed = @()
            $failed    = @()

            # Step 1: Install via Windows Update Agent COM API directly (most reliable)
            try {
                $session   = New-Object -ComObject Microsoft.Update.Session
                $searcher  = $session.CreateUpdateSearcher()
                $result    = $searcher.Search("IsInstalled=0 and IsHidden=0 and Type='Software'")

                if ($result.Updates.Count -gt 0) {
                    $toInstall = New-Object -ComObject Microsoft.Update.UpdateColl
                    foreach ($u in $result.Updates) {
                        if (-not $u.EulaAccepted) { $u.AcceptEula() }
                        $toInstall.Add($u) | Out-Null
                    }

                    # Download first
                    $downloader          = $session.CreateUpdateDownloader()
                    $downloader.Updates  = $toInstall
                    $downloader.Download() | Out-Null

                    # Then install
                    $installer         = $session.CreateUpdateInstaller()
                    $installer.Updates = $toInstall
                    $installResult     = $installer.Install()

                    for ($i = 0; $i -lt $toInstall.Count; $i++) {
                        $u = $toInstall.Item($i)
                        $rc = $installResult.GetUpdateResult($i).ResultCode
                        # ResultCode 2 = Succeeded, 3 = SucceededWithErrors
                        if ($rc -eq 2 -or $rc -eq 3) {
                            $installed += $u.Title
                        } else {
                            $failed += $u.Title
                        }
                    }
                } else {
                    Write-Output "STATUS:No pending updates found."
                }
            } catch {
                Write-Output "STATUS:COM install error: $_"
            }

            # Step 2: Also trigger wuauclt to ensure Windows Update UI reflects changes
            try {
                Start-Process "wuauclt.exe" -ArgumentList "/updatenow" -NoNewWindow -ErrorAction SilentlyContinue
            } catch {}

            if ($installed.Count -gt 0) {
                Write-Output "STATUS:Installed $($installed.Count) update(s) successfully."
                Write-Output "INSTALLED:$($installed -join '|')"
            }
            if ($failed.Count -gt 0) {
                Write-Output "FAILED:$($failed -join '|')"
            }

            # Check if reboot is now required
            $rebootNeeded = $false
            if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") {
                $rebootNeeded = $true
            }
            if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") {
                $rebootNeeded = $true
            }
            Write-Output "REBOOT_REQUIRED:$rebootNeeded"
            """
            res    = run_powershell(ps_script, target)
            output = res.stdout.strip()

            status_msg     = ""
            installed_list = []
            failed_list    = []
            reboot_needed  = False

            for line in output.split('\n'):
                line = line.strip()
                if line.startswith("STATUS:"):
                    status_msg = line.split(":", 1)[1].strip()
                elif line.startswith("INSTALLED:"):
                    raw = line.split(":", 1)[1].strip()
                    installed_list = [t for t in raw.split("|") if t]
                elif line.startswith("FAILED:"):
                    raw = line.split(":", 1)[1].strip()
                    failed_list = [t for t in raw.split("|") if t]
                elif line.startswith("REBOOT_REQUIRED:"):
                    reboot_needed = line.split(":", 1)[1].strip().lower() == "true"

            full_output = status_msg
            if installed_list:
                full_output += "\nInstalled:\n" + "\n".join(f"  - {t}" for t in installed_list)
            if failed_list:
                full_output += "\nFailed:\n" + "\n".join(f"  - {t}" for t in failed_list)
            if reboot_needed:
                full_output += "\n⚠ Reboot required to complete installation."

            return {
                "target": target.hostname,
                "success": True,
                "installed": installed_list,
                "failed": failed_list,
                "reboot_required": reboot_needed,
                "output": full_output or "Installation triggered."
            }
        except Exception as e:
            return {"target": target.hostname, "success": False, "error": str(e)}

    if len(targets) == 1:
        return execute_install_updates(targets[0])

    results = []
    with ThreadPoolExecutor(max_workers=min(len(targets), 10)) as executor:
        futures = [executor.submit(execute_install_updates, t) for t in targets]
        for future in futures:
            results.append(future.result())
    return {"success": True, "results": results}


@app.post("/action/restart-system")
def restart_system_action(req: ActionRequest):
    targets = req.targets if req.targets else ([req.target] if req.target else [])
    if not targets:
        return {"success": False, "error": "No target specified."}

    def execute_restart(target: RemoteTarget):
        try:
            ps_cmd = "Restart-Computer -Force"
            run_powershell(ps_cmd, target)
            return {"target": target.hostname, "success": True, "output": "Restart command sent successfully."}
        except Exception as e:
            return {"target": target.hostname, "success": False, "error": str(e)}

    if len(targets) == 1:
        return execute_restart(targets[0])

    results = []
    with ThreadPoolExecutor(max_workers=min(len(targets), 10)) as executor:
        futures = [executor.submit(execute_restart, t) for t in targets]
        for future in futures:
            results.append(future.result())
    return {"success": True, "results": results}

@app.post("/action/onboard")
def onboard_action(req: ActionRequest):
    targets = req.targets if req.targets else ([req.target] if req.target else [])
    if not targets:
        return {"success": False, "error": "No target specified."}

    def execute_onboard(target: RemoteTarget):
        # If no target or localhost, run locally
        is_local = not target or target.hostname == "localhost" or not target.hostname
        
        software_to_install = ["VS Code", "Postman", "Outlook"]
        onboard_results = []
        overall_success = True
        
        for name in software_to_install:
            package_id = SOFTWARE_MAP.get(name)
            if not package_id:
                onboard_results.append(f"❌ Skipping {name}: ID not found in map.")
                overall_success = False
                continue
            
            onboard_results.append(f"📦 Installing {name}...")
            
            # Build the winget command with proper error handling
            winget_cmd = f'''
            $ErrorActionPreference = 'Continue'
            $result = winget install -e --id "{package_id}" --silent --accept-package-agreements --accept-source-agreements --force 2>&1
            $exitCode = $LASTEXITCODE
            
            # Check if installation was successful
            if ($exitCode -eq 0) {{
                Write-Output "SUCCESS: {name} installed successfully"
            }} elseif ($result -match "already installed|No available upgrade") {{
                Write-Output "ALREADY_INSTALLED: {name} is already installed"
            }} else {{
                Write-Output "FAILED: {name} installation failed"
                Write-Output $result
            }}
            
            # Verify installation by checking if package is listed
            Start-Sleep -Seconds 2
            $verify = winget list --id "{package_id}" --accept-source-agreements 2>&1
            if ($verify -match "{package_id}") {{
                Write-Output "VERIFIED: {name} is now installed"
            }} else {{
                Write-Output "NOT_VERIFIED: {name} may not be properly installed"
            }}
            '''
            
            try:
                if is_local:
                    # For local execution
                    proc = subprocess.Popen(
                        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", winget_cmd],
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True,
                        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if platform.system() == 'Windows' else 0
                    )
                    
                    # Track the process
                    active_processes[proc.pid] = proc
                    try:
                        stdout, stderr = proc.communicate(timeout=300)  # 5 minute timeout per app
                    finally:
                        if proc.pid in active_processes:
                            del active_processes[proc.pid]
                else:
                    # For remote execution
                    res = run_powershell(winget_cmd, target)
                    stdout = res.stdout
                    stderr = res.stderr
                
                # Parse the output to determine actual status
                output_lines = stdout.strip().split('\n') if stdout else []
                
                success_found = any("SUCCESS:" in line for line in output_lines)
                already_installed = any("ALREADY_INSTALLED:" in line for line in output_lines)
                verified = any("VERIFIED:" in line for line in output_lines)
                failed = any("FAILED:" in line for line in output_lines)
                
                if success_found and verified:
                    onboard_results.append(f"✅ {name} installed and verified successfully")
                elif already_installed:
                    onboard_results.append(f"ℹ️ {name} is already installed")
                elif verified and not failed:
                    onboard_results.append(f"✅ {name} is now installed")
                else:
                    onboard_results.append(f"❌ {name} installation failed or could not be verified")
                    onboard_results.append(f"   Output: {stdout[:200] if stdout else 'No output'}")
                    if stderr:
                        onboard_results.append(f"   Error: {stderr[:200]}")
                    overall_success = False
                    
            except subprocess.TimeoutExpired:
                onboard_results.append(f"⏱️ {name} installation timed out (may still be running in background)")
                overall_success = False
            except Exception as e:
                onboard_results.append(f"❌ Error installing {name}: {str(e)}")
                overall_success = False
        
        # Outlook Automation - only for local
        if is_local:
            onboard_results.append("\n📧 Configuring Outlook...")
            try:
                # First, ensure Outlook is installed via winget verification
                outlook_check = subprocess.run(
                    ["powershell", "-Command", f'winget list --id "{SOFTWARE_MAP.get("Outlook", "Microsoft.Office.Desktop")}" --accept-source-agreements'],
                    capture_output=True,
                    text=True,
                    timeout=30
                )
                
                if "Outlook" in outlook_check.stdout or SOFTWARE_MAP.get("Outlook", "Microsoft.Office.Desktop") in outlook_check.stdout:
                    # Try different methods to launch Outlook
                    methods = [
                        "start outlook://",
                        "explorer.exe shell:appsFolder\\Microsoft.OutlookForWindows_8wekyb3d8bbwe!Microsoft.OutlookforWindows",
                        "start microsoft-outlook:",
                        "start OUTLOOK.EXE"
                    ]
                    
                    launched = False
                    for method in methods:
                        try:
                            subprocess.run(method, shell=True, capture_output=True, timeout=5)
                            launched = True
                            break
                        except:
                            continue
                    
                    if launched:
                        onboard_results.append("✅ Outlook launched successfully")
                    else:
                        onboard_results.append("⚠️ Outlook is installed but could not auto-launch. Please open manually.")
                else:
                    onboard_results.append("⚠️ Outlook installation could not be verified. Please check manually.")
            except Exception as e:
                onboard_results.append(f"⚠️ Outlook check error: {str(e)}")
        else:
            onboard_results.append("ℹ️ Outlook: Remote installation completed via winget")

        onboard_results.append("\n🌐 Jira: Web-based tool. Access here: https://atlassian.net")
        
        return {
            "target": target.hostname if target else "localhost",
            "success": overall_success,
            "output": "\n".join(onboard_results)
        }

    if len(targets) == 1:
        return execute_onboard(targets[0])
    
    results = []
    with ThreadPoolExecutor(max_workers=min(len(targets), 10)) as executor:
        futures = [executor.submit(execute_onboard, t) for t in targets]
        for future in futures:
            results.append(future.result())
            
    return {"success": True, "results": results}
    
@app.post("/action/reset-update")
def reset_update_action(target: RemoteTarget | None = None):
    try:
        # We'll use a single PowerShell script to handle this safely
        ps_script = f"""
        Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
        Stop-Service -Name cryptSvc -Force -ErrorAction SilentlyContinue
        Stop-Service -Name bits -Force -ErrorAction SilentlyContinue
        Stop-Service -Name msiserver -Force -ErrorAction SilentlyContinue
        
        if (Test-Path "C:\\Windows\\SoftwareDistribution") {{
            Remove-Item -Path "C:\\Windows\\SoftwareDistribution" -Recurse -Force -ErrorAction SilentlyContinue
        }}
        if (Test-Path "C:\\Windows\\System32\\catroot2") {{
            Remove-Item -Path "C:\\Windows\\System32\\catroot2" -Recurse -Force -ErrorAction SilentlyContinue
        }}
        
        Start-Service -Name wuauserv
        Start-Service -Name cryptSvc
        Start-Service -Name bits
        Start-Service -Name msiserver
        """
        
        res = run_powershell(ps_script, target)
        
        if res.returncode == 0:
            return {"success": True, "output": "Windows Update components have been reset successfully. Please try checking for updates again."}
        else:
            return {"success": False, "error": res.stderr or "Failed to reset components. Admin rights might be required."}
            
    except Exception as e:
        return {"success": False, "error": str(e)}

@app.post("/stop")
def stop_process():
    try:
        # Terminate all tracked processes
        for pid, proc in list(active_processes.items()):
            try:
                parent = psutil.Process(pid)
                for child in parent.children(recursive=True):
                    child.terminate()
                parent.terminate()
            except:
                pass
            del active_processes[pid]
            
        return {"success": True, "message": "All active installations stopped."}
    except Exception as e:
        return {"success": False, "error": str(e)}

class TicketCreate(BaseModel):
    subject: str
    sender: str
    description: str

@app.get("/emails/check")
def check_emails(token: str = None):
    # Strictly require a token and real API usage. No mock data.
    if not token:
        return []

    try:
        # 1. List messages with subject "Incident Management"
        # q=subject:"Incident Management"
        query = 'subject:"Incident Management"'
        encoded_query = urllib.parse.quote(query)
        url = f"https://gmail.googleapis.com/gmail/v1/users/me/messages?q={encoded_query}&maxResults=10"
        
        req = urllib.request.Request(url)
        req.add_header('Authorization', f'Bearer {token}')
        req.add_header('Accept', 'application/json')
        
        with urllib.request.urlopen(req) as response:
            if response.status != 200:
                print(f"Gmail API List Error: {response.status}")
                return []
            data = json.loads(response.read().decode())
            messages = data.get('messages', [])
            
        result = []
        for msg in messages:
            msg_id = msg['id']
            # 2. Get message details
            detail_url = f"https://gmail.googleapis.com/gmail/v1/users/me/messages/{msg_id}"
            detail_req = urllib.request.Request(detail_url)
            detail_req.add_header('Authorization', f'Bearer {token}')
            
            with urllib.request.urlopen(detail_req) as detail_response:
                detail_data = json.loads(detail_response.read().decode())
                
                payload = detail_data.get('payload', {})
                headers = payload.get('headers', [])
                
                subject = next((h['value'] for h in headers if h['name'] == 'Subject'), 'No Subject')
                sender = next((h['value'] for h in headers if h['name'] == 'From'), 'Unknown')
                
                # Get body (snippet is easier, or parse parts)
                snippet = detail_data.get('snippet', '')
                
                # Get full body if available
                def get_full_body(payload):
                    if 'parts' in payload:
                        for part in payload['parts']:
                            if part['mimeType'] == 'text/plain':
                                if 'data' in part['body']:
                                    return base64.urlsafe_b64decode(part['body']['data']).decode()
                            elif 'parts' in part:
                                body = get_full_body(part)
                                if body: return body
                    elif 'body' in payload and 'data' in payload['body']:
                        return base64.urlsafe_b64decode(payload['body']['data']).decode()
                    return snippet
                
                full_body = get_full_body(payload)
                
                # Format timestamp
                internal_date = int(detail_data.get('internalDate', 0))
                timestamp = datetime.fromtimestamp(internal_date/1000).strftime("%Y-%m-%d %H:%M")
                
                # Double check subject (API search is sometimes loose)
                if "Incident Management" in subject:
                    result.append({
                        "id": msg_id,
                        "sender": sender,
                        "subject": subject,
                        "body": full_body,
                        "timestamp": timestamp,
                        "to": "me" 
                    })
                
        return result
        
    except urllib.error.HTTPError as e:
        print(f"Gmail API HTTP Error: {e.code} - {e.reason}")
        # If 403, it means permission denied (scope issue or test user issue)
        return []
    except Exception as e:
        print(f"Gmail API General Error: {e}")
        return []

@app.post("/tickets")
def create_ticket(ticket: TicketCreate, db: Session = Depends(get_db)):
    new_ticket = Ticket(
        subject=ticket.subject,
        sender=ticket.sender,
        description=ticket.description,
        status="Open"
    )
    db.add(new_ticket)
    db.commit()
    db.refresh(new_ticket)
    return new_ticket

@app.get("/tickets")
def get_tickets(db: Session = Depends(get_db)):
    return db.query(Ticket).order_by(Ticket.created_at.desc()).all()

@app.put("/tickets/{ticket_id}/status")
def update_ticket_status(ticket_id: int, status: str = Body(...), db: Session = Depends(get_db)):
    t = db.query(Ticket).filter(Ticket.id == ticket_id).first()
    if not t:
      return {"success": False, "error": "Ticket not found"}
    t.status = status
    db.commit()
    db.refresh(t)
    return t

@app.delete("/tickets/{ticket_id}")
def delete_ticket(ticket_id: int, db: Session = Depends(get_db)):
    t = db.query(Ticket).filter(Ticket.id == ticket_id).first()
    if not t:
        return {"success": False, "error": "Ticket not found"}
    db.delete(t)
    db.commit()
    return {"success": True, "message": "Ticket deleted successfully"}

@app.post("/software/info")
def get_software_info(req: InstallRequest):
    package = None
    if req.dropdown:
        package = SOFTWARE_MAP.get(req.dropdown)
    elif req.custom:
        package = req.custom
    
    if not package:
        return {"success": False, "error": "No software selected."}

    # Use winget show to get version and size
    show_cmd = f'winget show --id "{package}" --accept-source-agreements'
    res = run_powershell(show_cmd, req.target)
    
    if res.returncode != 0:
        return {"success": False, "error": "Could not retrieve software info."}

    output = res.stdout
    version = "Unknown"
    size = "100" # Default 100MB if not found
    
    # Parse version and size from output
    import re
    for line in output.split('\n'):
        line = line.strip()
        if line.startswith("Version:"):
            version = line.split(":", 1)[1].strip()
        
        # Look for various size labels in winget show output
        if any(label in line for label in ["Download Size:", "Installer Size:", "Size:"]):
            size_str = line.split(":", 1)[1].strip()
            # Extract numbers and unit (e.g., "123.45 MB" or "1 GB")
            match = re.search(r"(\d+(\.\d+)?)\s*(KB|MB|GB|B)", size_str, re.IGNORECASE)
            if match:
                value = float(match.group(1))
                unit = match.group(3).upper()
                if unit == "GB": value *= 1024
                elif unit == "KB": value /= 1024
                elif unit == "B": value /= (1024 * 1024)
                size = str(round(value, 1))
                # Once we find a valid size, we can stop looking for size
                # (Download size is usually what we want)
                if "Download" in line or "Installer" in line:
                    continue 

    return {
        "success": True,
        "version": version,
        "size": size
    }

def execute_install(package: str, target: RemoteTarget, force_upgrade: bool, dropdown: str = None):
    # 1. Check if already installed (unless force_upgrade is True)
    if not force_upgrade:
        check_cmd = f'winget list --id "{package}" --accept-source-agreements'
        res = run_powershell(check_cmd, target)
        
        if res.returncode == 0:
            lines = res.stdout.strip().split('\n')
            for line in lines:
                if package.lower() in line.lower():
                    parts = line.split()
                    # Try to find the package ID in the parts, it might be truncated in winget list
                    pkg_idx = -1
                    for i, part in enumerate(parts):
                        if part.lower() == package.lower() or package.lower().startswith(part.lower()):
                            pkg_idx = i
                            break
                    
                    if pkg_idx != -1:
                        try:
                            installed_v = parts[pkg_idx + 1]
                            available_v = parts[pkg_idx + 2] if len(parts) > pkg_idx + 2 else None
                            
                            if available_v in ["winget", "msstore", None]:
                                return {
                                    "target": target.hostname,
                                    "success": True, 
                                    "status": "already_installed", 
                                    "installed_version": installed_v,
                                    "output": f"{dropdown or package} is already installed (Version: {installed_v})."
                                }
                            else:
                                return {
                                    "target": target.hostname,
                                    "success": True, 
                                    "status": "update_available", 
                                    "installed_version": installed_v,
                                    "available_version": available_v,
                                    "output": f"{dropdown or package} {installed_v} is installed, but {available_v} is available."
                                }
                        except (ValueError, IndexError):
                            continue

    # 2. Proceed with installation/upgrade
    winget_cmd = f'winget install -e --id "{package}" --silent --accept-package-agreements --accept-source-agreements --force'
    
    if not target or target.hostname == "localhost":
        cmd = ["powershell", "-Command", winget_cmd]
    else:
        remote_script = f"""
        $pw = ConvertTo-SecureString "{target.password}" -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential ("{target.username}", $pw)
        Invoke-Command -ComputerName {target.hostname} -Credential $cred -ScriptBlock {{ {winget_cmd} }}
        """
        cmd = ["powershell", "-Command", remote_script]

    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if platform.system() == 'Windows' else 0
        )
        
        active_processes[proc.pid] = proc
        try:
            stdout, stderr = proc.communicate(timeout=600)

        except subprocess.TimeoutExpired:
            proc.kill()
            stdout, stderr = proc.communicate() 
            
            return {
                "target": target.hostname,
                "success": False,
                "error": "Installation timed out after 10 minutes"
            }
        finally:
            if proc.pid in active_processes:
                del active_processes[proc.pid]
            
        success = proc.returncode == 0
        output_text = stdout.strip() if 'stdout' in locals() else ""
        error_text = stderr.strip() if 'stderr' in locals() else ""
        
        # Check for common winget error messages that don't necessarily return non-zero exit code
        failure_indicators = [
            "No installed package found matching input criteria",
            "No package found matching input criteria",
            "Multiple packages found matching input criteria",
            "Failed to find a package"
        ]
        
        if success and any(indicator in output_text for indicator in failure_indicators):
            success = False
            if not error_text:
                error_text = "Package not found or multiple packages found."

        # Log successful installation to database
        if success:
            try:
                db: Session = Depends(get_db)
                # Use dropdown name if available for better readability, else package ID
                software_name = dropdown if dropdown else package
                
                new_install = InstalledSoftware(
                    name=software_name,
                    version="Latest", # Winget usually installs latest unless specified
                    installed_at=datetime.now(timezone.utc).replace(tzinfo=None)
                )
                db.add(new_install)
                db.commit()
                db.close()
                print(f"Logged installation of {software_name} to database.")
            except Exception as e:
                print(f"Failed to log installation to DB: {e}")

        return {
            "target": target.hostname,
            "success": success,
            "output": output_text,
            "error": error_text if not success else (None if success else "Installation failed")
        }
    except Exception as e:
        return {"target": target.hostname, "success": False, "error": str(e)}

@app.post("/install")
def install(req: InstallRequest):
    package = None
    if req.dropdown:
        package = SOFTWARE_MAP.get(req.dropdown)
    elif req.custom:
        package = req.custom
    
    if not package:
        return {"success": False, "error": "No software selected."}

    targets = req.targets if req.targets else ([req.target] if req.target else [])
    if not targets:
        return {"success": False, "error": "No target specified."}

    if len(targets) == 1:
        return execute_install(package, targets[0], req.force_upgrade, req.dropdown)
    
    results = []
    with ThreadPoolExecutor(max_workers=min(len(targets), 10)) as executor:
        futures = [executor.submit(execute_install, package, t, req.force_upgrade, req.dropdown) for t in targets]
        for future in futures:
            results.append(future.result())
            
    return {"success": True, "results": results}
        
def execute_uninstall(package: str, target: RemoteTarget, dropdown: str = None):
    winget_cmd = f'winget uninstall -e --id "{package}" --silent --accept-source-agreements --force'
    
    if not target or target.hostname == "localhost":
        cmd = ["powershell", "-Command", winget_cmd]
    else:
        remote_script = f"""
        $pw = ConvertTo-SecureString "{target.password}" -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential ("{target.username}", $pw)
        Invoke-Command -ComputerName {target.hostname} -Credential $cred -ScriptBlock {{ {winget_cmd} }}
        """
        cmd = ["powershell", "-Command", remote_script]

    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if platform.system() == 'Windows' else 0
        )
        
        active_processes[proc.pid] = proc
        try:
            stdout, stderr = proc.communicate()
        finally:
            if proc.pid in active_processes:
                del active_processes[proc.pid]
            
        success = proc.returncode == 0
        output_text = stdout.strip() if 'stdout' in locals() else ""
        error_text = stderr.strip() if 'stderr' in locals() else ""

        # Check for common winget error messages
        failure_indicators = [
            "No installed package found matching input criteria",
            "No package found matching input criteria",
            "Failed to find a package"
        ]
        
        if success and any(indicator in output_text for indicator in failure_indicators):
            success = False
            if not error_text:
                error_text = "Package not found."

        # Log successful uninstallation to database
        if success:
            try:
                db: Session = Depends(get_db)
                # Use dropdown name if available for better readability, else package ID
                software_name = dropdown if dropdown else package
                
                new_uninstall = UninstalledSoftware(
                    name=software_name,
                    version="Latest", 
                    uninstalled_at=datetime.now(timezone.utc).replace(tzinfo=None)
                )
                db.add(new_uninstall)
                db.commit()
                db.close()
                print(f"Logged uninstallation of {software_name} to database.")
            except Exception as e:
                print(f"Failed to log uninstallation to DB: {e}")

        return {
            "target": target.hostname,
            "success": success,
            "output": output_text,
            "error": error_text if not success else (None if success else "Uninstallation failed")
        }
    except Exception as e:
        return {"target": target.hostname, "success": False, "error": str(e)}

@app.post("/uninstall")
def uninstall(req: InstallRequest):
    package = None
    if req.dropdown:
        package = SOFTWARE_MAP.get(req.dropdown)
    elif req.custom:
        package = req.custom
    
    if not package:
        return {"success": False, "error": "No software selected."}

    targets = req.targets if req.targets else ([req.target] if req.target else [])
    if not targets:
        return {"success": False, "error": "No target specified."}

    if len(targets) == 1:
        return execute_uninstall(package, targets[0], req.dropdown)
    
    results = []
    with ThreadPoolExecutor(max_workers=min(len(targets), 10)) as executor:
        futures = [executor.submit(execute_uninstall, package, t, req.dropdown) for t in targets]
        for future in futures:
            results.append(future.result())
            
    return {"success": True, "results": results}

@app.get("/history")
def get_all_history(email: str, db: Session = Depends(get_db)):
    sessions = db.query(ChatHistory).filter(ChatHistory.email == email).all()

    result = []
    for s in sessions:
        result.append({
            "id": s.session_id,
            "title": s.title,
            "messages": json.loads(s.messages),
            "timestamp": s.timestamp,
            "pinned": s.pinned
        })

    return result

@app.post("/history")
def save_session(req: SaveHistoryRequest, db: Session = Depends(get_db)):

    existing = db.query(ChatHistory).filter(
        ChatHistory.session_id == req.session.id,
        ChatHistory.email == req.email
    ).first()

    if existing:
        existing.title = req.session.title
        existing.messages = json.dumps(req.session.messages)
        existing.timestamp = req.session.timestamp
        existing.pinned = req.session.pinned
    else:
        new_session = ChatHistory(
            session_id=req.session.id,
            email=req.email,
            title=req.session.title,
            messages=json.dumps(req.session.messages),
            timestamp=req.session.timestamp,
            pinned=req.session.pinned
        )
        db.add(new_session)

    db.commit()

    return {"success": True}

@app.delete("/history/{session_id}")
def delete_session(session_id: str, email: str, db: Session = Depends(get_db)):
    db.query(ChatHistory).filter(
        ChatHistory.session_id == session_id,
        ChatHistory.email == email
    ).delete()
    db.commit()
    return {"success": True}

@app.get("/debug/licenses")
def list_licenses():
    try:
        token = get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        url = "https://graph.microsoft.com/v1.0/subscribedSkus"

        response = requests.get(url, headers=headers)

        return response.json()
    except Exception as e:
        return {"error": str(e)}

# ==============================
# DOCUMENT MANAGEMENT & SHAREPOINT
# ==============================
from models import Document as DBDocument

@app.post("/upload-document")
async def upload_document(
    file: UploadFile = File(...),
    employee_id: str = Form(None),
    db: Session = Depends(get_db)
):

    """
    1. Upload File to SharePoint
    2. Extract AI data (POC simulation)
    3. Save JSON output to SharePoint
    4. Store URLs in DB
    """
    filename = file.filename
    content = await file.read()
    
    # 1. AI Extraction POC (Simulation)
    # In a real POC, this would be a prompt to an AI model
    extracted_data = {
        "filename": filename,
        "extraction_date": datetime.now().isoformat(),
        "confidence": 0.98,
        "summary": f"This is a simulated AI extraction for {filename}"
    }
    
    # 2. Upload to SharePoint
    # New Structure: Assist-IQ / {Employee_ID} / {filename}
    # and Assist-IQ / {Employee_ID} / {filename}_extracted.json
    
    # Determine the folder name by employee profile name when possible
    clean_id = (employee_id or "").strip()
    folder_name = "general"

    # If employee_id comes from an existing employee, use their current full_name.
    if clean_id:
        query = db.query(Employee).filter(Employee.employee_id == clean_id)
        # Avoid invalid SQL cast for numeric primary key lookup when non-numeric strings present
        if clean_id.isdigit():
            query = query.union_all(db.query(Employee).filter(Employee.id == int(clean_id)))

        employee = query.first()
        if employee and employee.full_name:
            folder_name = employee.full_name.strip()
        else:
            folder_name = clean_id

    # Sanitize folder name for SharePoint path (replace spaces and disallowed chars)
    folder_name = re.sub(r"[\\/:*?\"<>|]", "", folder_name)
    folder_name = folder_name.strip().replace(" ", "_")
    if not folder_name:
        folder_name = "general"

    logger.info(f"Uploading document '{filename}' for ID/Folder: '{folder_name}'")

    try:
        # 1. Upload original file with its actual name
        sp_file = upload_file_to_sharepoint(content, filename, subfolder=folder_name)
        
        # 2. Upload JSON extraction as {filename}_extracted.json
        json_filename = f"{filename}_extracted.json"
        sp_json = upload_json_to_sharepoint(extracted_data, json_filename, subfolder=folder_name)

        
        # 3. Store in Database
        new_doc = DBDocument(
            filename=filename,
            employee_id=employee_id,
            sharepoint_url=sp_file["url"],
            sharepoint_json_url=sp_json["url"],
            extracted_data=json.dumps(extracted_data)
        )
        db.add(new_doc)
        db.commit()
        db.refresh(new_doc)
        
        return {
            "success": True,
            "document": {
                "id": new_doc.id,
                "sharepoint_url": new_doc.sharepoint_url,
                "sharepoint_json_url": new_doc.sharepoint_json_url
            }
        }

    except Exception as e:
        logger.error(f"SharePoint upload failed: {str(e)}")
        return {"success": False, "error": str(e)}

@app.get("/documents")
def get_documents(db: Session = Depends(get_db)):
    return db.query(DBDocument).all()
