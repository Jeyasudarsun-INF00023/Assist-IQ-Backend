"""
Quick test to verify /activity endpoint is working correctly.
Run this from c:\auto_installer_project:
  python test_activity.py
"""
import requests

BACKEND_URL = "http://localhost:8000/activity"
TOKEN = "SECRET123"
EMP_ID = "INFI0023"

payload = {
    "employee_id": EMP_ID,
    "app": "TestApp.exe",
    "window": "Test Window - Verifying Live Update",
    "device_id": "TEST-DEVICE-ID-123"
}

headers = {
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type": "application/json"
}

print(f"Sending POST to {BACKEND_URL} ...")
try:
    resp = requests.post(BACKEND_URL, json=payload, headers=headers, timeout=5)
    print(f"Status: {resp.status_code}")
    print(f"Response: {resp.text}")
    if resp.status_code == 200:
        print("\n✅ SUCCESS! Backend received the activity.")
        print("   Check the dashboard - it should update NOW.")
    elif resp.status_code == 403:
        print("\n❌ FORBIDDEN! Token mismatch or employee not found.")
        print("   Run: python backend/migrate_v4.py  to reset tokens.")
    elif resp.status_code == 401:
        print("\n❌ UNAUTHORIZED! Token header missing.")
except Exception as e:
    print(f"\n❌ CONNECTION ERROR: {e}")
    print("   Is the backend running? Try: uvicorn main:app --reload")
