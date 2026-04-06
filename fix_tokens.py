import sys
import os
import secrets
sys.path.append(os.path.join(os.getcwd(), 'backend'))
from database import SessionLocal
from models import Employee

db = SessionLocal()
emps = db.query(Employee).filter(Employee.api_token == None).all()
print(f"Found {len(emps)} employees without tokens. Fixing...")

for e in emps:
    e.api_token = secrets.token_hex(16)
    print(f"Assigned token to {e.employee_id}")

db.commit()
db.close()
print("Done.")
