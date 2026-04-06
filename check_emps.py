import sys
import os
sys.path.append(os.path.join(os.getcwd(), 'backend'))
from database import SessionLocal
from models import Employee

db = SessionLocal()
emps = db.query(Employee).all()
for e in emps:
    print(f"ID: {e.employee_id}, Token: {e.api_token}")
db.close()
