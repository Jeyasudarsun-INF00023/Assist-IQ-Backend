from sqlalchemy import Column, Integer, String, DateTime, Boolean, Text
from datetime import datetime, timezone
from database import Base
import json


def utcnow_naive() -> datetime:
    # SQLite DateTime columns are naive; store UTC consistently without utcnow().
    return datetime.now(timezone.utc).replace(tzinfo=None)

class InstalledSoftware(Base):
    __tablename__ = "installed_software"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)
    version = Column(String)
    installed_at = Column(DateTime, default=utcnow_naive)

class UninstalledSoftware(Base):
    __tablename__ = "uninstalled_software"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)
    version = Column(String)
    uninstalled_at = Column(DateTime, default=utcnow_naive)

class ChatHistory(Base):
    __tablename__ = "chat_history"

    id = Column(Integer, primary_key=True, index=True)
    session_id = Column(String, index=True)
    email = Column(String, index=True)
    title = Column(String)
    messages = Column(Text)  # We store JSON as string
    timestamp = Column(String)
    pinned = Column(Boolean, default=False)

class Device(Base):
    __tablename__ = "devices"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=True)
    ip_address = Column(String, unique=True, index=True, nullable=False)
    username = Column(String, nullable=True)
    password = Column(String, nullable=True)

class Ticket(Base):
    __tablename__ = "tickets"

    id = Column(Integer, primary_key=True, index=True)
    sender = Column(String)
    subject = Column(String)
    description = Column(String)
    status = Column(String, default="Open")
    created_at = Column(DateTime, default=utcnow_naive)

class Employee(Base):
    __tablename__ = "employees"

    id = Column(Integer, primary_key=True, index=True)
    full_name = Column(String, nullable=False)
    first_name = Column(String, nullable=True)
    last_name = Column(String, nullable=True)
    phone_number = Column(String, nullable=True)
    personal_email = Column(String, nullable=True)
    address = Column(String, nullable=True)
    experience_level = Column(String, nullable=True)
    role = Column(String)
    employee_id = Column(String, unique=True, index=True)
    email = Column(String, nullable=True)
    temp_password = Column(String, nullable=True)
    avatar = Column(String, nullable=True)
    laptop = Column(String, nullable=True)
    mouse = Column(String, nullable=True)
    headphone = Column(String, nullable=True)
    department = Column(String, nullable=True)
    seat_id = Column(Integer, nullable=True)
    documents = Column(Text, nullable=True)  # Store as JSON string
    created_at = Column(DateTime, default=utcnow_naive)
    offboarded = Column(Boolean, default=False)
    offboarded_at = Column(DateTime, nullable=True)
    last_app = Column(String, nullable=True)
    last_window = Column(String, nullable=True)
    is_online = Column(Boolean, default=False)
    last_seen = Column(DateTime, nullable=True)
    api_token = Column(String, nullable=True)
    device_id = Column(String, nullable=True)
    verification_code = Column(String, nullable=True)
    verification_code_expires = Column(DateTime, nullable=True)

class ActivityLog(Base):
    __tablename__ = "activity_logs"
    id = Column(Integer, primary_key=True, index=True)
    employee_id = Column(String, index=True)
    app = Column(String)
    window = Column(String)
    start_time = Column(DateTime, default=utcnow_naive)
    end_time = Column(DateTime, nullable=True)

class Asset(Base):
    __tablename__ = "assets"

    id = Column(Integer, primary_key=True, index=True)
    type = Column(String, nullable=False) # Laptop, Mouse, Headset, etc.
    category = Column(String, nullable=True)  # Employee or Office
    brand = Column(String)
    model = Column(String)
    sn = Column(String, unique=True, index=True)
    processor = Column(String, nullable=True)
    ram = Column(String, nullable=True)
    storage = Column(String, nullable=True)
    os = Column(String, nullable=True)
    assignee = Column(String, nullable=True)
    assigned_date = Column(String, nullable=True)
    remarks = Column(Text, nullable=True)
    price = Column(String, nullable=True)
    custom_fields = Column(Text, nullable=True)  # saved JSON object as string
    created_at = Column(DateTime, default=utcnow_naive)

class Document(Base):
    __tablename__ = "documents"
    id = Column(Integer, primary_key=True, index=True)
    filename = Column(String)
    employee_id = Column(String, nullable=True) # Linked to employee_id
    sharepoint_url = Column(String, nullable=True)
    sharepoint_json_url = Column(String, nullable=True)
    extracted_data = Column(Text, nullable=True) # JSON of AI extracted info
    created_at = Column(DateTime, default=utcnow_naive)


